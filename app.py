"""
flask_app.py
============
Production Flask application for Batching Incentive & Deduction Calculator.
Replaces the Streamlit front-end; all Python business-logic modules are
kept exactly as-is (calculator.py, database.py, oracle_connector.py …).

Run with:
    python flask_app.py
Opens on http://localhost:2001
"""

from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import threading
from datetime import date as _date, datetime as _dt, timedelta

import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from flask import (Flask, flash, jsonify, redirect, render_template,
                   request, send_file, url_for)

import cache_helpers
import calculator
import config
import data_loader
import database
import email_helper
import google_sheets
import oracle_connector
import report_generator
import tp_calculator
import validations

# ── App setup ────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "rdc-incentive-local-key-2025")

# ── Single-user in-memory state (no session serialization overhead) ──────────
_S: dict = {}


def _s(key, default=None):
    return _S.get(key, default)


def _ss(key, val):
    _S[key] = val


# ── Module-scoped session state ───────────────────────────────────────────────
# New modules MUST use _ms/_mss instead of _s/_ss so keys never collide.
# Keys are stored as "{module}.{key}" inside the shared _S dict.
# Example:  _mss("tp", "results", df)  →  _S["tp.results"] = df
#           _ms("tp", "results")       →  _S.get("tp.results")
def _ms(module: str, key: str, default=None):
    return _S.get(f"{module}.{key}", default)


def _mss(module: str, key: str, val):
    _S[f"{module}.{key}"] = val


# ── Boot ─────────────────────────────────────────────────────────────────────
database.init_db()

# Apply the rolling Oracle-data retention once at startup (covers the case
# where the app was restarted after a month rolled over).
try:
    _purge = database.purge_old_oracle_data()
    if any(_purge["deleted"].values()):
        print(f"[retention] purged Oracle data before {_purge['cutoff']}: {_purge['deleted']}")
except Exception as _exc:
    print(f"[retention] startup purge skipped: {_exc}")

# ── Monthly email scheduler ───────────────────────────────────────────────────
_scheduler: BackgroundScheduler | None = None


def _retention_job():
    """Daily: enforce the shared rolling Oracle-data retention for all modules."""
    try:
        database.purge_old_oracle_data()
    except Exception:
        pass


def _scheduled_email_job():
    """Fires on 1st of each month — sends previous month's report."""
    if database.get_setting("email_schedule_enabled", "false") != "true":
        return
    today         = _date.today()
    first_this    = today.replace(day=1)
    last_prev     = first_this - timedelta(days=1)
    first_prev    = last_prev.replace(day=1)
    from_s, to_s  = str(first_prev), str(last_prev)
    try:
        res = calculator.run_calculation(
            month=first_prev.month, year=first_prev.year,
            start_date=from_s, end_date=to_s, persist=False,
        )
        if res["error"]:
            database.set_setting("email_schedule_last_status",
                f"FAIL {_dt.now():%Y-%m-%d %H:%M} — {res['error']}")
            return
        rows     = res.get("results_rows", [])
        unmapped = res.get("unmapped_rows", [])
        month_label = first_prev.strftime("%B %Y")
        meta    = _build_meta(from_s, to_s, rows, unmapped, "Scheduled monthly report")
        df_f    = pd.DataFrame(rows)    if rows     else pd.DataFrame(columns=RESULT_COLS)
        df_u    = pd.DataFrame(unmapped) if unmapped else pd.DataFrame()
        val_df  = database.read_table("validation_errors")
        xlsx    = report_generator.generate_excel_report(df_f, df_u, val_df, meta)
        fname   = f"incentive_report_{from_s}_to_{to_s}.xlsx"
        to_addr = database.get_setting("email_schedule_to", "")
        cc_addr = database.get_setting("email_schedule_cc", "")
        result  = email_helper.send_report_email(
            to_emails=to_addr, cc_emails=cc_addr,
            subject=email_helper.compose_report_subject(month_label),
            body=email_helper.compose_report_body(month_label),
            attachment_bytes=xlsx, attachment_name=fname,
        )
        status = (f"OK {_dt.now():%Y-%m-%d %H:%M} — sent to {to_addr}" if result["success"]
                  else f"FAIL {_dt.now():%Y-%m-%d %H:%M} — {result.get('error','')}")
        database.set_setting("email_schedule_last_status", status)
    except Exception as exc:
        database.set_setting("email_schedule_last_status",
            f"ERROR {_dt.now():%Y-%m-%d %H:%M} — {exc}")


def _build_tp_excel(plant_rows, location_rows) -> bytes:
    """Build the RDC-TP report workbook (Plant + Location sheets) as bytes."""
    plant_df = pd.DataFrame(plant_rows)[[
        "lookup_code","plant_name","exco_location","business_head",
        "plant_manager","mixer_theo_cap","total_quantity","total_time_min",
        "throughput_pct","batch_count"
    ]].rename(columns={
        "lookup_code":"Plant Code","plant_name":"Plant","exco_location":"Exco Location",
        "business_head":"Business Head","plant_manager":"Plant Manager",
        "mixer_theo_cap":"Mixer Capacity","total_quantity":"Total Qty",
        "total_time_min":"Total Time (min)","throughput_pct":"Throughput %",
        "batch_count":"Batches",
    })
    loc_df = pd.DataFrame(location_rows)[[
        "exco_location","plant_count","total_quantity","avg_throughput_pct"
    ]].rename(columns={
        "exco_location":"Exco Location","plant_count":"Plants",
        "total_quantity":"Total Qty","avg_throughput_pct":"Avg Throughput %",
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        plant_df.to_excel(writer, sheet_name="Plant Throughput", index=False)
        loc_df.to_excel(writer, sheet_name="Location Throughput", index=False)
    buf.seek(0)
    return buf.read()


def _tp_daily_oracle_fetch_job():
    """
    Runs every day at 00:10.
    Fetches Oracle TP data for the previous month (full) and the current month
    (1st → today), replaces what is stored, then purges data older than the
    previous month. This keeps the local store always current with minimal load.
    """
    if not oracle_connector.is_configured():
        return
    today      = _date.today()
    # Current month: 1st → today
    cur_from   = today.replace(day=1)
    cur_to     = today
    # Previous month: 1st → last day
    if today.month == 1:
        prev_from = _date(today.year - 1, 12, 1)
        import calendar as _cal
        prev_to   = _date(today.year - 1, 12, _cal.monthrange(today.year - 1, 12)[1])
    else:
        import calendar as _cal
        prev_from = _date(today.year, today.month - 1, 1)
        prev_to   = _date(today.year, today.month - 1,
                          _cal.monthrange(today.year, today.month - 1)[1])
    try:
        for fd, td in [(str(prev_from), str(prev_to)), (str(cur_from), str(cur_to))]:
            raw_df, _ = oracle_connector.fetch_tp_data(fd, td)
            if not raw_df.empty:
                parsed, _ = tp_calculator.parse_oracle_df(raw_df)
                oracle_connector.save_tp_oracle_data(raw_df, fd, td, parsed, replace=True)
        database.purge_old_oracle_data()
        print(f"[tp-daily-fetch] done — prev: {prev_from}→{prev_to}, cur: {cur_from}→{cur_to}")
    except Exception as exc:
        print(f"[tp-daily-fetch] error: {exc}")


def _tp_scheduled_email_job():
    """Fires on 1st of each month — emails the previous month's TP report."""
    if database.get_module_setting("tp", "email_schedule_enabled", "false") != "true":
        return
    today      = _date.today()
    first_this = today.replace(day=1)
    last_prev  = first_this - timedelta(days=1)
    first_prev = last_prev.replace(day=1)
    fd, td     = str(first_prev), str(last_prev)
    try:
        month, year = first_prev.month, first_prev.year
        plant_rows, loc_rows, _w = tp_calculator.run_tp_calculation(
            month, year, from_date=fd, to_date=td)
        if not plant_rows:
            database.set_module_setting("tp", "email_schedule_last_status",
                f"FAIL {_dt.now():%Y-%m-%d %H:%M} — no data for {fd} → {td}")
            return
        import calendar
        label   = f"{calendar.month_name[month]} {year}"
        subject = f"RDC-TP Plant Throughput Report — {label}"
        body    = (f"Dear Team,\n\nPlease find attached the Plant Throughput Report "
                   f"for {label}.\n\nRegards,\nRDC Operations")
        res = email_helper.send_report_email(
            to_emails=database.get_module_setting("tp", "email_schedule_to", ""),
            cc_emails=database.get_module_setting("tp", "email_schedule_cc", ""),
            subject=subject, body=body,
            attachment_bytes=_build_tp_excel(plant_rows, loc_rows),
            attachment_name=f"RDC_TP_{year}_{month:02d}.xlsx",
        )
        status = (f"OK {_dt.now():%Y-%m-%d %H:%M} — sent" if res["success"]
                  else f"FAIL {_dt.now():%Y-%m-%d %H:%M} — {res.get('error','')}")
        database.set_module_setting("tp", "email_schedule_last_status", status)
    except Exception as exc:
        database.set_module_setting("tp", "email_schedule_last_status",
            f"ERROR {_dt.now():%Y-%m-%d %H:%M} — {exc}")


def _start_scheduler():
    global _scheduler
    sched_time = database.get_setting("email_schedule_time", "08:00")
    try:
        h, m = map(int, sched_time.split(":"))
    except Exception:
        h, m = 8, 0
    _scheduler = BackgroundScheduler(daemon=True)
    _scheduler.add_job(_scheduled_email_job,
                       CronTrigger(day=1, hour=h, minute=m),
                       id="monthly_report", replace_existing=True)
    # Daily rolling retention for all modules' Oracle data (runs 00:05 + on 1st).
    _scheduler.add_job(_retention_job,
                       CronTrigger(hour=0, minute=5),
                       id="oracle_retention", replace_existing=True)
    # RDC-TP monthly report (own schedule time).
    tp_time = database.get_module_setting("tp", "email_schedule_time", "08:00")
    try:
        th, tm = map(int, tp_time.split(":"))
    except Exception:
        th, tm = 8, 0
    _scheduler.add_job(_tp_scheduled_email_job,
                       CronTrigger(day=1, hour=th, minute=tm),
                       id="tp_monthly_email", replace_existing=True)
    # RDC-TP daily Oracle fetch — keeps previous month + current month data fresh.
    _scheduler.add_job(_tp_daily_oracle_fetch_job,
                       CronTrigger(hour=0, minute=10),
                       id="tp_daily_oracle_fetch", replace_existing=True)
    _scheduler.start()


_start_scheduler()

# ── Jinja filters / globals ───────────────────────────────────────────────────

@app.template_filter("format_int")
def _fmt_int(v):
    try:
        return f"{int(v):,}"
    except Exception:
        return v


@app.template_global()
def bg_auto():
    return database.get_setting("bg_auto_theme", "true") == "true"


@app.template_global()
def bg_animate():
    return database.get_setting("bg_animate", "true") == "true"


@app.template_global()
def bg_theme():
    return database.get_setting("bg_manual_theme", "Daytime")


# ── Context processor (sidebar active-page + bg settings) ──────────────────

@app.context_processor
def _ctx():
    return dict(
        active_page=_s("active_page", ""),
        bg_auto=bg_auto(),
        bg_animate=bg_animate(),
        bg_theme=bg_theme(),
    )


# ── Helpers ───────────────────────────────────────────────────────────────────

def _records(df: pd.DataFrame) -> list:
    if df is None or df.empty:
        return []
    return json.loads(df.to_json(orient="records", date_format="iso"))


def _row_cls(row: dict) -> str:
    if (row.get("deduction_amount") or 0) > 0:
        return "row-red"
    if (row.get("incentive_amount") or 0) > 0:
        return "row-green"
    return ""


def _sort_rows(rows: list) -> list:
    def _key(r):
        d = r.get("deduction_amount") or 0
        i = r.get("incentive_amount") or 0
        if d > 0:
            return (0, -d, -i)
        if i > 0:
            return (1, 0, -i)
        return (2, 0, 0)
    return sorted(rows, key=_key)


def _apply_filters(rows: list, cats, desigs, plants, elig, outcome, search):
    out = rows
    if cats:
        out = [r for r in out if r.get("category") in cats]
    if desigs:
        out = [r for r in out if r.get("designation") in desigs]
    if plants:
        out = [r for r in out if r.get("plant") in plants]
    if elig == "Yes":
        out = [r for r in out if r.get("incentive_eligible") == "Yes"]
    elif elig == "No":
        out = [r for r in out if r.get("incentive_eligible") == "No"]
    if outcome == "incentive":
        out = [r for r in out if (r.get("incentive_amount") or 0) > 0]
    elif outcome == "deduction":
        out = [r for r in out if (r.get("deduction_amount") or 0) > 0]
    elif outcome == "both":
        out = [r for r in out if (r.get("incentive_amount") or 0) > 0 and (r.get("deduction_amount") or 0) > 0]
    elif outcome == "neither":
        out = [r for r in out if (r.get("incentive_amount") or 0) == 0 and (r.get("deduction_amount") or 0) == 0]
    if search.strip():
        s = search.strip().lower()
        out = [r for r in out if s in str(r.get("employee_code", "")).lower()
               or s in str(r.get("employee_name", "")).lower()]
    return out


RESULT_COLS = [
    "employee_code", "employee_name", "designation",
    "plant", "plant_code",
    "total_quantity", "ytd_maintenance_cost",
    "incentive_eligible", "incentive_rate", "incentive_amount",
    "deduction_target", "shortfall_quantity", "deduction_amount",
    "remarks",
]
RESULT_LABELS = {
    "employee_code":        "Emp Code",
    "employee_name":        "Name",
    "designation":          "Designation",
    "plant":                "Plant",
    "plant_code":           "Plant Code",
    "total_quantity":       "Total Qty",
    "ytd_maintenance_cost": "R&M Cost YTD",
    "incentive_eligible":   "Eligible",
    "incentive_rate":       "Inc Rate",
    "incentive_amount":     "Incentive Amount (Rs)",
    "deduction_target":     "Ded Target",
    "shortfall_quantity":   "Shortfall",
    "deduction_amount":     "Deduction Amount (Rs)",
    "remarks":              "Remarks",
}

# Reports page gets two extra plant-aggregate columns
REPORT_COLS = RESULT_COLS + ["plant_total_incentive", "plant_total_deduction"]
REPORT_LABELS = {
    **RESULT_LABELS,
    "plant_total_incentive": "Plant Incentive (Rs)",
    "plant_total_deduction": "Plant Deduction (Rs)",
}
CAT_TABS = {
    "All Trainees":       ["Civil Trainee", "Non-Civil Trainee"],
    "Plant Mgr & PI":     ["PM & API"],
    "QCI":                ["QCI"],
    "All MO":             ["MO"],
    "SPE":                ["SPE"],
    "TL Employee":        ["TL BPO"],
    "Production Officer": ["Production Officer"],
    "NA":                 ["NA"],
}


def _build_meta(from_date, to_date, filtered, unmapped, applied_filters):
    return {
        "generated_on":     _dt.now().isoformat(timespec="seconds"),
        "date_range":       f"{from_date} to {to_date}",
        "applied_filters":  applied_filters,
        "total_employees":  len(filtered),
        "total_quantity":   sum(r.get("total_quantity") or 0 for r in filtered),
        "total_incentive":  sum(r.get("incentive_amount") or 0 for r in filtered),
        "total_deduction":  sum(r.get("deduction_amount") or 0 for r in filtered),
        "unmapped_count":   len(unmapped),
        "validation_count": database.get_table_counts().get("validation_errors", 0),
    }


def _rows_to_df(rows: list, cols) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows)[list(cols)]


# ── PAGES ─────────────────────────────────────────────────────────────────────

@app.route("/")
def page_home():
    return render_template("home.html")

@app.route("/id")
def page_id_dashboard():
    return redirect(url_for("page_dashboard"))

@app.route("/dashboard")
def page_dashboard():
    _ss("active_page", "dashboard")
    counts = database.get_table_counts()
    return render_template("dashboard.html",
                           counts=counts,
                           db_path=database.DB_PATH)

# ═══════════════════════════════════════════════════════════════════════════
# RDC-TP MODULE — Plant Throughput Calculator
# ═══════════════════════════════════════════════════════════════════════════

def _tp_ctx():
    """Base context dict for every TP page."""
    return dict(
        active_page="",
        bg_auto=bg_auto(), bg_animate=bg_animate(), bg_theme=bg_theme(),
    )


@app.route("/tp")
def page_tp():
    return redirect(url_for("tp_dashboard"))


# ── Dashboard ─────────────────────────────────────────────────────────────────
@app.route("/tp/dashboard")
def tp_dashboard():
    counts = {
        "tp_plant_data":  database.get_table_counts().get("tp_plant_data", 0),
        "tp_oracle_data": database.get_table_counts().get("tp_oracle_data", 0),
        "tp_results":     database.get_table_counts().get("tp_results", 0),
    }
    last_sync  = google_sheets.get_tp_last_sync_info()
    ora_ready  = oracle_connector.is_configured()
    ctx = _tp_ctx()
    ctx["active_page"] = "dashboard"
    return render_template("tp_dashboard.html", counts=counts,
                           last_sync=last_sync, ora_ready=ora_ready, **ctx)


# ── Data Uploader ─────────────────────────────────────────────────────────────
@app.route("/tp/data-uploader")
def tp_data_uploader():
    ora_ready    = oracle_connector.is_configured()
    last_sync    = google_sheets.get_tp_last_sync_info()
    sheet_id     = database.get_module_setting("tp", "gsheet_id",
                                               database.get_setting("gsheet_id", ""))
    tp_worksheet = database.get_module_setting("tp", "gsheet_worksheet", "Plant Data for TP")
    ora_counts   = database.get_table_counts().get("tp_oracle_data", 0)

    plant_rows  = database.get_tp_plants()
    plant_count = len(plant_rows)
    tp_codes    = database.get_tp_plant_codes()
    log_rows    = database.get_tp_plant_log()

    edit_code  = request.args.get("edit_code", "").strip()
    del_code   = request.args.get("del_code", "").strip()
    edit_plant = database.get_tp_plant(edit_code) if edit_code else None
    del_plant  = database.get_tp_plant(del_code)  if del_code  else None

    ctx = _tp_ctx()
    ctx["active_page"] = "data_uploader"
    return render_template("tp_data_uploader.html",
                           ora_ready=ora_ready, last_sync=last_sync,
                           sheet_id=sheet_id, tp_worksheet=tp_worksheet,
                           ora_counts=ora_counts,
                           plant_rows=plant_rows, plant_count=plant_count,
                           tp_codes=tp_codes, log_rows=log_rows,
                           edit_plant=edit_plant, del_plant=del_plant,
                           **ctx)


@app.route("/tp/action/add-plant", methods=["POST"])
def tp_add_plant():
    plant_data = {
        "plant_code":     request.form.get("plant_code", ""),
        "exco_location":  request.form.get("exco_location", ""),
        "plant_name":     request.form.get("plant_name", ""),
        "business_head":  request.form.get("business_head", ""),
        "plant_manager":  request.form.get("plant_manager", ""),
        "mixer_theo_cap": request.form.get("mixer_theo_cap", 0),
    }
    try:
        database.add_tp_plant(**plant_data)
        flash("✅ Plant added successfully.", "success")
        res = google_sheets.push_tp_plant_add(plant_data)
        if res["ok"]:
            flash(f"☁️ {res['message']}", "info")
        else:
            flash(f"⚠️ Saved locally but Google Sheet not updated: {res['message']}", "warning")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("tp_data_uploader") + "?open_modal=plant&tab=add")


@app.route("/tp/action/update-plant/<code>", methods=["POST"])
def tp_update_plant(code):
    plant_data = {
        "plant_code":     code,
        "exco_location":  request.form.get("exco_location", ""),
        "plant_name":     request.form.get("plant_name", ""),
        "business_head":  request.form.get("business_head", ""),
        "plant_manager":  request.form.get("plant_manager", ""),
        "mixer_theo_cap": request.form.get("mixer_theo_cap", 0),
    }
    try:
        database.update_tp_plant(**plant_data)
        flash("✅ Plant updated successfully.", "success")
        res = google_sheets.push_tp_plant_update(plant_data)
        if res["ok"]:
            flash(f"☁️ {res['message']}", "info")
        else:
            flash(f"⚠️ Saved locally but Google Sheet not updated: {res['message']}", "warning")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("tp_data_uploader") + "?open_modal=plant&tab=edit")


@app.route("/tp/action/delete-plant/<code>", methods=["POST"])
def tp_delete_plant(code):
    try:
        database.delete_tp_plant(code)
        flash(f"✅ Plant '{code}' deleted.", "success")
        res = google_sheets.push_tp_plant_delete(code)
        if res["ok"]:
            flash(f"☁️ {res['message']}", "info")
        else:
            flash(f"⚠️ Deleted locally but Google Sheet not updated: {res['message']}", "warning")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("tp_data_uploader") + "?open_modal=plant&tab=delete")


@app.route("/tp/action/fetch-oracle", methods=["POST"])
def tp_fetch_oracle():
    from_date = request.form.get("from_date", "")
    to_date   = request.form.get("to_date", "")
    if not from_date or not to_date:
        flash("Please select both From and To dates.", "error")
        return redirect(url_for("tp_data_uploader"))
    try:
        raw_df, ora_warnings = oracle_connector.fetch_tp_data(from_date, to_date)
        parsed, skip_log     = tp_calculator.parse_oracle_df(raw_df)
        oracle_connector.save_tp_oracle_data(raw_df, from_date, to_date, parsed, replace=True)
        database.purge_old_oracle_data()   # keep only previous + current month
        _mss("tp", "skip_log", skip_log)
        _mss("tp", "ora_from", from_date)
        _mss("tp", "ora_to",   to_date)
        for w in ora_warnings:
            flash(w, "warning")
        flash(f"✅ {len(parsed)} rows fetched & saved ({len(skip_log)} skipped).", "success")
    except Exception as exc:
        flash(f"Oracle error: {exc}", "error")
    return redirect(url_for("tp_data_uploader"))


@app.route("/tp/action/sync-sheets", methods=["POST"])
def tp_sync_sheets():
    sheet_id    = (request.form.get("sheet_id", "").strip()
                   or database.get_module_setting("tp", "gsheet_id",
                                                   database.get_setting("gsheet_id", "")))
    worksheet   = request.form.get("worksheet", "Plant Data for TP").strip()
    result      = google_sheets.sync_tp_plant_data(sheet_id, worksheet)
    if result["error"]:
        flash(f"Sync failed: {result['error']}", "error")
    else:
        # Remember the TP sheet id/worksheet so future syncs reuse them.
        database.set_module_setting("tp", "gsheet_id", google_sheets.extract_sheet_id(sheet_id))
        database.set_module_setting("tp", "gsheet_worksheet", worksheet)
        flash(f"✅ {result['rows_synced']} plant rows synced from Google Sheets ({result['mode']} mode).", "success")
    return redirect(url_for("tp_data_uploader"))


# ── Calculate ─────────────────────────────────────────────────────────────────
@app.route("/tp/calculate", methods=["GET", "POST"])
def tp_calculate():
    plant_rows    = _ms("tp", "plant_rows", [])
    location_rows = _ms("tp", "location_rows", [])
    warnings      = _ms("tp", "calc_warnings", [])
    ran           = _ms("tp", "calc_ran", False)

    if request.method == "POST":
        from_s = request.form.get("from_date", "")
        to_s   = request.form.get("to_date", "")
        try:
            fd = _date.fromisoformat(from_s)
            td = _date.fromisoformat(to_s)
        except (ValueError, TypeError):
            flash("Please select valid From and To dates.", "error")
            return redirect(url_for("tp_calculate"))
        if fd > td:
            flash("'From Date' must be on or before 'To Date'.", "error")
            return redirect(url_for("tp_calculate"))

        month, year = fd.month, fd.year
        plant_rows, location_rows, warnings = tp_calculator.run_tp_calculation(
            month, year, from_date=str(fd), to_date=str(td))
        _mss("tp", "plant_rows",    plant_rows)
        _mss("tp", "location_rows", location_rows)
        _mss("tp", "calc_warnings", warnings)
        _mss("tp", "calc_month",    month)
        _mss("tp", "calc_year",     year)
        _mss("tp", "calc_from",     str(fd))
        _mss("tp", "calc_to",       str(td))
        _mss("tp", "calc_ran",      True)
        ran = True

        if plant_rows:
            tp_calculator.save_tp_results(plant_rows, month, year)
            flash(f"✅ Calculated throughput for {fd} → {td} — {len(plant_rows)} plants, "
                  f"{len([l for l in location_rows if not l['is_pan_india']])} locations.", "success")
        else:
            flash("No results produced — check data and warnings.", "warning")
        return redirect(url_for("tp_calculate"))

    today = _date.today()
    ctx = _tp_ctx()
    ctx["active_page"] = "calculate"
    return render_template("tp_calculate.html",
                           plant_rows=plant_rows, location_rows=location_rows,
                           warnings=warnings, ran=ran,
                           default_from=_ms("tp", "calc_from", str(today.replace(day=1))),
                           default_to=_ms("tp", "calc_to", str(today)),
                           **ctx)


# ── Reports ───────────────────────────────────────────────────────────────────
def _tp_band(pct):
    """Throughput colour band for a percentage."""
    if pct < 60:
        return "red"
    if pct < 75:
        return "yellow"
    return "green"


def _apply_tp_filters(plant_rows, excos, bheads, plants, band, search):
    """Filter plant rows by Exco Location, Business Head, Plant, throughput band
    and a free-text search over plant code / plant name."""
    out = []
    s = (search or "").strip().lower()
    for r in plant_rows:
        if excos and r.get("exco_location") not in excos:
            continue
        if bheads and r.get("business_head") not in bheads:
            continue
        if plants and r.get("plant_name") not in plants:
            continue
        if band and band != "All" and _tp_band(r.get("throughput_pct", 0)) != band:
            continue
        if s and s not in str(r.get("lookup_code", "")).lower() \
             and s not in str(r.get("plant_name", "")).lower():
            continue
        out.append(r)
    return out


@app.route("/tp/reports")
def tp_reports():
    today = _date.today()

    # Date range — default to the last calculation's range, else this month.
    from_s = request.args.get("from_date",
                              _ms("tp", "calc_from", str(today.replace(day=1))))
    to_s   = request.args.get("to_date",
                              _ms("tp", "calc_to", str(today)))
    try:
        fd = _date.fromisoformat(from_s)
        td = _date.fromisoformat(to_s)
    except ValueError:
        fd, td = today.replace(day=1), today

    ora_note = None

    # Re-run the calculation for the chosen range when the user loads a report,
    # otherwise fall back to the rows from the last Calculate run.
    if "from_date" in request.args:
        month, year = fd.month, fd.year

        # Auto-fetch Oracle data for this range in-memory (no DB write, no purge)
        # so historical date ranges are not blocked by the 2-month retention window.
        ora_df_live = None
        if oracle_connector.is_configured():
            try:
                raw_df, ora_warns = oracle_connector.fetch_tp_data(str(fd), str(td))
                if not raw_df.empty:
                    parsed, skip_log = tp_calculator.parse_oracle_df(raw_df)
                    _mss("tp", "skip_log", skip_log)
                    import pandas as _pd
                    ora_df_live = _pd.DataFrame(parsed)
                    ora_note = f"Oracle: {len(raw_df):,} rows loaded for {fd} → {td}."
                else:
                    ora_note = "Oracle returned no rows for this range."
                    _mss("tp", "skip_log", [])
            except Exception as exc:
                ora_note = f"Oracle fetch failed: {exc}"

        all_plants, all_locs, calc_warns = tp_calculator.run_tp_calculation(
            month, year, from_date=str(fd), to_date=str(td), ora_df=ora_df_live)
        _mss("tp", "calc_warnings", calc_warns)
        _mss("tp", "plant_rows", all_plants)
        _mss("tp", "location_rows", all_locs)
        _mss("tp", "calc_from", str(fd))
        _mss("tp", "calc_to", str(td))
        _mss("tp", "calc_month", month)
        _mss("tp", "calc_year", year)
    else:
        all_plants = _ms("tp", "plant_rows", [])
        all_locs   = _ms("tp", "location_rows", [])

    # Filter inputs
    excos  = request.args.getlist("exco")
    bheads = request.args.getlist("bhead")
    plants = request.args.getlist("plant")
    band   = request.args.get("band", "All")
    search = request.args.get("search", "")

    # Distinct filter options (from the full, unfiltered set)
    unique_excos  = sorted({r.get("exco_location", "") for r in all_plants if r.get("exco_location")})
    unique_bheads = sorted({r.get("business_head", "") for r in all_plants if r.get("business_head")})
    unique_plants = sorted({r.get("plant_name", "")    for r in all_plants if r.get("plant_name")})

    # Apply filters → plant rows, then rebuild location rows from the filtered set
    plant_rows = _apply_tp_filters(all_plants, excos, bheads, plants, band, search)
    location_rows = tp_calculator.build_location_rows(plant_rows, fd.month, fd.year)

    # Keep a filtered snapshot for downloads / email
    _mss("tp", "report_plant_rows", plant_rows)
    _mss("tp", "report_location_rows", location_rows)

    smtp        = email_helper.get_smtp_config()
    email_ready = email_helper.is_configured()

    # TP-only email log (filter shared table by RDC_TP_ filename prefix)
    _elog = database.read_table("email_log", order_by="id DESC")
    if not _elog.empty and "report_file_name" in _elog.columns:
        _elog = _elog[_elog["report_file_name"].str.startswith("RDC_TP_", na=False)]
    tp_email_log = _records(_elog.drop(columns=["id"], errors="ignore").head(20)) \
        if not _elog.empty else []

    if (fd.year, fd.month) == (td.year, td.month):
        import calendar
        month_label = f"{calendar.month_name[fd.month]} {fd.year}"
    else:
        month_label = f"{fd:%d %b %Y} → {td:%d %b %Y}"

    # TP-specific email defaults (set in TP Settings → Email)
    default_to      = database.get_module_setting("tp", "email_default_to", "") or smtp.get("default_to", "")
    default_cc      = database.get_module_setting("tp", "email_default_cc", "") or smtp.get("default_cc", "")
    default_subject = (database.get_module_setting("tp", "email_default_subject", "")
                       or f"RDC-TP Plant Throughput Report — {month_label}")
    default_body    = (database.get_module_setting("tp", "email_default_body", "")
                       or f"Dear Team,\n\nPlease find attached the Plant Throughput "
                          f"Report for {month_label}.\n\nRegards,\nRDC Operations")

    ctx = _tp_ctx()
    ctx["active_page"] = "reports"
    return render_template("tp_reports.html",
                           plant_rows=plant_rows, location_rows=location_rows,
                           total_plants=len(all_plants),
                           from_date=str(fd), to_date=str(td),
                           unique_excos=unique_excos, unique_bheads=unique_bheads,
                           unique_plants=unique_plants,
                           excos=excos, bheads=bheads, plants=plants,
                           band=band, search=search,
                           month_label=month_label, email_cfg=smtp,
                           email_ready=email_ready,
                           default_to=default_to, default_cc=default_cc,
                           default_subject=default_subject,
                           default_body=default_body,
                           tp_email_log=tp_email_log,
                           ora_note=ora_note, **ctx)


def _tp_mon_tag(month, year):
    import calendar as _cal
    return f"{_cal.month_abbr[month]}'{str(year)[2:]}"


def _tp_build_excel(plant_rows, location_rows, month, year):
    """Return color-coded, wrapped-header Excel bytes for the TP report."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    mon_tag = _tp_mon_tag(month, year)

    RED_FILL    = PatternFill("solid", fgColor="FFD5D5")
    YEL_FILL    = PatternFill("solid", fgColor="FFF3CC")
    GRN_FILL    = PatternFill("solid", fgColor="D5F0D5")
    HDR_FILL    = PatternFill("solid", fgColor="1F4E79")
    TITLE_FILL  = PatternFill("solid", fgColor="0D2B52")
    HDR_FONT    = Font(bold=True, color="FFFFFF", size=10)
    TITLE_FONT  = Font(bold=True, color="FFFFFF", size=11)
    WRAP        = Alignment(wrap_text=True, vertical="center", horizontal="center")
    CTR         = Alignment(vertical="center", horizontal="center")
    LEFT        = Alignment(vertical="center", horizontal="left")
    thin        = Side(style="thin", color="CCCCCC")
    BDR         = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _fill(pct):
        return RED_FILL if pct < 60 else (YEL_FILL if pct < 75 else GRN_FILL)

    wb = Workbook()

    # ── Location sheet ──────────────────────────────────────────────────────────
    ws_l = wb.active
    ws_l.title = "Location Throughput"
    loc_heads = ["Exco Location", "Plants", "Total Qty", "Avg Throughput %"]
    ws_l.merge_cells(f"A1:{get_column_letter(len(loc_heads))}1")
    t = ws_l.cell(1, 1, f"Location wise Throughput - {mon_tag}")
    t.fill = TITLE_FILL; t.font = TITLE_FONT; t.alignment = CTR
    ws_l.row_dimensions[1].height = 22
    for ci, h in enumerate(loc_heads, 1):
        c = ws_l.cell(2, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    ws_l.row_dimensions[2].height = 32
    for ri, row in enumerate(location_rows, 3):
        pct = float(row.get("avg_throughput_pct", 0))
        f   = _fill(pct)
        vals = [row.get("exco_location",""), row.get("plant_count",0),
                round(float(row.get("total_quantity",0)), 1), f"{round(pct)}%"]
        for ci, v in enumerate(vals, 1):
            c = ws_l.cell(ri, ci, v)
            c.fill = f; c.border = BDR
            c.alignment = LEFT if ci == 1 else CTR
    for ci, w in enumerate([24, 9, 12, 18], 1):
        ws_l.column_dimensions[get_column_letter(ci)].width = w

    # ── Plant sheet ─────────────────────────────────────────────────────────────
    ws_p = wb.create_sheet("Plant Throughput")
    plant_heads = ["Plant Code","Plant","Exco Location","Business Head",
                   "Plant Manager","Mixer Cap","Total Qty","Time (min)","Throughput %","Batches"]
    ws_p.merge_cells(f"A1:{get_column_letter(len(plant_heads))}1")
    t2 = ws_p.cell(1, 1, f"Plant Throughput report - {mon_tag}")
    t2.fill = TITLE_FILL; t2.font = TITLE_FONT; t2.alignment = CTR
    ws_p.row_dimensions[1].height = 22
    for ci, h in enumerate(plant_heads, 1):
        c = ws_p.cell(2, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    ws_p.row_dimensions[2].height = 40
    for ri, row in enumerate(plant_rows, 3):
        pct = float(row.get("throughput_pct", 0))
        f   = _fill(pct)
        vals = [row.get("lookup_code",""), row.get("plant_name",""),
                row.get("exco_location",""), row.get("business_head",""),
                row.get("plant_manager",""), row.get("mixer_theo_cap",""),
                round(float(row.get("total_quantity",0)), 1),
                round(float(row.get("total_time_min",0)), 1),
                f"{round(pct)}%", row.get("batch_count",0)]
        for ci, v in enumerate(vals, 1):
            c = ws_p.cell(ri, ci, v)
            c.fill = f; c.border = BDR
            c.alignment = CTR if ci > 2 else LEFT
    for ci, w in enumerate([12,22,18,20,20,10,11,11,13,9], 1):
        ws_p.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _tp_build_html_tables(plant_rows, location_rows, month, year):
    """Return HTML string with two color-coded tables for TP report email."""
    mon_tag = _tp_mon_tag(month, year)

    def _bg(pct):
        if pct < 60:   return "#FFB3B3"
        if pct < 75:   return "#FFE066"
        return "#92D492"

    def _fg(bg):
        """Dark complementary text color for each background."""
        return {"#FFB3B3": "#7B1F1F", "#FFE066": "#5C4200",
                "#92D492": "#1A5C1A", "#D9D9D9": "#222222"}.get(bg, "#19263A")

    BORDER = "2px solid #9A9A9A"
    HDR_BG = "#D9D9D9"
    FONT   = "font-family:Arial,sans-serif;font-size:10px;"

    TTL = (f'style="{FONT}background:#0A2540;color:#fff;font-weight:bold;'
           f'padding:6px 8px;font-size:11px;border:{BORDER};text-align:left"')

    def _th(w):
        return (f'style="{FONT}background:#0A2540;color:#fff;font-weight:bold;'
                f'padding:3px 5px;border:{BORDER};text-align:center;'
                f'white-space:normal;word-break:break-word;line-height:1.2;width:{w}"')

    # ── Location table — plain rows, PAN India row colored ───────────────────
    L_COLS = [("Sr. no.","3%"),("Exco Location","27%"),("Plants","8%"),
              ("Total Qty","16%"),("Time (min)","16%"),("Avg TP %","10%")]
    ths = "".join(f'<th {_th(w)}>{h}</th>' for h, w in L_COLS)

    loc_body = ""
    for i, r in enumerate(location_rows, 1):
        pct = float(r.get("avg_throughput_pct", 0))
        pan = bool(r.get("is_pan_india"))
        if pan:
            pct_bg  = _bg(pct)
            PAN_BASE = (f'{FONT}background:{HDR_BG};color:{_fg(HDR_BG)};font-weight:bold;'
                        f'padding:4px 5px;border:{BORDER};border-top:2px solid #555;'
                        f'vertical-align:middle')
            loc_body += (
                f'<tr>'
                f'<td style="{PAN_BASE};text-align:center">—</td>'
                f'<td style="{PAN_BASE};text-align:left">&#127988; PAN India</td>'
                f'<td style="{PAN_BASE};text-align:center">{r.get("plant_count",0)}</td>'
                f'<td style="{PAN_BASE};text-align:center">{round(float(r.get("total_quantity",0)),1)}</td>'
                f'<td style="{PAN_BASE};text-align:center">{round(float(r.get("total_time_min",0)),1)}</td>'
                f'<td style="{FONT}background:{pct_bg};color:{_fg(pct_bg)};font-weight:bold;'
                f'padding:4px 5px;border:{BORDER};border-top:2px solid #555;'
                f'text-align:center;vertical-align:middle">{round(pct)}%</td>'
                f'</tr>'
            )
        else:
            PLAIN = (f'{FONT}background:#ffffff;color:#19263A;'
                     f'padding:3px 5px;border:{BORDER};vertical-align:middle')
            loc_body += (
                f'<tr>'
                f'<td style="{PLAIN};text-align:center">{i}</td>'
                f'<td style="{PLAIN};text-align:left">{r.get("exco_location","")}</td>'
                f'<td style="{PLAIN};text-align:center">{r.get("plant_count",0)}</td>'
                f'<td style="{PLAIN};text-align:center">{round(float(r.get("total_quantity",0)),1)}</td>'
                f'<td style="{PLAIN};text-align:center">{round(float(r.get("total_time_min",0)),1)}</td>'
                f'<td style="{PLAIN};text-align:center;font-weight:bold">{round(pct)}%</td>'
                f'</tr>'
            )

    colspan_l = len(L_COLS)
    loc_html = (
        f'<table cellpadding="0" cellspacing="0" '
        f'style="border-collapse:collapse;width:100%;max-width:560px;'
        f'margin:14px 0 10px;table-layout:fixed">'
        f'<tr><td colspan="{colspan_l}" {TTL}>Location wise Throughput - {mon_tag}</td></tr>'
        f'<tr>{ths}</tr>{loc_body}</table>'
    )

    # ── Plant table — full list, color-coded by TP % ─────────────────────────
    P_COLS = [("Sr. no.","3%"),("Plant","23%"),("Exco Location","11%"),
              ("Business Head","11%"),("Plant Manager","11%"),
              ("Mixer Cap","7%"),("Total Qty","10%"),("Time (min)","10%"),("TP %","8%")]
    ths2 = "".join(f'<th {_th(w)}>{h}</th>' for h, w in P_COLS)

    plant_body = ""
    for i, r in enumerate(plant_rows, 1):
        pct = float(r.get("throughput_pct", 0))
        bg  = _bg(pct)
        fg  = _fg(bg)
        TD  = (f'{FONT}background:{bg};color:{fg};'
               f'padding:3px 5px;border:{BORDER};vertical-align:middle')
        plant_body += (
            f'<tr>'
            f'<td style="{TD};text-align:center">{i}</td>'
            f'<td style="{TD};text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{r.get("plant_name","")}</td>'
            f'<td style="{TD};text-align:left">{r.get("exco_location","")}</td>'
            f'<td style="{TD};text-align:left">{r.get("business_head","")}</td>'
            f'<td style="{TD};text-align:left">{r.get("plant_manager","")}</td>'
            f'<td style="{TD};text-align:center">{r.get("mixer_theo_cap","")}</td>'
            f'<td style="{TD};text-align:center">{round(float(r.get("total_quantity",0)),1)}</td>'
            f'<td style="{TD};text-align:center">{round(float(r.get("total_time_min",0)),1)}</td>'
            f'<td style="{TD};text-align:center;font-weight:bold">{round(pct)}%</td>'
            f'</tr>'
        )

    colspan_p = len(P_COLS)
    plant_html = (
        f'<table cellpadding="0" cellspacing="0" '
        f'style="border-collapse:collapse;width:100%;max-width:860px;'
        f'margin:10px 0 20px;table-layout:fixed">'
        f'<tr><td colspan="{colspan_p}" {TTL}>Plant Throughput report - {mon_tag}</td></tr>'
        f'<tr>{ths2}</tr>{plant_body}</table>'
    )

    return loc_html + plant_html


@app.route("/tp/download-excel")
def tp_download_excel():
    plant_rows    = _ms("tp", "report_plant_rows", _ms("tp", "plant_rows", []))
    location_rows = _ms("tp", "report_location_rows", _ms("tp", "location_rows", []))
    month = _ms("tp", "calc_month", _date.today().month)
    year  = _ms("tp", "calc_year",  _date.today().year)
    if not plant_rows:
        flash("No data to download — run Calculate first.", "warning")
        return redirect(url_for("tp_reports"))
    excel_bytes = _tp_build_excel(plant_rows, location_rows, month, year)
    fname = f"RDC_TP_{year}_{month:02d}.xlsx"
    return send_file(io.BytesIO(excel_bytes), as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/tp/download-csv")
def tp_download_csv():
    plant_rows = _ms("tp", "report_plant_rows", _ms("tp", "plant_rows", []))
    month = _ms("tp", "calc_month", _date.today().month)
    year  = _ms("tp", "calc_year",  _date.today().year)
    if not plant_rows:
        flash("No data to download — run Calculate first.", "warning")
        return redirect(url_for("tp_reports"))
    df = pd.DataFrame(plant_rows)
    buf = io.BytesIO(df.to_csv(index=False).encode("utf-8"))
    return send_file(buf, as_attachment=True,
                     download_name=f"RDC_TP_{year}_{month:02d}.csv",
                     mimetype="text/csv")


@app.route("/tp/action/send-email", methods=["POST"])
def tp_send_email():
    plant_rows    = _ms("tp", "report_plant_rows", _ms("tp", "plant_rows", []))
    location_rows = _ms("tp", "report_location_rows", _ms("tp", "location_rows", []))
    month = _ms("tp", "calc_month", _date.today().month)
    year  = _ms("tp", "calc_year",  _date.today().year)
    to_addr = request.form.get("to", "").strip()
    cc_addr = request.form.get("cc", "").strip()
    subject = request.form.get("subject", "").strip()
    body    = request.form.get("body", "").strip()
    if not plant_rows:
        flash("No results to email — run Calculate first.", "warning")
        return redirect(url_for("tp_reports"))
    if not to_addr:
        flash("Please enter at least one To email address.", "error")
        return redirect(url_for("tp_reports"))

    import calendar as _cal
    month_name = _cal.month_name[month]
    if not subject:
        subject = f"RDC-TP Plant Throughput Report — {month_name} {year}"
    if not body:
        body = f"Dear Team,\n\nPlease find the Plant Throughput Report for {month_name} {year} below."

    # Build color-coded Excel attachment
    excel_bytes = _tp_build_excel(plant_rows, location_rows, month, year)
    fname = f"RDC_TP_{year}_{month:02d}.xlsx"

    # HTML tables in body — no inline images
    tables_html = _tp_build_html_tables(plant_rows, location_rows, month, year)
    html_body   = email_helper.wrap_html_body(body, tables_html)

    result = email_helper.send_report_email(
        to_emails=to_addr, cc_emails=cc_addr,
        subject=subject, body=body,
        attachment_bytes=excel_bytes, attachment_name=fname,
        html_body=html_body,
    )
    if result.get("success"):
        flash(f"✅ Report emailed to {to_addr}.", "success")
    else:
        flash(f"Email failed: {result.get('error')}", "error")
    return redirect(url_for("tp_reports"))


# ── Validation ────────────────────────────────────────────────────────────────
@app.route("/tp/validation")
def tp_validation():
    skip_log = _ms("tp", "skip_log", [])
    calc_warnings = _ms("tp", "calc_warnings", [])
    ctx = _tp_ctx()
    ctx["active_page"] = "validation"
    return render_template("tp_validation.html",
                           skip_log=skip_log, calc_warnings=calc_warnings, **ctx)


# ── Settings ──────────────────────────────────────────────────────────────────
@app.route("/tp/settings", methods=["GET"])
def tp_settings():
    sheet_id    = database.get_module_setting("tp", "gsheet_id",
                                              database.get_setting("gsheet_id", ""))
    worksheet   = database.get_module_setting("tp", "gsheet_worksheet", "Plant Data for TP")
    plant_col   = database.get_module_setting("tp", "oracle_plant_col", "PLANTNO")
    batch_col   = database.get_module_setting("tp", "oracle_batch_col", "BATCHCODE")
    time_col    = database.get_module_setting("tp", "oracle_time_col",  "TIMETAKEN")
    smtp        = email_helper.get_smtp_config()
    email_configured = bool(smtp.get("host") and smtp.get("sender"))
    ora_configured   = oracle_connector.is_configured()
    last_sync        = google_sheets.get_tp_last_sync_info()
    # Monthly-email schedule (TP-scoped)
    sched_enabled     = database.get_module_setting("tp", "email_schedule_enabled", "false") == "true"
    sched_time        = database.get_module_setting("tp", "email_schedule_time", "08:00")
    sched_to          = database.get_module_setting("tp", "email_schedule_to", "")
    sched_cc          = database.get_module_setting("tp", "email_schedule_cc", "")
    sched_last_status  = database.get_module_setting("tp", "email_schedule_last_status", "")
    tp_email_to        = database.get_module_setting("tp", "email_default_to", "")
    tp_email_cc        = database.get_module_setting("tp", "email_default_cc", "")
    tp_email_subject   = database.get_module_setting("tp", "email_default_subject", "")
    tp_email_body      = database.get_module_setting("tp", "email_default_body", "")
    ctx = _tp_ctx()
    ctx["active_page"] = "settings"
    return render_template("tp_settings.html",
                           sheet_id=sheet_id, worksheet=worksheet,
                           plant_col=plant_col, batch_col=batch_col, time_col=time_col,
                           smtp=smtp, email_configured=email_configured,
                           ora_configured=ora_configured, last_sync=last_sync,
                           sched_enabled=sched_enabled, sched_time=sched_time,
                           sched_to=sched_to, sched_cc=sched_cc,
                           sched_last_status=sched_last_status,
                           tp_email_to=tp_email_to, tp_email_cc=tp_email_cc,
                           tp_email_subject=tp_email_subject, tp_email_body=tp_email_body,
                           **ctx)


@app.route("/tp/settings/save-oracle-cols", methods=["POST"])
def tp_save_oracle_cols():
    database.set_module_settings_bulk("tp", {
        "oracle_plant_col": request.form.get("plant_col", "PLANTNO").strip(),
        "oracle_batch_col": request.form.get("batch_col", "BATCHCODE").strip(),
        "oracle_time_col":  request.form.get("time_col",  "TIMETAKEN").strip(),
    })
    flash("Oracle column names saved.", "success")
    return redirect(url_for("tp_settings", m="oracle-cols"))


@app.route("/tp/settings/save-worksheet", methods=["POST"])
def tp_save_worksheet():
    worksheet = request.form.get("worksheet", "Plant Data for TP").strip()
    sheet_id  = request.form.get("sheet_id", "").strip()
    database.set_module_setting("tp", "gsheet_worksheet", worksheet)
    if sheet_id:
        # Save to the TP-scoped id so we never disturb RDC-I&D's sheet.
        database.set_module_setting("tp", "gsheet_id", google_sheets.extract_sheet_id(sheet_id))
    flash("Sheet settings saved.", "success")
    return redirect(url_for("tp_settings", m="sheet"))


@app.route("/tp/settings/save-schedule", methods=["POST"])
def tp_save_email_schedule():
    sched_time = request.form.get("sched_time", "08:00").strip()
    database.set_module_settings_bulk("tp", {
        "email_schedule_time": sched_time,
        "email_schedule_to":   request.form.get("sched_to", "").strip(),
        "email_schedule_cc":   request.form.get("sched_cc", "").strip(),
    })
    if _scheduler:
        try:
            h, m = map(int, sched_time.split(":"))
        except Exception:
            h, m = 8, 0
        _scheduler.reschedule_job("tp_monthly_email",
                                  trigger=CronTrigger(day=1, hour=h, minute=m))
    flash("TP schedule settings saved.", "success")
    return redirect(url_for("tp_settings", m="schedule"))


@app.route("/tp/settings/save-email-defaults", methods=["POST"])
def tp_save_email_defaults():
    database.set_module_settings_bulk("tp", {
        "email_default_to":      request.form.get("default_to", "").strip(),
        "email_default_cc":      request.form.get("default_cc", "").strip(),
        "email_default_subject": request.form.get("default_subject", "").strip(),
        "email_default_body":    request.form.get("default_body", "").strip(),
    })
    flash("TP email defaults saved.", "success")
    return redirect(url_for("tp_settings", m="email"))


@app.route("/tp/settings/toggle-schedule", methods=["POST"])
def tp_toggle_email_schedule():
    current = database.get_module_setting("tp", "email_schedule_enabled", "false")
    new_val = "false" if current == "true" else "true"
    database.set_module_setting("tp", "email_schedule_enabled", new_val)
    flash(f"RDC-TP scheduled monthly report {'enabled' if new_val == 'true' else 'disabled'}.", "success")
    return redirect(url_for("tp_settings", m="schedule"))


@app.route("/btrtp")
def page_btrtp():
    return render_template("module_placeholder.html",
                           module_name="RDC-BTRTP",
                           module_desc="Calculates batcher wise Throughput for individual efficiency tracking")

@app.route("/jldc")
def page_jldc():
    return render_template("module_placeholder.html",
                           module_name="RDC-JLDC",
                           module_desc="Calculates LJCB & Loader Diesel Consumption")


@app.route("/data-uploader")
def page_data_uploader():
    _ss("active_page", "data_uploader")
    last_sync = google_sheets.get_last_sync_info()
    has_creds = google_sheets.credentials_exist()
    ora_cfg   = oracle_connector.get_oracle_config()
    ora_ready = oracle_connector.is_configured(ora_cfg)
    auto_sync = database.get_setting("gsheet_auto_sync", "false") == "true"

    m_count = database.get_table_counts().get("master_data", 0)
    master_cols, master_rows = [], []
    if m_count > 0:
        df = database.read_table_limited("master_data", order_by="employee_code", limit=500)
        df = df.drop(columns=["id"], errors="ignore")
        master_cols = df.columns.tolist()
        master_rows = _records(df)

    b_count = database.get_table_counts().get("backend_data", 0)
    b_earliest = b_latest = ""
    backend_preview = []
    if b_count > 0:
        df_e = database.read_table_limited("backend_data", order_by="date", limit=1)
        df_l = database.read_table_limited("backend_data", order_by="date DESC", limit=1)
        b_earliest = df_e["date"].iloc[0] if not df_e.empty else ""
        b_latest   = df_l["date"].iloc[0] if not df_l.empty else ""
        df_p = database.read_table_limited("backend_data", order_by="date", limit=200)
        backend_preview = _records(df_p.drop(columns=["id"], errors="ignore"))

    maint_df = database.read_table("maintenance_cost", order_by="plant_code")
    maint_cols, maint_rows, maint_avg, maint_above = [], [], 0, 0
    if not maint_df.empty:
        df_m2 = maint_df.drop(columns=["id"], errors="ignore")
        maint_cols = df_m2.columns.tolist()
        maint_rows = _records(df_m2)
        maint_avg  = round(maint_df["ytd_maintenance_cost"].mean(), 2)
        maint_above = int((maint_df["ytd_maintenance_cost"] > config.MAINTENANCE_COST_THRESHOLD).sum())

    codes_df = database.read_table_limited("master_data", order_by="employee_code", limit=100000)
    codes = codes_df["employee_code"].astype(str).tolist() if not codes_df.empty else []

    log_df   = database.read_table("master_data_change_log", order_by="id DESC")
    log_rows = _records(log_df.drop(columns=["id"], errors="ignore")) if not log_df.empty else []

    ora_b_preview = []
    if b_count > 0:
        df_ob = database.read_table_limited("backend_data", order_by="date DESC", limit=10)
        ora_b_preview = _records(df_ob.drop(columns=["id"], errors="ignore"))

    # Edit / delete employee (loaded when query params present)
    edit_code = request.args.get("edit_code")
    del_code  = request.args.get("del_code")
    edit_emp  = database.get_employee(edit_code) if edit_code else None
    del_emp   = database.get_employee(del_code)  if del_code  else None

    return render_template("data_uploader.html",
                           last_sync=last_sync, has_creds=has_creds,
                           ora_cfg=ora_cfg, ora_ready=ora_ready, auto_sync=auto_sync,
                           m_count=m_count, master_cols=master_cols, master_rows=master_rows,
                           b_count=b_count, b_earliest=b_earliest, b_latest=b_latest,
                           backend_preview=backend_preview,
                           maint_cols=maint_cols, maint_rows=maint_rows,
                           maint_avg=maint_avg, maint_above=maint_above,
                           maint_count=len(maint_rows),
                           codes=codes, log_rows=log_rows, ora_b_preview=ora_b_preview,
                           edit_emp=edit_emp, del_emp=del_emp,
                           categories=config.CATEGORIES,
                           today=str(_date.today()),
                           today_first=str(_date.today().replace(day=1)))


@app.route("/calculate")
def page_calculate():
    _ss("active_page", "calculate")
    available = calculator.get_available_months()

    default_from = default_to = str(_date.today())
    if available:
        conn = database.get_connection()
        try:
            mm = pd.read_sql_query(
                "SELECT MIN(date) AS mn, MAX(date) AS mx FROM backend_data", conn)
        finally:
            conn.close()
        default_from = str(pd.to_datetime(mm["mn"].iloc[0]).date())
        default_to   = str(pd.to_datetime(mm["mx"].iloc[0]).date())

    last_calc  = calculator.get_last_calculation_info()
    results_df = database.read_table("calculation_results") \
        if database.get_table_counts().get("calculation_results", 0) > 0 \
        else pd.DataFrame()

    results       = _records(results_df)
    unmapped_rows = _records(database.read_table("unmapped_employees")
                             .drop(columns=["id"], errors="ignore"))

    cat_results = {}
    for label, cats in CAT_TABS.items():
        grp = [r for r in results if r.get("category") in cats]
        grp = _sort_rows(grp)
        for r in grp:
            r["_cls"] = _row_cls(r)
        cat_results[label] = grp

    total_emps = len(results)
    elig_count = sum(1 for r in results if r.get("incentive_eligible") == "Yes")
    total_inc  = sum((r.get("incentive_amount") or 0) for r in results)
    total_ded  = sum((r.get("deduction_amount") or 0) for r in results)

    return render_template("calculate.html",
                           available=available,
                           default_from=default_from, default_to=default_to,
                           last_calc=last_calc,
                           results=results if not results_df.empty else None,
                           ran=_s("calc_ran", False),
                           cat_tabs=CAT_TABS, cat_results=cat_results,
                           unmapped=unmapped_rows,
                           result_cols=RESULT_COLS, result_labels=RESULT_LABELS,
                           total_emps=total_emps, elig_count=elig_count,
                           total_inc=total_inc, total_ded=total_ded)


@app.route("/reports")
def page_reports():
    _ss("active_page", "reports")

    available  = calculator.get_available_months()
    ora_live   = oracle_connector.is_configured()
    no_backend = (not available) and (not ora_live)

    from_date_s = request.args.get("from_date", str(_date.today().replace(day=1)))
    to_date_s   = request.args.get("to_date",   str(_date.today()))
    try:
        from_date = _date.fromisoformat(from_date_s)
        to_date   = _date.fromisoformat(to_date_s)
    except ValueError:
        from_date = _date.today().replace(day=1)
        to_date   = _date.today()

    cats   = request.args.getlist("category")
    desigs = request.args.getlist("designation")
    plants = request.args.getlist("plant")
    elig   = request.args.get("elig",    "All")
    outcome = request.args.get("outcome", "All")
    search  = request.args.get("search",  "")

    results = None
    unmapped = []
    error_msg = ""
    ora_note  = ""
    unique_cats = unique_desigs = unique_plants = []
    total_rows = 0

    has_params = "from_date" in request.args

    # Always pre-populate filter dropdowns from the persisted calculation_results table
    # so filter selects are usable before the user clicks "Load Report"
    _cr_all = database.read_table("calculation_results")
    if not _cr_all.empty:
        unique_cats   = sorted(_cr_all["category"].dropna().unique().tolist())   if "category"    in _cr_all.columns else []
        unique_desigs = sorted(_cr_all["designation"].dropna().unique().tolist()) if "designation" in _cr_all.columns else []
        unique_plants = sorted(_cr_all["plant"].dropna().unique().tolist())       if "plant"       in _cr_all.columns else []

    if has_params and not no_backend:
        range_key = f"{from_date}|{to_date}"
        if _s("rpt_range_key") != range_key:
            # Optionally fetch Oracle data first
            if ora_live:
                try:
                    df_ora, _ = oracle_connector.fetch_backend_data(from_date, to_date)
                    if not df_ora.empty:
                        oracle_connector.save_oracle_backend_data(
                            df_ora, from_date, to_date, replace=True)
                        ora_note = f"Oracle: {len(df_ora):,} rows loaded."
                    else:
                        ora_note = "Oracle returned no rows for this range."
                except Exception as exc:
                    ora_note = f"Oracle fetch failed: {exc}"

            res = calculator.run_calculation(
                month=from_date.month, year=from_date.year,
                start_date=str(from_date), end_date=str(to_date),
                persist=False,
            )
            if res["error"]:
                error_msg = res["error"]
                _ss("rpt_all", [])
                _ss("rpt_unmapped", [])
            else:
                _ss("rpt_all", res.get("results_rows", []))
                _ss("rpt_unmapped", res.get("unmapped_rows", []))
            _ss("rpt_range_key", range_key)
            _ss("rpt_ora_note", ora_note)

        all_rows = _s("rpt_all", [])
        unmapped  = _s("rpt_unmapped", [])
        ora_note  = _s("rpt_ora_note", "")
        total_rows = len(all_rows)

        unique_cats   = sorted({r.get("category", "")    for r in all_rows if r.get("category")})
        unique_desigs = sorted({r.get("designation", "") for r in all_rows if r.get("designation")})
        unique_plants = sorted({r.get("plant", "")       for r in all_rows if r.get("plant")})

        filtered = _apply_filters(all_rows, cats, desigs, plants, elig, outcome, search)
        filtered = _sort_rows(filtered)

        # Plant-wise totals (from full unfiltered set so plant sum is always complete)
        plant_inc_map: dict = {}
        plant_ded_map: dict = {}
        for r in all_rows:
            p = r.get("plant", "")
            plant_inc_map[p] = plant_inc_map.get(p, 0.0) + (r.get("incentive_amount") or 0)
            plant_ded_map[p] = plant_ded_map.get(p, 0.0) + (r.get("deduction_amount") or 0)

        for r in filtered:
            r["_cls"] = _row_cls(r)
            r["plant_total_incentive"] = plant_inc_map.get(r.get("plant", ""), 0.0)
            r["plant_total_deduction"] = plant_ded_map.get(r.get("plant", ""), 0.0)
        results = filtered

        # Build applied-filters string
        parts = []
        if cats:   parts.append(f"Category: {', '.join(cats)}")
        if desigs: parts.append(f"Designation: {', '.join(desigs)}")
        if plants: parts.append(f"Plant: {', '.join(plants)}")
        if elig != "All":     parts.append(f"Eligibility: {elig}")
        if outcome != "All":  parts.append(f"Outcome: {outcome}")
        if search.strip():    parts.append(f"Search: {search.strip()}")
        applied_filters = " | ".join(parts) if parts else "None"

        # Store snapshot for downloads / email
        _ss("rpt_filtered",        filtered)
        _ss("rpt_unmapped_snap",   unmapped)
        _ss("rpt_from",            str(from_date))
        _ss("rpt_to",              str(to_date))
        _ss("rpt_applied_filters", applied_filters)

    # Month label for email
    if (from_date.year, from_date.month) == (to_date.year, to_date.month):
        month_label = from_date.strftime("%B %Y")
    else:
        month_label = f"{from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"

    email_cfg    = email_helper.get_smtp_config()
    email_ready  = bool(email_cfg["host"] and email_cfg["sender"] and email_cfg["password"])
    default_subj = email_helper.compose_report_subject(month_label)
    default_body = email_helper.compose_report_body(month_label)

    elog_df    = database.read_table("email_log", order_by="id DESC")
    email_log  = _records(elog_df.drop(columns=["id"], errors="ignore").head(20)) \
        if not elog_df.empty else []

    total_inc  = sum((r.get("incentive_amount") or 0) for r in (results or []))
    total_ded  = sum((r.get("deduction_amount") or 0) for r in (results or []))
    elig_count = sum(1 for r in (results or []) if r.get("incentive_eligible") == "Yes")

    # Group filtered rows by CAT_TABS for the tab UI (mirrors calculate page)
    cat_results = {}
    for label, cats in CAT_TABS.items():
        cat_results[label] = [r for r in (results or []) if r.get("category") in cats]

    return render_template("reports.html",
                           from_date=str(from_date), to_date=str(to_date),
                           results=results,
                           cat_tabs=CAT_TABS, cat_results=cat_results,
                           unmapped=unmapped,
                           total_rows=total_rows,
                           result_cols=REPORT_COLS, result_labels=REPORT_LABELS,
                           unique_cats=unique_cats, unique_desigs=unique_desigs,
                           unique_plants=unique_plants,
                           cats=cats, desigs=desigs, plants=plants,
                           elig=elig, outcome=outcome, search=search,
                           total_inc=total_inc, total_ded=total_ded,
                           elig_count=elig_count,
                           error_msg=error_msg, ora_note=ora_note,
                           ora_live=ora_live, no_backend=no_backend,
                           month_label=month_label,
                           email_cfg=email_cfg, email_ready=email_ready,
                           default_subject=default_subj, default_body=default_body,
                           email_log=email_log)


@app.route("/validation")
def page_validation():
    _ss("active_page", "validation")
    last    = validations.get_last_validation_info()
    counts  = database.get_table_counts()
    total_errors = counts.get("validation_errors", 0)

    summary_rows = summary_cols = []
    master_n = backend_n = maint_n = 0
    errors   = []
    sources  = []
    etypes   = []
    sel_src  = request.args.get("src", "")
    sel_etype = request.args.get("etype", "")

    if total_errors > 0:
        summary_df = validations.get_validation_summary()
        if not summary_df.empty:
            summary_cols = summary_df.columns.tolist()
            summary_rows = _records(summary_df)
            master_n  = int(summary_df[summary_df["Source"] == "master_data"]["Count"].sum())
            backend_n = int(summary_df[summary_df["Source"] == "backend_data"]["Count"].sum())
            maint_n   = int(summary_df[summary_df["Source"] == "maintenance_cost"]["Count"].sum())

        all_err_df = database.read_table("validation_errors")
        sources = sorted(all_err_df["source"].unique().tolist()) if not all_err_df.empty else []
        etypes  = sorted(all_err_df["error_type"].unique().tolist()) if not all_err_df.empty else []

        filtered_err = all_err_df.copy()
        if sel_src:   filtered_err = filtered_err[filtered_err["source"] == sel_src]
        if sel_etype: filtered_err = filtered_err[filtered_err["error_type"] == sel_etype]
        errors = _records(filtered_err.drop(columns=["id"], errors="ignore"))

    return render_template("validation.html",
                           last=last, total_errors=total_errors,
                           summary_rows=summary_rows, summary_cols=summary_cols,
                           master_n=master_n, backend_n=backend_n, maint_n=maint_n,
                           errors=errors, sources=sources, etypes=etypes,
                           sel_src=sel_src, sel_etype=sel_etype)


@app.route("/settings")
def page_settings():
    _ss("active_page", "settings")
    smtp       = email_helper.get_smtp_config()
    email_configured = bool(smtp["host"] and smtp["sender"] and smtp["password"])
    ora        = oracle_connector.get_oracle_config()
    ora_configured = oracle_connector.is_configured(ora)

    last_cache_clear   = database.get_setting("last_cache_clear_date", "")
    cache_cleared_today = last_cache_clear == str(_date.today())

    sched_enabled     = database.get_setting("email_schedule_enabled", "false") == "true"
    sched_time        = database.get_setting("email_schedule_time", "08:00")
    sched_to          = database.get_setting("email_schedule_to", "")
    sched_cc          = database.get_setting("email_schedule_cc", "")
    sched_last_status = database.get_setting("email_schedule_last_status", "")

    return render_template("settings.html",
                           smtp=smtp, email_configured=email_configured,
                           ora=ora, ora_configured=ora_configured,
                           db_size_mb=cache_helpers.get_db_size_mb(),
                           bg_auto=bg_auto(), bg_animate=bg_animate(), bg_theme=bg_theme(),
                           last_cache_clear=last_cache_clear,
                           cache_cleared_today=cache_cleared_today,
                           sched_enabled=sched_enabled, sched_time=sched_time,
                           sched_to=sched_to, sched_cc=sched_cc,
                           sched_last_status=sched_last_status)


# ── ACTIONS ───────────────────────────────────────────────────────────────────

@app.route("/action/sync-gsheet", methods=["POST"])
def sync_gsheet():
    sheet_id  = request.form.get("sheet_id", "").strip()
    worksheet = request.form.get("worksheet", "Sheet1").strip()
    if not sheet_id:
        flash("Please enter a Google Sheet ID.", "error")
        return redirect(url_for("page_data_uploader"))
    try:
        clean_id = google_sheets.extract_sheet_id(sheet_id)
        df_sync  = google_sheets.fetch_master_data(clean_id, worksheet)
        now = _dt.now().isoformat(timespec="seconds")
        rows = [{"employee_code": str(r["Employee Code"]),
                 "employee_name": str(r["Employee Name"]),
                 "designation":   str(r["Designation"]),
                 "category":      str(r["Category"]),
                 "plant":         str(r["Plant"]),
                 "plant_code":    str(r["Plant Code"]),
                 "updated_at":    now}
                for _, r in df_sync.iterrows()]
        inserted = database.replace_table_rows("master_data", rows)
        database.set_settings_bulk({"gsheet_id": clean_id, "gsheet_worksheet": worksheet,
                                    "gsheet_last_sync": now, "gsheet_last_count": str(inserted)})
        flash(f"Synced {inserted:,} employees from Google Sheets.", "success")
    except Exception as exc:
        flash(f"Sync failed: {exc}", "error")
    return redirect(url_for("page_data_uploader"))


@app.route("/action/toggle-auto-sync", methods=["POST"])
def toggle_auto_sync():
    enabled = "enabled" in request.form
    database.set_setting("gsheet_auto_sync", "true" if enabled else "false")
    return redirect(url_for("page_data_uploader"))


@app.route("/action/add-employee", methods=["POST"])
def add_employee():
    emp = {
        "employee_code": request.form["code"].strip(),
        "employee_name": request.form.get("name", "").strip(),
        "designation":   request.form.get("designation", "").strip(),
        "category":      request.form.get("category", ""),
        "plant":         request.form.get("plant", "").strip(),
        "plant_code":    request.form.get("plant_code", "").strip(),
    }
    try:
        database.add_employee(
            emp["employee_code"], emp["employee_name"], emp["designation"],
            emp["category"], emp["plant"], emp["plant_code"],
        )
        flash(f"Added employee {emp['employee_code']}.", "success")
        res = google_sheets.push_id_employee_add(emp)
        if res["ok"]:
            flash(f"☁️ {res['message']}", "info")
        else:
            flash(f"⚠️ Saved locally but Google Sheet not updated: {res['message']}", "warning")
    except ValueError as e:
        flash(str(e), "error")
    return redirect(url_for("page_data_uploader") + "?open_modal=employee&tab=add")


@app.route("/action/update-employee/<code>", methods=["POST"])
def update_employee(code):
    emp = {
        "employee_code": code,
        "employee_name": request.form.get("name", "").strip(),
        "designation":   request.form.get("designation", "").strip(),
        "category":      request.form.get("category", ""),
        "plant":         request.form.get("plant", "").strip(),
        "plant_code":    request.form.get("plant_code", "").strip(),
    }
    try:
        database.update_employee(
            code, emp["employee_name"], emp["designation"],
            emp["category"], emp["plant"], emp["plant_code"],
        )
        flash(f"Updated employee {code}.", "success")
        res = google_sheets.push_id_employee_update(emp)
        if res["ok"]:
            flash(f"☁️ {res['message']}", "info")
        else:
            flash(f"⚠️ Saved locally but Google Sheet not updated: {res['message']}", "warning")
    except ValueError as e:
        flash(str(e), "error")
    return redirect(url_for("page_data_uploader") + "?open_modal=employee&tab=edit")


@app.route("/action/delete-employee/<code>", methods=["POST"])
def delete_employee(code):
    try:
        database.delete_employee(code)
        flash(f"Deleted employee {code}.", "success")
        res = google_sheets.push_id_employee_delete(code)
        if res["ok"]:
            flash(f"☁️ {res['message']}", "info")
        else:
            flash(f"⚠️ Deleted locally but Google Sheet not updated: {res['message']}", "warning")
    except ValueError as e:
        flash(str(e), "error")
    return redirect(url_for("page_data_uploader") + "?open_modal=employee&tab=delete")


@app.route("/action/upload-backend", methods=["POST"])
def upload_backend():
    if "file" not in request.files or request.files["file"].filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("page_data_uploader") + "#backend")
    f = request.files["file"]
    replace = request.form.get("mode", "replace") == "replace"
    try:
        df, warns = data_loader.load_backend_data(f)
        for w in warns:
            flash(w, "warning")
        saved = data_loader.save_backend_data(df, source_file=f.filename, replace=replace)
        flash(f"{'Replaced' if replace else 'Appended'} {saved:,} backend rows.", "success")
    except Exception as exc:
        flash(f"Upload failed: {exc}", "error")
    return redirect(url_for("page_data_uploader") + "#backend")


@app.route("/action/upload-maintenance", methods=["POST"])
def upload_maintenance():
    if "file" not in request.files or request.files["file"].filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("page_data_uploader") + "#maintenance")
    f = request.files["file"]
    try:
        df, warns = data_loader.load_maintenance_cost(f)
        for w in warns:
            flash(w, "warning")
        saved = data_loader.save_maintenance_cost(df)
        flash(f"Saved {saved:,} plant maintenance cost rows.", "success")
    except Exception as exc:
        flash(f"Upload failed: {exc}", "error")
    return redirect(url_for("page_data_uploader") + "#maintenance")


@app.route("/action/fetch-oracle", methods=["POST"])
def fetch_oracle():
    from_s = request.form.get("from_date", str(_date.today().replace(day=1)))
    to_s   = request.form.get("to_date",   str(_date.today()))
    replace = request.form.get("mode", "replace") == "replace"
    try:
        fd = _date.fromisoformat(from_s)
        td = _date.fromisoformat(to_s)
        result = oracle_connector.test_connection()
        if not result["success"]:
            flash(f"Oracle connection failed: {result['error']}", "error")
            return redirect(url_for("page_data_uploader") + "#oracle")
        df_ora, warns = oracle_connector.fetch_backend_data(fd, td)
        for w in warns:
            flash(w, "warning")
        if df_ora.empty:
            flash("Oracle returned no rows for the selected date range.", "warning")
        else:
            saved = oracle_connector.save_oracle_backend_data(df_ora, fd, td, replace=replace)
            database.purge_old_oracle_data()   # keep only previous + current month
            flash(f"{'Replaced' if replace else 'Appended'} {saved:,} rows from Oracle.", "success")
    except Exception as exc:
        flash(f"Oracle fetch failed: {exc}", "error")
    return redirect(url_for("page_data_uploader") + "#oracle")


@app.route("/action/run-calculation", methods=["POST"])
def run_calculation():
    from_s = request.form.get("from_date")
    to_s   = request.form.get("to_date")
    try:
        fd = _date.fromisoformat(from_s)
        td = _date.fromisoformat(to_s)
        if fd > td:
            flash("'From Date' must be on or before 'To Date'.", "error")
            return redirect(url_for("page_calculate"))
        result = calculator.run_calculation(
            month=fd.month, year=fd.year,
            start_date=str(fd), end_date=str(td),
        )
        if result["error"]:
            flash(f"Calculation failed: {result['error']}", "error")
        else:
            flash(f"Calculation complete — {result['mapped']:,} employees, {result['unmapped']} unmapped.", "success")
        _ss("calc_ran", True)
    except Exception as exc:
        flash(f"Calculation error: {exc}", "error")
    return redirect(url_for("page_calculate"))


@app.route("/action/run-validation", methods=["POST"])
def run_validation():
    try:
        errs: list = []
        validations._validate_master_data(errs)
        validations._validate_backend_data(errs)
        validations._validate_maintenance_cost(errs)
        validations.clear_validation_errors()
        if errs:
            database.insert_rows("validation_errors", errs)
        ran_at = _dt.now().isoformat(timespec="seconds")
        database.set_settings_bulk({
            "last_validation_at":     ran_at,
            "last_validation_errors": str(len(errs)),
        })
        flash(f"Validation complete — {len(errs)} error(s) found.", "success" if not errs else "warning")
    except Exception as exc:
        flash(f"Validation error: {exc}", "error")
    return redirect(url_for("page_validation"))


@app.route("/action/clear-validation", methods=["POST"])
def clear_validation():
    validations.clear_validation_errors()
    database.set_settings_bulk({"last_validation_at": "", "last_validation_errors": "0"})
    flash("All validation errors cleared.", "success")
    return redirect(url_for("page_validation"))


@app.route("/action/save-smtp", methods=["POST"])
def save_smtp():
    pwd = request.form.get("password", "").strip()
    to_save = {
        "smtp_host":        request.form.get("host", "").strip(),
        "smtp_port":        request.form.get("port", "587").strip(),
        "smtp_sender":      request.form.get("sender", "").strip(),
        "smtp_use_tls":     "true" if "use_tls" in request.form else "false",
        "email_default_to": request.form.get("default_to", "").strip(),
        "email_default_cc": request.form.get("default_cc", "").strip(),
        "email_subject":    request.form.get("subject", "").strip(),
    }
    if pwd:
        to_save["smtp_password"] = pwd
    database.set_settings_bulk(to_save)
    flash("Email settings saved.", "success")
    return redirect(url_for("page_settings"))


@app.route("/action/test-smtp", methods=["POST"])
def test_smtp():
    cfg = email_helper.get_smtp_config()
    try:
        res = email_helper.send_report_email(
            to_emails=cfg["sender"], cc_emails="",
            subject="Test email — Batching Incentive Calculator",
            body="This is a test email confirming your SMTP settings work.",
        )
        if res["success"]:
            flash(f"Test email sent to {cfg['sender']}. Check the inbox.", "success")
        else:
            flash(f"Test failed: {res['error']}", "error")
    except Exception as exc:
        flash(f"Test failed: {exc}", "error")
    return redirect(url_for("page_settings"))


@app.route("/action/save-oracle", methods=["POST"])
def save_oracle():
    pwd = request.form.get("password", "").strip()
    to_save = {
        "oracle_host":              request.form.get("host", "").strip(),
        "oracle_port":              request.form.get("port", "").strip(),
        "oracle_service":           request.form.get("service", "").strip(),
        "oracle_user":              request.form.get("user", "").strip(),
        "oracle_status_filter":     request.form.get("status_filter", "").strip(),
        "oracle_instantclient_dir": request.form.get("instantclient", "").strip(),
    }
    if pwd:
        to_save["oracle_password"] = pwd
    database.set_settings_bulk(to_save)
    flash("Oracle settings saved.", "success")
    return redirect(url_for("page_settings"))


@app.route("/action/test-oracle", methods=["POST"])
def test_oracle():
    try:
        result = oracle_connector.test_connection()
        if result["success"]:
            flash(f"Connection successful! {result.get('version', '')}", "success")
        else:
            flash(f"Connection failed: {result['error']}", "error")
    except Exception as exc:
        flash(f"Test failed: {exc}", "error")
    return redirect(url_for("page_settings"))


@app.route("/action/save-bg-settings", methods=["POST"])
def save_bg_settings():
    auto    = "auto_theme" in request.form
    animate = "animate"    in request.form
    theme   = request.form.get("manual_theme", "Daytime")
    database.set_settings_bulk({
        "bg_auto_theme":    "true" if auto    else "false",
        "bg_animate":       "true" if animate else "false",
        "bg_manual_theme":  theme,
    })
    flash("Background settings saved.", "success")
    return redirect(url_for("page_settings"))


@app.route("/action/clear-cache", methods=["POST"])
def clear_cache():
    today = str(_date.today())
    last  = database.get_setting("last_cache_clear_date", "")
    if last == today:
        flash("Cache already compacted today — come back tomorrow.", "warning")
        return redirect(url_for("page_settings"))
    try:
        summary = cache_helpers.clear_cache()
        if summary["error"]:
            flash(f"Could not compact database: {summary['error']}", "error")
        else:
            database.set_setting("last_cache_clear_date", today)
            freed = summary["freed_mb"]
            note  = f"Freed {freed} MB." if freed > 0 else "Database was already compact."
            flash(f"Done! {note} Database size now {summary['after_mb']} MB.", "success")
    except Exception as exc:
        flash(f"Cache clear failed: {exc}", "error")
    return redirect(url_for("page_settings"))


# ── AJAX endpoint (topbar theme) ─────────────────────────────────────────────

@app.route("/api/bg-settings", methods=["POST"])
def api_bg_settings():
    data = request.get_json(silent=True) or {}
    to_save = {}
    if "auto_theme" in data:
        to_save["bg_auto_theme"] = "true" if data.get("auto_theme") else "false"
    if "manual_theme" in data:
        to_save["bg_manual_theme"] = data.get("manual_theme", "Daytime")
    if "animate" in data:
        to_save["bg_animate"] = "true" if data.get("animate") else "false"
    if to_save:
        database.set_settings_bulk(to_save)
    return jsonify({"ok": True})


# ── Oracle live status (shared by all modules + launcher icon) ───────────────
@app.route("/api/oracle-status")
def api_oracle_status():
    """Live, honest Oracle status. Used by the launcher icon and every module
    topbar. One shared Oracle connection serves all four modules."""
    cfg = oracle_connector.get_oracle_config()
    configured = oracle_connector.is_configured(cfg)
    reachable  = oracle_connector.is_reachable(cfg) if configured else False
    if not configured:
        state, label = "unconfigured", "Oracle not configured"
    elif reachable:
        state, label = "connected", f"Oracle connected — {cfg['host']}"
    else:
        state, label = "unreachable", "Oracle unreachable (check office network / VPN)"
    return jsonify({"configured": configured, "reachable": reachable,
                    "state": state, "label": label})


@app.route("/tp/api/table-columns")
def tp_api_table_columns():
    """Return all column names in rdc_batch_trx_headers so the user can find the right plant column."""
    try:
        cols = oracle_connector.get_table_columns()
        return jsonify({"ok": True, "columns": cols})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)})


# ── DOWNLOAD ENDPOINTS ────────────────────────────────────────────────────────

def _snapshot_dfs():
    filtered = _s("rpt_filtered", [])
    unmapped  = _s("rpt_unmapped_snap", [])
    from_s    = _s("rpt_from",  str(_date.today()))
    to_s      = _s("rpt_to",    str(_date.today()))
    applied   = _s("rpt_applied_filters", "None")
    meta      = _build_meta(from_s, to_s, filtered, unmapped, applied)

    col_set = set(RESULT_COLS) | {"category", "month", "year"}
    all_cols = [c for c in (RESULT_COLS + ["category", "month", "year"])
                if c not in ("_cls",)]

    df_f = pd.DataFrame(filtered) if filtered else pd.DataFrame()
    df_u = pd.DataFrame(unmapped) if unmapped else pd.DataFrame()
    val_df = database.read_table("validation_errors")
    return df_f, df_u, val_df, meta, from_s, to_s


@app.route("/download/excel")
def download_excel():
    df_f, df_u, val_df, meta, from_s, to_s = _snapshot_dfs()
    if df_f.empty:
        flash("No report data. Load a report on the View Reports page first.", "warning")
        return redirect(url_for("page_reports"))
    try:
        xlsx = report_generator.generate_excel_report(df_f, df_u, val_df, meta)
        fname = f"incentive_report_{from_s}_to_{to_s}.xlsx"
        return send_file(
            io.BytesIO(xlsx), as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        flash(f"Could not build Excel: {exc}", "error")
        return redirect(url_for("page_reports"))


@app.route("/download/csv")
def download_csv():
    filtered = _s("rpt_filtered", [])
    from_s   = _s("rpt_from",  str(_date.today()))
    to_s     = _s("rpt_to",    str(_date.today()))
    if not filtered:
        flash("No report data. Load a report on the View Reports page first.", "warning")
        return redirect(url_for("page_reports"))
    df = pd.DataFrame(filtered)
    # Rename columns for display
    df_show = df[[c for c in RESULT_COLS if c in df.columns]].rename(columns=RESULT_LABELS)
    csv_bytes = df_show.to_csv(index=False).encode("utf-8")
    fname = f"incentive_report_{from_s}_to_{to_s}.csv"
    return send_file(io.BytesIO(csv_bytes), as_attachment=True,
                     download_name=fname, mimetype="text/csv")


@app.route("/action/send-email", methods=["POST"])
def send_email():
    df_f, df_u, val_df, meta, from_s, to_s = _snapshot_dfs()
    if df_f.empty:
        flash("No report data. Load a report on the View Reports page first.", "warning")
        return redirect(url_for("page_reports"))

    to_addr  = request.form.get("to", "")
    cc_addr  = request.form.get("cc", "")
    subject  = request.form.get("subject", "")
    body     = request.form.get("body", "")
    incl_tables = "include_tables" in request.form

    try:
        fname     = f"incentive_report_{from_s}_to_{to_s}.xlsx"
        xlsx_data = report_generator.generate_excel_report(df_f, df_u, val_df, meta)

        # Generate PNG preview image and embed it inline in the email body
        exports_dir = os.path.join(os.path.dirname(__file__), "exports")
        img_path = email_helper.create_report_preview_image(
            df_f,
            os.path.join(exports_dir, "email_preview_id.png"),
            month_label=meta.get("date_range", ""))
        html_body = email_helper.wrap_html_body_with_image(body,
                                                           excel_attached=True)

        res = email_helper.send_report_email(
            to_emails=to_addr, cc_emails=cc_addr,
            subject=subject, body=body,
            attachment_bytes=xlsx_data, attachment_name=fname,
            html_body=html_body, inline_image_path=img_path,
        )
        if res["success"]:
            flash(f"Report emailed to {to_addr}.", "success")
        else:
            flash(f"Email failed: {res['error']}", "error")
    except Exception as exc:
        flash(f"Could not send: {exc}", "error")

    return redirect(url_for("page_reports"))


@app.route("/action/restart-server", methods=["POST"])
def restart_server():
    def _restart():
        import time
        time.sleep(1)
        # Spawn a new process first, then exit this one (works on Windows)
        subprocess.Popen([sys.executable] + sys.argv,
                         creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0)
        os._exit(0)
    threading.Thread(target=_restart, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/action/save-email-schedule", methods=["POST"])
def save_email_schedule():
    sched_time = request.form.get("sched_time", "08:00").strip()
    sched_to   = request.form.get("sched_to",   "").strip()
    sched_cc   = request.form.get("sched_cc",   "").strip()
    database.set_settings_bulk({
        "email_schedule_time": sched_time,
        "email_schedule_to":   sched_to,
        "email_schedule_cc":   sched_cc,
    })
    if _scheduler:
        try:
            h, m = map(int, sched_time.split(":"))
        except Exception:
            h, m = 8, 0
        _scheduler.reschedule_job(
            "monthly_report",
            trigger=CronTrigger(day=1, hour=h, minute=m),
        )
    flash("Schedule settings saved.", "success")
    return redirect(url_for("page_settings"))


@app.route("/action/toggle-email-schedule", methods=["POST"])
def toggle_email_schedule():
    current = database.get_setting("email_schedule_enabled", "false")
    new_val = "false" if current == "true" else "true"
    database.set_setting("email_schedule_enabled", new_val)
    flash(f"Scheduled monthly report {'enabled' if new_val == 'true' else 'disabled'}.", "success")
    return redirect(url_for("page_settings"))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("  RDC Batching Incentive Calculator")
    print("  Flask server starting on http://localhost:2001")
    print("=" * 60)
    app.run(host="0.0.0.0", port=2001, debug=True, use_reloader=False)
