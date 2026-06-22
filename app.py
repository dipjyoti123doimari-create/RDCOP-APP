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
import logging
import os
import subprocess
import sys
import threading
import traceback
import uuid
from datetime import date as _date, datetime as _dt, timedelta

# Load .env if present (for SECRET_KEY, ADMIN_USERNAME, etc.)
try:
    from dotenv import load_dotenv as _load_dotenv
    _load_dotenv()
except ImportError:
    pass  # python-dotenv not installed; rely on environment variables set externally

import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from flask import (Flask, flash, g, jsonify, make_response, redirect, render_template,
                   request, send_file, url_for)

import auth
import cache_helpers
import calculator
import config
import data_loader
import database
import email_helper
import google_sheets
import oracle_connector
import report_generator
import btrtp_calculator
import ecmd_calculator
import tp_calculator
import validations

# ── File logging (survives minimised/closed terminal windows) ────────────────
_LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.log")
logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
_log = logging.getLogger(__name__)

def _handle_unhandled_exception(exc_type, exc_value, exc_tb):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    _log.critical("Unhandled exception — server will exit:\n%s",
                  "".join(traceback.format_exception(exc_type, exc_value, exc_tb)))

sys.excepthook = _handle_unhandled_exception

# ── App setup ────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "rdc-incentive-local-key-2025")
app.config["SESSION_PERMANENT"] = False
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(minutes=20)

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


# ── Real-time operation progress (single-user, single active job at a time) ───
_progress: dict = {"pct": 0, "msg": ""}
_progress_lock = threading.Lock()


def _set_progress(pct: int, msg: str = ""):
    with _progress_lock:
        _progress["pct"] = pct
        _progress["msg"] = msg


# ── Boot ─────────────────────────────────────────────────────────────────────
database.init_db()
auth.bootstrap_admin(app)

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


def _id_daily_oracle_fetch_job():
    """
    Runs every day at 00:15.
    Fetches I&D Oracle backend_data for previous month (full) + current month
    (1st → today). Keeps local store always at 2-month rolling window.
    """
    if not oracle_connector.is_configured():
        return
    import calendar as _cal
    today    = _date.today()
    cur_from = str(today.replace(day=1))
    cur_to   = str(today)
    if today.month == 1:
        prev_from = _date(today.year - 1, 12, 1)
        prev_to   = _date(today.year - 1, 12, _cal.monthrange(today.year - 1, 12)[1])
    else:
        prev_from = _date(today.year, today.month - 1, 1)
        prev_to   = _date(today.year, today.month - 1,
                          _cal.monthrange(today.year, today.month - 1)[1])
    try:
        for fd, td in [(str(prev_from), str(prev_to)), (cur_from, cur_to)]:
            df_ora, _ = oracle_connector.fetch_backend_data(fd, td)
            if not df_ora.empty:
                oracle_connector.save_oracle_backend_data(df_ora, fd, td, replace=True)
        database.purge_old_oracle_data()
        print(f"[id-daily-fetch] done — prev: {prev_from}→{prev_to}, cur: {cur_from}→{cur_to}")
    except Exception as exc:
        print(f"[id-daily-fetch] error: {exc}")


def _btrtp_daily_oracle_fetch_job():
    """
    Runs every day at 00:20.
    Fetches BTRTP Oracle data for previous month (full) + current month (1st → today).
    """
    if not oracle_connector.is_configured():
        return
    import calendar as _cal
    today    = _date.today()
    cur_from = str(today.replace(day=1))
    cur_to   = str(today)
    if today.month == 1:
        prev_from = _date(today.year - 1, 12, 1)
        prev_to   = _date(today.year - 1, 12, _cal.monthrange(today.year - 1, 12)[1])
    else:
        prev_from = _date(today.year, today.month - 1, 1)
        prev_to   = _date(today.year, today.month - 1,
                          _cal.monthrange(today.year, today.month - 1)[1])
    try:
        for fd, td in [(str(prev_from), str(prev_to)), (cur_from, cur_to)]:
            raw_df, _ = oracle_connector.fetch_btrtp_data(fd, td)
            if not raw_df.empty:
                parsed, _ = btrtp_calculator.parse_btrtp_oracle_df(raw_df)
                oracle_connector.save_btrtp_oracle_data(parsed, replace=True)
        database.purge_old_oracle_data()
        print(f"[btrtp-daily-fetch] done — prev: {prev_from}→{prev_to}, cur: {cur_from}→{cur_to}")
    except Exception as exc:
        print(f"[btrtp-daily-fetch] error: {exc}")


def _startup_oracle_fetch():
    """
    Runs once 45 seconds after startup in a background thread.
    Seeds local Oracle cache for all modules if Oracle is reachable.
    This ensures data is fresh even when the midnight cron was missed
    (e.g. server was off, VPN wasn't connected at midnight).
    """
    import time as _time
    _time.sleep(45)
    if not oracle_connector.is_configured() or not oracle_connector.is_reachable():
        print("[startup-fetch] Oracle not reachable — skipping startup seed")
        return
    print("[startup-fetch] Oracle reachable — seeding all modules")
    _tp_daily_oracle_fetch_job()
    _id_daily_oracle_fetch_job()
    _btrtp_daily_oracle_fetch_job()
    print("[startup-fetch] done")


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
    # RDC-I&D daily Oracle fetch
    _scheduler.add_job(_id_daily_oracle_fetch_job,
                       CronTrigger(hour=0, minute=15),
                       id="id_daily_oracle_fetch", replace_existing=True)
    # RDC-BTRTP daily Oracle fetch
    _scheduler.add_job(_btrtp_daily_oracle_fetch_job,
                       CronTrigger(hour=0, minute=20),
                       id="btrtp_daily_oracle_fetch", replace_existing=True)
    _scheduler.start()


_start_scheduler()

# Seed Oracle cache on startup (45s delay — waits for network/VPN to settle)
import threading as _threading
_threading.Thread(target=_startup_oracle_fetch, daemon=True).start()

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


# ── Global auth gate ─────────────────────────────────────────────────────────
# Enforce login for every request except /login, /logout and static assets.
# This is the "minimum disruption" approach — existing routes need no changes.

_AUTH_EXEMPT = {"/login", "/logout"}

@app.before_request
def _require_login():
    if request.path.startswith("/static/"):
        return
    if request.path in _AUTH_EXEMPT:
        return
    user = auth.get_current_user()
    if user is None:
        if request.path.startswith("/api/"):
            return jsonify({"error": "Not authenticated"}), 401
        from flask import flash as _flash
        _flash("Please log in to continue.", "warning")
        return redirect(url_for("page_login", next=request.path))
    # Inject into g so route handlers can read it without calling get_current_user() again
    g.current_user = user

    # ── Role-based path restrictions ─────────────────────────────────────────
    role = user.get("role", "")
    path = request.path

    # SUPER_ADMIN-only: settings, calculate, upload/sync actions, admin panel
    _admin_only_patterns = (
        "/settings", "/calculate", "/data-uploader",
        "/action/upload-", "/action/sync-", "/action/fetch-oracle",
        "/action/run-calculation", "/action/run-validation",
        "/action/save-smtp", "/action/save-oracle",
        "/action/add-employee", "/action/update-employee", "/action/delete-employee",
        "/action/add-waiver", "/action/delete-waiver",
        "/action/restart-server", "/action/save-email-schedule",
        "/action/toggle-email-schedule", "/action/clear-",
        "/action/assign-maintenance", "/action/delete-maintenance",
        "/action/toggle-auto-sync", "/action/save-",
        "/tp/action/add-plant", "/tp/action/update-plant", "/tp/action/delete-plant",
        "/btrtp/settings", "/tp/settings",
    )
    if role != auth.SUPER_ADMIN:
        for pattern in _admin_only_patterns:
            if pattern in path:
                auth.log_activity(user, "ACCESS_DENIED",
                                  details={"path": path, "method": request.method})
                return render_template("access_denied.html",
                                       current_user=user,
                                       required_roles=[auth.SUPER_ADMIN]), 403

    # Email send: PLANT_USER cannot send emails
    if role == auth.PLANT_USER and "/action/send-email" in path:
        return render_template("access_denied.html",
                               current_user=user,
                               required_roles=[auth.SUPER_ADMIN, auth.HO_VIEWER,
                                               auth.FINANCE_VIEWER, auth.REGIONAL_USER]), 403

    # Non-SUPER_ADMIN users must not enter any module page — redirect to user dashboard
    _module_prefixes = ("/dashboard", "/id", "/tp/", "/btrtp/", "/ecmd/")
    if role != auth.SUPER_ADMIN and request.method == "GET":
        for pfx in _module_prefixes:
            if path == pfx.rstrip("/") or path.startswith(pfx):
                return redirect(url_for("page_home"))


# ── Context processor (sidebar active-page + bg settings + auth) ────────────

@app.context_processor
def _ctx():
    auth_ctx = auth.inject_auth_context()
    return dict(
        active_page=_s("active_page", ""),
        bg_auto=bg_auto(),
        bg_animate=bg_animate(),
        bg_theme=bg_theme(),
        now=_dt.now(),
        **auth_ctx,
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

# ── Auth routes ───────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def page_login():
    return auth.login_view()


@app.route("/logout", methods=["POST"])
def page_logout():
    return auth.logout_view()


@app.route("/change-password", methods=["GET", "POST"])
@auth.login_required
def page_change_password():
    return auth.change_password_view()


# ── Admin / User Management routes (SUPER_ADMIN only) ─────────────────────────

@app.route("/admin/users")
@auth.admin_required
def admin_users():
    users = database.get_all_users()
    plant_map = {u["id"]: database.get_user_plant_access(u["id"]) for u in users}
    return render_template("admin/users.html",
                           active_page="users",
                           users=users,
                           plant_map=plant_map)


@app.route("/admin/users/create", methods=["GET", "POST"])
@auth.admin_required
def admin_create_user():
    all_plants = database.get_tp_plants()
    if request.method == "GET":
        return render_template("admin/user_form.html",
                               active_page="create",
                               edit_mode=False,
                               user={},
                               all_plants=all_plants,
                               assigned_plants=[],
                               assigned_plant_names=[],
                               allowed_roles=auth.ALLOWED_ROLES)
    # POST
    from werkzeug.security import generate_password_hash as _hash
    full_name  = request.form.get("full_name", "").strip()
    username   = request.form.get("username", "").strip()
    email      = request.form.get("email", "").strip()
    role       = request.form.get("role", "PLANT_USER")
    password   = request.form.get("password", "")
    is_active  = request.form.get("is_active", "1") == "1"
    must_chg   = request.form.get("must_change_password", "1") == "1"
    plant_names = request.form.getlist("plant_names")

    if not all([full_name, username, email, password]):
        flash("All required fields must be filled.", "error")
        return render_template("admin/user_form.html",
                               active_page="create",
                               edit_mode=False,
                               user=dict(full_name=full_name, username=username, email=email, role=role),
                               all_plants=all_plants,
                               assigned_plants=[],
                               assigned_plant_names=[],
                               allowed_roles=auth.ALLOWED_ROLES), 400
    if len(password) < 8:
        flash("Password must be at least 8 characters.", "error")
        return render_template("admin/user_form.html",
                               active_page="create",
                               edit_mode=False,
                               user=dict(full_name=full_name, username=username, email=email, role=role),
                               all_plants=all_plants,
                               assigned_plants=[],
                               assigned_plant_names=[],
                               allowed_roles=auth.ALLOWED_ROLES), 400
    try:
        user_id = database.create_user(
            full_name=full_name, email=email, username=username,
            password_hash=_hash(password),
            role=role, is_active=is_active, must_change_password=must_chg,
        )
        if role in auth.RESTRICTED_ROLES and plant_names:
            code_map = {p["plant_name"]: p["plant_code"] for p in all_plants}
            plant_list = [{"plant_name": n, "plant_code": code_map.get(n, "")} for n in plant_names]
            database.set_user_plants(user_id, plant_list)
        auth.log_activity(auth.get_current_user(), "CREATE_USER",
                          details={"username": username, "role": role})
        flash(f"User '{username}' created successfully.", "success")
        return redirect(url_for("admin_users"))
    except Exception as exc:
        flash(f"Error creating user: {exc}", "error")
        return render_template("admin/user_form.html",
                               active_page="create",
                               edit_mode=False,
                               user=dict(full_name=full_name, username=username, email=email, role=role),
                               all_plants=all_plants,
                               assigned_plants=[],
                               assigned_plant_names=[],
                               allowed_roles=auth.ALLOWED_ROLES), 400


@app.route("/admin/users/<int:user_id>/edit", methods=["GET", "POST"])
@auth.admin_required
def admin_edit_user(user_id):
    from werkzeug.security import generate_password_hash as _hash
    user = database.get_user_by_id(user_id)
    if not user:
        flash("User not found.", "error")
        return redirect(url_for("admin_users"))
    all_plants      = database.get_tp_plants()
    assigned_plants = database.get_user_plant_access(user_id)
    assigned_names  = [p["plant_name"] for p in assigned_plants]

    if request.method == "GET":
        return render_template("admin/user_form.html",
                               active_page="users",
                               edit_mode=True,
                               user=user,
                               all_plants=all_plants,
                               assigned_plants=assigned_plants,
                               assigned_plant_names=assigned_names,
                               allowed_roles=auth.ALLOWED_ROLES)
    # POST
    full_name  = request.form.get("full_name", "").strip()
    email      = request.form.get("email", "").strip()
    role       = request.form.get("role", user["role"])
    is_active  = request.form.get("is_active", "1") == "1"
    must_chg   = request.form.get("must_change_password", "0") == "1"
    password   = request.form.get("password", "").strip()
    plant_names = request.form.getlist("plant_names")

    database.update_user(user_id, full_name=full_name, email=email, role=role,
                         is_active=is_active, must_change_password=must_chg)
    if password:
        if len(password) < 8:
            flash("Password must be at least 8 characters.", "error")
            return redirect(url_for("admin_edit_user", user_id=user_id))
        database.update_user_password(user_id, _hash(password))

    code_map = {p["plant_name"]: p["plant_code"] for p in all_plants}
    plant_list = [{"plant_name": n, "plant_code": code_map.get(n, "")} for n in plant_names]
    database.set_user_plants(user_id, plant_list)

    auth.log_activity(auth.get_current_user(), "EDIT_USER",
                      details={"user_id": user_id, "role": role})
    flash(f"User '{user['username']}' updated.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:user_id>/toggle", methods=["POST"])
@auth.admin_required
def admin_toggle_user(user_id):
    current = auth.get_current_user()
    if current and current["id"] == user_id:
        flash("You cannot deactivate your own account.", "error")
        return redirect(url_for("admin_users"))
    user = database.get_user_by_id(user_id)
    if not user:
        flash("User not found.", "error")
        return redirect(url_for("admin_users"))
    database.update_user(user_id,
                         full_name=user["full_name"], email=user["email"],
                         role=user["role"],
                         is_active=not user["is_active"],
                         must_change_password=user["must_change_password"])
    status = "activated" if not user["is_active"] else "deactivated"
    auth.log_activity(current, "TOGGLE_USER",
                      details={"user_id": user_id, "new_status": status})
    flash(f"User '{user['username']}' {status}.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:user_id>/reset-password", methods=["GET", "POST"])
@auth.admin_required
def admin_reset_password(user_id):
    from werkzeug.security import generate_password_hash as _hash
    target_user = database.get_user_by_id(user_id)
    if not target_user:
        flash("User not found.", "error")
        return redirect(url_for("admin_users"))
    if request.method == "GET":
        return render_template("admin/reset_password.html",
                               active_page="users",
                               target_user=target_user)
    new_pw  = request.form.get("new_password", "")
    conf_pw = request.form.get("confirm_password", "")
    must_chg = request.form.get("must_change", "1") == "1"
    if len(new_pw) < 8 or new_pw != conf_pw:
        flash("Passwords must match and be at least 8 characters.", "error")
        return redirect(url_for("admin_reset_password", user_id=user_id))
    database.update_user_password(user_id, _hash(new_pw), must_change_password=must_chg)
    auth.log_activity(auth.get_current_user(), "RESET_PASSWORD",
                      details={"target_user_id": user_id})
    flash(f"Password for '{target_user['username']}' has been reset.", "success")
    return redirect(url_for("admin_edit_user", user_id=user_id))


@app.route("/admin/audit-log")
@auth.admin_required
def admin_audit_log():
    rows = database.get_login_audit_log(limit=500)
    return render_template("admin/audit_log.html",
                           active_page="audit",
                           rows=rows)


@app.route("/admin/activity-log")
@auth.admin_required
def admin_activity_log():
    rows = database.get_user_activity_log(limit=500)
    return render_template("admin/activity_log.html",
                           active_page="activity",
                           rows=rows)


# ── Home page ─────────────────────────────────────────────────────────────────

@app.route("/")
@auth.login_required
def page_home():
    user = auth.get_current_user()
    # SUPER_ADMIN gets the module launcher as before
    if user and user.get("role") == "SUPER_ADMIN":
        resp = make_response(render_template("home.html"))
        resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
        return resp

    # All other roles get the unified user dashboard
    from datetime import date as _today
    MONTHS = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
    def _mn(m): return MONTHS[m-1] if m else "—"

    now       = _today.today()
    cur_ym    = f"{now.year}-{now.month:02d}"
    # previous calendar month
    if now.month == 1:
        prev_m, prev_y = 12, now.year - 1
    else:
        prev_m, prev_y = now.month - 1, now.year
    prev_ym = f"{prev_y}-{prev_m:02d}"

    conn = database.get_connection()
    # defaults — overwritten below as each section runs
    tp_last_known_sync = bt_last_known_sync = None
    try:
        cur = conn.cursor()

        # ── I&D last calculated ──────────────────────────────
        cur.execute("""SELECT month, year,
            COUNT(*) as total_emp,
            SUM(CASE WHEN incentive_eligible='Yes' THEN 1 ELSE 0 END) as inc_emp,
            SUM(CASE WHEN deduction_amount > 0   THEN 1 ELSE 0 END) as ded_emp,
            ROUND(SUM(COALESCE(incentive_amount,0)),0) as total_inc,
            ROUND(SUM(COALESCE(deduction_amount,0)),0) as total_ded
            FROM calculation_results GROUP BY year,month
            ORDER BY year DESC, month DESC LIMIT 1""")
        row = cur.fetchone()
        id_last = dict(row) if row else None
        if id_last:
            id_last["month_name"] = _mn(id_last["month"])

        # Allowed plants for this user (empty list = all plants for global roles)
        allowed_plants  = auth.get_user_allowed_plants(user)
        user_plant_rows = database.get_user_plant_access(user["id"]) if allowed_plants else []
        is_restricted   = bool(allowed_plants)  # True for REGIONAL_USER / PLANT_USER

        def _plant_filter(rows, col="plant"):
            if not is_restricted:
                return rows
            al = [p.lower() for p in allowed_plants]
            return [r for r in rows if str(r.get(col, "")).lower() in al]

        # I&D last calculated — full detail rows for drill-down
        id_last_rows = []
        if id_last:
            cur.execute("""SELECT employee_code, employee_name, designation, plant,
                total_quantity, incentive_eligible, incentive_amount, deduction_amount, remarks
                FROM calculation_results WHERE month=? AND year=?
                ORDER BY plant, employee_name""",
                (id_last["month"], id_last["year"]))
            cols = [d[0] for d in cur.description]
            id_last_rows = _plant_filter(
                [dict(zip(cols, r)) for r in cur.fetchall()], col="plant")
            # Recompute summary counts from filtered rows so the card matches
            if is_restricted:
                id_last["total_emp"] = len(id_last_rows)
                id_last["inc_emp"]   = sum(1 for r in id_last_rows if r.get("incentive_amount") and r["incentive_amount"] > 0)
                id_last["ded_emp"]   = sum(1 for r in id_last_rows if r.get("deduction_amount") and r["deduction_amount"] > 0)
                id_last["total_inc"] = sum(r.get("incentive_amount") or 0 for r in id_last_rows)

        # I&D current month raw (backend_data) — no plant column, show pan-India total
        cur.execute("""SELECT COUNT(DISTINCT created_by) as emp_count,
            ROUND(SUM(quantity),0) as total_qty,
            MAX(uploaded_at) as last_sync
            FROM backend_data WHERE substr(date,1,7)=?""", (cur_ym,))
        row = cur.fetchone()
        id_cur = dict(row) if (row and row[0]) else None
        if id_cur:
            id_cur["month_name"] = _mn(now.month)
            id_cur["month"] = now.month
            id_cur["year"]  = now.year

        # ── TP last calculated ───────────────────────────────
        cur.execute("""SELECT month, year,
            COUNT(DISTINCT lookup_code) as plants,
            ROUND(AVG(throughput_pct),1) as avg_tp,
            SUM(CASE WHEN throughput_pct >= 75 THEN 1 ELSE 0 END) as above_target,
            SUM(CASE WHEN throughput_pct <  75 THEN 1 ELSE 0 END) as below_target,
            ROUND(SUM(total_quantity),0) as total_qty
            FROM tp_results GROUP BY year,month
            ORDER BY year DESC, month DESC LIMIT 1""")
        row = cur.fetchone()
        tp_last = dict(row) if row else None
        if tp_last:
            tp_last["month_name"] = _mn(tp_last["month"])

        # TP last — full detail rows
        tp_last_rows = []
        if tp_last:
            cur.execute("""SELECT plant_name, exco_location, business_head,
                plant_manager, total_quantity, throughput_pct, batch_count
                FROM tp_results WHERE month=? AND year=?
                ORDER BY throughput_pct DESC""",
                (tp_last["month"], tp_last["year"]))
            cols = [d[0] for d in cur.description]
            tp_last_rows = _plant_filter(
                [dict(zip(cols, r)) for r in cur.fetchall()], col="plant_name")
            if is_restricted:
                tp_last["plants"]       = len(tp_last_rows)
                tp_last["above_target"] = sum(1 for r in tp_last_rows if (r.get("throughput_pct") or 0) >= 75)
                tp_last["below_target"] = sum(1 for r in tp_last_rows if (r.get("throughput_pct") or 0) < 75)
                tp_last["avg_tp"]       = round(sum(r.get("throughput_pct") or 0 for r in tp_last_rows) / len(tp_last_rows), 1) if tp_last_rows else 0

        # TP current month raw (tp_oracle_data)
        cur.execute("""SELECT COUNT(DISTINCT lookup_code) as plants,
            ROUND(SUM(quantity),0) as total_qty,
            MAX(fetched_at) as last_sync
            FROM tp_oracle_data WHERE substr(production_date,1,7)=?""", (cur_ym,))
        row = cur.fetchone()
        tp_cur = {"plants": row[0] or 0, "total_qty": row[1] or 0,
                  "last_sync": row[2]} if row else None
        if tp_cur:
            tp_cur["month_name"] = _mn(now.month)
            tp_cur["month"] = now.month
            tp_cur["year"]  = now.year
        # If no current-month data yet but there IS older local data, show last known sync
        if tp_cur is None:
            cur.execute("SELECT MAX(fetched_at) FROM tp_oracle_data")
            r = cur.fetchone()
            tp_last_known_sync = r[0] if r else None
        else:
            tp_last_known_sync = None

        # ── BTRTP last calculated ────────────────────────────
        cur.execute("""SELECT month, year,
            COUNT(*) as batchers,
            ROUND(AVG(throughput_pct),1) as avg_tp,
            SUM(CASE WHEN throughput_pct >= 75 THEN 1 ELSE 0 END) as above_target,
            SUM(CASE WHEN throughput_pct <  75 THEN 1 ELSE 0 END) as below_target
            FROM btrtp_results GROUP BY year,month
            ORDER BY year DESC, month DESC LIMIT 1""")
        row = cur.fetchone()
        bt_last = dict(row) if row else None
        if bt_last:
            bt_last["month_name"] = _mn(bt_last["month"])

        # BTRTP last — full detail rows
        bt_last_rows = []
        if bt_last:
            cur.execute("""SELECT batcher_name, batcher_id, plant_name, exco_location,
                total_quantity, throughput_pct, batch_count
                FROM btrtp_results WHERE month=? AND year=?
                ORDER BY throughput_pct DESC""",
                (bt_last["month"], bt_last["year"]))
            cols = [d[0] for d in cur.description]
            bt_last_rows = _plant_filter(
                [dict(zip(cols, r)) for r in cur.fetchall()], col="plant_name")
            if is_restricted:
                bt_last["batchers"]     = len(bt_last_rows)
                bt_last["above_target"] = sum(1 for r in bt_last_rows if (r.get("throughput_pct") or 0) >= 75)
                bt_last["below_target"] = sum(1 for r in bt_last_rows if (r.get("throughput_pct") or 0) < 75)
                bt_last["avg_tp"]       = round(sum(r.get("throughput_pct") or 0 for r in bt_last_rows) / len(bt_last_rows), 1) if bt_last_rows else 0

        # BTRTP current — use a fresh cursor to avoid any state from previous queries
        bt_last_known_sync = None
        try:
            cur2 = conn.cursor()
            cur2.execute("""SELECT COUNT(DISTINCT lookup_code) as batchers,
                MAX(fetched_at) as last_sync
                FROM btrtp_oracle_data WHERE substr(production_date,1,7)=?""", (cur_ym,))
            row = cur2.fetchone()
            bt_cur = {"batchers": row[0], "last_sync": row[1]} if (row and row[0]) else None
            if bt_cur is None:
                cur2.execute("SELECT MAX(fetched_at) FROM btrtp_oracle_data")
                r = cur2.fetchone()
                bt_last_known_sync = r[0] if r else None
        except Exception:
            bt_cur = None
        if bt_cur:
            bt_cur["month_name"] = _mn(now.month)
            bt_cur["month"] = now.month
            bt_cur["year"]  = now.year

        # ── ECMD last calculated ─────────────────────────────
        cur.execute("""SELECT month, year,
            COUNT(*) as plants,
            ROUND(AVG(energy_per_mt),2) as avg_energy,
            ROUND(AVG(mixer_dg_ratio),1) as avg_dg
            FROM ecmd_results GROUP BY year,month
            ORDER BY year DESC, month DESC LIMIT 1""")
        row = cur.fetchone()
        ec_last = dict(row) if row else None
        if ec_last:
            ec_last["month_name"] = _mn(ec_last["month"])

        # ECMD last — full detail rows
        ec_last_rows = []
        if ec_last:
            cur.execute("""SELECT plant_name, exco_location, plant_manager,
                eb_kwh, dg_kwh, total_kwh, total_volume, energy_per_mt,
                mixer_dg_ratio, diesel_issued_ltrs
                FROM ecmd_results WHERE month=? AND year=?
                ORDER BY plant_name""",
                (ec_last["month"], ec_last["year"]))
            cols = [d[0] for d in cur.description]
            ec_last_rows = _plant_filter(
                [dict(zip(cols, r)) for r in cur.fetchall()], col="plant_name")
            if is_restricted:
                ec_last["plants"] = len(ec_last_rows)

        # ECMD current — readings submitted this month
        cur.execute("""SELECT COUNT(*) as plants FROM ecmd_readings
            WHERE month=? AND year=?""", (now.month, now.year))
        row = cur.fetchone()
        ec_cur = {"plants": row[0], "month_name": _mn(now.month),
                  "month": now.month, "year": now.year} if row else None

    finally:
        conn.close()

    # ECMD entry window
    entry_month = int(database.get_setting("ecmd_entry_open_month", 0) or 0)
    entry_year  = int(database.get_setting("ecmd_entry_open_year",  0) or 0)
    entry_open  = bool(entry_month and entry_year)

    resp = make_response(render_template("user_dashboard.html",
        id_last=id_last, id_last_rows=id_last_rows, id_cur=id_cur,
        tp_last=tp_last, tp_last_rows=tp_last_rows, tp_cur=tp_cur,
        bt_last=bt_last, bt_last_rows=bt_last_rows, bt_cur=bt_cur,
        ec_last=ec_last, ec_last_rows=ec_last_rows, ec_cur=ec_cur,
        cur_month_name=_mn(now.month), cur_year=now.year,
        user_plant_rows=user_plant_rows,
        tp_last_known_sync=tp_last_known_sync,
        bt_last_known_sync=bt_last_known_sync,
        entry_open=entry_open,
        entry_month_name=_mn(entry_month) if entry_open else "",
        entry_year=entry_year,
        bg_auto=bg_auto(), bg_animate=bg_animate(), bg_theme=bg_theme(),
    ))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

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
        return jsonify({"ok": False, "redirect": url_for("tp_data_uploader")})
    try:
        _set_progress(20, "Fetching TP data from Oracle…")
        raw_df, ora_warnings = oracle_connector.fetch_tp_data(from_date, to_date)
        _set_progress(55, "Parsing plant records…")
        parsed, skip_log     = tp_calculator.parse_oracle_df(raw_df)
        _set_progress(75, "Saving to database…")
        oracle_connector.save_tp_oracle_data(raw_df, from_date, to_date, parsed, replace=True)
        _set_progress(90, "Cleaning old records…")
        database.purge_old_oracle_data()
        _mss("tp", "skip_log", skip_log)
        _mss("tp", "ora_from", from_date)
        _mss("tp", "ora_to",   to_date)
        for w in ora_warnings:
            flash(w, "warning")
        flash(f"✅ {len(parsed)} rows fetched & saved ({len(skip_log)} skipped).", "success")
    except Exception as exc:
        flash(f"Oracle error: {exc}", "error")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("tp_data_uploader")})


@app.route("/tp/action/sync-sheets", methods=["POST"])
def tp_sync_sheets():
    sheet_id    = (request.form.get("sheet_id", "").strip()
                   or database.get_module_setting("tp", "gsheet_id",
                                                   database.get_setting("gsheet_id", "")))
    worksheet   = request.form.get("worksheet", "Plant Data for TP").strip()
    _set_progress(20, "Syncing plant data from Google Sheets…")
    result      = google_sheets.sync_tp_plant_data(sheet_id, worksheet)
    _set_progress(85, "Saving settings…")
    if result["error"]:
        flash(f"Sync failed: {result['error']}", "error")
    else:
        database.set_module_setting("tp", "gsheet_id", google_sheets.extract_sheet_id(sheet_id))
        database.set_module_setting("tp", "gsheet_worksheet", worksheet)
        flash(f"✅ {result['rows_synced']} plant rows synced from Google Sheets ({result['mode']} mode).", "success")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("tp_data_uploader")})


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
            return jsonify({"ok": False, "redirect": url_for("tp_calculate")})
        if fd > td:
            flash("'From Date' must be on or before 'To Date'.", "error")
            return jsonify({"ok": False, "redirect": url_for("tp_calculate")})

        month, year = fd.month, fd.year
        _set_progress(20, "Loading Oracle data…")
        plant_rows, location_rows, warnings = tp_calculator.run_tp_calculation(
            month, year, from_date=str(fd), to_date=str(td))
        _set_progress(80, "Saving throughput results…")
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
        _set_progress(100, "Complete")
        return jsonify({"ok": True, "redirect": url_for("tp_calculate")})

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
    """Return Excel bytes matching exactly the HTML email tables (same headers, columns, colors)."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    mon_tag = _tp_mon_tag(month, year)

    # Colors matching the HTML email exactly
    RED_FILL   = PatternFill("solid", fgColor="FFB3B3")
    YEL_FILL   = PatternFill("solid", fgColor="FFE066")
    GRN_FILL   = PatternFill("solid", fgColor="92D492")
    PAN_FILL   = PatternFill("solid", fgColor="D9D9D9")
    HDR_FILL   = PatternFill("solid", fgColor="0A2540")
    PLAIN_FILL = PatternFill("solid", fgColor="FFFFFF")

    # Font colors complementing each background (matching HTML _fg())
    RED_FONT   = Font(color="7B1F1F", size=10)
    YEL_FONT   = Font(color="5C4200", size=10)
    GRN_FONT   = Font(color="1A5C1A", size=10)
    PAN_FONT   = Font(bold=True, color="222222", size=10)
    HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
    TITLE_FONT = Font(bold=True, color="FFFFFF", size=11)
    PLAIN_FONT = Font(color="19263A", size=10)

    CTR   = Alignment(vertical="center", horizontal="center")
    LEFT  = Alignment(vertical="center", horizontal="left")
    WRAP  = Alignment(wrap_text=True, vertical="center", horizontal="center")
    thin  = Side(style="thin", color="9A9A9A")
    BDR   = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _fill_font(pct):
        if pct < 60:  return RED_FILL, RED_FONT
        if pct < 75:  return YEL_FILL, YEL_FONT
        return GRN_FILL, GRN_FONT

    wb = Workbook()

    # ── Location sheet — Sr. no., Exco Location, Plants, Total Qty, Time (min), Avg TP %
    ws_l = wb.active
    ws_l.title = "Location Throughput"
    LOC_HEADS = ["Sr. no.", "Exco Location", "Plants", "Total Qty", "Time (min)", "Avg TP %"]
    ws_l.merge_cells(f"A1:{get_column_letter(len(LOC_HEADS))}1")
    t = ws_l.cell(1, 1, f"Location wise Throughput - {mon_tag}")
    t.fill = HDR_FILL; t.font = TITLE_FONT; t.alignment = CTR
    ws_l.row_dimensions[1].height = 22
    for ci, h in enumerate(LOC_HEADS, 1):
        c = ws_l.cell(2, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    ws_l.row_dimensions[2].height = 28
    srno = 1
    for ri, row in enumerate(location_rows, 3):
        pct = float(row.get("avg_throughput_pct", 0))
        pan = bool(row.get("is_pan_india"))
        pct_fill, pct_font = _fill_font(pct)
        if pan:
            row_fill, row_font = PAN_FILL, PAN_FONT
            srno_val = "—"
        else:
            row_fill, row_font = PLAIN_FILL, PLAIN_FONT
            srno_val = srno; srno += 1
        vals = [srno_val, row.get("exco_location",""), row.get("plant_count",0),
                round(float(row.get("total_quantity",0)), 1),
                round(float(row.get("total_time_min",0)), 1), f"{round(pct)}%"]
        for ci, v in enumerate(vals, 1):
            c = ws_l.cell(ri, ci, v)
            c.border = BDR
            c.alignment = LEFT if ci == 2 else CTR
            if ci == len(LOC_HEADS):          # Avg TP % cell always gets pct color
                c.fill = pct_fill; c.font = Font(bold=pan, color=pct_font.color, size=10)
            else:
                c.fill = row_fill; c.font = row_font
    for ci, w in enumerate([30, 22, 8, 12, 12, 10], 1):
        ws_l.column_dimensions[get_column_letter(ci)].width = w

    # ── Plant sheet — Sr. no., Plant, Exco Location, Business Head, Plant Manager,
    #                 Mixer Cap, Total Qty, Time (min), TP %
    ws_p = wb.create_sheet("Plant Throughput")
    PLT_HEADS = ["Sr. no.", "Plant", "Exco Location", "Business Head",
                 "Plant Manager", "Mixer Cap", "Total Qty", "Time (min)", "TP %"]
    ws_p.merge_cells(f"A1:{get_column_letter(len(PLT_HEADS))}1")
    t2 = ws_p.cell(1, 1, f"Plant Throughput report - {mon_tag}")
    t2.fill = HDR_FILL; t2.font = TITLE_FONT; t2.alignment = CTR
    ws_p.row_dimensions[1].height = 22
    for ci, h in enumerate(PLT_HEADS, 1):
        c = ws_p.cell(2, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    ws_p.row_dimensions[2].height = 28
    for ri, row in enumerate(plant_rows, 3):
        pct  = float(row.get("throughput_pct", 0))
        fill, fnt = _fill_font(pct)
        vals = [ri - 2, row.get("plant_name",""), row.get("exco_location",""),
                row.get("business_head",""), row.get("plant_manager",""),
                row.get("mixer_theo_cap",""),
                round(float(row.get("total_quantity",0)), 1),
                round(float(row.get("total_time_min",0)), 1),
                f"{round(pct)}%"]
        for ci, v in enumerate(vals, 1):
            c = ws_p.cell(ri, ci, v)
            c.fill = fill; c.font = fnt; c.border = BDR
            c.alignment = LEFT if ci == 2 else CTR
    for ci, w in enumerate([30, 40, 16, 18, 18, 10, 11, 11, 8], 1):
        ws_p.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _tp_build_html_tables(plant_rows, location_rows, month, year):
    """Return HTML tables for TP report email.
    width:100% + table-layout:auto lets each column size itself to its content.
    Plant name column wraps; all other columns stay on one line (nowrap).
    """
    mon_tag = _tp_mon_tag(month, year)

    def _bg(pct):
        if pct < 60:  return "#FFB3B3"
        if pct < 75:  return "#FFE066"
        return "#92D492"

    def _fg(bg):
        return {"#FFB3B3": "#7B1F1F", "#FFE066": "#5C4200",
                "#92D492": "#1A5C1A", "#D9D9D9": "#222222"}.get(bg, "#19263A")

    F       = "font-family:Arial,sans-serif;font-size:11px;"
    TBL_LOC = "border-collapse:collapse;width:auto;margin:4px 0 10px"
    TBL_PLT = "border-collapse:collapse;width:auto;margin:4px 0 10px"
    TTL     = (f'style="{F}font-size:12px;font-weight:bold;background:#082B49;color:#fff;'
               f'padding:3px 6px;border:1px solid #7A7A7A;text-align:left"')

    def _th(align="center", wrap=False, w=None):
        ws = "" if wrap else "white-space:nowrap;"
        wd = f"width:{w}px;" if w else ""
        wa = f'width="{w}" ' if w else ""
        return (f'{wa}style="{F}background:#082B49;color:#fff;font-weight:bold;'
                f'padding:3px 6px;border:1px solid #7A7A7A;text-align:{align};'
                f'{ws}{wd}line-height:1.2;vertical-align:middle"')

    def _td(bg, fg, align, bold=False, top_bdr="", wrap=False):
        fw  = "font-weight:bold;" if bold else ""
        top = f"border-top:{top_bdr};" if top_bdr else ""
        ws  = "" if wrap else "white-space:nowrap;"
        return (f'style="{F}padding:2px 5px;border:1px solid #9E9E9E;{top}'
                f'background:{bg};color:{fg};text-align:{align};'
                f'{ws}line-height:1.2;vertical-align:middle;{fw}"')

    # ── Location table ────────────────────────────────────────────────────────
    loc_body = ""
    srno = 1
    for r in location_rows:
        pct    = float(r.get("avg_throughput_pct", 0))
        pan    = bool(r.get("is_pan_india"))
        pct_bg = _bg(pct); pct_fg = _fg(pct_bg)
        qty    = round(float(r.get("total_quantity", 0)), 1)
        mins   = round(float(r.get("total_time_min", 0)), 1)
        plants = r.get("plant_count", 0)
        loc    = r.get("exco_location", "")
        if pan:
            bg = "#D9D9D9"; fg = "#222222"
            loc_body += (
                f'<tr>'
                f'<td {_td(bg, fg, "center", bold=True, top_bdr="2px solid #555")}>—</td>'
                f'<td {_td(bg, fg, "left",   bold=True, top_bdr="2px solid #555")}>&#127988; PAN India</td>'
                f'<td {_td(bg, fg, "center", bold=True, top_bdr="2px solid #555")}>{plants}</td>'
                f'<td {_td(bg, fg, "right",  bold=True, top_bdr="2px solid #555")}>{qty}</td>'
                f'<td {_td(bg, fg, "right",  bold=True, top_bdr="2px solid #555")}>{mins}</td>'
                f'<td {_td(pct_bg, pct_fg, "center", bold=True, top_bdr="2px solid #555")}>{round(pct)}%</td>'
                f'</tr>'
            )
        else:
            bg = "#ffffff"; fg = "#19263A"
            loc_body += (
                f'<tr>'
                f'<td {_td(bg, fg, "center")}>{srno}</td>'
                f'<td {_td(bg, fg, "left")}>{loc}</td>'
                f'<td {_td(bg, fg, "center")}>{plants}</td>'
                f'<td {_td(bg, fg, "right")}>{qty}</td>'
                f'<td {_td(bg, fg, "right")}>{mins}</td>'
                f'<td {_td(bg, fg, "center", bold=True)}>{round(pct)}%</td>'
                f'</tr>'
            )
            srno += 1

    loc_html = (
        f'<table cellpadding="0" cellspacing="0" style="{TBL_LOC}">'
        f'<tr><td colspan="6" {TTL}>Location wise Throughput - {mon_tag}</td></tr>'
        f'<tr>'
        f'<th {_th("center", w=45)}>Sr. no.</th>'
        f'<th {_th("left")}>Exco Location</th>'
        f'<th {_th("center")}>Plants</th>'
        f'<th {_th("right")}>Total Qty</th>'
        f'<th {_th("right")}>Time (min)</th>'
        f'<th {_th("center")}>Avg TP %</th>'
        f'</tr>{loc_body}</table>'
    )

    # ── Plant table ───────────────────────────────────────────────────────────
    plant_body = ""
    for i, r in enumerate(plant_rows, 1):
        pct  = float(r.get("throughput_pct", 0))
        bg   = _bg(pct); fg = _fg(bg)
        name = str(r.get("plant_name", ""))
        bh   = str(r.get("business_head", ""))
        pm   = str(r.get("plant_manager", ""))
        cap  = str(r.get("mixer_theo_cap", ""))
        qty  = round(float(r.get("total_quantity", 0)), 1)
        mins = round(float(r.get("total_time_min", 0)), 1)
        plant_body += (
            f'<tr>'
            f'<td {_td(bg, fg, "center")}>{i}</td>'
            f'<td {_td(bg, fg, "left", wrap=True)}>{name}</td>'
            f'<td {_td(bg, fg, "left")}>{bh}</td>'
            f'<td {_td(bg, fg, "left")}>{pm}</td>'
            f'<td {_td(bg, fg, "right")}>{cap}</td>'
            f'<td {_td(bg, fg, "right")}>{qty}</td>'
            f'<td {_td(bg, fg, "right")}>{mins}</td>'
            f'<td {_td(bg, fg, "center", bold=True)}>{round(pct)}%</td>'
            f'</tr>'
        )

    plant_html = (
        f'<table cellpadding="0" cellspacing="0" style="{TBL_PLT}">'
        f'<tr><td colspan="8" {TTL}>Plant Throughput Report - {mon_tag}</td></tr>'
        f'<tr>'
        f'<th {_th("center", w=45)}>Sr. no.</th>'
        f'<th {_th("center", wrap=True)}>Plant</th>'
        f'<th {_th("center")}>Business Head</th>'
        f'<th {_th("center")}>Plant Manager</th>'
        f'<th {_th("center")}>Mixer Cap</th>'
        f'<th {_th("center")}>Total Qty</th>'
        f'<th {_th("center")}>Time (min)</th>'
        f'<th {_th("center")}>TP %</th>'
        f'</tr>{plant_body}</table>'
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

    # Build color-coded Excel attachment
    excel_bytes = _tp_build_excel(plant_rows, location_rows, month, year)
    fname = f"RDC_TP_{year}_{month:02d}.xlsx"

    # Compact HTML email body — same style as BTRTP
    _body_font = "font-family:Arial,Calibri,sans-serif;font-size:12px;color:#000000;"
    tables_html = _tp_build_html_tables(plant_rows, location_rows, month, year)
    html_body = (
        f'<html><body style="margin:0;padding:8px 10px;{_body_font}">'
        f'<p style="margin:0 0 3px 0;">Dear Team,</p>'
        f'<p style="margin:0 0 3px 0;">Please find attached the Plant Throughput Report for {month_name} {year}.</p>'
        f'<p style="margin:0 0 6px 0;">Regards,<br>RDC Operations</p>'
        f'{tables_html}'
        f'</body></html>'
    )

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


# ═══════════════════════════════════════════════════════════════════════════
# RDC-BTRTP MODULE — Batcher-wise Throughput Calculator
# ═══════════════════════════════════════════════════════════════════════════

def _btrtp_ctx():
    """Base context dict for every BTRTP page."""
    return dict(
        active_page="",
        bg_auto=bg_auto(), bg_animate=bg_animate(), bg_theme=bg_theme(),
    )


def _btrtp_band(pct):
    """Throughput colour band for a percentage (same thresholds as TP)."""
    if pct < 60:
        return "red"
    if pct < 75:
        return "yellow"
    return "green"


@app.route("/btrtp")
def page_btrtp():
    return redirect(url_for("btrtp_dashboard"))


# ── Dashboard ─────────────────────────────────────────────────────────────────
@app.route("/btrtp/dashboard")
def btrtp_dashboard():
    all_counts = database.get_table_counts()
    counts = {
        "btrtp_oracle_data": all_counts.get("btrtp_oracle_data", 0),
        "btrtp_master_data": all_counts.get("btrtp_master_data", 0),
        "btrtp_results":     all_counts.get("btrtp_results",     0),
        "tp_plant_data":     all_counts.get("tp_plant_data",     0),
    }
    last_sync = google_sheets.get_btrtp_last_sync_info()
    ora_ready = oracle_connector.is_configured()
    ctx = _btrtp_ctx()
    ctx["active_page"] = "dashboard"
    return render_template("btrtp_dashboard.html", counts=counts,
                           last_sync=last_sync, ora_ready=ora_ready, **ctx)


# ── Data Uploader ─────────────────────────────────────────────────────────────
@app.route("/btrtp/data-uploader")
def btrtp_data_uploader():
    ora_ready    = oracle_connector.is_configured()
    last_master  = google_sheets.get_btrtp_last_sync_info()
    last_plant   = google_sheets.get_tp_last_sync_info()

    # BT Master sheet settings
    master_sheet_id  = database.get_module_setting("btrtp", "gsheet_id",
                                                    database.get_setting("gsheet_id", ""))
    master_worksheet = database.get_module_setting("btrtp", "gsheet_worksheet", "BT Master Data")

    # Plant Data sheet settings (shared reference with TP)
    plant_sheet_id   = database.get_module_setting("tp", "gsheet_id",
                                                    database.get_setting("gsheet_id", ""))
    plant_worksheet  = database.get_module_setting("tp", "gsheet_worksheet", "Plant Data for TP")

    all_counts   = database.get_table_counts()
    ora_counts   = all_counts.get("btrtp_oracle_data", 0)
    master_count = all_counts.get("btrtp_master_data", 0)
    plant_count  = all_counts.get("tp_plant_data", 0)

    ctx = _btrtp_ctx()
    ctx["active_page"] = "data_uploader"
    return render_template("btrtp_data_uploader.html",
                           ora_ready=ora_ready,
                           last_master=last_master, last_plant=last_plant,
                           master_sheet_id=master_sheet_id, master_worksheet=master_worksheet,
                           plant_sheet_id=plant_sheet_id, plant_worksheet=plant_worksheet,
                           ora_counts=ora_counts, master_count=master_count,
                           plant_count=plant_count, **ctx)


@app.route("/btrtp/action/fetch-oracle", methods=["POST"])
def btrtp_fetch_oracle():
    from_date = request.form.get("from_date", "")
    to_date   = request.form.get("to_date", "")
    if not from_date or not to_date:
        flash("Please select both From and To dates.", "error")
        return jsonify({"ok": False, "redirect": url_for("btrtp_data_uploader")})
    try:
        _set_progress(20, "Fetching BTRTP data from Oracle…")
        raw_df, ora_warnings = oracle_connector.fetch_btrtp_data(from_date, to_date)
        _set_progress(55, "Parsing batcher records…")
        parsed, skip_log     = btrtp_calculator.parse_btrtp_oracle_df(raw_df)
        _set_progress(75, "Saving to database…")
        oracle_connector.save_btrtp_oracle_data(parsed, replace=True)
        _set_progress(90, "Cleaning old records…")
        database.purge_old_oracle_data()
        _mss("btrtp", "skip_log",  skip_log)
        _mss("btrtp", "ora_from",  from_date)
        _mss("btrtp", "ora_to",    to_date)
        for w in ora_warnings:
            flash(w, "warning")
        flash(f"✅ {len(parsed)} rows fetched & saved ({len(skip_log)} skipped).", "success")
    except Exception as exc:
        flash(f"Oracle error: {exc}", "error")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("btrtp_data_uploader")})


@app.route("/btrtp/action/sync-master", methods=["POST"])
def btrtp_sync_master():
    sheet_id  = (request.form.get("sheet_id", "").strip()
                 or database.get_module_setting("btrtp", "gsheet_id",
                                                database.get_setting("gsheet_id", "")))
    worksheet = request.form.get("worksheet", "BT Master Data").strip()
    _set_progress(20, "Syncing BT Master Data from Google Sheets…")
    result = google_sheets.sync_btrtp_master_data(sheet_id, worksheet)
    _set_progress(85, "Saving settings…")
    if result["error"]:
        flash(f"Sync failed: {result['error']}", "error")
    elif result["rows_synced"] == 0:
        flash("⚠️ Sync completed but 0 batcher rows were saved. "
              "Check that the sheet tab name, 'Batcher ID' and 'Batcher Name' columns exist and have data.",
              "warning")
    else:
        database.set_module_setting("btrtp", "gsheet_id",
                                    google_sheets.extract_sheet_id(sheet_id))
        database.set_module_setting("btrtp", "gsheet_worksheet", worksheet)
        flash(f"✅ {result['rows_synced']} batcher rows synced from Google Sheets ({result['mode']} mode).",
              "success")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("btrtp_data_uploader")})


@app.route("/btrtp/action/sync-plant", methods=["POST"])
def btrtp_sync_plant():
    sheet_id  = (request.form.get("sheet_id", "").strip()
                 or database.get_module_setting("tp", "gsheet_id",
                                                database.get_setting("gsheet_id", "")))
    worksheet = request.form.get("worksheet", "Plant Data for TP").strip()
    _set_progress(20, "Syncing Plant Data from Google Sheets…")
    result = google_sheets.sync_tp_plant_data(sheet_id, worksheet)
    _set_progress(85, "Saving settings…")
    if result["error"]:
        flash(f"Sync failed: {result['error']}", "error")
    else:
        database.set_module_setting("tp", "gsheet_id", google_sheets.extract_sheet_id(sheet_id))
        database.set_module_setting("tp", "gsheet_worksheet", worksheet)
        flash(f"✅ {result['rows_synced']} plant rows synced from Google Sheets ({result['mode']} mode).",
              "success")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("btrtp_data_uploader")})


# ── Calculate ─────────────────────────────────────────────────────────────────
@app.route("/btrtp/calculate", methods=["GET", "POST"])
def btrtp_calculate():
    batcher_rows = _ms("btrtp", "batcher_rows", [])
    warnings     = _ms("btrtp", "calc_warnings", [])
    ran          = _ms("btrtp", "calc_ran", False)

    if request.method == "POST":
        from_s = request.form.get("from_date", "")
        to_s   = request.form.get("to_date", "")
        try:
            fd = _date.fromisoformat(from_s)
            td = _date.fromisoformat(to_s)
        except (ValueError, TypeError):
            flash("Please select valid From and To dates.", "error")
            return jsonify({"ok": False, "redirect": url_for("btrtp_calculate")})
        if fd > td:
            flash("'From Date' must be on or before 'To Date'.", "error")
            return jsonify({"ok": False, "redirect": url_for("btrtp_calculate")})

        month, year = fd.month, fd.year
        _set_progress(20, "Loading Oracle data…")
        batcher_rows, warnings = btrtp_calculator.run_btrtp_calculation(
            month, year, from_date=str(fd), to_date=str(td))
        _set_progress(80, "Saving results…")
        _mss("btrtp", "batcher_rows",  batcher_rows)
        _mss("btrtp", "calc_warnings", warnings)
        _mss("btrtp", "calc_month",    month)
        _mss("btrtp", "calc_year",     year)
        _mss("btrtp", "calc_from",     str(fd))
        _mss("btrtp", "calc_to",       str(td))
        _mss("btrtp", "calc_ran",      True)
        ran = True

        if batcher_rows:
            btrtp_calculator.save_btrtp_results(batcher_rows, month, year)
            flash(f"✅ Calculated batcher throughput for {fd} → {td} — "
                  f"{len(batcher_rows)} batcher-plant rows.", "success")
        else:
            flash("No results produced — check data and warnings.", "warning")
        _set_progress(100, "Complete")
        return jsonify({"ok": True, "redirect": url_for("btrtp_calculate")})

    today = _date.today()
    plant_groups = _group_btrtp_by_plant(batcher_rows)
    ctx = _btrtp_ctx()
    ctx["active_page"] = "calculate"
    return render_template("btrtp_calculate.html",
                           batcher_rows=batcher_rows, plant_groups=plant_groups,
                           warnings=warnings, ran=ran,
                           default_from=_ms("btrtp", "calc_from", str(today.replace(day=1))),
                           default_to=_ms("btrtp", "calc_to", str(today)),
                           **ctx)


# ── Reports ───────────────────────────────────────────────────────────────────
def _apply_btrtp_filters(rows, excos, bheads, plants, batchers, band, search):
    """Filter batcher rows by Exco, Business Head, Plant, Batcher, band, and search."""
    out = []
    s = (search or "").strip().lower()
    for r in rows:
        if excos   and r.get("exco_location") not in excos:
            continue
        if bheads  and r.get("business_head") not in bheads:
            continue
        if plants  and r.get("plant_name") not in plants:
            continue
        if batchers and r.get("batcher_name") not in batchers:
            continue
        if band and band != "All" and _btrtp_band(r.get("throughput_pct", 0)) != band:
            continue
        if s and s not in str(r.get("batcher_id", "")).lower() \
             and s not in str(r.get("batcher_name", "")).lower() \
             and s not in str(r.get("plant_name", "")).lower():
            continue
        out.append(r)
    return out


@app.route("/btrtp/reports")
def btrtp_reports():
    today = _date.today()

    from_s = request.args.get("from_date",
                              _ms("btrtp", "calc_from", str(today.replace(day=1))))
    to_s   = request.args.get("to_date",
                              _ms("btrtp", "calc_to", str(today)))
    try:
        fd = _date.fromisoformat(from_s)
        td = _date.fromisoformat(to_s)
    except ValueError:
        fd, td = today.replace(day=1), today

    # Re-run calculation if user picks a date range directly from Reports
    if "from_date" in request.args:
        month, year = fd.month, fd.year
        all_rows, calc_warns = btrtp_calculator.run_btrtp_calculation(
            month, year, from_date=str(fd), to_date=str(td))
        _mss("btrtp", "batcher_rows",  all_rows)
        _mss("btrtp", "calc_warnings", calc_warns)
        _mss("btrtp", "calc_from",  str(fd))
        _mss("btrtp", "calc_to",    str(td))
        _mss("btrtp", "calc_month", month)
        _mss("btrtp", "calc_year",  year)
    else:
        all_rows = _ms("btrtp", "batcher_rows", [])

    excos   = request.args.getlist("exco")
    bheads  = request.args.getlist("bhead")
    plants  = request.args.getlist("plant")
    batchers = request.args.getlist("batcher")
    band    = request.args.get("band", "All")
    search  = request.args.get("search", "")

    unique_excos    = sorted({r.get("exco_location", "") for r in all_rows if r.get("exco_location")})
    unique_bheads   = sorted({r.get("business_head", "")  for r in all_rows if r.get("business_head")})
    unique_plants   = sorted({r.get("plant_name", "")     for r in all_rows if r.get("plant_name")})
    unique_batchers = sorted({r.get("batcher_name", "")   for r in all_rows if r.get("batcher_name")})

    batcher_rows = _apply_btrtp_filters(all_rows, excos, bheads, plants, batchers, band, search)

    _mss("btrtp", "report_rows", batcher_rows)

    smtp        = email_helper.get_smtp_config()
    email_ready = email_helper.is_configured()

    _elog = database.read_table("email_log", order_by="id DESC")
    if not _elog.empty and "report_file_name" in _elog.columns:
        _elog = _elog[_elog["report_file_name"].str.startswith("RDC_BTRTP_", na=False)]
    btrtp_email_log = _records(_elog.drop(columns=["id"], errors="ignore").head(20)) \
        if not _elog.empty else []

    if (fd.year, fd.month) == (td.year, td.month):
        import calendar
        month_label = f"{calendar.month_name[fd.month]} {fd.year}"
    else:
        month_label = f"{fd:%d %b %Y} → {td:%d %b %Y}"

    default_to      = database.get_module_setting("btrtp", "email_default_to", "") or smtp.get("default_to", "")
    default_cc      = database.get_module_setting("btrtp", "email_default_cc", "") or smtp.get("default_cc", "")
    default_subject = (database.get_module_setting("btrtp", "email_default_subject", "")
                       or f"RDC-BTRTP Batcher Throughput Report — {month_label}")
    default_body    = (database.get_module_setting("btrtp", "email_default_body", "")
                       or f"Dear Team,\n\nPlease find attached the Batcher Throughput "
                          f"Report for {month_label}.\n\nRegards,\nRDC Operations")

    plant_groups = _group_btrtp_by_plant(batcher_rows)
    ctx = _btrtp_ctx()
    ctx["active_page"] = "reports"
    return render_template("btrtp_reports.html",
                           batcher_rows=batcher_rows, plant_groups=plant_groups,
                           total_rows=len(all_rows),
                           from_date=str(fd), to_date=str(td),
                           unique_excos=unique_excos, unique_bheads=unique_bheads,
                           unique_plants=unique_plants, unique_batchers=unique_batchers,
                           excos=excos, bheads=bheads, plants=plants, batchers=batchers,
                           band=band, search=search,
                           month_label=month_label, email_cfg=smtp,
                           email_ready=email_ready,
                           default_to=default_to, default_cc=default_cc,
                           default_subject=default_subject, default_body=default_body,
                           btrtp_email_log=btrtp_email_log, **ctx)


def _group_btrtp_by_plant(batcher_rows: list) -> list:
    """
    Group flat batcher rows by plant (lookup_code), return list of plant-group dicts.
    Each group: {plant_name, lookup_code, mixer_label, exco_location, business_head,
                 plant_manager, mixer_theo_cap, rows, avg_tp, total_qty, total_time_hrs}
    Rows within each group sorted by TP% descending.
    """
    groups: dict = {}
    for r in batcher_rows:
        key = r["lookup_code"]
        if key not in groups:
            parts = key.split("_", 1)
            mixer_label = parts[1] if len(parts) > 1 else ""
            groups[key] = {
                "plant_name":     r["plant_name"],
                "lookup_code":    key,
                "mixer_label":    mixer_label,
                "exco_location":  r["exco_location"],
                "business_head":  r["business_head"],
                "plant_manager":  r["plant_manager"],
                "mixer_theo_cap": r["mixer_theo_cap"],
                "rows": [],
            }
        groups[key]["rows"].append(r)

    # Compute aggregates and sort rows within each group by TP% descending
    result = sorted(groups.values(), key=lambda g: (g["plant_name"], g["mixer_label"]))
    for g in result:
        g["rows"].sort(key=lambda r: -r["throughput_pct"])
        tps   = [r["throughput_pct"] for r in g["rows"]]
        g["avg_tp"]        = round(sum(tps) / len(tps), 1) if tps else 0
        g["total_qty"]     = round(sum(r["total_quantity"]  for r in g["rows"]), 1)
        g["total_time_hrs"]= round(sum(r["total_time_hrs"]  for r in g["rows"]), 2)
        g["batch_count"]   = sum(r["batch_count"] for r in g["rows"])
    return result


def _btrtp_mon_tag(month, year):
    import calendar as _cal
    return f"{_cal.month_abbr[month]}'{str(year)[2:]}"


def _btrtp_build_excel(batcher_rows, month, year):
    """Build colour-coded Excel bytes for BTRTP batcher results."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    mon_tag = _btrtp_mon_tag(month, year)

    RED_FILL   = PatternFill("solid", fgColor="FFB3B3")
    YEL_FILL   = PatternFill("solid", fgColor="FFE066")
    GRN_FILL   = PatternFill("solid", fgColor="92D492")
    HDR_FILL   = PatternFill("solid", fgColor="0A2540")
    PLAIN_FILL = PatternFill("solid", fgColor="FFFFFF")

    RED_FONT   = Font(color="7B1F1F", size=10)
    YEL_FONT   = Font(color="5C4200", size=10)
    GRN_FONT   = Font(color="1A5C1A", size=10)
    HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
    TITLE_FONT = Font(bold=True, color="FFFFFF", size=11)
    PLAIN_FONT = Font(color="19263A", size=10)

    CTR  = Alignment(vertical="center", horizontal="center")
    LEFT = Alignment(vertical="center", horizontal="left")
    WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
    thin = Side(style="thin", color="9A9A9A")
    BDR  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _fill_font(pct):
        if pct < 60:  return RED_FILL, RED_FONT
        if pct < 75:  return YEL_FILL, YEL_FONT
        return GRN_FILL, GRN_FONT

    wb = Workbook()
    ws = wb.active
    ws.title = "Batcher Throughput"

    HEADS = ["Sr. no.", "Batcher ID", "Batcher Name", "Plant", "Business Head",
             "Plant Manager", "Mixer Cap", "Total Qty", "Time (hr)", "TP %", "Batches"]

    ws.merge_cells(f"A1:{get_column_letter(len(HEADS))}1")
    t = ws.cell(1, 1, f"Batcher wise Throughput - {mon_tag}")
    t.fill = HDR_FILL; t.font = TITLE_FONT; t.alignment = CTR
    ws.row_dimensions[1].height = 22

    for ci, h in enumerate(HEADS, 1):
        c = ws.cell(2, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    ws.row_dimensions[2].height = 28

    srno = 1
    for ri, row in enumerate(batcher_rows, 3):
        pct        = float(row.get("throughput_pct", 0))
        row_fill, row_font = _fill_font(pct)
        vals = [
            srno,
            row.get("batcher_id", ""),
            row.get("batcher_name", ""),
            row.get("plant_name", ""),
            row.get("business_head", ""),
            row.get("plant_manager", ""),
            round(float(row.get("mixer_theo_cap", 0)), 1),
            round(float(row.get("total_quantity", 0)), 1),
            round(float(row.get("total_time_hrs", 0)), 2),
            f"{round(pct)}%",
            int(row.get("batch_count", 0)),
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v)
            c.border = BDR
            c.fill   = row_fill
            c.font   = row_font
            c.alignment = LEFT if ci in (2, 3, 4, 5, 6) else CTR
        srno += 1

    for ci, w in enumerate([8, 14, 20, 22, 18, 18, 10, 10, 10, 8, 8], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


@app.route("/btrtp/download-excel")
def btrtp_download_excel():
    rows  = _ms("btrtp", "report_rows", _ms("btrtp", "batcher_rows", []))
    month = _ms("btrtp", "calc_month", _date.today().month)
    year  = _ms("btrtp", "calc_year",  _date.today().year)
    if not rows:
        flash("No data to download — run Calculate first.", "warning")
        return redirect(url_for("btrtp_reports"))
    excel_bytes = _btrtp_build_excel(rows, month, year)
    fname = f"RDC_BTRTP_{year}_{month:02d}.xlsx"
    return send_file(io.BytesIO(excel_bytes), as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/btrtp/download-csv")
def btrtp_download_csv():
    rows  = _ms("btrtp", "report_rows", _ms("btrtp", "batcher_rows", []))
    month = _ms("btrtp", "calc_month", _date.today().month)
    year  = _ms("btrtp", "calc_year",  _date.today().year)
    if not rows:
        flash("No data to download — run Calculate first.", "warning")
        return redirect(url_for("btrtp_reports"))
    df = pd.DataFrame(rows)
    buf = io.BytesIO(df.to_csv(index=False).encode("utf-8"))
    return send_file(buf, as_attachment=True,
                     download_name=f"RDC_BTRTP_{year}_{month:02d}.csv",
                     mimetype="text/csv")


@app.route("/btrtp/action/send-email", methods=["POST"])
def btrtp_send_email():
    rows  = _ms("btrtp", "report_rows", _ms("btrtp", "batcher_rows", []))
    month = _ms("btrtp", "calc_month", _date.today().month)
    year  = _ms("btrtp", "calc_year",  _date.today().year)
    to_addr = request.form.get("to", "").strip()
    cc_addr = request.form.get("cc", "").strip()
    subject = request.form.get("subject", "").strip()
    body    = request.form.get("body", "").strip()

    if not rows:
        flash("No results to email — run Calculate first.", "warning")
        return redirect(url_for("btrtp_reports"))
    if not to_addr:
        flash("Please enter at least one To email address.", "error")
        return redirect(url_for("btrtp_reports"))

    import calendar as _cal
    month_name = _cal.month_name[month]
    mon_tag = _btrtp_mon_tag(month, year)
    if not subject:
        subject = f"RDC-BTRTP Batcher Throughput Report — {month_name} {year}"
    if not body:
        body = f"Dear Team,\n\nPlease find the Batcher Throughput Report for {month_name} {year} below."

    excel_bytes = _btrtp_build_excel(rows, month, year)
    fname = f"RDC_BTRTP_{year}_{month:02d}.xlsx"

    # Compact HTML email body with inline CSS (Gmail / Outlook safe)
    _TH = ('style="background-color:#082B49;color:#ffffff;font-weight:bold;'
           'text-align:center;padding:3px 6px;border:1px solid #808080;'
           'font-size:11px;font-family:Arial,sans-serif;line-height:1.2;white-space:nowrap;"')
    def _td_b(bg, fg, align="center"):
        return (f'style="background-color:{bg};color:{fg};padding:2px 5px;'
                f'border:1px solid #999999;text-align:{align};'
                f'font-size:11px;font-family:Arial,sans-serif;line-height:1.2;white-space:nowrap;"')
    def _row_color(pct):
        if pct < 60: return "#FFB3B3", "#7B1F1F"
        if pct < 75: return "#FFE066", "#5C4200"
        return "#92D492", "#1A5C1A"

    head_row = (f'<tr><th {_TH}>Sr.</th><th {_TH}>Batcher ID</th>'
                f'<th {_TH}>Batcher Name</th><th {_TH}>Plant</th>'
                f'<th {_TH}>Mixer Cap</th><th {_TH}>Total Qty</th>'
                f'<th {_TH}>Time (hr)</th><th {_TH}>TP %</th><th {_TH}>Batches</th></tr>')
    body_rows = ""
    for i, r in enumerate(rows, 1):
        pct = float(r.get("throughput_pct", 0))
        bg, fg = _row_color(pct)
        body_rows += (
            f'<tr>'
            f'<td {_td_b(bg,fg,"center")}>{i}</td>'
            f'<td {_td_b(bg,fg,"left")}>{r.get("batcher_id","")}</td>'
            f'<td {_td_b(bg,fg,"left")}>{r.get("batcher_name","")}</td>'
            f'<td {_td_b(bg,fg,"left")}>{r.get("plant_name","")}</td>'
            f'<td {_td_b(bg,fg,"center")}>{round(float(r.get("mixer_theo_cap",0)),1)}</td>'
            f'<td {_td_b(bg,fg,"center")}>{round(float(r.get("total_quantity",0)),1)}</td>'
            f'<td {_td_b(bg,fg,"center")}>{round(float(r.get("total_time_hrs",0)),2)}</td>'
            f'<td {_td_b(bg,fg,"center")}><b>{round(pct)}%</b></td>'
            f'<td {_td_b(bg,fg,"center")}>{int(r.get("batch_count",0))}</td>'
            f'</tr>'
        )
    _tbl_css = ('border-collapse:collapse;width:auto;font-family:Arial,sans-serif;'
                'font-size:11px;margin:0;padding:0;')
    _body_font = 'font-family:Arial,Calibri,sans-serif;font-size:12px;color:#000000;'
    html_body = (
        f'<html><body style="margin:0;padding:8px 10px;{_body_font}">'
        f'<p style="margin:0 0 3px 0;">Dear Team,</p>'
        f'<p style="margin:0 0 3px 0;">Please find attached the Batcher Throughput Report for {month_name} {year}.</p>'
        f'<p style="margin:0 0 6px 0;">Regards,<br>RDC Operations</p>'
        f'<p style="margin:0 0 4px 0;font-size:13px;font-weight:bold;color:#082B49;">'
        f'Batcher Throughput Report - {mon_tag}</p>'
        f'<table style="{_tbl_css}">{head_row}{body_rows}</table>'
        f'</body></html>'
    )

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
    return redirect(url_for("btrtp_reports"))


# ── Validation ────────────────────────────────────────────────────────────────
@app.route("/btrtp/validation")
def btrtp_validation():
    skip_log      = _ms("btrtp", "skip_log", [])
    calc_warnings = _ms("btrtp", "calc_warnings", [])
    ctx = _btrtp_ctx()
    ctx["active_page"] = "validation"
    return render_template("btrtp_validation.html",
                           skip_log=skip_log, calc_warnings=calc_warnings, **ctx)


# ── Settings ──────────────────────────────────────────────────────────────────
@app.route("/btrtp/settings", methods=["GET"])
def btrtp_settings():
    # BT Master sheet
    master_sheet_id  = database.get_module_setting("btrtp", "gsheet_id",
                                                    database.get_setting("gsheet_id", ""))
    master_worksheet = database.get_module_setting("btrtp", "gsheet_worksheet", "BT Master Data")
    # Plant Data sheet (shared with TP)
    plant_sheet_id   = database.get_module_setting("tp", "gsheet_id",
                                                    database.get_setting("gsheet_id", ""))
    plant_worksheet  = database.get_module_setting("tp", "gsheet_worksheet", "Plant Data for TP")

    batcher_col = database.get_module_setting("btrtp", "oracle_batcher_col", "CREATED_BY")
    smtp        = email_helper.get_smtp_config()
    email_configured = bool(smtp.get("host") and smtp.get("sender"))
    ora_configured   = oracle_connector.is_configured()
    last_master = google_sheets.get_btrtp_last_sync_info()
    last_plant  = google_sheets.get_tp_last_sync_info()
    btrtp_email_to      = database.get_module_setting("btrtp", "email_default_to", "")
    btrtp_email_cc      = database.get_module_setting("btrtp", "email_default_cc", "")
    btrtp_email_subject = database.get_module_setting("btrtp", "email_default_subject", "")
    btrtp_email_body    = database.get_module_setting("btrtp", "email_default_body", "")
    ctx = _btrtp_ctx()
    ctx["active_page"] = "settings"
    return render_template("btrtp_settings.html",
                           master_sheet_id=master_sheet_id, master_worksheet=master_worksheet,
                           plant_sheet_id=plant_sheet_id, plant_worksheet=plant_worksheet,
                           batcher_col=batcher_col,
                           smtp=smtp, email_configured=email_configured,
                           ora_configured=ora_configured,
                           last_master=last_master, last_plant=last_plant,
                           btrtp_email_to=btrtp_email_to, btrtp_email_cc=btrtp_email_cc,
                           btrtp_email_subject=btrtp_email_subject,
                           btrtp_email_body=btrtp_email_body, **ctx)


@app.route("/btrtp/settings/save-oracle-cols", methods=["POST"])
def btrtp_save_oracle_cols():
    database.set_module_setting("btrtp", "oracle_batcher_col",
                                request.form.get("batcher_col", "CREATED_BY").strip())
    flash("Oracle column settings saved.", "success")
    return redirect(url_for("btrtp_settings", m="oracle-cols"))


@app.route("/btrtp/settings/save-master-sheet", methods=["POST"])
def btrtp_save_master_sheet():
    worksheet = request.form.get("worksheet", "BT Master Data").strip()
    sheet_id  = request.form.get("sheet_id", "").strip()
    database.set_module_setting("btrtp", "gsheet_worksheet", worksheet)
    if sheet_id:
        database.set_module_setting("btrtp", "gsheet_id",
                                    google_sheets.extract_sheet_id(sheet_id))
    flash("BT Master Data sheet settings saved.", "success")
    return redirect(url_for("btrtp_settings", m="master-sheet"))


@app.route("/btrtp/settings/save-plant-sheet", methods=["POST"])
def btrtp_save_plant_sheet():
    worksheet = request.form.get("worksheet", "Plant Data for TP").strip()
    sheet_id  = request.form.get("sheet_id", "").strip()
    database.set_module_setting("tp", "gsheet_worksheet", worksheet)
    if sheet_id:
        database.set_module_setting("tp", "gsheet_id",
                                    google_sheets.extract_sheet_id(sheet_id))
    flash("Plant Data sheet settings saved.", "success")
    return redirect(url_for("btrtp_settings", m="plant-sheet"))


@app.route("/btrtp/settings/save-email-defaults", methods=["POST"])
def btrtp_save_email_defaults():
    database.set_module_settings_bulk("btrtp", {
        "email_default_to":      request.form.get("default_to", "").strip(),
        "email_default_cc":      request.form.get("default_cc", "").strip(),
        "email_default_subject": request.form.get("default_subject", "").strip(),
        "email_default_body":    request.form.get("default_body", "").strip(),
    })
    flash("BTRTP email defaults saved.", "success")
    return redirect(url_for("btrtp_settings", m="email"))


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
    maint_cols, maint_months_data, maint_avg, maint_above = [], {}, 0, 0
    if not maint_df.empty:
        maint_avg   = round(maint_df["ytd_maintenance_cost"].mean(), 2)
        maint_above = int((maint_df["ytd_maintenance_cost"] > config.MAINTENANCE_COST_THRESHOLD).sum())
        # Group rows by (year, month) for the template
        import calendar as _cal
        for _, r in maint_df.iterrows():
            m, y = int(r.get("month") or 0), int(r.get("year") or 0)
            key = (y, m)
            if key not in maint_months_data:
                label = f"{_cal.month_name[m]} {y}" if m else "Unassigned"
                maint_months_data[key] = {"label": label, "month": m, "year": y, "rows": []}
            maint_months_data[key]["rows"].append({
                "plant_code":           r["plant_code"],
                "ytd_maintenance_cost": r["ytd_maintenance_cost"],
                "uploaded_at":          r.get("uploaded_at", ""),
            })
    maint_month_groups = sorted(maint_months_data.values(),
                                key=lambda g: (g["year"], g["month"]), reverse=True)
    maint_count = len(maint_df) if not maint_df.empty else 0
    # Set of (month, year) tuples already uploaded — used to lock the upload form
    maint_uploaded_keys = {(g["month"], g["year"]) for g in maint_month_groups
                           if g["month"] and g["year"]}
    has_unassigned = any(g["month"] == 0 for g in maint_month_groups)

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
                           maint_month_groups=maint_month_groups,
                           maint_avg=maint_avg, maint_above=maint_above,
                           maint_count=maint_count,
                           current_month=_date.today().month,
                           current_year=_date.today().year,
                           maint_uploaded_keys=maint_uploaded_keys,
                           has_unassigned=has_unassigned,
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

    import calendar as _cal
    _maint_months = database.get_maintenance_months()
    if _maint_months:
        _mm, _my = _maint_months[0]
        maint_month_label = f"{_cal.month_name[_mm]} {_my}" if _mm else "Unassigned"
        _lc_month = int(last_calc.get("month") or 0)
        _lc_year  = int(last_calc.get("year")  or 0)
        maint_mismatch = bool(_lc_month and _lc_year and (_mm != _lc_month or _my != _lc_year))
    else:
        maint_month_label = None
        maint_mismatch    = False

    all_waivers = database.get_all_waivers()
    import calendar as _cal2
    waiver_month_opts = [(m, _cal2.month_name[m]) for m in range(1, 13)]
    # Employee list for searchable waiver LOV
    _emp_df = database.read_table_limited("master_data", order_by="employee_code", limit=100000)
    waiver_employees = (
        [{"code": str(r["employee_code"]), "name": str(r.get("employee_name", ""))}
         for _, r in _emp_df.iterrows()]
        if not _emp_df.empty else []
    )

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
                           total_inc=total_inc, total_ded=total_ded,
                           maint_month_label=maint_month_label,
                           maint_mismatch=maint_mismatch,
                           all_waivers=all_waivers,
                           waiver_month_opts=waiver_month_opts,
                           waiver_employees=waiver_employees,
                           current_year=_date.today().year)


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

    import calendar as _cal
    _maint_months = database.get_maintenance_months()
    if _maint_months:
        _mm, _my = _maint_months[0]
        maint_month_label = f"{_cal.month_name[_mm]} {_my}" if _mm else "Unassigned"
        # Only flag mismatch after user has explicitly chosen a date range
        maint_mismatch = has_params and (_mm != from_date.month or _my != from_date.year)
    else:
        maint_month_label = None
        maint_mismatch    = False

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
                           email_log=email_log,
                           maint_month_label=maint_month_label,
                           maint_mismatch=maint_mismatch)


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

    # Human-readable labels and fix instructions for each error type
    _ERR_META = {
        "BLANK_FIELD":          ("🔴 Blank Field",          "A required column is empty. Fill it in the Google Sheet and re-sync."),
        "INVALID_CATEGORY":     ("🔴 Invalid Category",     f"Category value is not in the allowed list: {config.CATEGORIES}. Fix spelling in the Google Sheet and re-sync."),
        "DUPLICATE_CODE":       ("🔴 Duplicate Employee Code", "The same Employee Code appears more than once. Remove the duplicate row from the Google Sheet and re-sync."),
        "NO_DATA":              ("⚠️ No Data",              "This data source is empty. Upload or sync data first."),
        "BLANK_EMPLOYEE_CODE":  ("🔴 Blank Employee Code",  "The 'Created by' column is blank for this batch row. Fix in the Backend Data Excel."),
        "UNMAPPED_EMPLOYEE":    ("🟡 Unmapped Employee",    "Employee Code exists in Backend Data but not in Master Data. Add the employee to the Master Data Google Sheet and re-sync."),
        "INVALID_QUANTITY":     ("🔴 Invalid Quantity",     "Quantity is zero, negative, or not a number. Fix in the Backend Data Excel."),
        "INVALID_DATE":         ("🔴 Invalid Date",         "Date is not in YYYY-MM-DD format. Fix in the Backend Data Excel."),
        "BLANK_PLANT_CODE":     ("🔴 Blank Plant Code",     "Plant Code is missing in the Maintenance Cost file."),
        "UNMAPPED_PLANT":       ("🟡 Unmapped Plant",       "Plant Code in Maintenance Cost file is not in Master Data. Check spelling."),
        "INVALID_COST":         ("🔴 Invalid Cost",         "YTD Maintenance Cost is negative or not a number."),
    }
    _SRC_LABEL = {
        "master_data":       "Master Data",
        "backend_data":      "Backend Data",
        "maintenance_cost":  "Maintenance Cost",
    }

    if total_errors > 0:
        all_err_df = database.read_table("validation_errors")
        if not all_err_df.empty:
            master_n  = int((all_err_df["source"] == "master_data").sum())
            backend_n = int((all_err_df["source"] == "backend_data").sum())
            maint_n   = int((all_err_df["source"] == "maintenance_cost").sum())

            # Build grouped sections for the template
            # Section = one card per (source, error_type) combination
            from collections import OrderedDict
            sections = OrderedDict()
            for _, r in all_err_df.iterrows():
                src   = r["source"]
                etype = r["error_type"]
                key   = (src, etype)
                if key not in sections:
                    label, fix = _ERR_META.get(etype, (etype, "Review and correct this data."))
                    sections[key] = {
                        "source":       src,
                        "source_label": _SRC_LABEL.get(src, src),
                        "error_type":   etype,
                        "label":        label,
                        "fix":          fix,
                        "rows":         [],
                    }
                sections[key]["rows"].append({
                    "row_number":  r.get("row_number", ""),
                    "column_name": r.get("column_name", ""),
                    "message":     r.get("error_message", ""),
                })

            # Apply source filter
            if sel_src:
                sections = {k: v for k, v in sections.items() if v["source"] == sel_src}
            if sel_etype:
                sections = {k: v for k, v in sections.items() if v["error_type"] == sel_etype}

            err_sections = list(sections.values())
            sources = sorted(all_err_df["source"].unique().tolist())
            etypes  = sorted(all_err_df["error_type"].unique().tolist())
        else:
            err_sections = []
    else:
        err_sections = []

    return render_template("validation.html",
                           last=last, total_errors=total_errors,
                           master_n=master_n, backend_n=backend_n, maint_n=maint_n,
                           err_sections=err_sections,
                           sources=sources, etypes=etypes,
                           sel_src=sel_src, sel_etype=sel_etype,
                           src_label=_SRC_LABEL)


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
        return jsonify({"ok": False, "redirect": url_for("page_data_uploader")})
    try:
        _set_progress(15, "Connecting to Google Sheets…")
        clean_id = google_sheets.extract_sheet_id(sheet_id)
        _set_progress(35, "Fetching master data…")
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
        _set_progress(70, "Saving employee records…")
        inserted = database.replace_table_rows("master_data", rows)
        _set_progress(90, "Updating settings…")
        database.set_settings_bulk({"gsheet_id": clean_id, "gsheet_worksheet": worksheet,
                                    "gsheet_last_sync": now, "gsheet_last_count": str(inserted)})
        flash(f"Synced {inserted:,} employees from Google Sheets.", "success")
    except Exception as exc:
        flash(f"Sync failed: {exc}", "error")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("page_data_uploader")})


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
        month = int(request.form.get("maint_month", _date.today().month))
        year  = int(request.form.get("maint_year",  _date.today().year))
    except (ValueError, TypeError):
        flash("Invalid month or year selected.", "error")
        return redirect(url_for("page_data_uploader") + "#maintenance")
    try:
        import calendar as _cal
        df, warns = data_loader.load_maintenance_cost(f)
        for w in warns:
            flash(w, "warning")
        saved = data_loader.save_maintenance_cost(df, month, year)
        flash(f"Saved {saved:,} plant maintenance cost rows for {_cal.month_name[month]} {year}.", "success")
    except Exception as exc:
        flash(f"Upload failed: {exc}", "error")
    return redirect(url_for("page_data_uploader") + "#maintenance")


@app.route("/action/assign-maintenance-month", methods=["POST"])
def assign_maintenance_month():
    try:
        month = int(request.form.get("month"))
        year  = int(request.form.get("year"))
        import calendar as _cal
        updated = database.assign_maintenance_month(month, year)
        flash(f"✅ Assigned {updated} existing rows to {_cal.month_name[month]} {year}.", "success")
    except Exception as exc:
        flash(f"Assignment failed: {exc}", "error")
    return redirect(url_for("page_data_uploader") + "#maintenance")


@app.route("/action/delete-maintenance-month", methods=["POST"])
def delete_maintenance_month():
    try:
        month = int(request.form.get("month"))
        year  = int(request.form.get("year"))
        import calendar as _cal
        deleted = database.delete_maintenance_month(month, year)
        flash(f"Deleted {deleted} maintenance cost rows for {_cal.month_name[month]} {year}.", "success")
    except Exception as exc:
        flash(f"Delete failed: {exc}", "error")
    return redirect(url_for("page_data_uploader") + "#maintenance")


@app.route("/action/fetch-oracle", methods=["POST"])
def fetch_oracle():
    from_s = request.form.get("from_date", str(_date.today().replace(day=1)))
    to_s   = request.form.get("to_date",   str(_date.today()))
    replace = request.form.get("mode", "replace") == "replace"
    try:
        fd = _date.fromisoformat(from_s)
        td = _date.fromisoformat(to_s)
        _set_progress(15, "Testing Oracle connection…")
        result = oracle_connector.test_connection()
        if not result["success"]:
            flash(f"Oracle connection failed: {result['error']}", "error")
            _set_progress(100, "Complete")
            return jsonify({"ok": False, "redirect": url_for("page_data_uploader") + "#oracle"})
        _set_progress(35, "Fetching data from Oracle…")
        df_ora, warns = oracle_connector.fetch_backend_data(fd, td)
        for w in warns:
            flash(w, "warning")
        _set_progress(70, "Saving to database…")
        if df_ora.empty:
            flash("Oracle returned no rows for the selected date range.", "warning")
        else:
            saved = oracle_connector.save_oracle_backend_data(df_ora, fd, td, replace=replace)
            _set_progress(90, "Cleaning old records…")
            database.purge_old_oracle_data()
            flash(f"{'Replaced' if replace else 'Appended'} {saved:,} rows from Oracle.", "success")
    except Exception as exc:
        flash(f"Oracle fetch failed: {exc}", "error")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("page_data_uploader") + "#oracle"})


@app.route("/action/run-calculation", methods=["POST"])
def run_calculation():
    from_s = request.form.get("from_date")
    to_s   = request.form.get("to_date")
    try:
        fd = _date.fromisoformat(from_s)
        td = _date.fromisoformat(to_s)
        if fd > td:
            flash("'From Date' must be on or before 'To Date'.", "error")
            return jsonify({"ok": False, "redirect": url_for("page_calculate")})
        _set_progress(15, "Loading employee data…")
        result = calculator.run_calculation(
            month=fd.month, year=fd.year,
            start_date=str(fd), end_date=str(td),
        )
        _set_progress(90, "Saving results…")
        if result["error"]:
            flash(f"Calculation failed: {result['error']}", "error")
        else:
            flash(f"Calculation complete — {result['mapped']:,} employees, {result['unmapped']} unmapped.", "success")
            for w in result.get("calc_warnings", []):
                flash(w, "warning")
        _ss("calc_ran", True)
    except Exception as exc:
        flash(f"Calculation error: {exc}", "error")
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("page_calculate")})


@app.route("/action/run-validation", methods=["POST"])
def run_validation():
    try:
        errs: list = []
        _set_progress(15, "Validating master data…")
        validations._validate_master_data(errs)
        _set_progress(45, "Validating backend data…")
        validations._validate_backend_data(errs)
        _set_progress(70, "Validating maintenance costs…")
        validations._validate_maintenance_cost(errs)
        _set_progress(85, "Saving validation results…")
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
    _set_progress(100, "Complete")
    return jsonify({"ok": True, "redirect": url_for("page_validation")})


@app.route("/validation/download-excel")
def download_validation_excel():
    """Download I&D validation errors as an Excel file."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    import io

    all_err_df = database.read_table("validation_errors")
    errors = _records(all_err_df.drop(columns=["id"], errors="ignore")) \
             if not all_err_df.empty else []

    wb = Workbook()
    ws = wb.active
    ws.title = "Validation Errors"

    hdr_fill = PatternFill("solid", fgColor="1E3A5F")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    cols    = ["source", "row_number", "column_name", "error_type", "error_message", "created_at"]
    headers = ["Source", "Row #", "Column", "Issue Type", "What's Wrong — What to Fix", "Detected At"]

    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    red_fill    = PatternFill("solid", fgColor="FFE0E0")
    yellow_fill = PatternFill("solid", fgColor="FFF3CD")
    for ri, row in enumerate(errors, 2):
        for ci, col in enumerate(cols, 1):
            cell = ws.cell(row=ri, column=ci, value=str(row.get(col, "") or ""))
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        etype = str(row.get("error_type", "")).upper()
        fill  = yellow_fill if "UNMAPPED" in etype or "NO_DATA" in etype else red_fill
        for ci in range(1, len(cols) + 1):
            ws.cell(row=ri, column=ci).fill = fill

    col_widths = [18, 8, 18, 24, 65, 20]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    from flask import send_file
    return send_file(buf, as_attachment=True,
                     download_name="RDC_ID_Validation_Errors.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/tp/validation/download-excel")
def tp_download_validation_excel():
    """Download TP skip log + calc warnings as Excel."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    import io

    skip_log      = _ms("tp", "skip_log", [])
    calc_warnings = _ms("tp", "calc_warnings", [])

    wb = Workbook()

    # Sheet 1 — Skip Log
    ws1 = wb.active
    ws1.title = "Skip Log"
    hdr_fill = PatternFill("solid", fgColor="1E3A5F")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    skip_cols    = ["batch_ref", "plant_code", "production_date", "quantity", "time_taken_min", "reason"]
    skip_headers = ["Batch Ref", "Plant Code", "Date", "Quantity", "Time (min)", "Reason"]
    for ci, h in enumerate(skip_headers, 1):
        c = ws1.cell(row=1, column=ci, value=h)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal="center")
    row_fill = PatternFill("solid", fgColor="FFE0E0")
    for ri, row in enumerate(skip_log, 2):
        for ci, col in enumerate(skip_cols, 1):
            cell = ws1.cell(row=ri, column=ci, value=str(row.get(col, "") or ""))
            cell.fill = row_fill
    for ci, w in enumerate([18, 14, 14, 12, 12, 40], 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w

    # Sheet 2 — Calc Warnings
    ws2 = wb.create_sheet("Calc Warnings")
    ws2.cell(row=1, column=1, value="Warning").fill = hdr_fill
    ws2.cell(row=1, column=1).font = hdr_font
    ws2.column_dimensions["A"].width = 80
    warn_fill = PatternFill("solid", fgColor="FFF3CD")
    for ri, w in enumerate(calc_warnings, 2):
        cell = ws2.cell(row=ri, column=1, value=w)
        cell.fill = warn_fill
        cell.alignment = Alignment(wrap_text=True)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    from flask import send_file
    return send_file(buf, as_attachment=True,
                     download_name="RDC_TP_Validation_SkipLog.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/btrtp/validation/download-excel")
def btrtp_download_validation_excel():
    """Download BTRTP skip log + calc warnings as Excel."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    import io

    skip_log      = _ms("btrtp", "skip_log", [])
    calc_warnings = _ms("btrtp", "calc_warnings", [])

    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Skip Log"
    hdr_fill = PatternFill("solid", fgColor="1E3A5F")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    skip_cols    = ["batcher_id", "batch_ref", "plant_code", "production_date", "quantity", "time_taken_min", "reason"]
    skip_headers = ["Batcher ID", "Batch Ref", "Plant Code", "Date", "Quantity", "Time (min)", "Reason"]
    for ci, h in enumerate(skip_headers, 1):
        c = ws1.cell(row=1, column=ci, value=h)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = Alignment(horizontal="center")
    row_fill = PatternFill("solid", fgColor="FFE0E0")
    for ri, row in enumerate(skip_log, 2):
        for ci, col in enumerate(skip_cols, 1):
            cell = ws1.cell(row=ri, column=ci, value=str(row.get(col, "") or ""))
            cell.fill = row_fill
    for ci, w in enumerate([16, 18, 14, 14, 12, 12, 40], 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w

    ws2 = wb.create_sheet("Calc Warnings")
    ws2.cell(row=1, column=1, value="Warning").fill = hdr_fill
    ws2.cell(row=1, column=1).font = hdr_font
    ws2.column_dimensions["A"].width = 80
    warn_fill = PatternFill("solid", fgColor="FFF3CD")
    for ri, w in enumerate(calc_warnings, 2):
        cell = ws2.cell(row=ri, column=1, value=w)
        cell.fill = warn_fill
        cell.alignment = Alignment(wrap_text=True)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    from flask import send_file
    return send_file(buf, as_attachment=True,
                     download_name="RDC_BTRTP_Validation_SkipLog.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


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


# ── Real-time progress polling endpoint ───────────────────────────────────────
@app.route("/api/progress")
def api_progress():
    with _progress_lock:
        return jsonify(dict(_progress))


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


@app.route("/action/add-waiver", methods=["POST"])
def action_add_waiver():
    emp  = request.form.get("waiver_emp_code", "").strip()
    mon  = request.form.get("waiver_month", "").strip()
    yr   = request.form.get("waiver_year", "").strip()
    rsn  = request.form.get("waiver_reason", "").strip()
    cust = request.form.get("waiver_custom", "").strip()
    if not emp or not mon or not yr or not rsn:
        flash("All waiver fields are required.", "error")
        return redirect(url_for("page_calculate", anchor="waivers"))
    try:
        database.upsert_waiver(emp, int(mon), int(yr), rsn, cust)
        flash(f"✅ Waiver saved for {emp}.", "success")
    except Exception as e:
        flash(f"Could not save waiver: {e}", "error")
    return redirect(url_for("page_calculate", anchor="waivers"))


@app.route("/action/delete-waiver", methods=["POST"])
def action_delete_waiver():
    wid = request.form.get("waiver_id", "").strip()
    if not wid:
        flash("Invalid waiver ID.", "error")
        return redirect(url_for("page_calculate", anchor="waivers"))
    database.delete_waiver(int(wid))
    flash("Waiver removed.", "success")
    return redirect(url_for("page_calculate", anchor="waivers"))


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
        fname       = f"incentive_report_{from_s}_to_{to_s}.xlsx"
        xlsx_data   = report_generator.generate_excel_report(df_f, df_u, val_df, meta)
        tables_html = report_generator.build_email_tables_html(df_f)
        import html as _html_mod
        _safe_body  = _html_mod.escape(body or "").replace("\n", "<br>")
        _body_font  = "font-family:Arial,Calibri,sans-serif;font-size:12px;color:#000000;"
        html_body   = (
            f'<html><body style="margin:0;padding:8px 10px;{_body_font}">'
            f'<p style="margin:0 0 3px 0;">{_safe_body}</p>'
            f'{tables_html}'
            f'</body></html>'
        )

        res = email_helper.send_report_email(
            to_emails=to_addr, cc_emails=cc_addr,
            subject=subject, body=body,
            attachment_bytes=xlsx_data, attachment_name=fname,
            html_body=html_body,
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


# ═══════════════════════════════════════════════════════════════════════════
# RDC-ECMD MODULE — Energy Consumption & Mixer DG Ratio
# ═══════════════════════════════════════════════════════════════════════════

def _ecmd_ctx():
    """Base context dict for every ECMD page."""
    return dict(
        active_page="",
        bg_auto=bg_auto(), bg_animate=bg_animate(), bg_theme=bg_theme(),
    )


def _ecmd_mon_tag(month, year):
    import calendar as _cal
    return f"{_cal.month_abbr[month]}'{str(year)[2:]}"


@app.route("/ecmd")
def page_ecmd():
    return redirect(url_for("ecmd_dashboard"))


# ── Dashboard ─────────────────────────────────────────────────────────────────
@app.route("/ecmd/dashboard")
def ecmd_dashboard():
    all_counts  = database.get_table_counts()
    counts = {
        "ecmd_readings": all_counts.get("ecmd_readings", 0),
        "ecmd_results":  all_counts.get("ecmd_results",  0),
        "tp_plant_data": all_counts.get("tp_plant_data", 0),
    }
    ora_ready   = oracle_connector.is_configured()
    ecmd_months = database.get_ecmd_months()
    ctx = _ecmd_ctx()
    ctx["active_page"] = "dashboard"
    return render_template("ecmd_dashboard.html", counts=counts,
                           ora_ready=ora_ready, ecmd_months=ecmd_months, **ctx)


# ── Data Entry ────────────────────────────────────────────────────────────────
@app.route("/ecmd/data-entry", methods=["GET", "POST"])
def ecmd_data_entry():
    import calendar as _cal

    today = _date.today()
    sel_month = int(request.args.get("month", today.month))
    sel_year  = int(request.args.get("year",  today.year))

    all_plant_rows = database.get_tp_plants()
    # Filter plants by user access (PLANT_USER / REGIONAL_USER see only their plants)
    plant_rows = auth.apply_plant_filter_rows(all_plant_rows, g.current_user)
    readings_map = {r["plant_code"]: r
                    for r in database.get_ecmd_readings_for_month(sel_month, sel_year)}

    months_list = [(m, _cal.month_name[m]) for m in range(1, 13)]
    years_list  = list(range(today.year - 2, today.year + 2))

    ctx = _ecmd_ctx()
    ctx["active_page"] = "data_entry"
    return render_template("ecmd_data_entry.html",
                           plant_rows=plant_rows,
                           readings_map=readings_map,
                           sel_month=sel_month, sel_year=sel_year,
                           months_list=months_list, years_list=years_list,
                           **ctx)


@app.route("/ecmd/action/save-reading", methods=["POST"])
def ecmd_save_reading():
    plant_code = request.form.get("plant_code", "").strip()
    month      = int(request.form.get("month", _date.today().month))
    year       = int(request.form.get("year",  _date.today().year))
    # Plant-level access check for ECMD entry
    user = g.current_user
    plant_row = next((p for p in database.get_tp_plants() if p["plant_code"] == plant_code), None)
    if plant_row and not auth.user_can_access_plant(user, plant_row["plant_name"]):
        flash("You are not authorised to save readings for this plant.", "error")
        return redirect(url_for("ecmd_data_entry", month=month, year=year))

    def _fv(key):
        v = request.form.get(key, "").strip()
        return float(v) if v != "" else None

    data = {
        "eb_kwh_open":        _fv("eb_kwh_open"),
        "eb_kwh_close":       _fv("eb_kwh_close"),
        "eb_kvah_open":       _fv("eb_kvah_open"),
        "eb_kvah_close":      _fv("eb_kvah_close"),
        "mf":                 _fv("mf") or 1.0,
        "dg_hr_open":         _fv("dg_hr_open"),
        "dg_hr_close":        _fv("dg_hr_close"),
        "dg_kwh_open":        _fv("dg_kwh_open"),
        "dg_kwh_close":       _fv("dg_kwh_close"),
        "mixer_dg_hr_open":   _fv("mixer_dg_hr_open"),
        "mixer_dg_hr_close":  _fv("mixer_dg_hr_close"),
        "diesel_issued_ltrs": _fv("diesel_issued_ltrs"),
        "volume_on_dg":       _fv("volume_on_dg"),
    }
    if not plant_code:
        flash("Plant Code is required.", "error")
        return redirect(url_for("ecmd_data_entry", month=month, year=year))
    try:
        database.upsert_ecmd_reading(plant_code, month, year, data)
        flash(f"✅ Readings saved for plant {plant_code} — {month}/{year}.", "success")
    except Exception as exc:
        flash(f"Save failed: {exc}", "error")
    return redirect(url_for("ecmd_data_entry", month=month, year=year))


@app.route("/ecmd/action/delete-reading", methods=["POST"])
def ecmd_delete_reading():
    plant_code = request.form.get("plant_code", "").strip()
    month      = int(request.form.get("month", _date.today().month))
    year       = int(request.form.get("year",  _date.today().year))
    n = database.delete_ecmd_reading(plant_code, month, year)
    if n:
        flash(f"✅ Reading for {plant_code} ({month}/{year}) deleted.", "success")
    else:
        flash("Reading not found.", "warning")
    return redirect(url_for("ecmd_data_entry", month=month, year=year))


# ── Calculate ─────────────────────────────────────────────────────────────────
@app.route("/ecmd/calculate", methods=["GET", "POST"])
def ecmd_calculate():
    import calendar as _cal

    today = _date.today()
    plant_rows = _ms("ecmd", "plant_rows", [])
    loc_rows   = _ms("ecmd", "loc_rows",   [])
    warnings   = _ms("ecmd", "calc_warnings", [])
    ran        = _ms("ecmd", "calc_ran", False)

    months_list = [(m, _cal.month_name[m]) for m in range(1, 13)]
    years_list  = list(range(today.year - 2, today.year + 2))

    if request.method == "POST":
        from_s = request.form.get("from_date", "")
        to_s   = request.form.get("to_date",   "")
        try:
            fd = _date.fromisoformat(from_s)
            td = _date.fromisoformat(to_s)
        except (ValueError, TypeError):
            flash("Please select valid From and To dates.", "error")
            return jsonify({"ok": False, "redirect": url_for("ecmd_calculate")})
        if fd > td:
            flash("'From Date' must be on or before 'To Date'.", "error")
            return jsonify({"ok": False, "redirect": url_for("ecmd_calculate")})

        month, year = fd.month, fd.year
        _set_progress(20, "Running ECMD calculations…")
        plant_rows, warnings = ecmd_calculator.run_ecmd_calculation(
            month, year, from_date=str(fd), to_date=str(td))
        _set_progress(70, "Building location summary…")
        loc_rows = ecmd_calculator.build_location_summary(plant_rows)
        _set_progress(85, "Saving results…")
        _mss("ecmd", "plant_rows",     plant_rows)
        _mss("ecmd", "loc_rows",       loc_rows)
        _mss("ecmd", "calc_warnings",  warnings)
        _mss("ecmd", "calc_month",     month)
        _mss("ecmd", "calc_year",      year)
        _mss("ecmd", "calc_from",      str(fd))
        _mss("ecmd", "calc_to",        str(td))
        _mss("ecmd", "calc_ran",       True)
        ran = True

        if plant_rows:
            database.save_ecmd_results(plant_rows, month, year)
            flash(f"✅ Calculated {len(plant_rows)} plant(s) for {fd} → {td}.", "success")
        else:
            flash("No results — enter readings first or check warnings.", "warning")
        _set_progress(100, "Complete")
        return jsonify({"ok": True, "redirect": url_for("ecmd_calculate")})

    ctx = _ecmd_ctx()
    ctx["active_page"] = "calculate"
    return render_template("ecmd_calculate.html",
                           plant_rows=plant_rows, loc_rows=loc_rows,
                           warnings=warnings, ran=ran,
                           months_list=months_list, years_list=years_list,
                           default_from=_ms("ecmd", "calc_from", str(today.replace(day=1))),
                           default_to=_ms("ecmd", "calc_to", str(today)),
                           **ctx)


# ── Reports ───────────────────────────────────────────────────────────────────
@app.route("/ecmd/reports")
def ecmd_reports():
    import calendar as _cal

    today = _date.today()
    from_s = request.args.get("from_date",
                              _ms("ecmd", "calc_from", str(today.replace(day=1))))
    to_s   = request.args.get("to_date",
                              _ms("ecmd", "calc_to", str(today)))
    try:
        fd = _date.fromisoformat(from_s)
        td = _date.fromisoformat(to_s)
    except ValueError:
        fd, td = today.replace(day=1), today

    if "from_date" in request.args:
        month, year = fd.month, fd.year
        all_plants, warnings = ecmd_calculator.run_ecmd_calculation(
            month, year, from_date=str(fd), to_date=str(td))
        loc_rows = ecmd_calculator.build_location_summary(all_plants)
        _mss("ecmd", "plant_rows",    all_plants)
        _mss("ecmd", "loc_rows",      loc_rows)
        _mss("ecmd", "calc_warnings", warnings)
        _mss("ecmd", "calc_from",     str(fd))
        _mss("ecmd", "calc_to",       str(td))
        _mss("ecmd", "calc_month",    month)
        _mss("ecmd", "calc_year",     year)
    else:
        all_plants = _ms("ecmd", "plant_rows", [])
        loc_rows   = _ms("ecmd", "loc_rows",   [])

    # Filters
    excos  = request.args.getlist("exco")
    bheads = request.args.getlist("bhead")
    search = request.args.get("search", "")

    unique_excos  = sorted({r.get("exco_location", "") for r in all_plants if r.get("exco_location")})
    unique_bheads = sorted({r.get("business_head", "") for r in all_plants if r.get("business_head")})

    plant_rows = all_plants
    if excos:
        plant_rows = [r for r in plant_rows if r.get("exco_location") in excos]
    if bheads:
        plant_rows = [r for r in plant_rows if r.get("business_head") in bheads]
    if search.strip():
        s = search.strip().lower()
        plant_rows = [r for r in plant_rows
                      if s in str(r.get("plant_code", "")).lower()
                      or s in str(r.get("plant_name", "")).lower()]

    if plant_rows is not all_plants:
        loc_rows = ecmd_calculator.build_location_summary(plant_rows)

    _mss("ecmd", "report_plant_rows", plant_rows)
    _mss("ecmd", "report_loc_rows",   loc_rows)

    smtp        = email_helper.get_smtp_config()
    email_ready = email_helper.is_configured()
    month = _ms("ecmd", "calc_month", fd.month)
    year  = _ms("ecmd", "calc_year",  fd.year)

    _elog = database.read_table("email_log", order_by="id DESC")
    if not _elog.empty and "report_file_name" in _elog.columns:
        _elog = _elog[_elog["report_file_name"].str.startswith("RDC_ECMD_", na=False)]
    ecmd_email_log = _records(_elog.drop(columns=["id"], errors="ignore").head(20)) \
        if not _elog.empty else []

    if (fd.year, fd.month) == (td.year, td.month):
        month_label = f"{_cal.month_name[fd.month]} {fd.year}"
    else:
        month_label = f"{fd:%d %b %Y} → {td:%d %b %Y}"

    default_to      = database.get_module_setting("ecmd", "email_default_to", "") or smtp.get("default_to", "")
    default_cc      = database.get_module_setting("ecmd", "email_default_cc", "") or smtp.get("default_cc", "")
    default_subject = (database.get_module_setting("ecmd", "email_default_subject", "")
                       or f"RDC-ECMD Energy & DG Report — {month_label}")
    default_body    = (database.get_module_setting("ecmd", "email_default_body", "")
                       or f"Dear Team,\n\nPlease find attached the Energy Consumption & "
                          f"Mixer DG Ratio Report for {month_label}.\n\nRegards,\nRDC Operations")

    ctx = _ecmd_ctx()
    ctx["active_page"] = "reports"
    return render_template("ecmd_reports.html",
                           plant_rows=plant_rows, loc_rows=loc_rows,
                           total_plants=len(all_plants),
                           from_date=str(fd), to_date=str(td),
                           unique_excos=unique_excos, unique_bheads=unique_bheads,
                           excos=excos, bheads=bheads, search=search,
                           month_label=month_label,
                           email_cfg=smtp, email_ready=email_ready,
                           default_to=default_to, default_cc=default_cc,
                           default_subject=default_subject, default_body=default_body,
                           ecmd_email_log=ecmd_email_log, **ctx)


def _ecmd_build_excel(plant_rows, loc_rows, month, year):
    """Build colour-coded Excel report for ECMD (Energy + DG tabs)."""
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    mon_tag    = _ecmd_mon_tag(month, year)
    HDR_FILL   = PatternFill("solid", fgColor="0A2540")
    PAN_FILL   = PatternFill("solid", fgColor="D9D9D9")
    PLAIN_FILL = PatternFill("solid", fgColor="FFFFFF")
    ALT_FILL   = PatternFill("solid", fgColor="F0F4FA")
    HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
    TITLE_FONT = Font(bold=True, color="FFFFFF", size=11)
    PAN_FONT   = Font(bold=True, color="222222", size=10)
    PLAIN_FONT = Font(color="19263A", size=10)
    ALT_FONT   = Font(color="19263A", size=10)
    CTR  = Alignment(vertical="center", horizontal="center")
    LEFT = Alignment(vertical="center", horizontal="left")
    WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
    thin = Side(style="thin", color="9A9A9A")
    BDR  = Border(left=thin, right=thin, top=thin, bottom=thin)

    wb = Workbook()

    # ── Sheet 1: Energy Consumption ───────────────────────────────────────────
    ws_e = wb.active
    ws_e.title = "Energy Consumption"
    E_HEADS = ["Sr.", "Plant Code", "Plant Name", "Exco Location", "Business Head",
               "EB KWh", "DG KWh", "Total KWh", "Total Vol (MT)", "Energy/MT (KWh)"]
    ws_e.merge_cells(f"A1:{get_column_letter(len(E_HEADS))}1")
    t = ws_e.cell(1, 1, f"Plant Energy Consumption - {mon_tag}")
    t.fill = HDR_FILL; t.font = TITLE_FONT; t.alignment = CTR
    ws_e.row_dimensions[1].height = 22
    for ci, h in enumerate(E_HEADS, 1):
        c = ws_e.cell(2, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    ws_e.row_dimensions[2].height = 28
    for ri, row in enumerate(plant_rows, 3):
        fill = ALT_FILL if ri % 2 == 0 else PLAIN_FILL
        font = ALT_FONT if ri % 2 == 0 else PLAIN_FONT
        vals = [ri - 2, row.get("plant_code",""), row.get("plant_name",""),
                row.get("exco_location",""), row.get("business_head",""),
                round(float(row.get("eb_kwh",0)),2),
                round(float(row.get("dg_kwh",0)),2),
                round(float(row.get("total_kwh",0)),2),
                round(float(row.get("total_volume",0)),2),
                round(float(row.get("energy_per_mt",0)),4)]
        for ci, v in enumerate(vals, 1):
            c = ws_e.cell(ri, ci, v)
            c.fill = fill; c.font = font; c.border = BDR
            c.alignment = LEFT if ci in (2, 3, 4, 5) else CTR
    # Location summary rows at bottom
    ws_e.append([])
    loc_hdr_row = len(plant_rows) + 4
    ws_e.merge_cells(f"A{loc_hdr_row}:{get_column_letter(len(E_HEADS))}{loc_hdr_row}")
    th = ws_e.cell(loc_hdr_row, 1, f"Location Summary - {mon_tag}")
    th.fill = HDR_FILL; th.font = TITLE_FONT; th.alignment = CTR
    for ci, h in enumerate(["Sr.", "Exco Location", "Plants", "EB KWh", "DG KWh",
                              "Total KWh", "Total Vol (MT)", "Energy/MT"], 1):
        c = ws_e.cell(loc_hdr_row + 1, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    srno = 1
    for ri, row in enumerate(loc_rows, loc_hdr_row + 2):
        pan  = bool(row.get("is_pan_india"))
        fill = PAN_FILL if pan else PLAIN_FILL
        font = PAN_FONT if pan else PLAIN_FONT
        srno_val = "—" if pan else srno
        vals = [srno_val, row.get("exco_location",""), row.get("plant_count",0),
                round(float(row.get("eb_kwh", row.get("total_kwh",0)) - float(row.get("dg_kwh",0))),2)
                    if "eb_kwh" not in row else round(float(row.get("eb_kwh",0)),2),
                round(float(row.get("dg_kwh",0) if "dg_kwh" in row else 0),2),
                round(float(row.get("total_kwh",0)),2),
                round(float(row.get("total_volume",0)),2),
                round(float(row.get("energy_per_mt",0)),4)]
        for ci, v in enumerate(vals, 1):
            c = ws_e.cell(ri, ci, v)
            c.fill = fill; c.font = font; c.border = BDR
            c.alignment = LEFT if ci == 2 else CTR
        if not pan:
            srno += 1
    for ci, w in enumerate([5, 12, 22, 20, 18, 12, 12, 14, 14, 14], 1):
        ws_e.column_dimensions[get_column_letter(ci)].width = w

    # ── Sheet 2: Mixer DG Ratio ───────────────────────────────────────────────
    ws_d = wb.create_sheet("Mixer DG Ratio")
    D_HEADS = ["Sr.", "Plant Code", "Plant Name", "Exco Location", "Business Head",
               "DG Run Hrs", "Mixer DG Hrs", "Mixer DG %", "Diesel (L)", "L/Hr",
               "Vol on DG (MT)"]
    ws_d.merge_cells(f"A1:{get_column_letter(len(D_HEADS))}1")
    t2 = ws_d.cell(1, 1, f"Mixer DG Ratio Report - {mon_tag}")
    t2.fill = HDR_FILL; t2.font = TITLE_FONT; t2.alignment = CTR
    ws_d.row_dimensions[1].height = 22
    for ci, h in enumerate(D_HEADS, 1):
        c = ws_d.cell(2, ci, h)
        c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = WRAP; c.border = BDR
    ws_d.row_dimensions[2].height = 28
    for ri, row in enumerate(plant_rows, 3):
        fill = ALT_FILL if ri % 2 == 0 else PLAIN_FILL
        font = ALT_FONT if ri % 2 == 0 else PLAIN_FONT
        pct  = float(row.get("mixer_dg_ratio", 0))
        vals = [ri - 2, row.get("plant_code",""), row.get("plant_name",""),
                row.get("exco_location",""), row.get("business_head",""),
                round(float(row.get("dg_run_hrs",0)),2),
                round(float(row.get("mixer_dg_hrs",0)),2),
                f"{round(pct,1)}%",
                round(float(row.get("diesel_issued_ltrs",0)),2),
                round(float(row.get("ltr_per_hr",0)),3),
                round(float(row.get("volume_on_dg",0)),2)]
        for ci, v in enumerate(vals, 1):
            c = ws_d.cell(ri, ci, v)
            c.fill = fill; c.font = font; c.border = BDR
            c.alignment = LEFT if ci in (2, 3, 4, 5) else CTR
    for ci, w in enumerate([5, 12, 22, 20, 18, 10, 11, 10, 10, 8, 12], 1):
        ws_d.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


@app.route("/ecmd/download-excel")
def ecmd_download_excel():
    plant_rows = _ms("ecmd", "report_plant_rows", _ms("ecmd", "plant_rows", []))
    loc_rows   = _ms("ecmd", "report_loc_rows",   _ms("ecmd", "loc_rows",   []))
    month = _ms("ecmd", "calc_month", _date.today().month)
    year  = _ms("ecmd", "calc_year",  _date.today().year)
    if not plant_rows:
        flash("No data to download — run Calculate first.", "warning")
        return redirect(url_for("ecmd_reports"))
    excel_bytes = _ecmd_build_excel(plant_rows, loc_rows, month, year)
    fname = f"RDC_ECMD_{year}_{month:02d}.xlsx"
    return send_file(io.BytesIO(excel_bytes), as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/ecmd/download-csv")
def ecmd_download_csv():
    plant_rows = _ms("ecmd", "report_plant_rows", _ms("ecmd", "plant_rows", []))
    month = _ms("ecmd", "calc_month", _date.today().month)
    year  = _ms("ecmd", "calc_year",  _date.today().year)
    if not plant_rows:
        flash("No data to download — run Calculate first.", "warning")
        return redirect(url_for("ecmd_reports"))
    buf = io.BytesIO(pd.DataFrame(plant_rows).to_csv(index=False).encode("utf-8"))
    return send_file(buf, as_attachment=True,
                     download_name=f"RDC_ECMD_{year}_{month:02d}.csv",
                     mimetype="text/csv")


@app.route("/ecmd/action/send-email", methods=["POST"])
def ecmd_send_email():
    plant_rows = _ms("ecmd", "report_plant_rows", _ms("ecmd", "plant_rows", []))
    loc_rows   = _ms("ecmd", "report_loc_rows",   _ms("ecmd", "loc_rows",   []))
    month      = _ms("ecmd", "calc_month", _date.today().month)
    year       = _ms("ecmd", "calc_year",  _date.today().year)
    to_addr    = request.form.get("to", "").strip()
    cc_addr    = request.form.get("cc", "").strip()
    subject    = request.form.get("subject", "").strip()
    body       = request.form.get("body", "").strip()

    if not plant_rows:
        flash("No results to email — run Calculate first.", "warning")
        return redirect(url_for("ecmd_reports"))
    if not to_addr:
        flash("Please enter at least one To email address.", "error")
        return redirect(url_for("ecmd_reports"))

    import calendar as _cal
    month_name = _cal.month_name[month]
    mon_tag    = _ecmd_mon_tag(month, year)
    if not subject:
        subject = f"RDC-ECMD Energy & DG Report — {month_name} {year}"

    excel_bytes = _ecmd_build_excel(plant_rows, loc_rows, month, year)
    fname = f"RDC_ECMD_{year}_{month:02d}.xlsx"

    _TH = ('style="background-color:#082B49;color:#fff;font-weight:bold;'
           'text-align:center;padding:3px 6px;border:1px solid #808080;'
           'font-size:11px;font-family:Arial,sans-serif;white-space:nowrap;"')
    def _td_e(align="center"):
        return (f'style="background-color:#fff;color:#19263A;padding:2px 5px;'
                f'border:1px solid #999;text-align:{align};'
                f'font-size:11px;font-family:Arial,sans-serif;white-space:nowrap;"')
    def _td_pan(align="center"):
        return (f'style="background-color:#D9D9D9;color:#222;font-weight:bold;'
                f'padding:2px 5px;border:1px solid #999;text-align:{align};'
                f'font-size:11px;font-family:Arial,sans-serif;white-space:nowrap;"')

    _tbl = 'border-collapse:collapse;width:auto;margin:4px 0 10px'
    _bf  = 'font-family:Arial,Calibri,sans-serif;font-size:12px;color:#000;'

    # Energy table
    e_head = (f'<tr><th {_TH}>Sr.</th><th {_TH}>Plant</th><th {_TH}>Location</th>'
              f'<th {_TH}>EB KWh</th><th {_TH}>DG KWh</th><th {_TH}>Total KWh</th>'
              f'<th {_TH}>Total Vol (MT)</th><th {_TH}>Energy/MT</th></tr>')
    e_rows = ""
    for i, r in enumerate(plant_rows, 1):
        e_rows += (f'<tr><td {_td_e("center")}>{i}</td>'
                   f'<td {_td_e("left")}>{r.get("plant_name","")}</td>'
                   f'<td {_td_e("left")}>{r.get("exco_location","")}</td>'
                   f'<td {_td_e("right")}>{round(float(r.get("eb_kwh",0)),2)}</td>'
                   f'<td {_td_e("right")}>{round(float(r.get("dg_kwh",0)),2)}</td>'
                   f'<td {_td_e("right")}><b>{round(float(r.get("total_kwh",0)),2)}</b></td>'
                   f'<td {_td_e("right")}>{round(float(r.get("total_volume",0)),2)}</td>'
                   f'<td {_td_e("right")}>{round(float(r.get("energy_per_mt",0)),4)}</td></tr>')

    # DG table
    d_head = (f'<tr><th {_TH}>Sr.</th><th {_TH}>Plant</th><th {_TH}>Location</th>'
              f'<th {_TH}>DG Run Hrs</th><th {_TH}>Mixer DG Hrs</th>'
              f'<th {_TH}>Mixer DG %</th><th {_TH}>Diesel (L)</th>'
              f'<th {_TH}>L/Hr</th><th {_TH}>Vol on DG</th></tr>')
    d_rows = ""
    for i, r in enumerate(plant_rows, 1):
        d_rows += (f'<tr><td {_td_e("center")}>{i}</td>'
                   f'<td {_td_e("left")}>{r.get("plant_name","")}</td>'
                   f'<td {_td_e("left")}>{r.get("exco_location","")}</td>'
                   f'<td {_td_e("right")}>{round(float(r.get("dg_run_hrs",0)),2)}</td>'
                   f'<td {_td_e("right")}>{round(float(r.get("mixer_dg_hrs",0)),2)}</td>'
                   f'<td {_td_e("center")}><b>{round(float(r.get("mixer_dg_ratio",0)),1)}%</b></td>'
                   f'<td {_td_e("right")}>{round(float(r.get("diesel_issued_ltrs",0)),2)}</td>'
                   f'<td {_td_e("right")}>{round(float(r.get("ltr_per_hr",0)),3)}</td>'
                   f'<td {_td_e("right")}>{round(float(r.get("volume_on_dg",0)),2)}</td></tr>')

    html_body = (
        f'<html><body style="margin:0;padding:8px 10px;{_bf}">'
        f'<p style="margin:0 0 3px 0;">Dear Team,</p>'
        f'<p style="margin:0 0 3px 0;">Please find attached the Energy Consumption &amp; '
        f'Mixer DG Ratio Report for {month_name} {year}.</p>'
        f'<p style="margin:0 0 6px 0;">Regards,<br>RDC Operations</p>'
        f'<p style="margin:0 0 4px 0;font-size:13px;font-weight:bold;color:#082B49;">'
        f'Plant Energy Consumption - {mon_tag}</p>'
        f'<table style="{_tbl}">{e_head}{e_rows}</table>'
        f'<p style="margin:10px 0 4px 0;font-size:13px;font-weight:bold;color:#082B49;">'
        f'Mixer DG Ratio - {mon_tag}</p>'
        f'<table style="{_tbl}">{d_head}{d_rows}</table>'
        f'</body></html>'
    )

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
    return redirect(url_for("ecmd_reports"))


# ── Settings ──────────────────────────────────────────────────────────────────
@app.route("/ecmd/settings", methods=["GET"])
def ecmd_settings():
    smtp             = email_helper.get_smtp_config()
    email_configured = bool(smtp.get("host") and smtp.get("sender"))
    ora_configured   = oracle_connector.is_configured()
    sched_enabled    = database.get_module_setting("ecmd", "email_schedule_enabled", "false") == "true"
    sched_time       = database.get_module_setting("ecmd", "email_schedule_time", "08:00")
    sched_to         = database.get_module_setting("ecmd", "email_schedule_to", "")
    sched_cc         = database.get_module_setting("ecmd", "email_schedule_cc", "")
    sched_last_status = database.get_module_setting("ecmd", "email_schedule_last_status", "")
    ecmd_email_to      = database.get_module_setting("ecmd", "email_default_to", "")
    ecmd_email_cc      = database.get_module_setting("ecmd", "email_default_cc", "")
    ecmd_email_subject = database.get_module_setting("ecmd", "email_default_subject", "")
    ecmd_email_body    = database.get_module_setting("ecmd", "email_default_body", "")
    entry_open_month = int(database.get_setting("ecmd_entry_open_month", 0) or 0)
    entry_open_year  = int(database.get_setting("ecmd_entry_open_year",  0) or 0)
    _MN = ["","January","February","March","April","May","June",
           "July","August","September","October","November","December"]
    entry_open_month_name = _MN[entry_open_month] if entry_open_month else ""
    ctx = _ecmd_ctx()
    ctx["active_page"] = "settings"
    return render_template("ecmd_settings.html",
                           smtp=smtp, email_configured=email_configured,
                           ora_configured=ora_configured,
                           sched_enabled=sched_enabled, sched_time=sched_time,
                           sched_to=sched_to, sched_cc=sched_cc,
                           sched_last_status=sched_last_status,
                           ecmd_email_to=ecmd_email_to, ecmd_email_cc=ecmd_email_cc,
                           ecmd_email_subject=ecmd_email_subject,
                           ecmd_email_body=ecmd_email_body,
                           entry_open_month=entry_open_month,
                           entry_open_year=entry_open_year,
                           entry_open_month_name=entry_open_month_name,
                           **ctx)


@app.route("/ecmd/settings/set-entry-period", methods=["POST"])
@auth.admin_required
def ecmd_set_entry_period():
    month = request.form.get("entry_month", "").strip()
    year  = request.form.get("entry_year",  "").strip()
    database.set_setting("ecmd_entry_open_month", month if month else "0")
    database.set_setting("ecmd_entry_open_year",  year  if month else "0")
    if month and year:
        flash(f"Entry window opened for month {month}/{year}.", "success")
    else:
        flash("Entry window closed.", "success")
    return redirect(url_for("ecmd_settings"))


@app.route("/ecmd/settings/save-schedule", methods=["POST"])
def ecmd_save_schedule():
    sched_time = request.form.get("sched_time", "08:00").strip()
    database.set_module_settings_bulk("ecmd", {
        "email_schedule_time": sched_time,
        "email_schedule_to":   request.form.get("sched_to", "").strip(),
        "email_schedule_cc":   request.form.get("sched_cc", "").strip(),
    })
    if _scheduler:
        try:
            h, m = map(int, sched_time.split(":"))
        except Exception:
            h, m = 8, 0
        try:
            _scheduler.reschedule_job("ecmd_monthly_email",
                                      trigger=CronTrigger(day=1, hour=h, minute=m))
        except Exception:
            pass
    flash("ECMD schedule settings saved.", "success")
    return redirect(url_for("ecmd_settings", m="schedule"))


@app.route("/ecmd/settings/toggle-schedule", methods=["POST"])
def ecmd_toggle_schedule():
    current = database.get_module_setting("ecmd", "email_schedule_enabled", "false")
    new_val = "false" if current == "true" else "true"
    database.set_module_setting("ecmd", "email_schedule_enabled", new_val)
    flash(f"RDC-ECMD scheduled report {'enabled' if new_val == 'true' else 'disabled'}.", "success")
    return redirect(url_for("ecmd_settings", m="schedule"))


@app.route("/ecmd/settings/save-email-defaults", methods=["POST"])
def ecmd_save_email_defaults():
    database.set_module_settings_bulk("ecmd", {
        "email_default_to":      request.form.get("default_to", "").strip(),
        "email_default_cc":      request.form.get("default_cc", "").strip(),
        "email_default_subject": request.form.get("default_subject", "").strip(),
        "email_default_body":    request.form.get("default_body", "").strip(),
    })
    flash("ECMD email defaults saved.", "success")
    return redirect(url_for("ecmd_settings", m="email"))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("  RDC Batching Incentive Calculator")
    print("  Flask server starting on http://localhost:2001")
    print("=" * 60)
    app.run(host="0.0.0.0", port=2001, debug=False, use_reloader=False)
