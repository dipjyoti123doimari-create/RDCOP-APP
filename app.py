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
from datetime import date as _date, datetime as _dt

import pandas as pd
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


# ── Boot ─────────────────────────────────────────────────────────────────────
database.init_db()

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
    "ytd_maintenance_cost": "YTD Cost",
    "incentive_eligible":   "Eligible",
    "incentive_rate":       "Inc Rate",
    "incentive_amount":     "Incentive Amount (Rs)",
    "deduction_target":     "Ded Target",
    "shortfall_quantity":   "Shortfall",
    "deduction_amount":     "Deduction Amount (Rs)",
    "remarks":              "Remarks",
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
def page_dashboard():
    _ss("active_page", "dashboard")
    counts = database.get_table_counts()
    return render_template("dashboard.html",
                           counts=counts,
                           db_path=database.DB_PATH)


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
        for r in filtered:
            r["_cls"] = _row_cls(r)
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
                           result_cols=RESULT_COLS, result_labels=RESULT_LABELS,
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

    return render_template("settings.html",
                           smtp=smtp, email_configured=email_configured,
                           ora=ora, ora_configured=ora_configured,
                           db_size_mb=cache_helpers.get_db_size_mb(),
                           bg_auto=bg_auto(), bg_animate=bg_animate(), bg_theme=bg_theme(),
                           last_cache_clear=last_cache_clear,
                           cache_cleared_today=cache_cleared_today)


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
    try:
        database.add_employee(
            request.form["code"].strip(),
            request.form.get("name", "").strip(),
            request.form.get("designation", "").strip(),
            request.form.get("category", ""),
            request.form.get("plant", "").strip(),
            request.form.get("plant_code", "").strip(),
        )
        flash(f"Added employee {request.form['code'].strip()}.", "success")
    except ValueError as e:
        flash(str(e), "error")
    return redirect(url_for("page_data_uploader"))


@app.route("/action/update-employee/<code>", methods=["POST"])
def update_employee(code):
    try:
        database.update_employee(
            code,
            request.form.get("name", "").strip(),
            request.form.get("designation", "").strip(),
            request.form.get("category", ""),
            request.form.get("plant", "").strip(),
            request.form.get("plant_code", "").strip(),
        )
        flash(f"Updated employee {code}.", "success")
    except ValueError as e:
        flash(str(e), "error")
    return redirect(url_for("page_data_uploader"))


@app.route("/action/delete-employee/<code>", methods=["POST"])
def delete_employee(code):
    try:
        database.delete_employee(code)
        flash(f"Deleted employee {code}.", "success")
    except ValueError as e:
        flash(str(e), "error")
    return redirect(url_for("page_data_uploader"))


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
    database.set_settings_bulk({
        "bg_auto_theme":   "true" if data.get("auto_theme") else "false",
        "bg_manual_theme": data.get("manual_theme", "Daytime"),
    })
    return jsonify({"ok": True})


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

        html_body = None
        if incl_tables:
            tables_html = report_generator.build_email_tables_html(df_f)
            html_body   = email_helper.wrap_html_body(body, tables_html)

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


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("  RDC Batching Incentive Calculator")
    print("  Flask server starting on http://localhost:2001")
    print("=" * 60)
    app.run(host="0.0.0.0", port=2001, debug=False)
