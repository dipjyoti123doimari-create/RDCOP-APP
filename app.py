"""
app.py
======
This is the MAIN file of the app. You run the whole app with this command
(from inside the project folder, in the VS Code terminal):

    streamlit run app.py

What this file does (and ONLY this):
- Sets up the page (title, icon, layout).
- Applies the theme (from ui_helpers.py).
- Shows the logo (if one exists) and the sidebar navigation menu.
- Looks at which page the user picked and calls the matching page function.

Important design rule for this project:
app.py should stay SMALL. It should NOT contain heavy business logic
(no calculations, no database code). Those live in their own files
(calculator.py, database.py, etc.) which we build in later phases.

Right now (Phase 1) every page is just a friendly placeholder so you can
see the navigation working.
"""

import os
import time
from datetime import datetime as _dt

import pandas as pd
import streamlit as st

import cache_helpers
import calculator
import config
import data_loader
import database
import email_helper
import google_sheets
import report_generator
import ui_helpers
import validations
from utils.animated_background import render_interactive_background, THEME_NAMES


# ---------------------------------------------------------------------------
# STEP 1: Configure the browser tab (title + icon) and page layout.
# ---------------------------------------------------------------------------
# We try to use the logo as the browser tab icon. If the logo file is missing,
# we fall back to a simple emoji so the app never crashes.
def _get_page_icon():
    """Return the logo path if it exists, otherwise a fallback emoji."""
    if os.path.exists(config.LOGO_PATH):
        return config.LOGO_PATH
    return "📊"  # fallback icon shown in the browser tab


st.set_page_config(
    page_title=config.APP_NAME,
    page_icon=_get_page_icon(),
    layout="wide",                 # use the full width of the browser
    initial_sidebar_state="expanded",
)

# Apply our custom theme. This must come right after set_page_config().
ui_helpers.inject_custom_css()


# ---------------------------------------------------------------------------
# STEP 2: Build the sidebar (logo + navigation menu).
# ---------------------------------------------------------------------------
def render_sidebar():
    """
    Draw the sidebar and return the name of the page the user selected.
    """
    with st.sidebar:
        # Show the logo at the top IF the file exists. If not, show the app name.
        if os.path.exists(config.LOGO_PATH):
            st.image(config.LOGO_PATH, use_container_width=True)
        else:
            st.markdown(f"### 📊 {config.APP_NAME}")

        st.caption(config.APP_TAGLINE)
        st.divider()

        # The navigation menu. st.radio shows one clickable item per page.
        selected_page = st.radio(
            "Navigation",
            options=config.PAGES,
            label_visibility="collapsed",
        )

        st.divider()
        st.caption("Calculate • Reports • Email • Cache (Phases 1–13)")

    return selected_page


# ---------------------------------------------------------------------------
# STEP 3: Define a placeholder for each page.
# ---------------------------------------------------------------------------
# In later phases each of these functions will be replaced with the real page.
# For now they just show the page header and a "coming soon" note so you can
# confirm the navigation works.

def _placeholder(title, subtitle, coming_in_phase):
    """A small reusable placeholder body used by every page for now."""
    ui_helpers.render_page_header(title, subtitle)
    ui_helpers.render_glass_card(
        "🚧 Coming soon",
        f"This page is a placeholder for now. The real feature will be built in "
        f"<b>{coming_in_phase}</b>.",
    )


def page_dashboard():
    ui_helpers.render_page_header(
        "RDC Operations Reports & Calculators",
        "Centralized reporting and calculation tools for plant operations",
    )

    # A simple row of KPI cards. The numbers are all 0 for now because we have
    # not loaded any data yet. They become real in later phases.
    # Live row counts straight from the SQLite database.
    counts = database.get_table_counts()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        ui_helpers.render_kpi_card("Master Data Rows", f"{counts['master_data']:,}",
                                   "Employees synced/stored")
    with col2:
        ui_helpers.render_kpi_card("Backend Rows", f"{counts['backend_data']:,}",
                                   "Uploaded batching records")
    with col3:
        ui_helpers.render_kpi_card("Maintenance Rows",
                                   f"{counts['maintenance_cost']:,}",
                                   "Plant maintenance costs")
    with col4:
        ui_helpers.render_kpi_card("Calculated Rows",
                                   f"{counts['calculation_results']:,}",
                                   "Saved results")

    st.write("")  # small spacer

    # ---- Database status panel (proves Phase 2 tables exist) ----
    st.subheader("🗄️ Database status")
    st.caption(f"Local SQLite database: `{database.DB_PATH}`")

    status_df = pd.DataFrame(
        {"Table": list(counts.keys()), "Rows": list(counts.values())}
    )
    st.dataframe(status_df, hide_index=True, use_container_width=True)

    ui_helpers.render_success_message(
        "Database ready — all 9 tables are created. They are empty for now and "
        "will fill up as you sync Master Data (Phase 3) and upload files (Phase 4)."
    )


def page_data_uploader():
    """
    One combined page for all data sources. We use three tabs so the single
    "Data Uploader" sidebar button still gives access to all three jobs:
      1. Master Data Sync & Management
      2. Upload Backend Data
      3. Upload Maintenance Cost Data
    """
    ui_helpers.render_page_header(
        "Data Uploader",
        "Sync master data and upload backend & maintenance cost files in one place",
    )

    tab_master, tab_backend, tab_maintenance = st.tabs(
        [
            "🧑‍💼 Master Data Sync & Management",
            "📦 Backend Data",
            "🛠️ Maintenance Cost Data",
        ]
    )

    with tab_master:
        st.subheader("🔄 Sync from Google Sheets")

        # --- Mode indicator ---
        if google_sheets.credentials_exist():
            st.success(
                "🔐 **Private mode** — service account credentials found. "
                "Private sheets are supported.",
                icon="✅",
            )
        else:
            st.info(
                "🌐 **Public mode** — no service account credentials. "
                "Works perfectly for sheets shared as 'Anyone with the link can view'. "
                "To enable private sheets later, add `credentials/service_account.json`.",
                icon="ℹ️",
            )

        # --- Connection settings ---
        last_sync_info = google_sheets.get_last_sync_info()

        with st.form("gsheet_sync_form"):
            st.caption(
                "Paste your Google Sheet **URL or ID** below and click **Sync Now**. "
                "A full URL like `https://docs.google.com/spreadsheets/d/…/edit` also works."
            )
            sheet_id = st.text_input(
                "Google Sheet URL or ID",
                value=last_sync_info["sheet_id"],
                placeholder="https://docs.google.com/spreadsheets/d/1BxiMVs0.../edit",
                help="Paste the full URL from your browser, or just the Sheet ID.",
            )
            worksheet_name = st.text_input(
                "Worksheet / Tab Name",
                value=last_sync_info["worksheet"] or "Sheet1",
                placeholder="Sheet1",
                help="The exact name of the tab inside the spreadsheet (case-sensitive).",
            )
            sync_clicked = st.form_submit_button("🔄 Sync Now", type="primary")

        if sync_clicked:
            if not sheet_id.strip():
                ui_helpers.render_error_message("Please enter a Google Sheet ID.")
            else:
                tracker = ui_helpers.ProgressTracker(
                    ["📥 Fetching from Sheets", "🧹 Cleaning data", "💾 Saving to DB"]
                )

                # Step 0 — fetch rows from Google Sheets (network call, takes a moment)
                tracker.start(0)
                try:
                    clean_id = google_sheets.extract_sheet_id(sheet_id.strip())
                    df_sync = google_sheets.fetch_master_data(
                        clean_id, worksheet_name.strip()
                    )
                    tracker.complete(0, f"{len(df_sync):,} employees")
                except Exception as exc:
                    tracker.fail(0, "Failed")
                    ui_helpers.render_error_message(
                        f"Sync failed:<br><code>{exc}</code>"
                    )
                    st.stop()

                # Step 1 — cleaning (already done inside fetch_master_data)
                tracker.start(1)
                time.sleep(0.25)          # let the spinner flash briefly
                tracker.complete(1, "Columns validated")

                # Step 2 — save to SQLite
                tracker.start(2)
                now = _dt.now().isoformat(timespec="seconds")
                rows_to_save = [
                    {
                        "employee_code": str(row["Employee Code"]),
                        "employee_name": str(row["Employee Name"]),
                        "designation":   str(row["Designation"]),
                        "category":      str(row["Category"]),
                        "plant":         str(row["Plant"]),
                        "plant_code":    str(row["Plant Code"]),
                        "updated_at":    now,
                    }
                    for _, row in df_sync.iterrows()
                ]
                inserted = database.replace_table_rows("master_data", rows_to_save)
                database.set_setting("gsheet_id",         clean_id)
                database.set_setting("gsheet_worksheet",  worksheet_name.strip())
                database.set_setting("gsheet_last_sync",  now)
                database.set_setting("gsheet_last_count", str(inserted))
                tracker.complete(2, f"{inserted:,} rows saved")

                time.sleep(0.8)   # let the user see the all-green state
                st.rerun()

        # --- Startup auto-sync banner (shown if auto-sync just ran) ---
        startup_result = st.session_state.get("startup_sync_result")
        if startup_result:
            if startup_result["error"]:
                ui_helpers.render_error_message(
                    f"Auto-sync on startup failed:<br><code>{startup_result['error']}</code>"
                )
            else:
                ui_helpers.render_success_message(
                    f"✅ Auto-synced on startup — <b>{startup_result['rows_synced']}</b> "
                    f"employees loaded at {startup_result['synced_at']}."
                )
            # Clear it so it doesn't show again on the next rerun.
            del st.session_state["startup_sync_result"]

        # --- Last sync info ---
        if last_sync_info["last_sync"]:
            st.info(
                f"Last synced: **{last_sync_info['last_sync']}** · "
                f"**{last_sync_info['last_count']}** rows",
                icon="🕐",
            )

        # --- Auto-sync toggle (saved permanently to app_settings) ---
        current_auto = database.get_setting("gsheet_auto_sync", "false") == "true"
        new_auto = st.toggle(
            "Auto-sync every time the app starts",
            value=current_auto,
            help="When ON, the app fetches the latest data from Google Sheets "
                 "automatically each time you open it — no need to click Sync Now.",
        )
        if new_auto != current_auto:
            database.set_setting("gsheet_auto_sync", "true" if new_auto else "false")
            st.rerun()

        st.divider()

        # --- Current master data table ---
        st.subheader("👥 Current Master Data")
        _m_count = database.get_table_counts().get("master_data", 0)
        if _m_count == 0:
            ui_helpers.render_glass_card(
                "No data yet",
                "The master_data table is empty. Enter your Sheet ID above and click "
                "<b>Sync Now</b> to load employees.",
            )
        else:
            master_df = database.read_table_limited(
                "master_data", order_by="employee_code", limit=500
            )
            st.caption(
                f"Showing first 500 of **{_m_count:,}** employees · "
                "use Export to see all"
            )
            st.dataframe(master_df.drop(columns=["id"], errors="ignore"),
                         hide_index=True, use_container_width=True)

        st.divider()
        st.subheader("✏️ Manage Employees (Add / Edit / Delete)")
        ui_helpers.render_warning_message(
            "Manual changes here can be <b>overwritten the next time you sync</b> "
            "from Google Sheets (a sync does a full replace of master data)."
        )

        # List of current employee codes for the Edit / Delete pickers.
        _codes_df = database.read_table_limited(
            "master_data", order_by="employee_code", limit=100000
        )
        _codes = (_codes_df["employee_code"].astype(str).tolist()
                  if not _codes_df.empty else [])

        m_add, m_edit, m_del, m_log = st.tabs(
            ["➕ Add", "✏️ Edit", "🗑️ Delete", "📜 Change log"]
        )

        # --- Add ----------------------------------------------------------------
        with m_add:
            with st.form("add_emp_form", clear_on_submit=True):
                a1, a2 = st.columns(2)
                with a1:
                    add_code  = st.text_input("Employee Code *")
                    add_name  = st.text_input("Employee Name")
                    add_desig = st.text_input("Designation")
                with a2:
                    add_cat   = st.selectbox("Category", options=config.CATEGORIES)
                    add_plant = st.text_input("Plant")
                    add_pcode = st.text_input("Plant Code")
                add_submit = st.form_submit_button("➕ Add employee", type="primary")
            if add_submit:
                try:
                    database.add_employee(add_code, add_name, add_desig,
                                          add_cat, add_plant, add_pcode)
                    ui_helpers.render_success_message(
                        f"Added employee <b>{add_code}</b>."
                    )
                    st.rerun()
                except ValueError as err:
                    ui_helpers.render_error_message(str(err))

        # --- Edit ---------------------------------------------------------------
        with m_edit:
            if not _codes:
                ui_helpers.render_glass_card(
                    "No employees", "Add or sync employees first."
                )
            else:
                sel_code = st.selectbox("Select Employee Code", options=_codes,
                                        key="edit_pick")
                emp = database.get_employee(sel_code)
                if emp:
                    with st.form("edit_emp_form"):
                        e1, e2 = st.columns(2)
                        with e1:
                            ed_name  = st.text_input("Employee Name",
                                                     value=emp["employee_name"] or "")
                            ed_desig = st.text_input("Designation",
                                                     value=emp["designation"] or "")
                            ed_plant = st.text_input("Plant",
                                                     value=emp["plant"] or "")
                        with e2:
                            _cats = config.CATEGORIES
                            _idx = (_cats.index(emp["category"])
                                    if emp["category"] in _cats else 0)
                            ed_cat   = st.selectbox("Category", options=_cats,
                                                    index=_idx)
                            ed_pcode = st.text_input("Plant Code",
                                                     value=emp["plant_code"] or "")
                        ed_submit = st.form_submit_button("💾 Save changes",
                                                          type="primary")
                    if ed_submit:
                        try:
                            database.update_employee(sel_code, ed_name, ed_desig,
                                                     ed_cat, ed_plant, ed_pcode)
                            ui_helpers.render_success_message(
                                f"Updated employee <b>{sel_code}</b>."
                            )
                            st.rerun()
                        except ValueError as err:
                            ui_helpers.render_error_message(str(err))

        # --- Delete -------------------------------------------------------------
        with m_del:
            if not _codes:
                ui_helpers.render_glass_card("No employees", "Nothing to delete yet.")
            else:
                del_code = st.selectbox("Select Employee Code", options=_codes,
                                        key="del_pick")
                emp = database.get_employee(del_code)
                if emp:
                    st.dataframe(
                        pd.DataFrame([{
                            "Code": emp["employee_code"],
                            "Name": emp["employee_name"],
                            "Designation": emp["designation"],
                            "Category": emp["category"],
                            "Plant": emp["plant"],
                            "Plant Code": emp["plant_code"],
                        }]),
                        hide_index=True, use_container_width=True,
                    )
                    confirm = st.checkbox(
                        "Yes, I really want to delete this employee",
                        key="del_confirm",
                    )
                    if st.button("🗑️ Delete employee", disabled=not confirm,
                                 key="del_btn"):
                        try:
                            database.delete_employee(del_code)
                            ui_helpers.render_success_message(
                                f"Deleted employee <b>{del_code}</b>."
                            )
                            st.rerun()
                        except ValueError as err:
                            ui_helpers.render_error_message(str(err))

        # --- Change log ---------------------------------------------------------
        with m_log:
            log_df = database.read_table("master_data_change_log", order_by="id DESC")
            if log_df.empty:
                ui_helpers.render_glass_card(
                    "No changes yet",
                    "Add, edit or delete an employee and it will be recorded here.",
                )
            else:
                show_log = (
                    log_df.rename(columns={
                        "action_type":    "Action",
                        "employee_code":  "Emp Code",
                        "old_value_json": "Old values",
                        "new_value_json": "New values",
                        "changed_at":     "When",
                        "changed_by":     "By",
                    }).drop(columns=["id"], errors="ignore")
                )
                st.caption(f"{len(show_log):,} change(s), newest first")
                st.dataframe(show_log, hide_index=True, use_container_width=True)

    with tab_backend:
        st.subheader("📦 Upload Backend Data")
        st.caption(
            f"Required columns: **{', '.join(config.BACKEND_REQUIRED_COLUMNS)}**  "
            "· Extra columns in the file are ignored."
        )

        uploaded_backend = st.file_uploader(
            "Choose the Backend Data Excel file (.xlsx)",
            type=["xlsx"],
            key="backend_file_uploader",
        )

        if uploaded_backend:
            # Cache the parsed dataframe in session_state so re-renders
            # (caused by radio clicks, etc.) don't re-read the large Excel file.
            _bkey = f"backend_{uploaded_backend.name}_{uploaded_backend.size}"
            if st.session_state.get("_backend_cache_key") != _bkey:
                try:
                    _df, _w = data_loader.load_backend_data(uploaded_backend)
                    st.session_state["_backend_cache_key"] = _bkey
                    st.session_state["_backend_df"]   = _df
                    st.session_state["_backend_warns"] = _w
                except ValueError as err:
                    ui_helpers.render_error_message(str(err))
                    st.stop()

            df_backend = st.session_state["_backend_df"]
            warnings   = st.session_state["_backend_warns"]

            for w in warnings:
                ui_helpers.render_warning_message(w)

            st.success(
                f"✅ File read — **{len(df_backend):,} rows** ready to save."
            )

            with st.expander("🔍 Preview (first 10 rows)", expanded=True):
                st.dataframe(df_backend.head(10), hide_index=True,
                             use_container_width=True)

            replace_mode = st.radio(
                "Upload mode",
                options=["Replace existing data", "Append to existing data"],
                index=0, horizontal=True,
                help="Replace clears the previous upload first.",
            )

            if st.button("💾 Save to Database", type="primary",
                         key="save_backend_btn"):
                replace = replace_mode == "Replace existing data"
                tracker = ui_helpers.ProgressTracker(
                    ["📂 Reading file", "🧹 Cleaning data", "💾 Saving to DB"]
                )
                tracker.start(0)
                tracker.complete(0, f"{len(df_backend):,} rows")
                tracker.start(1)
                time.sleep(0.2)
                lbl = f"{len(warnings)} skipped" if warnings else "✓ All clean"
                tracker.complete(1, lbl)
                tracker.start(2)
                saved = data_loader.save_backend_data(
                    df_backend, source_file=uploaded_backend.name, replace=replace
                )
                tracker.complete(2, f"{saved:,} rows {'replaced' if replace else 'appended'}")
                # Clear cache so the summary refreshes after save
                st.session_state.pop("_backend_cache_key", None)
                time.sleep(0.8)
                st.rerun()

        # --- Current data summary (never loads all rows) ---
        st.divider()
        st.subheader("📋 Current Backend Data")
        _b_count = database.get_table_counts().get("backend_data", 0)
        if _b_count == 0:
            ui_helpers.render_glass_card(
                "No data yet",
                "Upload an Excel file above to populate backend data.",
            )
        else:
            _b_first = database.read_table_limited("backend_data",
                                                    order_by="date", limit=1)
            _b_last  = database.read_table_limited("backend_data",
                                                    order_by="date DESC", limit=1)
            _b_preview = database.read_table_limited("backend_data",
                                                      order_by="date", limit=200)
            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Total rows", f"{_b_count:,}")
            col_b.metric("Earliest date", _b_first["date"].iloc[0])
            col_c.metric("Latest date",   _b_last["date"].iloc[0])
            st.caption(f"Showing first 200 of **{_b_count:,}** rows")
            st.dataframe(_b_preview.drop(columns=["id"], errors="ignore"),
                         hide_index=True, use_container_width=True)

    with tab_maintenance:
        st.subheader("🛠️ Upload Maintenance Cost Data")
        st.caption(
            f"Required columns: **{', '.join(config.MAINTENANCE_REQUIRED_COLUMNS)}**  "
            "· Each upload replaces the previous data."
        )

        uploaded_maint = st.file_uploader(
            "Choose the Maintenance Cost Excel file (.xlsx)",
            type=["xlsx"],
            key="maint_file_uploader",
        )

        if uploaded_maint:
            _mkey = f"maint_{uploaded_maint.name}_{uploaded_maint.size}"
            if st.session_state.get("_maint_cache_key") != _mkey:
                try:
                    _dfm, _wm = data_loader.load_maintenance_cost(uploaded_maint)
                    st.session_state["_maint_cache_key"] = _mkey
                    st.session_state["_maint_df"]   = _dfm
                    st.session_state["_maint_warns"] = _wm
                except ValueError as err:
                    ui_helpers.render_error_message(str(err))
                    st.stop()

            df_maint = st.session_state["_maint_df"]
            warns_m  = st.session_state["_maint_warns"]

            for w in warns_m:
                ui_helpers.render_warning_message(w)

            st.success(
                f"✅ File read — **{len(df_maint):,} plants** ready to save."
            )

            with st.expander("🔍 Preview (all rows)", expanded=True):
                st.dataframe(df_maint, hide_index=True, use_container_width=True)

            if st.button("💾 Save to Database", type="primary",
                         key="save_maint_btn"):
                tracker = ui_helpers.ProgressTracker(
                    ["📂 Reading file", "🧹 Cleaning data", "💾 Saving to DB"]
                )
                tracker.start(0)
                tracker.complete(0, f"{len(df_maint):,} plants")
                tracker.start(1)
                time.sleep(0.2)
                lbl_m = f"{len(warns_m)} skipped" if warns_m else "✓ All clean"
                tracker.complete(1, lbl_m)
                tracker.start(2)
                saved_m = data_loader.save_maintenance_cost(df_maint)
                tracker.complete(2, f"{saved_m:,} plants saved")
                st.session_state.pop("_maint_cache_key", None)
                time.sleep(0.8)
                st.rerun()

        # --- Current data summary ---
        st.divider()
        st.subheader("📋 Current Maintenance Cost Data")
        maint_df = database.read_table("maintenance_cost", order_by="plant_code")
        if maint_df.empty:
            ui_helpers.render_glass_card(
                "No data yet",
                "Upload an Excel file above to populate maintenance cost data.",
            )
        else:
            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Plants", f"{len(maint_df):,}")
            col_b.metric("Avg cost (Rs/cum)",
                         f"{maint_df['ytd_maintenance_cost'].mean():.2f}")
            col_c.metric(
                f"Above >{config.MAINTENANCE_COST_THRESHOLD}",
                int((maint_df["ytd_maintenance_cost"]
                     > config.MAINTENANCE_COST_THRESHOLD).sum()),
            )
            st.dataframe(maint_df.drop(columns=["id"], errors="ignore"),
                         hide_index=True, use_container_width=True)


def _style_result_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Return a same-shaped DataFrame of CSS strings (the row highlighting).

    Spec rule:
      - Incentive Amount > 0  -> light green row
      - Deduction Amount > 0  -> light red row
      - If BOTH are > 0       -> light red wins (deduction takes priority)
      - Both 0                -> normal (no colour)

    The amount columns are found by their name PREFIX so this keeps working
    even when the deduction column is renamed to "Deduction Amount @ Rs 10".
    """
    styles = pd.DataFrame("", index=df.index, columns=df.columns)

    inc_col = next((c for c in df.columns if c.startswith("Incentive Amount")), None)
    ded_col = next((c for c in df.columns if c.startswith("Deduction Amount")), None)

    has_inc = (df[inc_col] > 0) if inc_col else pd.Series(False, index=df.index)
    has_ded = (df[ded_col] > 0) if ded_col else pd.Series(False, index=df.index)

    # Green for incentive-only rows first...
    styles.loc[has_inc & ~has_ded, :] = "background-color: #D4EDDA"  # light green
    # ...then red for ANY row with a deduction (this also overrides "both").
    styles.loc[has_ded, :] = "background-color: #F8D7DA"             # light red
    return styles


# ---------------------------------------------------------------------------
# Shared result-table helpers (used by BOTH the Calculate and View Reports pages)
# ---------------------------------------------------------------------------
# The columns we show (in order) and their friendly display names.
RESULT_SHOW_COLS = [
    "employee_code", "employee_name", "designation",
    "plant", "plant_code",
    "total_quantity", "ytd_maintenance_cost",
    "incentive_eligible", "incentive_rate", "incentive_amount",
    "deduction_target", "shortfall_quantity", "deduction_amount",
    "remarks",
]
RESULT_RENAME = {
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


def _sort_report_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sort result rows for display/export:
      1. Deduction (red) rows first — highest → lowest deduction amount.
      2. Incentive (green) rows next — highest → lowest incentive amount.
      3. Plain rows last.
    A row with BOTH counts as a deduction row (red takes priority).
    Works on the internal column names (incentive_amount / deduction_amount).
    """
    out = df.copy()
    out["_band"] = 2                                    # plain rows
    out.loc[out["incentive_amount"] > 0, "_band"] = 1   # green
    out.loc[out["deduction_amount"] > 0, "_band"] = 0   # red (priority)
    out = out.sort_values(
        by=["_band", "deduction_amount", "incentive_amount"],
        ascending=[True, False, False],
    ).drop(columns="_band")
    return out


def page_calculate():
    import calendar as _cal

    ui_helpers.render_page_header(
        "Calculate Incentive & Deduction",
        "Choose a date range and run the full category-wise incentive & deduction calculation",
    )

    # ── Guard: need backend data ─────────────────────────────────────────────
    available = calculator.get_available_months()
    if not available:
        ui_helpers.render_glass_card(
            "No backend data",
            "Upload a Backend Data file on the <b>Data Uploader</b> page first.",
        )
        return

    # ── Date range picker ────────────────────────────────────────────────────
    # Default: first and last date found in backend_data
    conn = database.get_connection()
    try:
        minmax = pd.read_sql_query(
            "SELECT MIN(date) AS mn, MAX(date) AS mx FROM backend_data", conn
        )
    finally:
        conn.close()

    from datetime import date as _date
    default_from = pd.to_datetime(minmax["mn"].iloc[0]).date()
    default_to   = pd.to_datetime(minmax["mx"].iloc[0]).date()

    col_f, col_t, col_btn = st.columns([2, 2, 3])
    with col_f:
        from_date = st.date_input("From Date", value=default_from, key="calc_from")
    with col_t:
        to_date = st.date_input("To Date", value=default_to, key="calc_to")

    if from_date > to_date:
        ui_helpers.render_error_message("'From Date' must be on or before 'To Date'.")
        return

    # Row count for the selected range
    conn = database.get_connection()
    try:
        cnt = pd.read_sql_query(
            "SELECT COUNT(*) AS c FROM backend_data WHERE date >= ? AND date <= ?",
            conn, params=(str(from_date), str(to_date)),
        )
    finally:
        conn.close()
    period_rows = int(cnt["c"].iloc[0])

    with col_btn:
        st.write("")
        st.caption(f"**{period_rows:,}** backend rows in selected range")

    if period_rows == 0:
        ui_helpers.render_warning_message(
            "No backend data found for the selected date range."
        )
        return

    last = calculator.get_last_calculation_info()
    if last["ran_at"]:
        st.caption(
            f"Last run: {last['ran_at']}  ·  "
            f"{last['mapped']} employees  ·  {last['unmapped']} unmapped"
        )

    # ── Run button ───────────────────────────────────────────────────────────
    st.write("")
    if st.button("⚡ Run Calculation", type="primary", key="run_calc_btn"):
        tracker = ui_helpers.ProgressTracker([
            "📦 Aggregating data",
            "🧑‍💼 Matching employees",
            "💰 Calculating",
            "💾 Saving",
        ])
        tracker.start(0)
        time.sleep(0.2)
        tracker.complete(0, f"{period_rows:,} rows")

        tracker.start(1)
        time.sleep(0.1)
        tracker.complete(1, f"{database.get_table_counts().get('master_data',0):,} in master")

        tracker.start(2)
        result = calculator.run_calculation(
            month=from_date.month,
            year=from_date.year,
            start_date=str(from_date),
            end_date=str(to_date),
        )
        if result["error"]:
            tracker.fail(2, "Error")
            ui_helpers.render_error_message(result["error"])
            st.stop()
        tracker.complete(2, f"{result['mapped']:,} employees")

        tracker.start(3)
        time.sleep(0.2)
        tracker.complete(3, f"done · {result['unmapped']} unmapped")
        time.sleep(0.8)
        st.rerun()

    st.divider()

    # ── Results ──────────────────────────────────────────────────────────────
    if database.get_table_counts().get("calculation_results", 0) == 0:
        ui_helpers.render_glass_card(
            "No results yet",
            "Select a date range and click <b>Run Calculation</b> above.",
        )
        return

    results_df = database.read_table("calculation_results")

    # ── Summary KPI row ───────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Employees", f"{len(results_df):,}")
    c2.metric("Incentive Eligible", f"{(results_df['incentive_eligible']=='Yes').sum():,}")
    c3.metric("Total Incentive", f"₹{results_df['incentive_amount'].sum():,.0f}")
    c4.metric("Total Deduction", f"₹{results_df['deduction_amount'].sum():,.0f}")

    st.write("")

    # ── Legend ────────────────────────────────────────────────────────────────
    _result_legend()
    st.write("")

    # ── Category-wise tabs ────────────────────────────────────────────────────
    # Map tab label → list of categories it covers (from config.REPORT_SHEETS)
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

    # Only show tabs that have at least one row in the results
    active_tabs = {
        label: cats
        for label, cats in CAT_TABS.items()
        if not results_df[results_df["category"].isin(cats)].empty
    }

    # Display columns + friendly names come from the shared module constants.
    SHOW_COLS = RESULT_SHOW_COLS
    RENAME = RESULT_RENAME

    if active_tabs:
        tabs = st.tabs(list(active_tabs.keys()))
        for tab_widget, (label, cats) in zip(tabs, active_tabs.items()):
            with tab_widget:
                grp = results_df[results_df["category"].isin(cats)].copy()

                # Category summary row at top of tab
                c_a, c_b, c_c, c_d = st.columns(4)
                c_a.metric("Employees", f"{len(grp):,}")
                c_b.metric("Eligible", f"{(grp['incentive_eligible']=='Yes').sum():,}")
                c_c.metric("Incentive", f"₹{grp['incentive_amount'].sum():,.0f}")
                c_d.metric("Deduction", f"₹{grp['deduction_amount'].sum():,.0f}")

                # Show the deduction RATE in the column header (spec rule).
                # Every category inside one tab shares the same rate, so we can
                # label the header e.g. "Deduction Amount @ Rs 10".
                ded_rates = {
                    config.DEDUCTION_RULES.get(c, {}).get("rate", 0) for c in cats
                }
                rename_tab = dict(RENAME)
                if len(ded_rates) == 1 and next(iter(ded_rates)):
                    rename_tab["deduction_amount"] = (
                        f"Deduction Amount @ Rs {next(iter(ded_rates))}"
                    )

                # Sort: deduction (red) rows first, then incentive (green), then
                # plain — biggest amount on top within each band.
                grp = _sort_report_rows(grp)

                # Build display df
                show = (
                    grp[[c for c in SHOW_COLS if c in grp.columns]]
                    .rename(columns=rename_tab)
                    .reset_index(drop=True)
                )

                # Apply row highlighting
                styled = show.style.apply(_style_result_rows, axis=None)
                st.dataframe(styled, hide_index=True, use_container_width=True)

    # ── Unmapped employees — small footnote ──────────────────────────────────
    unmapped_df = database.read_table("unmapped_employees")
    if not unmapped_df.empty:
        st.divider()
        with st.expander(
            f"⚠️ {len(unmapped_df)} unmapped employee(s) — codes in backend data not found in master data",
            expanded=False,
        ):
            st.dataframe(
                unmapped_df.drop(columns=["id"], errors="ignore"),
                hide_index=True, use_container_width=True,
            )


def _result_legend():
    """The little green/red colour key used on the results pages."""
    st.markdown(
        '<span style="background:#D4EDDA;padding:4px 10px;border-radius:6px;'
        'font-size:13px;margin-right:8px">🟢 Incentive earned</span>'
        '<span style="background:#F8D7DA;padding:4px 10px;border-radius:6px;'
        'font-size:13px">🔴 Deduction applied (takes priority if both)</span>',
        unsafe_allow_html=True,
    )


def page_view_reports():
    ui_helpers.render_page_header(
        "View Reports",
        "Pick a date range, then browse the report with live multi-select filters",
    )

    # ── Report period (From / To date range) ─────────────────────────────────
    # View Reports builds the report LIVE for the chosen range from Backend Data,
    # so you can browse any period. It does NOT overwrite the saved calculation
    # (it calls the engine with persist=False). It needs backend data to work.
    available = calculator.get_available_months()
    if not available:
        ui_helpers.render_glass_card(
            "No backend data",
            "Upload a Backend Data file on the <b>Data Uploader</b> page first.",
        )
        return

    conn = database.get_connection()
    try:
        minmax = pd.read_sql_query(
            "SELECT MIN(date) AS mn, MAX(date) AS mx FROM backend_data", conn
        )
    finally:
        conn.close()
    default_from = pd.to_datetime(minmax["mn"].iloc[0]).date()
    default_to   = pd.to_datetime(minmax["mx"].iloc[0]).date()

    st.subheader("📅 Report period")
    col_f, col_t = st.columns(2)
    with col_f:
        from_date = st.date_input("From Date", value=default_from, key="vr_from")
    with col_t:
        to_date = st.date_input("To Date", value=default_to, key="vr_to")

    if from_date > to_date:
        ui_helpers.render_error_message("'From Date' must be on or before 'To Date'.")
        return

    # Recompute ONLY when the date range changes — so changing the filters below
    # is instant and never re-runs the calculation.
    range_key = f"{from_date}|{to_date}"
    if (st.session_state.get("vr_range_key") != range_key
            or "vr_results_df" not in st.session_state):
        with st.spinner("Calculating the report for the selected period…"):
            res = calculator.run_calculation(
                month=from_date.month, year=from_date.year,
                start_date=str(from_date), end_date=str(to_date),
                persist=False,                      # never overwrite saved results
            )
        if res["error"]:
            st.session_state["vr_results_df"]  = pd.DataFrame()
            st.session_state["vr_unmapped_df"] = pd.DataFrame()
            st.session_state["vr_range_note"]  = res["error"]
        else:
            st.session_state["vr_results_df"]  = pd.DataFrame(res["results_rows"])
            st.session_state["vr_unmapped_df"] = pd.DataFrame(res["unmapped_rows"])
            st.session_state["vr_range_note"]  = ""
        st.session_state["vr_range_key"] = range_key

    results_df = st.session_state["vr_results_df"]
    if results_df.empty:
        ui_helpers.render_warning_message(
            st.session_state.get("vr_range_note")
            or "No mapped employees were found for this date range."
        )
        return

    st.caption(
        f"Live report for **{from_date} → {to_date}**  ·  "
        f"**{len(results_df):,}** employees  ·  _(preview only — not saved)_"
    )
    st.divider()

    # ── Phase 8: dynamic multi-select filters ────────────────────────────────
    # The results table carries the Master Data columns (Category, Designation,
    # Plant, Plant Code), so the dropdown options come straight from real data.
    st.subheader("🔎 Filters")
    filtered = ui_helpers.render_dynamic_filters(
        results_df,
        [
            ("category",   "Category"),
            ("designation", "Designation"),
            ("plant",      "Plant"),
            ("plant_code", "Plant Code"),
        ],
        key_prefix="vr",
    )

    # Extra filters: eligibility, outcome, and a free-text search box.
    col1, col2, col3 = st.columns(3)
    with col1:
        elig = st.selectbox(
            "Incentive eligibility",
            ["All", "Eligible (Yes)", "Not eligible (No)"],
            key="vr_elig",
        )
    with col2:
        outcome = st.selectbox(
            "Outcome",
            ["All", "Has incentive", "Has deduction", "Has both", "Neither"],
            key="vr_outcome",
        )
    with col3:
        search = st.text_input(
            "Search Emp Code / Name",
            key="vr_search",
            placeholder="e.g. 1023 or Ramesh",
        )

    # Apply eligibility filter
    if elig == "Eligible (Yes)":
        filtered = filtered[filtered["incentive_eligible"] == "Yes"]
    elif elig == "Not eligible (No)":
        filtered = filtered[filtered["incentive_eligible"] == "No"]

    # Apply outcome filter
    if outcome == "Has incentive":
        filtered = filtered[filtered["incentive_amount"] > 0]
    elif outcome == "Has deduction":
        filtered = filtered[filtered["deduction_amount"] > 0]
    elif outcome == "Has both":
        filtered = filtered[(filtered["incentive_amount"] > 0) &
                            (filtered["deduction_amount"] > 0)]
    elif outcome == "Neither":
        filtered = filtered[(filtered["incentive_amount"] == 0) &
                            (filtered["deduction_amount"] == 0)]

    # Apply text search (matches Employee Code OR Name, case-insensitive)
    if search.strip():
        s = search.strip().lower()
        mask = (
            filtered["employee_code"].astype(str).str.lower().str.contains(s) |
            filtered["employee_name"].astype(str).str.lower().str.contains(s)
        )
        filtered = filtered[mask]

    # Build a readable "applied filters" description (used in the table caption,
    # the Excel Summary sheet and the email body).
    filter_parts = []
    for col_name, lbl in [("category", "Category"), ("designation", "Designation"),
                          ("plant", "Plant"), ("plant_code", "Plant Code")]:
        sel = st.session_state.get(f"vr_{col_name}")
        if sel:
            filter_parts.append(f"{lbl}: {', '.join(map(str, sel))}")
    if elig != "All":
        filter_parts.append(f"Eligibility: {elig}")
    if outcome != "All":
        filter_parts.append(f"Outcome: {outcome}")
    if search.strip():
        filter_parts.append(f"Search: {search.strip()}")
    applied_filters = " | ".join(filter_parts) if filter_parts else "None"

    st.divider()

    # ── Generate the report (explicit button) ────────────────────────────────
    # Nothing below is shown until you press Generate, and what IS shown always
    # matches the date range + filters that were active at that moment — so the
    # table, summary, Excel and email body can never drift out of sync.
    selection_sig = f"{from_date}|{to_date}|{applied_filters}|{len(filtered)}"
    if st.button("🔄 Generate Report", type="primary", key="vr_generate"):
        st.session_state["vr_report"] = {
            "filtered":        filtered.copy(),
            "unmapped":        st.session_state.get("vr_unmapped_df", pd.DataFrame()).copy(),
            "from_date":       from_date,
            "to_date":         to_date,
            "applied_filters": applied_filters,
            "total_rows":      len(results_df),
            "sig":             selection_sig,
        }

    report = st.session_state.get("vr_report")
    if not report:
        ui_helpers.render_glass_card(
            "Ready when you are",
            "Choose your <b>date range</b> and <b>filters</b> above, then click "
            "<b>🔄 Generate Report</b> to build the table, summary and export options.",
        )
        return

    # If the date range / filters changed since the last Generate, nudge the user.
    if report["sig"] != selection_sig:
        ui_helpers.render_warning_message(
            "Your date range or filters changed since this report was generated. "
            "Click <b>🔄 Generate Report</b> again to refresh it."
        )

    # From here on, use the GENERATED snapshot for everything (table, KPIs,
    # Excel, email) so they always agree with each other.
    filtered        = report["filtered"]
    unmapped_df     = report["unmapped"]
    from_date       = report["from_date"]
    to_date         = report["to_date"]
    applied_filters = report["applied_filters"]
    gen_total_rows  = report["total_rows"]

    # ── Summary KPIs (reflect the generated report) ──────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Employees (filtered)", f"{len(filtered):,}")
    c2.metric("Incentive eligible",
              f"{(filtered['incentive_eligible'] == 'Yes').sum():,}")
    c3.metric("Total Incentive", f"₹{filtered['incentive_amount'].sum():,.0f}")
    c4.metric("Total Deduction", f"₹{filtered['deduction_amount'].sum():,.0f}")

    st.write("")
    _result_legend()
    st.write("")

    if filtered.empty:
        ui_helpers.render_warning_message(
            "No rows match the current filters. Clear a filter to see more."
        )
        return

    # ── Sorted + colour-highlighted table ────────────────────────────────────
    sorted_df = _sort_report_rows(filtered)
    show = (
        sorted_df[[c for c in RESULT_SHOW_COLS if c in sorted_df.columns]]
        .rename(columns=RESULT_RENAME)
        .reset_index(drop=True)
    )
    st.caption(f"Showing **{len(show):,}** of **{gen_total_rows:,}** total rows")
    styled = show.style.apply(_style_result_rows, axis=None)
    st.dataframe(styled, hide_index=True, use_container_width=True)

    # ─────────────────────────────────────────────────────────────────────────
    # 📤 EXPORT & EMAIL  (merged into this page — no separate sidebar item)
    # ─────────────────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("📤 Export & Email")

    # filtered / unmapped_df / applied_filters all come from the generated
    # snapshot above. Validation errors are date-independent, so read them live.
    validation_df = database.read_table("validation_errors")

    # Friendly month label for the subject/body: "May 2026" when the whole range
    # sits in one calendar month, otherwise the explicit date range.
    if (from_date.year, from_date.month) == (to_date.year, to_date.month):
        month_label = from_date.strftime("%B %Y")
    else:
        month_label = f"{from_date.strftime('%d %b %Y')} to {to_date.strftime('%d %b %Y')}"

    tab_excel, tab_csv, tab_email = st.tabs(
        ["📊 Excel report", "📄 CSV", "✉️ Email"]
    )

    # --- Excel (full multi-sheet, formatted) ---------------------------------
    with tab_excel:
        st.caption(
            "One workbook: **Summary** + a sheet per category + **Unmapped** + "
            "**Validation Errors**. Colour-highlighted, header frozen, and it "
            "**respects the filters and date range above**."
        )
        meta = {
            "generated_on":    _dt.now().isoformat(timespec="seconds"),
            "date_range":      f"{from_date} to {to_date}",
            "applied_filters": applied_filters,
            "total_employees": int(len(filtered)),
            "total_quantity":  float(filtered["total_quantity"].sum()),
            "total_incentive": float(filtered["incentive_amount"].sum()),
            "total_deduction": float(filtered["deduction_amount"].sum()),
            "unmapped_count":  int(len(unmapped_df)),
            "validation_count": int(len(validation_df)),
        }
        try:
            xlsx_bytes = report_generator.generate_excel_report(
                filtered, unmapped_df, validation_df, meta
            )
            st.download_button(
                "⬇️ Download Excel report (.xlsx)",
                data=xlsx_bytes,
                file_name=f"incentive_report_{from_date}_to_{to_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument."
                     "spreadsheetml.sheet",
                key="vr_xlsx",
                type="primary",
            )
        except Exception as exc:  # noqa: BLE001 - show any build error to the user
            ui_helpers.render_error_message(f"Could not build the Excel file:<br><code>{exc}</code>")

    # --- CSV (quick, single sheet of the filtered view) ----------------------
    with tab_csv:
        st.caption("A plain CSV of exactly the filtered table shown above.")
        csv_bytes = show.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download filtered data (CSV)",
            data=csv_bytes,
            file_name=f"incentive_report_{from_date}_to_{to_date}.csv",
            mime="text/csv",
            key="vr_download",
        )

    # --- Email (sends the Excel above as an attachment) ----------------------
    with tab_email:
        # Read config once; derive "configured" from it (no extra DB reads).
        email_cfg = email_helper.get_smtp_config()
        email_ready = bool(email_cfg["host"] and email_cfg["sender"]
                           and email_cfg["password"])
        if not email_ready:
            ui_helpers.render_warning_message(
                "Email is not set up yet. Add your SMTP settings on the "
                "<b>Settings</b> page, then come back here to send."
            )

        e_to = st.text_input("To (comma-separated)", value=email_cfg["default_to"],
                             key="vr_email_to", placeholder="name@company.com")
        e_cc = st.text_input("CC (optional)", value=email_cfg["default_cc"],
                             key="vr_email_cc", placeholder="manager@company.com")

        # Subject + body follow the company's standard monthly-mail format
        # (see the Mail Reference), with the report month filled in automatically.
        # IMPORTANT: a keyed text widget keeps its first cached value and ignores
        # later `value=` changes, which made the period go stale. We tie the keys
        # to the report month so they refresh when the period changes, while still
        # letting you edit the text for the same report.
        default_subject = email_helper.compose_report_subject(month_label)
        default_body    = email_helper.compose_report_body(month_label)
        e_subject = st.text_input(
            "Subject", value=default_subject,
            key=f"vr_email_subject::{month_label}",
        )
        e_body = st.text_area(
            "Message", value=default_body, height=300,
            key=f"vr_email_body::{month_label}",
        )
        include_tables = st.checkbox(
            "📋 Also paste the report tables inside the email body (optional)",
            value=False, key="vr_email_tables",
            help="Off by default: the email is the message text above, with the "
                 "report attached as Excel. Tick this to also paste a colour-coded "
                 "table per section into the body (makes the email much larger).",
        )

        if st.button("✉️ Send report", key="vr_send_email", type="primary",
                     disabled=not email_ready):
            try:
                fname = f"incentive_report_{from_date}_to_{to_date}.xlsx"
                xlsx_for_email = report_generator.generate_excel_report(
                    filtered, unmapped_df, validation_df, meta
                )
                # When requested, render the per-section tables into the email
                # body (HTML) after the message text — the Excel is still attached.
                html_body = None
                if include_tables:
                    tables_html = report_generator.build_email_tables_html(filtered)
                    html_body = email_helper.wrap_html_body(e_body, tables_html)
                with st.spinner("Sending email…"):
                    res = email_helper.send_report_email(
                        to_emails=e_to, cc_emails=e_cc, subject=e_subject,
                        body=e_body, attachment_bytes=xlsx_for_email,
                        attachment_name=fname, html_body=html_body,
                    )
                if res["success"]:
                    ui_helpers.render_success_message(f"Report emailed to {e_to}.")
                else:
                    ui_helpers.render_error_message(f"Email failed: {res['error']}")
            except Exception as exc:  # noqa: BLE001
                ui_helpers.render_error_message(f"Could not send: {exc}")

        with st.expander("📜 Recent email log"):
            elog = database.read_table("email_log", order_by="id DESC")
            if elog.empty:
                st.caption("No emails sent yet.")
            else:
                st.dataframe(
                    elog.drop(columns=["id"], errors="ignore").head(20),
                    hide_index=True, use_container_width=True,
                )


def page_validation():
    ui_helpers.render_page_header(
        "Error / Validation Report",
        "Run data checks on Master Data, Backend Data and Maintenance Cost before calculating",
    )

    # ── Last run info ────────────────────────────────────────────────────────
    last = validations.get_last_validation_info()
    if last["last_run"]:
        err_n = int(last["error_count"])
        if err_n == 0:
            ui_helpers.render_success_message(
                f"Last validation ran at **{last['last_run']}** — ✅ No errors found."
            )
        else:
            ui_helpers.render_warning_message(
                f"Last validation ran at **{last['last_run']}** — "
                f"⚠️ **{err_n}** error(s) found. Review below."
            )
    else:
        ui_helpers.render_glass_card(
            "Validation not run yet",
            "Click <b>Run Validation</b> below to check all three data sources.",
        )

    st.write("")

    # ── Run button ───────────────────────────────────────────────────────────
    if st.button("▶️ Run Validation Now", type="primary", key="run_validation_btn"):
        tracker = ui_helpers.ProgressTracker([
            "🧑‍💼 Master Data",
            "📦 Backend Data",
            "🛠️ Maintenance Cost",
            "💾 Saving errors",
        ])

        tracker.start(0)
        _errs: list = []
        m_count = validations._validate_master_data(_errs)
        tracker.complete(0, f"{m_count} issue(s)")

        tracker.start(1)
        b_count = validations._validate_backend_data(_errs)
        tracker.complete(1, f"{b_count} issue(s)")

        tracker.start(2)
        mc_count = validations._validate_maintenance_cost(_errs)
        tracker.complete(2, f"{mc_count} issue(s)")

        tracker.start(3)
        validations.clear_validation_errors()
        if _errs:
            database.insert_rows("validation_errors", _errs)
        ran_at = _dt.now().isoformat(timespec="seconds")
        database.set_setting("last_validation_at",     ran_at)
        database.set_setting("last_validation_errors", str(len(_errs)))
        tracker.complete(3, f"{len(_errs)} total")

        time.sleep(0.8)
        st.rerun()

    st.divider()

    # ── Summary cards ────────────────────────────────────────────────────────
    counts = database.get_table_counts()
    total_errors = counts.get("validation_errors", 0)

    if total_errors == 0:
        ui_helpers.render_glass_card(
            "All clear",
            "No validation errors are stored. "
            "Either validation has not been run yet, or all checks passed.",
        )
        return

    # Summary by source + type
    st.subheader("📊 Error Summary")
    summary_df = validations.get_validation_summary()
    if not summary_df.empty:
        col1, col2, col3, col4 = st.columns(4)
        master_n  = summary_df[summary_df["Source"] == "master_data"]["Count"].sum()
        backend_n = summary_df[summary_df["Source"] == "backend_data"]["Count"].sum()
        maint_n   = summary_df[summary_df["Source"] == "maintenance_cost"]["Count"].sum()
        col1.metric("Total errors",       int(total_errors))
        col2.metric("Master Data",        int(master_n))
        col3.metric("Backend Data",       int(backend_n))
        col4.metric("Maintenance Cost",   int(maint_n))
        st.write("")
        st.dataframe(summary_df, hide_index=True, use_container_width=True)

    st.divider()

    # ── Detailed error table with filters ───────────────────────────────────
    st.subheader("🔍 Error Details")

    all_errors_df = database.read_table("validation_errors")

    # Filters
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        source_opts = ["All"] + sorted(all_errors_df["source"].unique().tolist())
        sel_source = st.selectbox("Filter by source", options=source_opts,
                                  key="val_filter_source")
    with col_f2:
        type_opts = ["All"] + sorted(all_errors_df["error_type"].unique().tolist())
        sel_type = st.selectbox("Filter by error type", options=type_opts,
                                key="val_filter_type")

    filtered = all_errors_df.copy()
    if sel_source != "All":
        filtered = filtered[filtered["source"] == sel_source]
    if sel_type != "All":
        filtered = filtered[filtered["error_type"] == sel_type]

    st.caption(f"Showing **{len(filtered):,}** of **{total_errors:,}** errors")
    st.dataframe(
        filtered.drop(columns=["id"], errors="ignore"),
        hide_index=True,
        use_container_width=True,
    )

    # Clear button
    st.write("")
    if st.button("🗑️ Clear All Validation Errors", key="clear_val_btn"):
        validations.clear_validation_errors()
        database.set_setting("last_validation_at",     "")
        database.set_setting("last_validation_errors", "0")
        st.rerun()


def page_settings():
    ui_helpers.render_page_header(
        "Settings",
        "Control the animated background now; SMTP & cache settings come later",
    )

    # ---- Animated background controls ----
    st.subheader("🎨 Animated background")
    st.caption(
        "A Stripe-inspired ray-burst background. It reacts to your mouse and can "
        "change its theme automatically based on your computer's clock. You can "
        "also pick any theme directly from the dropdown in the top-right corner."
    )

    # Auto theme on/off. Using key=... binds the widget directly to session_state,
    # so the choice is saved automatically and applied on the next run.
    st.toggle(
        "Use automatic theme based on system time",
        key="bg_auto_theme",
        help="04-06 Pre-dawn · 06-09 Sunrise · 09-17 Daytime · "
             "17-18:30 Dusk · 18:30-20 Sunset · 20-04 Night",
    )

    # Manual theme dropdown — only shown when automatic theme is OFF.
    if not st.session_state.bg_auto_theme:
        st.selectbox("Manual theme", options=THEME_NAMES, key="bg_manual_theme")

    st.toggle("Background animation", key="bg_animate",
              help="Turn off for a calm, static gradient (no moving rays).")

    st.radio("Animation intensity", options=["Low", "Medium", "High"],
             key="bg_intensity", horizontal=True,
             help="Low = 70 rays · Medium = 120 rays · High = 180 rays")

    ui_helpers.render_warning_message(
        "Note: dark themes (Night, Pre-dawn) make the page darker. If any text "
        "feels hard to read, turn on a lighter theme or switch animation off."
    )

    st.divider()

    # ---- Email (SMTP) settings (Phase 12) ----
    st.subheader("✉️ Email (SMTP) settings")
    st.caption(
        "Used to email the Excel report from the View Reports page. For Gmail or "
        "Outlook, create an **App Password** and use that here — not your normal "
        "login password."
    )

    # Read the config ONCE and derive "configured" from it, so we don't hit the
    # database again for every is_configured() check on this page.
    cfg = email_helper.get_smtp_config()
    configured = bool(cfg["host"] and cfg["sender"] and cfg["password"])
    if configured:
        ui_helpers.render_success_message("SMTP is configured and ready to send.")
    else:
        ui_helpers.render_warning_message("SMTP is not fully configured yet.")

    with st.form("smtp_form"):
        s1, s2 = st.columns(2)
        with s1:
            smtp_host   = st.text_input("SMTP host", value=cfg["host"],
                                        placeholder="smtp.gmail.com")
            smtp_sender = st.text_input("Sender email", value=cfg["sender"],
                                        placeholder="you@company.com")
            default_to  = st.text_input("Default To (comma-separated)",
                                        value=cfg["default_to"])
        with s2:
            smtp_port = st.number_input("Port", value=int(cfg["port"]),
                                        min_value=1, max_value=65535, step=1)
            smtp_pwd  = st.text_input(
                "Password / App password", value="", type="password",
                help="Leave blank to keep the saved password. Stored locally in "
                     "your app database; use an app password where possible.",
            )
            default_cc = st.text_input("Default CC (optional)", value=cfg["default_cc"])
        smtp_subject = st.text_input("Default subject", value=cfg["subject"])
        use_tls = st.toggle("Use TLS (recommended)", value=cfg["use_tls"])
        saved = st.form_submit_button("💾 Save email settings", type="primary")

    if saved:
        # Save every key in ONE transaction (fast) instead of 8 separate writes.
        to_save = {
            "smtp_host":        smtp_host.strip(),
            "smtp_port":        str(int(smtp_port)),
            "smtp_sender":      smtp_sender.strip(),
            "smtp_use_tls":     "true" if use_tls else "false",
            "email_default_to": default_to.strip(),
            "email_default_cc": default_cc.strip(),
            "email_subject":    smtp_subject.strip(),
        }
        if smtp_pwd:  # only overwrite the password when a new one is typed
            to_save["smtp_password"] = smtp_pwd
        database.set_settings_bulk(to_save)
        ui_helpers.render_success_message("Email settings saved.")
        st.rerun()

    # Quick connection test (sends a tiny email to the sender address)
    if configured:
        if st.button("✉️ Send a test email to the sender", key="smtp_test_btn"):
            with st.spinner("Sending test email…"):
                res = email_helper.send_report_email(
                    to_emails=cfg["sender"], cc_emails="",
                    subject="Test email — Batching Incentive Calculator",
                    body="This is a test email confirming your SMTP settings work.",
                )
            if res["success"]:
                ui_helpers.render_success_message(
                    f"Test email sent to {cfg['sender']}. Check the inbox."
                )
            else:
                ui_helpers.render_error_message(f"Test failed: {res['error']}")

    st.divider()

    # ---- Cache & storage (Phase 13) ----
    st.subheader("🧹 Cache & storage")
    st.caption(
        "Housekeeping tools. These keep the app fast and reclaim disk space. "
        "They are completely safe — your Master Data, uploads, calculation "
        "results and settings are NOT deleted."
    )

    # Show the current database file size so the user can see the effect.
    st.metric("Database file size", f"{cache_helpers.get_db_size_mb()} MB")

    st.write(
        "**Clear Cache** does two things: it empties Streamlit's temporary "
        "in-memory results (so the next click reloads fresh data), and it "
        "compacts the database file to remove unused empty space."
    )

    if st.button("🧹 Clear cache & compact database", key="clear_cache_btn",
                 type="primary"):
        with st.spinner("Clearing cache and compacting the database…"):
            summary = cache_helpers.clear_cache()

        if summary["error"]:
            ui_helpers.render_error_message(
                f"Could not compact the database: {summary['error']}"
            )
        else:
            freed = summary["freed_mb"]
            freed_note = (
                f" Freed <b>{freed} MB</b> of unused space."
                if freed > 0 else
                " The database was already compact (nothing to free)."
            )
            cache_note = (
                "Streamlit memory cache cleared."
                if summary["streamlit_cache_cleared"] else
                "No Streamlit cache was active."
            )
            ui_helpers.render_success_message(
                f"Done! {cache_note}{freed_note} "
                f"Database size is now {summary['after_mb']} MB."
            )


# ---------------------------------------------------------------------------
# STEP 4: Connect each page name to its function.
# ---------------------------------------------------------------------------
# This dictionary maps the text shown in the sidebar to the function that
# draws that page. It keeps the routing tidy and easy to extend.

PAGE_FUNCTIONS = {
    "Dashboard": page_dashboard,
    "Data Uploader": page_data_uploader,
    "Calculate Incentive & Deduction": page_calculate,
    "View Reports": page_view_reports,
    "Error / Validation Report": page_validation,
    "Settings": page_settings,
}


# ---------------------------------------------------------------------------
# STEP 5: Run the app.
# ---------------------------------------------------------------------------
def _init_background_settings():
    """
    Make sure the animated-background settings exist in st.session_state.
    session_state is Streamlit's memory that survives between reruns, so the
    user's choices on the Settings page are remembered while the app is open.
    """
    ss = st.session_state
    ss.setdefault("bg_auto_theme", True)        # follow system time?
    ss.setdefault("bg_manual_theme", "Daytime")  # used when auto is off
    ss.setdefault("bg_animate", True)            # animate or static?
    ss.setdefault("bg_intensity", "Medium")      # Low / Medium / High


def _auto_sync_master_data():
    """
    If the user has turned on "Auto-sync on startup", fetch the latest master
    data from Google Sheets once per browser session (not on every rerun).

    We use the session_state flag 'startup_sync_done' to make sure we only
    run this once, even though Streamlit re-executes the whole script on
    every widget interaction.
    """
    ss = st.session_state

    # Already ran in this browser session — skip.
    if ss.get("startup_sync_done"):
        return

    ss["startup_sync_done"] = True  # mark as done immediately to avoid double-run

    # Check whether the user has opted in to auto-sync.
    auto_sync_enabled = database.get_setting("gsheet_auto_sync", "false") == "true"
    if not auto_sync_enabled:
        return

    # Need a sheet ID and credentials in place to actually sync.
    sheet_id  = database.get_setting("gsheet_id", "")
    worksheet = database.get_setting("gsheet_worksheet", "Sheet1")
    if not sheet_id or not google_sheets.credentials_exist():
        return

    # Run the sync silently in the background — store the result so the
    # Data Uploader page can show a banner if it wants to.
    result = google_sheets.sync_master_data(sheet_id, worksheet)
    ss["startup_sync_result"] = result


def main():
    # Create the database and its tables once per session. init_db() is safe to
    # call repeatedly, but the flag avoids doing it on every single rerun.
    if not st.session_state.get("db_ready"):
        database.init_db()
        st.session_state["db_ready"] = True

    # Auto-sync master data from Google Sheets if the setting is turned on.
    _auto_sync_master_data()

    _init_background_settings()
    ss = st.session_state

    # Draw the app-wide animated background ONCE, using the saved settings.
    # (Calling it every run is safe — it updates the existing background.)
    render_interactive_background(
        mode="auto" if ss.bg_auto_theme else "manual",
        theme=ss.bg_manual_theme,
        intensity=ss.bg_intensity.lower(),
        animate=ss.bg_animate,
    )

    selected_page = render_sidebar()

    # Find the function for the selected page and call it.
    page_function = PAGE_FUNCTIONS.get(selected_page, page_dashboard)
    page_function()


# This standard Python line means: "only run main() when this file is the one
# being executed" (which is exactly what `streamlit run app.py` does).
if __name__ == "__main__":
    main()
