"""
database.py
===========
Everything about the local SQLite database lives here.

SQLite is a tiny database that is stored in ONE file on your computer
(data/app.db). There is nothing extra to install or run — Python has SQLite
built in. That makes it perfect for a simple, local, single-admin app.

This file is the ONLY place that talks to the database. Other parts of the app
(calculations, uploads, email) call the helper functions below instead of
writing SQL themselves. That keeps the database logic in one tidy place.

What this file provides (Phase 2):
- get_connection()                 -> open a connection to data/app.db
- init_db()                        -> create all 9 tables if they don't exist
- Generic helpers:
    insert_rows(table, rows)
    replace_table_rows(table, rows)
    read_table(table)              -> returns a pandas DataFrame
    clear_table(table)
    get_table_counts()             -> {table_name: row_count}
- Settings helpers:
    set_setting(key, value), get_setting(key, default), get_all_settings()
- Log helpers:
    log_master_data_change(...), log_email(...)

The actual UPLOADING and CALCULATING that fill these tables come in later
phases (3, 4, 6, 10, 11, 12). Here we just build the storage and the tools to
read/write it.
"""

import json
import os
import sqlite3
from datetime import datetime

import pandas as pd


# ---------------------------------------------------------------------------
# 1. WHERE THE DATABASE FILE LIVES
# ---------------------------------------------------------------------------
# The database file is data/app.db, relative to this project folder.
DB_DIR = "data"
DB_PATH = os.path.join(DB_DIR, "app.db")


def _now():
    """Return the current time as a short, sortable text string."""
    return datetime.now().isoformat(timespec="seconds")


# ---------------------------------------------------------------------------
# 2. THE TABLE DEFINITIONS (the database "schema")
# ---------------------------------------------------------------------------
# Each entry is a CREATE TABLE statement. "IF NOT EXISTS" means it is safe to
# run every time the app starts — existing tables are left untouched.
#
# A quick note on column types in SQLite:
#   TEXT    = words / dates stored as text (we keep codes as TEXT so leading
#             zeros like "007" are never lost)
#   INTEGER = whole numbers
#   REAL    = numbers with decimals (quantities, amounts, rates)

TABLE_SCHEMAS = {
    # Employees synced from Google Sheets (Phase 3) and editable in-app (Phase 11)
    "master_data": """
        CREATE TABLE IF NOT EXISTS master_data (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_code TEXT UNIQUE NOT NULL,
            employee_name TEXT,
            designation   TEXT,
            category      TEXT,
            plant         TEXT,
            plant_code    TEXT,
            updated_at    TEXT
        )
    """,

    # Rows read from the uploaded Backend Data Excel (Phase 4)
    "backend_data": """
        CREATE TABLE IF NOT EXISTS backend_data (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            created_by  TEXT,
            quantity    REAL,
            date        TEXT,
            source_file TEXT,
            uploaded_at TEXT
        )
    """,

    # Rows read from the uploaded Maintenance Cost Excel (Phase 4)
    # month + year added so each calendar month can hold its own cost file.
    # UNIQUE on (plant_code, month, year) — one cost per plant per month.
    "maintenance_cost": """
        CREATE TABLE IF NOT EXISTS maintenance_cost (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_code           TEXT,
            month                INTEGER,
            year                 INTEGER,
            ytd_maintenance_cost REAL,
            uploaded_at          TEXT,
            UNIQUE(plant_code, month, year)
        )
    """,

    # The final calculated incentive/deduction results (Phase 6 / 10)
    "calculation_results": """
        CREATE TABLE IF NOT EXISTS calculation_results (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            month                INTEGER,
            year                 INTEGER,
            employee_code        TEXT,
            employee_name        TEXT,
            designation          TEXT,
            category             TEXT,
            plant                TEXT,
            plant_code           TEXT,
            total_quantity       REAL,
            ytd_maintenance_cost REAL,
            incentive_eligible   TEXT,
            incentive_rate       REAL,
            incentive_amount     REAL,
            deduction_target     REAL,
            shortfall_quantity   REAL,
            deduction_amount     REAL,
            remarks              TEXT,
            generated_at         TEXT
        )
    """,

    # Employee codes found in Backend Data but missing from Master Data (Phase 6)
    "unmapped_employees": """
        CREATE TABLE IF NOT EXISTS unmapped_employees (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_code  TEXT,
            month          INTEGER,
            year           INTEGER,
            total_quantity REAL,
            remarks        TEXT,
            generated_at   TEXT
        )
    """,

    # Problems found while validating data (Phase 5)
    "validation_errors": """
        CREATE TABLE IF NOT EXISTS validation_errors (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            source        TEXT,
            row_number    INTEGER,
            column_name   TEXT,
            error_type    TEXT,
            error_message TEXT,
            created_at    TEXT
        )
    """,

    # A record of every report email that was sent (Phase 12)
    "email_log": """
        CREATE TABLE IF NOT EXISTS email_log (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            report_file_name TEXT,
            to_emails        TEXT,
            cc_emails        TEXT,
            subject          TEXT,
            sent_at          TEXT,
            status           TEXT,
            error_message    TEXT
        )
    """,

    # An audit trail of master-data add/edit/delete actions (Phase 11)
    "master_data_change_log": """
        CREATE TABLE IF NOT EXISTS master_data_change_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            action_type     TEXT,
            employee_code   TEXT,
            old_value_json  TEXT,
            new_value_json  TEXT,
            changed_at      TEXT,
            changed_by      TEXT
        )
    """,

    # Simple key/value store for app settings (theme, SMTP, etc.)
    "app_settings": """
        CREATE TABLE IF NOT EXISTS app_settings (
            key        TEXT PRIMARY KEY,
            value      TEXT,
            updated_at TEXT
        )
    """,

    # Deduction waivers — employee approved for zero deduction for a given month
    "waivers": """
        CREATE TABLE IF NOT EXISTS waivers (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_code TEXT NOT NULL,
            month         INTEGER NOT NULL,
            year          INTEGER NOT NULL,
            reason        TEXT NOT NULL,
            custom_reason TEXT,
            created_at    TEXT,
            UNIQUE(employee_code, month, year)
        )
    """,

    # ── RDC-TP tables ────────────────────────────────────────────────────────

    # Plant reference data synced from "Plant Data for TP" Google Sheet tab
    "tp_plant_data": """
        CREATE TABLE IF NOT EXISTS tp_plant_data (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_code     TEXT UNIQUE NOT NULL,
            exco_location  TEXT,
            plant_name     TEXT,
            business_head  TEXT,
            plant_manager  TEXT,
            mixer_theo_cap REAL,
            updated_at     TEXT
        )
    """,

    # Audit trail for manual add/edit/delete of TP plant rows
    "tp_plant_change_log": """
        CREATE TABLE IF NOT EXISTS tp_plant_change_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            action_type     TEXT,
            plant_code      TEXT,
            old_value_json  TEXT,
            new_value_json  TEXT,
            changed_at      TEXT,
            changed_by      TEXT
        )
    """,

    # Raw production rows pulled from Oracle (after cleaning & filtering)
    "tp_oracle_data": """
        CREATE TABLE IF NOT EXISTS tp_oracle_data (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            production_date TEXT,
            plant_code      TEXT,
            mixer_variant   TEXT,
            lookup_code     TEXT,
            batch_ref       TEXT,
            quantity        REAL,
            time_taken_min  REAL,
            fetched_at      TEXT
        )
    """,

    # Final calculated throughput results per plant / mixer
    "tp_results": """
        CREATE TABLE IF NOT EXISTS tp_results (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            month          INTEGER,
            year           INTEGER,
            lookup_code    TEXT,
            plant_name     TEXT,
            exco_location  TEXT,
            business_head  TEXT,
            plant_manager  TEXT,
            mixer_theo_cap REAL,
            total_quantity REAL,
            total_time_hrs REAL,
            throughput_pct REAL,
            batch_count    INTEGER,
            generated_at   TEXT
        )
    """,

    # ── RDC-ECMD tables ─────────────────────────────────────────────────────

    # Plant energy & DG readings entered by plant persons each month
    # One row per plant per month. UNIQUE on (plant_code, month, year).
    "ecmd_readings": """
        CREATE TABLE IF NOT EXISTS ecmd_readings (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_code            TEXT NOT NULL,
            month                 INTEGER NOT NULL,
            year                  INTEGER NOT NULL,
            eb_kwh_open           REAL,
            eb_kwh_close          REAL,
            eb_kvah_open          REAL,
            eb_kvah_close         REAL,
            mf                    REAL DEFAULT 1.0,
            dg_hr_open            REAL,
            dg_hr_close           REAL,
            dg_kwh_open           REAL,
            dg_kwh_close          REAL,
            mixer_dg_hr_open      REAL,
            mixer_dg_hr_close     REAL,
            diesel_issued_ltrs    REAL,
            volume_on_dg          REAL,
            entered_by            TEXT DEFAULT 'Admin',
            entered_at            TEXT,
            UNIQUE(plant_code, month, year)
        )
    """,

    # Daily readings — one row per plant per day (alternative to monthly ecmd_readings)
    "ecmd_daily_readings": """
        CREATE TABLE IF NOT EXISTS ecmd_daily_readings (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_code            TEXT NOT NULL,
            month                 INTEGER NOT NULL,
            year                  INTEGER NOT NULL,
            day                   INTEGER NOT NULL,
            eb_kwh_open           REAL,
            eb_kwh_close          REAL,
            eb_kvah_open          REAL,
            eb_kvah_close         REAL,
            mf                    REAL DEFAULT 1.0,
            dg_hr_open            REAL,
            dg_hr_close           REAL,
            dg_kwh_open           REAL,
            dg_kwh_close          REAL,
            mixer_dg_hr_open      REAL,
            mixer_dg_hr_close     REAL,
            entered_by            TEXT DEFAULT 'Admin',
            entered_at            TEXT,
            UNIQUE(plant_code, month, year, day)
        )
    """,

    # Entry mode per plant+month+year: 'monthly' or 'daily'
    "ecmd_entry_mode": """
        CREATE TABLE IF NOT EXISTS ecmd_entry_mode (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            plant_code TEXT NOT NULL,
            month      INTEGER NOT NULL,
            year       INTEGER NOT NULL,
            mode       TEXT NOT NULL DEFAULT 'monthly',
            UNIQUE(plant_code, month, year)
        )
    """,

    # Dual Plant Utilisation — cached fortnightly report rows
    "ecmd_dual_plant": """
        CREATE TABLE IF NOT EXISTS ecmd_dual_plant (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            period_label TEXT NOT NULL,
            from_date    TEXT NOT NULL,
            to_date      TEXT NOT NULL,
            plant_code   TEXT NOT NULL,
            plant_name   TEXT,
            mixer        TEXT NOT NULL,
            quantity     REAL DEFAULT 0,
            pct_share    REAL DEFAULT 0,
            fetched_at   TEXT
        )
    """,

    # Invoice Final Submission Pending — cached fortnightly report rows
    "ecmd_invoice_pending": """
        CREATE TABLE IF NOT EXISTS ecmd_invoice_pending (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            period_label TEXT NOT NULL,
            from_date    TEXT NOT NULL,
            to_date      TEXT NOT NULL,
            plant_code   TEXT NOT NULL,
            plant_name   TEXT,
            sales_order  TEXT,
            line_number  TEXT,
            quantity     REAL DEFAULT 0,
            fetched_at   TEXT
        )
    """,

    # Calculated ECMD results — one row per plant per month
    "ecmd_results": """
        CREATE TABLE IF NOT EXISTS ecmd_results (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            month                 INTEGER,
            year                  INTEGER,
            plant_code            TEXT,
            plant_name            TEXT,
            exco_location         TEXT,
            business_head         TEXT,
            plant_manager         TEXT,
            eb_kwh                REAL,
            dg_kwh                REAL,
            total_kwh             REAL,
            total_volume          REAL,
            energy_per_mt         REAL,
            dg_run_hrs            REAL,
            mixer_dg_hrs          REAL,
            mixer_dg_ratio        REAL,
            diesel_issued_ltrs    REAL,
            ltr_per_hr            REAL,
            volume_on_dg          REAL,
            generated_at          TEXT
        )
    """,

    # ── Shared Oracle raw cache ──────────────────────────────────────────────
    # One table for ALL modules — fetched once, read by I&D / TP / BTRTP.
    # Columns are a superset of every module's needs.
    # Index on production_date keeps per-month queries fast.
    "oracle_raw_data": """
        CREATE TABLE IF NOT EXISTS oracle_raw_data (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            production_date TEXT NOT NULL,
            created_by      TEXT,
            plant_code      TEXT,
            batch_ref       TEXT,
            quantity        REAL,
            time_taken_min  REAL,
            fetched_at      TEXT
        )
    """,

    # ── RDC-BTRTP tables ────────────────────────────────────────────────────

    # Raw Oracle rows with batcher (CREATED_BY) column
    "btrtp_oracle_data": """
        CREATE TABLE IF NOT EXISTS btrtp_oracle_data (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            production_date TEXT,
            batcher_id      TEXT,
            plant_code      TEXT,
            mixer_variant   TEXT,
            lookup_code     TEXT,
            batch_ref       TEXT,
            quantity        REAL,
            time_taken_min  REAL,
            fetched_at      TEXT
        )
    """,

    # BT Master synced from Google Sheets (batcher_id → batcher_name)
    "btrtp_master_data": """
        CREATE TABLE IF NOT EXISTS btrtp_master_data (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            batcher_id   TEXT UNIQUE,
            batcher_name TEXT,
            updated_at   TEXT
        )
    """,

    # Calculated batcher-wise throughput results
    "btrtp_results": """
        CREATE TABLE IF NOT EXISTS btrtp_results (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            month          INTEGER,
            year           INTEGER,
            batcher_id     TEXT,
            batcher_name   TEXT,
            lookup_code    TEXT,
            plant_name     TEXT,
            exco_location  TEXT,
            business_head  TEXT,
            plant_manager  TEXT,
            mixer_theo_cap REAL,
            total_quantity REAL,
            total_time_hrs REAL,
            throughput_pct REAL,
            batch_count    INTEGER,
            generated_at   TEXT
        )
    """,

    # ── Authentication & Authorisation tables ────────────────────────────────

    "users": """
        CREATE TABLE IF NOT EXISTS users (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name            TEXT NOT NULL,
            email                TEXT NOT NULL,
            username             TEXT UNIQUE NOT NULL,
            password_hash        TEXT NOT NULL,
            role                 TEXT NOT NULL DEFAULT 'PLANT_USER',
            is_active            INTEGER NOT NULL DEFAULT 1,
            must_change_password INTEGER NOT NULL DEFAULT 0,
            created_at           TEXT,
            updated_at           TEXT,
            last_login_at        TEXT
        )
    """,

    "user_plant_access": """
        CREATE TABLE IF NOT EXISTS user_plant_access (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id      INTEGER NOT NULL,
            plant_code   TEXT NOT NULL,
            plant_name   TEXT NOT NULL,
            created_at   TEXT,
            UNIQUE(user_id, plant_name)
        )
    """,

    "login_audit_log": """
        CREATE TABLE IF NOT EXISTS login_audit_log (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id        INTEGER,
            login_time     TEXT,
            logout_time    TEXT,
            ip_address     TEXT,
            user_agent     TEXT,
            status         TEXT,
            failure_reason TEXT
        )
    """,

    "user_activity_log": """
        CREATE TABLE IF NOT EXISTS user_activity_log (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id      INTEGER,
            action       TEXT,
            module_name  TEXT,
            details_json TEXT,
            created_at   TEXT,
            ip_address   TEXT
        )
    """,

    # ── Slow Loading Alert (SLA) tables ─────────────────────────────────────

    # Mixer capacity → base allowed time lookup table
    "slow_loading_thresholds": """
        CREATE TABLE IF NOT EXISTS slow_loading_thresholds (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            mixer_capacity      REAL NOT NULL,
            reference_quantity  REAL NOT NULL DEFAULT 6,
            base_allowed_minutes REAL NOT NULL,
            is_active           INTEGER NOT NULL DEFAULT 1,
            created_at          TEXT,
            updated_at          TEXT
        )
    """,

    # One row per detected slow-loading event
    "slow_loading_alert_records": """
        CREATE TABLE IF NOT EXISTS slow_loading_alert_records (
            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
            alert_date              TEXT NOT NULL,
            alert_hour              INTEGER NOT NULL,
            plant_code              TEXT,
            plant_name              TEXT,
            customer                TEXT,
            grade                   TEXT,
            batcher_code            TEXT,
            batcher_name            TEXT,
            tm_number               TEXT,
            batched_quantity        REAL,
            mixer_capacity          REAL,
            loading_time_minutes    REAL,
            allowed_loading_minutes REAL,
            delay_minutes           REAL,
            alert_type              TEXT DEFAULT 'HOURLY',
            status                  TEXT DEFAULT 'OPEN',
            remarks                 TEXT,
            alert_key               TEXT,
            created_at              TEXT,
            updated_at              TEXT
        )
    """,

    # Email send log for SLA module
    "slow_loading_email_logs": """
        CREATE TABLE IF NOT EXISTS slow_loading_email_logs (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            alert_type    TEXT,
            alert_date    TEXT,
            alert_hour    INTEGER,
            plant_code    TEXT,
            plant_name    TEXT,
            to_emails     TEXT,
            cc_emails     TEXT,
            subject       TEXT,
            total_cases   INTEGER,
            status        TEXT,
            error_message TEXT,
            sent_at       TEXT,
            created_at    TEXT
        )
    """,

    # Scheduler run log for SLA module
    "slow_loading_scheduler_logs": """
        CREATE TABLE IF NOT EXISTS slow_loading_scheduler_logs (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            job_name              TEXT,
            run_started_at        TEXT,
            run_completed_at      TEXT,
            status                TEXT,
            total_records_checked INTEGER DEFAULT 0,
            total_alert_cases     INTEGER DEFAULT 0,
            total_emails_sent     INTEGER DEFAULT 0,
            error_message         TEXT
        )
    """,

    # Key-value config store for SLA module (mirrors module_settings pattern)
    "slow_loading_config": """
        CREATE TABLE IF NOT EXISTS slow_loading_config (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            config_key   TEXT UNIQUE NOT NULL,
            config_value TEXT,
            updated_at   TEXT
        )
    """,
}


# ---------------------------------------------------------------------------
# 3. CONNECTION + TABLE CREATION
# ---------------------------------------------------------------------------
def get_connection():
    """
    Open and return a connection to the SQLite database.

    - We make sure the data/ folder exists first.
    - check_same_thread=False lets Streamlit use the connection safely.
    - row_factory = sqlite3.Row lets us read columns by name (like a dict).

    Remember to close the connection when done (the helpers below do this).
    """
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """
    Create every table if it does not already exist, and run any one-time
    column migrations needed on older databases.
    """
    conn = get_connection()
    try:
        cur = conn.cursor()
        for create_sql in TABLE_SCHEMAS.values():
            cur.execute(create_sql)

        # Migration: add month + year columns to maintenance_cost if they don't exist
        cur.execute("PRAGMA table_info(maintenance_cost)")
        existing_cols = {row[1] for row in cur.fetchall()}
        if "month" not in existing_cols:
            cur.execute("ALTER TABLE maintenance_cost ADD COLUMN month INTEGER")
        if "year" not in existing_cols:
            cur.execute("ALTER TABLE maintenance_cost ADD COLUMN year  INTEGER")

        # Index for fast date-range queries on the shared Oracle raw cache
        cur.execute("""CREATE INDEX IF NOT EXISTS idx_oracle_raw_date
                       ON oracle_raw_data(production_date)""")

        # ── SLA indexes ──────────────────────────────────────────────────────
        cur.execute("""CREATE INDEX IF NOT EXISTS idx_sla_records_date
                       ON slow_loading_alert_records(alert_date)""")
        cur.execute("""CREATE INDEX IF NOT EXISTS idx_sla_records_hour
                       ON slow_loading_alert_records(alert_hour)""")
        cur.execute("""CREATE INDEX IF NOT EXISTS idx_sla_records_plant
                       ON slow_loading_alert_records(plant_code)""")
        cur.execute("""CREATE INDEX IF NOT EXISTS idx_sla_records_key
                       ON slow_loading_alert_records(alert_key)""")
        cur.execute("""CREATE INDEX IF NOT EXISTS idx_sla_records_status
                       ON slow_loading_alert_records(status)""")

        # Seed threshold table if empty
        cur.execute("SELECT COUNT(*) FROM slow_loading_thresholds")
        if cur.fetchone()[0] == 0:
            now_ts = datetime.now().isoformat(timespec="seconds")
            thresholds = [
                (30, 6, 16), (45, 6, 12), (56, 6, 10), (60, 6, 10),
                (65, 6, 10), (70, 6, 10), (75, 6, 10), (80, 6, 10),
                (86, 6, 10), (90, 6,  8), (100, 6, 8), (110, 6, 8),
                (120, 6,  8),
            ]
            cur.executemany(
                "INSERT INTO slow_loading_thresholds "
                "(mixer_capacity, reference_quantity, base_allowed_minutes, is_active, created_at, updated_at) "
                "VALUES (?, ?, ?, 1, ?, ?)",
                [(m, r, b, now_ts, now_ts) for m, r, b in thresholds]
            )

        conn.commit()
    finally:
        conn.close()


def get_maintenance_months() -> list:
    """Return list of (month, year) tuples that exist in maintenance_cost, newest first."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT DISTINCT month, year FROM maintenance_cost "
            "WHERE month IS NOT NULL AND year IS NOT NULL "
            "ORDER BY year DESC, month DESC"
        )
        return [(r[0], r[1]) for r in cur.fetchall()]
    finally:
        conn.close()


def assign_maintenance_month(month: int, year: int) -> int:
    """Set month+year on all rows where month IS NULL (legacy/unassigned rows)."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "UPDATE maintenance_cost SET month = ?, year = ? WHERE month IS NULL OR year IS NULL",
            (month, year)
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def delete_maintenance_month(month: int, year: int) -> int:
    """Delete all maintenance_cost rows for the given month+year. Returns rows deleted."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM maintenance_cost WHERE month = ? AND year = ?",
            (month, year)
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# 4. GENERIC READ / WRITE HELPERS
# ---------------------------------------------------------------------------
# These work for ANY table, which keeps us from writing nearly identical
# functions for every table. We only ever pass in our own fixed table names
# (never anything typed by a user), so building the SQL text is safe here.

def insert_rows(table, rows):
    """
    Insert a list of rows into a table.

    `rows` is a list of dictionaries, where each dictionary's keys are column
    names, e.g. [{"plant_code": "P1", "ytd_maintenance_cost": 12.5}, ...].
    Returns the number of rows inserted.
    """
    if not rows:
        return 0

    columns = list(rows[0].keys())
    col_text = ", ".join(columns)
    placeholders = ", ".join(["?"] * len(columns))
    sql = f"INSERT INTO {table} ({col_text}) VALUES ({placeholders})"
    values = [tuple(row.get(col) for col in columns) for row in rows]

    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.executemany(sql, values)
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def clear_table(table):
    """Delete ALL rows from a table (the table itself stays)."""
    conn = get_connection()
    try:
        conn.execute(f"DELETE FROM {table}")
        conn.commit()
    finally:
        conn.close()


def replace_table_rows(table, rows):
    """
    Replace a table's contents: clear it, then insert the new rows.
    Useful for data that is re-uploaded or re-synced (master data, backend
    data, maintenance cost). Returns the number of rows inserted.
    """
    clear_table(table)
    return insert_rows(table, rows)


# Oracle-sourced tables and the column that holds their production date.
# Shared by ALL modules — add new modules' Oracle tables here so the rolling
# retention policy covers them automatically.
ORACLE_DATA_TABLES = {
    "oracle_raw_data":   "production_date",  # shared cache — all modules
    "backend_data":      "date",             # RDC-I&D (Excel uploads kept separately)
    "tp_oracle_data":    "production_date",  # RDC-TP legacy (kept for compat)
    "btrtp_oracle_data": "production_date",  # RDC-BTRTP legacy (kept for compat)
}


def _prev_month_first_day(today=None):
    """First day of the PREVIOUS calendar month (the retention cut-off)."""
    from datetime import date
    t = today or date.today()
    if t.month == 1:
        return date(t.year - 1, 12, 1)
    return date(t.year, t.month - 1, 1)


def purge_old_oracle_data(today=None):
    """
    Rolling retention for ALL modules' Oracle data (shared policy).

    Keeps only rows dated on/after the 1st of the PREVIOUS month and deletes
    everything older. So in June you keep May + June; when July starts, May is
    purged and you keep June + July. This caps storage and lowers the load.

    Dates are stored as 'YYYY-MM-DD' strings, so a plain string comparison is
    correct (ISO dates sort lexicographically). Returns a summary dict.
    """
    cutoff = _prev_month_first_day(today).isoformat()
    deleted = {}
    conn = get_connection()
    try:
        existing = {
            r[0] for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        for table, date_col in ORACLE_DATA_TABLES.items():
            if table not in existing:
                continue
            cur = conn.execute(
                f"DELETE FROM {table} WHERE {date_col} IS NOT NULL "
                f"AND {date_col} < ?", (cutoff,)
            )
            deleted[table] = cur.rowcount
        conn.commit()
    finally:
        conn.close()
    return {"cutoff": cutoff, "deleted": deleted}


def read_table(table, order_by=None):
    """
    Read a whole table into a pandas DataFrame (a table in memory).
    Optionally order the rows by a column name.
    NOTE: avoid calling this on large tables (backend_data can have 60K+ rows).
    Use read_table_limited() for display purposes.
    """
    sql = f"SELECT * FROM {table}"
    if order_by:
        sql += f" ORDER BY {order_by}"
    conn = get_connection()
    try:
        return pd.read_sql_query(sql, conn)
    finally:
        conn.close()


def read_table_limited(table, order_by=None, limit=200):
    """
    Read only the first `limit` rows from a table — safe for large tables.
    Used for display/preview so the browser never has to render 60K+ rows.
    """
    sql = f"SELECT * FROM {table}"
    if order_by:
        sql += f" ORDER BY {order_by}"
    sql += f" LIMIT {limit}"
    conn = get_connection()
    try:
        return pd.read_sql_query(sql, conn)
    finally:
        conn.close()


def get_table_counts():
    """
    Return a dictionary of {table_name: number_of_rows} for every table.
    Handy for a "database status" panel so you can confirm the tables exist.
    """
    counts = {}
    conn = get_connection()
    try:
        cur = conn.cursor()
        for table in TABLE_SCHEMAS:
            cur.execute(f"SELECT COUNT(*) FROM {table}")
            counts[table] = cur.fetchone()[0]
    finally:
        conn.close()
    return counts


# ---------------------------------------------------------------------------
# 5. SETTINGS HELPERS (app_settings table)
# ---------------------------------------------------------------------------
def set_setting(key, value):
    """Save (or update) one setting. Values are stored as text."""
    conn = get_connection()
    try:
        conn.execute(
            "INSERT INTO app_settings (key, value, updated_at) VALUES (?, ?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value, "
            "updated_at = excluded.updated_at",
            (key, str(value), _now()),
        )
        conn.commit()
    finally:
        conn.close()


def set_settings_bulk(mapping):
    """
    Save many settings at once in a SINGLE connection/transaction.

    `mapping` is a {key: value} dictionary. This is much faster than calling
    set_setting() in a loop (which opens a new connection for every key) and is
    what the Settings page uses when you press "Save".
    """
    if not mapping:
        return
    now = _now()
    rows = [(k, str(v), now) for k, v in mapping.items()]
    conn = get_connection()
    try:
        conn.executemany(
            "INSERT INTO app_settings (key, value, updated_at) VALUES (?, ?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value, "
            "updated_at = excluded.updated_at",
            rows,
        )
        conn.commit()
    finally:
        conn.close()


def get_setting(key, default=None):
    """Read one setting. Returns `default` if the key was never saved."""
    conn = get_connection()
    try:
        cur = conn.execute("SELECT value FROM app_settings WHERE key = ?", (key,))
        row = cur.fetchone()
        return row[0] if row else default
    finally:
        conn.close()


def get_all_settings():
    """Return all settings as a {key: value} dictionary."""
    conn = get_connection()
    try:
        cur = conn.execute("SELECT key, value FROM app_settings")
        return {row[0]: row[1] for row in cur.fetchall()}
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# 5b. MODULE-SCOPED SETTINGS  (prefix = "{module}.")
# ---------------------------------------------------------------------------
# Every module must use these instead of get_setting/set_setting so that keys
# from different modules never collide in the shared app_settings table.
# Convention: module IDs are lowercase short codes — "id", "tp", "btrtp", "jldc".
#
# Example: database.get_module_setting("tp", "target_qty", 0)
#          → reads key  "tp.target_qty"

def get_module_setting(module: str, key: str, default=None):
    """Read one module-scoped setting.  Key stored as '{module}.{key}'."""
    return get_setting(f"{module}.{key}", default)


def set_module_setting(module: str, key: str, value):
    """Save one module-scoped setting.  Key stored as '{module}.{key}'."""
    set_setting(f"{module}.{key}", str(value))


def set_module_settings_bulk(module: str, mapping: dict):
    """Save many module-scoped settings at once (one DB round-trip)."""
    set_settings_bulk({f"{module}.{k}": v for k, v in mapping.items()})


# ---------------------------------------------------------------------------
# 6. LOG HELPERS
# ---------------------------------------------------------------------------
def log_master_data_change(action_type, employee_code,
                           old_value=None, new_value=None, changed_by="Admin"):
    """
    Record one master-data change (add / edit / delete) in the audit log.
    `old_value` and `new_value` are dictionaries; we store them as JSON text.
    Since there is no login, changed_by defaults to "Admin".
    """
    insert_rows("master_data_change_log", [{
        "action_type": action_type,
        "employee_code": employee_code,
        "old_value_json": json.dumps(old_value) if old_value is not None else None,
        "new_value_json": json.dumps(new_value) if new_value is not None else None,
        "changed_at": _now(),
        "changed_by": changed_by,
    }])


def log_email(report_file_name, to_emails, cc_emails, subject,
              status, error_message=None):
    """Record one report-email send attempt (success or failure)."""
    insert_rows("email_log", [{
        "report_file_name": report_file_name,
        "to_emails": to_emails,
        "cc_emails": cc_emails,
        "subject": subject,
        "sent_at": _now(),
        "status": status,
        "error_message": error_message,
    }])


# ---------------------------------------------------------------------------
# 7. MASTER DATA SINGLE-ROW HELPERS (Phase 11: Add / Edit / Delete)
# ---------------------------------------------------------------------------
# These let the admin manage ONE employee at a time (the sync in google_sheets
# handles bulk replace). Every change is written to master_data_change_log so
# there is a full audit trail. Employee codes are kept as TEXT.

# The fields the admin can edit (id and updated_at are managed automatically).
_EMPLOYEE_FIELDS = ["employee_name", "designation", "category", "plant", "plant_code"]


def get_employee(employee_code):
    """Return one employee as a dict, or None if the code is not found."""
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM master_data WHERE employee_code = ?",
            (str(employee_code).strip(),),
        )
        row = cur.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def _employee_snapshot(emp):
    """Pick just the editable fields from a full employee dict (for logging)."""
    if emp is None:
        return None
    return {f: emp.get(f) for f in _EMPLOYEE_FIELDS}


def add_employee(employee_code, employee_name, designation,
                 category, plant, plant_code, changed_by="Admin"):
    """
    Add ONE new employee. Raises ValueError if the code is blank or already
    exists. Logs an 'ADD' entry in the change log.
    """
    code = str(employee_code).strip()
    if not code:
        raise ValueError("Employee Code cannot be blank.")
    if get_employee(code) is not None:
        raise ValueError(f"Employee Code '{code}' already exists.")

    new_values = {
        "employee_name": employee_name.strip(),
        "designation":   designation.strip(),
        "category":      category.strip(),
        "plant":         plant.strip(),
        "plant_code":    str(plant_code).strip(),
    }
    row = {"employee_code": code, **new_values, "updated_at": _now()}
    insert_rows("master_data", [row])
    log_master_data_change("ADD", code, old_value=None,
                           new_value=new_values, changed_by=changed_by)


def update_employee(employee_code, employee_name, designation,
                    category, plant, plant_code, changed_by="Admin"):
    """
    Update ONE existing employee's editable fields. Raises ValueError if the
    code does not exist. Logs an 'EDIT' entry (old + new values).
    """
    code = str(employee_code).strip()
    old = get_employee(code)
    if old is None:
        raise ValueError(f"Employee Code '{code}' was not found.")

    new_values = {
        "employee_name": employee_name.strip(),
        "designation":   designation.strip(),
        "category":      category.strip(),
        "plant":         plant.strip(),
        "plant_code":    str(plant_code).strip(),
    }
    conn = get_connection()
    try:
        conn.execute(
            "UPDATE master_data SET employee_name=?, designation=?, category=?, "
            "plant=?, plant_code=?, updated_at=? WHERE employee_code=?",
            (new_values["employee_name"], new_values["designation"],
             new_values["category"], new_values["plant"],
             new_values["plant_code"], _now(), code),
        )
        conn.commit()
    finally:
        conn.close()

    log_master_data_change("EDIT", code, old_value=_employee_snapshot(old),
                           new_value=new_values, changed_by=changed_by)


def delete_employee(employee_code, changed_by="Admin"):
    """
    Delete ONE employee. Raises ValueError if the code does not exist.
    Logs a 'DELETE' entry (with the old values so it can be checked later).
    """
    code = str(employee_code).strip()
    old = get_employee(code)
    if old is None:
        raise ValueError(f"Employee Code '{code}' was not found.")

    conn = get_connection()
    try:
        conn.execute("DELETE FROM master_data WHERE employee_code = ?", (code,))
        conn.commit()
    finally:
        conn.close()

    log_master_data_change("DELETE", code, old_value=_employee_snapshot(old),
                           new_value=None, changed_by=changed_by)


# ---------------------------------------------------------------------------
# 8. TP PLANT DATA SINGLE-ROW HELPERS
# ---------------------------------------------------------------------------
_TP_PLANT_FIELDS = ["exco_location", "plant_name", "business_head",
                    "plant_manager", "mixer_theo_cap"]


def get_tp_plants():
    """Return all rows from tp_plant_data as list of dicts, ordered by plant_code."""
    df = read_table("tp_plant_data", order_by="plant_code")
    if df.empty:
        return []
    df = df.drop(columns=["id"], errors="ignore")
    return [dict(r) for _, r in df.iterrows()]


def get_tp_plant(plant_code):
    """Return one plant as a dict, or None if not found."""
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM tp_plant_data WHERE plant_code = ?",
            (str(plant_code).strip(),),
        )
        row = cur.fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_tp_plant_codes():
    """Return sorted list of all plant codes."""
    conn = get_connection()
    try:
        cur = conn.execute("SELECT plant_code FROM tp_plant_data ORDER BY plant_code")
        return [r[0] for r in cur.fetchall()]
    finally:
        conn.close()


def _tp_plant_snapshot(plant):
    if plant is None:
        return None
    return {f: plant.get(f) for f in _TP_PLANT_FIELDS}


def log_tp_plant_change(action_type, plant_code, old_value=None,
                        new_value=None, changed_by="Admin"):
    insert_rows("tp_plant_change_log", [{
        "action_type":    action_type,
        "plant_code":     plant_code,
        "old_value_json": json.dumps(old_value) if old_value is not None else None,
        "new_value_json": json.dumps(new_value) if new_value is not None else None,
        "changed_at":     _now(),
        "changed_by":     changed_by,
    }])


def get_tp_plant_log():
    """Return change log rows newest-first."""
    df = read_table("tp_plant_change_log", order_by="id DESC")
    if df.empty:
        return []
    return [dict(r) for _, r in df.drop(columns=["id"], errors="ignore").iterrows()]


def add_tp_plant(plant_code, exco_location, plant_name, business_head,
                 plant_manager, mixer_theo_cap, changed_by="Admin"):
    code = str(plant_code).strip()
    if not code:
        raise ValueError("Plant Code cannot be blank.")
    if get_tp_plant(code) is not None:
        raise ValueError(f"Plant Code '{code}' already exists.")
    new_values = {
        "exco_location":  str(exco_location).strip(),
        "plant_name":     str(plant_name).strip(),
        "business_head":  str(business_head).strip(),
        "plant_manager":  str(plant_manager).strip(),
        "mixer_theo_cap": float(mixer_theo_cap) if mixer_theo_cap else 0.0,
    }
    insert_rows("tp_plant_data", [{"plant_code": code, **new_values, "updated_at": _now()}])
    log_tp_plant_change("ADD", code, old_value=None, new_value=new_values,
                        changed_by=changed_by)


def update_tp_plant(plant_code, exco_location, plant_name, business_head,
                    plant_manager, mixer_theo_cap, changed_by="Admin"):
    code = str(plant_code).strip()
    old = get_tp_plant(code)
    if old is None:
        raise ValueError(f"Plant Code '{code}' was not found.")
    new_values = {
        "exco_location":  str(exco_location).strip(),
        "plant_name":     str(plant_name).strip(),
        "business_head":  str(business_head).strip(),
        "plant_manager":  str(plant_manager).strip(),
        "mixer_theo_cap": float(mixer_theo_cap) if mixer_theo_cap else 0.0,
    }
    conn = get_connection()
    try:
        conn.execute(
            "UPDATE tp_plant_data SET exco_location=?, plant_name=?, business_head=?, "
            "plant_manager=?, mixer_theo_cap=?, updated_at=? WHERE plant_code=?",
            (new_values["exco_location"], new_values["plant_name"],
             new_values["business_head"], new_values["plant_manager"],
             new_values["mixer_theo_cap"], _now(), code),
        )
        conn.commit()
    finally:
        conn.close()
    log_tp_plant_change("EDIT", code, old_value=_tp_plant_snapshot(old),
                        new_value=new_values, changed_by=changed_by)


def delete_tp_plant(plant_code, changed_by="Admin"):
    code = str(plant_code).strip()
    old = get_tp_plant(code)
    if old is None:
        raise ValueError(f"Plant Code '{code}' was not found.")
    conn = get_connection()
    try:
        conn.execute("DELETE FROM tp_plant_data WHERE plant_code = ?", (code,))
        conn.commit()
    finally:
        conn.close()
    log_tp_plant_change("DELETE", code, old_value=_tp_plant_snapshot(old),
                        new_value=None, changed_by=changed_by)


# ---------------------------------------------------------------------------
# Waiver CRUD helpers
# ---------------------------------------------------------------------------

def get_waivers(month: int, year: int) -> list:
    """Return all waivers for a given month/year as a list of dicts."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, employee_code, month, year, reason, custom_reason, created_at "
            "FROM waivers WHERE month = ? AND year = ? ORDER BY employee_code",
            (month, year)
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def get_all_waivers() -> list:
    """Return all waivers (all months) as a list of dicts."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, employee_code, month, year, reason, custom_reason, created_at "
            "FROM waivers ORDER BY year DESC, month DESC, employee_code"
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def upsert_waiver(employee_code: str, month: int, year: int,
                  reason: str, custom_reason: str = "") -> None:
    """Insert or replace a waiver for the given employee+month+year."""
    conn = get_connection()
    try:
        conn.execute(
            "INSERT INTO waivers (employee_code, month, year, reason, custom_reason, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(employee_code, month, year) DO UPDATE SET "
            "reason=excluded.reason, custom_reason=excluded.custom_reason, created_at=excluded.created_at",
            (str(employee_code).strip(), month, year, reason, custom_reason or "", _now())
        )
        conn.commit()
    finally:
        conn.close()


def delete_waiver(waiver_id: int) -> None:
    """Delete a waiver by its id."""
    conn = get_connection()
    try:
        conn.execute("DELETE FROM waivers WHERE id = ?", (waiver_id,))
        conn.commit()
    finally:
        conn.close()


def get_waiver_lookup(month: int, year: int) -> dict:
    """Return {employee_code: waiver_display_text} for the given month/year."""
    rows = get_waivers(month, year)
    result = {}
    for r in rows:
        if r["reason"] == "other":
            text = f"Waived, {r['custom_reason']}" if r["custom_reason"] else "Waived"
        elif r["reason"] == "dr_bhoon":
            text = "Waived by Dr. Bhoon"
        elif r["reason"] == "approved_leave":
            text = "Waived, on approved leave"
        else:
            text = f"Waived: {r['reason']}"
        result[str(r["employee_code"]).strip()] = text
    return result


# ---------------------------------------------------------------------------
# RDC-ECMD helpers
# ---------------------------------------------------------------------------

def get_ecmd_reading(plant_code: str, month: int, year: int) -> dict | None:
    """Return the readings row for a plant/month/year, or None."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM ecmd_readings WHERE plant_code=? AND month=? AND year=?",
            (str(plant_code).strip(), month, year)
        )
        row = cur.fetchone()
        if row is None:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))
    finally:
        conn.close()


def get_ecmd_readings_for_month(month: int, year: int) -> list:
    """Return all readings rows for a given month/year as list of dicts."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM ecmd_readings WHERE month=? AND year=? ORDER BY plant_code",
            (month, year)
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def get_ecmd_all_readings() -> list:
    """Return all readings rows (all months) newest-first."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM ecmd_readings ORDER BY year DESC, month DESC, plant_code"
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def upsert_ecmd_reading(plant_code: str, month: int, year: int, data: dict,
                        entered_by: str = "Admin") -> None:
    """Insert or update a reading row (UPSERT on plant_code+month+year)."""
    conn = get_connection()
    try:
        conn.execute(
            """INSERT INTO ecmd_readings
               (plant_code, month, year,
                eb_kwh_open, eb_kwh_close, eb_kvah_open, eb_kvah_close, mf,
                dg_hr_open, dg_hr_close, dg_kwh_open, dg_kwh_close,
                mixer_dg_hr_open, mixer_dg_hr_close,
                diesel_issued_ltrs, volume_on_dg,
                entered_by, entered_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
               ON CONFLICT(plant_code, month, year) DO UPDATE SET
                eb_kwh_open=excluded.eb_kwh_open, eb_kwh_close=excluded.eb_kwh_close,
                eb_kvah_open=excluded.eb_kvah_open, eb_kvah_close=excluded.eb_kvah_close,
                mf=excluded.mf,
                dg_hr_open=excluded.dg_hr_open, dg_hr_close=excluded.dg_hr_close,
                dg_kwh_open=excluded.dg_kwh_open, dg_kwh_close=excluded.dg_kwh_close,
                mixer_dg_hr_open=excluded.mixer_dg_hr_open,
                mixer_dg_hr_close=excluded.mixer_dg_hr_close,
                diesel_issued_ltrs=excluded.diesel_issued_ltrs,
                volume_on_dg=excluded.volume_on_dg,
                entered_by=excluded.entered_by, entered_at=excluded.entered_at
            """,
            (
                str(plant_code).strip(), month, year,
                data.get("eb_kwh_open"), data.get("eb_kwh_close"),
                data.get("eb_kvah_open"), data.get("eb_kvah_close"),
                data.get("mf", 1.0),
                data.get("dg_hr_open"), data.get("dg_hr_close"),
                data.get("dg_kwh_open"), data.get("dg_kwh_close"),
                data.get("mixer_dg_hr_open"), data.get("mixer_dg_hr_close"),
                data.get("diesel_issued_ltrs"), data.get("volume_on_dg"),
                entered_by, _now(),
            )
        )
        conn.commit()
    finally:
        conn.close()


def delete_ecmd_reading(plant_code: str, month: int, year: int) -> int:
    """Delete one reading row. Returns rows deleted."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM ecmd_readings WHERE plant_code=? AND month=? AND year=?",
            (str(plant_code).strip(), month, year)
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


# ── ECMD entry mode ──────────────────────────────────────────────────────────

def get_ecmd_entry_mode(plant_code: str, month: int, year: int) -> str:
    """Return 'monthly', 'daily', or 'none' for this plant+month+year. Default: 'none'."""
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT mode FROM ecmd_entry_mode WHERE plant_code=? AND month=? AND year=?",
            (str(plant_code).strip(), month, year)
        )
        row = cur.fetchone()
        return row[0] if row else "none"
    finally:
        conn.close()


def set_ecmd_entry_mode(plant_code: str, month: int, year: int, mode: str) -> None:
    """Set entry mode ('monthly' or 'daily') for this plant+month+year."""
    conn = get_connection()
    try:
        conn.execute(
            """INSERT INTO ecmd_entry_mode(plant_code, month, year, mode)
               VALUES(?,?,?,?)
               ON CONFLICT(plant_code, month, year) DO UPDATE SET mode=excluded.mode""",
            (str(plant_code).strip(), month, year, mode)
        )
        conn.commit()
    finally:
        conn.close()


# ── ECMD daily readings ───────────────────────────────────────────────────────

def get_ecmd_daily_readings(plant_code: str, month: int, year: int) -> list:
    """Return all daily reading rows for a plant+month+year, ordered by day."""
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM ecmd_daily_readings WHERE plant_code=? AND month=? AND year=? ORDER BY day",
            (str(plant_code).strip(), month, year)
        )
        rows = cur.fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_ecmd_daily_readings_for_month(month: int, year: int) -> list:
    """Return all daily readings for a month/year across all plants."""
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM ecmd_daily_readings WHERE month=? AND year=? ORDER BY plant_code, day",
            (month, year)
        )
        rows = cur.fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def upsert_ecmd_daily_reading(plant_code: str, month: int, year: int, day: int,
                               data: dict, entered_by: str = "Admin") -> None:
    """Insert or update a single day's reading."""
    import datetime as _dt
    fields = ["eb_kwh_open","eb_kwh_close","eb_kvah_open","eb_kvah_close","mf",
              "dg_hr_open","dg_hr_close","dg_kwh_open","dg_kwh_close",
              "mixer_dg_hr_open","mixer_dg_hr_close"]
    cols   = ", ".join(["plant_code","month","year","day","entered_by","entered_at"] + fields)
    vals   = ", ".join(["?"] * (6 + len(fields)))
    updates = ", ".join([f"{f}=excluded.{f}" for f in fields] + ["entered_by=excluded.entered_by","entered_at=excluded.entered_at"])
    row = ([str(plant_code).strip(), month, year, day,
            entered_by, str(_dt.datetime.now())] +
           [data.get(f) for f in fields])
    conn = get_connection()
    try:
        conn.execute(
            f"INSERT INTO ecmd_daily_readings({cols}) VALUES({vals}) "
            f"ON CONFLICT(plant_code,month,year,day) DO UPDATE SET {updates}",
            row
        )
        conn.commit()
    finally:
        conn.close()


def delete_ecmd_daily_reading(plant_code: str, month: int, year: int, day: int) -> int:
    """Delete one day's reading. Returns rows deleted."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM ecmd_daily_readings WHERE plant_code=? AND month=? AND year=? AND day=?",
            (str(plant_code).strip(), month, year, day)
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def delete_ecmd_all_daily_readings(plant_code: str, month: int, year: int) -> int:
    """Delete all daily readings for a plant+month+year. Returns rows deleted."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM ecmd_daily_readings WHERE plant_code=? AND month=? AND year=?",
            (str(plant_code).strip(), month, year)
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def get_ecmd_results_for_month(month: int, year: int) -> list:
    """Return calculated ECMD result rows for a month/year."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM ecmd_results WHERE month=? AND year=? ORDER BY plant_code",
            (month, year)
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def save_ecmd_results(rows: list, month: int, year: int) -> None:
    """Replace all ecmd_results for the given month/year with new rows."""
    conn = get_connection()
    try:
        conn.execute(
            "DELETE FROM ecmd_results WHERE month=? AND year=?", (month, year)
        )
        conn.commit()
    finally:
        conn.close()
    if rows:
        insert_rows("ecmd_results", rows)


def get_ecmd_months() -> list:
    """Return list of (month, year) tuples that have readings, newest-first."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT DISTINCT month, year FROM ecmd_readings "
            "ORDER BY year DESC, month DESC"
        )
        return [(r[0], r[1]) for r in cur.fetchall()]
    finally:
        conn.close()


def get_ecmd_mf(plant_code: str) -> float:
    """Return the most recently entered MF for a plant (default 1.0)."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT mf FROM ecmd_readings WHERE plant_code=? "
            "ORDER BY year DESC, month DESC LIMIT 1",
            (str(plant_code).strip(),)
        )
        row = cur.fetchone()
        return float(row[0]) if row and row[0] else 1.0
    finally:
        conn.close()


def save_dual_plant_report(period_label: str, from_date: str, to_date: str, rows: list):
    """Replace cached dual-plant rows for a given period_label."""
    conn = get_connection()
    try:
        conn.execute("DELETE FROM ecmd_dual_plant WHERE period_label=?", (period_label,))
        for r in rows:
            conn.execute(
                """INSERT INTO ecmd_dual_plant
                   (period_label,from_date,to_date,plant_code,plant_name,mixer,quantity,pct_share,fetched_at)
                   VALUES(?,?,?,?,?,?,?,?,?)""",
                (period_label, from_date, to_date,
                 r["plant_code"], r.get("plant_name",""), r["mixer"],
                 r["quantity"], r["pct_share"], r.get("fetched_at",""))
            )
        conn.commit()
    finally:
        conn.close()


def get_dual_plant_report(period_label: str) -> list:
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM ecmd_dual_plant WHERE period_label=? ORDER BY plant_code, mixer",
            (period_label,)
        )
        return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def get_dual_plant_periods() -> list:
    """Return distinct period_labels available, newest first."""
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT period_label, from_date, to_date, MAX(fetched_at) AS fetched_at "
            "FROM ecmd_dual_plant GROUP BY period_label, from_date, to_date "
            "ORDER BY from_date DESC"
        )
        return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def save_invoice_pending_report(period_label: str, from_date: str, to_date: str, rows: list):
    """Replace cached invoice-pending rows for a given period_label."""
    conn = get_connection()
    try:
        # Add new columns to existing DB if missing (migration)
        for col, coltype in [("sales_order", "TEXT"), ("line_number", "TEXT")]:
            try:
                conn.execute(f"ALTER TABLE ecmd_invoice_pending ADD COLUMN {col} {coltype}")
                conn.commit()
            except Exception:
                pass
        conn.execute("DELETE FROM ecmd_invoice_pending WHERE period_label=?", (period_label,))
        for r in rows:
            conn.execute(
                """INSERT INTO ecmd_invoice_pending
                   (period_label,from_date,to_date,plant_code,plant_name,sales_order,line_number,quantity,fetched_at)
                   VALUES(?,?,?,?,?,?,?,?,?)""",
                (period_label, from_date, to_date,
                 r["plant_code"], r.get("plant_name",""),
                 r.get("sales_order",""), r.get("line_number",""),
                 r["quantity"], r.get("fetched_at",""))
            )
        conn.commit()
    finally:
        conn.close()


def get_invoice_pending_report(period_label: str) -> list:
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM ecmd_invoice_pending WHERE period_label=? ORDER BY quantity DESC",
            (period_label,)
        )
        return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def get_invoice_pending_periods() -> list:
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT period_label, from_date, to_date, MAX(fetched_at) AS fetched_at "
            "FROM ecmd_invoice_pending GROUP BY period_label, from_date, to_date "
            "ORDER BY from_date DESC"
        )
        return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def get_ecmd_allowed_months() -> list:
    """Return list of (month, year) tuples admin has unlocked for data entry."""
    raw = get_module_setting("ecmd", "allowed_months", "")
    result = []
    for part in str(raw).split(","):
        part = part.strip()
        if not part:
            continue
        try:
            m, y = part.split("-")
            result.append((int(m), int(y)))
        except Exception:
            pass
    return result


def set_ecmd_allowed_months(pairs: list):
    """Save list of (month, year) tuples. Pass [] to allow all."""
    encoded = ",".join(f"{m}-{y}" for m, y in pairs)
    set_module_setting("ecmd", "allowed_months", encoded)


# ---------------------------------------------------------------------------
# Authentication / Authorisation DB helpers
# ---------------------------------------------------------------------------

def _user_row_to_dict(row) -> dict:
    """Convert a sqlite3.Row or tuple+description to a plain dict."""
    if row is None:
        return None
    cols = [d[0] for d in row.cursor_description] if hasattr(row, "cursor_description") else None
    if hasattr(row, "keys"):
        return dict(row)
    return dict(row)


def count_users() -> int:
    """Return total number of rows in users table."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM users")
        return cur.fetchone()[0]
    finally:
        conn.close()


def create_user(full_name: str, email: str, username: str, password_hash: str,
                role: str, is_active: bool = True,
                must_change_password: bool = True) -> int:
    """Insert a new user. Returns the new user id."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO users
               (full_name, email, username, password_hash, role,
                is_active, must_change_password, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (full_name.strip(), email.strip(), username.strip(), password_hash,
             role, 1 if is_active else 0, 1 if must_change_password else 0,
             _now(), _now())
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def get_user_by_id(user_id: int) -> dict | None:
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE id = ?", (user_id,))
        row = cur.fetchone()
        if row is None:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))
    finally:
        conn.close()


def get_user_by_username(username: str) -> dict | None:
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username = ?",
                    (username.strip(),))
        row = cur.fetchone()
        if row is None:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))
    finally:
        conn.close()


def get_user_by_email(email: str) -> dict | None:
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE LOWER(email) = LOWER(?)",
                    (email.strip(),))
        row = cur.fetchone()
        if row is None:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))
    finally:
        conn.close()


def get_all_users() -> list:
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM users ORDER BY role, full_name"
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def update_user(user_id: int, full_name: str, email: str, role: str,
                is_active: bool, must_change_password: bool) -> None:
    conn = get_connection()
    try:
        conn.execute(
            """UPDATE users SET full_name=?, email=?, role=?,
               is_active=?, must_change_password=?, updated_at=?
               WHERE id=?""",
            (full_name.strip(), email.strip(), role,
             1 if is_active else 0, 1 if must_change_password else 0,
             _now(), user_id)
        )
        conn.commit()
    finally:
        conn.close()


def update_user_password(user_id: int, password_hash: str,
                         must_change_password: bool = False) -> None:
    conn = get_connection()
    try:
        conn.execute(
            "UPDATE users SET password_hash=?, must_change_password=?, updated_at=? WHERE id=?",
            (password_hash, 1 if must_change_password else 0, _now(), user_id)
        )
        conn.commit()
    finally:
        conn.close()


def update_last_login(user_id: int) -> None:
    conn = get_connection()
    try:
        conn.execute(
            "UPDATE users SET last_login_at=? WHERE id=?", (_now(), user_id)
        )
        conn.commit()
    finally:
        conn.close()


def delete_user(user_id: int) -> None:
    conn = get_connection()
    try:
        conn.execute("DELETE FROM user_plant_access WHERE user_id=?", (user_id,))
        conn.execute("DELETE FROM users WHERE id=?", (user_id,))
        conn.commit()
    finally:
        conn.close()


# ── User Plant Access ─────────────────────────────────────────────────────────

def get_user_plant_access(user_id: int) -> list:
    """Return all plant rows assigned to a user."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM user_plant_access WHERE user_id=? ORDER BY plant_name",
            (user_id,)
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def assign_plant_to_user(user_id: int, plant_code: str, plant_name: str) -> None:
    """Add one plant assignment (ignore if already exists)."""
    conn = get_connection()
    try:
        conn.execute(
            """INSERT OR IGNORE INTO user_plant_access
               (user_id, plant_code, plant_name, created_at)
               VALUES (?,?,?,?)""",
            (user_id, plant_code.strip(), plant_name.strip(), _now())
        )
        conn.commit()
    finally:
        conn.close()


def remove_plant_from_user(user_id: int, plant_name: str) -> None:
    conn = get_connection()
    try:
        conn.execute(
            "DELETE FROM user_plant_access WHERE user_id=? AND plant_name=?",
            (user_id, plant_name.strip())
        )
        conn.commit()
    finally:
        conn.close()


def set_user_plants(user_id: int, plant_list: list[dict]) -> None:
    """Replace all plant assignments for a user. plant_list = [{plant_code, plant_name}]."""
    conn = get_connection()
    try:
        conn.execute("DELETE FROM user_plant_access WHERE user_id=?", (user_id,))
        for p in plant_list:
            conn.execute(
                "INSERT OR IGNORE INTO user_plant_access (user_id, plant_code, plant_name, created_at) VALUES (?,?,?,?)",
                (user_id, p.get("plant_code", "").strip(), p.get("plant_name", "").strip(), _now())
            )
        conn.commit()
    finally:
        conn.close()


# ── Audit Logs ────────────────────────────────────────────────────────────────

def log_login_attempt(user_id, ip_address: str, user_agent: str,
                      status: str, failure_reason: str = "") -> None:
    conn = get_connection()
    try:
        conn.execute(
            """INSERT INTO login_audit_log
               (user_id, login_time, ip_address, user_agent, status, failure_reason)
               VALUES (?,?,?,?,?,?)""",
            (user_id, _now(), ip_address[:45], user_agent, status, failure_reason)
        )
        conn.commit()
    finally:
        conn.close()


def log_user_activity(user_id, action: str, module_name: str = "",
                      details_json: str = "{}", ip_address: str = "") -> None:
    conn = get_connection()
    try:
        conn.execute(
            """INSERT INTO user_activity_log
               (user_id, action, module_name, details_json, created_at, ip_address)
               VALUES (?,?,?,?,?,?)""",
            (user_id, action, module_name, details_json, _now(), ip_address[:45])
        )
        conn.commit()
    finally:
        conn.close()


def get_login_audit_log(limit: int = 200) -> list:
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute(
            """SELECT l.*, u.username, u.full_name, u.role
               FROM login_audit_log l
               LEFT JOIN users u ON u.id = l.user_id
               ORDER BY l.id DESC LIMIT ?""",
            (limit,)
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def get_user_activity_log(limit: int = 500, user_id: int = None) -> list:
    conn = get_connection()
    try:
        cur = conn.cursor()
        if user_id:
            cur.execute(
                """SELECT a.*, u.username, u.full_name
                   FROM user_activity_log a
                   LEFT JOIN users u ON u.id = a.user_id
                   WHERE a.user_id=?
                   ORDER BY a.id DESC LIMIT ?""",
                (user_id, limit)
            )
        else:
            cur.execute(
                """SELECT a.*, u.username, u.full_name
                   FROM user_activity_log a
                   LEFT JOIN users u ON u.id = a.user_id
                   ORDER BY a.id DESC LIMIT ?""",
                (limit,)
            )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# SLA — Slow Loading Alert helpers
# ---------------------------------------------------------------------------

def sla_get_thresholds() -> list:
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM slow_loading_thresholds WHERE is_active=1 ORDER BY mixer_capacity"
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def sla_get_all_thresholds() -> list:
    conn = get_connection()
    try:
        cur = conn.execute("SELECT * FROM slow_loading_thresholds ORDER BY mixer_capacity")
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def sla_upsert_threshold(mixer_capacity, reference_quantity, base_allowed_minutes, threshold_id=None):
    now_ts = _now()
    conn = get_connection()
    try:
        if threshold_id:
            conn.execute(
                "UPDATE slow_loading_thresholds "
                "SET mixer_capacity=?, reference_quantity=?, base_allowed_minutes=?, updated_at=? WHERE id=?",
                (mixer_capacity, reference_quantity, base_allowed_minutes, now_ts, threshold_id)
            )
        else:
            conn.execute(
                "INSERT INTO slow_loading_thresholds "
                "(mixer_capacity, reference_quantity, base_allowed_minutes, is_active, created_at, updated_at) "
                "VALUES (?, ?, ?, 1, ?, ?)",
                (mixer_capacity, reference_quantity, base_allowed_minutes, now_ts, now_ts)
            )
        conn.commit()
    finally:
        conn.close()


def sla_delete_threshold(threshold_id):
    conn = get_connection()
    try:
        conn.execute("DELETE FROM slow_loading_thresholds WHERE id=?", (threshold_id,))
        conn.commit()
    finally:
        conn.close()


def sla_alert_key_exists_this_hour(alert_key, alert_date, alert_hour):
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT 1 FROM slow_loading_alert_records "
            "WHERE alert_key=? AND alert_date=? AND alert_hour=? AND status='SENT' LIMIT 1",
            (alert_key, alert_date, alert_hour)
        )
        return cur.fetchone() is not None
    finally:
        conn.close()


def sla_bulk_insert_records(rows):
    if not rows:
        return 0
    now_ts = _now()
    for r in rows:
        r.setdefault("created_at", now_ts)
        r.setdefault("updated_at", now_ts)
    return insert_rows("slow_loading_alert_records", rows)


def sla_mark_records_sent(record_ids):
    if not record_ids:
        return
    now_ts = _now()
    conn = get_connection()
    try:
        conn.executemany(
            "UPDATE slow_loading_alert_records SET status='SENT', updated_at=? WHERE id=?",
            [(now_ts, rid) for rid in record_ids]
        )
        conn.commit()
    finally:
        conn.close()


def sla_mark_records_failed(record_ids, error=""):
    if not record_ids:
        return
    now_ts = _now()
    conn = get_connection()
    try:
        conn.executemany(
            "UPDATE slow_loading_alert_records SET status='FAILED', remarks=?, updated_at=? WHERE id=?",
            [(error[:500], now_ts, rid) for rid in record_ids]
        )
        conn.commit()
    finally:
        conn.close()


def sla_log_email(alert_type, alert_date, alert_hour, plant_code, plant_name,
                   to_emails, cc_emails, subject, total_cases, status, error_message=""):
    now_ts = _now()
    insert_rows("slow_loading_email_logs", [{
        "alert_type": alert_type, "alert_date": alert_date, "alert_hour": alert_hour,
        "plant_code": plant_code, "plant_name": plant_name,
        "to_emails": to_emails, "cc_emails": cc_emails,
        "subject": subject, "total_cases": total_cases,
        "status": status, "error_message": error_message,
        "sent_at": now_ts, "created_at": now_ts,
    }])


def sla_log_scheduler(job_name, started_at, completed_at, status,
                       total_checked=0, total_alerts=0, total_sent=0, error_message=""):
    insert_rows("slow_loading_scheduler_logs", [{
        "job_name": job_name, "run_started_at": started_at,
        "run_completed_at": completed_at, "status": status,
        "total_records_checked": total_checked, "total_alert_cases": total_alerts,
        "total_emails_sent": total_sent, "error_message": error_message,
    }])


def sla_get_report(from_date=None, to_date=None, plant_code=None, batcher_code=None,
                    tm_number=None, grade=None, customer=None,
                    alert_type=None, status=None, limit=500):
    clauses, params = [], []
    if from_date:
        clauses.append("alert_date >= ?"); params.append(from_date)
    if to_date:
        clauses.append("alert_date <= ?"); params.append(to_date)
    if plant_code:
        clauses.append("plant_code = ?"); params.append(plant_code)
    if batcher_code:
        clauses.append("batcher_code = ?"); params.append(batcher_code)
    if tm_number:
        clauses.append("tm_number LIKE ?"); params.append(f"%{tm_number}%")
    if grade:
        clauses.append("grade LIKE ?"); params.append(f"%{grade}%")
    if customer:
        clauses.append("customer LIKE ?"); params.append(f"%{customer}%")
    if alert_type:
        clauses.append("alert_type = ?"); params.append(alert_type)
    if status:
        clauses.append("status = ?"); params.append(status)
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    sql = (f"SELECT * FROM slow_loading_alert_records {where} "
           f"ORDER BY alert_date DESC, alert_hour DESC, delay_minutes DESC LIMIT ?")
    params.append(limit)
    conn = get_connection()
    try:
        cur = conn.execute(sql, params)
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def sla_get_dashboard_kpis(today_str, plant_codes=None):
    conn = get_connection()
    try:
        plant_filter = ""
        params_base = [today_str]
        if plant_codes:
            placeholders = ",".join("?" * len(plant_codes))
            plant_filter = f"AND plant_code IN ({placeholders})"
            params_base = [today_str] + list(plant_codes)

        def _one(sql, p):
            r = conn.execute(sql, p).fetchone()
            return r[0] if r else 0

        vehicles_today = _one(
            f"SELECT COUNT(DISTINCT tm_number) FROM slow_loading_alert_records WHERE alert_date=? {plant_filter}",
            params_base)
        slow_today = _one(
            f"SELECT COUNT(*) FROM slow_loading_alert_records WHERE alert_date=? AND status!='SKIPPED_DUPLICATE' {plant_filter}",
            params_base)
        hourly_sent = _one(
            f"SELECT COUNT(*) FROM slow_loading_email_logs WHERE alert_date=? AND alert_type='HOURLY' AND status='SENT'",
            [today_str])
        daily_sent = _one(
            f"SELECT COUNT(*) FROM slow_loading_email_logs WHERE alert_date=? AND alert_type='DAILY_SUMMARY' AND status='SENT'",
            [today_str])

        row = conn.execute(
            f"SELECT AVG(loading_time_minutes), AVG(allowed_loading_minutes), MAX(delay_minutes) "
            f"FROM slow_loading_alert_records WHERE alert_date=? {plant_filter}",
            params_base).fetchone()
        avg_load    = round(row[0] or 0, 1)
        avg_allowed = round(row[1] or 0, 1)
        max_delay   = round(row[2] or 0, 1)

        last_run = conn.execute(
            "SELECT run_completed_at FROM slow_loading_scheduler_logs "
            "WHERE status='SUCCESS' ORDER BY id DESC LIMIT 1"
        ).fetchone()
        last_run_time = last_run[0] if last_run else "Never"

        return {
            "vehicles_today": vehicles_today, "slow_today": slow_today,
            "hourly_sent": hourly_sent, "daily_sent": daily_sent,
            "avg_load": avg_load, "avg_allowed": avg_allowed,
            "max_delay": max_delay, "last_run_time": last_run_time,
        }
    finally:
        conn.close()


def sla_get_email_logs(limit=200):
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM slow_loading_email_logs ORDER BY id DESC LIMIT ?", (limit,)
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def sla_get_scheduler_logs(limit=100):
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT * FROM slow_loading_scheduler_logs ORDER BY id DESC LIMIT ?", (limit,)
        )
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
    finally:
        conn.close()


def sla_get_distinct_plants():
    conn = get_connection()
    try:
        cur = conn.execute(
            "SELECT DISTINCT plant_code, plant_name FROM slow_loading_alert_records "
            "WHERE plant_code IS NOT NULL ORDER BY plant_name"
        )
        return [{"plant_code": r[0], "plant_name": r[1]} for r in cur.fetchall()]
    finally:
        conn.close()
