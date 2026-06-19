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
    "backend_data":      "date",             # RDC-I&D
    "tp_oracle_data":    "production_date",  # RDC-TP
    "btrtp_oracle_data": "production_date",  # RDC-BTRTP
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
