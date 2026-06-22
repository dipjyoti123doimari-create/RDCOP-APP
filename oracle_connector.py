"""
oracle_connector.py
===================
Connects to the RDC enterprise Oracle database and fetches backend production
data directly from rdc_batch_trx_headers — no Excel upload needed.

Uses python-oracledb in THICK mode (requires Oracle Instant Client).
The Instant Client folder path is stored in app_settings.

Column mapping (Oracle → app):
    CREATED_BY        -> created_by  (employee code)
    PRODDATE          -> date        (production date, YYYY-MM-DD)
    PRODUCED_QUANTITY -> quantity    (quantity produced)
    STATUS filter     -> 'Processed' by default

Public functions:
    get_oracle_config()              -> dict of connection settings from DB
    test_connection()                -> {"success": bool, "error": str, "version": str}
    fetch_backend_data(from, to)     -> (DataFrame, warnings_list)
    save_oracle_backend_data(df, ..) -> rows_inserted (int)
"""

import socket

import oracledb
import pandas as pd
from datetime import datetime

import database

# Schema + table in Oracle that holds the production data.
_TABLE = "APPSREAD.rdc_batch_trx_headers"

# Default path to Oracle Instant Client — overridden by app_settings.
_DEFAULT_INSTANTCLIENT = r"D:\AI Project\Incentive Calculator\instantclient"

# oracledb thick mode can only be initialised once per Python process.
_thick_initialized = False


def get_oracle_config() -> dict:
    """Read Oracle connection settings from app_settings (single DB call)."""
    s = database.get_all_settings()
    return {
        "host":          (s.get("oracle_host", "192.168.100.11") or "").strip(),
        "port":          (s.get("oracle_port", "1528") or "1528").strip(),
        "service":       (s.get("oracle_service", "RDCAZPRD") or "").strip(),
        "user":          (s.get("oracle_user", "RDCREAD") or "").strip(),
        "password":      (s.get("oracle_password", "") or ""),
        "instantclient": (s.get("oracle_instantclient_dir", _DEFAULT_INSTANTCLIENT) or _DEFAULT_INSTANTCLIENT).strip(),
        "status_filter": (s.get("oracle_status_filter", "Processed") or "").strip(),
    }


def is_configured(cfg: dict = None) -> bool:
    """True when host, user and password are all set (config present only —
    does NOT mean the server is reachable). Use is_reachable() for that."""
    cfg = cfg or get_oracle_config()
    return bool(cfg["host"] and cfg["user"] and cfg["password"])


def is_reachable(cfg: dict = None, timeout: float = 2.0) -> bool:
    """
    Fast, honest liveness check: open a raw TCP socket to host:port.

    Returns True only if the Oracle listener actually answers. On the office
    network this returns in a few milliseconds; off-network it fails after
    `timeout` seconds. This is NOT a full Oracle login — it just proves the
    server is reachable, which is what the status indicator needs.
    """
    cfg = cfg or get_oracle_config()
    host, port = cfg.get("host", ""), cfg.get("port", "")
    if not host or not port:
        return False
    try:
        with socket.create_connection((host, int(port)), timeout=timeout):
            return True
    except Exception:
        return False


def _init_thick(lib_dir: str):
    """Initialise oracledb thick mode — safe to call multiple times."""
    global _thick_initialized
    if not _thick_initialized:
        oracledb.init_oracle_client(lib_dir=lib_dir)
        _thick_initialized = True


def _dsn(cfg: dict) -> str:
    return f"{cfg['host']}:{cfg['port']}/{cfg['service']}"


def test_connection() -> dict:
    """
    Try to open a connection and return basic Oracle version info.
    Returns {"success": bool, "error": str|None, "version": str}.
    """
    cfg = get_oracle_config()
    try:
        _init_thick(cfg["instantclient"])
        conn = oracledb.connect(
            user=cfg["user"], password=cfg["password"], dsn=_dsn(cfg)
        )
        cur = conn.cursor()
        cur.execute("SELECT * FROM v$version WHERE rownum = 1")
        version = cur.fetchone()[0]
        conn.close()
        return {"success": True, "error": None, "version": version}
    except Exception as exc:
        return {"success": False, "error": str(exc), "version": ""}


def fetch_backend_data(from_date, to_date) -> tuple:
    """
    Fetch production rows from Oracle for the given date range.

    Returns:
        (DataFrame, warnings_list)

    The DataFrame has columns:
        created_by, date (YYYY-MM-DD string), quantity
    ready to be passed straight to save_oracle_backend_data().
    """
    cfg = get_oracle_config()
    warnings = []

    _init_thick(cfg["instantclient"])
    conn = oracledb.connect(
        user=cfg["user"], password=cfg["password"], dsn=_dsn(cfg)
    )
    try:
        cur = conn.cursor()

        # Build status filter clause only when a filter value is configured.
        # PRODDATE is VARCHAR2 stored as 'YYYY-MM-DD' strings, so plain
        # string comparison works correctly for date ranges.
        fd = str(from_date)
        td = str(to_date)

        status_clause = ""
        params = {"from_date": fd, "to_date": td}
        if cfg["status_filter"]:
            status_clause = "AND STATUS = :status"
            params["status"] = cfg["status_filter"]

        # Use the same configurable time column as TP (default TIMETAKEN)
        time_col = (database.get_module_setting("tp", "oracle_time_col", "TIMETAKEN") or "TIMETAKEN").strip()

        sql = f"""
            SELECT
                CREATED_BY        AS created_by,
                PRODDATE          AS prod_date,
                PRODUCED_QUANTITY AS quantity,
                {time_col}        AS time_taken
            FROM {_TABLE}
            WHERE PRODDATE >= :from_date
              AND PRODDATE <= :to_date
              {status_clause}
            ORDER BY PRODDATE, CREATED_BY
        """
        cur.execute(sql, params)
        rows = cur.fetchall()
        cols = [d[0].lower() for d in cur.description]

        if not rows:
            warnings.append(
                f"No rows found in Oracle for {from_date} → {to_date}"
                + (f" with STATUS = '{cfg['status_filter']}'." if cfg["status_filter"] else ".")
            )
            return pd.DataFrame(columns=["created_by", "date", "quantity"]), warnings

        df = pd.DataFrame(rows, columns=cols)
        df = df.rename(columns={"prod_date": "date"})
        df["quantity"]   = pd.to_numeric(df["quantity"],   errors="coerce").fillna(0)
        df["time_taken"] = pd.to_numeric(df["time_taken"], errors="coerce").fillna(0)

        # Drop rows where time_taken = 0 (quantity should not be counted per business rule)
        zero_time = df["time_taken"] == 0
        if zero_time.sum():
            warnings.append(f"{zero_time.sum()} row(s) skipped — time taken = 0.")
        df = df[~zero_time].reset_index(drop=True)

        # Drop rows with blank employee code.
        before = len(df)
        df["created_by"] = df["created_by"].fillna("").astype(str).str.strip()
        df = df[df["created_by"] != ""].reset_index(drop=True)
        dropped = before - len(df)
        if dropped:
            warnings.append(f"{dropped} row(s) skipped — empty CREATED_BY.")

        # Drop zero/negative quantity rows.
        bad_qty = df["quantity"] <= 0
        if bad_qty.sum():
            warnings.append(f"{bad_qty.sum()} row(s) skipped — zero or negative PRODUCED_QUANTITY.")
        df = df[~bad_qty].reset_index(drop=True)

        # Drop the time_taken column before returning (backend_data table has no such column)
        df = df.drop(columns=["time_taken"], errors="ignore")

        return df, warnings

    finally:
        conn.close()


# ── RDC-TP: fetch throughput data from same Oracle table ─────────────────────

def get_table_columns() -> list:
    """
    Return all column names in rdc_batch_trx_headers from Oracle's data dict.
    Used to help the user identify the correct plant/batch/time column names.
    """
    cfg = get_oracle_config()
    _init_thick(cfg["instantclient"])
    conn = oracledb.connect(user=cfg["user"], password=cfg["password"], dsn=_dsn(cfg))
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT COLUMN_NAME FROM ALL_TAB_COLUMNS "
            "WHERE TABLE_NAME = 'RDC_BATCH_TRX_HEADERS' "
            "ORDER BY COLUMN_ID"
        )
        return [row[0] for row in cur.fetchall()]
    finally:
        conn.close()


def get_tp_oracle_cols() -> dict:
    """
    Return the Oracle column names used for the TP module.
    Defaults can be overridden in Settings → RDC-TP → Oracle Column Names.
    """
    return {
        "plant":  (database.get_module_setting("tp", "oracle_plant_col", "PLANTNO")   or "PLANTNO").strip(),
        "batch":  (database.get_module_setting("tp", "oracle_batch_col", "BATCHCODE") or "BATCHCODE").strip(),
        "time":   (database.get_module_setting("tp", "oracle_time_col",  "TIMETAKEN") or "TIMETAKEN").strip(),
    }


def fetch_tp_data(from_date, to_date) -> tuple:
    """
    Fetch production rows for RDC-TP from Oracle.

    Returns (DataFrame, warnings_list).
    DataFrame columns: production_date, plant_col, batch_ref, quantity, time_taken_min
    """
    cfg  = get_oracle_config()
    cols = get_tp_oracle_cols()
    warnings = []

    _init_thick(cfg["instantclient"])
    conn = oracledb.connect(user=cfg["user"], password=cfg["password"], dsn=_dsn(cfg))
    try:
        cur = conn.cursor()
        fd, td = str(from_date), str(to_date)
        params = {"from_date": fd, "to_date": td}
        status_clause = ""
        if cfg["status_filter"]:
            status_clause = "AND STATUS = :status"
            params["status"] = cfg["status_filter"]

        sql = f"""
            SELECT
                PRODDATE                   AS production_date,
                {cols['plant']}            AS plant_col,
                {cols['batch']}            AS batch_ref,
                PRODUCED_QUANTITY          AS quantity,
                {cols['time']}             AS time_taken_min
            FROM {_TABLE}
            WHERE PRODDATE >= :from_date
              AND PRODDATE <= :to_date
              {status_clause}
            ORDER BY PRODDATE, {cols['plant']}
        """
        cur.execute(sql, params)
        rows = cur.fetchall()
        col_names = [d[0].lower() for d in cur.description]

        if not rows:
            warnings.append(f"No rows found in Oracle for {from_date} → {to_date}.")
            return pd.DataFrame(columns=["production_date","plant_col","batch_ref","quantity","time_taken_min"]), warnings

        df = pd.DataFrame(rows, columns=col_names)
        df["quantity"]      = pd.to_numeric(df["quantity"],      errors="coerce").fillna(0)
        df["time_taken_min"] = pd.to_numeric(df["time_taken_min"], errors="coerce").fillna(0)

        # Warn about missing batch/plant values
        blank_batch = df["batch_ref"].isna() | (df["batch_ref"].astype(str).str.strip() == "")
        if blank_batch.sum():
            warnings.append(f"{blank_batch.sum()} row(s) have blank batch reference — skipped.")
        df = df[~blank_batch].reset_index(drop=True)

        return df, warnings

    finally:
        conn.close()


def fetch_btrtp_data(from_date, to_date) -> tuple:
    """
    Fetch production rows for RDC-BTRTP from Oracle.
    Same table as TP but adds CREATED_BY as batcher_id.

    Returns (DataFrame, warnings_list).
    DataFrame columns: production_date, batcher_id, plant_col, batch_ref, quantity, time_taken_min
    """
    cfg  = get_oracle_config()
    cols = get_tp_oracle_cols()
    batcher_col = (database.get_module_setting("btrtp", "oracle_batcher_col", "CREATED_BY") or "CREATED_BY").strip()
    warnings = []

    _init_thick(cfg["instantclient"])
    conn = oracledb.connect(user=cfg["user"], password=cfg["password"], dsn=_dsn(cfg))
    try:
        cur = conn.cursor()
        params = {"from_date": str(from_date), "to_date": str(to_date)}
        status_clause = ""
        if cfg["status_filter"]:
            status_clause = "AND STATUS = :status"
            params["status"] = cfg["status_filter"]

        sql = f"""
            SELECT
                PRODDATE                   AS production_date,
                {batcher_col}              AS batcher_id,
                {cols['plant']}            AS plant_col,
                {cols['batch']}            AS batch_ref,
                PRODUCED_QUANTITY          AS quantity,
                {cols['time']}             AS time_taken_min
            FROM {_TABLE}
            WHERE PRODDATE >= :from_date
              AND PRODDATE <= :to_date
              {status_clause}
            ORDER BY PRODDATE, {batcher_col}
        """
        cur.execute(sql, params)
        rows = cur.fetchall()
        col_names = [d[0].lower() for d in cur.description]

        if not rows:
            warnings.append(f"No rows found in Oracle for {from_date} → {to_date}.")
            return pd.DataFrame(columns=["production_date","batcher_id","plant_col","batch_ref","quantity","time_taken_min"]), warnings

        df = pd.DataFrame(rows, columns=col_names)
        df["quantity"]       = pd.to_numeric(df["quantity"],       errors="coerce").fillna(0)
        df["time_taken_min"] = pd.to_numeric(df["time_taken_min"], errors="coerce").fillna(0)

        blank_batcher = df["batcher_id"].isna() | (df["batcher_id"].astype(str).str.strip() == "")
        if blank_batcher.sum():
            warnings.append(f"{blank_batcher.sum()} row(s) have blank {batcher_col} — skipped.")
        df = df[~blank_batcher].reset_index(drop=True)

        blank_batch = df["batch_ref"].isna() | (df["batch_ref"].astype(str).str.strip() == "")
        if blank_batch.sum():
            warnings.append(f"{blank_batch.sum()} row(s) have blank batch reference — skipped.")
        df = df[~blank_batch].reset_index(drop=True)

        return df, warnings
    finally:
        conn.close()


def save_btrtp_oracle_data(parsed_rows: list, replace: bool = True) -> int:
    """Save parsed BTRTP rows into btrtp_oracle_data table."""
    now = datetime.now().isoformat(timespec="seconds")
    for r in parsed_rows:
        r["fetched_at"] = now
    if replace:
        return database.replace_table_rows("btrtp_oracle_data", parsed_rows)
    return database.insert_rows("btrtp_oracle_data", parsed_rows)


def save_tp_oracle_data(df: pd.DataFrame, from_date, to_date,
                         parsed_rows: list, replace: bool = True) -> int:
    """
    Save parsed TP rows into tp_oracle_data table.
    parsed_rows is the list of dicts produced by tp_calculator.parse_oracle_df().
    """
    now = datetime.now().isoformat(timespec="seconds")
    for r in parsed_rows:
        r["fetched_at"] = now

    if replace:
        return database.replace_table_rows("tp_oracle_data", parsed_rows)
    return database.insert_rows("tp_oracle_data", parsed_rows)


def fetch_oracle_raw_data(from_date, to_date) -> tuple:
    """
    Single Oracle query that fetches ALL columns needed by every module.
    I&D uses: created_by, production_date, quantity, time_taken_min
    TP  uses: plant_code, batch_ref, production_date, quantity, time_taken_min
    BTRTP uses: all of the above together

    Returns (DataFrame, warnings_list).
    DataFrame columns: production_date, created_by, plant_code, batch_ref,
                       quantity, time_taken_min
    """
    cfg  = get_oracle_config()
    cols = get_tp_oracle_cols()
    batcher_col = (database.get_module_setting("btrtp", "oracle_batcher_col", "CREATED_BY") or "CREATED_BY").strip()
    warnings = []

    _init_thick(cfg["instantclient"])
    conn = oracledb.connect(user=cfg["user"], password=cfg["password"], dsn=_dsn(cfg))
    try:
        cur = conn.cursor()
        fd, td = str(from_date), str(to_date)
        params = {"from_date": fd, "to_date": td}
        status_clause = ""
        if cfg["status_filter"]:
            status_clause = "AND STATUS = :status"
            params["status"] = cfg["status_filter"]

        sql = f"""
            SELECT
                PRODDATE              AS production_date,
                {batcher_col}         AS created_by,
                {cols['plant']}       AS plant_code,
                {cols['batch']}       AS batch_ref,
                PRODUCED_QUANTITY     AS quantity,
                {cols['time']}        AS time_taken_min
            FROM {_TABLE}
            WHERE PRODDATE >= :from_date
              AND PRODDATE <= :to_date
              {status_clause}
            ORDER BY PRODDATE, {batcher_col}
        """
        cur.execute(sql, params)
        rows = cur.fetchall()
        col_names = [d[0].lower() for d in cur.description]

        if not rows:
            warnings.append(f"No rows found in Oracle for {from_date} → {to_date}.")
            return pd.DataFrame(columns=["production_date","created_by","plant_code",
                                         "batch_ref","quantity","time_taken_min"]), warnings

        df = pd.DataFrame(rows, columns=col_names)
        df["quantity"]       = pd.to_numeric(df["quantity"],       errors="coerce").fillna(0)
        df["time_taken_min"] = pd.to_numeric(df["time_taken_min"], errors="coerce").fillna(0)

        # Drop rows with zero time (business rule: no time = don't count)
        zero_time = df["time_taken_min"] == 0
        if zero_time.sum():
            warnings.append(f"{zero_time.sum()} row(s) skipped — time_taken = 0.")
        df = df[~zero_time].reset_index(drop=True)

        # Drop rows with blank created_by
        df["created_by"] = df["created_by"].fillna("").astype(str).str.strip()
        blank_by = df["created_by"] == ""
        if blank_by.sum():
            warnings.append(f"{blank_by.sum()} row(s) skipped — blank {batcher_col}.")
        df = df[~blank_by].reset_index(drop=True)

        # Drop zero/negative quantity
        bad_qty = df["quantity"] <= 0
        if bad_qty.sum():
            warnings.append(f"{bad_qty.sum()} row(s) skipped — zero/negative quantity.")
        df = df[~bad_qty].reset_index(drop=True)

        return df, warnings
    finally:
        conn.close()


def save_oracle_raw_data(df: pd.DataFrame, replace: bool = True) -> int:
    """
    Save a DataFrame from fetch_oracle_raw_data() into the shared oracle_raw_data table.
    replace=True clears existing data for the rows' date range first.
    Returns rows saved.
    """
    if df.empty:
        return 0
    now = datetime.now().isoformat(timespec="seconds")
    rows = []
    for _, row in df.iterrows():
        rows.append({
            "production_date": str(row["production_date"]),
            "created_by":      str(row["created_by"]),
            "plant_code":      str(row.get("plant_code", "") or ""),
            "batch_ref":       str(row.get("batch_ref", "") or ""),
            "quantity":        float(row["quantity"]),
            "time_taken_min":  float(row["time_taken_min"]),
            "fetched_at":      now,
        })
    if replace:
        # Clear rows in the same date range before inserting
        if rows:
            min_d = min(r["production_date"] for r in rows)
            max_d = max(r["production_date"] for r in rows)
            conn = database.get_connection()
            try:
                conn.execute(
                    "DELETE FROM oracle_raw_data WHERE production_date >= ? AND production_date <= ?",
                    (min_d, max_d)
                )
                conn.commit()
            finally:
                conn.close()
    return database.insert_rows("oracle_raw_data", rows)


def save_oracle_backend_data(df: pd.DataFrame, from_date, to_date,
                              replace: bool = True) -> int:
    """
    Save a DataFrame returned by fetch_backend_data() into the backend_data table.

    Args:
        df         – DataFrame with columns: created_by, date, quantity
        from_date  – start of the fetched range (used as source label)
        to_date    – end of the fetched range (used as source label)
        replace    – True = clear existing data first; False = append

    Returns the number of rows saved.
    """
    now = datetime.now().isoformat(timespec="seconds")
    source_label = f"Oracle {from_date} to {to_date}"

    rows = [
        {
            "created_by":  str(row["created_by"]),
            "quantity":    float(row["quantity"]),
            "date":        str(row["date"]),
            "source_file": source_label,
            "uploaded_at": now,
        }
        for _, row in df.iterrows()
    ]

    if replace:
        return database.replace_table_rows("backend_data", rows)
    return database.insert_rows("backend_data", rows)
