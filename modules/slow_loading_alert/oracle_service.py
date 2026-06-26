"""
oracle_service.py
=================
Fetches slow-loading source data from Oracle via the existing shared
oracle_raw_data cache or a direct Oracle query.

Required Oracle columns (mapped from rdc_batch_trx_headers):
  Plant          → plant_code    (PLANTNO by default)
  Customer       → customer      (CUSTOMERNAME or similar)
  Item Desc      → grade         (ITEMDESCRIPTION or similar)
  Created by     → batcher_code  (CREATED_BY)
  Truck number   → tm_number     (TRUCKNUMBER or similar)
  Quantity       → batched_quantity (PRODUCED_QUANTITY)
  Time taken     → loading_time  (TIMETAKEN — abs value enforced)

Column name overrides are stored in module_settings with prefix "sla.":
  sla.oracle_customer_col    default: CUSTOMERNAME
  sla.oracle_grade_col       default: ITEMDESCRIPTION
  sla.oracle_truck_col       default: TRUCKNUMBER
  sla.oracle_mixer_cap_col   default: MIXERCAPACITY

If those columns do not exist in the Oracle table the query will fail —
the admin must override them in SLA Settings.
"""

import pandas as pd
from datetime import date as _date

import database
import oracle_connector


_TABLE = "APPSREAD.rdc_batch_trx_headers"


def _sla_oracle_cols() -> dict:
    def _ms(k, d):
        return (database.get_module_setting("sla", k, d) or d).strip()
    # Defaults match the REAL rdc_batch_trx_headers columns (verified against the
    # live Oracle data dictionary):
    #   PLANTNO, SALESORDER, ITEMNAME, CREATED_BY, TRUCK_CODE,
    #   PRODUCED_QUANTITY, TIMETAKEN
    # Notes:
    #   customer  → the trx table has no customer-NAME column, only SALESORDER.
    #               The customer name shown in reports is resolved separately.
    #               Default to SALESORDER; admin can override in SLA Settings if
    #               a customer-name column exists in another joined source.
    #   grade     → ITEMNAME (FG code). The friendly grade label (e.g. M25) is a
    #               lookup; admin can override the column if needed.
    #   mixer_cap → NOT in Oracle — comes from the Google Sheet plant mapping.
    #               Always selected as NULL here and filled in mapping_service.
    # An EMPTY override means "this column is not available" → selected as NULL.
    return {
        "plant":      _ms("oracle_plant_col",      "PLANTNO"),
        "customer":   _ms("oracle_customer_col",   "SALESORDER"),
        "grade":      _ms("oracle_grade_col",       "ITEMNAME"),
        "batcher":    _ms("oracle_batcher_col",     "CREATED_BY"),
        "truck":      _ms("oracle_truck_col",       "TRUCK_CODE"),
        "quantity":   _ms("oracle_quantity_col",    "PRODUCED_QUANTITY"),
        "time":       _ms("oracle_time_col",        "TIMETAKEN"),
        "mixer_cap":  _ms("oracle_mixer_cap_col",   ""),  # from Google Sheet
    }


def _plant_mixer_caps() -> dict:
    """
    Mixer capacity per plant code, sourced from the existing TP plant master.
    Returns {plant_code: mixer_theo_cap}. Empty dict if table missing/empty.
    """
    caps = {}
    try:
        conn = database.get_connection()
        try:
            cur = conn.execute(
                "SELECT plant_code, mixer_theo_cap FROM tp_plant_data "
                "WHERE mixer_theo_cap IS NOT NULL"
            )
            for code, cap in cur.fetchall():
                if code is not None and cap is not None:
                    caps[str(code).strip()] = cap
        finally:
            conn.close()
    except Exception:
        pass
    return caps


def fetch_loading_data(from_date=None, to_date=None) -> tuple:
    """
    Fetch today's loading rows from Oracle directly.

    Returns (DataFrame, warnings_list).

    DataFrame columns:
        production_date, plant_code, customer, grade, batcher_code,
        tm_number, batched_quantity, loading_time_minutes, mixer_capacity
    """
    cfg = oracle_connector.get_oracle_config()
    cols = _sla_oracle_cols()
    # Reuse shared TP column overrides for plant and time as a fallback
    tp_cols = oracle_connector.get_tp_oracle_cols()
    plant_col  = cols["plant"]   or tp_cols["plant"]
    time_col   = cols["time"]    or tp_cols["time"]

    warnings = []

    if from_date is None:
        from_date = str(_date.today())
    if to_date is None:
        to_date = str(_date.today())

    # Build the SELECT list. Any column whose override is EMPTY is treated as
    # "not available in Oracle" and selected as a literal NULL so the query
    # never references a non-existent column. Mixer capacity always comes from
    # the Google Sheet mapping, so it is always NULL here.
    def _sel(expr_col, alias):
        return f"{expr_col} AS {alias}" if expr_col else f"NULL AS {alias}"

    select_list = ",\n                ".join([
        "PRODDATE AS production_date",
        _sel(plant_col,         "plant_code"),
        _sel(cols["customer"],  "customer"),
        _sel(cols["grade"],     "grade"),
        _sel(cols["batcher"],   "batcher_code"),
        _sel(cols["truck"],     "tm_number"),
        _sel(cols["quantity"],  "batched_quantity"),
        _sel(time_col,          "loading_time_minutes"),
        "NULL AS mixer_capacity",  # filled from Google Sheet in mapping_service
    ])

    oracle_connector._init_thick(cfg["instantclient"])
    import oracledb
    conn = oracledb.connect(user=cfg["user"], password=cfg["password"],
                            dsn=oracle_connector._dsn(cfg))
    try:
        cur = conn.cursor()
        params = {"from_date": str(from_date), "to_date": str(to_date)}
        status_clause = ""
        if cfg["status_filter"]:
            status_clause = "AND STATUS = :status"
            params["status"] = cfg["status_filter"]

        sql = f"""
            SELECT
                {select_list}
            FROM {_TABLE}
            WHERE PRODDATE >= :from_date
              AND PRODDATE <= :to_date
              {status_clause}
            ORDER BY PRODDATE, {plant_col}
        """
        cur.execute(sql, params)

        rows = cur.fetchall()
        col_names = [d[0].lower() for d in cur.description]

        if not rows:
            warnings.append(f"No loading rows found in Oracle for {from_date} → {to_date}.")
            return _empty_df(), warnings

        df = pd.DataFrame(rows, columns=col_names)
        df = _clean_df(df)
        return df, warnings

    finally:
        conn.close()


def _empty_df() -> pd.DataFrame:
    return pd.DataFrame(columns=[
        "production_date", "plant_code", "customer", "grade",
        "batcher_code", "tm_number", "batched_quantity",
        "loading_time_minutes", "mixer_capacity"
    ])


def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """Standardise types and apply the negative-time rule."""
    df["batched_quantity"]    = pd.to_numeric(df["batched_quantity"],    errors="coerce").fillna(0)
    df["loading_time_minutes"] = pd.to_numeric(df["loading_time_minutes"], errors="coerce").fillna(0)
    df["mixer_capacity"]      = pd.to_numeric(df["mixer_capacity"],      errors="coerce")

    # Negative time → take absolute value
    df["loading_time_minutes"] = df["loading_time_minutes"].abs()

    # Drop rows with zero or blank key fields
    for col in ("plant_code", "batcher_code", "tm_number"):
        df[col] = df[col].fillna("").astype(str).str.strip()
    df = df[df["plant_code"] != ""].reset_index(drop=True)
    df = df[df["tm_number"]  != ""].reset_index(drop=True)
    df = df[df["batched_quantity"] > 0].reset_index(drop=True)
    df = df[df["loading_time_minutes"] > 0].reset_index(drop=True)

    # Coerce string columns
    for col in ("customer", "grade", "production_date"):
        df[col] = df[col].fillna("").astype(str).str.strip()

    return df
