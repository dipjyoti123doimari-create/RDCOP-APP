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
    return {
        "plant":      _ms("oracle_plant_col",      "PLANTNO"),
        "customer":   _ms("oracle_customer_col",   "CUSTOMERNAME"),
        "grade":      _ms("oracle_grade_col",       "ITEMDESCRIPTION"),
        "batcher":    _ms("oracle_batcher_col",     "CREATED_BY"),
        "truck":      _ms("oracle_truck_col",       "TRUCKNUMBER"),
        "quantity":   _ms("oracle_quantity_col",    "PRODUCED_QUANTITY"),
        "time":       _ms("oracle_time_col",        "TIMETAKEN"),
        "mixer_cap":  _ms("oracle_mixer_cap_col",   "MIXERCAPACITY"),
    }


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
    # Reuse shared TP column overrides for plant and time
    tp_cols = oracle_connector.get_tp_oracle_cols()
    plant_col  = cols["plant"]   or tp_cols["plant"]
    time_col   = cols["time"]    or tp_cols["time"]

    warnings = []

    if from_date is None:
        from_date = str(_date.today())
    if to_date is None:
        to_date = str(_date.today())

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

        # MIXER_CAP is optional — wrap in NVL so missing column doesn't crash
        mixer_cap_select = f"{cols['mixer_cap']} AS mixer_capacity" if cols["mixer_cap"] else "NULL AS mixer_capacity"

        sql = f"""
            SELECT
                PRODDATE              AS production_date,
                {plant_col}           AS plant_code,
                {cols['customer']}    AS customer,
                {cols['grade']}       AS grade,
                {cols['batcher']}     AS batcher_code,
                {cols['truck']}       AS tm_number,
                {cols['quantity']}    AS batched_quantity,
                {time_col}            AS loading_time_minutes,
                {mixer_cap_select}
            FROM {_TABLE}
            WHERE PRODDATE >= :from_date
              AND PRODDATE <= :to_date
              {status_clause}
            ORDER BY PRODDATE, {plant_col}
        """
        try:
            cur.execute(sql, params)
        except Exception as exc:
            # Mixer capacity column may not exist — retry without it
            warnings.append(f"Mixer capacity column '{cols['mixer_cap']}' not found in Oracle — defaulting to NULL. ({exc})")
            sql_no_mixer = f"""
                SELECT
                    PRODDATE              AS production_date,
                    {plant_col}           AS plant_code,
                    {cols['customer']}    AS customer,
                    {cols['grade']}       AS grade,
                    {cols['batcher']}     AS batcher_code,
                    {cols['truck']}       AS tm_number,
                    {cols['quantity']}    AS batched_quantity,
                    {time_col}            AS loading_time_minutes,
                    NULL                  AS mixer_capacity
                FROM {_TABLE}
                WHERE PRODDATE >= :from_date
                  AND PRODDATE <= :to_date
                  {status_clause}
                ORDER BY PRODDATE, {plant_col}
            """
            cur.execute(sql_no_mixer, params)

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
