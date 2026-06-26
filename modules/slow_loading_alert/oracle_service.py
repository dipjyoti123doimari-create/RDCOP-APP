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
# Master header holds the friendly customer NAME and grade NAME, joined by
# sales order + line number (verified live):
#   T002=plant  T003=line no  T004=sales order  T006=grade name  T009=customer name
_MASTER_TABLE = "APPSREAD.rdc_batch_master_header"


def _sla_oracle_cols() -> dict:
    def _ms(k, d):
        return (database.get_module_setting("sla", k, d) or d).strip()
    # Defaults match the REAL rdc_batch_trx_headers columns (verified against the
    # live Oracle data dictionary):
    #   PLANTNO, SALESORDER, LINENUMBER, ITEMNAME, CREATED_BY, TRUCK_CODE,
    #   PRODUCED_QUANTITY, TIMETAKEN
    # Customer NAME and friendly grade NAME are NOT in the trx table; they are
    # joined from rdc_batch_master_header (T009 customer, T006 grade) on
    # SALESORDER=T004 and LINENUMBER=T003. See _build_query().
    #   mixer_cap → NOT in Oracle — comes from the Google Sheet plant mapping.
    # An EMPTY override means "this column is not available" → selected as NULL.
    return {
        "plant":         _ms("oracle_plant_col",      "PLANTNO"),
        "salesorder":    _ms("oracle_salesorder_col", "SALESORDER"),
        "linenumber":    _ms("oracle_linenumber_col", "LINENUMBER"),
        "batcher":       _ms("oracle_batcher_col",     "CREATED_BY"),
        "truck":         _ms("oracle_truck_col",       "TRUCK_CODE"),
        "quantity":      _ms("oracle_quantity_col",    "PRODUCED_QUANTITY"),
        "time":          _ms("oracle_time_col",        "TIMETAKEN"),
        "mixer_cap":     _ms("oracle_mixer_cap_col",   ""),  # from Google Sheet
        # Master-header columns for the customer/grade lookup (override if the
        # T-column positions ever differ).
        "cust_name_col": _ms("oracle_cust_name_col",   "T009"),
        "grade_name_col":_ms("oracle_grade_name_col",  "T006"),
        "master_so_col": _ms("oracle_master_so_col",   "T004"),
        "master_ln_col": _ms("oracle_master_ln_col",   "T003"),
        # Fallbacks used when the master row is missing (LEFT JOIN gives NULL).
        "grade_fallback":_ms("oracle_grade_col",       "ITEMNAME"),
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

    # Customer NAME (m.T009) and grade NAME (m.T006) are LEFT-JOINed from the
    # master header on sales order + line number. When a master row is missing,
    # customer falls back to the sales-order number and grade to ITEMNAME, so the
    # query never breaks and a value is always shown.
    cust_expr  = (f"NVL(TO_CHAR(m.{cols['cust_name_col']}), TO_CHAR(h.{cols['salesorder']}))"
                  if cols["cust_name_col"] else f"TO_CHAR(h.{cols['salesorder']})")
    grade_expr = (f"NVL(TO_CHAR(m.{cols['grade_name_col']}), h.{cols['grade_fallback']})"
                  if cols["grade_name_col"] else f"h.{cols['grade_fallback']}")

    def _hsel(expr_col, alias):
        return f"h.{expr_col} AS {alias}" if expr_col else f"NULL AS {alias}"

    select_list = ",\n                ".join([
        "h.PRODDATE AS production_date",
        _hsel(plant_col,        "plant_code"),
        f"{cust_expr} AS customer",
        f"{grade_expr} AS grade",
        _hsel(cols["batcher"],  "batcher_code"),
        _hsel(cols["truck"],    "tm_number"),
        _hsel(cols["quantity"], "batched_quantity"),
        _hsel(time_col,         "loading_time_minutes"),
        "NULL AS mixer_capacity",  # filled from Google Sheet in mapping_service
    ])

    # Build the LEFT JOIN only when master-header lookup columns are configured.
    join_clause = ""
    if cols["cust_name_col"] or cols["grade_name_col"]:
        join_clause = (
            f"LEFT JOIN {_MASTER_TABLE} m "
            f"ON TO_CHAR(h.{cols['salesorder']}) = TO_CHAR(m.{cols['master_so_col']}) "
            f"AND TO_CHAR(h.{cols['linenumber']}) = TO_CHAR(m.{cols['master_ln_col']})"
        )

    oracle_connector._init_thick(cfg["instantclient"])
    import oracledb
    conn = oracledb.connect(user=cfg["user"], password=cfg["password"],
                            dsn=oracle_connector._dsn(cfg))
    try:
        cur = conn.cursor()
        params = {"from_date": str(from_date), "to_date": str(to_date)}
        status_clause = ""
        if cfg["status_filter"]:
            status_clause = "AND h.STATUS = :status"
            params["status"] = cfg["status_filter"]

        sql = f"""
            SELECT
                {select_list}
            FROM {_TABLE} h
            {join_clause}
            WHERE h.PRODDATE >= :from_date
              AND h.PRODDATE <= :to_date
              {status_clause}
            ORDER BY h.PRODDATE, h.{plant_col}
        """
        try:
            cur.execute(sql, params)
        except Exception as exc:
            # If the master-header join fails (e.g. column mismatch), fall back to
            # a join-free query so alerts still go out with sales order + item.
            warnings.append(f"Customer/grade join failed, using sales order + item instead. ({exc})")
            fb_select = ",\n                ".join([
                "PRODDATE AS production_date",
                _hsel(plant_col, "plant_code").replace("h.", ""),
                f"TO_CHAR({cols['salesorder']}) AS customer",
                f"{cols['grade_fallback']} AS grade",
                _hsel(cols["batcher"],  "batcher_code").replace("h.", ""),
                _hsel(cols["truck"],    "tm_number").replace("h.", ""),
                _hsel(cols["quantity"], "batched_quantity").replace("h.", ""),
                _hsel(time_col,         "loading_time_minutes").replace("h.", ""),
                "NULL AS mixer_capacity",
            ])
            sql_fb = f"""
                SELECT
                    {fb_select}
                FROM {_TABLE}
                WHERE PRODDATE >= :from_date AND PRODDATE <= :to_date
                  {status_clause.replace('h.STATUS', 'STATUS')}
                ORDER BY PRODDATE, {plant_col}
            """
            cur.execute(sql_fb, params)

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
