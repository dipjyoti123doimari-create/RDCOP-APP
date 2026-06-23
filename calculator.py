"""
calculator.py
=============
Phase 6 — Incentive & Deduction Calculation Engine

Takes the three data sources from SQLite and produces a full per-employee
incentive and deduction result for a chosen month/year (or date range).

BUSINESS RULES (all rates/targets come from config.py):
  Incentive:
    - Employee must have Total Quantity >= 1000 to be eligible.
    - Production Officer category is NEVER eligible for incentive.
    - If the employee's plant YTD Maintenance Cost <= 18  → "low cost" rate.
    - If the employee's plant YTD Maintenance Cost  > 18  → "high cost" rate.
    - SPE rate:   low = Rs 5.0/unit,  high = Rs 2.5/unit
    - Other rate: low = Rs 3.0/unit,  high = Rs 1.5/unit
    - NA category is eligible exactly like the "other" categories: it needs
      Total Quantity >= 1000 and uses the 3.0/1.5 rate (NOT the SPE rate).
      NA has NO deduction (its target is 0).
    - If a plant's YTD Maintenance Cost is missing, treat it as 0 (low-cost
      rate) and add the remark "Maintenance cost missing, treated as 0".

  Deduction:
    - Each category has a monthly target quantity and a per-unit shortfall rate.
    - Deduction = max(0, target - total_quantity) * rate
    - NA has target 0, so deduction is always 0.

  Unmapped employees:
    - Employee codes in Backend Data with no matching Master Data record.
    - Saved to the unmapped_employees table for reporting.

PUBLIC FUNCTIONS:
  run_calculation(month, year, start_date=None, end_date=None) -> dict
  get_available_months()   -> list of (year, month) tuples found in backend_data
  get_last_calculation_info() -> dict
"""

from datetime import datetime

import pandas as pd

import config
import database


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


# ---------------------------------------------------------------------------
# 1. DATA FETCH HELPERS
# ---------------------------------------------------------------------------

def get_available_months() -> list:
    """
    Return a sorted list of (year, month) tuples where production data exists.
    Checks oracle_raw_data first (Oracle-sourced), falls back to backend_data (Excel uploads).
    """
    conn = database.get_connection()
    try:
        # Prefer shared Oracle cache; fall back to manual Excel uploads
        df = pd.read_sql_query(
            """SELECT DISTINCT substr(production_date,1,7) AS ym
               FROM oracle_raw_data
               UNION
               SELECT DISTINCT substr(date,1,7) AS ym
               FROM backend_data
               ORDER BY ym""",
            conn,
        )
    finally:
        conn.close()

    result = []
    for ym in df["ym"].tolist():
        try:
            y, m = ym.split("-")
            result.append((int(y), int(m)))
        except Exception:
            pass
    return result


def _fetch_backend_for_period(month: int, year: int,
                               start_date: str = None,
                               end_date: str = None) -> pd.DataFrame:
    """
    Return production rows for the given period for I&D calculation.

    Priority: if oracle_raw_data has rows for this period, use ONLY that table.
    Fall back to backend_data (Excel uploads) only when oracle has nothing.
    This prevents double-counting when Oracle data was also manually uploaded.
    """
    conn = database.get_connection()
    try:
        if start_date and end_date:
            # Check if Oracle has data for this date range
            oracle_check = pd.read_sql_query(
                "SELECT COUNT(*) AS cnt FROM oracle_raw_data "
                "WHERE production_date >= ? AND production_date <= ?",
                conn, params=(start_date, end_date)
            )
            if oracle_check["cnt"].iloc[0] > 0:
                df = pd.read_sql_query(
                    "SELECT created_by, quantity, production_date AS date "
                    "FROM oracle_raw_data "
                    "WHERE production_date >= ? AND production_date <= ?",
                    conn, params=(start_date, end_date)
                )
            else:
                df = pd.read_sql_query(
                    "SELECT created_by, quantity, date "
                    "FROM backend_data "
                    "WHERE date >= ? AND date <= ?",
                    conn, params=(start_date, end_date)
                )
        else:
            period = f"{year:04d}-{month:02d}"
            oracle_check = pd.read_sql_query(
                "SELECT COUNT(*) AS cnt FROM oracle_raw_data "
                "WHERE substr(production_date,1,7) = ?",
                conn, params=(period,)
            )
            if oracle_check["cnt"].iloc[0] > 0:
                df = pd.read_sql_query(
                    "SELECT created_by, quantity, production_date AS date "
                    "FROM oracle_raw_data "
                    "WHERE substr(production_date,1,7) = ?",
                    conn, params=(period,)
                )
            else:
                df = pd.read_sql_query(
                    "SELECT created_by, quantity, date "
                    "FROM backend_data "
                    "WHERE substr(date,1,7) = ?",
                    conn, params=(period,)
                )
    finally:
        conn.close()
    return df


# ---------------------------------------------------------------------------
# 2. CORE CALCULATION
# ---------------------------------------------------------------------------

_VALID_CATEGORIES = set(config.CATEGORIES)


def _calculate_incentive(total_qty: float, category: str,
                          ytd_cost: float) -> tuple:
    """
    Return (incentive_eligible, incentive_rate, incentive_amount, remarks).
    """
    # Unknown / blank category — flag but do not calculate
    if not category or category.lower() in ("nan", "") or \
            category not in _VALID_CATEGORIES:
        return "No", 0.0, 0.0, f"Unknown category '{category}' — fix in Master Data"

    if category in config.NO_INCENTIVE_CATEGORIES:
        return "No", 0.0, 0.0, "Category not eligible for incentive"

    low_cost = (ytd_cost <= config.MAINTENANCE_COST_THRESHOLD)

    # Every eligible category (including NA) must meet the minimum quantity.
    if total_qty < config.INCENTIVE_MIN_QUANTITY:
        return (
            "No",
            0.0,
            0.0,
            f"Quantity {total_qty:.0f} < minimum {config.INCENTIVE_MIN_QUANTITY}",
        )

    if category == "SPE":
        rate = config.SPE_RATE_LOW_COST if low_cost else config.SPE_RATE_HIGH_COST
    else:
        rate = config.OTHER_RATE_LOW_COST if low_cost else config.OTHER_RATE_HIGH_COST

    amount = round(total_qty * rate, 2)
    cost_label = "low" if low_cost else "high"
    return "Yes", rate, amount, f"{cost_label} maintenance cost rate"


def _calculate_deduction(total_qty: float, category: str) -> tuple:
    """
    Return (deduction_target, shortfall_qty, deduction_amount).
    """
    # Unknown category — no deduction applied
    if not category or category.lower() in ("nan", "") or \
            category not in _VALID_CATEGORIES:
        return 0.0, 0.0, 0.0

    rule = config.DEDUCTION_RULES.get(category, {"target": 0, "rate": 0})
    target = rule["target"]
    rate   = rule["rate"]

    if target == 0:
        return 0.0, 0.0, 0.0

    shortfall = max(0.0, target - total_qty)
    deduction = round(shortfall * rate, 2)
    return float(target), shortfall, deduction


# ---------------------------------------------------------------------------
# 3. MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def run_calculation(month: int, year: int,
                    start_date: str = None,
                    end_date: str = None,
                    persist: bool = True) -> dict:
    """
    Run the full incentive and deduction calculation for the given period.

    persist=True  (default) -> save the results to the database (used by the
                               Calculate page; this is the "official" run).
    persist=False           -> calculate and RETURN the rows only, without
                               touching the database. Used by the View Reports
                               page so browsing a date range never overwrites
                               the saved calculation.

    Steps:
      1. Fetch backend rows for the period → sum quantity per employee.
      2. Separate unmapped employees (not in master data).
      3. For each mapped employee: look up plant maintenance cost,
         calculate incentive and deduction.
      4. (If persist) Save results to calculation_results / unmapped_employees.

    Returns:
        {
            "total_employees":   int,
            "mapped":            int,
            "unmapped":          int,
            "total_incentive":   float,
            "total_deduction":   float,
            "generated_at":      str,
            "results_rows":      list[dict],   # the calculated rows
            "error":             str | None,
        }
    """
    try:
        # ── Step 1: Fetch and aggregate backend data ─────────────────────────
        backend_df = _fetch_backend_for_period(month, year, start_date, end_date)
        if backend_df.empty:
            return {
                "total_employees": 0, "mapped": 0, "unmapped": 0,
                "total_incentive": 0, "total_deduction": 0,
                "generated_at": _now(), "results_rows": [],
                "error": "No backend data found for the selected period.",
            }

        bad_qty_mask = pd.to_numeric(backend_df["quantity"], errors="coerce").isna()
        bad_qty_count = int(bad_qty_mask.sum())
        backend_df["quantity"] = pd.to_numeric(backend_df["quantity"], errors="coerce").fillna(0)
        backend_df["created_by"] = backend_df["created_by"].astype(str).str.strip()

        aggregated = (
            backend_df.groupby("created_by", as_index=False)["quantity"]
            .sum()
            .rename(columns={"created_by": "employee_code", "quantity": "total_quantity"})
        )

        # ── Step 2: Load master & maintenance data ───────────────────────────
        master_df = database.read_table("master_data")
        master_df["employee_code"] = master_df["employee_code"].astype(str).str.strip()

        # Load maintenance cost — try exact month/year first, fall back to whatever is uploaded.
        import calendar as _cal
        calc_month_label = f"{_cal.month_name[month]} {year}"
        conn_m = database.get_connection()
        try:
            import pandas as _pd2
            maint_df = _pd2.read_sql_query(
                "SELECT plant_code, ytd_maintenance_cost, month, year "
                "FROM maintenance_cost WHERE month = ? AND year = ?",
                conn_m, params=(month, year)
            )
            if maint_df.empty:
                # Fall back to all uploaded rows (use whatever month is available)
                maint_df = _pd2.read_sql_query(
                    "SELECT plant_code, ytd_maintenance_cost, month, year "
                    "FROM maintenance_cost",
                    conn_m
                )
        finally:
            conn_m.close()

        # If still no maintenance data at all, proceed with 0 cost for every plant
        if maint_df.empty:
            maint_applied_label = "None (₹0)"
            maint_lookup = {}
        else:
            # Detect which month's data is actually being used and warn if it differs
            _m_used = int(maint_df["month"].dropna().mode().iloc[0]) if not maint_df["month"].dropna().empty else 0
            _y_used = int(maint_df["year"].dropna().mode().iloc[0])  if not maint_df["year"].dropna().empty else 0
            maint_applied_label = (f"{_cal.month_name[_m_used]} {_y_used}"
                                   if _m_used else "Unassigned")
            maint_df["plant_code"] = maint_df["plant_code"].astype(str).str.strip()
            maint_lookup = dict(
                zip(maint_df["plant_code"], maint_df["ytd_maintenance_cost"])
            )

        # ── Step 3: Separate mapped vs unmapped ──────────────────────────────
        known_codes = set(master_df["employee_code"].tolist())
        mapped_agg   = aggregated[aggregated["employee_code"].isin(known_codes)]
        unmapped_agg = aggregated[~aggregated["employee_code"].isin(known_codes)]

        now = _now()

        # ── Step 4: Build unmapped-employee rows (save only when persisting) ──
        unmapped_rows = [
            {
                "employee_code":  row["employee_code"],
                "month":          month,
                "year":           year,
                "total_quantity": float(row["total_quantity"]),
                "remarks":        "Employee code not found in Master Data",
                "generated_at":   now,
            }
            for _, row in unmapped_agg.iterrows()
        ]
        # unmapped rows saved in Step 6 together with calculation_results

        # ── Step 5: Calculate incentive & deduction for mapped employees ─────
        merged = mapped_agg.merge(master_df, on="employee_code", how="left")

        # Load waivers for this month/year — {employee_code: display_text}
        waiver_lookup = database.get_waiver_lookup(month, year)

        results = []
        for _, row in merged.iterrows():
            emp_code  = str(row["employee_code"])
            total_qty = float(row["total_quantity"])
            category  = str(row.get("category", "")).strip()
            plant_code = str(row.get("plant_code", "")).strip()

            # If this plant has no row in the maintenance file, the rule says:
            # treat the cost as 0 (counts as low-cost) AND add a clear remark.
            cost_missing = plant_code not in maint_lookup
            ytd_cost = float(maint_lookup.get(plant_code, 0.0))

            inc_elig, inc_rate, inc_amount, remark = _calculate_incentive(
                total_qty, category, ytd_cost
            )
            if cost_missing:
                miss_note = "Maintenance cost missing, treated as 0"
                remark = f"{remark}; {miss_note}" if remark else miss_note
            ded_target, shortfall, ded_amount = _calculate_deduction(
                total_qty, category
            )
            # Override remark with context-specific message
            if emp_code in waiver_lookup:
                # Waived employee — keep red row (ded_target/shortfall intact for display)
                # but zero the actual deduction amount
                ded_amount = 0.0
                remark = waiver_lookup[emp_code]
            elif ded_amount > 0:
                # Red row — quantity below deduction target
                remark = (f"Quantity {total_qty:.0f} < minimum {int(ded_target)} "
                          f"(deduction target)")
            elif inc_amount == 0 and inc_elig == "No" and ded_amount == 0:
                # White row — not eligible for incentive, no deduction either
                if total_qty < config.INCENTIVE_MIN_QUANTITY:
                    remark = (f"Quantity {total_qty:.0f} < minimum "
                              f"{config.INCENTIVE_MIN_QUANTITY} for incentive")

            results.append({
                "month":                month,
                "year":                 year,
                "employee_code":        emp_code,
                "employee_name":        str(row.get("employee_name", "")),
                "designation":          str(row.get("designation", "")),
                "category":             category,
                "plant":                str(row.get("plant", "")),
                "plant_code":           plant_code,
                "total_quantity":       total_qty,
                "ytd_maintenance_cost": ytd_cost,
                "incentive_eligible":   inc_elig,
                "incentive_rate":       inc_rate,
                "incentive_amount":     inc_amount,
                "deduction_target":     ded_target,
                "shortfall_quantity":   shortfall,
                "deduction_amount":     ded_amount,
                "remarks":              remark,
                "generated_at":         now,
            })

        # ── Step 6: Save results (only when persisting) ──────────────────────
        if persist:
            # Delete only this month/year — keeps other months intact
            conn_d = database.get_connection()
            try:
                conn_d.execute(
                    "DELETE FROM calculation_results WHERE month = ? AND year = ?",
                    (month, year))
                conn_d.execute(
                    "DELETE FROM unmapped_employees WHERE month = ? AND year = ?",
                    (month, year))
                conn_d.commit()
            finally:
                conn_d.close()
            if results:
                database.insert_rows("calculation_results", results)
            if unmapped_rows:
                database.insert_rows("unmapped_employees", unmapped_rows)

        total_incentive = sum(r["incentive_amount"] for r in results)
        total_deduction = sum(r["deduction_amount"] for r in results)

        # ── Data quality warnings ────────────────────────────────────────────
        calc_warnings = []
        if maint_applied_label == "None (₹0)":
            calc_warnings.append(
                f"⚠️ No maintenance cost data found — all plants calculated with ₹0 cost. "
                f"Upload maintenance cost for {calc_month_label} for accurate results."
            )
        elif _m_used != month or _y_used != year:
            calc_warnings.append(
                f"⚠️ Maintenance cost month mismatch: calculation is for {calc_month_label} "
                f"but maintenance cost data applied is from {maint_applied_label}. "
                f"Upload {calc_month_label} maintenance cost for accurate results."
            )
        if bad_qty_count > 0:
            calc_warnings.append(
                f"⚠️ {bad_qty_count} backend row(s) had non-numeric Quantity and were "
                f"counted as 0. Check your Backend Data file for corrupt entries."
            )
        unknown_cat = [r for r in results if r["remarks"]
                       and "Unknown category" in r["remarks"]]
        if unknown_cat:
            names = ", ".join(
                f"{r['employee_code']} ({r['category']})" for r in unknown_cat[:5]
            )
            extra = f" …and {len(unknown_cat) - 5} more" if len(unknown_cat) > 5 else ""
            calc_warnings.append(
                f"⚠️ {len(unknown_cat)} employee(s) have an unrecognised category "
                f"and received ₹0 incentive + ₹0 deduction — fix in Master Data: "
                f"{names}{extra}"
            )

        # Persist last-run info for the dashboard (only on an official run).
        if persist:
            database.set_setting("last_calc_month",     str(month))
            database.set_setting("last_calc_year",       str(year))
            database.set_setting("last_calc_at",         now)
            database.set_setting("last_calc_mapped",     str(len(results)))
            database.set_setting("last_calc_unmapped",   str(len(unmapped_agg)))

        return {
            "total_employees": len(aggregated),
            "mapped":          len(results),
            "unmapped":        len(unmapped_agg),
            "total_incentive": round(total_incentive, 2),
            "total_deduction": round(total_deduction, 2),
            "generated_at":    now,
            "results_rows":    results,
            "unmapped_rows":   unmapped_rows,
            "calc_warnings":   calc_warnings,
            "error":           None,
        }

    except Exception as exc:
        return {
            "total_employees": 0, "mapped": 0, "unmapped": 0,
            "total_incentive": 0, "total_deduction": 0,
            "generated_at": _now(), "results_rows": [],
            "error": str(exc),
        }


# ---------------------------------------------------------------------------
# 4. HELPERS FOR UI
# ---------------------------------------------------------------------------

def get_last_calculation_info() -> dict:
    """Return the last calculation run details stored in app_settings."""
    return {
        "month":    database.get_setting("last_calc_month",   None),
        "year":     database.get_setting("last_calc_year",    None),
        "ran_at":   database.get_setting("last_calc_at",      None),
        "mapped":   database.get_setting("last_calc_mapped",  "0"),
        "unmapped": database.get_setting("last_calc_unmapped","0"),
    }
