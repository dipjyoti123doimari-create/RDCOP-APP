"""
tp_calculator.py
================
All calculation logic for the RDC-TP (Plant Throughput) module.

PUBLIC FUNCTIONS
    parse_oracle_df(df)           -> list[dict]   clean & parse raw Oracle rows
    run_tp_calculation(month, year) -> (plant_rows, location_rows, warnings)
    save_tp_results(plant_rows, month, year) -> int

THROUGHPUT FORMULA (per plant/mixer)
    throughput % = (Total Qty / Total Time in Hours) / Mixer Theo. Capacity × 100

TIME TAKEN RULES
    1. Take abs() — negative values are valid, just remove the sign
    2. Skip rows where abs(time) == 0
    3. Skip rows where abs(time) > 100 (outlier batches)
    Both quantity and time are excluded for skipped rows.

BATCH PARSING
    Format A: AY1/2026/2167        → plant_code=AY1, mixer=None,  lookup=AY1
    Format B: AY1/BP1/2026/2167    → plant_code=AY1, mixer=BP1,   lookup=AY1_BP1

LOCATION TP
    Simple average of throughput % of all plants under the same Exco Location.
    (Weighted average not applicable because mixer capacities differ per plant.)

GRADE-ADJUSTED THROUGHPUT (additive KPI — does not alter the formula above)
    1. weighted_avg_grade = SUM(grade × qty) / SUM(qty), grade parsed from the
       Oracle grade/item description ("M20" → 20). Rows with an unparseable
       grade are excluded from this weighted average but still counted in the
       existing throughput (they carry no grade info, not a grade of 0).
    2. RAG (Rational Average Grade) = REFERENCE_GRADE + (weighted_avg_grade - REFERENCE_GRADE) / 2
       (dampens the grade deviation by half so the adjustment doesn't swing
       throughput_pct too far past 100%; weighted_avg_grade itself is still
       reported unchanged as its own KPI)
    3. grade_adjusted_capacity = mixer_theo_cap × (REFERENCE_GRADE / RAG)
    4. grade_adjusted_throughput_pct = (actual_rate / grade_adjusted_capacity) × 100
       where actual_rate is exactly the same Total Qty / Total Time (hrs) used
       above.
"""

import re
from datetime import datetime

import pandas as pd

import database

# Reference concrete grade the Grade-Adjusted Throughput KPI is normalised
# against (M20). Numeric grade, not a display string.
REFERENCE_GRADE = 20.0

# Matches "M20", "m 30", "M25/M7.5", "M30-MIVAN" etc. — first standalone
# M<number> wins. The (?<![A-Za-z0-9]) lookbehind stops "TM1028/M40" or
# "TM20/M50" (truck/mixer reference codes, e.g. "TM260") from being misread
# as grade M1028 or M20 — the M there is part of "TM", not a grade prefix.
# Confirmed against real Oracle grade_desc values (verified 2026-08-25).
_GRADE_RE = re.compile(r"(?<![A-Za-z0-9])M\s*(\d+(?:\.\d+)?)", re.IGNORECASE)


# ── Batch parsing ─────────────────────────────────────────────────────────────

def _parse_batch(batch_str: str) -> tuple:
    """
    Parse a batch reference string into (plant_code, mixer_variant, lookup_code).

    Returns:
        plant_code    – base plant code (e.g. "AY1")
        mixer_variant – "BP1"/"BP2"/"BP3" or None
        lookup_code   – key used to join with tp_plant_data (e.g. "AY1" or "AY1_BP1")
    """
    parts = str(batch_str).strip().split("/")
    plant_code = parts[0].strip() if parts else ""

    if len(parts) >= 4:
        # Format B: PLANT/BPn/YEAR/BATCHNUM
        second = parts[1].strip().upper()
        if second.startswith("BP"):
            return plant_code, second, f"{plant_code}_{second}"
        # Unusual 4-part without BP — treat like Format A
        return plant_code, None, plant_code

    # Format A: PLANT/YEAR/BATCHNUM (3 parts) or any other
    return plant_code, None, plant_code


def parse_grade(desc) -> float:
    """
    Extract the numeric concrete grade from an item/grade description
    (e.g. "M30-MIVAN" → 30.0, "M25/M7.5" → 25.0, "FGJKL079" → None).

    Returns None when no confident M<number> pattern is found — callers must
    treat that as "unknown grade", never as grade 0.
    """
    if desc is None:
        return None
    m = _GRADE_RE.search(str(desc))
    if not m:
        return None
    try:
        val = float(m.group(1))
    except ValueError:
        return None
    return val if val > 0 else None


# ── Oracle data cleaning & parsing ───────────────────────────────────────────

def parse_oracle_df(raw_df: pd.DataFrame) -> tuple:
    """
    Clean a raw Oracle DataFrame (from oracle_connector.fetch_tp_data) and
    return (parsed_rows, skip_log) ready for tp_oracle_data table.

    Rules applied:
        - batch_ref must be non-blank
        - time_taken_min: take abs(); skip row if == 0 or > 100
        - quantity: skip row if <= 0
    """
    rows = []
    skip_log = []   # list of dicts for tp_validation_errors style log

    for _, r in raw_df.iterrows():
        batch_ref = str(r.get("batch_ref", "")).strip()
        production_date = str(r.get("production_date", "")).strip()
        plant_code_raw  = str(r.get("plant_code", "")).strip()

        if not batch_ref or batch_ref == "nan":
            skip_log.append({"reason": "Blank batch reference", "batch_ref": batch_ref,
                             "plant_code": plant_code_raw, "production_date": production_date,
                             "quantity": "", "time_taken_min": ""})
            continue

        raw_time = r.get("time_taken_min", 0)
        try:
            time_min = abs(float(raw_time))
        except (TypeError, ValueError):
            time_min = 0.0

        qty = float(r.get("quantity", 0) or 0)

        if time_min == 0:
            skip_log.append({"reason": "Time taken = 0", "batch_ref": batch_ref,
                             "plant_code": plant_code_raw, "production_date": production_date,
                             "quantity": round(qty, 1), "time_taken_min": round(time_min, 1)})
            continue
        if time_min > 100:
            skip_log.append({"reason": f"Time taken > 100 min ({time_min:.1f})", "batch_ref": batch_ref,
                             "plant_code": plant_code_raw, "production_date": production_date,
                             "quantity": round(qty, 1), "time_taken_min": round(time_min, 1)})
            continue
        if qty <= 0:
            skip_log.append({"reason": f"Quantity ≤ 0 ({qty})", "batch_ref": batch_ref,
                             "plant_code": plant_code_raw, "production_date": production_date,
                             "quantity": round(qty, 1), "time_taken_min": round(time_min, 1)})
            continue

        plant_code, mixer_variant, lookup_code = _parse_batch(batch_ref)

        rows.append({
            "production_date": str(r.get("production_date", "")),
            "plant_code":      plant_code,
            "mixer_variant":   mixer_variant or "",
            "lookup_code":     lookup_code,
            "batch_ref":       batch_ref,
            "quantity":        qty,
            "time_taken_min":  time_min,
            # Grade-Adjusted Throughput only — not used by the existing
            # throughput calculation. None when ungraded/unparseable.
            "grade":           parse_grade(r.get("grade_desc")),
            "grade_desc":      "" if r.get("grade_desc") is None else str(r.get("grade_desc")),
        })

    return rows, skip_log


# ── Main calculation ──────────────────────────────────────────────────────────

def run_tp_calculation(month: int, year: int,
                        from_date: str = None, to_date: str = None,
                        ora_df=None) -> tuple:
    """
    Calculate plant-wise and location-wise throughput.

    ora_df – optional in-memory DataFrame (already parsed). When supplied the
             DB is not read at all, enabling historical queries without
             persisting data. When None, data is read from tp_oracle_data.

    Returns:
        plant_rows    – list[dict]  one row per plant/mixer
        location_rows – list[dict]  one row per Exco Location (simple avg)
        warnings      – list[str]
    """
    warnings = []

    # 1. Load oracle data (from memory or shared DB cache)
    if ora_df is None:
        # Read from shared oracle_raw_data; fall back to legacy tp_oracle_data
        import sqlite3 as _sq
        conn = database.get_connection()
        try:
            fd_str = str(from_date) if from_date else "0000-00-00"
            td_str = str(to_date)   if to_date   else "9999-12-31"
            ora_df = pd.read_sql_query(
                """SELECT production_date, plant_code AS plant_col,
                          batch_ref, quantity, time_taken_min
                   FROM oracle_raw_data
                   WHERE production_date >= ? AND production_date <= ?""",
                conn, params=(fd_str, td_str)
            )
            if ora_df.empty:
                # fall back to legacy table
                ora_df = pd.read_sql_query(
                    """SELECT production_date, plant_col, batch_ref,
                              quantity, time_taken_min
                       FROM tp_oracle_data
                       WHERE production_date >= ? AND production_date <= ?""",
                    conn, params=(fd_str, td_str)
                )
        finally:
            conn.close()
        if ora_df.empty:
            warnings.append("No Oracle data found — fetch from Oracle first.")
            return [], [], warnings
    else:
        if ora_df.empty:
            warnings.append(f"No Oracle data in range {from_date} → {to_date}.")
            return [], [], warnings

    # 2. Load plant reference data
    plant_df = database.read_table("tp_plant_data")
    if plant_df.empty:
        warnings.append("No Plant Data found — sync from Google Sheets first.")
        return [], [], warnings

    plant_map = {
        str(r["plant_code"]): r.to_dict()
        for _, r in plant_df.iterrows()
    }

    # 3. Derive lookup_code from batch_ref if not present (shared cache doesn't have it)
    ora_df["quantity"]       = pd.to_numeric(ora_df["quantity"],       errors="coerce").fillna(0)
    ora_df["time_taken_min"] = pd.to_numeric(ora_df["time_taken_min"], errors="coerce").fillna(0)

    if "lookup_code" not in ora_df.columns:
        ora_df["lookup_code"] = ora_df["batch_ref"].apply(
            lambda b: _parse_batch(str(b))[2]
        )

    # Grade-Adjusted Throughput: "grade" is only present when the caller passed
    # an in-memory df from parse_oracle_df() (the live Reports/Calculate path).
    # The persisted oracle_raw_data / tp_oracle_data cache tables carry no
    # grade column, so grade-adjusted figures are simply unavailable there —
    # never treated as grade 0.
    has_grade = "grade" in ora_df.columns
    if has_grade:
        ora_df["grade"] = pd.to_numeric(ora_df["grade"], errors="coerce")
        ora_df["graded_qty"] = ora_df["quantity"].where(ora_df["grade"].notna(), 0.0)
        ora_df["grade_x_qty"] = (ora_df["grade"] * ora_df["quantity"]).where(ora_df["grade"].notna(), 0.0)

    # Drop zero-time and zero-qty rows (already filtered in fetch but guard here too)
    ora_df = ora_df[(ora_df["time_taken_min"] > 0) & (ora_df["time_taken_min"] <= 100)
                    & (ora_df["quantity"] > 0)]

    agg_kwargs = dict(
        total_quantity=("quantity",      "sum"),
        total_time_min=("time_taken_min","sum"),
        batch_count=("quantity",         "count"),
    )
    if has_grade:
        agg_kwargs["graded_quantity"] = ("graded_qty",   "sum")
        agg_kwargs["grade_x_qty_sum"] = ("grade_x_qty",  "sum")

    grouped = (
        ora_df.groupby("lookup_code", sort=True)
        .agg(**agg_kwargs)
        .reset_index()
    )

    now = datetime.now().isoformat(timespec="seconds")
    plant_rows = []

    for _, row in grouped.iterrows():
        lookup = str(row["lookup_code"])
        info   = plant_map.get(lookup)

        if not info:
            warnings.append(f"Plant code '{lookup}' not found in Plant Data — skipped.")
            continue

        total_time_hrs = row["total_time_min"] / 60.0
        if total_time_hrs <= 0:
            warnings.append(f"'{lookup}' has zero total time after filtering — skipped.")
            continue

        mixer_cap = float(info.get("mixer_theo_cap") or 0)
        actualProductionRate = row["total_quantity"] / total_time_hrs   # units/hr
        if mixer_cap <= 0:
            warnings.append(f"'{lookup}' has zero Mixer Theo. Capacity — throughput set to 0%.")
            throughput_pct = 0.0
        else:
            avg_rate       = actualProductionRate
            throughput_pct = (avg_rate / mixer_cap) * 100.0
        existingThroughput = throughput_pct

        # ── Grade-Adjusted Throughput (additive; does not affect the above) ──
        weightedAverageGrade = None
        gradeAdjustedCapacity = None
        gradeAdjustedThroughput = None
        if has_grade:
            graded_qty = float(row.get("graded_quantity") or 0)
            if graded_qty > 0:
                weightedAverageGrade = float(row["grade_x_qty_sum"]) / graded_qty
                # RAG = Rational Average Grade — dampens the raw weighted-average
                # grade's deviation from REFERENCE_GRADE by half before it's used
                # to scale capacity, so grade_adjusted_throughput_pct stops
                # swinging past 100%.
                rationalAverageGrade = REFERENCE_GRADE + (weightedAverageGrade - REFERENCE_GRADE) / 2
                if mixer_cap > 0 and rationalAverageGrade > 0:
                    gradeAdjustedCapacity = mixer_cap * (REFERENCE_GRADE / rationalAverageGrade)
                    if gradeAdjustedCapacity > 0:
                        gradeAdjustedThroughput = (actualProductionRate / gradeAdjustedCapacity) * 100.0

        plant_rows.append({
            "month":          month,
            "year":           year,
            "lookup_code":    lookup,
            "plant_name":     str(info.get("plant_name", "")),
            "exco_location":  str(info.get("exco_location", "")),
            "business_head":  str(info.get("business_head", "")),
            "plant_manager":  str(info.get("plant_manager", "")),
            "mixer_theo_cap": mixer_cap,
            "total_quantity": round(float(row["total_quantity"]), 2),
            "total_time_min": round(float(row["total_time_min"]), 1),
            "total_time_hrs": round(total_time_hrs, 2),
            "throughput_pct": round(throughput_pct, 2),
            "batch_count":    int(row["batch_count"]),
            "generated_at":   now,
            # Grade-Adjusted Throughput KPI — None when no graded rows exist
            # for this plant (never displayed as 0%, see templates).
            "weighted_avg_grade":          round(weightedAverageGrade, 2) if weightedAverageGrade is not None else None,
            "rational_average_grade":      round(rationalAverageGrade, 2) if weightedAverageGrade is not None else None,
            "grade_adjusted_capacity":     round(gradeAdjustedCapacity, 2) if gradeAdjustedCapacity is not None else None,
            "grade_adjusted_throughput_pct": round(gradeAdjustedThroughput, 2) if gradeAdjustedThroughput is not None else None,
            "graded_quantity":             round(float(row.get("graded_quantity") or 0), 2) if has_grade else None,
        })

    # 4. Sort plant rows lowest → highest and build location summary
    plant_rows.sort(key=lambda r: r["throughput_pct"])
    location_rows = build_location_rows(plant_rows, month, year)

    return plant_rows, location_rows, warnings


def build_location_rows(plant_rows: list, month: int, year: int) -> list:
    """
    Build Exco-Location rows (simple average of plant throughput %) from a list
    of plant rows, sorted lowest → highest with a PAN India total at the bottom.
    Reused by both the calculation and the filtered Reports view.
    """
    location_map = {}
    for pr in plant_rows:
        loc = pr.get("exco_location", "")
        if not loc:
            continue
        if loc not in location_map:
            location_map[loc] = {"plant_count": 0, "total_throughput_pct": 0.0,
                                 "total_quantity": 0.0, "total_time_min": 0.0,
                                 "grade_x_qty_sum": 0.0, "graded_quantity": 0.0,
                                 "gat_pct_sum": 0.0, "gat_pct_count": 0}
        location_map[loc]["plant_count"]          += 1
        location_map[loc]["total_throughput_pct"] += pr["throughput_pct"]
        location_map[loc]["total_quantity"]        += pr["total_quantity"]
        location_map[loc]["total_time_min"]        += pr.get("total_time_min", 0.0)
        # Grade-Adjusted Throughput: recompute the location's weighted grade
        # from underlying qty×grade sums — never average plant weighted grades.
        gqty = pr.get("graded_quantity")
        wag  = pr.get("weighted_avg_grade")
        if gqty and wag is not None:
            location_map[loc]["grade_x_qty_sum"] += wag * gqty
            location_map[loc]["graded_quantity"]  += gqty
        gat = pr.get("grade_adjusted_throughput_pct")
        if gat is not None:
            location_map[loc]["gat_pct_sum"]   += gat
            location_map[loc]["gat_pct_count"] += 1

    location_rows = []
    all_pct_sum = all_qty_sum = all_time_sum = 0.0
    all_grade_x_qty = all_graded_qty = 0.0
    all_gat_sum = 0.0
    all_gat_count = 0
    all_count = 0
    for loc, d in sorted(location_map.items()):
        cnt = d["plant_count"]
        loc_weighted_grade = (d["grade_x_qty_sum"] / d["graded_quantity"]) if d["graded_quantity"] > 0 else None
        loc_rag = (REFERENCE_GRADE + (loc_weighted_grade - REFERENCE_GRADE) / 2) if loc_weighted_grade is not None else None
        loc_avg_gat = (d["gat_pct_sum"] / d["gat_pct_count"]) if d["gat_pct_count"] > 0 else None
        location_rows.append({
            "exco_location":      loc,
            "plant_count":        cnt,
            "avg_throughput_pct": round(d["total_throughput_pct"] / cnt, 2) if cnt else 0.0,
            "total_quantity":     round(d["total_quantity"], 2),
            "total_time_min":     round(d["total_time_min"], 1),
            "month": month, "year": year, "is_pan_india": False,
            "weighted_avg_grade":               round(loc_weighted_grade, 2) if loc_weighted_grade is not None else None,
            "rational_average_grade":           round(loc_rag, 2) if loc_rag is not None else None,
            "avg_grade_adjusted_throughput_pct": round(loc_avg_gat, 2) if loc_avg_gat is not None else None,
        })
        all_pct_sum     += d["total_throughput_pct"]
        all_qty_sum     += d["total_quantity"]
        all_time_sum    += d["total_time_min"]
        all_grade_x_qty += d["grade_x_qty_sum"]
        all_graded_qty  += d["graded_quantity"]
        all_gat_sum     += d["gat_pct_sum"]
        all_gat_count   += d["gat_pct_count"]
        all_count       += cnt

    location_rows.sort(key=lambda r: r["avg_throughput_pct"])

    if all_count:
        pan_weighted_grade = (all_grade_x_qty / all_graded_qty) if all_graded_qty > 0 else None
        pan_rag = (REFERENCE_GRADE + (pan_weighted_grade - REFERENCE_GRADE) / 2) if pan_weighted_grade is not None else None
        pan_avg_gat = (all_gat_sum / all_gat_count) if all_gat_count > 0 else None
        location_rows.append({
            "exco_location":      "PAN India",
            "plant_count":        all_count,
            "avg_throughput_pct": round(all_pct_sum / all_count, 2),
            "total_quantity":     round(all_qty_sum, 2),
            "total_time_min":     round(all_time_sum, 1),
            "month": month, "year": year, "is_pan_india": True,
            "weighted_avg_grade":               round(pan_weighted_grade, 2) if pan_weighted_grade is not None else None,
            "rational_average_grade":           round(pan_rag, 2) if pan_rag is not None else None,
            "avg_grade_adjusted_throughput_pct": round(pan_avg_gat, 2) if pan_avg_gat is not None else None,
        })

    return location_rows


# ── Persist results ───────────────────────────────────────────────────────────

def save_tp_results(plant_rows: list, month: int, year: int) -> int:
    """Replace stored results for this month/year with new plant_rows."""
    conn = database.get_connection()
    try:
        conn.execute(
            "DELETE FROM tp_results WHERE month = ? AND year = ?", (month, year)
        )
        conn.commit()
    finally:
        conn.close()

    # Only persist columns that exist in the tp_results table (total_time_min is
    # a display-only field kept in the in-memory rows, not stored).
    keep = {"month", "year", "lookup_code", "plant_name", "exco_location",
            "business_head", "plant_manager", "mixer_theo_cap", "total_quantity",
            "total_time_hrs", "throughput_pct", "batch_count", "generated_at"}
    db_rows = [{k: v for k, v in r.items() if k in keep} for r in plant_rows]
    return database.insert_rows("tp_results", db_rows)
