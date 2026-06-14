"""
validations.py
==============
Phase 5 — Data Validation

Checks the three data sources against business rules and stores any
problems found in the `validation_errors` SQLite table so they can be
reviewed and fixed before the incentive calculation is run.

WHAT IS VALIDATED:
  1. Master Data
       - Required fields (Employee Code, Name, Designation, Category,
         Plant, Plant Code) must not be blank.
       - Category must be one of the allowed values in config.CATEGORIES.
       - Employee Code must be unique (no duplicates).

  2. Backend Data
       - Employee Code (created_by) must exist in Master Data.
       - Quantity must be > 0.
       - Date must be a valid date (YYYY-MM-DD).

  3. Maintenance Cost
       - Plant Code must exist in at least one Master Data record.
       - YTD Maintenance Cost must be >= 0.

PUBLIC FUNCTIONS:
  run_all_validations()  -> dict with counts and list of errors
  clear_validation_errors()
  get_validation_summary() -> DataFrame of error counts by source/type
  get_last_validation_info() -> dict
"""

from datetime import datetime

import pandas as pd

import config
import database


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def clear_validation_errors():
    """Remove all rows from the validation_errors table."""
    database.clear_table("validation_errors")


# ---------------------------------------------------------------------------
# 1. MASTER DATA VALIDATION
# ---------------------------------------------------------------------------

def _validate_master_data(errors: list) -> int:
    """
    Validate all rows in the master_data table.
    Appends error dicts to `errors`. Returns count of problems found.
    """
    df = database.read_table("master_data")
    if df.empty:
        errors.append({
            "source": "master_data",
            "row_number": 0,
            "column_name": "—",
            "error_type": "NO_DATA",
            "error_message": "Master Data table is empty. Sync from Google Sheets first.",
            "created_at": _now(),
        })
        return 1

    count = 0
    valid_cats = {c.lower() for c in config.CATEGORIES}
    required_fields = ["employee_code", "employee_name", "designation",
                       "category", "plant", "plant_code"]

    for i, row in df.iterrows():
        row_num = i + 1

        for field in required_fields:
            val = str(row.get(field, "")).strip()
            if not val or val.lower() == "nan":
                errors.append({
                    "source": "master_data",
                    "row_number": row_num,
                    "column_name": field,
                    "error_type": "BLANK_FIELD",
                    "error_message": f"Row {row_num}: '{field}' is blank "
                                     f"(Employee Code: {row.get('employee_code', '?')}).",
                    "created_at": _now(),
                })
                count += 1

        cat = str(row.get("category", "")).strip()
        if cat and cat.lower() != "nan" and cat.lower() not in valid_cats:
            errors.append({
                "source": "master_data",
                "row_number": row_num,
                "column_name": "category",
                "error_type": "INVALID_CATEGORY",
                "error_message": f"Row {row_num}: Category '{cat}' is not in the allowed list. "
                                 f"Allowed: {config.CATEGORIES}",
                "created_at": _now(),
            })
            count += 1

    # Duplicate employee codes
    dupes = df[df["employee_code"].duplicated(keep=False)]
    if not dupes.empty:
        for code in dupes["employee_code"].unique():
            rows_with_code = df[df["employee_code"] == code].index.tolist()
            errors.append({
                "source": "master_data",
                "row_number": rows_with_code[0] + 1,
                "column_name": "employee_code",
                "error_type": "DUPLICATE_CODE",
                "error_message": f"Employee Code '{code}' appears {len(rows_with_code)} times "
                                 f"(rows {[r+1 for r in rows_with_code]}).",
                "created_at": _now(),
            })
            count += 1

    return count


# ---------------------------------------------------------------------------
# 2. BACKEND DATA VALIDATION
# ---------------------------------------------------------------------------

def _validate_backend_data(errors: list) -> int:
    """
    Validate all rows in the backend_data table.
    Checks employee codes against master_data, quantity and date validity.
    Returns count of problems found.
    """
    backend_df = database.read_table("backend_data")
    if backend_df.empty:
        errors.append({
            "source": "backend_data",
            "row_number": 0,
            "column_name": "—",
            "error_type": "NO_DATA",
            "error_message": "Backend Data table is empty. Upload a file first.",
            "created_at": _now(),
        })
        return 1

    master_df = database.read_table("master_data")
    known_codes = set(master_df["employee_code"].astype(str).str.strip().tolist()) \
        if not master_df.empty else set()

    count = 0

    # Use vectorised checks for speed on 60K+ rows ─────────────────────────

    # Blank / nan employee codes
    codes = backend_df["created_by"].fillna("").astype(str).str.strip()
    blank_mask = (codes == "") | (codes.str.lower() == "nan")
    for i in backend_df[blank_mask].index:
        errors.append({
            "source": "backend_data",
            "row_number": int(i) + 1,
            "column_name": "created_by",
            "error_type": "BLANK_EMPLOYEE_CODE",
            "error_message": f"Row {int(i)+1}: 'created_by' is blank.",
            "created_at": _now(),
        })
        count += 1

    # Unmapped employee codes (not in master data) — deduplicated for readability
    valid_rows = backend_df[~blank_mask].copy()
    valid_rows["_code"] = valid_rows["created_by"].astype(str).str.strip()
    unmapped_codes = valid_rows[~valid_rows["_code"].isin(known_codes)]["_code"].unique()
    for code in unmapped_codes:
        row_count = (valid_rows["_code"] == code).sum()
        first_row = valid_rows[valid_rows["_code"] == code].index[0]
        errors.append({
            "source": "backend_data",
            "row_number": int(first_row) + 1,
            "column_name": "created_by",
            "error_type": "UNMAPPED_EMPLOYEE",
            "error_message": f"Employee Code '{code}' not found in Master Data "
                             f"({row_count} row(s) affected, first at row {int(first_row)+1}).",
            "created_at": _now(),
        })
        count += 1

    # Invalid quantity
    qtys = pd.to_numeric(backend_df["quantity"], errors="coerce")
    bad_qty_mask = qtys.isna() | (qtys <= 0)
    for i in backend_df[bad_qty_mask].index:
        errors.append({
            "source": "backend_data",
            "row_number": int(i) + 1,
            "column_name": "quantity",
            "error_type": "INVALID_QUANTITY",
            "error_message": f"Row {int(i)+1}: Quantity '{backend_df.at[i,'quantity']}' "
                             f"is not a positive number.",
            "created_at": _now(),
        })
        count += 1

    # Invalid dates
    dates = pd.to_datetime(backend_df["date"], format="%Y-%m-%d", errors="coerce")
    bad_date_mask = dates.isna()
    for i in backend_df[bad_date_mask].index:
        errors.append({
            "source": "backend_data",
            "row_number": int(i) + 1,
            "column_name": "date",
            "error_type": "INVALID_DATE",
            "error_message": f"Row {int(i)+1}: Date '{backend_df.at[i,'date']}' "
                             f"is not a valid YYYY-MM-DD date.",
            "created_at": _now(),
        })
        count += 1

    return count


# ---------------------------------------------------------------------------
# 3. MAINTENANCE COST VALIDATION
# ---------------------------------------------------------------------------

def _validate_maintenance_cost(errors: list) -> int:
    """
    Validate all rows in the maintenance_cost table.
    Checks plant codes against master_data and that costs are >= 0.
    Returns count of problems found.
    """
    maint_df = database.read_table("maintenance_cost")
    if maint_df.empty:
        errors.append({
            "source": "maintenance_cost",
            "row_number": 0,
            "column_name": "—",
            "error_type": "NO_DATA",
            "error_message": "Maintenance Cost table is empty. Upload a file first.",
            "created_at": _now(),
        })
        return 1

    master_df = database.read_table("master_data")
    known_plant_codes = set(master_df["plant_code"].astype(str).str.strip().tolist()) \
        if not master_df.empty else set()

    count = 0

    for i, row in maint_df.iterrows():
        row_num = int(i) + 1

        pc = str(row.get("plant_code", "")).strip()
        if not pc or pc.lower() == "nan":
            errors.append({
                "source": "maintenance_cost",
                "row_number": row_num,
                "column_name": "plant_code",
                "error_type": "BLANK_PLANT_CODE",
                "error_message": f"Row {row_num}: 'plant_code' is blank.",
                "created_at": _now(),
            })
            count += 1
        elif pc not in known_plant_codes:
            errors.append({
                "source": "maintenance_cost",
                "row_number": row_num,
                "column_name": "plant_code",
                "error_type": "UNMAPPED_PLANT",
                "error_message": f"Row {row_num}: Plant Code '{pc}' not found in Master Data.",
                "created_at": _now(),
            })
            count += 1

        try:
            cost = float(row.get("ytd_maintenance_cost", 0))
            if cost < 0:
                raise ValueError
        except (ValueError, TypeError):
            errors.append({
                "source": "maintenance_cost",
                "row_number": row_num,
                "column_name": "ytd_maintenance_cost",
                "error_type": "INVALID_COST",
                "error_message": f"Row {row_num}: YTD Maintenance Cost "
                                 f"'{row.get('ytd_maintenance_cost')}' is not valid.",
                "created_at": _now(),
            })
            count += 1

    return count


# ---------------------------------------------------------------------------
# 4. MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def run_all_validations() -> dict:
    """
    Run all three validations, save errors to SQLite, and return a summary.

    Returns:
        {
            "total_errors":       int,
            "master_errors":      int,
            "backend_errors":     int,
            "maintenance_errors": int,
            "ran_at":             str,
        }
    """
    clear_validation_errors()
    errors: list = []

    master_count      = _validate_master_data(errors)
    backend_count     = _validate_backend_data(errors)
    maintenance_count = _validate_maintenance_cost(errors)

    if errors:
        database.insert_rows("validation_errors", errors)

    ran_at = _now()
    database.set_setting("last_validation_at",     ran_at)
    database.set_setting("last_validation_errors", str(len(errors)))

    return {
        "total_errors":       len(errors),
        "master_errors":      master_count,
        "backend_errors":     backend_count,
        "maintenance_errors": maintenance_count,
        "ran_at":             ran_at,
    }


# ---------------------------------------------------------------------------
# 5. SUMMARY HELPERS
# ---------------------------------------------------------------------------

def get_validation_summary() -> pd.DataFrame:
    """
    Return a DataFrame of error counts grouped by source and error_type.
    Returns an empty DataFrame if no errors are stored.
    """
    df = database.read_table("validation_errors")
    if df.empty:
        return pd.DataFrame(columns=["Source", "Error Type", "Count"])
    summary = (
        df.groupby(["source", "error_type"])
        .size()
        .reset_index(name="Count")
        .rename(columns={"source": "Source", "error_type": "Error Type"})
        .sort_values(["Source", "Count"], ascending=[True, False])
        .reset_index(drop=True)
    )
    return summary


def get_last_validation_info() -> dict:
    """Return when validation was last run and how many errors were found."""
    return {
        "last_run":    database.get_setting("last_validation_at",     None),
        "error_count": database.get_setting("last_validation_errors", "0"),
    }
