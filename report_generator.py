"""
report_generator.py
===================
Phase 10 — turn the calculation results into ONE formatted Excel workbook.

The workbook has these sheets (see config.REPORT_SHEETS):
  - Summary                (report info + totals)
  - one sheet per category group (All Trainees, SPE, QCI, ...)
  - Unmapped Employees
  - Validation Errors

Formatting applied to every data sheet:
  - Bold header row with a dark fill and white font.
  - Header row frozen (stays visible while scrolling).
  - Column widths auto-adjusted to the content.
  - Number formatting on quantity / amount columns.
  - Row colours: incentive row -> light green, deduction row -> light red.
    If BOTH are present, light red wins (deduction priority).
  - Rows sorted: deductions (red) first (largest first), then incentives
    (green, largest first), then plain rows.

The single public function is generate_excel_report(...), which returns the
finished workbook as bytes (ready for a Streamlit download button or email
attachment). All Excel code lives HERE so the rest of the app stays clean.
"""

import io

import pandas as pd

import config


# ---------------------------------------------------------------------------
# Column mapping: internal DB name -> friendly Excel header
# ---------------------------------------------------------------------------
_COL_MAP = {
    "month":                "Month",
    "year":                 "Year",
    "employee_code":        "Employee Code",
    "employee_name":        "Employee Name",
    "designation":          "Designation",
    "category":             "Category",
    "plant":                "Plant",
    "plant_code":           "Plant Code",
    "total_quantity":       "Total Quantity",
    "ytd_maintenance_cost": "YTD Maintenance Cost",
    "incentive_eligible":   "Incentive Eligible",
    "incentive_rate":       "Incentive Rate",
    "incentive_amount":     "Incentive Amount",
    "deduction_target":     "Deduction Target",
    "shortfall_quantity":   "Shortfall Quantity",
    "deduction_amount":     "Deduction Amount",
    "remarks":              "Remarks",
}

# Which Excel columns are numbers (so we right-align + number-format them).
_INT_COLS = {"Month", "Year", "Row Number"}
_NUM_COLS = {
    "Total Quantity", "YTD Maintenance Cost", "Incentive Rate",
    "Incentive Amount", "Deduction Target", "Shortfall Quantity",
    "Deduction Amount",
}


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def _safe_sheet_name(name: str) -> str:
    """Excel sheet names must be <= 31 characters and avoid : \\ / ? * [ ]."""
    for bad in r":\/?*[]":
        name = name.replace(bad, " ")
    return name[:31]


def _coltype(col: str) -> str:
    """Return 'int', 'num' or 'text' for a given Excel column header."""
    if col in _INT_COLS:
        return "int"
    if col in _NUM_COLS:
        return "num"
    return "text"


def _sort_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Deduction rows first (largest first), then incentive rows, then plain."""
    out = df.copy()
    out["_band"] = 2
    out.loc[out["incentive_amount"] > 0, "_band"] = 1
    out.loc[out["deduction_amount"] > 0, "_band"] = 0
    out = out.sort_values(
        by=["_band", "deduction_amount", "incentive_amount"],
        ascending=[True, False, False],
    ).drop(columns="_band")
    return out


def _ded_header_label(categories) -> str:
    """e.g. 'Deduction Amount @ Rs 20' when all categories share one rate."""
    rates = {config.DEDUCTION_RULES.get(c, {}).get("rate", 0) for c in categories}
    if len(rates) == 1 and next(iter(rates)):
        return f"Deduction Amount @ Rs {next(iter(rates))}"
    return "Deduction Amount"


def _build_formats(wb):
    """Create and return the dictionary of reusable cell formats."""
    fmts = {}
    fmts["header"] = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": "#0A2540",
        "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1,
    })
    fmts["label"] = wb.add_format({"bold": True, "align": "left", "border": 1})

    bands = {"normal": None, "green": "#D4EDDA", "red": "#F8D7DA"}
    types = {"text": None, "int": "#,##0", "num": "#,##0.00"}
    for bname, bcolor in bands.items():
        for tname, numfmt in types.items():
            spec = {"border": 1, "valign": "vcenter"}
            if bcolor:
                spec["bg_color"] = bcolor
            if numfmt:
                spec["num_format"] = numfmt
            spec["align"] = "left" if tname == "text" else "right"
            fmts[f"{bname}_{tname}"] = wb.add_format(spec)
    return fmts


def _write_table(ws, fmts, out_df, header_overrides=None):
    """Write a DataFrame to a worksheet with header, colours and widths."""
    headers = list(out_df.columns)
    overrides = header_overrides or {}

    # Header row (frozen)
    for c, h in enumerate(headers):
        ws.write(0, c, overrides.get(h, h), fmts["header"])
    ws.freeze_panes(1, 0)

    if out_df.empty:
        ws.write(1, 0, "No records.", fmts["normal_text"])
        for c, h in enumerate(headers):
            ws.set_column(c, c, max(len(str(overrides.get(h, h))) + 2, 10))
        return

    # Data rows
    for r, (_, row) in enumerate(out_df.iterrows(), start=1):
        inc = row.get("Incentive Amount", 0) or 0
        ded = row.get("Deduction Amount", 0) or 0
        band = "red" if ded > 0 else ("green" if inc > 0 else "normal")
        for c, h in enumerate(headers):
            fmt = fmts[f"{band}_{_coltype(h)}"]
            val = row[h]
            if pd.isna(val):
                ws.write_blank(r, c, None, fmt)
            else:
                ws.write(r, c, val, fmt)

    # Auto column widths (header length vs longest value)
    for c, h in enumerate(headers):
        longest = out_df[h].astype(str).str.len().max()
        longest = 0 if pd.isna(longest) else int(longest)
        width = max(len(str(overrides.get(h, h))), longest) + 2
        ws.set_column(c, c, min(max(width, 9), 42))


def _prepare_category_df(results_df, categories) -> pd.DataFrame:
    """Filter results to the given categories, sort, and rename to Excel headers."""
    empty = pd.DataFrame(columns=config.REPORT_OUTPUT_COLUMNS)
    if results_df is None or results_df.empty or not categories:
        return empty
    df = results_df[results_df["category"].isin(categories)].copy()
    if df.empty:
        return empty
    df = _sort_rows(df)
    cols = [c for c in _COL_MAP if c in df.columns]
    out = df[cols].rename(columns=_COL_MAP)
    return out[[c for c in config.REPORT_OUTPUT_COLUMNS if c in out.columns]] \
        .reset_index(drop=True)


# The seven email sections in order, each mapped to the categories it covers.
# Titles match email_helper.REPORT_SECTIONS (and the company Mail Reference).
EMAIL_SECTIONS = [
    ("Production report of all Trainees", ["Civil Trainee", "Non-Civil Trainee"]),
    ("Production report of PM/API", ["PM & API"]),
    ("Production Report of QC Person", ["QCI"]),
    ("Production report of MO", ["MO"]),
    ("Production report of SPE", ["SPE"]),
    ("Production report of Officer Production (Onroll Batcher)", ["Production Officer"]),
    ("Production report of TL Batcher, Mechanic", ["TL BPO"]),
]


def _cell_value(col: str, val) -> str:
    """Format one value for an HTML table cell (numbers nicely, blanks as '')."""
    if pd.isna(val):
        return ""
    if col in _INT_COLS:
        # Month / Year / Row Number are plain integers — no thousands comma
        # (so the year reads "2026", not "2,026").
        try:
            return str(int(val))
        except (ValueError, TypeError):
            return str(val)
    if col in _NUM_COLS:
        try:
            return f"{float(val):,.2f}"
        except (ValueError, TypeError):
            return str(val)
    return str(val)


# Compact CSS (one <style> block per email) keeps the message small enough to
# avoid email clipping. Row colour matches the Excel: red = deduction row,
# green = incentive row. Emitted once at the top of build_email_tables_html.
_EMAIL_TABLE_CSS = (
    "<style>"
    ".rpt-h{font-family:Arial,sans-serif;color:#1a56db;margin:18px 0 6px;font-size:15px}"
    ".rpt{border-collapse:collapse;font-family:Arial,sans-serif;margin:0 0 12px}"
    ".rpt th{background:#0A2540;color:#fff;border:1px solid #2a4660;padding:4px 7px;"
    "font-size:11px;text-align:left;white-space:nowrap}"
    ".rpt td{border:1px solid #cbd5e1;padding:3px 7px;font-size:11px}"
    ".rpt tr.r td{background:#F8D7DA}"      # deduction row (red)
    ".rpt tr.g td{background:#D4EDDA}"      # incentive row (green)
    "</style>"
)


def build_email_tables_html(results_df, sections=None) -> str:
    """
    Build the HTML for the per-section report tables shown INSIDE the email body
    (in addition to the Excel attachment). One heading + one colour-coded table
    per section, matching the Excel: red row = deduction, green row = incentive.

    Uses a single <style> block + CSS classes (not per-cell inline styles) so the
    email stays small enough not to be clipped by Gmail.
    """
    import html as _html
    sections = sections or EMAIL_SECTIONS
    parts = [_EMAIL_TABLE_CSS]

    for idx, (title, cats) in enumerate(sections, start=1):
        parts.append(f'<h3 class="rpt-h">{idx}. {_html.escape(title)}</h3>')
        out = _prepare_category_df(results_df, cats)
        if out.empty:
            parts.append('<p style="font-family:Arial,sans-serif;font-size:12px;'
                         'color:#777;margin:0 0 8px">No records for this section.</p>')
            continue

        headers = list(out.columns)
        ded_label = _ded_header_label(cats)
        head_cells = "".join(
            f'<th>{_html.escape(ded_label if h == "Deduction Amount" else str(h))}</th>'
            for h in headers
        )

        body_rows = []
        for _, row in out.iterrows():
            inc = row.get("Incentive Amount", 0) or 0
            ded = row.get("Deduction Amount", 0) or 0
            cls = ' class="r"' if ded > 0 else (' class="g"' if inc > 0 else "")
            cells = "".join(
                f"<td>{_html.escape(_cell_value(h, row[h]))}</td>" for h in headers
            )
            body_rows.append(f"<tr{cls}>{cells}</tr>")

        parts.append(
            f'<table class="rpt"><thead><tr>{head_cells}</tr></thead>'
            f'<tbody>{"".join(body_rows)}</tbody></table>'
        )

    return "".join(parts)


def _prepare_unmapped_df(unmapped_df) -> pd.DataFrame:
    cols = {"employee_code": "Employee Code", "month": "Month", "year": "Year",
            "total_quantity": "Total Quantity", "remarks": "Remarks"}
    if unmapped_df is None or unmapped_df.empty:
        return pd.DataFrame(columns=list(cols.values()))
    keep = [c for c in cols if c in unmapped_df.columns]
    return unmapped_df[keep].rename(columns=cols).reset_index(drop=True)


def _prepare_validation_df(validation_df) -> pd.DataFrame:
    cols = {"source": "Source", "row_number": "Row Number",
            "column_name": "Column", "error_type": "Error Type",
            "error_message": "Error Message", "created_at": "Created At"}
    if validation_df is None or validation_df.empty:
        return pd.DataFrame(columns=list(cols.values()))
    keep = [c for c in cols if c in validation_df.columns]
    return validation_df[keep].rename(columns=cols).reset_index(drop=True)


def _write_summary(ws, fmts, meta):
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 46)
    ws.write(0, 0, "Summary", fmts["header"])
    ws.write(0, 1, "", fmts["header"])

    rows = [
        ("Report Generated On",     meta.get("generated_on", "")),
        ("Selected Date Range",     meta.get("date_range", "")),
        ("Applied Filters",         meta.get("applied_filters", "None")),
        ("Total Employees",         meta.get("total_employees", 0)),
        ("Total Quantity",          meta.get("total_quantity", 0)),
        ("Total Incentive Amount",  meta.get("total_incentive", 0)),
        ("Total Deduction Amount",  meta.get("total_deduction", 0)),
        ("Unmapped Employee Count", meta.get("unmapped_count", 0)),
        ("Validation Error Count",  meta.get("validation_count", 0)),
    ]
    for i, (k, v) in enumerate(rows, start=1):
        ws.write(i, 0, k, fmts["label"])
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            fmt = fmts["normal_num"] if isinstance(v, float) else fmts["normal_int"]
            ws.write_number(i, 1, v, fmt)
        else:
            ws.write(i, 1, str(v), fmts["normal_text"])


# ---------------------------------------------------------------------------
# PUBLIC ENTRY POINT
# ---------------------------------------------------------------------------
def generate_excel_report(results_df, unmapped_df, validation_df, meta) -> bytes:
    """
    Build the full multi-sheet Excel workbook and return it as bytes.

    results_df    -> the (already filtered) calculation results DataFrame
                     using the internal column names.
    unmapped_df   -> unmapped employees DataFrame (internal names) or None.
    validation_df -> validation_errors DataFrame (internal names) or None.
    meta          -> dict with summary info (generated_on, date_range,
                     applied_filters, totals, counts).
    """
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    wb = writer.book
    fmts = _build_formats(wb)

    # 1) Summary sheet
    _write_summary(wb.add_worksheet("Summary"), fmts, meta)

    # 2) One sheet per category group (skip the special sheets)
    special = {"Summary", "Unmapped Employees", "Validation Errors"}
    for sheet_name, categories in config.REPORT_SHEETS.items():
        if sheet_name in special:
            continue
        out_df = _prepare_category_df(results_df, categories)
        ws = wb.add_worksheet(_safe_sheet_name(sheet_name))
        _write_table(ws, fmts, out_df,
                     header_overrides={"Deduction Amount": _ded_header_label(categories)})

    # 3) Unmapped Employees
    _write_table(wb.add_worksheet("Unmapped Employees"), fmts,
                 _prepare_unmapped_df(unmapped_df))

    # 4) Validation Errors
    _write_table(wb.add_worksheet("Validation Errors"), fmts,
                 _prepare_validation_df(validation_df))

    writer.close()
    return output.getvalue()
