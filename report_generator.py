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
    "total_quantity":       "Batching Quantity",
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
_INT_COLS = {"Month", "Year", "Row Number", "Sr. no."}
_NUM_COLS = {
    "Batching Quantity", "YTD Maintenance Cost", "Incentive Rate",
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

    bands = {"normal": None, "green": "#92D492", "red": "#FFB3B3"}
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


def _write_table(ws, fmts, out_df, header_overrides=None, add_srno=True):
    """Write a DataFrame to a worksheet with header, colours and widths."""
    headers = list(out_df.columns)
    overrides = header_overrides or {}

    # Prepend Sr. No. column
    if add_srno:
        headers = ["Sr. No."] + headers

    # Header row (row height 30 to allow text wrap on long headers)
    ws.set_row(0, 30)
    for c, h in enumerate(headers):
        ws.write(0, c, overrides.get(h, h), fmts["header"])
    ws.freeze_panes(1, 0)

    if out_df.empty:
        ws.write(1, 0, "No records.", fmts["normal_text"])
        for c, h in enumerate(headers):
            ws.set_column(c, c, max(12, 8))
        return

    # Data rows
    data_headers = headers[1:] if add_srno else headers
    for r, (_, row) in enumerate(out_df.iterrows(), start=1):
        inc = row.get("Incentive Amount", 0) or 0
        ded = row.get("Deduction Amount", 0) or 0
        band = "red" if ded > 0 else ("green" if inc > 0 else "normal")
        col_offset = 0
        if add_srno:
            ws.write(r, 0, r, fmts[f"{band}_int"])
            col_offset = 1
        for c, h in enumerate(data_headers):
            fmt = fmts[f"{band}_{_coltype(h)}"]
            val = row[h]
            if pd.isna(val):
                ws.write_blank(r, c + col_offset, None, fmt)
            else:
                ws.write(r, c + col_offset, val, fmt)

    # Column widths — driven by data length only (headers wrap, so ignore header length)
    for c, h in enumerate(headers):
        if add_srno and c == 0:
            ws.set_column(0, 0, 6)   # Sr. No. — narrow
            continue
        data_col = h
        if data_col in out_df.columns:
            longest = out_df[data_col].astype(str).str.len().max()
            longest = 0 if pd.isna(longest) else int(longest)
        else:
            longest = 0
        ws.set_column(c, c, min(max(longest + 1, 10), 30))


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


def build_email_tables_html(results_df, sections=None) -> str:
    """
    Build per-section color-coded HTML tables for the I&D report email body.
    Uses fully inline styles (no <style> block) to survive Gmail's CSS stripper.
    Layout: report name above the table (separate <p>), then the table with
    Sr. No. as first column (no Month/Year/Plant Code columns).
    TH cells wrap text so long headers don't blow out column widths.
    """
    import html as _html
    sections = sections or EMAIL_SECTIONS

    F        = "font-family:Arial,sans-serif;font-size:11px;"
    HDR_BDR  = "border:1px solid #7A7A7A"
    CELL_BDR = "border:1px solid #9E9E9E"

    def _row_bg(inc, ded, remark=""):
        if ded > 0 or "waiv" in str(remark).lower(): return "#FFB3B3"
        if inc > 0: return "#92D492"
        return "#ffffff"

    parts = []

    for idx, (title, cats) in enumerate(sections, start=1):
        out       = _prepare_category_df(results_df, cats)
        headers   = ["Sr. No."] + list(out.columns) if not out.empty else ["Sr. No."]
        ded_label = _ded_header_label(cats)
        num_cols  = len(headers)

        # TH: wrap text (no nowrap) so long headers don't force wide columns
        th_base = (
            f'{F}background-color:#082B49;color:#ffffff;font-weight:bold;'
            f'padding:3px 5px;{HDR_BDR};white-space:normal;line-height:1.2;'
            f'vertical-align:middle;text-align:center;'
        )

        thead_cells = ""
        for h in headers:
            label = "Sr. No." if h == "Sr. No." else (ded_label if h == "Deduction Amount" else str(h))
            # width goes INSIDE the style string — two style attrs would make the second win
            extra = "width:32px;" if h == "Sr. No." else ""
            thead_cells += f'<th style="{th_base}{extra}">{_html.escape(label)}</th>'

        tbody_rows = []
        if out.empty:
            tbody_rows.append(
                f'<tr><td colspan="{num_cols}" style="{F}padding:2px 5px;{CELL_BDR};'
                f'color:#777;text-align:center;background:#ffffff">'
                f'No records for this section.</td></tr>'
            )
        else:
            data_headers = list(out.columns)
            for srno, (_, row) in enumerate(out.iterrows(), start=1):
                inc    = row.get("Incentive Amount", 0) or 0
                ded    = row.get("Deduction Amount", 0) or 0
                remark = row.get("Remarks", "") or ""
                bg     = _row_bg(inc, ded, remark)
                # Sr. No. cell
                sr_s = (f'{F}background:{bg};color:#000;padding:2px 4px;{CELL_BDR};'
                        f'text-align:center;line-height:1.2;vertical-align:middle;white-space:nowrap;')
                cells = f'<td style="{sr_s}">{srno}</td>'
                for h in data_headers:
                    align = "right" if _coltype(h) in ("int", "num") else "left"
                    td_s  = (f'{F}background:{bg};padding:2px 5px;{CELL_BDR};'
                             f'text-align:{align};white-space:nowrap;line-height:1.2;vertical-align:middle')
                    cells += f'<td style="{td_s}">{_html.escape(_cell_value(h, row[h]))}</td>'
                tbody_rows.append(f"<tr>{cells}</tr>")

        # Report name: navy bar above the column header row, same width as the table
        title_style = (
            f'{F}font-size:12px;font-weight:bold;'
            f'background-color:#082B49;color:#ffffff;'
            f'padding:4px 6px;border:1px solid #7A7A7A;'
            f'display:block;margin:8px 0 0 0;'
        )
        parts.append(
            f'<div style="{title_style}">{idx}. {_html.escape(title)}</div>'
            f'<table cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;width:auto;margin:0 0 10px">'
            f'<thead><tr>{thead_cells}</tr></thead>'
            f'<tbody>{"".join(tbody_rows)}</tbody>'
            f'</table>'
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
                 _prepare_unmapped_df(unmapped_df), add_srno=False)

    # 4) Validation Errors
    _write_table(wb.add_worksheet("Validation Errors"), fmts,
                 _prepare_validation_df(validation_df), add_srno=False)

    writer.close()
    return output.getvalue()
