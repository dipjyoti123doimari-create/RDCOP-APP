"""
email_helper.py
===============
Phase 12 — send the generated Excel report by email.

How the SMTP settings are found (in priority order):
  - The password comes from the environment variable SMTP_PASSWORD if set,
    otherwise from the value saved on the Settings page (stored locally in the
    app database). The password is NEVER hardcoded and NEVER printed/logged.
  - Host, port, sender, TLS and default recipients come from the Settings page
    (the app_settings table).

Public functions:
  get_smtp_config()              -> dict of current SMTP settings
  is_configured()                -> True when host + sender + password are set
  wrap_html_body_with_image(...) -> HTML body string with inline CID image
  create_tp_preview_image(...)   -> generate TP report PNG, return path or None
  create_report_preview_image(.) -> generate I&D report PNG, return path or None
  send_report_email(...)         -> {"success": bool, "error": str | None}

Email design: PNG preview image embedded inline (CID) so colors and borders
display correctly in Gmail/Outlook regardless of CSS stripping. Full Excel
report is attached separately.
"""

import os
import smtplib
from email.message import EmailMessage

import database


_DEFAULT_PORT  = 587
_XLSX_MAINTYPE = "application"
_XLSX_SUBTYPE  = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_smtp_config() -> dict:
    """Read all email settings. Password falls back to the SMTP_PASSWORD env var."""
    s = database.get_all_settings()
    password = os.environ.get("SMTP_PASSWORD") or s.get("smtp_password", "")
    raw_port = s.get("smtp_port", str(_DEFAULT_PORT))
    try:
        port = int(raw_port)
    except (TypeError, ValueError):
        port = _DEFAULT_PORT

    return {
        "host":       (s.get("smtp_host",   "") or "").strip(),
        "port":       port,
        "sender":     (s.get("smtp_sender", "") or "").strip(),
        "password":   password,
        "use_tls":    s.get("smtp_use_tls", "true") == "true",
        "default_to": s.get("email_default_to", "") or "",
        "default_cc": s.get("email_default_cc", "") or "",
        "subject":    s.get("email_subject", "Batching Incentive & Deduction Report"),
    }


def is_configured() -> bool:
    """True when the essentials (host, sender, password) are all set."""
    c = get_smtp_config()
    return bool(c["host"] and c["sender"] and c["password"])


REPORT_SECTIONS = [
    "Production report of all Trainees",
    "Production report of PM/API",
    "Production Report of QC Person",
    "Production report of MO",
    "Production report of SPE",
    "Production report of Officer Production (Onroll Batcher)",
    "Production report of TL Batcher, Mechanic",
]


def compose_report_subject(month_label) -> str:
    return ("Report of Production, Penalty & Incentive summary for "
            f"(PI/QCI, MO, Teamlease Employee & All Trainees) - {month_label}")


def compose_report_body(month_label, sections=None) -> str:
    sections = sections or REPORT_SECTIONS
    numbered = "\n".join(f"{i}. {s}" for i, s in enumerate(sections, start=1))
    return (
        "Dear Sir,\n\n"
        f"Please find below the compiled report for the month of {month_label}, "
        "covering the following:\n\n"
        f"{numbered}\n\n"
        "Kindly review the attached Excel file below and in case of any "
        "clarifications or corrections, please mail me and Kanhaiya sir"
    )


def _split_emails(text) -> list:
    if not text:
        return []
    parts = [p.strip() for p in str(text).replace(";", ",").split(",")]
    return [p for p in parts if p]


def wrap_html_body(message_text, tables_html="") -> str:
    """Legacy plain HTML body (used as fallback when image generation fails)."""
    import html as _html
    safe = _html.escape(message_text or "").replace("\n", "<br>")
    return (
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;'
        'color:#0A2540;line-height:1.45">'
        f'<div>{safe}</div>'
        f'{tables_html or ""}'
        '</div>'
    )


def wrap_html_body_with_image(message_text, excel_attached=True) -> str:
    """
    HTML email body that references an inline PNG via CID (Content-ID).
    The caller must attach the image separately with Content-ID <report_preview>.
    """
    import html as _html
    safe = _html.escape(message_text or "").replace("\n", "<br>")
    note = (
        '<p style="color:#555;font-size:12px;margin-top:14px">'
        '&#128206; Full detailed report is attached as an Excel file.</p>'
        if excel_attached else ""
    )
    return (
        '<!DOCTYPE html><html><body>'
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;'
        'color:#0A2540;line-height:1.6;max-width:960px;margin:0 auto;padding:20px">'
        f'<div style="margin-bottom:18px">{safe}</div>'
        '<div style="margin:18px 0">'
        '<img src="cid:report_preview" alt="Report Preview" '
        'style="max-width:100%;height:auto;border-radius:14px;display:block;margin:auto">'
        '</div>'
        f'{note}'
        '<p style="color:#333;margin-top:18px">Regards</p>'
        '</div></body></html>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# PREMIUM DARK GLOSSY PNG PREVIEW IMAGE (Pillow)
# Dark gradient background + visible spreadsheet-style cell borders +
# glossy dark-tinted conditional row colors (red/amber/green on dark bg).
# White text, high-DPI 2× scale — looks great in Gmail / Outlook.
# ─────────────────────────────────────────────────────────────────────────────

_S    = 2    # 2× Retina scale factor
_OP   = 26   # outer image padding (virtual px)
_CP   = 8    # cell inner padding (virtual px)
_BH   = 62   # main banner height (virtual px)
_SH   = 38   # section title row height
_HH   = 42   # column header row height
_DROW = 30   # base data row height
_LH   = 17   # text line height (virtual px)
_GAP  = 16   # gap between table sections

# Premium dark palette — all opaque RGB, simulating glass/gloss on dark bg
_D = {
    "bg_top":  ( 8,  26,  51),   # #081A33 gradient top
    "bg_bot":  (11,  22,  40),   # #0B1628 gradient bottom
    "banner":  ( 4,  14,  30),   # #060E1E main title bar
    "sec":     ( 8,  43,  73),   # #082B49 section title bg
    "hdr":     (10,  37,  64),   # #0A2540 column header bg
    "row_odd": (14,  33,  56),   # #0E2138 neutral odd row
    "row_even":(11,  26,  46),   # #0B1A2E neutral even row
    "pan_row": (18,  45,  74),   # #122D4A PAN India special row
    "red":     (90,  20,  30),   # #5A141E deduction / below threshold
    "amber":   (90,  62,  14),   # #5A3E0E warning / mid range
    "green":   (14,  72,  46),   # #0E482E good / incentive
    "bdr_out": (143, 163, 184),  # #8FA3B8 outer border (2 px)
    "bdr_hdr": ( 95, 120, 149),  # #5F7895 header cell borders
    "bdr_in":  ( 49,  68,  92),  # #31445C inner data cell borders
    "txt":     (248, 250, 252),  # #F8FAFC main white text
    "txt_dim": (203, 213, 225),  # #CBD5E1 secondary text
    "txt_mut": (148, 163, 184),  # #94A3B8 muted footer text
}


def _pil_font(pt, bold=False):
    from PIL import ImageFont
    face = "arialbd.ttf" if bold else "arial.ttf"
    for path in [
        f"C:/Windows/Fonts/{face}",
        f"/usr/share/fonts/truetype/liberation/Liberation{'Bold' if bold else 'Regular'}.ttf",
    ]:
        try:
            return ImageFont.truetype(path, pt * _S)
        except OSError:
            pass
    return ImageFont.load_default()


def _fw(font, text):
    """Text width from font metrics (no ImageDraw needed)."""
    try:
        bb = font.getbbox(str(text))
        return bb[2] - bb[0]
    except Exception:
        try:
            return font.getsize(str(text))[0]
        except Exception:
            return len(str(text)) * max(6, getattr(font, "size", 10) // 2)


def _wrap(text, font, max_px):
    """Split text into wrapped lines, each fitting within max_px."""
    words = str(text).split()
    if not words:
        return [""]
    lines, cur = [], words[0]
    for w in words[1:]:
        test = cur + " " + w
        if _fw(font, test) <= max_px:
            cur = test
        else:
            lines.append(cur)
            cur = w
    lines.append(cur)
    return lines or [""]


def _clip(text, font, max_px):
    """Truncate text with ellipsis to fit max_px."""
    s = str(text)
    if _fw(font, s) <= max_px:
        return s
    while s and _fw(font, s + "…") > max_px:
        s = s[:-1]
    return (s + "…") if s else ""


def _img_fmt(val):
    try:
        f = float(val)
        return f"{int(f):,}" if f == int(f) else f"{f:,.1f}"
    except (TypeError, ValueError):
        return str(val) if val is not None else ""


def _fill_gradient(img, c_top, c_bot):
    """Fill image with a vertical gradient from c_top to c_bot."""
    from PIL import Image as _Img
    w, h = img.size
    strip = _Img.new("RGB", (1, h))
    strip.putdata([
        tuple(int(c_top[j] + (c_bot[j] - c_top[j]) * y / max(h - 1, 1))
              for j in range(3))
        for y in range(h)
    ])
    img.paste(strip.resize((w, h), _Img.NEAREST))


def _draw_premium_table(draw, sec_title, col_defs, data_rows, row_color_fn,
                        x, y, f_sec, f_hdr, f_data):
    """
    Draw one premium dark table section.

    col_defs  = [(label, scaled_px_width, align, wrap_data), ...]
                align: 'l'|'c'|'r'   wrap_data: True wraps text in data cells
    data_rows = [[str, ...], ...]
    row_color_fn(row_idx, row_list) -> RGB tuple

    Returns new y position after drawing.
    """
    S    = _S
    cp   = _CP   * S
    sh   = _SH   * S
    lh   = _LH   * S
    drow = _DROW * S
    hh_b = _HH   * S
    total_w = sum(d[1] for d in col_defs)
    x1 = x + total_w

    # Pre-calculate header height (always word-wrap header labels)
    hdr_lines = [_wrap(d[0], f_hdr, d[1] - cp * 2) for d in col_defs]
    hh = max(hh_b, max(len(ls) for ls in hdr_lines) * lh + cp * 2)

    # Pre-calculate data row heights
    all_rh, all_wrapped = [], []
    for row in data_rows:
        cells_lines, max_l = [], 1
        for ci, (_, w, _, do_wrap) in enumerate(col_defs):
            txt = str(row[ci]) if ci < len(row) else ""
            ls  = (_wrap(txt, f_data, w - cp * 2) if do_wrap
                   else [_clip(txt, f_data, w - cp * 2)])
            cells_lines.append(ls)
            max_l = max(max_l, len(ls))
        all_rh.append(max(drow, max_l * lh + cp * 2))
        all_wrapped.append(cells_lines)

    # Section title bar
    draw.rectangle([x, y, x1, y + sh], fill=_D["sec"])
    draw.text((x + cp, y + (sh - 11 * S) // 2),
              str(sec_title), fill=_D["txt"], font=f_sec)
    draw.line([x, y,   x1, y],   fill=_D["bdr_out"], width=S)
    draw.line([x, y,   x,  y + sh], fill=_D["bdr_out"], width=S)
    draw.line([x1, y,  x1, y + sh], fill=_D["bdr_out"], width=S)
    draw.line([x, y + sh, x1, y + sh], fill=_D["bdr_hdr"], width=1)
    y += sh

    # Column header row
    cx = x
    for (_, w, align, _), lines in zip(col_defs, hdr_lines):
        draw.rectangle([cx, y, cx + w, y + hh], fill=_D["hdr"])
        draw.line([cx + w, y, cx + w, y + hh], fill=_D["bdr_hdr"], width=1)
        tby = y + max(cp // 2, (hh - len(lines) * lh) // 2)
        for line in lines:
            tw = _fw(f_hdr, line)
            tx = (cx + max(cp, (w - tw) // 2) if align == 'c'
                  else (cx + w - tw - cp if align == 'r' else cx + cp))
            draw.text((tx, tby), line, fill=_D["txt"], font=f_hdr)
            tby += lh
        cx += w
    draw.line([x, y,    x,  y + hh], fill=_D["bdr_out"], width=S)
    draw.line([x1, y,   x1, y + hh], fill=_D["bdr_out"], width=S)
    draw.line([x, y + hh, x1, y + hh], fill=_D["bdr_hdr"], width=S)
    y += hh

    # Data rows: fill → border → text (order ensures borders stay visible)
    for ri, (row, rh, cells_lines) in enumerate(
            zip(data_rows, all_rh, all_wrapped)):
        bg = row_color_fn(ri, row)
        cx = x
        for (_, w, align, _), lines in zip(col_defs, cells_lines):
            draw.rectangle([cx, y, cx + w, y + rh], fill=bg)
            draw.line([cx + w, y, cx + w, y + rh], fill=_D["bdr_in"], width=1)
            draw.line([cx, y + rh, cx + w, y + rh], fill=_D["bdr_in"], width=1)
            tby = y + max(cp // 2, (rh - len(lines) * lh) // 2)
            for line in lines:
                tw = _fw(f_data, line)
                tx = (cx + max(cp, (w - tw) // 2) if align == 'c'
                      else (cx + w - tw - cp if align == 'r' else cx + cp))
                draw.text((tx, tby), line, fill=_D["txt"], font=f_data)
                tby += lh
            cx += w
        draw.line([x, y,   x,  y + rh], fill=_D["bdr_out"], width=S)
        draw.line([x1, y,  x1, y + rh], fill=_D["bdr_out"], width=S)
        y += rh

    draw.line([x, y, x1, y], fill=_D["bdr_out"], width=S)
    return y


def _tp_row_color(pct_str, name):
    """Map TP % string + location name to dark glossy RGB color."""
    if "■" in name or "PAN" in name.upper():
        return _D["pan_row"]
    try:
        pct = float(str(pct_str).replace("%", "").strip())
    except ValueError:
        return _D["row_odd"]
    if pct < 50: return _D["red"]
    if pct < 70: return _D["amber"]
    return _D["green"]


def create_tp_preview_image(plant_rows, location_rows, month, year,
                            output_path, max_rows=30):
    """
    Generate TP Plant Throughput Report preview PNG — premium dark glossy style.
    Dark gradient background, white text, visible cell borders, colored rows.
    Returns output_path on success, None on failure.
    """
    try:
        import calendar
        from PIL import Image, ImageDraw

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        S  = _S
        op = _OP * S
        bh = _BH * S
        gh = _GAP * S

        mon_name = calendar.month_name[month]
        mon_tag  = f"{calendar.month_abbr[month]}'{str(year)[-2:]}"

        # Column defs: (label, base_width, align, wrap_data)
        LOC_COLS = [
            ("#",                 28, 'c', False),
            ("Exco Location",    140, 'l', False),
            ("Plants",            50, 'c', False),
            ("Total Qty",         85, 'r', False),
            ("Total Time (min)", 115, 'r', False),
            ("Avg TP %",          65, 'c', False),
        ]
        PLT_COLS = [
            ("#",                 28, 'c', False),
            ("Plant",            195, 'l', True),   # ← wraps long plant names
            ("Exco Location",     95, 'l', False),
            ("Business Head",    115, 'l', False),
            ("Total Qty",         85, 'r', False),
            ("Time (min)",        80, 'r', False),
            ("TP %",              55, 'c', False),
        ]

        f_ban  = _pil_font(14, bold=True)
        f_sec  = _pil_font(11, bold=True)
        f_hdr  = _pil_font(9,  bold=True)
        f_data = _pil_font(9,  bold=False)
        f_foot = _pil_font(8,  bold=False)

        loc_defs = [(h, w * S, a, wp) for h, w, a, wp in LOC_COLS]
        plt_defs = [(h, w * S, a, wp) for h, w, a, wp in PLT_COLS]
        loc_w = sum(d[1] for d in loc_defs)
        plt_w = sum(d[1] for d in plt_defs)
        tbl_w = max(loc_w, plt_w)
        img_w = tbl_w + 2 * op

        # Build row data
        loc_rows_out = []
        for i, r in enumerate(location_rows, 1):
            pan = r.get("is_pan_india", False)
            pct = round(float(r.get("avg_throughput_pct", 0) or 0))
            loc_rows_out.append([
                "—" if pan else str(i),
                ("■ " if pan else "") + str(r.get("exco_location", "")),
                str(r.get("plant_count", "")),
                _img_fmt(r.get("total_quantity", 0)),
                _img_fmt(r.get("total_time_min", 0)),
                f"{pct}%",
            ])

        plt_rows_out = []
        for i, r in enumerate(plant_rows[:max_rows], 1):
            pct = round(float(r.get("throughput_pct", 0) or 0))
            plt_rows_out.append([
                str(i),
                str(r.get("plant_name", "")),
                str(r.get("exco_location", "")),
                str(r.get("business_head", "")),
                _img_fmt(r.get("total_quantity", 0)),
                _img_fmt(r.get("total_time_min", 0)),
                f"{pct}%",
            ])

        def _loc_color(ri, row):
            return _tp_row_color(row[-1], row[1] if len(row) > 1 else "")

        def _plt_color(ri, row):
            return _tp_row_color(row[-1], "")

        # Pre-calc table heights before creating image
        cp   = _CP   * S
        lh   = _LH   * S
        drow = _DROW * S
        hh_b = _HH   * S
        sh   = _SH   * S

        def _tbl_height(col_defs, data_rows):
            hdr_lines = [_wrap(d[0], f_hdr, d[1] - cp * 2) for d in col_defs]
            hh = max(hh_b, max(len(ls) for ls in hdr_lines) * lh + cp * 2)
            row_hs = []
            for row in data_rows:
                max_l = max(
                    len(_wrap(str(row[ci]), f_data, d[1] - cp * 2) if d[3]
                        else [_clip(str(row[ci]), f_data, d[1] - cp * 2)])
                    for ci, d in enumerate(col_defs) if ci < len(row)
                )
                row_hs.append(max(drow, max_l * lh + cp * 2))
            return sh + hh + sum(row_hs) + S * 2   # +2 bottom border

        loc_h = _tbl_height(loc_defs, loc_rows_out)
        plt_h = _tbl_height(plt_defs, plt_rows_out)
        footer_h = 30 * S if len(plant_rows) > max_rows else 0

        total_h = 2 * op + bh + gh + loc_h + gh + plt_h + gh + footer_h

        img  = Image.new("RGB", (img_w, total_h), _D["bg_top"])
        _fill_gradient(img, _D["bg_top"], _D["bg_bot"])
        draw = ImageDraw.Draw(img)

        # Main banner (dark bar at top)
        draw.rectangle([0, 0, img_w, bh], fill=_D["banner"])
        draw.line([0, bh, img_w, bh], fill=_D["bdr_hdr"], width=S)
        draw.text((op, (bh - 14 * S) // 2),
                  f"RDC-TP Plant Throughput Report  —  {mon_name} {year}",
                  fill=_D["txt"], font=f_ban)
        y = bh + gh

        y = _draw_premium_table(
            draw, f"Location wise Throughput — {mon_tag}",
            loc_defs, loc_rows_out, _loc_color, op, y, f_sec, f_hdr, f_data)
        y += gh
        y = _draw_premium_table(
            draw, f"Plant Throughput report — {mon_tag}",
            plt_defs, plt_rows_out, _plt_color, op, y, f_sec, f_hdr, f_data)

        if len(plant_rows) > max_rows:
            y += 6 * S
            draw.text((op, y),
                      (f"Showing first {max_rows} of {len(plant_rows)} plants. "
                       "Full report in attached Excel."),
                      fill=_D["txt_mut"], font=f_foot)

        img.save(output_path, "PNG", dpi=(144, 144))
        return output_path

    except Exception as exc:
        print(f"[email_helper] TP preview image error: {exc}")
        import traceback; traceback.print_exc()
        return None


def create_report_preview_image(df, output_path, max_rows=30, month_label=""):
    """
    Generate I&D Incentive & Deduction Report preview PNG — premium dark glossy style.
    Category-wise sections, dark gradient bg, white text, colored rows.
    Returns output_path on success, None on failure.
    """
    try:
        from PIL import Image, ImageDraw
        from report_generator import EMAIL_SECTIONS, _prepare_category_df

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        S  = _S
        op = _OP * S
        bh = _BH * S
        gh = _GAP * S
        cp = _CP  * S
        lh = _LH  * S
        drow = _DROW * S
        hh_b = _HH   * S
        sh   = _SH   * S

        # Columns: (renamed_key, display_header, base_width, align, wrap_data)
        COLS = [
            ("#",               "#",             28,  'c', False),
            ("Employee Name",   "Employee Name", 145, 'l', True),   # wrap names
            ("Category",        "Category",      118, 'l', False),
            ("Plant",           "Plant",          68, 'l', False),
            ("Total Quantity",  "Total Qty",      80, 'r', False),
            ("Incentive Amount","Incentive Amt",  95, 'r', False),
            ("Deduction Amount","Deduction Amt",  95, 'r', False),
            ("Remarks",         "Remarks",         62, 'l', False),
        ]
        col_keys = [c[0] for c in COLS]
        col_defs = [(c[1], c[2] * S, c[3], c[4]) for c in COLS]
        img_w    = sum(d[1] for d in col_defs) + 2 * op

        ded_idx = col_keys.index("Deduction Amount")
        inc_idx = col_keys.index("Incentive Amount")

        def _id_color(ri, row):
            try:
                ded = float(str(row[ded_idx]).replace(",", "")) if ded_idx < len(row) else 0
            except ValueError:
                ded = 0
            try:
                inc = float(str(row[inc_idx]).replace(",", "")) if inc_idx < len(row) else 0
            except ValueError:
                inc = 0
            if ded > 0: return _D["red"]
            if inc > 0: return _D["green"]
            return _D["row_odd"] if ri % 2 == 0 else _D["row_even"]

        f_ban  = _pil_font(14, bold=True)
        f_sec  = _pil_font(11, bold=True)
        f_hdr  = _pil_font(9,  bold=True)
        f_data = _pil_font(9,  bold=False)
        f_foot = _pil_font(8,  bold=False)

        # Gather section data
        sections_out = []
        total_shown  = 0
        for title, cats in EMAIL_SECTIONS:
            if total_shown >= max_rows:
                break
            sec_df = _prepare_category_df(df, cats)
            if sec_df.empty:
                continue
            sliced = sec_df.head(max_rows - total_shown)
            total_shown += len(sliced)
            rows_out = []
            for rn, (_, row) in enumerate(sliced.iterrows(), 1):
                cells = []
                for k in col_keys:
                    if k == "#":
                        cells.append(str(rn))
                    elif k in ("Total Quantity", "Incentive Amount", "Deduction Amount"):
                        val = row.get(k, "") if hasattr(row, "get") else getattr(row, k, "")
                        cells.append(_img_fmt(val))
                    else:
                        val = row.get(k, "") if hasattr(row, "get") else getattr(row, k, "")
                        cells.append(str(val) if val is not None else "")
                rows_out.append(cells)
            sections_out.append((title, rows_out))

        # Pre-calc total image height
        def _sec_height(data_rows):
            hdr_lines = [_wrap(d[0], f_hdr, d[1] - cp * 2) for d in col_defs]
            hh = max(hh_b, max(len(ls) for ls in hdr_lines) * lh + cp * 2)
            row_hs = []
            for row in data_rows:
                max_l = max(
                    len(_wrap(str(row[ci]), f_data, d[1] - cp * 2) if d[3]
                        else [_clip(str(row[ci]), f_data, d[1] - cp * 2)])
                    for ci, d in enumerate(col_defs) if ci < len(row)
                )
                row_hs.append(max(drow, max_l * lh + cp * 2))
            return sh + hh + sum(row_hs) + S * 2

        footer_h = 30 * S if total_shown >= max_rows else 0
        total_h  = 2 * op + bh + gh
        for _, rows in sections_out:
            total_h += _sec_height(rows) + gh
        total_h += footer_h + (60 * S if not sections_out else 0)

        img  = Image.new("RGB", (img_w, total_h), _D["bg_top"])
        _fill_gradient(img, _D["bg_top"], _D["bg_bot"])
        draw = ImageDraw.Draw(img)

        # Main banner
        draw.rectangle([0, 0, img_w, bh], fill=_D["banner"])
        draw.line([0, bh, img_w, bh], fill=_D["bdr_hdr"], width=S)
        draw.text((op, (bh - 14 * S) // 2),
                  ("Batching Incentive & Deduction Report"
                   + (f"  —  {month_label}" if month_label else "")),
                  fill=_D["txt"], font=f_ban)
        y = bh + gh

        for title, rows in sections_out:
            y = _draw_premium_table(
                draw, title, col_defs, rows, _id_color, op, y,
                f_sec, f_hdr, f_data)
            y += gh

        if total_shown >= max_rows:
            draw.text((op, y + 4 * S),
                      (f"Showing first {max_rows} rows only. "
                       "Full report is in the attached Excel file."),
                      fill=_D["txt_mut"], font=f_foot)

        img.save(output_path, "PNG", dpi=(144, 144))
        return output_path

    except Exception as exc:
        print(f"[email_helper] I&D preview image error: {exc}")
        import traceback; traceback.print_exc()
        return None


# ─────────────────────────────────────────────────────────────────────────────
# EMAIL SENDING
# ─────────────────────────────────────────────────────────────────────────────

def send_report_email(to_emails, cc_emails, subject, body,
                      attachment_bytes=None, attachment_name=None,
                      html_body=None, inline_image_path=None) -> dict:
    """
    Send the report email with optional inline PNG preview and Excel attachment.

    When inline_image_path points to an existing PNG file the message uses a
    multipart/related structure so the image is embedded inline (CID reference).
    html_body must include <img src="cid:report_preview"> to display it.

    Falls back to a plain EmailMessage (no inline image) when the image file
    is missing or image generation failed.

    Returns {"success": bool, "error": str | None}.
    Every attempt is recorded in the email_log table.
    """
    cfg     = get_smtp_config()
    to_list = _split_emails(to_emails)
    cc_list = _split_emails(cc_emails)
    result  = {"success": False, "error": None}

    if not is_configured():
        result["error"] = "SMTP is not configured. Add settings on the Settings page."
    elif not to_list:
        result["error"] = "Please enter at least one 'To' email address."
    else:
        try:
            use_image = bool(inline_image_path and os.path.isfile(inline_image_path))

            if use_image:
                msg = _build_mime_with_inline_image(
                    cfg, to_list, cc_list, subject, body,
                    html_body, inline_image_path,
                    attachment_bytes, attachment_name)
            else:
                # Fallback: standard EmailMessage (no inline image)
                msg = EmailMessage()
                msg["From"]    = cfg["sender"]
                msg["To"]      = ", ".join(to_list)
                if cc_list:
                    msg["Cc"]  = ", ".join(cc_list)
                msg["Subject"] = subject or cfg["subject"]
                msg.set_content(body or "Please find the attached report.")
                if html_body:
                    msg.add_alternative(html_body, subtype="html")
                if attachment_bytes:
                    msg.add_attachment(
                        attachment_bytes,
                        maintype=_XLSX_MAINTYPE, subtype=_XLSX_SUBTYPE,
                        filename=attachment_name or "report.xlsx")

            recipients = to_list + cc_list
            with smtplib.SMTP(cfg["host"], cfg["port"], timeout=30) as server:
                if cfg["use_tls"]:
                    server.starttls()
                server.login(cfg["sender"], cfg["password"])
                server.send_message(msg, to_addrs=recipients)

            result["success"] = True
        except Exception as exc:
            result["error"] = str(exc)

    database.log_email(
        report_file_name=attachment_name or "",
        to_emails=", ".join(to_list),
        cc_emails=", ".join(cc_list),
        subject=subject or cfg["subject"],
        status="Success" if result["success"] else "Failed",
        error_message=result["error"],
    )
    return result


def _build_mime_with_inline_image(cfg, to_list, cc_list, subject, body,
                                  html_body, image_path,
                                  attachment_bytes, attachment_name):
    """
    Build multipart/mixed message:
      mixed
      ├── related
      │   ├── HTML (with <img src="cid:report_preview">)
      │   └── PNG  (Content-ID: <report_preview>, inline)
      └── Excel attachment (if provided)
    """
    from email.mime.multipart import MIMEMultipart
    from email.mime.text      import MIMEText
    from email.mime.image     import MIMEImage
    from email.mime.base      import MIMEBase
    from email                import encoders

    outer = MIMEMultipart("mixed")
    outer["From"]    = cfg["sender"]
    outer["To"]      = ", ".join(to_list)
    if cc_list:
        outer["Cc"] = ", ".join(cc_list)
    outer["Subject"] = subject or cfg["subject"]

    # related = HTML body + inline image
    related = MIMEMultipart("related")
    html = html_body or f"<p>{body or 'Please find the report preview below.'}</p>"
    related.attach(MIMEText(html, "html", "utf-8"))

    with open(image_path, "rb") as fh:
        img_data = fh.read()
    img_part = MIMEImage(img_data, _subtype="png")
    img_part.add_header("Content-ID", "<report_preview>")
    img_part.add_header("Content-Disposition", "inline",
                        filename="report_preview.png")
    related.attach(img_part)
    outer.attach(related)

    # Excel attachment
    if attachment_bytes:
        xlsx = MIMEBase(_XLSX_MAINTYPE, _XLSX_SUBTYPE)
        xlsx.set_payload(attachment_bytes)
        encoders.encode_base64(xlsx)
        xlsx.add_header("Content-Disposition", "attachment",
                        filename=attachment_name or "report.xlsx")
        outer.attach(xlsx)

    return outer
