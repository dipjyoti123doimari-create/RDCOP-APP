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


def compose_report_body(month_label, sections=None, waivers=None) -> str:
    sections = sections or REPORT_SECTIONS
    numbered = "\n".join(f"{i}. {s}" for i, s in enumerate(sections, start=1))
    waiver_note = ""
    if waivers:
        lines = []
        for w in waivers:
            reason_map = {
                "dr_bhoon":       "Waived by Dr. Bhoon",
                "approved_leave": "Waived, on approved leave",
            }
            reason = reason_map.get(w.get("reason", ""), w.get("custom_reason") or w.get("reason", ""))
            lines.append(f"  - {w.get('employee_code','')} : {reason}")
        waiver_note = (
            f"\n\nNote: {len(waivers)} employee(s) have deduction waiver(s) applied this month "
            f"(deduction shown as ₹0 in red row):\n" + "\n".join(lines)
        )
    return (
        "Dear Sir,\n\n"
        f"Please find below the compiled report for the month of {month_label}, "
        "covering the following:\n\n"
        f"{numbered}"
        f"{waiver_note}\n\n"
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


def wrap_html_body_with_image(message_text, excel_attached=True,
                              num_images=1) -> str:
    """
    HTML email body with one or two inline CID images.

    num_images=1  → single <img src="cid:report_preview">   (I&D, TP location-only)
    num_images=2  → two images: cid:report_preview (location) +
                                cid:report_preview_2 (plant)
    """
    import html as _html
    safe = _html.escape(message_text or "").replace("\n", "<br>")
    note = (
        '<p style="color:#555;font-size:12px;margin-top:14px">'
        '&#128206; Full detailed report is attached as an Excel file.</p>'
        if excel_attached else ""
    )
    img_style = ('max-width:100%;height:auto;border-radius:10px;'
                 'display:block;margin:auto;box-shadow:0 2px 8px rgba(0,0,0,0.12)')
    if num_images >= 2:
        imgs = (
            f'<div style="margin:14px 0">'
            f'<img src="cid:report_preview" alt="Location Throughput" style="{img_style}">'
            f'</div>'
            f'<div style="margin:14px 0">'
            f'<img src="cid:report_preview_2" alt="Plant Throughput" style="{img_style}">'
            f'</div>'
        )
    else:
        imgs = (f'<div style="margin:14px 0">'
                f'<img src="cid:report_preview" alt="Report Preview" style="{img_style}">'
                f'</div>')
    return (
        '<!DOCTYPE html><html><body>'
        '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;'
        'color:#0A2540;line-height:1.6;max-width:960px;margin:0 auto;padding:20px">'
        f'<div style="margin-bottom:14px">{safe}</div>'
        f'{imgs}'
        f'{note}'
        '<p style="color:#333;margin-top:14px">Regards</p>'
        '</div></body></html>'
    )


# ─────────────────────────────────────────────────────────────────────────────
# PREMIUM DARK GLOSSY PNG PREVIEW IMAGE (Pillow)
# Dark gradient background + visible spreadsheet-style cell borders +
# glossy dark-tinted conditional row colors (red/amber/green on dark bg).
# White text, high-DPI 2× scale — looks great in Gmail / Outlook.
# ─────────────────────────────────────────────────────────────────────────────

_S    = 2    # 2× Retina scale factor
_OP   = 10   # outer image padding (virtual px)
_CP   = 4    # cell inner padding (virtual px)
_SH   = 24   # section title row height
_HH   = 28   # column header row height
_DROW = 14   # base data row height (compact)
_LH   = 12   # text line height (virtual px)
_GAP  = 6    # gap between table sections

# Frosted-glass palette — light tinted rows on white background, dark navy headers
_D = {
    "bg":       (255, 255, 255),  # white image background
    "sec":      (  8,  47,  84),  # #082F54 section title — dark steel blue
    "hdr":      ( 12,  56, 100),  # #0C3864 column header — slightly lighter steel blue
    "row_odd":  (247, 249, 252),  # very light gray (neutral odd row)
    "row_even": (255, 255, 255),  # white (neutral even row)
    "pan_row":  (218, 232, 250),  # #DAE8FA light blue (PAN India)
    "red":      (255, 210, 214),  # #FFD2D6 frosted rose-red (below threshold)
    "amber":    (255, 238, 190),  # #FEEEBE frosted gold-amber (mid range)
    "green":    (196, 244, 218),  # #C4F4DA frosted mint-green (good/incentive)
    "bdr_out":  ( 60,  90, 130),  # outer table border
    "bdr_hdr":  ( 88, 118, 158),  # header cell dividers
    "bdr_in":   (140, 155, 180),  # inner gridlines — clearly visible on white/light rows
    "txt":      (248, 250, 252),  # white text (dark header/sec elements)
    "txt_dark": ( 25,  38,  58),  # #19263A dark navy text (on light rows)
    "txt_mut":  (100, 116, 139),  # muted footer text
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


def _auto_txt(bg):
    """White text on dark backgrounds, dark text on light backgrounds."""
    r, g, b = bg
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    return _D["txt"] if lum < 140 else _D["txt_dark"]


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
                        x, y, f_sec, f_hdr, f_data, f_bold=None):
    """
    Draw one frosted-glass table section on white background.

    col_defs     = [(label, scaled_px_width, align, wrap_data), ...]
    data_rows    = [[str, ...], ...]
    row_color_fn(ri, row) -> RGB tuple OR (RGB, is_bold bool)
    f_bold       = bold font for special rows (PAN India etc.); falls back to f_data

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
    draw.line([x, y + sh, x1, y + sh], fill=_D["bdr_hdr"], width=S)
    y += sh

    # Column header row
    cx = x
    for (_, w, align, _), lines in zip(col_defs, hdr_lines):
        draw.rectangle([cx, y, cx + w, y + hh], fill=_D["hdr"])
        draw.line([cx + w, y, cx + w, y + hh], fill=_D["bdr_hdr"], width=S)
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
        result  = row_color_fn(ri, row)
        bg, is_bold = (result if isinstance(result, tuple) and len(result) == 2
                       and isinstance(result[1], bool) else (result, False))
        f_cell  = (f_bold or f_data) if is_bold else f_data
        txt_col = _auto_txt(bg)
        cx = x
        for (_, w, align, _), lines in zip(col_defs, cells_lines):
            draw.rectangle([cx, y, cx + w, y + rh], fill=bg)
            draw.line([cx + w, y, cx + w, y + rh], fill=_D["bdr_in"], width=S)
            draw.line([cx, y + rh, cx + w, y + rh], fill=_D["bdr_in"], width=S)
            tby = y + max(cp // 2, (rh - len(lines) * lh) // 2)
            for line in lines:
                tw = _fw(f_cell, line)
                tx = (cx + max(cp, (w - tw) // 2) if align == 'c'
                      else (cx + w - tw - cp if align == 'r' else cx + cp))
                draw.text((tx, tby), line, fill=txt_col, font=f_cell)
                tby += lh
            cx += w
        draw.line([x, y,   x,  y + rh], fill=_D["bdr_out"], width=S)
        draw.line([x1, y,  x1, y + rh], fill=_D["bdr_out"], width=S)
        y += rh

    draw.line([x, y, x1, y], fill=_D["bdr_out"], width=S)
    return y


def _tp_row_color(pct_str, name):
    """Return (bg_color, is_bold) for a TP row based on % and name."""
    if "■" in name or "PAN" in name.upper():
        return (_D["pan_row"], True)   # PAN India → bold
    try:
        pct = float(str(pct_str).replace("%", "").strip())
    except ValueError:
        return (_D["row_odd"], False)
    if pct < 50: return (_D["red"],   False)
    if pct < 70: return (_D["amber"], False)
    return (_D["green"], False)


def _tp_build_image(col_defs_raw, data_rows_out, row_color_fn,
                    sec_title, output_path, footnote=None):
    """
    Internal helper: render a single TP table (no banner) into output_path.
    col_defs_raw = [(label, base_width_px, align, wrap_data), ...]
    Returns output_path on success, None on failure.
    """
    try:
        from PIL import Image, ImageDraw
        S  = _S
        op = _OP * S
        gh = _GAP * S
        cp = _CP  * S
        lh = _LH  * S
        drow = _DROW * S
        hh_b = _HH   * S
        sh   = _SH   * S

        f_sec  = _pil_font(9,  bold=True)
        f_hdr  = _pil_font(7,  bold=True)
        f_data = _pil_font(7,  bold=False)
        f_bold = _pil_font(7,  bold=True)
        f_foot = _pil_font(6,  bold=False)

        col_defs = [(h, w * S, a, wp) for h, w, a, wp in col_defs_raw]
        tbl_w    = sum(d[1] for d in col_defs)
        img_w    = tbl_w + 2 * op

        # Pre-calc table height
        hdr_lines = [_wrap(d[0], f_hdr, d[1] - cp * 2) for d in col_defs]
        hh = max(hh_b, max(len(ls) for ls in hdr_lines) * lh + cp * 2)
        row_hs = []
        for row in data_rows_out:
            max_l = max(
                len(_wrap(str(row[ci]), f_data, d[1] - cp * 2) if d[3]
                    else [_clip(str(row[ci]), f_data, d[1] - cp * 2)])
                for ci, d in enumerate(col_defs) if ci < len(row)
            )
            row_hs.append(max(drow, max_l * lh + cp * 2))
        foot_h = (18 * S if footnote else 0)
        total_h = op + sh + hh + sum(row_hs) + S * 2 + gh + foot_h + op

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
        img  = Image.new("RGB", (img_w, total_h), _D["bg"])
        draw = ImageDraw.Draw(img)

        y = op
        y = _draw_premium_table(draw, sec_title, col_defs, data_rows_out,
                                row_color_fn, op, y, f_sec, f_hdr, f_data, f_bold)
        if footnote:
            y += 4 * S
            draw.text((op, y), footnote, fill=_D["txt_mut"], font=f_foot)

        img.save(output_path, "PNG", dpi=(144, 144))
        return output_path
    except Exception as exc:
        print(f"[email_helper] TP image error ({output_path}): {exc}")
        import traceback; traceback.print_exc()
        return None


def create_tp_location_image(location_rows, month, year, output_path):
    """Generate Location wise Throughput table PNG (no top banner)."""
    import calendar
    mon_tag = f"{calendar.month_abbr[month]}'{str(year)[-2:]}"
    LOC_COLS = [
        ("#",              24, 'c', False),
        ("Exco Location", 110, 'l', False),
        ("Plants",         40, 'c', False),
        ("Total Qty",      70, 'r', False),
        ("Time (min)",     82, 'r', False),
        ("Avg TP %",       54, 'c', False),
    ]
    rows_out = []
    for i, r in enumerate(location_rows, 1):
        pan = r.get("is_pan_india", False)
        pct = round(float(r.get("avg_throughput_pct", 0) or 0))
        rows_out.append([
            "—" if pan else str(i),
            ("■ " if pan else "") + str(r.get("exco_location", "")),
            str(r.get("plant_count", "")),
            _img_fmt(r.get("total_quantity", 0)),
            _img_fmt(r.get("total_time_min", 0)),
            f"{pct}%",
        ])
    def _loc_color(ri, row):
        name = row[1] if len(row) > 1 else ""
        if "■" in name or "PAN" in name.upper():
            return (_D["pan_row"], True)
        return (_D["row_odd"] if ri % 2 == 0 else _D["row_even"], False)
    return _tp_build_image(LOC_COLS, rows_out, _loc_color,
                           f"Location wise Throughput — {mon_tag}",
                           output_path)


def create_tp_plant_image(plant_rows, month, year, output_path):
    """Generate Plant Throughput table PNG — full list, no row limit."""
    import calendar
    mon_tag = f"{calendar.month_abbr[month]}'{str(year)[-2:]}"
    PLT_COLS = [
        ("#",               22, 'c', False),
        ("Plant",          160, 'l', True),
        ("Exco Location",   76, 'l', False),
        ("Business Head",   90, 'l', False),
        ("Total Qty",       68, 'r', False),
        ("Time (min)",      64, 'r', False),
        ("TP %",            46, 'c', False),
    ]
    rows_out = []
    for i, r in enumerate(plant_rows, 1):
        pct = round(float(r.get("throughput_pct", 0) or 0))
        rows_out.append([
            str(i),
            str(r.get("plant_name", "")),
            str(r.get("exco_location", "")),
            str(r.get("business_head", "")),
            _img_fmt(r.get("total_quantity", 0)),
            _img_fmt(r.get("total_time_min", 0)),
            f"{pct}%",
        ])
    def _plt_color(ri, row):
        return _tp_row_color(row[-1], "")
    return _tp_build_image(PLT_COLS, rows_out, _plt_color,
                           f"Plant Throughput — {mon_tag}",
                           output_path)


def create_tp_preview_image(plant_rows, location_rows, month, year,
                            output_path, max_rows=50):
    """
    Generate two TP PNG images (location + plant) and return list of paths.
    output_path base is used to derive _loc.png and _plant.png filenames.
    """
    base     = output_path.replace(".png", "")
    loc_path = base + "_loc.png"
    plt_path = base + "_plant.png"
    paths = []
    r1 = create_tp_location_image(location_rows, month, year, loc_path)
    if r1: paths.append(r1)
    r2 = create_tp_plant_image(plant_rows, month, year, plt_path)
    if r2: paths.append(r2)
    return paths if paths else None


def create_report_preview_image(df, output_path, max_rows=30, month_label=""):
    """
    Generate I&D Incentive & Deduction Report preview PNG — frosted glass style.
    Category-wise sections, white background, colored rows, dark navy headers.
    Returns output_path on success, None on failure.
    """
    try:
        from PIL import Image, ImageDraw
        from report_generator import EMAIL_SECTIONS, _prepare_category_df

        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        S    = _S
        op   = _OP   * S
        gh   = _GAP  * S
        cp   = _CP   * S
        lh   = _LH   * S
        drow = _DROW * S
        hh_b = _HH   * S
        sh   = _SH   * S

        # Columns: (renamed_key, display_header, base_width, align, wrap_data)
        COLS = [
            ("#",               "#",             26,  'c', False),
            ("Employee Name",   "Employee Name", 140, 'l', True),
            ("Category",        "Category",      115, 'l', False),
            ("Plant",           "Plant",          65, 'l', False),
            ("Total Quantity",  "Total Qty",      75, 'r', False),
            ("Incentive Amount","Incentive Amt",  90, 'r', False),
            ("Deduction Amount","Deduction Amt",  90, 'r', False),
            ("Remarks",         "Remarks",         60, 'l', False),
        ]
        col_keys = [c[0] for c in COLS]
        col_defs = [(c[1], c[2] * S, c[3], c[4]) for c in COLS]
        img_w    = sum(d[1] for d in col_defs) + 2 * op

        ded_idx = col_keys.index("Deduction Amount")
        inc_idx = col_keys.index("Incentive Amount")
        rem_idx = col_keys.index("Remarks") if "Remarks" in col_keys else -1

        def _id_color(ri, row):
            try:
                ded = float(str(row[ded_idx]).replace(",", "")) if ded_idx < len(row) else 0
            except ValueError:
                ded = 0
            try:
                inc = float(str(row[inc_idx]).replace(",", "")) if inc_idx < len(row) else 0
            except ValueError:
                inc = 0
            remark = str(row[rem_idx]) if rem_idx >= 0 and rem_idx < len(row) else ""
            is_waived = "waiv" in remark.lower()
            if ded > 0 or is_waived: return (_D["red"],   False)
            if inc > 0: return (_D["green"], False)
            return (_D["row_odd"] if ri % 2 == 0 else _D["row_even"], False)

        f_sec  = _pil_font(10, bold=True)
        f_hdr  = _pil_font(8,  bold=True)
        f_data = _pil_font(8,  bold=False)
        f_foot = _pil_font(7,  bold=False)

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

        # Pre-calc total image height (no banner — starts with first section)
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

        footer_h = 18 * S if total_shown >= max_rows else 0
        total_h  = op
        for _, rows in sections_out:
            total_h += _sec_height(rows) + gh
        total_h += footer_h + op + (60 * S if not sections_out else 0)

        img  = Image.new("RGB", (img_w, total_h), _D["bg"])
        draw = ImageDraw.Draw(img)
        y    = op

        for title, rows in sections_out:
            y = _draw_premium_table(
                draw, title, col_defs, rows, _id_color, op, y,
                f_sec, f_hdr, f_data)
            y += gh

        if total_shown >= max_rows:
            draw.text((op, y + 2 * S),
                      (f"Showing first {max_rows} rows. Full report in attached Excel."),
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
                      html_body=None, inline_image_path=None,
                      extra_inline_image_path=None) -> dict:
    """
    Send the report email with optional inline PNG preview(s) and Excel attachment.

    inline_image_path       → first inline image  (CID: report_preview)
    extra_inline_image_path → second inline image (CID: report_preview_2) — TP plant table

    Falls back to plain EmailMessage when no valid image paths are provided.
    Returns {"success": bool, "error": str | None}.
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
                img2 = (extra_inline_image_path
                        if extra_inline_image_path and os.path.isfile(extra_inline_image_path)
                        else None)
                msg = _build_mime_with_inline_image(
                    cfg, to_list, cc_list, subject, body,
                    html_body, inline_image_path,
                    attachment_bytes, attachment_name,
                    image_path_2=img2)
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
                                  attachment_bytes, attachment_name,
                                  image_path_2=None):
    """
    Build multipart/mixed message:
      mixed
      ├── related
      │   ├── HTML (with <img src="cid:report_preview"> and optionally cid:report_preview_2)
      │   ├── PNG  (Content-ID: <report_preview>)
      │   └── PNG  (Content-ID: <report_preview_2>)  ← optional second image
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

    if image_path_2:
        with open(image_path_2, "rb") as fh:
            img_data2 = fh.read()
        img_part2 = MIMEImage(img_data2, _subtype="png")
        img_part2.add_header("Content-ID", "<report_preview_2>")
        img_part2.add_header("Content-Disposition", "inline",
                             filename="report_preview_2.png")
        related.attach(img_part2)

    outer.attach(related)

    if attachment_bytes:
        xlsx = MIMEBase(_XLSX_MAINTYPE, _XLSX_SUBTYPE)
        xlsx.set_payload(attachment_bytes)
        encoders.encode_base64(xlsx)
        xlsx.add_header("Content-Disposition", "attachment",
                        filename=attachment_name or "report.xlsx")
        outer.attach(xlsx)

    return outer
