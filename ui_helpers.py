"""
ui_helpers.py
=============
This file controls how the app LOOKS (the "theme").

It keeps all the styling (colors, cards, headers, buttons) in one place so
that app.py and the page files stay clean and focused on content.

The most important idea here is "inject_custom_css()". Streamlit lets us send
our own CSS (the language that styles web pages) into the app using
st.markdown(..., unsafe_allow_html=True). We build one big block of CSS using
the colors from config.py and inject it once when the app starts.

Reusable helper functions provided:
- inject_custom_css()                  -> applies the whole theme
- render_page_header(title, subtitle)  -> a nice gradient page title
- render_kpi_card(label, value, helper)-> a small "stat" card (used on Dashboard)
- render_glass_card(title, content)    -> a soft glass-style content card
- render_success_message(message)      -> green message box
- render_warning_message(message)      -> yellow message box
- render_error_message(message)        -> red message box
"""

import streamlit as st

import config  # our colors and app constants live here


# ---------------------------------------------------------------------------
# 1. THE THEME (CSS)
# ---------------------------------------------------------------------------
def inject_custom_css():
    """
    Inject the custom theme CSS into the page.

    Call this ONCE near the top of app.py, right after st.set_page_config().
    """
    t = config.THEME  # short name so the CSS below is easier to read

    css = f"""
    <style>
    /* ---- Overall page background ---- */
    .stApp {{
        background: {t['bg']};
        color: {t['text_primary']};
    }}

    /* ---- Sidebar look ---- */
    section[data-testid="stSidebar"] {{
        background: rgba(255, 255, 255, 0.85);
        border-right: 1px solid {t['card_border']};
    }}

    /* ---- Gradient page header card ---- */
    .rdc-page-header {{
        background: linear-gradient(120deg, {t['accent_purple']}, {t['accent_blue']});
        color: white;
        padding: 26px 30px;
        border-radius: 18px;
        margin-bottom: 22px;
        box-shadow: 0 10px 30px rgba(99, 91, 255, 0.25);
    }}
    .rdc-page-header h1 {{
        margin: 0;
        font-size: 28px;
        font-weight: 700;
        color: white;
    }}
    .rdc-page-header p {{
        margin: 6px 0 0 0;
        font-size: 15px;
        opacity: 0.95;
        color: white;
    }}

    /* ---- KPI (stat) card ---- */
    .rdc-kpi-card {{
        background: {t['card_bg']};
        border: 1px solid {t['card_border']};
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 6px 18px rgba(10, 37, 64, 0.06);
        backdrop-filter: blur(6px);
    }}
    .rdc-kpi-label {{
        color: {t['text_secondary']};
        font-size: 13px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.4px;
        margin-bottom: 6px;
    }}
    .rdc-kpi-value {{
        color: {t['text_primary']};
        font-size: 26px;
        font-weight: 700;
    }}
    .rdc-kpi-helper {{
        color: {t['text_secondary']};
        font-size: 12px;
        margin-top: 4px;
    }}

    /* ---- Soft "glass" content card ---- */
    .rdc-glass-card {{
        background: {t['card_bg']};
        border: 1px solid {t['card_border']};
        border-radius: 16px;
        padding: 20px 24px;
        box-shadow: 0 6px 18px rgba(10, 37, 64, 0.06);
        backdrop-filter: blur(6px);
        margin-bottom: 18px;
    }}
    .rdc-glass-card h3 {{
        margin-top: 0;
        color: {t['text_primary']};
    }}

    /* ---- Coloured status messages ---- */
    .rdc-msg {{
        border-radius: 12px;
        padding: 12px 16px;
        margin: 8px 0;
        font-size: 14px;
        border: 1px solid rgba(10, 37, 64, 0.08);
    }}
    .rdc-msg-success {{ background: {t['success']}; color: #0B6B3A; }}
    .rdc-msg-warning {{ background: {t['warning']}; color: #8A6D00; }}
    .rdc-msg-error   {{ background: {t['error']};   color: #B00020; }}

    /* ---- Transparent / glass-style buttons ---- */
    .stButton > button {{
        background: rgba(99, 91, 255, 0.10);
        color: {t['accent_purple']};
        border: 1px solid rgba(99, 91, 255, 0.35);
        border-radius: 12px;
        padding: 8px 18px;
        font-weight: 600;
        transition: all 0.2s ease;
    }}
    .stButton > button:hover {{
        background: rgba(99, 91, 255, 0.20);
        border-color: {t['accent_purple']};
        color: {t['accent_purple']};
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# 2. PAGE HEADER
# ---------------------------------------------------------------------------
def render_page_header(title, subtitle=None):
    """
    Show a large gradient header at the top of a page.

    title    -> the big heading text
    subtitle -> optional smaller line below the heading
    """
    subtitle_html = f"<p>{subtitle}</p>" if subtitle else ""
    st.markdown(
        f"""
        <div class="rdc-page-header">
            <h1>{title}</h1>
            {subtitle_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 3. KPI (STAT) CARD
# ---------------------------------------------------------------------------
def render_kpi_card(label, value, helper_text=None):
    """
    Show a small statistic card, for example: "Total Employees -> 0".

    Tip: put several of these inside st.columns(...) to lay them out in a row.
    """
    helper_html = f'<div class="rdc-kpi-helper">{helper_text}</div>' if helper_text else ""
    st.markdown(
        f"""
        <div class="rdc-kpi-card">
            <div class="rdc-kpi-label">{label}</div>
            <div class="rdc-kpi-value">{value}</div>
            {helper_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 4. GLASS CONTENT CARD
# ---------------------------------------------------------------------------
def render_glass_card(title, content):
    """
    Show a soft glass-style card with a title and some HTML/text content.
    """
    st.markdown(
        f"""
        <div class="rdc-glass-card">
            <h3>{title}</h3>
            <div>{content}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 5. STATUS MESSAGES (success / warning / error)
# ---------------------------------------------------------------------------
def render_success_message(message):
    """Green box for good news."""
    st.markdown(f'<div class="rdc-msg rdc-msg-success">✅ {message}</div>',
                unsafe_allow_html=True)


def render_warning_message(message):
    """Yellow box for warnings."""
    st.markdown(f'<div class="rdc-msg rdc-msg-warning">⚠️ {message}</div>',
                unsafe_allow_html=True)


def render_error_message(message):
    """Red box for errors."""
    st.markdown(f'<div class="rdc-msg rdc-msg-error">⛔ {message}</div>',
                unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# 6. DYNAMIC LOV MULTI-SELECT FILTERS  (Phase 8)
# ---------------------------------------------------------------------------
def render_dynamic_filters(df, filter_columns, key_prefix="flt"):
    """
    Show one multi-select dropdown per column and return the filtered rows.

    "LOV" = List Of Values. The options inside each dropdown are pulled LIVE
    from the data, so they always match what is actually present (no hardcoding).

    df             -> the pandas DataFrame to filter
    filter_columns -> list of (column_name, friendly_label) tuples, e.g.
                      [("category", "Category"), ("plant", "Plant")]
    key_prefix     -> a unique text prefix so the widget keys never clash
                      with widgets on another page

    Rule: if a dropdown has NOTHING selected, that column is NOT filtered
    (i.e. "show all"). Selecting one or more values keeps only matching rows.
    Returns a NEW filtered DataFrame (the original is left untouched).
    """
    if df is None or df.empty:
        return df

    filtered = df.copy()

    # Lay the dropdowns out neatly in rows of up to 4 side by side.
    per_row = 4
    for start in range(0, len(filter_columns), per_row):
        chunk = filter_columns[start:start + per_row]
        slots = st.columns(len(chunk))
        for slot, (col_name, label) in zip(slots, chunk):
            if col_name not in df.columns:
                continue  # safely skip a column that isn't in the data
            # Build the option list live from the (already filtered? no — full)
            # data so every real value is selectable.
            options = sorted(df[col_name].dropna().astype(str).unique().tolist())
            with slot:
                chosen = st.multiselect(
                    label, options=options, key=f"{key_prefix}_{col_name}"
                )
            if chosen:
                filtered = filtered[filtered[col_name].astype(str).isin(chosen)]

    return filtered


# ---------------------------------------------------------------------------
# 7. ANIMATED PROGRESS TRACKER
# ---------------------------------------------------------------------------
class ProgressTracker:
    """
    Animated multi-step progress bar for upload and sync operations.

    Shows a row of step cards connected by lines. Each card has:
      - A spinning arc   while the step is active
      - A green check ✓  when the step is done  (with a pop animation)
      - A red  cross ✗   if the step failed
      - A grey circle    while the step is pending

    Usage:
        tracker = ProgressTracker(["📂 Reading", "🧹 Cleaning", "💾 Saving"])
        tracker.start(0)
        df = do_some_work()
        tracker.complete(0, f"{len(df)} rows")
        tracker.start(1)
        ...
    """

    def __init__(self, steps: list):
        self._steps   = steps
        self._states  = ["pending"] * len(steps)
        self._results = [""] * len(steps)
        self._ph      = st.empty()
        self._render()

    # -- public API ----------------------------------------------------------

    def start(self, idx: int):
        """Mark step idx as actively running (shows spinner)."""
        self._states[idx] = "active"
        self._render()

    def complete(self, idx: int, result: str = ""):
        """Mark step idx as done (shows green ✓ + optional result text)."""
        self._states[idx] = "done"
        self._results[idx] = result
        self._render()

    def fail(self, idx: int, result: str = ""):
        """Mark step idx as failed (shows red ✗ + optional error text)."""
        self._states[idx] = "error"
        self._results[idx] = result
        self._render()

    # -- internal rendering --------------------------------------------------

    def _render(self):
        parts = []

        for i, (label, state, result) in enumerate(
            zip(self._steps, self._states, self._results)
        ):
            # Connecting line between cards
            if i > 0:
                clr = "#10b981" if self._states[i - 1] == "done" else "#e2e8f0"
                parts.append(
                    f'<div style="flex:1;min-width:14px;max-width:52px;height:2px;'
                    f'background:{clr};margin-bottom:36px;border-radius:2px;'
                    f'transition:background .5s ease"></div>'
                )

            # Icon + label + result card
            if state == "active":
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'border:3px solid rgba(99,91,255,.18);border-top-color:#635BFF;'
                    'animation:rdcSpin .75s linear infinite"></div>'
                )
                nc = "#635BFF"
            elif state == "done":
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'background:linear-gradient(135deg,#10b981,#059669);'
                    'display:flex;align-items:center;justify-content:center;'
                    'color:white;font-size:20px;font-weight:800;'
                    'box-shadow:0 4px 12px rgba(16,185,129,.35);'
                    'animation:rdcPop .35s cubic-bezier(.175,.885,.32,1.275)">✓</div>'
                )
                nc = "#10b981"
            elif state == "error":
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'background:linear-gradient(135deg,#ef4444,#dc2626);'
                    'display:flex;align-items:center;justify-content:center;'
                    'color:white;font-size:20px;font-weight:800;'
                    'box-shadow:0 4px 12px rgba(239,68,68,.35)">✗</div>'
                )
                nc = "#ef4444"
            else:  # pending
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'border:2.5px dashed #e2e8f0;background:#f8fafc;'
                    'display:flex;align-items:center;justify-content:center;'
                    'color:#cbd5e1;font-size:18px">○</div>'
                )
                nc = "#94a3b8"

            parts.append(
                f'<div style="display:flex;flex-direction:column;align-items:center;'
                f'gap:6px;min-width:82px">'
                f'  {icon}'
                f'  <div style="font-size:11px;font-weight:700;color:{nc};'
                f'       text-align:center;line-height:1.3;max-width:78px">{label}</div>'
                f'  <div style="font-size:10px;color:#64748b;text-align:center;'
                f'       min-height:13px;line-height:1.3">{result}</div>'
                f'</div>'
            )

        html = (
            '<style>'
            '@keyframes rdcSpin{to{transform:rotate(360deg)}}'
            '@keyframes rdcPop{'
            '  0%{transform:scale(0);opacity:0}'
            '  70%{transform:scale(1.2)}'
            '  100%{transform:scale(1);opacity:1}}'
            '</style>'
            '<div style="display:flex;align-items:center;justify-content:center;'
            'padding:22px 20px;gap:0;'
            'font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;'
            'background:rgba(255,255,255,.85);border-radius:18px;'
            'border:1px solid rgba(255,255,255,.5);'
            'box-shadow:0 6px 24px rgba(10,37,64,.08);margin:10px 0">'
            + "".join(parts)
            + '</div>'
        )

        self._ph.markdown(html, unsafe_allow_html=True)
