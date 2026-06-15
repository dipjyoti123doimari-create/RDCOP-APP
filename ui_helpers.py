"""
ui_helpers.py
=============
Controls how the app LOOKS — Stripe-exact dark theme.

Stripe palette:
  #635BFF  purple accent          #0A2540  dark navy base
  #00D4FF  cyan accent            #8898AA  muted secondary text
  rgba(255,255,255,0.04/.08)      frosted glass card bg / border

Typography: Inter (Google Fonts), matching Stripe's clean sans-serif weight.

Reusable helpers:
- inject_custom_css()                  -> applies the whole theme (call once)
- render_page_header(title, subtitle)  -> gradient page title bar
- render_kpi_card(label, value, helper)-> stat card (Dashboard)
- render_glass_card(title, content)    -> frosted-glass content card
- render_success_message(message)      -> green message box
- render_warning_message(message)      -> amber message box
- render_error_message(message)        -> red message box
- render_dynamic_filters(df, cols, ..) -> multi-select filter row
- ProgressTracker class                -> animated step-tracker
"""

import streamlit as st

import config


# ---------------------------------------------------------------------------
# 1. THEME (CSS)
# ---------------------------------------------------------------------------
def inject_custom_css():
    """
    Inject Stripe dark theme CSS + Inter font.
    Call ONCE near the top of app.py, right after st.set_page_config().
    """
    t = config.THEME

    st.markdown(
        # Google Fonts — Inter (matches Stripe's typography)
        '<link rel="preconnect" href="https://fonts.googleapis.com">'
        '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
        '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">',
        unsafe_allow_html=True,
    )

    css = f"""
    <style>
    /* ---- Inter font everywhere ---- */
    html, body, [class*="css"], .stMarkdown, .stText,
    button, input, select, textarea {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont,
                     'Segoe UI', sans-serif !important;
    }}

    /* ---- Transparent page — canvas blob animation shows through ---- */
    .stApp {{
        background: transparent !important;
    }}
    [data-testid="stHeader"] {{
        background: transparent !important;
        border-bottom: 1px solid rgba(255,255,255,0.06) !important;
    }}

    /* ---- Main content area — dark veil so text stays readable over any blob ---- */
    section.main {{
        background: rgba(5, 10, 20, 0.42) !important;
    }}

    /* ---- Main content padding ---- */
    .block-container {{
        padding-top: 2rem !important;
    }}

    /* ---- Force white text on all native Streamlit elements ---- */
    .stMarkdown p, .stMarkdown li, .stMarkdown h1, .stMarkdown h2,
    .stMarkdown h3, .stMarkdown h4,
    [data-testid="stText"], label, .stSelectbox label,
    .stTextInput label, .stNumberInput label, .stTextArea label,
    .stMultiSelect label, .stRadio label, .stCheckbox label,
    .stToggle label, .stDateInput label,
    [data-baseweb="form-control-label"],
    .stSubheader, .stHeader, p {{
        color: rgba(255,255,255,0.92) !important;
    }}

    /* ---- Sidebar — dark frosted glass ---- */
    section[data-testid="stSidebar"] {{
        background: rgba(5, 6, 15, 0.80) !important;
        border-right: 1px solid rgba(255, 255, 255, 0.06) !important;
        backdrop-filter: blur(20px) !important;
        -webkit-backdrop-filter: blur(20px) !important;
    }}
    /* Sidebar text */
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] label {{
        color: rgba(255,255,255,0.75) !important;
    }}

    /* ---- Gradient page-header card ---- */
    .rdc-page-header {{
        background: linear-gradient(135deg, {t['accent_purple']} 0%, #7A73FF 50%, {t['accent_blue']} 100%);
        color: white;
        padding: 28px 32px;
        border-radius: 16px;
        margin-bottom: 24px;
        box-shadow: 0 8px 32px rgba(99, 91, 255, 0.35),
                    0 2px 8px  rgba(0, 0, 0, 0.30);
    }}
    .rdc-page-header h1 {{
        margin: 0;
        font-size: 28px;
        font-weight: 800;
        letter-spacing: -0.02em;
        color: white;
    }}
    .rdc-page-header p {{
        margin: 8px 0 0 0;
        font-size: 15px;
        font-weight: 500;
        opacity: 0.90;
        color: white;
    }}

    /* ---- KPI (stat) card ---- */
    .rdc-kpi-card {{
        background: {t['card_bg']};
        border: 1px solid {t['card_border']};
        border-radius: 16px;
        padding: 20px 22px;
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.25);
        transition: border-color 0.2s ease, box-shadow 0.2s ease;
    }}
    .rdc-kpi-card:hover {{
        border-color: rgba(99, 91, 255, 0.40);
        box-shadow: 0 4px 24px rgba(99, 91, 255, 0.15);
    }}
    .rdc-kpi-label {{
        color: {t['text_secondary']};
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 8px;
    }}
    .rdc-kpi-value {{
        color: {t['text_primary']};
        font-size: 28px;
        font-weight: 700;
        letter-spacing: -0.02em;
    }}
    .rdc-kpi-helper {{
        color: {t['text_secondary']};
        font-size: 12px;
        margin-top: 6px;
    }}

    /* ---- Frosted glass content card ---- */
    .rdc-glass-card {{
        background: {t['card_bg']};
        border: 1px solid {t['card_border']};
        border-radius: 16px;
        padding: 22px 26px;
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.25);
        margin-bottom: 18px;
    }}
    .rdc-glass-card h3 {{
        margin-top: 0;
        color: {t['text_primary']};
        font-weight: 700;
        letter-spacing: -0.01em;
    }}

    /* ---- Coloured status message boxes ---- */
    .rdc-msg {{
        border-radius: 10px;
        padding: 12px 16px;
        margin: 8px 0;
        font-size: 14px;
        font-weight: 500;
        border: 1px solid transparent;
    }}
    .rdc-msg-success {{
        background: {t['success']};
        color: #6EE7B7;
        border-color: rgba(16, 185, 129, 0.25);
    }}
    .rdc-msg-warning {{
        background: {t['warning']};
        color: #FCD34D;
        border-color: rgba(245, 158, 11, 0.25);
    }}
    .rdc-msg-error {{
        background: {t['error']};
        color: #FCA5A5;
        border-color: rgba(239, 68, 68, 0.25);
    }}

    /* ---- Stripe-style buttons ---- */
    .stButton > button {{
        background: linear-gradient(135deg, #635BFF 0%, #7A73FF 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 10px 22px !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        letter-spacing: 0.01em !important;
        box-shadow: 0 4px 16px rgba(99, 91, 255, 0.35) !important;
        transition: all 0.2s ease !important;
    }}
    .stButton > button:hover {{
        background: linear-gradient(135deg, #5851EA 0%, #6D64FF 100%) !important;
        box-shadow: 0 6px 24px rgba(99, 91, 255, 0.50) !important;
        transform: translateY(-1px) !important;
    }}
    .stButton > button:active {{
        transform: translateY(0px) !important;
        box-shadow: 0 2px 8px rgba(99, 91, 255, 0.35) !important;
    }}

    /* ---- Tab bar ---- */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 2px;
        background: rgba(255,255,255,0.04) !important;
        border-radius: 10px !important;
        padding: 4px !important;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 8px !important;
        color: {t['text_secondary']} !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        padding: 8px 16px !important;
    }}
    .stTabs [aria-selected="true"] {{
        background: rgba(99, 91, 255, 0.20) !important;
        color: {t['accent_purple']} !important;
    }}

    /* ---- Metric widget ---- */
    [data-testid="stMetricValue"] {{
        color: white !important;
        font-weight: 700 !important;
    }}
    [data-testid="stMetricLabel"] {{
        color: {t['text_secondary']} !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.06em !important;
    }}

    /* ---- Expander ---- */
    [data-testid="stExpander"] {{
        background: {t['card_bg']} !important;
        border: 1px solid {t['card_border']} !important;
        border-radius: 12px !important;
    }}

    /* ---- Caption / helper text ---- */
    .stCaption, [data-testid="stCaptionContainer"] p,
    [data-testid="stCaptionContainer"] {{
        color: rgba(255,255,255,0.65) !important;
    }}

    /* ---- Info / status text inside st.success / st.info / st.warning ---- */
    [data-testid="stNotification"] p,
    [data-testid="stAlert"] p {{
        color: rgba(255,255,255,0.92) !important;
    }}

    /* ---- Divider ---- */
    hr {{
        border-color: rgba(255,255,255,0.06) !important;
    }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# 2. PAGE HEADER
# ---------------------------------------------------------------------------
def render_page_header(title, subtitle=None):
    """Large gradient header at the top of each page."""
    subtitle_html = f"<p>{subtitle}</p>" if subtitle else ""
    st.markdown(
        f'<div class="rdc-page-header">'
        f'  <h1>{title}</h1>'
        f'  {subtitle_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 3. KPI (STAT) CARD
# ---------------------------------------------------------------------------
def render_kpi_card(label, value, helper_text=None):
    """Small statistic card — place several inside st.columns() for a row."""
    helper_html = f'<div class="rdc-kpi-helper">{helper_text}</div>' if helper_text else ""
    st.markdown(
        f'<div class="rdc-kpi-card">'
        f'  <div class="rdc-kpi-label">{label}</div>'
        f'  <div class="rdc-kpi-value">{value}</div>'
        f'  {helper_html}'
        f'</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 4. GLASS CONTENT CARD
# ---------------------------------------------------------------------------
def render_glass_card(title, content):
    """Frosted-glass card with a title and HTML content."""
    st.markdown(
        f'<div class="rdc-glass-card">'
        f'  <h3>{title}</h3>'
        f'  <div>{content}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 5. STATUS MESSAGES
# ---------------------------------------------------------------------------
def render_success_message(message):
    """Green box — good news."""
    st.markdown(
        f'<div class="rdc-msg rdc-msg-success">✅ {message}</div>',
        unsafe_allow_html=True,
    )


def render_warning_message(message):
    """Amber box — warnings."""
    st.markdown(
        f'<div class="rdc-msg rdc-msg-warning">⚠️ {message}</div>',
        unsafe_allow_html=True,
    )


def render_error_message(message):
    """Red box — errors."""
    st.markdown(
        f'<div class="rdc-msg rdc-msg-error">⛔ {message}</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# 6. DYNAMIC LOV MULTI-SELECT FILTERS
# ---------------------------------------------------------------------------
def render_dynamic_filters(df, filter_columns, key_prefix="flt"):
    """
    Show one multi-select per column and return the filtered DataFrame.

    filter_columns -> list of (column_name, friendly_label) tuples.
    Selecting nothing = show all (no filter applied for that column).
    """
    if df is None or df.empty:
        return df

    filtered = df.copy()

    per_row = 4
    for start in range(0, len(filter_columns), per_row):
        chunk = filter_columns[start:start + per_row]
        slots = st.columns(len(chunk))
        for slot, (col_name, label) in zip(slots, chunk):
            if col_name not in df.columns:
                continue
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
    Animated multi-step progress tracker for upload / sync flows.

    Shows a row of step cards joined by lines.  Each card shows:
      ○  pending   — dashed grey circle
      ◉  active    — purple spinning arc
      ✓  done      — green circle with pop animation
      ✗  error     — red circle

    Usage:
        tracker = ProgressTracker(["📂 Reading", "🧹 Cleaning", "💾 Saving"])
        tracker.start(0)
        df = do_work()
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

    def start(self, idx: int):
        """Mark step idx as actively running."""
        self._states[idx] = "active"
        self._render()

    def complete(self, idx: int, result: str = ""):
        """Mark step idx as done with optional result text."""
        self._states[idx] = "done"
        self._results[idx] = result
        self._render()

    def fail(self, idx: int, result: str = ""):
        """Mark step idx as failed with optional error text."""
        self._states[idx] = "error"
        self._results[idx] = result
        self._render()

    def _render(self):
        parts = []

        for i, (label, state, result) in enumerate(
            zip(self._steps, self._states, self._results)
        ):
            # Connector line between cards
            if i > 0:
                done_prev = self._states[i - 1] == "done"
                line_bg = (
                    "linear-gradient(90deg,#635BFF,#10b981)"
                    if done_prev else
                    "rgba(255,255,255,0.08)"
                )
                parts.append(
                    f'<div style="flex:1;min-width:14px;max-width:52px;height:2px;'
                    f'background:{line_bg};margin-bottom:36px;border-radius:2px;'
                    f'transition:background .5s ease"></div>'
                )

            if state == "active":
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'border:3px solid rgba(99,91,255,0.20);border-top-color:#635BFF;'
                    'animation:rdcSpin .75s linear infinite"></div>'
                )
                nc = "#635BFF"
            elif state == "done":
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'background:linear-gradient(135deg,#10b981,#059669);'
                    'display:flex;align-items:center;justify-content:center;'
                    'color:white;font-size:20px;font-weight:800;'
                    'box-shadow:0 4px 14px rgba(16,185,129,0.40);'
                    'animation:rdcPop .35s cubic-bezier(.175,.885,.32,1.275)">✓</div>'
                )
                nc = "#10b981"
            elif state == "error":
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'background:linear-gradient(135deg,#ef4444,#dc2626);'
                    'display:flex;align-items:center;justify-content:center;'
                    'color:white;font-size:20px;font-weight:800;'
                    'box-shadow:0 4px 14px rgba(239,68,68,0.40)">✗</div>'
                )
                nc = "#ef4444"
            else:  # pending
                icon = (
                    '<div style="width:44px;height:44px;border-radius:50%;'
                    'border:2px dashed rgba(255,255,255,0.15);'
                    'background:rgba(255,255,255,0.04);'
                    'display:flex;align-items:center;justify-content:center;'
                    'color:rgba(255,255,255,0.20);font-size:18px">○</div>'
                )
                nc = "rgba(255,255,255,0.30)"

            parts.append(
                f'<div style="display:flex;flex-direction:column;align-items:center;'
                f'gap:6px;min-width:82px">'
                f'  {icon}'
                f'  <div style="font-size:11px;font-weight:700;color:{nc};'
                f'       text-align:center;line-height:1.3;max-width:78px">{label}</div>'
                f'  <div style="font-size:10px;color:rgba(255,255,255,0.45);'
                f'       text-align:center;min-height:13px;line-height:1.3">{result}</div>'
                f'</div>'
            )

        html = (
            '<style>'
            '@keyframes rdcSpin{to{transform:rotate(360deg)}}'
            '@keyframes rdcPop{'
            '0%{transform:scale(0);opacity:0}'
            '70%{transform:scale(1.2)}'
            '100%{transform:scale(1);opacity:1}}'
            '</style>'
            '<div style="display:flex;align-items:center;justify-content:center;'
            'padding:22px 20px;gap:0;'
            "font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;"
            'background:rgba(255,255,255,0.04);border-radius:16px;'
            'border:1px solid rgba(255,255,255,0.08);'
            'box-shadow:0 4px 24px rgba(0,0,0,0.30);margin:10px 0">'
            + "".join(parts)
            + '</div>'
        )

        self._ph.markdown(html, unsafe_allow_html=True)
