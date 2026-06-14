"""
utils/animated_background.py
============================
A reusable, ORIGINAL animated background for the whole app, inspired (visually
only) by the ray-burst hero on stripe.com — no Stripe code/assets are used.

What it draws:
- A fan of thin rays bursting upward from the bottom-center of the screen,
  each tipped with a small dot, over a soft radial gradient.
- The burst gently reacts to the mouse (rays tilt/parallax toward the cursor and
  ease back when the mouse leaves).
- Six time-of-day themes (Pre-dawn, Sunrise, Daytime, Dusk, Sunset, Night) that
  can be picked automatically from the user's local system time, or manually.
- A small theme-toggle button (top-right) to cycle themes by hand.

HOW IT WORKS (important Streamlit detail):
Streamlit's st.markdown() strips <script>, and a components.html() iframe is an
inline element that can't sit *behind* your widgets. So this component renders a
tiny (0-height) iframe whose JavaScript "breaks out" into the PARENT page
(window.parent.document) and appends a fixed, full-screen <canvas> there with
`pointer-events: none`. That makes it a true app-wide background that never
blocks buttons, forms, uploads, filters or tables, while a global mousemove
listener on the parent drives the reactivity.

Public API:
- THEME_NAMES                      -> list of the six theme names
- get_theme_from_time(now=None)    -> Python fallback theme picker
- render_interactive_background(...) -> render the background (call once in app.py)
"""

import json
from datetime import datetime

import streamlit.components.v1 as components


# ---------------------------------------------------------------------------
# 1. THEME PALETTES
# ---------------------------------------------------------------------------
# For each theme:
#   "base" -> [inner glow colour (at the burst origin), outer fill colour]
#   "rays" -> a list of colours the rays are randomly picked from
#   "icon" -> the emoji shown for that theme in the dropdown
THEME_PALETTES = {
    "Pre-dawn": {
        "base": ["#2a2150", "#0e1030"],
        "rays": ["#7c5cff", "#3aa0ff", "#ff7cc8"],
        "icon": "🌌",
    },
    "Sunrise": {
        "base": ["#ffd9a8", "#fff4e6"],
        "rays": ["#ff8fab", "#ff9e3d", "#ffd23d", "#9b6bff"],
        "icon": "🌅",
    },
    "Daytime": {
        "base": ["#cfe8ff", "#ffffff"],
        "rays": ["#3aa0ff", "#00d4ff", "#9b8bff"],
        "icon": "☀️",
    },
    "Dusk": {
        "base": ["#c9b8ff", "#ffe0bf"],
        "rays": ["#8a6bff", "#ff8fc8", "#ffb24d"],
        "icon": "🌆",
    },
    "Sunset": {
        "base": ["#ff8a5c", "#b5179e"],
        "rays": ["#ff2d95", "#9b2cff", "#ff7a3d"],
        "icon": "🌇",
    },
    "Night": {
        "base": ["#101a3a", "#05060f"],
        "rays": ["#3a7bff", "#7c5cff", "#00d4ff", "#ff5cc8"],
        "icon": "🌙",
    },
}

# A fixed order used when the toggle button cycles through themes.
THEME_NAMES = list(THEME_PALETTES.keys())

# How many rays each intensity uses.
INTENSITY_RAYS = {"low": 70, "medium": 120, "high": 180}


# ---------------------------------------------------------------------------
# 2. PYTHON FALLBACK: pick a theme from the time of day
# ---------------------------------------------------------------------------
def get_theme_from_time(now=None):
    """
    Return a theme name based on the given time (defaults to now).

    This is a Python fallback. The component normally detects the user's LOCAL
    browser time in JavaScript, which is more accurate, but this keeps a sane
    default if JavaScript time is ever unavailable.

    Mapping:
        04:00-05:59 Pre-dawn | 06:00-08:59 Sunrise | 09:00-16:59 Daytime
        17:00-18:29 Dusk     | 18:30-20:00 Sunset  | 20:01-03:59 Night
    """
    if now is None:
        now = datetime.now()
    minutes = now.hour * 60 + now.minute  # minutes since midnight

    if 240 <= minutes <= 359:      # 04:00 - 05:59
        return "Pre-dawn"
    if 360 <= minutes <= 539:      # 06:00 - 08:59
        return "Sunrise"
    if 540 <= minutes <= 1019:     # 09:00 - 16:59
        return "Daytime"
    if 1020 <= minutes <= 1109:    # 17:00 - 18:29
        return "Dusk"
    if 1110 <= minutes <= 1200:    # 18:30 - 20:00
        return "Sunset"
    return "Night"                 # 20:01 - 03:59


# ---------------------------------------------------------------------------
# 3. THE BROWSER CODE (HTML + JS that runs inside the tiny iframe)
# ---------------------------------------------------------------------------
# We keep the JavaScript as a normal string with literal { } braces and then
# swap in a few __TOKENS__ with .replace(). (Using .replace instead of an
# f-string avoids having to escape every brace in the JS.)
_BG_TEMPLATE = """
<!doctype html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;overflow:hidden;background:transparent;">
<script>
(function () {
  // ---- Settings sent from Python ----
  var CONFIG = {
    mode: "__MODE__",            // "auto" or "manual"
    theme: "__THEME__",          // used when mode === "manual"
    intensity: "__INTENSITY__",  // "low" | "medium" | "high"
    animate: __ANIMATE__         // true / false
  };
  var PALETTES = __PALETTES__;
  var THEME_ORDER = __THEME_ORDER__;
  var RAY_COUNTS = __RAY_COUNTS__;

  // We must reach the PARENT page to draw a true full-app background.
  var doc, win;
  try {
    win = window.parent;
    doc = window.parent.document;
    // Touch it to make sure we are allowed (throws if cross-origin).
    var _probe = doc.body;
  } catch (e) {
    // If we cannot reach the parent, give up quietly (app still works fine).
    return;
  }

  var reduceMotion = win.matchMedia &&
      win.matchMedia('(prefers-reduced-motion: reduce)').matches;

  // Cap the animation to ~40 frames/sec. The eye barely notices the difference
  // from 60fps for this slow drift, but it cuts the CPU work by a third and
  // keeps the page smooth on modest laptops.
  var FRAME_MS = 1000 / 40;

  function pickThemeByTime() {
    var d = new Date();
    var m = d.getHours() * 60 + d.getMinutes();
    if (m >= 240 && m <= 359)  return "Pre-dawn";
    if (m >= 360 && m <= 539)  return "Sunrise";
    if (m >= 540 && m <= 1019) return "Daytime";
    if (m >= 1020 && m <= 1109) return "Dusk";
    if (m >= 1110 && m <= 1200) return "Sunset";
    return "Night";
  }

  // ---- If the background already exists (Streamlit re-ran), just update it ----
  if (win.__RDC_BG__) {
    win.__RDC_BG__.update(CONFIG);
    return;
  }

  // ---- One-time setup ----
  // CSS: make Streamlit transparent so the canvas shows through, place the
  // canvas behind everything, and style the toggle button.
  var style = doc.createElement('style');
  style.id = 'rdc-bg-style';
  style.textContent =
    '.stApp{background:transparent !important;}' +
    '[data-testid="stHeader"]{background:transparent !important;}' +
    '#rdc-bg-canvas{position:fixed;top:0;left:0;width:100vw;height:100vh;' +
    'z-index:-1;pointer-events:none;}' +
    '#rdc-bg-select{position:fixed;top:62px;right:16px;z-index:9999;' +
    'padding:7px 12px;border-radius:12px;border:1px solid rgba(10,37,64,0.14);' +
    'background:rgba(255,255,255,0.80);color:#0A2540;cursor:pointer;' +
    'font-size:13px;font-weight:600;font-family:inherit;outline:none;' +
    'box-shadow:0 4px 14px rgba(10,37,64,0.14);' +
    '-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px);' +
    'transition:background .2s ease;}' +
    '#rdc-bg-select:hover{background:rgba(255,255,255,0.96);}';
  doc.head.appendChild(style);

  var canvas = doc.createElement('canvas');
  canvas.id = 'rdc-bg-canvas';
  doc.body.appendChild(canvas);
  var ctx = canvas.getContext('2d');

  // Theme LOV (list-of-values dropdown): "Auto" + the six themes.
  var sel = doc.createElement('select');
  sel.id = 'rdc-bg-select';
  sel.title = 'Background theme';
  var optAuto = doc.createElement('option');
  optAuto.value = 'auto';
  optAuto.textContent = '🕒 Auto theme';
  sel.appendChild(optAuto);
  for (var ti = 0; ti < THEME_ORDER.length; ti++) {
    var oName = THEME_ORDER[ti];
    var o = doc.createElement('option');
    o.value = oName;
    o.textContent = ((PALETTES[oName] || {}).icon || '') + ' ' + oName;
    sel.appendChild(o);
  }
  doc.body.appendChild(sel);

  // Keep the dropdown showing the theme that is actually active.
  function syncSelect() {
    if (state.manualOverride) sel.value = state.manualOverride;
    else if (state.cfg.mode === 'auto') sel.value = 'auto';
    else sel.value = state.cfg.theme;
  }

  var state = {
    cfg: CONFIG,
    manualOverride: null,        // set when the user clicks the toggle button
    rays: [],
    mouse: { x: 0.5, y: 0.5 },   // eased position (normalised 0..1)
    target: { x: 0.5, y: 0.5 },  // where the cursor actually is
    W: 0, H: 0, dpr: 1,
    running: false,
    rafId: null,
    lastSig: JSON.stringify(CONFIG),
    lastAuto: null,
    grad: null,        // cached background gradient (rebuilt only on resize/theme)
    baseFill: '#000',  // cached outer fill colour
    lastFrame: 0       // timestamp of the last painted frame (for the FPS cap)
  };

  function currentTheme() {
    if (state.manualOverride) return state.manualOverride;
    if (state.cfg.mode === 'auto') return pickThemeByTime();
    return state.cfg.theme;
  }

  function rayCount() {
    return RAY_COUNTS[state.cfg.intensity] || 120;
  }

  function hexA(hex, a) {
    var h = hex.replace('#', '');
    var r = parseInt(h.substring(0, 2), 16);
    var g = parseInt(h.substring(2, 4), 16);
    var b = parseInt(h.substring(4, 6), 16);
    return 'rgba(' + r + ',' + g + ',' + b + ',' + a + ')';
  }

  function buildRays() {
    var pal = PALETTES[currentTheme()] || PALETTES['Daytime'];
    var n = rayCount();
    var rays = [];
    var spread = Math.PI * 1.1;          // ~198 degree fan
    var start = -Math.PI / 2 - spread / 2;
    for (var i = 0; i < n; i++) {
      var t = n > 1 ? i / (n - 1) : 0.5;
      var angle = start + spread * t + (Math.random() - 0.5) * 0.02;
      var color = pal.rays[Math.floor(Math.random() * pal.rays.length)];
      var op = 0.22 + Math.random() * 0.5;               // opacity
      rays.push({
        baseAngle: angle,
        len: 0.55 + Math.random() * 0.45,                // fraction of max length
        speed: 0.2 + Math.random() * 0.6,                // idle drift speed
        phase: Math.random() * Math.PI * 2,
        // Pre-compute the rgba colour strings ONCE here, so paint() does not
        // have to parse the hex (3x parseInt) for every ray on every frame.
        stroke: hexA(color, op),
        dot: hexA(color, Math.min(1, op + 0.25))
      });
    }
    state.rays = rays;
    buildGradient();   // theme may have changed, so refresh the cached gradient
  }

  // Build the full-screen radial gradient ONCE (per resize / theme change)
  // instead of recreating it on every animation frame — this was the main
  // source of the lag. The origin is fixed at the bottom-centre, so the
  // gradient only needs rebuilding when the size or theme actually changes.
  function buildGradient() {
    var pal = PALETTES[currentTheme()] || PALETTES['Daytime'];
    var ox = state.W / 2, oy = state.H;
    var g = ctx.createRadialGradient(ox, oy, 0, ox, oy, state.H * 1.15);
    g.addColorStop(0, pal.base[0]);
    g.addColorStop(1, pal.base[1]);
    state.grad = g;
    state.baseFill = pal.base[1];
  }

  function resize() {
    // Cap the pixel ratio at 2: on 3x phone screens drawing 9x the pixels every
    // frame is needlessly heavy and adds no visible quality for soft rays.
    state.dpr = Math.min(win.devicePixelRatio || 1, 2);
    state.W = win.innerWidth;
    state.H = win.innerHeight;
    canvas.width = state.W * state.dpr;
    canvas.height = state.H * state.dpr;
    ctx.setTransform(state.dpr, 0, 0, state.dpr, 0, 0);
    buildGradient();   // the gradient depends on size, so refresh it here
  }

  function paint(time) {
    var W = state.W, H = state.H;

    // Soft radial gradient base (cached — see buildGradient). One full-screen
    // fill covers everything, so we no longer need a separate flat fill first.
    if (!state.grad) buildGradient();
    ctx.fillStyle = state.grad;
    ctx.fillRect(0, 0, W, H);

    // Rays.
    var maxLen = Math.sqrt(W * W + H * H) * 0.62;
    var tilt = state.mouse.x - 0.5;      // -0.5 .. 0.5
    var lift = 0.5 - state.mouse.y;      // rays reach higher when cursor is high
    var origX = W / 2 + tilt * 42;       // gentle parallax of the origin
    var origY = H + 8;
    ctx.lineWidth = 1;                   // constant for every ray — set once

    for (var i = 0; i < state.rays.length; i++) {
      var r = state.rays[i];
      var idle = Math.sin(time * r.speed + r.phase) * 0.015;
      var a = r.baseAngle + idle + tilt * 0.28 * Math.cos(r.baseAngle);
      var len = maxLen * r.len * (1 + lift * 0.18);
      var ex = origX + Math.cos(a) * len;
      var ey = origY + Math.sin(a) * len;

      ctx.beginPath();
      ctx.moveTo(origX, origY);
      ctx.lineTo(ex, ey);
      ctx.strokeStyle = r.stroke;        // pre-computed in buildRays
      ctx.stroke();

      ctx.beginPath();
      ctx.arc(ex, ey, 1.5, 0, Math.PI * 2);
      ctx.fillStyle = r.dot;             // pre-computed in buildRays
      ctx.fill();
    }
  }

  function frame(ts) {
    if (!state.running) return;
    // Queue the next frame first, then skip the heavy paint if we are ahead of
    // our ~40fps budget. This throttles CPU use without ever stalling.
    state.rafId = win.requestAnimationFrame(frame);
    if (ts - state.lastFrame < FRAME_MS) return;
    state.lastFrame = ts;

    state.mouse.x += (state.target.x - state.mouse.x) * 0.10;  // smooth easing
    state.mouse.y += (state.target.y - state.mouse.y) * 0.10;
    paint(ts * 0.001);
  }

  function applyRunning() {
    var shouldRun = state.cfg.animate && !reduceMotion;
    if (shouldRun) {
      if (!state.running) {
        state.running = true;
        state.rafId = win.requestAnimationFrame(frame);
      }
    } else {
      state.running = false;
      paint(0);   // draw a single static frame
    }
  }

  // ---- Event listeners (attached once) ----
  doc.addEventListener('mousemove', function (e) {
    state.target.x = e.clientX / state.W;
    state.target.y = e.clientY / state.H;
  });
  doc.addEventListener('mouseleave', function () {
    state.target.x = 0.5;
    state.target.y = 0.5;
  });
  win.addEventListener('resize', function () {
    resize();
    if (!state.running) paint(0);
  });
  sel.addEventListener('change', function () {
    if (sel.value === 'auto') {
      state.manualOverride = null;
      state.cfg.mode = 'auto';        // follow system time again
    } else {
      state.manualOverride = sel.value;
    }
    buildRays();
    if (!state.running) paint(0);
  });

  // ---- Re-check the auto theme every minute so it changes on its own ----
  win.setInterval(function () {
    if (state.cfg.mode === 'auto' && !state.manualOverride) {
      var th = pickThemeByTime();
      if (th !== state.lastAuto) {
        state.lastAuto = th;
        syncSelect();
        buildRays();
        if (!state.running) paint(0);
      }
    }
  }, 60000);

  // ---- Controller used by later Streamlit re-runs ----
  win.__RDC_BG__ = {
    update: function (cfg) {
      var sig = JSON.stringify(cfg);
      if (sig === state.lastSig) return;   // nothing changed (e.g. page switch)
      state.lastSig = sig;
      state.cfg = cfg;
      state.manualOverride = null;          // Settings choices win over the dropdown
      buildRays();
      syncSelect();
      applyRunning();
    }
  };

  // ---- Go ----
  resize();
  buildRays();
  state.lastAuto = pickThemeByTime();
  syncSelect();
  applyRunning();
})();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# 4. PUBLIC RENDER FUNCTION
# ---------------------------------------------------------------------------
def render_interactive_background(mode="auto", theme="Daytime",
                                  intensity="medium", animate=True):
    """
    Render the app-wide animated background. Call this ONCE near the top of
    app.py, on every run (it is safe to call repeatedly — it updates the
    existing background instead of stacking new ones).

    Parameters
    ----------
    mode      : "auto" (theme follows the user's system time) or "manual".
    theme     : which theme to use when mode == "manual"
                (one of THEME_NAMES).
    intensity : "low", "medium" or "high" (controls how many rays are drawn).
    animate   : True to animate; False shows a static gradient + rays.
    """
    # Make sure the values we inject are clean / expected.
    mode = "auto" if mode == "auto" else "manual"
    if theme not in THEME_PALETTES:
        theme = "Daytime"
    intensity = intensity if intensity in INTENSITY_RAYS else "medium"
    animate_js = "true" if animate else "false"

    html = (
        _BG_TEMPLATE
        .replace("__MODE__", mode)
        .replace("__THEME__", theme)
        .replace("__INTENSITY__", intensity)
        .replace("__ANIMATE__", animate_js)
        .replace("__PALETTES__", json.dumps(THEME_PALETTES))
        .replace("__THEME_ORDER__", json.dumps(THEME_NAMES))
        .replace("__RAY_COUNTS__", json.dumps(INTENSITY_RAYS))
    )

    # height=0 -> the iframe itself is invisible; the real canvas lives in the
    # parent page. We render it so its JavaScript runs.
    components.html(html, height=0, width=0)
