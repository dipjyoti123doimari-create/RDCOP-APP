"""
utils/animated_background.py
============================
Stripe-style gradient blob background for the whole app.

Five large, soft radial gradient blobs drift slowly across a dark base,
overlapping with 'screen' blending to produce Stripe's signature glowing
gradient mesh.  Six themes swap the color palette.  The default "Daytime"
theme matches stripe.com's hero: dark navy (#0A2540) + purple + cyan blobs.

HOW IT WORKS:
A 0-height components.html() iframe runs JS that breaks into the parent
Streamlit page (window.parent.document) and appends a fixed full-screen
<canvas> behind all content (pointer-events:none), exactly as before.

Public API (unchanged — keeps full backwards compatibility):
- THEME_NAMES                          -> list of six theme names
- INTENSITY_RAYS                       -> kept for API compat; ignored
- get_theme_from_time(now=None)        -> Python fallback theme picker
- render_interactive_background(...)   -> render the background (call once)
"""

import json
from datetime import datetime

import streamlit.components.v1 as components


# ---------------------------------------------------------------------------
# 1. THEME PALETTES
# ---------------------------------------------------------------------------
# "base"   -> dark fill drawn first each frame (single hex string)
# "colors" -> hex colors for the 5 blobs (list, left-to-right order)
# "icon"   -> emoji for the theme picker dropdown

THEME_PALETTES = {
    "Pre-dawn": {
        "base":   "#050A15",
        "colors": ["#3B3FA0", "#1A0F6B", "#2D1B8E", "#0D1F5C", "#4B3480"],
        "icon":   "🌌",
    },
    "Sunrise": {
        "base":   "#1A0A2E",
        "colors": ["#FF6B9D", "#FF8C42", "#9B4DFF", "#4D79FF", "#FF4DA6"],
        "icon":   "🌅",
    },
    "Daytime": {
        "base":   "#0A2540",
        "colors": ["#635BFF", "#00D4FF", "#7A5AF8", "#0EA5E9", "#8B5CF6"],
        "icon":   "✨",
    },
    "Dusk": {
        "base":   "#0F0820",
        "colors": ["#8A4FFF", "#FF61AB", "#5B21B6", "#7C3AED", "#DB2777"],
        "icon":   "🌆",
    },
    "Sunset": {
        "base":   "#1A0510",
        "colors": ["#FF2D6B", "#9B2CFF", "#FF7A3D", "#FF4DA6", "#C041FF"],
        "icon":   "🌇",
    },
    "Night": {
        "base":   "#05060F",
        "colors": ["#3A7BFF", "#7C5CFF", "#00D4FF", "#5856D6", "#2563EB"],
        "icon":   "🌙",
    },
}

THEME_NAMES = list(THEME_PALETTES.keys())

# Kept for backwards API compatibility — blobs don't use ray counts.
INTENSITY_RAYS = {"low": 70, "medium": 120, "high": 180}


# ---------------------------------------------------------------------------
# 2. PYTHON FALLBACK THEME PICKER
# ---------------------------------------------------------------------------
def get_theme_from_time(now=None):
    """Return a theme name based on the current time of day."""
    if now is None:
        now = datetime.now()
    minutes = now.hour * 60 + now.minute

    if 240 <= minutes <= 359:   return "Pre-dawn"
    if 360 <= minutes <= 539:   return "Sunrise"
    if 540 <= minutes <= 1019:  return "Daytime"
    if 1020 <= minutes <= 1109: return "Dusk"
    if 1110 <= minutes <= 1200: return "Sunset"
    return "Night"


# ---------------------------------------------------------------------------
# 3. HTML + JS TEMPLATE
# ---------------------------------------------------------------------------
# Token substitution keeps JS braces literal — no f-string escaping needed.

_BG_TEMPLATE = """
<!doctype html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;overflow:hidden;background:transparent;">
<script>
(function () {
  var CONFIG = {
    mode:    "__MODE__",
    theme:   "__THEME__",
    animate: __ANIMATE__
  };
  var PALETTES    = __PALETTES__;
  var THEME_ORDER = __THEME_ORDER__;

  // Blob geometry — fixed across all themes; only colors change.
  // bx/by  = base center as fraction of (W, H)
  // ox/oy  = orbital amplitude (fraction of W, H)
  // r      = blob radius as fraction of screen diagonal
  // speed  = orbital angular speed (radians/second)
  //          0.28 rad/s → one full orbit in ~22 seconds (Stripe-speed drift)
  // phase  = initial angle offset so blobs don't all start at the same point
  // alpha  = gradient opacity at the centre
  var GEOM = [
    {bx:0.25, by:0.35, ox:0.18, oy:0.14, r:0.55, speed:0.28, phase:0.00, alpha:0.65},
    {bx:0.72, by:0.28, ox:0.14, oy:0.18, r:0.50, speed:0.22, phase:2.09, alpha:0.60},
    {bx:0.50, by:0.70, ox:0.12, oy:0.10, r:0.58, speed:0.25, phase:4.19, alpha:0.60},
    {bx:0.18, by:0.65, ox:0.09, oy:0.14, r:0.48, speed:0.32, phase:1.05, alpha:0.55},
    {bx:0.82, by:0.72, ox:0.07, oy:0.10, r:0.45, speed:0.19, phase:3.49, alpha:0.50}
  ];

  // Reach into the parent Streamlit page.
  var doc, win;
  try {
    win = window.parent;
    doc = window.parent.document;
    var _probe = doc.body;  // throws if cross-origin
  } catch (e) { return; }

  var reduceMotion = win.matchMedia &&
      win.matchMedia('(prefers-reduced-motion: reduce)').matches;

  // Blobs drift so slowly that 30 fps looks identical to 60 fps.
  var FRAME_MS = 1000 / 30;

  function pickThemeByTime() {
    var d = new Date();
    var m = d.getHours() * 60 + d.getMinutes();
    if (m >= 240  && m <= 359)  return "Pre-dawn";
    if (m >= 360  && m <= 539)  return "Sunrise";
    if (m >= 540  && m <= 1019) return "Daytime";
    if (m >= 1020 && m <= 1109) return "Dusk";
    if (m >= 1110 && m <= 1200) return "Sunset";
    return "Night";
  }

  // On Streamlit re-runs (page switch, widget change) just update CONFIG.
  if (win.__RDC_BG__) { win.__RDC_BG__.update(CONFIG); return; }

  // ---- One-time DOM setup ------------------------------------------------
  var style = doc.createElement('style');
  style.id = 'rdc-bg-style';
  style.textContent =
    // Let the canvas show through Streamlit's wrapper.
    '.stApp{background:transparent !important;}' +
    '[data-testid="stHeader"]{background:transparent !important;}' +
    // Canvas: full-page, behind everything, never blocks clicks.
    '#rdc-bg-canvas{position:fixed;top:0;left:0;width:100vw;height:100vh;' +
      'z-index:-1;pointer-events:none;}' +
    // Theme picker dropdown — dark glass look.
    '#rdc-bg-select{position:fixed;top:62px;right:16px;z-index:9999;' +
      'padding:6px 12px;border-radius:8px;' +
      'border:1px solid rgba(255,255,255,0.10);' +
      'background:rgba(10,37,64,0.80);' +
      'color:rgba(255,255,255,0.85);cursor:pointer;' +
      'font-size:12px;font-weight:600;font-family:inherit;outline:none;' +
      'box-shadow:0 4px 20px rgba(0,0,0,0.40);' +
      '-webkit-backdrop-filter:blur(16px);backdrop-filter:blur(16px);' +
      'transition:background .2s ease;' +
      'appearance:none;-webkit-appearance:none;}' +
    '#rdc-bg-select:hover{background:rgba(10,37,64,0.96);}' +
    '#rdc-bg-select option{background:#0A2540;color:#fff;}';
  doc.head.appendChild(style);

  var canvas = doc.createElement('canvas');
  canvas.id = 'rdc-bg-canvas';
  doc.body.appendChild(canvas);
  var ctx = canvas.getContext('2d');

  // Theme picker: Auto + six named themes.
  var sel = doc.createElement('select');
  sel.id = 'rdc-bg-select';
  sel.title = 'Background theme';
  var optAuto = doc.createElement('option');
  optAuto.value = 'auto';
  optAuto.textContent = '🕒 Auto';
  sel.appendChild(optAuto);
  for (var ti = 0; ti < THEME_ORDER.length; ti++) {
    var oName = THEME_ORDER[ti];
    var o = doc.createElement('option');
    o.value = oName;
    o.textContent = ((PALETTES[oName] || {}).icon || '') + ' ' + oName;
    sel.appendChild(o);
  }
  doc.body.appendChild(sel);

  // ---- State -------------------------------------------------------------
  var state = {
    cfg:           CONFIG,
    manualOverride: null,
    mouse:  { x: 0.5, y: 0.5 },   // eased (smooth) mouse position 0..1
    target: { x: 0.5, y: 0.5 },   // raw mouse position
    W: 0, H: 0, dpr: 1,
    running:   false,
    rafId:     null,
    lastSig:   JSON.stringify(CONFIG),
    lastAuto:  null,
    lastFrame: 0
  };

  function currentTheme() {
    if (state.manualOverride) return state.manualOverride;
    if (state.cfg.mode === 'auto') return pickThemeByTime();
    return state.cfg.theme;
  }

  function syncSelect() {
    if (state.manualOverride)     sel.value = state.manualOverride;
    else if (state.cfg.mode === 'auto') sel.value = 'auto';
    else                          sel.value = state.cfg.theme;
  }

  // Convert "#RRGGBB" + alpha (0-1) to "rgba(r,g,b,a)".
  function hexA(hex, a) {
    var h = hex.replace('#', '');
    var r = parseInt(h.substring(0, 2), 16);
    var g = parseInt(h.substring(2, 4), 16);
    var b = parseInt(h.substring(4, 6), 16);
    return 'rgba(' + r + ',' + g + ',' + b + ',' + a + ')';
  }

  // ---- Core paint --------------------------------------------------------
  function paint(time) {
    var W = state.W, H = state.H;
    var diag = Math.sqrt(W * W + H * H);
    var pal    = PALETTES[currentTheme()] || PALETTES['Daytime'];
    var colors = pal.colors;

    // Dark base fill.
    ctx.globalCompositeOperation = 'source-over';
    ctx.fillStyle = pal.base;
    ctx.fillRect(0, 0, W, H);

    // Blobs rendered with 'screen' blending.  Where two blobs overlap the
    // screen formula brightens the result — producing Stripe's signature
    // glowing highlight at blob intersections.
    ctx.globalCompositeOperation = 'screen';

    for (var i = 0; i < GEOM.length; i++) {
      var g     = GEOM[i];
      var color = colors[i % colors.length];

      // Orbital drift: each blob follows its own slow sine/cosine path.
      var cx = (g.bx + Math.sin(time * g.speed + g.phase) * g.ox) * W;
      var cy = (g.by + Math.cos(time * g.speed * 0.8 + g.phase) * g.oy) * H;

      // Subtle mouse parallax — alternating direction per blob for depth.
      var pxDir = (i % 2 === 0) ? 1 : -1;
      var pyDir = (i % 3 === 0) ? -1 : 1;
      cx += (state.mouse.x - 0.5) * W * 0.05 * pxDir * (i + 1) / GEOM.length;
      cy += (state.mouse.y - 0.5) * H * 0.03 * pyDir * (i + 1) / GEOM.length;

      // Blob radius scales with the screen diagonal for any window size.
      var r = g.r * diag * 0.5;

      // Soft radial gradient: opaque at center, transparent at edge.
      var grd = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
      grd.addColorStop(0,   hexA(color, g.alpha));
      grd.addColorStop(0.5, hexA(color, g.alpha * 0.45));
      grd.addColorStop(1,   hexA(color, 0));

      ctx.fillStyle = grd;
      ctx.fillRect(0, 0, W, H);
    }

    ctx.globalCompositeOperation = 'source-over';
  }

  function resize() {
    state.dpr = Math.min(win.devicePixelRatio || 1, 2);
    state.W   = win.innerWidth;
    state.H   = win.innerHeight;
    canvas.width  = state.W * state.dpr;
    canvas.height = state.H * state.dpr;
    ctx.setTransform(state.dpr, 0, 0, state.dpr, 0, 0);
  }

  function frame(ts) {
    if (!state.running) return;
    state.rafId = win.requestAnimationFrame(frame);
    if (ts - state.lastFrame < FRAME_MS) return;   // 30 fps cap
    state.lastFrame = ts;
    // Ease mouse toward target — prevents jittery movement.
    state.mouse.x += (state.target.x - state.mouse.x) * 0.08;
    state.mouse.y += (state.target.y - state.mouse.y) * 0.08;
    paint(ts * 0.001);  // pass seconds to paint()
  }

  function applyRunning() {
    var shouldRun = state.cfg.animate && !reduceMotion;
    if (shouldRun) {
      if (!state.running) {
        state.running = true;
        state.rafId   = win.requestAnimationFrame(frame);
      }
    } else {
      state.running = false;
      paint(0);   // static frame when animation is off
    }
  }

  // ---- Event listeners ---------------------------------------------------
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
      state.cfg.mode = 'auto';
    } else {
      state.manualOverride = sel.value;
    }
    if (!state.running) paint(0);
  });

  // Auto-switch theme every minute when in auto mode.
  win.setInterval(function () {
    if (state.cfg.mode === 'auto' && !state.manualOverride) {
      var th = pickThemeByTime();
      if (th !== state.lastAuto) {
        state.lastAuto = th;
        syncSelect();
        if (!state.running) paint(0);
      }
    }
  }, 60000);

  // ---- Controller (called by Streamlit re-runs via win.__RDC_BG__) -------
  win.__RDC_BG__ = {
    update: function (cfg) {
      var sig = JSON.stringify(cfg);
      if (sig === state.lastSig) return;   // nothing changed
      state.lastSig      = sig;
      state.cfg          = cfg;
      state.manualOverride = null;
      syncSelect();
      applyRunning();
    }
  };

  // ---- Go ----------------------------------------------------------------
  resize();
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
    Render the app-wide animated gradient blob background.

    Parameters
    ----------
    mode      : "auto" (follows local time) or "manual".
    theme     : active theme when mode == "manual" (one of THEME_NAMES).
    intensity : kept for API compatibility — ignored by the blob animation.
    animate   : True animates; False shows a static gradient snapshot.
    """
    mode       = "auto" if mode == "auto" else "manual"
    theme      = theme if theme in THEME_PALETTES else "Daytime"
    animate_js = "true" if animate else "false"

    html = (
        _BG_TEMPLATE
        .replace("__MODE__",        mode)
        .replace("__THEME__",       theme)
        .replace("__ANIMATE__",     animate_js)
        .replace("__PALETTES__",    json.dumps(THEME_PALETTES))
        .replace("__THEME_ORDER__", json.dumps(THEME_NAMES))
    )

    # height=0  -> iframe is invisible; canvas lives in the parent page.
    components.html(html, height=0, width=0)
