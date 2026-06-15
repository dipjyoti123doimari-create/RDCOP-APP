/* app.js — RDC-OP by DJ */

/* ---- Cursor bloom (Stripe-like background glow) ---- */
(function () {
  var root = document.documentElement;
  document.addEventListener('mousemove', function (e) {
    root.style.setProperty('--cx', e.clientX + 'px');
    root.style.setProperty('--cy', e.clientY + 'px');
  });
}());

/* ---- Toast ---- */
function showToast(msg, type) {
  var c = document.getElementById('toast-container');
  if (!c) return;
  var t = document.createElement('div');
  t.className = 'toast ' + (type || 'info');
  t.textContent = msg;
  c.appendChild(t);
  setTimeout(function () {
    t.style.transition = 'all 0.28s';
    t.style.opacity = '0';
    t.style.transform = 'translateX(110%)';
    setTimeout(function () { t.remove(); }, 300);
  }, 4500);
}

/* ---- Tabs ---- */
function switchTab(set, id) {
  document.querySelectorAll('[data-tabset="' + set + '"][data-tab]').forEach(function (b) {
    b.classList.toggle('active', b.dataset.tab === id);
  });
  document.querySelectorAll('[data-tabpanel="' + set + '"]').forEach(function (p) {
    p.classList.toggle('active', p.dataset.tab === id);
  });
}

/* ---- File drop zone ---- */
function initDrop(dropId, inputId, labelId) {
  var drop = document.getElementById(dropId);
  var inp  = document.getElementById(inputId);
  var lbl  = labelId ? document.getElementById(labelId) : null;
  if (!drop || !inp) return;
  drop.addEventListener('click', function () { inp.click(); });
  drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragleave', function () { drop.classList.remove('drag-over'); });
  drop.addEventListener('drop', function (e) {
    e.preventDefault(); drop.classList.remove('drag-over');
    if (e.dataTransfer.files.length) {
      var dt = new DataTransfer(); dt.items.add(e.dataTransfer.files[0]);
      inp.files = dt.files; inp.dispatchEvent(new Event('change', { bubbles: true }));
    }
  });
  inp.addEventListener('change', function () {
    if (inp.files[0] && lbl) lbl.textContent = '📄 ' + inp.files[0].name;
  });
}

/* ---- AJAX helper ---- */
function apiFetch(url, opts) {
  opts = opts || {};
  var btn = opts.btn; var load = opts.load; var ok = opts.ok || function(){}; var err = opts.err;
  if (btn) btn.disabled = true;
  if (load) load.classList.add('active');
  var fo = { method: opts.method || 'POST', headers: {} };
  if (opts.form) { fo.body = new FormData(opts.form); }
  else if (opts.data) { fo.headers['Content-Type'] = 'application/json'; fo.body = JSON.stringify(opts.data); }
  fetch(url, fo)
    .then(function (r) { return r.json(); })
    .then(function (j) {
      if (btn) btn.disabled = false;
      if (load) load.classList.remove('active');
      if (j.error) { (err || function (m) { showToast(m, 'error'); })(j.error); }
      else ok(j);
    })
    .catch(function (e) {
      if (btn) btn.disabled = false;
      if (load) load.classList.remove('active');
      (err || function (m) { showToast(m, 'error'); })('Network error: ' + e.message);
    });
}

/* ---- Confirm dialog ---- */
function confirmAction(msg, callback) {
  if (window.confirm(msg)) callback();
}

/* ---- Custom multi-select ---- */
function msInit() {
  document.querySelectorAll('.ms-wrap').forEach(function (wrap) {
    var box      = wrap.querySelector('.ms-box');
    var panel    = wrap.querySelector('.ms-panel');
    var allCb    = wrap.querySelector('.ms-all-cb');
    var searchIn = wrap.querySelector('.ms-search-box input');
    var opts     = wrap.querySelectorAll('.ms-opts input[type=checkbox]');

    function updateLabel() {
      var checked = Array.from(opts).filter(function (o) { return o.checked; });
      var lbl = box.querySelector('.ms-label');
      if (checked.length === 0) {
        lbl.textContent = 'All';
        lbl.classList.add('placeholder');
      } else if (checked.length === opts.length) {
        lbl.textContent = 'All (' + opts.length + ')';
        lbl.classList.remove('placeholder');
      } else {
        lbl.textContent = checked.length + ' of ' + opts.length + ' selected';
        lbl.classList.remove('placeholder');
      }
      if (allCb) {
        allCb.checked       = checked.length === opts.length && opts.length > 0;
        allCb.indeterminate = checked.length > 0 && checked.length < opts.length;
      }
    }

    /* open / close */
    box.addEventListener('click', function (e) {
      e.stopPropagation();
      var wasOpen = panel.classList.contains('open');
      /* close all other open panels */
      document.querySelectorAll('.ms-panel.open').forEach(function (p) {
        p.classList.remove('open');
        p.closest('.ms-wrap').querySelector('.ms-box').classList.remove('open');
      });
      if (!wasOpen) {
        panel.classList.add('open');
        box.classList.add('open');
        if (searchIn) { searchIn.value = ''; msShowAll(wrap); searchIn.focus(); }
      }
    });

    /* keep panel open when clicking inside */
    panel.addEventListener('click', function (e) { e.stopPropagation(); });

    /* select all toggle */
    if (allCb) {
      allCb.addEventListener('change', function () {
        wrap.querySelectorAll('.ms-opts label:not(.ms-hidden) input[type=checkbox]').forEach(function (o) {
          o.checked = allCb.checked;
        });
        updateLabel();
      });
    }

    /* search filter */
    if (searchIn) {
      searchIn.addEventListener('input', function () {
        var q = searchIn.value.toLowerCase();
        wrap.querySelectorAll('.ms-opts label').forEach(function (lbl) {
          var val = lbl.textContent.trim().toLowerCase();
          lbl.classList.toggle('ms-hidden', q !== '' && !val.includes(q));
        });
        updateLabel();
      });
    }

    /* individual checkbox changes */
    opts.forEach(function (o) { o.addEventListener('change', updateLabel); });

    updateLabel();
  });

  /* global: click outside closes all panels */
  document.addEventListener('click', function () {
    document.querySelectorAll('.ms-panel.open').forEach(function (p) {
      p.classList.remove('open');
      p.closest('.ms-wrap').querySelector('.ms-box').classList.remove('open');
    });
  });
}

function msShowAll(wrap) {
  wrap.querySelectorAll('.ms-opts label').forEach(function (l) { l.classList.remove('ms-hidden'); });
}

/* ---- On DOM ready ---- */
document.addEventListener('DOMContentLoaded', function () {

  /* Custom multi-selects */
  msInit();

  /* Tab initialization */
  document.querySelectorAll('.tab-btn[data-tabset]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      switchTab(btn.dataset.tabset, btn.dataset.tab);
    });
  });

  /* Topbar theme selector (hidden native, driven by custom picker) */
  var sel = document.getElementById('topbar-theme');
  if (sel) {
    sel.addEventListener('change', function () {
      if (window.setBgTheme) setBgTheme(this.value, false);
      fetch('/api/bg-settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ manual_theme: this.value, auto_theme: false })
      });
    });
  }

  /* Custom theme picker open/close */
  document.addEventListener('click', function(e) {
    var picker = document.getElementById('theme-picker');
    if (!picker) return;
    var panel = document.getElementById('theme-picker-panel');
    if (!picker.contains(e.target)) panel.style.display = 'none';
  });

  /* Flash message auto-dismiss */
  document.querySelectorAll('.alert[data-auto]').forEach(function (el) {
    setTimeout(function () {
      el.style.transition = 'opacity 0.4s';
      el.style.opacity = '0';
      setTimeout(function () { el.remove(); }, 400);
    }, 6000);
  });

  /* Live Oracle status pill (launcher + every module topbar).
     Runs async so the page paints instantly and the dot updates when the
     reachability check returns. Shared connection across all four modules. */
  initOracleStatus();
});

/* ---- Live Oracle status (shared across all modules) ---- */
function initOracleStatus() {
  var pill = document.getElementById('ora-status');
  if (!pill) return;
  var txt = document.getElementById('ora-status-text');
  fetch('/api/oracle-status', { cache: 'no-store' })
    .then(function (r) { return r.json(); })
    .then(function (d) {
      pill.classList.remove('checking', 'connected', 'unreachable', 'unconfigured');
      pill.classList.add(d.state);
      pill.title = d.label;
      if (txt) {
        txt.textContent = d.state === 'connected' ? 'Oracle' :
                          d.state === 'unreachable' ? 'Oracle ✕' :
                          d.state === 'unconfigured' ? 'Oracle —' : 'Oracle…';
      }
    })
    .catch(function () {
      pill.classList.remove('checking');
      pill.classList.add('unreachable');
      pill.title = 'Oracle status check failed';
      if (txt) txt.textContent = 'Oracle ✕';
    });
}

/* ---- Custom theme picker ---- */
var _THEME_ICONS = {
  'Pre-dawn':'🌑','Sunrise':'🌅','Daytime':'☀️','Dusk':'🌆','Sunset':'🌇','Night':'🌙'
};
var _THEME_CLS = {
  'Pre-dawn':'ti-predawn','Sunrise':'ti-sunrise','Daytime':'ti-daytime','Dusk':'ti-dusk','Sunset':'ti-sunset','Night':''
};
function toggleThemePicker() {
  var panel = document.getElementById('theme-picker-panel');
  if (!panel) return;
  panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
}
function selectTheme(t) {
  var panel  = document.getElementById('theme-picker-panel');
  var icon   = document.getElementById('theme-icon-display');
  var label  = document.getElementById('theme-label-display');
  var sel    = document.getElementById('topbar-theme');
  if (panel)  panel.style.display = 'none';
  if (label)  label.textContent = t;
  if (icon) {
    var cls = _THEME_CLS[t] || '';
    icon.innerHTML = '<span class="' + cls + '">' + (_THEME_ICONS[t] || '') + '</span>';
    if (t === 'Night') icon.innerHTML += '<span class="star-sparkle">✦</span>';
  }
  /* Update active state in dropdown */
  document.querySelectorAll('.theme-picker-opt').forEach(function(o) {
    o.classList.toggle('active', o.getAttribute('onclick') === "selectTheme('" + t + "')");
  });
  /* Fire the hidden select so existing API call triggers */
  if (sel) { sel.value = t; sel.dispatchEvent(new Event('change')); }
}
