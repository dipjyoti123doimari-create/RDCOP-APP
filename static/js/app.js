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

/* ---- Modal portal: move all modals to <body> so position:fixed works
        regardless of parent transforms/filters (pageContentIn animation) ---- */
document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('.modal-overlay').forEach(function (m) {
    document.body.appendChild(m);
  });

  /* Inject [spinner+text row | % counter] + progress bar into every .loading */
  document.querySelectorAll('.loading').forEach(function (el) {
    /* wrap existing spinner + text in a row div */
    var row = document.createElement('div');
    row.className = 'loading-row';
    while (el.firstChild) row.appendChild(el.firstChild);
    var pct = document.createElement('span');
    pct.className = 'loading-pct';
    row.appendChild(pct);
    el.appendChild(row);

    /* bar below the row */
    var bar  = document.createElement('div');  bar.className  = 'loading-bar';
    var fill = document.createElement('div');  fill.className = 'loading-bar-fill';
    bar.appendChild(fill);
    el.appendChild(bar);

    /* drive width + percentage via setInterval when .active is toggled */
    var timer = null; var p = 0;
    new MutationObserver(function () {
      if (el.classList.contains('active')) {
        p = 0; fill.style.width = '0%'; pct.textContent = '0%';
        clearInterval(timer);
        timer = setInterval(function () {
          p = Math.min(87, p + Math.max(0.3, (87 - p) * 0.05));
          fill.style.width = p.toFixed(1) + '%';
          pct.textContent  = Math.floor(p) + '%';
        }, 100);
      } else {
        clearInterval(timer);
        fill.style.width = '0%';
        pct.textContent  = '';
      }
    }).observe(el, { attributes: true, attributeFilter: ['class'] });
  });
});

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

/* ---- Searchable SELECT (LOV combobox) ---- */
function initSearchableSelects() {
  document.querySelectorAll('select[data-searchable]').forEach(function (sel) {
    var allOpts = Array.from(sel.options).map(function (o) {
      return { value: o.value, text: o.text };
    });

    /* hide native select; keep in DOM for form submission */
    sel.style.display = 'none';

    /* wrapper */
    var wrap = document.createElement('div');
    wrap.className = 'ss-wrap';
    sel.parentNode.insertBefore(wrap, sel);
    wrap.appendChild(sel);

    /* trigger button */
    var btn = document.createElement('div');
    btn.className = 'ss-btn';
    btn.setAttribute('tabindex', '0');
    btn.innerHTML = '<span class="ss-label placeholder">— select —</span><span class="ss-arrow">▾</span>';
    wrap.insertBefore(btn, sel);
    var labelEl = btn.querySelector('.ss-label');

    /* panel (portalled to body on open) */
    var panel = document.createElement('div');
    panel.className = 'ss-panel';
    var searchRow = document.createElement('div');
    searchRow.className = 'ss-search-row';
    var inp = document.createElement('input');
    inp.type = 'text'; inp.className = 'ss-search';
    inp.placeholder = '🔍 Search…'; inp.autocomplete = 'off';
    searchRow.appendChild(inp);
    var list = document.createElement('div');
    list.className = 'ss-list';
    panel.appendChild(searchRow);
    panel.appendChild(list);

    /* sync button label from current select value */
    function syncLabel() {
      var cur = allOpts.find(function (o) { return o.value === sel.value; });
      if (cur && cur.value !== '') {
        labelEl.textContent = cur.text;
        labelEl.classList.remove('placeholder');
      } else {
        labelEl.textContent = '— select —';
        labelEl.classList.add('placeholder');
      }
    }
    syncLabel();

    /* render filtered list */
    function renderList(q) {
      q = (q || '').toLowerCase().trim();
      list.innerHTML = '';
      allOpts.forEach(function (o) {
        if (q && o.value !== '' &&
            !o.text.toLowerCase().includes(q) &&
            !o.value.toLowerCase().includes(q)) return;
        var item = document.createElement('div');
        item.className = 'ss-item' +
          (o.value === '' ? ' ss-placeholder' : '') +
          (o.value === sel.value ? ' ss-active' : '');
        item.textContent = o.text;
        item.addEventListener('mousedown', function (e) {
          e.preventDefault();
          sel.value = o.value;
          syncLabel();
          closePanel();
          sel.dispatchEvent(new Event('change', { bubbles: true }));
        });
        list.appendChild(item);
      });
    }

    var isOpen = false;
    function openPanel() {
      isOpen = true;
      btn.classList.add('open');
      inp.value = '';
      renderList('');
      /* position relative to button using fixed coords */
      var r = btn.getBoundingClientRect();
      panel.style.top    = (r.bottom + 4) + 'px';
      panel.style.left   = r.left + 'px';
      panel.style.width  = r.width + 'px';
      document.body.appendChild(panel);
      inp.focus();
    }
    function closePanel() {
      isOpen = false;
      btn.classList.remove('open');
      if (panel.parentNode) panel.parentNode.removeChild(panel);
    }

    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      if (isOpen) closePanel(); else openPanel();
    });
    btn.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); if (isOpen) closePanel(); else openPanel(); }
      if (e.key === 'Escape') closePanel();
    });
    panel.addEventListener('click', function (e) { e.stopPropagation(); });
    inp.addEventListener('input', function () { renderList(inp.value); });
    document.addEventListener('click', function () { if (isOpen) closePanel(); });
  });
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

  /* Searchable LOV selects */
  initSearchableSelects();

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
  /* Picking a theme means manual mode — turn Auto off. */
  var autoCb = document.getElementById('bg-auto-cb');
  if (autoCb) autoCb.checked = false;
  /* Fire the hidden select so existing API call triggers */
  if (sel) { sel.value = t; sel.dispatchEvent(new Event('change')); }
}

/* Auto-theme + animation toggles — full background control from the top-right
   picker (replaces the old Settings → Animated Background modal). */
function toggleBgAuto(cb) {
  var on = cb.checked;
  var label = document.getElementById('theme-label-display');
  var t = label ? label.textContent.trim() : 'Daytime';
  if (window.setBgTheme) setBgTheme(t, on);
  fetch('/api/bg-settings', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ auto_theme: on })
  });
}
function toggleBgAnimate(cb) {
  var on = cb.checked;
  if (window.setBgAnimate) setBgAnimate(on);
  fetch('/api/bg-settings', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ animate: on })
  });
}
