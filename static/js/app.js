/* app.js — RDC Operations App */

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

/* ---- On DOM ready ---- */
document.addEventListener('DOMContentLoaded', function () {

  /* Tab initialization */
  document.querySelectorAll('.tab-btn[data-tabset]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      switchTab(btn.dataset.tabset, btn.dataset.tab);
    });
  });

  /* Topbar theme selector */
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

  /* Flash message auto-dismiss */
  document.querySelectorAll('.alert[data-auto]').forEach(function (el) {
    setTimeout(function () {
      el.style.transition = 'opacity 0.4s';
      el.style.opacity = '0';
      setTimeout(function () { el.remove(); }, 400);
    }, 6000);
  });
});
