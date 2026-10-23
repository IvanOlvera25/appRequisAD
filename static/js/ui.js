/* AD17 · comportamiento global del layout: tema, menús, drawer y toasts. */
(function () {
  'use strict';
  var root = document.documentElement;

  /* ---------- Tema claro / oscuro ---------- */
  function setTheme(theme, animate) {
    if (animate) {
      root.classList.add('theme-anim');
      setTimeout(function () { root.classList.remove('theme-anim'); }, 450);
    }
    root.setAttribute('data-theme', theme);
    try { localStorage.setItem('ui-theme', theme); } catch (e) {}
    document.dispatchEvent(new CustomEvent('themechange', { detail: { theme: theme } }));
  }
  var themeBtn = document.getElementById('themeToggle');
  if (themeBtn) {
    themeBtn.addEventListener('click', function () {
      setTheme(root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark', true);
    });
  }
  // Si el usuario nunca eligió, seguir los cambios del sistema
  if (window.matchMedia) {
    var mq = matchMedia('(prefers-color-scheme: dark)');
    var follow = function (e) {
      var saved = null;
      try { saved = localStorage.getItem('ui-theme'); } catch (err) {}
      if (saved !== 'light' && saved !== 'dark') root.setAttribute('data-theme', e.matches ? 'dark' : 'light');
    };
    if (mq.addEventListener) mq.addEventListener('change', follow);
  }

  /* ---------- Menús desplegables ---------- */
  var dropdowns = Array.prototype.slice.call(document.querySelectorAll('.nav-dropdown'));
  dropdowns.forEach(function (dd) {
    var toggle = dd.querySelector('.nav-dropdown-toggle');
    if (!toggle) return;
    toggle.setAttribute('aria-haspopup', 'true');
    toggle.addEventListener('click', function (e) {
      e.stopPropagation();
      var wasOpen = dd.classList.contains('open');
      dropdowns.forEach(function (d) { d.classList.remove('open'); });
      if (!wasOpen) dd.classList.add('open');
      toggle.setAttribute('aria-expanded', String(!wasOpen));
    });
  });
  document.addEventListener('click', function () {
    dropdowns.forEach(function (d) { d.classList.remove('open'); });
  });

  /* ---------- Drawer móvil ---------- */
  function openDrawer() { document.body.classList.add('drawer-open'); }
  function closeDrawer() { document.body.classList.remove('drawer-open'); }
  document.querySelectorAll('[data-drawer-open]').forEach(function (b) { b.addEventListener('click', openDrawer); });
  document.querySelectorAll('[data-drawer-close]').forEach(function (b) { b.addEventListener('click', closeDrawer); });
  document.querySelectorAll('.drawer a').forEach(function (a) { a.addEventListener('click', closeDrawer); });

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      closeDrawer();
      dropdowns.forEach(function (d) { d.classList.remove('open'); });
    }
  });

  /* ---------- Toasts ---------- */
  var ICONS = { success: 'fa-circle-check', error: 'fa-circle-exclamation', warning: 'fa-triangle-exclamation', info: 'fa-circle-info' };
  function dismiss(t) {
    if (!t || t.classList.contains('leaving')) return;
    t.classList.add('leaving');
    setTimeout(function () { t.remove(); }, 320);
  }
  function wire(t) {
    var close = t.querySelector('.toast-close');
    if (close) close.addEventListener('click', function () { dismiss(t); });
    var ms = parseInt(t.getAttribute('data-autoclose') || '0', 10);
    if (ms > 0) setTimeout(function () { dismiss(t); }, ms);
  }
  function toast(message, type, ms) {
    type = ICONS[type] ? type : 'info';
    ms = ms == null ? 5000 : ms;
    var stack = document.getElementById('toastStack');
    if (!stack) { alert(message); return; }
    var t = document.createElement('div');
    t.className = 'toast ' + type;
    t.setAttribute('role', 'status');
    t.setAttribute('data-autoclose', String(ms));
    t.innerHTML = '<span class="toast-ico"><i class="fa-solid ' + ICONS[type] + '"></i></span>' +
      '<div class="toast-body"></div><button type="button" class="toast-close" aria-label="Cerrar">&times;</button>' +
      (ms > 0 ? '<span class="toast-progress" style="animation-duration:' + ms + 'ms"></span>' : '');
    t.querySelector('.toast-body').textContent = message;
    stack.appendChild(t);
    wire(t);
    return t;
  }
  document.querySelectorAll('#toastStack .toast').forEach(wire);

  /* ---------- Utilidades compartidas ---------- */
  function money(n) {
    return (Number(n) || 0).toLocaleString('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 2 });
  }
  // Anima un número (KPIs) desde 0 hasta su valor
  function countUp(el) {
    var target = parseFloat(el.getAttribute('data-count')) || 0;
    var isMoney = el.hasAttribute('data-money');
    var reduce = window.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches;
    var fmt = function (v) { return isMoney ? money(v) : Math.round(v).toLocaleString('es-MX'); };
    if (reduce || target === 0) { el.textContent = fmt(target); return; }
    var start = null, dur = 900;
    function step(ts) {
      if (!start) start = ts;
      var p = Math.min((ts - start) / dur, 1);
      var eased = 1 - Math.pow(1 - p, 3);
      el.textContent = fmt(target * eased);
      if (p < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }
  document.querySelectorAll('[data-count]').forEach(countUp);

  window.UI = { toast: toast, money: money, setTheme: setTheme, countUp: countUp };
})();
