/* Felles ikoner og verktoeytips for panelet og forhaandsvisningsvinduet. */
(function (global) {
  'use strict';

  var ICONS = {
    undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
    redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
    refresh: '<path d="M20 11a8 8 0 1 0-2.3 6.2"/><path d="M20 4v7h-7"/>',
    folder: '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    reset: '<path d="M4 12a8 8 0 1 0 2.4-5.7"/><path d="M4 4v5h5"/>',
    pipette: '<path d="m2 22 1-1h3l9-9"/><path d="M3 21v-3l9-9"/><path d="m15 6 3.4-3.4a2.1 2.1 0 1 1 3 3L18 9l.4.4a2.1 2.1 0 1 1-3 3l-3.8-3.8a2.1 2.1 0 1 1 3-3l.4.4z"/>',
    eraser: '<path d="m7 21-4.3-4.3a1 1 0 0 1 0-1.4l10-10a1 1 0 0 1 1.4 0l5.6 5.6a1 1 0 0 1 0 1.4L11 21"/><path d="M22 21H7"/><path d="m5 11 9 9"/>',
    brush: '<path d="m9.1 11.9 8-8.1a2.9 2.9 0 1 1 4.1 4.1l-8.1 8"/><path d="M7.1 14.9c-1.7 0-3 1.4-3 3 0 1.3-2.5 1.5-2 2 1.1 1.1 2.5 2 4 2 2.2 0 4-1.8 4-4a3 3 0 0 0-3-3z"/>',
    eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
    eyeOff: '<path d="m3 3 18 18"/><path d="M10.6 5.1A10 10 0 0 1 12 5c6.5 0 10 7 10 7a17 17 0 0 1-3.2 4.2"/><path d="M6.6 6.6C3.9 8.3 2 12 2 12s3.5 7 10 7a9.7 9.7 0 0 0 5.4-1.6"/><path d="M9.9 9.9a3 3 0 0 0 4.2 4.2"/>',
    wand: '<path d="m15 4 1 2 2 1-2 1-1 2-1-2-2-1 2-1z"/><path d="M3 21 14 10"/><path d="m19 13 .5 1 1 .5-1 .5-.5 1-.5-1-1-.5 1-.5z"/>',
    trash: '<path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="m6 6 1 14h10l1-14"/>',
    download: '<path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/>',
    close: '<path d="M6 6l12 12"/><path d="M18 6 6 18"/>',
    popout: '<path d="M14 4h6v6"/><path d="M20 4l-8 8"/><path d="M18 14v4a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4"/>'
  };

  function setIcon(el, name) {
    var old = el.querySelector('svg.ic');
    if (old) old.remove();
    el.insertAdjacentHTML('afterbegin', '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true">' + ICONS[name] + '</svg>');
  }

  function initIcons() {
    document.querySelectorAll('[data-icon]').forEach(function (el) {
      setIcon(el, el.dataset.icon);
      if (el.dataset.tip && !el.getAttribute('aria-label')) el.setAttribute('aria-label', el.dataset.tip);
    });
  }

  /** Egne verktoeytips: vises med en gang og holdes innenfor panelet. */
  function initTooltips() {
    var tip = document.getElementById('tip'), current = null, timer = null;
    function hide() { clearTimeout(timer); tip.hidden = true; current = null; }
    function show(el) {
      tip.textContent = el.dataset.tip;
      tip.hidden = false;
      var r = el.getBoundingClientRect(), t = tip.getBoundingClientRect();
      var left = Math.min(window.innerWidth - t.width - 6, Math.max(6, r.left + r.width / 2 - t.width / 2));
      var top = r.bottom + 6;
      if (top + t.height > window.innerHeight - 4) top = r.top - t.height - 6;
      tip.style.left = left + 'px'; tip.style.top = top + 'px';
    }
    document.addEventListener('mouseover', function (e) {
      var el = e.target.closest && e.target.closest('[data-tip]');
      if (el === current) return;
      hide();
      if (!el) return;
      current = el;
      timer = setTimeout(function () { if (current === el) show(el); }, 350);
    });
    document.addEventListener('focusin', function (e) {
      var el = e.target.closest && e.target.closest('[data-tip]');
      if (el && e.target.matches(':focus-visible')) { current = el; show(el); }
    });
    document.addEventListener('focusout', hide);
    document.addEventListener('pointerdown', hide);
    window.addEventListener('scroll', hide, true);
  }

  global.UI = { ICONS: ICONS, setIcon: setIcon, initIcons: initIcons, initTooltips: initTooltips };
})(window);
