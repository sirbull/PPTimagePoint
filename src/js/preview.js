/*
 * Stort forhaandsvisningsvindu (Office-dialog). Viser bildet stort og lar deg
 * velge bakgrunnsfarge og tegne med penselen. Alt sendes til panelet, som eier
 * tilstanden og sender oppdatert tilstand tilbake.
 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var CHUNK = 30000;
  var inbox = {};
  var pv = null;          // Pipeline.Previewer
  var state = null;       // siste tilstand fra panelet
  var local = { tool: 'none', pickMode: 'set', brush: 30, hardness: 65 };
  var drawing = null, compare = false, raf = false;

  /* ---------- Meldinger ---------- */

  function send(obj) {
    var str = JSON.stringify(obj);
    try {
      if (str.length <= CHUNK) { Office.context.ui.messageParent(str); return; }
      var id = Math.random().toString(36).slice(2), n = Math.ceil(str.length / CHUNK);
      for (var i = 0; i < n; i++) Office.context.ui.messageParent(JSON.stringify({ t: 'chunk', id: id, i: i, n: n, d: str.slice(i * CHUNK, (i + 1) * CHUNK) }));
    } catch (e) { /* panelet er borte */ }
  }

  function receive(raw) {
    var m;
    try { m = JSON.parse(raw); } catch (e) { return; }
    if (m.t === 'chunk') {
      var b = inbox[m.id] = inbox[m.id] || { n: m.n, got: 0, parts: [] };
      if (b.parts[m.i] == null) { b.parts[m.i] = m.d; b.got++; }
      if (b.got < b.n) return;
      delete inbox[m.id];
      try { m = JSON.parse(b.parts.join('')); } catch (e) { return; }
    }
    if (m.t === 'image') loadImage(m.url);
    else if (m.t === 'state') applyState(m);
  }

  function loadImage(url) {
    var img = new Image();
    img.onload = function () {
      var c = document.createElement('canvas');
      c.width = img.naturalWidth; c.height = img.naturalHeight;
      var x = c.getContext('2d', { willReadFrequently: true });
      x.drawImage(img, 0, 0);
      pv = new Pipeline.Previewer(x.getImageData(0, 0, c.width, c.height));
      var v = $('pvView');
      v.width = c.width; v.height = c.height;
      v.hidden = false; $('pvWait').hidden = true;
      draw();
    };
    img.src = url;
  }

  function applyState(m) {
    state = m.s;
    local.tool = m.tool; local.pickMode = m.pickMode || 'set'; local.brush = m.brush; local.hardness = m.hardness;
    if (!drawing) syncTools();
    $('pvSize').value = m.brush; $('pvSizeOut').textContent = m.brush;
    $('pvHard').value = m.hardness; $('pvHardOut').textContent = m.hardness;
    $('pvUndo').disabled = !m.canUndo; $('pvRedo').disabled = !m.canRedo;
    $('pvLive').textContent = m.live || '';
    draw();
  }

  /* ---------- Tegning ---------- */

  function draw() {
    if (raf) return;
    raf = true;
    requestAnimationFrame(function () {
      raf = false;
      if (!pv || !state) return;
      var st = state;
      if (drawing) st = Object.assign({}, state, { strokes: state.strokes.concat([drawing]) });
      pv.draw($('pvView').getContext('2d'), st, compare);
    });
  }

  function syncTools() {
    document.querySelectorAll('[data-tool]').forEach(function (b) { b.classList.toggle('on', b.dataset.tool === local.tool); });
    var brush = local.tool === 'erase' || local.tool === 'restore';
    $('pvBrush').hidden = !brush;
    $('pvStage').classList.toggle('tool-erase', local.tool === 'erase');
    $('pvStage').classList.toggle('tool-restore', local.tool === 'restore');
    $('pvStage').classList.toggle('tool-none', local.tool === 'none');
    $('pvStage').classList.toggle('tool-pick', local.tool === 'pick');
    $('pvPick').hidden = local.tool !== 'pick';
    ['pm-set', 'pm-add', 'pm-sub'].forEach(function (c) { $('pvStage').classList.toggle(c, c === 'pm-' + local.pickMode); });
    document.querySelectorAll('[data-pickmode]').forEach(function (b) { b.classList.toggle('on', b.dataset.pickmode === local.pickMode); });
    $('pvHint').textContent = local.tool === 'erase' ? 'Paint over what should become transparent.'
      : local.tool === 'restore' ? 'Paint back parts of the subject.'
      : local.tool === 'pick' ? 'Click the background color to remove it. Shift+click adds more colors.'
      : 'Choose a tool to edit the picture.';
    if (!brush) $('pvCursor').hidden = true;
  }

  function eventPos(e) {
    var r = $('pvView').getBoundingClientRect();
    return {
      x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)),
      y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)),
      cx: e.clientX, cy: e.clientY, rect: r
    };
  }

  function placeCursor(pos) {
    var c = $('pvCursor'), d = local.brush / 1000 * Math.max(pos.rect.width, pos.rect.height) * 2;
    var stage = $('pvStage').getBoundingClientRect();
    c.style.width = c.style.height = d + 'px';
    c.style.left = (pos.cx - stage.left) + 'px';
    c.style.top = (pos.cy - stage.top) + 'px';
    c.firstElementChild.style.width = c.firstElementChild.style.height = Math.max(2, local.hardness) + '%';
    c.hidden = false;
  }

  function bind() {
    var v = $('pvView');
    v.addEventListener('pointerdown', function (e) {
      if (!pv || !state) return;
      e.preventDefault();
      var pos = eventPos(e);
      if (local.tool === 'none') return;
      if (local.tool === 'pick') {
        var mode = e.altKey ? 'sub' : (e.shiftKey || e.ctrlKey || e.metaKey) ? 'add' : local.pickMode;
        send({ t: 'pick', x: pos.x, y: pos.y, mode: mode });
        return;
      }
      v.setPointerCapture(e.pointerId);
      drawing = { mode: local.tool, r: local.brush / 1000, h: local.hardness / 100, pts: [{ x: pos.x, y: pos.y }] };
      draw();
    });
    v.addEventListener('pointermove', function (e) {
      if (!pv) return;
      var pos = eventPos(e);
      if (local.tool === 'erase' || local.tool === 'restore') placeCursor(pos);
      if (drawing) { drawing.pts.push({ x: pos.x, y: pos.y }); draw(); }
    });
    var end = function () {
      if (!drawing) return;
      var st = drawing;
      // vis stroeket til panelet svarer med ny tilstand
      state = Object.assign({}, state, { strokes: state.strokes.concat([st]) });
      drawing = null;
      send({ t: 'stroke', stroke: st });
      draw();
    };
    v.addEventListener('pointerup', end);
    v.addEventListener('pointercancel', end);
    v.addEventListener('pointerleave', function () { $('pvCursor').hidden = true; });

    document.querySelectorAll('[data-tool]').forEach(function (b) {
      b.addEventListener('click', function () { local.tool = local.tool === b.dataset.tool ? 'none' : b.dataset.tool; syncTools(); send({ t: 'tool', tool: local.tool }); });
    });
    document.querySelectorAll('[data-pickmode]').forEach(function (b) {
      b.addEventListener('click', function () { local.pickMode = b.dataset.pickmode; syncTools(); send({ t: 'pickmode', v: local.pickMode }); });
    });
    $('pvSize').addEventListener('input', function (e) { local.brush = +e.target.value; $('pvSizeOut').textContent = e.target.value; });
    $('pvSize').addEventListener('change', function (e) { send({ t: 'brush', v: +e.target.value }); });
    $('pvHard').addEventListener('input', function (e) { local.hardness = +e.target.value; $('pvHardOut').textContent = e.target.value; });
    $('pvHard').addEventListener('change', function (e) { send({ t: 'hardness', v: +e.target.value }); });

    var cmp = $('pvCompare');
    cmp.addEventListener('pointerdown', function (e) { e.preventDefault(); compare = true; draw(); });
    var off = function () { if (compare) { compare = false; draw(); } };
    cmp.addEventListener('pointerup', off);
    cmp.addEventListener('pointerleave', off);
    UI.bgPicker($('pvBg'), $('pvStage'));
    $('pvUndo').addEventListener('click', function () { send({ t: 'undo' }); });
    $('pvRedo').addEventListener('click', function () { send({ t: 'redo' }); });
    $('pvClose').addEventListener('click', function () { send({ t: 'close' }); });
    document.addEventListener('keydown', function (e) {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      var k = (e.key || '').toLowerCase();
      if (k === 'z' && !e.shiftKey) { e.preventDefault(); send({ t: 'undo' }); }
      else if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); send({ t: 'redo' }); }
    });
  }

  function start() {
    UI.initIcons();
    UI.initTooltips();
    bind();
    syncTools();
    Office.onReady(function () {
      Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, function (arg) { receive(arg.message); }, function () {
        send({ t: 'ready' });
      });
    });
  }

  start();
})();
