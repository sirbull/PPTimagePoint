/* PPTimagePoint: brukergrensesnitt og arbeidsflyt. */
(function () {
  'use strict';

  var Ops = window.ImageOps;
  var PROXY_MAX = 900;     // lengste side i forhaandsvisningen
  var FULL_MAX = 6000;     // lengste side vi behandler i full opploesning

  var $ = function (id) { return document.getElementById(id); };

  var S = {
    src: null, // { full, proxy, name, mime, meta }
    adjust: Object.assign({}, Ops.DEFAULT_ADJUST),
    key: { enabled: false, colors: [], seeds: [], tolerance: 18, softness: 10, mode: 'connected', fromEdges: true, choke: 1, feather: 1, defringe: true },
    strokes: [],
    trim: false,
    brush: 30,
    tool: 'pick',
    out: { w: 0, h: 0, lock: true },
    format: 'auto',
    compare: false
  };

  var cache = {};
  var sliders = {};

  /* ---------------- Hjelpere ---------------- */

  function status(msg, kind) {
    var el = $('status');
    el.textContent = msg || '';
    el.className = 'status' + (kind ? ' ' + kind : '');
    if (kind === 'ok') setTimeout(function () { if (el.textContent === msg) status(''); }, 5000);
  }

  function busy(on, text) {
    $('busy').hidden = !on;
    if (text) $('busyText').textContent = text;
    return new Promise(function (r) { setTimeout(r, on ? 40 : 0); });
  }

  function scaleOf(img) { return Math.max(img.width, img.height) / 1000; }

  function imageDataFromSource(source, sx, sy, sw, sh, maxSide) {
    var f = Math.min(1, maxSide / Math.max(sw, sh));
    var w = Math.max(1, Math.round(sw * f)), h = Math.max(1, Math.round(sh * f));
    var c = document.createElement('canvas');
    c.width = w; c.height = h;
    var ctx = c.getContext('2d', { willReadFrequently: true });
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(source, sx, sy, sw, sh, 0, 0, w, h);
    return ctx.getImageData(0, 0, w, h);
  }

  function decodeBlob(blob) {
    if (window.createImageBitmap && blob.type !== 'image/svg+xml') {
      return createImageBitmap(blob).catch(function () { return decodeViaImg(blob); });
    }
    return decodeViaImg(blob);
  }

  function decodeViaImg(blob) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(blob);
      var img = new Image();
      img.onload = function () { resolve(img); setTimeout(function () { URL.revokeObjectURL(url); }, 1000); };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error('Could not read the image file.')); };
      img.src = url;
    });
  }

  function toCanvas(img) {
    var c = document.createElement('canvas');
    c.width = img.width; c.height = img.height;
    c.getContext('2d').putImageData(img, 0, 0);
    return c;
  }

  function canvasToBlob(canvas, type, quality) {
    return new Promise(function (resolve, reject) {
      canvas.toBlob(function (b) { b ? resolve(b) : reject(new Error('Could not create the image file.')); }, type, quality);
    });
  }

  function blobToBase64(blob) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(String(fr.result).split(',')[1]); };
      fr.onerror = function () { reject(fr.error); };
      fr.readAsDataURL(blob);
    });
  }

  /* ---------------- Innlasting ---------------- */

  async function loadBlob(blob, opt) {
    opt = opt || {};
    var bmp = await decodeBlob(blob);
    var W = bmp.width || bmp.naturalWidth, H = bmp.height || bmp.naturalHeight;
    if (!W || !H) throw new Error('The image appears to be empty.');
    var c = opt.crop || { l: 0, t: 0, r: 0, b: 0 };
    var sx = Math.round(Math.max(0, c.l) * W), sy = Math.round(Math.max(0, c.t) * H);
    var sw = Math.max(1, W - sx - Math.round(Math.max(0, c.r) * W));
    var sh = Math.max(1, H - sy - Math.round(Math.max(0, c.b) * H));
    var full = imageDataFromSource(bmp, sx, sy, sw, sh, FULL_MAX);
    var proxy = imageDataFromSource(bmp, sx, sy, sw, sh, PROXY_MAX);
    if (bmp.close) bmp.close();

    S.src = { full: full, proxy: proxy, name: opt.name || 'image', mime: blob.type || 'image/png', meta: opt.meta || null, capped: Math.max(sw, sh) > FULL_MAX };
    S.strokes = [];
    S.key.colors = []; S.key.seeds = []; S.key.enabled = false;
    S.out.w = full.width; S.out.h = full.height;
    cache = {};

    var view = $('view');
    view.width = proxy.width; view.height = proxy.height;
    $('empty').hidden = true;
    ['stage', 'toolbar', 'info', 'panels', 'actions'].forEach(function (id) { $(id).hidden = false; });
    syncUI();
    render();
    var msg = 'Image loaded (' + W + ' × ' + H + ' px).';
    if (S.src.capped) msg += ' Processing at ' + full.width + ' × ' + full.height + ' px.';
    status(msg, 'ok');
  }

  async function getFromPowerPoint() {
    status('');
    await busy(true, 'Getting the image from PowerPoint ...');
    try {
      var pic = await Bridge.getSelectedImage();
      await loadBlob(new Blob([pic.bytes], { type: pic.mime }), pic);
    } catch (e) {
      status(e.message || String(e), 'err');
    } finally {
      busy(false);
    }
  }

  /* ---------------- Behandling ---------------- */

  function keyParamsKey() {
    var k = S.key;
    return JSON.stringify([k.colors.map(function (c) { return [c.r, c.g, c.b]; }), k.seeds, k.tolerance, k.softness, k.mode, k.fromEdges, k.choke, k.feather]);
  }

  function keyActive() { return S.key.enabled && S.key.colors.length > 0; }

  /** Lager alfakart for et bilde. region kan gis inn (fra forhaandsvisningen). */
  function buildAlpha(img, isProxy) {
    var w = img.width, h = img.height, s = scaleOf(img);
    var alpha = null;
    if (keyActive()) {
      var dist = Ops.distanceMap(img, S.key.colors);
      var region;
      if (isProxy || S.key.mode === 'global') {
        region = Ops.computeRegion(dist, w, h, S.key);
        if (isProxy) cache.region = region;
      } else {
        var P = S.src.proxy;
        var pr = cache.region || Ops.computeRegion(Ops.distanceMap(P, S.key.colors), P.width, P.height, S.key);
        region = Ops.upscaleRegion(pr, P.width, P.height, w, h);
      }
      alpha = Ops.alphaFromDist(dist, region, S.key.tolerance, S.key.softness);
      alpha = Ops.refineAlpha(alpha, w, h, S.key.choke * s, S.key.feather * s);
    }
    return alpha;
  }

  function finalize(colorImg, baseAlpha) {
    var w = colorImg.width, h = colorImg.height;
    var alpha = baseAlpha;
    if (S.strokes.length) {
      if (!alpha) { alpha = new Uint8Array(w * h); alpha.fill(255); }
      alpha = Ops.applyStrokes(alpha, w, h, S.strokes);
    }
    var keyRgb = keyActive() ? S.key.colors.map(function (c) { return Ops.adjustColor(c, S.adjust); }) : null;
    return Ops.compose(colorImg, alpha, keyRgb, S.key.defringe);
  }

  var rafPending = false;
  function render() {
    if (rafPending) return;
    rafPending = true;
    requestAnimationFrame(function () { rafPending = false; renderNow(); });
  }

  function renderNow() {
    if (!S.src) return;
    var P = S.src.proxy;
    var ctx = $('view').getContext('2d');
    if (S.compare) { ctx.clearRect(0, 0, P.width, P.height); ctx.putImageData(P, 0, 0); return; }

    var ak = JSON.stringify(S.adjust);
    if (cache.adjKey !== ak) { cache.adjusted = Ops.adjust(P, S.adjust, scaleOf(P)); cache.adjKey = ak; }
    var kk = keyActive() ? keyParamsKey() : 'off';
    if (cache.maskKey !== kk) { cache.region = null; cache.alpha = buildAlpha(P, true); cache.maskKey = kk; }
    var out = finalize(cache.adjusted, cache.alpha);
    cache.preview = out;
    ctx.clearRect(0, 0, P.width, P.height);
    ctx.putImageData(out, 0, 0);
    updateInfo();
  }

  function trimRectFor(img) {
    if (!S.trim) return null;
    var r = Ops.opaqueBounds(img, 8);
    if (!r || (r.w === img.width && r.h === img.height)) return null;
    return r;
  }

  /** Full opploesning. Returnerer { image, rect } der rect er i full-koordinater foer skalering. */
  function renderFull() {
    var F = S.src.full;
    var adjusted = Ops.adjust(F, S.adjust, scaleOf(F));
    var alpha = buildAlpha(F, false);
    var out = finalize(adjusted, alpha);
    var rect = { x: 0, y: 0, w: F.width, h: F.height };
    var tr = trimRectFor(out);
    if (tr) { out = Ops.crop(out, tr); rect = tr; }
    var factor = S.out.w / F.width;
    var nw = Math.max(1, Math.round(rect.w * factor));
    var nh = Math.max(1, Math.round(rect.h * (S.out.h / F.height)));
    if (nw !== out.width || nh !== out.height) out = Ops.resize(out, nw, nh);
    return { image: out, rect: rect };
  }

  async function encode(img) {
    var fmt = S.format;
    var transparent = Ops.hasTransparency(img);
    if (fmt === 'auto') fmt = transparent ? 'png' : (S.src.mime === 'image/jpeg' ? 'jpeg' : 'png');
    var canvas = toCanvas(img);
    if (fmt === 'jpeg') {
      if (transparent) {
        var c2 = document.createElement('canvas');
        c2.width = canvas.width; c2.height = canvas.height;
        var x = c2.getContext('2d');
        x.fillStyle = '#ffffff'; x.fillRect(0, 0, c2.width, c2.height); x.drawImage(canvas, 0, 0);
        canvas = c2;
      }
      return { blob: await canvasToBlob(canvas, 'image/jpeg', 0.92), ext: 'jpg' };
    }
    return { blob: await canvasToBlob(canvas, 'image/png'), ext: 'png' };
  }

  function placementFor(rect) {
    var m = S.src.meta;
    if (!m || !m.width) return null;
    var F = S.src.full;
    var px = m.width / F.width, py = m.height / F.height;
    return { left: m.left + rect.x * px, top: m.top + rect.y * py, width: rect.w * px, height: rect.h * py };
  }

  async function doInsert() {
    if (!S.src) return;
    status('');
    await busy(true, 'Processing at full resolution ...');
    try {
      var res = renderFull();
      var enc = await encode(res.image);
      if (!Bridge.available) { download(enc); return; }
      await busy(true, 'Inserting into the slide ...');
      var b64 = await blobToBase64(enc.blob);
      var replace = $('replace').checked && S.src.meta;
      var r = await Bridge.insertImage(b64, placementFor(res.rect), replace ? S.src.meta : null);
      if (r && r.meta) {
        // Videre arbeid gjelder det nye bildet, men vi beholder originalpikslene i panelet
        S.src.meta = Object.assign({}, r.meta, {
          left: r.meta.left - res.rect.x * (r.meta.width / res.rect.w),
          top: r.meta.top - res.rect.y * (r.meta.height / res.rect.h),
          width: r.meta.width * S.src.full.width / res.rect.w,
          height: r.meta.height * S.src.full.height / res.rect.h
        });
      } else if (replace) {
        S.src.meta = null;
      }
      var msg = 'Image inserted (' + res.image.width + ' × ' + res.image.height + ' px).';
      if (replace && r && !r.deleted) msg += ' The original could not be removed automatically. Delete it manually if needed.';
      status(msg, 'ok');
    } catch (e) {
      status(e.message || String(e), 'err');
    } finally {
      busy(false);
    }
  }

  async function doDownload() {
    if (!S.src) return;
    await busy(true, 'Processing at full resolution ...');
    try { download(await encode(renderFull().image)); }
    catch (e) { status(e.message || String(e), 'err'); }
    finally { busy(false); }
  }

  function download(enc) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(enc.blob);
    a.download = (S.src.name || 'image').replace(/[^\w\-æøåÆØÅ ]+/g, '').trim().replace(/\s+/g, '-') + '-edited.' + enc.ext;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
    status('Image downloaded.', 'ok');
  }

  /* ---------------- Grensesnitt ---------------- */

  function fmtVal(v, step) {
    var dec = String(step).indexOf('.') >= 0 ? String(step).split('.')[1].length : 0;
    var s = Number(v).toFixed(dec);
    return Number(v) > 0 && !/^-/.test(s) ? '+' + s : s;
  }

  function buildSliders() {
    document.querySelectorAll('.slider').forEach(function (el) {
      var key = el.dataset.key;
      var min = +el.dataset.min, max = +el.dataset.max, step = el.dataset.step, def = +el.dataset.def;
      var id = 'sl-' + key;
      el.innerHTML = '<div class="lbl"><label for="' + id + '">' + el.dataset.label + '</label><output></output></div>' +
        '<input type="range" id="' + id + '" min="' + min + '" max="' + max + '" step="' + step + '" value="' + def + '">';
      if (el.dataset.track) el.classList.add('track-' + el.dataset.track);
      var input = el.querySelector('input'), out = el.querySelector('output');
      var isAdj = el.classList.contains('adj');
      var signed = min < 0;
      var show = function () { out.textContent = signed ? fmtVal(input.value, step) : input.value; };
      var apply = function () {
        var v = +input.value;
        if (isAdj) S.adjust[key] = v;
        else if (key === 'brush') S.brush = v;
        else S.key[key] = v;
        show();
        if (key === 'brush') updateCursorSize(); else render();
      };
      input.addEventListener('input', apply);
      input.addEventListener('dblclick', function () { input.value = def; apply(); });
      sliders[key] = { input: input, show: show, def: def };
      show();
    });
  }

  function setSlider(key, v) {
    var s = sliders[key];
    if (!s) return;
    s.input.value = v; s.show();
  }

  function syncUI() {
    $('keyEnabled').checked = S.key.enabled;
    $('keyMode').value = S.key.mode;
    $('fromEdges').checked = S.key.fromEdges;
    $('edgesWrap').hidden = S.key.mode !== 'connected';
    $('defringe').checked = S.key.defringe;
    $('trim').checked = S.trim;
    ['tolerance', 'softness', 'choke', 'feather'].forEach(function (k) { setSlider(k, S.key[k]); });
    Object.keys(Ops.DEFAULT_ADJUST).forEach(function (k) { setSlider(k, S.adjust[k]); });
    setSlider('brush', S.brush);
    $('outW').value = S.out.w; $('outH').value = S.out.h;
    $('lockBtn').classList.toggle('on', S.out.lock);
    $('lockBtn').textContent = S.out.lock ? 'Locked' : 'Free';
    $('lockBtn').setAttribute('aria-pressed', String(S.out.lock));
    $('format').value = S.format;
    $('btnOptimal').disabled = !(S.src && S.src.meta && S.src.meta.width);
    renderChips();
  }

  function renderChips() {
    var box = $('chips');
    box.innerHTML = '';
    S.key.colors.forEach(function (c, i) {
      var chip = document.createElement('span');
      chip.className = 'chip';
      var hex = '#' + [c.r, c.g, c.b].map(function (v) { return ('0' + v.toString(16)).slice(-2); }).join('');
      chip.innerHTML = '<i style="background:' + hex + '"></i>' + hex.toUpperCase() + '<button type="button" title="Remove color" aria-label="Remove color">×</button>';
      chip.querySelector('button').addEventListener('click', function () {
        S.key.colors.splice(i, 1);
        S.key.seeds.splice(i, 1);
        if (!S.key.colors.length) S.key.enabled = false;
        syncUI(); render();
      });
      box.appendChild(chip);
    });
  }

  function updateInfo() {
    if (!S.src) return;
    var F = S.src.full;
    var txt = 'Original: ' + F.width + ' × ' + F.height + ' px';
    var w = S.out.w, h = S.out.h;
    if (S.trim && cache.preview) {
      var r = Ops.opaqueBounds(cache.preview, 8);
      if (r) {
        var P = S.src.proxy;
        w = Math.round(r.w / P.width * S.out.w); h = Math.round(r.h / P.height * S.out.h);
      }
    }
    txt += '   Output: approx. ' + w + ' × ' + h + ' px';
    $('info').textContent = txt;
  }

  function setTool(t) {
    S.tool = t;
    document.querySelectorAll('[data-tool]').forEach(function (b) { b.classList.toggle('on', b.dataset.tool === t); });
    var st = $('stage');
    st.classList.toggle('tool-erase', t === 'erase');
    st.classList.toggle('tool-restore', t === 'restore');
    $('brushCursor').hidden = true;
  }

  function updateCursorSize() {
    var c = $('brushCursor'), v = $('view');
    if (!S.src) return;
    var rect = v.getBoundingClientRect();
    var d = S.brush / 1000 * Math.max(rect.width, rect.height) * 2;
    c.style.width = d + 'px'; c.style.height = d + 'px';
  }

  function eventPos(e) {
    var rect = $('view').getBoundingClientRect();
    return {
      x: Math.min(1, Math.max(0, (e.clientX - rect.left) / rect.width)),
      y: Math.min(1, Math.max(0, (e.clientY - rect.top) / rect.height)),
      cx: e.clientX - rect.left, cy: e.clientY - rect.top, rect: rect
    };
  }

  function pickAt(pos, add) {
    var c = Ops.sampleColor(S.src.proxy, pos.x, pos.y, 2);
    if (add) { S.key.colors.push(c); S.key.seeds.push({ x: pos.x, y: pos.y }); }
    else { S.key.colors = [c]; S.key.seeds = [{ x: pos.x, y: pos.y }]; }
    S.key.enabled = true;
    syncUI(); render();
  }

  function bindStage() {
    var view = $('view'), stage = $('stage'), cursor = $('brushCursor');
    var drawing = null;
    view.addEventListener('pointerdown', function (e) {
      if (!S.src) return;
      e.preventDefault();
      var pos = eventPos(e);
      if (S.tool === 'pick') { pickAt(pos, e.shiftKey || e.ctrlKey || e.metaKey); return; }
      view.setPointerCapture(e.pointerId);
      drawing = { mode: S.tool, r: S.brush / 1000, pts: [{ x: pos.x, y: pos.y }] };
      S.strokes.push(drawing);
      render();
    });
    view.addEventListener('pointermove', function (e) {
      if (!S.src) return;
      var pos = eventPos(e);
      if (S.tool !== 'pick') {
        cursor.hidden = false;
        var off = view.offsetLeft, top = view.offsetTop;
        cursor.style.left = (off + pos.cx) + 'px'; cursor.style.top = (top + pos.cy) + 'px';
        updateCursorSize();
      }
      if (drawing) { drawing.pts.push({ x: pos.x, y: pos.y }); render(); }
    });
    var end = function () { drawing = null; };
    view.addEventListener('pointerup', end);
    view.addEventListener('pointercancel', end);
    view.addEventListener('pointerleave', function () { cursor.hidden = true; });

    // Dra og slipp
    [stage, $('empty'), document.body].forEach(function (el) {
      el.addEventListener('dragover', function (e) { e.preventDefault(); $('empty').classList.add('drag'); });
      el.addEventListener('dragleave', function () { $('empty').classList.remove('drag'); });
      el.addEventListener('drop', function (e) {
        e.preventDefault(); e.stopPropagation();
        $('empty').classList.remove('drag');
        var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
        if (f) openFile(f);
      });
    });
    document.addEventListener('paste', function (e) {
      var items = (e.clipboardData && e.clipboardData.items) || [];
      for (var i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image/') === 0) { openFile(items[i].getAsFile()); e.preventDefault(); return; }
      }
    });
  }

  function openFile(file) {
    if (!file) return;
    busy(true, 'Reading the image ...').then(function () {
      return loadBlob(file, { name: (file.name || 'image').replace(/\.[^.]+$/, '') });
    }).catch(function (e) { status(e.message || String(e), 'err'); })
      .then(function () { busy(false); });
  }

  function setOutW(w) {
    var F = S.src.full;
    S.out.w = Math.max(1, Math.round(w));
    if (S.out.lock) S.out.h = Math.max(1, Math.round(S.out.w * F.height / F.width));
    syncUI(); updateInfo();
  }

  function bindControls() {
    $('btnGet').addEventListener('click', getFromPowerPoint);
    $('fileInput').addEventListener('change', function (e) { openFile(e.target.files[0]); e.target.value = ''; });

    document.querySelectorAll('[data-tool]').forEach(function (b) {
      b.addEventListener('click', function () { setTool(b.dataset.tool); });
    });
    var cmp = $('btnCompare');
    var on = function (e) { e.preventDefault(); S.compare = true; renderNow(); };
    var off = function () { if (S.compare) { S.compare = false; renderNow(); } };
    cmp.addEventListener('pointerdown', on);
    cmp.addEventListener('pointerup', off);
    cmp.addEventListener('pointerleave', off);
    $('bgSelect').addEventListener('change', function (e) {
      var st = $('stage');
      st.classList.remove('checker', 'bg-white', 'bg-black', 'bg-slide');
      st.classList.add(e.target.value === 'checker' ? 'checker' : 'bg-' + e.target.value);
    });

    $('btnAuto').addEventListener('click', function () {
      if (!S.src) return;
      var c = Ops.detectBorderColor(S.src.proxy);
      if (!c) { status('No clear background color found along the edges.', 'err'); return; }
      S.key.colors = [c]; S.key.seeds = [{ x: 0, y: 0 }];
      S.key.mode = 'connected'; S.key.fromEdges = true; S.key.enabled = true;
      syncUI(); render();
    });
    $('keyEnabled').addEventListener('change', function (e) {
      if (e.target.checked && !S.key.colors.length) {
        e.target.checked = false;
        status('Click the background in the image, or use "Auto detect background".');
        return;
      }
      S.key.enabled = e.target.checked; render();
    });
    $('keyMode').addEventListener('change', function (e) { S.key.mode = e.target.value; syncUI(); render(); });
    $('fromEdges').addEventListener('change', function (e) { S.key.fromEdges = e.target.checked; render(); });
    $('defringe').addEventListener('change', function (e) { S.key.defringe = e.target.checked; render(); });
    $('trim').addEventListener('change', function (e) { S.trim = e.target.checked; updateInfo(); });
    $('btnUndoStroke').addEventListener('click', function () { S.strokes.pop(); render(); });
    $('btnClearStrokes').addEventListener('click', function () { S.strokes = []; render(); });

    $('btnResetAdjust').addEventListener('click', function () {
      S.adjust = Object.assign({}, Ops.DEFAULT_ADJUST); syncUI(); render();
    });

    $('outW').addEventListener('change', function (e) { if (S.src) setOutW(+e.target.value || 1); });
    $('outH').addEventListener('change', function (e) {
      if (!S.src) return;
      var F = S.src.full;
      S.out.h = Math.max(1, Math.round(+e.target.value || 1));
      if (S.out.lock) S.out.w = Math.max(1, Math.round(S.out.h * F.width / F.height));
      syncUI(); updateInfo();
    });
    $('lockBtn').addEventListener('click', function () {
      S.out.lock = !S.out.lock;
      if (S.out.lock && S.src) setOutW(S.out.w); else syncUI();
    });
    document.querySelectorAll('[data-pct]').forEach(function (b) {
      b.addEventListener('click', function () { if (S.src) { S.out.lock = true; setOutW(S.src.full.width * (+b.dataset.pct) / 100); } });
    });
    $('btnOptimal').addEventListener('click', function () {
      var m = S.src && S.src.meta;
      if (!m) return;
      var target = Math.round(m.width / 72 * 220);
      var F = S.src.full;
      S.out.lock = true;
      if (target >= F.width) {
        setOutW(F.width);
        status('The image already has the right or a lower resolution for its size on the slide.', 'ok');
      } else {
        setOutW(target);
        status('Set to ' + S.out.w + ' × ' + S.out.h + ' px (220 ppi on the slide).', 'ok');
      }
    });
    $('format').addEventListener('change', function (e) { S.format = e.target.value; });

    $('btnInsert').addEventListener('click', doInsert);
    $('btnDownload').addEventListener('click', doDownload);
    $('btnResetAll').addEventListener('click', function () {
      if (!S.src) return;
      S.adjust = Object.assign({}, Ops.DEFAULT_ADJUST);
      S.key.colors = []; S.key.seeds = []; S.key.enabled = false;
      S.strokes = []; S.trim = false;
      S.out = { w: S.src.full.width, h: S.src.full.height, lock: true };
      cache = {};
      syncUI(); render();
    });
    window.addEventListener('resize', updateCursorSize);
  }

  async function start() {
    buildSliders();
    bindControls();
    bindStage();
    var inPPT = await Bridge.init();
    if (!inPPT) {
      $('btnGet').hidden = true;
      $('btnInsert').hidden = true;
      $('replace').parentElement.hidden = true;
      $('btnDownload').classList.add('primary', 'grow');
      $('btnDownload').textContent = 'Download image';
      status('Running outside PowerPoint. You can try the tools and download the result.');
    }
  }

  window.__pip = { S: S, renderFull: renderFull, loadBlob: loadBlob, renderNow: renderNow, cache: function () { return cache; }, pickAt: pickAt, syncUI: syncUI };
  start();
})();
