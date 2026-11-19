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

    S.src = { full: full, proxy: proxy, name: opt.name || 'image', mime: blob.type || 'image/png', meta: opt.meta || null, capped: Math.max(sw, sh) > FULL_MAX,
      // Originalen slik den var da den ble hentet, for "Reset"
      origBlob: blob, origLossless: !c.l && !c.t && !c.r && !c.b && Math.max(sw, sh) <= FULL_MAX,
      origMeta: opt.meta ? Object.assign({}, opt.meta) : null, slideChanged: false,
      origCrop: c, sessionId: newSessionId() };
    S.strokes = [];
    S.key.colors = []; S.key.seeds = []; S.key.enabled = false;
    S.adjust = Object.assign({}, Ops.DEFAULT_ADJUST);
    S.trim = false;
    S.out = { w: full.width, h: full.height, lock: true };
    // Beskjaering fra PowerPoint (i pikselkoordinater) og hva som sist ble lagt paa lysbildet
    S.crop = { x: 0, y: 0, w: full.width, h: full.height };
    S.slideSync = S.src.meta ? {
      rect: { x: 0, y: 0, w: full.width, h: full.height },
      geom: { left: S.src.meta.left, top: S.src.meta.top, width: S.src.meta.width, height: S.src.meta.height }
    } : null;
    cache = {};
    live_.applied = editKey(); live_.lastKey = live_.applied;
    resetHistory();
    // Versjon 0 er bildet slik det var da det ble hentet
    S.src.versionCounter = 0; S.src.currentVersion = 0;
    S.src.versions = { 0: versionRecord(snapshotKey()) };
    setLiveState('');

    var view = $('view');
    view.width = proxy.width; view.height = proxy.height;
    $('empty').hidden = true;
    ['stage', 'toolbar', 'infoRow', 'panels', 'actions'].forEach(function (id) { $(id).hidden = false; });
    applyPreviewPref();
    syncUI();
    render();
    var msg = 'Image loaded (' + W + ' × ' + H + ' px).';
    if (S.src.capped) msg += ' Processing at ' + full.width + ' × ' + full.height + ' px.';
    status(msg, 'ok');
  }

  async function getFromPowerPoint() {
    status('');
    loading = true;
    await busy(true, 'Getting the image from PowerPoint ...');
    try {
      var pic = await Bridge.getSelectedImage();
      var tag = parseTag(pic.meta && pic.meta.sessionId);
      var rec = tag ? await Store.get(tag.sid) : null;
      if (rec && rec.versions) await restoreSession(rec, pic, tag.v);
      else await loadBlob(new Blob([pic.bytes], { type: pic.mime }), pic);
    } catch (e) {
      status(e.message || String(e), 'err');
    } finally {
      loading = false;
      busy(false);
    }
  }

  /* ---------------- Huske redigeringer (oekter) ---------------- */

  // Originalen, innstillingene og angrehistorikken lagres lokalt (IndexedDB),
  // knyttet til et usynlig merke paa bildet i presentasjonen. Da kan panelet
  // fortsette der det slapp selv om det har vaert lukket eller lastet paa nytt.
  var Store = (function () {
    var mem = {};
    var dbp = null;
    function db() {
      if (dbp) return dbp;
      dbp = new Promise(function (resolve) {
        try {
          var req = indexedDB.open('pptimagepoint', 1);
          req.onupgradeneeded = function () { req.result.createObjectStore('sessions', { keyPath: 'id' }); };
          req.onsuccess = function () { resolve(req.result); };
          req.onerror = function () { resolve(null); };
        } catch (e) { resolve(null); }
      });
      return dbp;
    }
    function tx(mode, fn) {
      return db().then(function (d) {
        if (!d) return null;
        return new Promise(function (resolve) {
          try {
            var t = d.transaction('sessions', mode);
            var r = fn(t.objectStore('sessions'));
            t.oncomplete = function () { resolve(r && r.result); };
            t.onerror = t.onabort = function () { resolve(null); };
          } catch (e) { resolve(null); }
        });
      });
    }
    return {
      get: async function (id) {
        if (mem[id]) return mem[id];
        return (await tx('readonly', function (st) { return st.get(id); })) || null;
      },
      put: async function (rec) {
        mem[rec.id] = rec;
        await tx('readwrite', function (st) { return st.put(rec); });
        // behold bare de 30 nyeste
        var all = await tx('readonly', function (st) { return st.getAll(); });
        if (all && all.length > 30) {
          all.sort(function (a, b) { return a.savedAt - b.savedAt; });
          var old = all.slice(0, all.length - 30);
          await tx('readwrite', function (st) { old.forEach(function (o) { st.delete(o.id); }); });
        }
      }
    };
  })();

  function newSessionId() {
    return 'pip-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }

  function saveSession(stateStr) {
    var src = S.src, m = src.meta;
    var rec = {
      id: src.sessionId, savedAt: Date.now(), name: src.name, mime: src.mime,
      origBlob: src.origBlob, origCrop: src.origCrop, origMeta: src.origMeta, origLossless: src.origLossless,
      versions: src.versions, versionCounter: src.versionCounter, currentVersion: src.currentVersion,
      hist: { past: hist.past.slice(-100), future: hist.future.slice() }
    };
    Store.put(rec).catch(function () { /* ikke kritisk */ });
  }

  async function restoreSession(rec, pic, version) {
    var v = rec.versions && rec.versions[version] ? version : rec.currentVersion;
    var vr = rec.versions[v];
    await loadBlob(rec.origBlob, { crop: rec.origCrop, name: rec.name, meta: Object.assign({}, pic.meta) });
    S.src.sessionId = rec.id;
    S.src.origMeta = rec.origMeta;
    S.src.origLossless = rec.origLossless;
    S.src.slideChanged = v !== 0;
    S.src.versions = rec.versions;
    S.src.versionCounter = rec.versionCounter;
    hist.past = (rec.hist && rec.hist.past) || [];
    hist.future = (rec.hist && rec.hist.future) || [];
    useVersion(v, vr, pic.meta.shapeId);
    hist.committed = hist.lastKey = vr.state;
    updateHistoryButtons();
    status('Your earlier edits are restored. Keep adjusting, undo, or press Reset.', 'ok');
  }

  /** Det som skal til for aa fortsette fra en bestemt versjon paa lysbildet. */
  function versionRecord(stateStr) {
    var m = S.src.meta;
    return {
      state: stateStr, crop: S.crop,
      syncRect: S.slideSync ? S.slideSync.rect : null, syncGeom: S.slideSync ? S.slideSync.geom : null,
      frame: m ? { left: m.left, top: m.top, width: m.width, height: m.height } : null
    };
  }

  function trimVersions() {
    var keys = Object.keys(S.src.versions).map(Number).filter(function (k) { return k > 0; }).sort(function (a, b) { return a - b; });
    while (keys.length > 60) delete S.src.versions[keys.shift()];
  }

  /** Setter panelet til innstillingene for en versjon som ligger paa lysbildet (uten aa endre lysbildet). */
  function useVersion(v, vr, shapeId) {
    applySnapshot(vr.state);
    S.crop = vr.crop || S.crop;
    S.slideSync = vr.syncRect && vr.syncGeom ? { rect: vr.syncRect, geom: vr.syncGeom } : null;
    if (S.src.meta && vr.frame) Object.assign(S.src.meta, vr.frame);
    if (S.src.meta && shapeId) S.src.meta.shapeId = shapeId;
    S.src.currentVersion = v;
    live_.applied = editKey(); live_.lastKey = live_.applied;
    render();
  }

  function parseTag(tag) {
    if (!tag) return null;
    var i = tag.lastIndexOf('|');
    return i < 0 ? { sid: tag, v: null } : { sid: tag.slice(0, i), v: parseInt(tag.slice(i + 1), 10) };
  }

  /* ---------------- Automatisk henting ved markering ---------------- */

  var loading = false;        // henter et bilde akkurat naa
  var suppressUntil = 0;      // ignorer markeringer vi selv har laget (ved innsetting)
  var selTimer = null;

  async function autoLoad() {
    if (loading || live_.running || Date.now() < suppressUntil) { retryLater(); return; }
    var info;
    try { info = await Bridge.getSelectedImageId(); } catch (e) { return; }
    if (loading || live_.running || Date.now() < suppressUntil) { retryLater(); return; }
    if (!info) return;
    var tag = parseTag(info.sessionId);
    if (S.src && S.src.meta && tag && tag.sid === S.src.sessionId) {
      // Samme bilde. Har PowerPoint angret eller gjort om, ligger en annen versjon paa lysbildet.
      var vr = S.src.versions[tag.v];
      if (tag.v !== S.src.currentVersion && vr) {
        useVersion(tag.v, vr, info.id);
        status('Matched the picture on the slide (undo or redo in PowerPoint).', 'ok');
      } else {
        S.src.meta.shapeId = info.id;
      }
      return;
    }
    if (S.src && S.src.meta && !tag && S.src.origMeta && info.id === S.src.origMeta.shapeId && S.src.currentVersion !== 0) {
      // PowerPoint har angret helt tilbake til originalbildet
      useVersion(0, S.src.versions[0], info.id);
      S.src.slideChanged = false;
      status('Matched the original picture on the slide (undo in PowerPoint).', 'ok');
      return;
    }
    if (S.src && S.src.meta && S.src.meta.shapeId === info.id) return;
    await getFromPowerPoint();
  }

  function retryLater() {
    clearTimeout(selTimer);
    selTimer = setTimeout(autoLoad, 400);
  }

  function onSelectionChanged() {
    clearTimeout(selTimer);
    selTimer = setTimeout(autoLoad, 250);
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
    drawCropOverlay(ctx, P);
    updateInfo();
  }

  /** Demper omraadet som er beskaaret bort i PowerPoint. */
  function drawCropOverlay(ctx, P) {
    var F = S.src.full, c = S.crop;
    if (!c || (!c.x && !c.y && c.w === F.width && c.h === F.height)) return;
    var k = P.width / F.width;
    var x0 = c.x * k, y0 = c.y * k, x1 = (c.x + c.w) * k, y1 = (c.y + c.h) * k;
    ctx.save();
    ctx.fillStyle = 'rgba(0,0,0,0.55)';
    ctx.fillRect(0, 0, P.width, y0);
    ctx.fillRect(0, y1, P.width, P.height - y1);
    ctx.fillRect(0, y0, x0, y1 - y0);
    ctx.fillRect(x1, y0, P.width - x1, y1 - y0);
    ctx.strokeStyle = 'rgba(255,255,255,0.9)';
    ctx.setLineDash([4, 3]);
    ctx.strokeRect(x0 + 0.5, y0 + 0.5, x1 - x0 - 1, y1 - y0 - 1);
    ctx.restore();
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
    var c = S.crop || { x: 0, y: 0, w: F.width, h: F.height };
    if (c.x || c.y || c.w !== F.width || c.h !== F.height) out = Ops.crop(out, c);
    var rect = { x: c.x, y: c.y, w: c.w, h: c.h };
    var tr = trimRectFor(out);
    if (tr) { out = Ops.crop(out, tr); rect = { x: c.x + tr.x, y: c.y + tr.y, w: tr.w, h: tr.h }; }
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

  /**
   * Leser av om bildet er flyttet, skalert eller beskaaret i PowerPoint siden
   * sist, og tar det med videre. Returnerer false hvis bildet er slettet.
   */
  async function syncFromSlide() {
    var m = S.src.meta, sync = S.slideSync;
    if (!m || !sync) return true;
    var g;
    try { g = await Bridge.getShapeGeometry(m); } catch (e) { return true; }
    if (g === undefined) return true;
    if (g === null) return false;
    var same = ['left', 'top', 'width', 'height'].every(function (k) { return Math.abs(g[k] - sync.geom[k]) < 0.05; });
    if (same) return true;
    var c = { l: 0, t: 0, r: 0, b: 0 };
    try { c = (await Bridge.getShapeCrop(m)) || c; } catch (e) { /* uten beskjaering: bare flyttet/skalert */ }
    var R = sync.rect;
    var cx = R.x + c.l * R.w, cy = R.y + c.t * R.h;
    var cw = R.w * (1 - c.l - c.r), ch = R.h * (1 - c.t - c.b);
    if (cw < 1 || ch < 1) return true;
    var F = S.src.full;
    var sx = g.width / cw, sy = g.height / ch;
    // Rammen for hele kildebildet slik det ville ligget paa lysbildet
    m.left = g.left - cx * sx; m.top = g.top - cy * sy;
    m.width = F.width * sx; m.height = F.height * sy;
    var x0 = Math.max(0, Math.round(cx)), y0 = Math.max(0, Math.round(cy));
    var x1 = Math.min(F.width, Math.round(cx + cw)), y1 = Math.min(F.height, Math.round(cy + ch));
    S.crop = { x: x0, y: y0, w: Math.max(1, x1 - x0), h: Math.max(1, y1 - y0) };
    S.slideSync = { rect: { x: cx, y: cy, w: cw, h: ch }, geom: g };
    render();
    return true;
  }

  /**
   * Behandler bildet i full opploesning og legger det paa lysbildet.
   * live = true: stille oppdatering i bakgrunnen (ingen sperre av panelet).
   */
  async function applyToSlide(live, opts) {
    opts = opts || {};
    if (!S.src) return;
    if (!live) status('');
    var key = editKey();
    var snap = snapshotKey();
    suppressUntil = Infinity;
    if (!live) await busy(true, 'Processing at full resolution ...');
    else setLiveState('Updating slide ...');
    try {
      if (Bridge.available && S.src.meta && !opts.original) {
        var present = await syncFromSlide();
        if (!present) {
          S.src.meta = null; S.slideSync = null;
          if (live) {
            live_.on = false; $('live').checked = false; syncLiveUI(); setLiveState('');
            status('The picture was removed from the slide. Live update is turned off.', 'err');
            return;
          }
        }
      }
      var res, enc;
      if (opts.original && S.src.origLossless) {
        // Legg tilbake den opprinnelige filen uendret (ingen ny komprimering)
        res = { image: S.src.full, rect: { x: 0, y: 0, w: S.src.full.width, h: S.src.full.height } };
        enc = { blob: S.src.origBlob, ext: (S.src.mime.split('/')[1] || 'png').replace('jpeg', 'jpg') };
      } else {
        res = renderFull();
        enc = await encode(res.image);
      }
      if (!Bridge.available) { download(enc); return; }
      if (!live) await busy(true, 'Inserting into the slide ...');
      var b64 = await blobToBase64(enc.blob);
      var replace = (live || $('replace').checked) && S.src.meta;
      var version = ++S.src.versionCounter;
      var r = await Bridge.insertImage(b64, placementFor(res.rect), replace ? S.src.meta : null, S.src.sessionId + '|' + version);
      S.src.slideChanged = !opts.original;
      S.slideSync = r && r.meta ? {
        rect: res.rect,
        geom: { left: r.meta.left, top: r.meta.top, width: r.meta.width, height: r.meta.height }
      } : null;
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
      live_.applied = key;
      if (r && r.meta && S.src.meta) {
        // Husk hvilke innstillinger denne versjonen paa lysbildet har (for Ctrl+Z i PowerPoint)
        S.src.currentVersion = version;
        S.src.versions[version] = versionRecord(snap);
        trimVersions();
        saveSession();
      }
      if (live) {
        setLiveState('Slide updated');
        if (replace && r && !r.deleted) status('The previous version could not be removed automatically. Delete it manually.', 'err');
      } else {
        var msg = 'Image inserted (' + res.image.width + ' × ' + res.image.height + ' px).';
        if (replace && r && !r.deleted) msg += ' The original could not be removed automatically. Delete it manually if needed.';
        status(msg, 'ok');
      }
    } catch (e) {
      if (live) { live_.on = false; $('live').checked = false; syncLiveUI(); setLiveState(''); }
      status(e.message || String(e), 'err');
    } finally {
      // Det nye bildet er markert og har samme id som vi lagret, saa det lastes ikke paa nytt.
      // Fikk vi ikke vite id-en, venter vi litt foer markeringen tolkes som et nytt valg.
      suppressUntil = S.src && S.src.meta ? 0 : Date.now() + 1500;
      if (!live) busy(false);
    }
  }

  function doInsert() { return applyToSlide(false); }

  /** Setter alt tilbake til slik bildet var da det ble hentet, ogsaa paa lysbildet. */
  async function resetToOriginal() {
    if (!S.src || live_.running) return;
    var F = S.src.full;
    S.adjust = Object.assign({}, Ops.DEFAULT_ADJUST);
    S.key.colors = []; S.key.seeds = []; S.key.enabled = false;
    S.strokes = []; S.trim = false;
    S.out = { w: F.width, h: F.height, lock: true };
    S.crop = { x: 0, y: 0, w: F.width, h: F.height };
    cache = {};
    syncUI(); render();
    if (Bridge.available && S.src.meta && S.src.origMeta && S.src.slideChanged) {
      var om = S.src.origMeta;
      S.src.meta = Object.assign({}, S.src.meta, { left: om.left, top: om.top, width: om.width, height: om.height });
      live_.running = true;
      try { await applyToSlide(false, { original: true }); } finally { live_.running = false; }
      live_.applied = editKey(); live_.lastKey = live_.applied;
      status('The picture is back to the original.', 'ok');
    } else {
      status('All changes are reset.', 'ok');
    }
  }

  /* ---------------- Angre / gjoer om ---------------- */

  // Historikken lagrer hele redigeringstilstanden som JSON. En endring
  // lagres som eget steg naar den har staatt i ro litt (slik at et helt
  // drag med en glidebryter blir ett steg).
  var hist = { past: [], future: [], committed: null, lastKey: null, changedAt: 0 };

  function snapshotKey() {
    return JSON.stringify({ adjust: S.adjust, key: S.key, strokes: S.strokes, trim: S.trim, out: S.out, format: S.format });
  }

  function resetHistory() {
    hist.past = []; hist.future = [];
    hist.committed = hist.lastKey = snapshotKey();
    updateHistoryButtons();
  }

  function commitHistory(k) {
    hist.past.push(hist.committed);
    if (hist.past.length > 100) hist.past.shift();
    hist.committed = k; hist.lastKey = k;
    hist.future = [];
    updateHistoryButtons();
  }

  function commitNow() {
    if (!S.src || live_.drawing) return;
    var k = snapshotKey();
    if (k !== hist.committed) commitHistory(k);
  }

  function historyTick() {
    if (!S.src) return;
    var k = snapshotKey();
    if (k !== hist.lastKey) { hist.lastKey = k; hist.changedAt = Date.now(); return; }
    if (k === hist.committed || live_.drawing) return;
    if (Date.now() - hist.changedAt < 400) return;
    commitHistory(k);
  }

  function applySnapshot(str) {
    var o = JSON.parse(str);
    S.adjust = o.adjust; S.key = o.key; S.strokes = o.strokes;
    S.trim = o.trim; S.out = o.out; S.format = o.format;
    hist.lastKey = str;
    syncUI(); render();
    updateHistoryButtons();
  }

  function undo() {
    if (!S.src || live_.drawing) return;
    var k = snapshotKey();
    if (k !== hist.committed) commitHistory(k);
    if (!hist.past.length) return;
    hist.future.push(hist.committed);
    hist.committed = hist.past.pop();
    applySnapshot(hist.committed);
  }

  function redo() {
    if (!S.src || live_.drawing || !hist.future.length) return;
    if (snapshotKey() !== hist.committed) return; // ny endring siden angring: ingenting aa gjoere om
    hist.past.push(hist.committed);
    hist.committed = hist.future.pop();
    applySnapshot(hist.committed);
  }

  function updateHistoryButtons() {
    var dirty = S.src && snapshotKey() !== hist.committed;
    $('btnUndo').disabled = !S.src || (!hist.past.length && !dirty);
    $('btnRedo').disabled = !S.src || !hist.future.length;
    $('btnResetAll').disabled = !S.src;
  }

  /* ---------------- Skjul forhaandsvisning ---------------- */

  var previewHidden = false;
  try { previewHidden = localStorage.getItem('pip.hidePreview') === '1'; } catch (e) { /* ikke kritisk */ }

  function applyPreviewPref() {
    var hasImg = !!S.src;
    $('stage').hidden = !hasImg || previewHidden;
    document.querySelector('.seg').hidden = previewHidden;
    $('btnCompare').hidden = previewHidden;
    $('bgSelect').hidden = previewHidden;
    $('previewHint').hidden = !hasImg || !previewHidden;
    setIcon($('btnTogglePreview'), previewHidden ? 'eye' : 'eyeOff');
    $('btnTogglePreview').dataset.tip = previewHidden ? 'Show preview' : 'Hide preview';
    $('btnTogglePreview').setAttribute('aria-label', $('btnTogglePreview').dataset.tip);
  }

  function bindPreviewToggle() {
    $('btnTogglePreview').addEventListener('click', function () {
      previewHidden = !previewHidden;
      try { localStorage.setItem('pip.hidePreview', previewHidden ? '1' : '0'); } catch (e) { /* ikke kritisk */ }
      if (previewHidden) setTool('pick');
      applyPreviewPref();
      updateCursorSize();
    });
  }

  function bindHistory() {
    $('btnUndo').addEventListener('click', undo);
    $('btnRedo').addEventListener('click', redo);
    document.addEventListener('keydown', function (e) {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      var t = e.target;
      // la tallfelt beholde sin egen angring av tekst
      if (t && t.tagName === 'INPUT' && t.type === 'number') return;
      var key = (e.key || '').toLowerCase();
      if (key === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
      else if (key === 'y' || (key === 'z' && e.shiftKey)) { e.preventDefault(); redo(); }
    });
    setInterval(function () { historyTick(); updateHistoryButtons(); }, 200);
    // Klikk, valg og slipp av glidebryter blir egne steg med en gang (ikke avhengig av tid)
    var commitSoon = function () { setTimeout(commitNow, 0); };
    document.addEventListener('change', commitSoon, true);
    document.addEventListener('click', commitSoon, true);
    $('view').addEventListener('pointerup', commitSoon);
  }

  /* ---------------- Live-oppdatering av lysbildet ---------------- */

  var live_ = { on: false, applied: null, lastKey: null, changedAt: 0, running: false, drawing: false };

  function editKey() {
    if (!S.src) return '';
    return JSON.stringify([S.adjust, keyActive() ? keyParamsKey() : 0, S.key.defringe,
      S.strokes.map(function (st) { return st.pts.length; }), S.trim, S.out.w, S.out.h, S.format]);
  }

  function setLiveState(t) { $('liveState').textContent = t || ''; }

  function syncLiveUI() {
    var on = live_.on && Bridge.available;
    $('replace').parentElement.hidden = on || !Bridge.available;
    $('btnInsert').textContent = on ? 'Apply now' : 'Insert into slide';
  }

  async function liveTick() {
    if (!live_.on || !S.src || !S.src.meta || live_.running || loading) return;
    var key = editKey();
    if (key !== live_.lastKey) { live_.lastKey = key; live_.changedAt = Date.now(); return; }
    if (key === live_.applied || live_.drawing || !$('busy').hidden) return;
    if (Date.now() - live_.changedAt < 600) return;
    live_.running = true;
    try { await applyToSlide(true); } finally { live_.running = false; }
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
    var txt = 'Original ' + F.width + ' × ' + F.height + ' px.';
    var c = S.crop || { w: F.width, h: F.height };
    var w = Math.round(S.out.w * c.w / F.width), h = Math.round(S.out.h * c.h / F.height);
    if (c.w !== F.width || c.h !== F.height) txt += ' Cropped in PowerPoint.';
    if (S.trim && cache.preview) {
      var r = Ops.opaqueBounds(cache.preview, 8);
      if (r) {
        var P = S.src.proxy;
        w = Math.round(r.w / P.width * S.out.w); h = Math.round(r.h / P.height * S.out.h);
      }
    }
    txt += ' Output approx. ' + w + ' × ' + h + ' px.';
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
      live_.drawing = true;
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
    var end = function () { drawing = null; live_.drawing = false; };
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

  function bindControls() {
    $('btnGet').addEventListener('click', getFromPowerPoint);
    document.querySelector('label[for="fileInput"]').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('fileInput').click(); }
    });
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
    $('btnClearStrokes').addEventListener('click', function () { S.strokes = []; render(); });

    $('btnResetAdjust').addEventListener('click', function () {
      S.adjust = Object.assign({}, Ops.DEFAULT_ADJUST); syncUI(); render();
    });

    $('btnInsert').addEventListener('click', doInsert);
    $('btnDownload').addEventListener('click', doDownload);
    $('btnResetAll').addEventListener('click', resetToOriginal);
    window.addEventListener('resize', updateCursorSize);
  }

  /* ---------------- Ikoner og verktoeytips ---------------- */

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
    download: '<path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/>'
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
    var tip = $('tip'), current = null, timer = null;
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

  async function start() {
    initIcons();
    initTooltips();
    buildSliders();
    bindControls();
    bindStage();
    bindHistory();
    bindPreviewToggle();
    var inPPT = await Bridge.init();
    if (!inPPT) {
      $('btnGet').hidden = true;
      $('btnInsert').textContent = 'Download image';
      $('btnDownload').hidden = true;
      $('replace').parentElement.hidden = true;
      status('Running outside PowerPoint. You can try the tools and download the result.');
      return;
    }
    $('emptyTitle').textContent = 'Select a picture on the slide.';
    $('emptyText').textContent = 'It opens here automatically. You can also drop an image here or paste one with Ctrl+V.';
    Bridge.onSelectionChanged(onSelectionChanged);
    $('liveWrap').hidden = false;
    live_.on = $('live').checked;
    $('live').addEventListener('change', function (e) { live_.on = e.target.checked; syncLiveUI(); if (!live_.on) setLiveState(''); });
    syncLiveUI();
    setInterval(liveTick, 250);
    // Naar du kommer tilbake til panelet (f.eks. etter Ctrl+Z i lysbildet), sjekk hva som ligger der
    window.addEventListener('focus', onSelectionChanged);
    document.documentElement.addEventListener('mouseenter', onSelectionChanged);
    autoLoad();
  }

  window.__pip = { hist: function () { return hist; }, undo: undo, redo: redo, S: S, renderFull: renderFull, loadBlob: loadBlob, renderNow: renderNow, cache: function () { return cache; }, pickAt: pickAt, syncUI: syncUI };
  start();
})();
