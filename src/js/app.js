/* PPTimagePoint: brukergrensesnitt og arbeidsflyt. */
(function () {
  'use strict';

  var Ops = window.ImageOps;
  var PROXY_MAX = 900;     // lengste side i forhaandsvisningen
  var FULL_MAX = 6000;     // lengste side vi behandler i full opploesning

  var $ = function (id) { return document.getElementById(id); };

  var S = {
    src: null, // { full, proxy, name, mime, meta }
    adjust: Ops.defaultAdjust(),
    key: { enabled: false, colors: [], seeds: [], tolerance: 18, softness: 10, mode: 'connected', fromEdges: true, choke: 1, feather: 1, defringe: true },
    strokes: [],
    trim: false,
    brush: 30,
    hardness: 65,
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
    S.adjust = Ops.defaultAdjust();
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
    var msg = 'Image loaded (' + sw + ' × ' + sh + ' px).';
    if (S.src.capped) msg += ' Processing at ' + full.width + ' × ' + full.height + ' px.';
    status(msg, 'ok');
  }

  async function getFromPowerPoint() {
    status('');
    loading = true;
    await busy(true, 'Getting the image from PowerPoint ...');
    try {
      var pic = await Bridge.getSelectedImage();
      var hash = await hashBytes(pic.bytes);
      // Kjenner vi igjen bildet? Foerst paa innholdet, saa paa det usynlige merket.
      var known = await Store.get('h:' + hash);
      var tag = known ? { sid: known.sid, v: known.v } : parseTag(pic.meta && pic.meta.sessionId);
      if (tag && S.src && tag.sid === S.src.sessionId && S.src.versions[tag.v] && pic.meta) {
        // Et av panelets egne bilder (f.eks. etter angring i PowerPoint): bare bytt versjon
        if (!S.src.meta) S.src.meta = Object.assign({}, pic.meta);
        S.src.meta.slideId = pic.meta.slideId;
        useVersion(tag.v, S.src.versions[tag.v], pic.meta.shapeId);
        S.src.slideChanged = tag.v !== 0 || S.src.slideChanged;
        return;
      }
      var rec = tag ? await Store.get(tag.sid) : null;
      if (rec && rec.versions) await restoreSession(rec, pic, tag.v);
      else {
        await loadBlob(new Blob([pic.bytes], { type: pic.mime }), pic);
        S.src.origHash = hash;
      }
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
        // behold de 30 nyeste oektene og de 800 nyeste fingeravtrykkene
        var keys = await tx('readonly', function (st) { return st.getAllKeys(); });
        if (!keys) return;
        var sessions = keys.filter(function (k) { return String(k).indexOf('pip-') === 0; });
        var hashes = keys.filter(function (k) { return String(k).indexOf('h:') === 0; });
        var drop = [];
        // id-ene for oekter starter med tidsstempel, saa sortering gir eldste foerst
        if (sessions.length > 30) drop = drop.concat(sessions.sort().slice(0, sessions.length - 30));
        if (hashes.length > 800) {
          var all = await tx('readonly', function (st) { return st.getAll(IDBKeyRange.bound('h:', 'h:\uffff')); });
          (all || []).sort(function (a, b) { return a.savedAt - b.savedAt; })
            .slice(0, hashes.length - 800).forEach(function (o) { drop.push(o.id); });
        }
        if (drop.length) await tx('readwrite', function (st) { drop.forEach(function (id) { st.delete(id); }); });
      }
    };
  })();

  /** Fingeravtrykk av en bildefil, slik at panelet kjenner igjen sine egne bilder. */
  async function hashBytes(data) {
    var u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
    try {
      if (window.crypto && crypto.subtle) {
        var d = new Uint8Array(await crypto.subtle.digest('SHA-256', u8));
        return Array.prototype.map.call(d.subarray(0, 16), function (b) { return ('0' + b.toString(16)).slice(-2); }).join('');
      }
    } catch (e) { /* faller tilbake under */ }
    var h1 = 0x811c9dc5, h2 = 0x01000193;
    for (var i = 0; i < u8.length; i++) { h1 = Math.imul(h1 ^ u8[i], 16777619) >>> 0; h2 = Math.imul(h2 + u8[i], 2246822519) >>> 0; }
    return 'f' + h1.toString(16) + h2.toString(16) + u8.length.toString(16);
  }

  function rememberHash(hash, sid, v) {
    if (!hash) return;
    Store.put({ id: 'h:' + hash, sid: sid, v: v, savedAt: Date.now() }).catch(function () { /* ikke kritisk */ });
  }

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
    if (src.origHash && !src.origHashSaved) { src.origHashSaved = true; rememberHash(src.origHash, src.sessionId, 0); }
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

  function keyActive() { return Pipeline.keyActive(S); }

  function pipeState() {
    return { adjust: S.adjust, key: S.key, strokes: S.strokes, crop: S.crop, fullW: S.src.full.width, fullH: S.src.full.height };
  }

  /** Alfakart i full opploesning, med omraadet fra forhaandsvisningen. */
  function buildAlpha(img) {
    return Pipeline.alphaFor(img, S, { proxy: S.src.proxy, proxyRegion: cache.pv && cache.pv.region }).alpha;
  }

  function finalize(colorImg, baseAlpha) { return Pipeline.finalize(colorImg, baseAlpha, S); }

  var rafPending = false;
  function render() {
    if (rafPending) return;
    rafPending = true;
    requestAnimationFrame(function () { rafPending = false; renderNow(); });
  }

  function renderNow() {
    if (!S.src) return;
    if (!cache.pv) cache.pv = new Pipeline.Previewer(S.src.proxy);
    cache.pv.draw($('view').getContext('2d'), pipeState(), S.compare);
    if (!S.compare) cache.preview = cache.pv.last;
    updateInfo();
    Popout.push();
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
    var alpha = buildAlpha(F);
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
        rememberHash(await hashBytes(await enc.blob.arrayBuffer()), S.src.sessionId, version);
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
    S.adjust = Ops.defaultAdjust();
    S.key.colors = []; S.key.seeds = []; S.key.enabled = false;
    S.strokes = []; S.trim = false;
    S.out = { w: F.width, h: F.height, lock: true };
    S.crop = { x: 0, y: 0, w: F.width, h: F.height };
    cache = {};
    syncUI(); render();
    if (Bridge.available && S.src.origMeta && (S.src.slideChanged || S.src.currentVersion !== 0)) {
      var om = S.src.origMeta;
      if (!S.src.meta) {
        // Mistet sporet av bildet paa lysbildet: bruk det som er markert, ellers sett inn ved siden av
        var sel = null;
        try { sel = await Bridge.getSelectedImageId(); } catch (e) { sel = null; }
        S.src.meta = Object.assign({}, om, { shapeId: sel ? sel.id : null });
      }
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
    var flags = $('btnUndo').disabled + ',' + $('btnRedo').disabled;
    if (flags !== lastHistFlags) { lastHistFlags = flags; Popout.push(); }
  }
  var lastHistFlags = '';

  /* ---------------- Skjul forhaandsvisning ---------------- */

  var previewHidden = false;
  try { previewHidden = localStorage.getItem('pip.hidePreview') === '1'; } catch (e) { /* ikke kritisk */ }

  function applyPreviewPref() {
    var hasImg = !!S.src;
    $('stage').hidden = !hasImg || previewHidden;
    document.querySelector('.seg').hidden = previewHidden;
    $('btnCompare').hidden = previewHidden;
    $('brushBar').hidden = !hasImg || previewHidden || (S.tool !== 'erase' && S.tool !== 'restore');
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
        else if (el.classList.contains('mix')) setMixerValue(key, v);
        else if (el.classList.contains('lvl')) setLevelsValue(key, v);
        else if (key === 'brush') S.brush = v;
        else if (key === 'hardness') S.hardness = v;
        else S.key[key] = v;
        show();
        if (key === 'brush' || key === 'hardness') { updateCursorSize(); Popout.push(); } else render();
      };
      input.addEventListener('input', apply);
      input.addEventListener('dblclick', function () { input.value = def; apply(); });
      sliders[key] = { input: input, show: show, def: def };
      show();
    });
  }

  /* ---------------- Color mixer, Levels og Curves ---------------- */

  var mixSel = 0;           // valgt fargeomraade i Color mixer
  var curveChan = 'rgb';    // valgt kanal i Curves
  var MIX_PART = { mixH: 0, mixS: 1, mixL: 2 };
  var LV_FIELD = { lvInBlack: 'inBlack', lvGamma: 'gamma', lvInWhite: 'inWhite', lvOutBlack: 'outBlack', lvOutWhite: 'outWhite' };

  function adjustExtras() {
    var d = Ops.defaultAdjust();
    if (!S.adjust.mixer) S.adjust.mixer = d.mixer;
    if (!S.adjust.levels) S.adjust.levels = d.levels;
    if (!S.adjust.curves) S.adjust.curves = d.curves;
  }

  function setMixerValue(key, v) {
    adjustExtras();
    var m = S.adjust.mixer.slice();
    m[mixSel * 3 + MIX_PART[key]] = v;
    S.adjust.mixer = m;
    renderSwatches();
  }

  function setLevelsValue(key, v) {
    adjustExtras();
    var l = Object.assign({}, S.adjust.levels);
    l[LV_FIELD[key]] = v;
    // hold svart under hvitt
    if (l.inWhite <= l.inBlack) { if (key === 'lvInBlack') l.inWhite = Math.min(255, l.inBlack + 1); else l.inBlack = Math.max(0, l.inWhite - 1); }
    S.adjust.levels = l;
    setSlider('lvInBlack', l.inBlack); setSlider('lvInWhite', l.inWhite);
    drawHistogram();
  }

  function hsl(h, s, l) { return 'hsl(' + h + ',' + s + '%,' + l + '%)'; }

  function renderSwatches() {
    var box = $('mixSwatches');
    if (!box.childElementCount) {
      Ops.MIX_COLORS.forEach(function (c, i) {
        var b = document.createElement('button');
        b.type = 'button'; b.className = 'swatch'; b.setAttribute('role', 'radio');
        b.style.setProperty('--sw', hsl(c.h, 85, 52));
        b.dataset.tip = c.name; b.setAttribute('aria-label', c.name);
        b.addEventListener('click', function () { mixSel = i; syncAdvancedUI(); });
        box.appendChild(b);
      });
    }
    var m = S.adjust.mixer || [];
    Array.prototype.forEach.call(box.children, function (b, i) {
      b.classList.toggle('on', i === mixSel);
      b.setAttribute('aria-checked', String(i === mixSel));
      b.classList.toggle('changed', !!(m[i * 3] || m[i * 3 + 1] || m[i * 3 + 2]));
    });
  }

  function syncAdvancedUI() {
    adjustExtras();
    var m = S.adjust.mixer, c = Ops.MIX_COLORS[mixSel].h;
    setSlider('mixH', m[mixSel * 3]); setSlider('mixS', m[mixSel * 3 + 1]); setSlider('mixL', m[mixSel * 3 + 2]);
    // fargede spor for valgt farge
    var hEl = $('sl-mixH').parentElement, sEl = $('sl-mixS').parentElement, lEl = $('sl-mixL').parentElement;
    var wheel = [];
    for (var a = -180; a <= 180; a += 30) wheel.push(hsl(c + a, 85, 52));
    hEl.style.setProperty('--track-bg', 'linear-gradient(90deg,' + wheel.join(',') + ')');
    sEl.style.setProperty('--track-bg', 'linear-gradient(90deg,' + hsl(c, 0, 55) + ',' + hsl(c, 95, 50) + ')');
    lEl.style.setProperty('--track-bg', 'linear-gradient(90deg,' + hsl(c, 80, 15) + ',' + hsl(c, 85, 50) + ',' + hsl(c, 90, 88) + ')');
    $('sl-mixH').previousElementSibling.firstElementChild.textContent = Ops.MIX_COLORS[mixSel].name + ' hue';
    renderSwatches();
    var l = S.adjust.levels;
    Object.keys(LV_FIELD).forEach(function (k) { setSlider(k, l[LV_FIELD[k]]); });
    drawHistogram();
    drawCurve();
  }

  /** Histogram for originalbildet (lysstyrke og hver kanal), regnet ut en gang per bilde. */
  function histogram() {
    if (!S.src) return null;
    if (S.src.hist) return S.src.hist;
    var d = S.src.proxy.data, H = { rgb: new Uint32Array(256), r: new Uint32Array(256), g: new Uint32Array(256), b: new Uint32Array(256) };
    for (var q = 0; q < d.length; q += 4) {
      if (d[q + 3] < 16) continue;
      H.r[d[q]]++; H.g[d[q + 1]]++; H.b[d[q + 2]]++;
      H.rgb[(d[q] * 54 + d[q + 1] * 183 + d[q + 2] * 19) >> 8]++;
    }
    S.src.hist = H;
    return H;
  }

  var CHAN_COLOR = { rgb: '#8a8886', r: '#e0453a', g: '#3fae49', b: '#3a7bd5' };

  function drawHist(ctx, W, H, hist, color) {
    var max = 1;
    for (var i = 2; i < 254; i++) if (hist[i] > max) max = hist[i];
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(0, H);
    for (i = 0; i < 256; i++) ctx.lineTo(i / 255 * W, H - Math.min(1, hist[i] / max) * H * 0.95);
    ctx.lineTo(W, H);
    ctx.closePath();
    ctx.fill();
  }

  function drawHistogram() {
    var cv = $('histLevels'), ctx = cv.getContext('2d'), W = cv.width, H = cv.height;
    ctx.clearRect(0, 0, W, H);
    var h = histogram();
    if (!h) return;
    ctx.globalAlpha = 0.55;
    drawHist(ctx, W, H, h.rgb, getComputedStyle(document.body).color);
    ctx.globalAlpha = 1;
    var l = S.adjust.levels || Ops.defaultAdjust().levels;
    var mark = function (v, col) {
      var x = v / 255 * W;
      ctx.fillStyle = col; ctx.strokeStyle = '#000';
      ctx.beginPath(); ctx.moveTo(x, H - 16); ctx.lineTo(x - 9, H); ctx.lineTo(x + 9, H); ctx.closePath(); ctx.fill(); ctx.stroke();
    };
    var mid = l.inBlack + (l.inWhite - l.inBlack) * Math.pow(0.5, l.gamma);
    mark(l.inBlack, '#000'); mark(mid, '#888'); mark(l.inWhite, '#fff');
  }

  function curvePoints() { adjustExtras(); return S.adjust.curves[curveChan]; }

  function drawCurve() {
    var cv = $('curveCanvas'), ctx = cv.getContext('2d'), W = cv.width, H = cv.height;
    var dark = matchMedia('(prefers-color-scheme: dark)').matches;
    ctx.clearRect(0, 0, W, H);
    ctx.fillStyle = dark ? '#1f1f1f' : '#fbfbfb';
    ctx.fillRect(0, 0, W, H);
    var h = histogram();
    if (h) { ctx.globalAlpha = 0.28; drawHist(ctx, W, H, h[curveChan], CHAN_COLOR[curveChan]); ctx.globalAlpha = 1; }
    ctx.strokeStyle = dark ? 'rgba(255,255,255,.12)' : 'rgba(0,0,0,.1)';
    ctx.lineWidth = 2;
    for (var i = 1; i < 4; i++) {
      ctx.beginPath(); ctx.moveTo(i * W / 4, 0); ctx.lineTo(i * W / 4, H); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(0, i * H / 4); ctx.lineTo(W, i * H / 4); ctx.stroke();
    }
    ctx.setLineDash([8, 8]);
    ctx.beginPath(); ctx.moveTo(0, H); ctx.lineTo(W, 0); ctx.stroke();
    ctx.setLineDash([]);
    // de andre kanalene svakt i bakgrunnen
    adjustExtras();
    ['r', 'g', 'b', 'rgb'].forEach(function (ch) {
      if (ch === curveChan) return;
      var pts = S.adjust.curves[ch];
      if (pts.length === 2 && pts[0][1] === 0 && pts[1][1] === 255 && pts[0][0] === 0 && pts[1][0] === 255) return;
      plot(ctx, Ops.curveLUT(pts), CHAN_COLOR[ch], 2, 0.45, W, H);
    });
    plot(ctx, Ops.curveLUT(curvePoints()), curveChan === 'rgb' ? (dark ? '#f3f2f1' : '#201f1e') : CHAN_COLOR[curveChan], 4, 1, W, H);
    curvePoints().forEach(function (p) {
      ctx.beginPath();
      ctx.arc(p[0] / 255 * W, H - p[1] / 255 * H, 10, 0, Math.PI * 2);
      ctx.fillStyle = '#fff'; ctx.fill();
      ctx.lineWidth = 3; ctx.strokeStyle = CHAN_COLOR[curveChan] === '#8a8886' ? '#201f1e' : CHAN_COLOR[curveChan]; ctx.stroke();
    });
  }

  function plot(ctx, lut, color, width, alpha, W, H) {
    ctx.globalAlpha = alpha; ctx.strokeStyle = color; ctx.lineWidth = width;
    ctx.beginPath();
    for (var x = 0; x < 256; x++) {
      var px = x / 255 * W, py = H - lut[x] / 255 * H;
      if (x) ctx.lineTo(px, py); else ctx.moveTo(px, py);
    }
    ctx.stroke(); ctx.globalAlpha = 1;
  }

  function setCurve(pts) {
    adjustExtras();
    var c = Object.assign({}, S.adjust.curves);
    c[curveChan] = pts;
    S.adjust.curves = c;
    drawCurve(); render();
  }

  function bindCurves() {
    var cv = $('curveCanvas'), drag = -1;
    var toVal = function (e) {
      var r = cv.getBoundingClientRect();
      return { x: (e.clientX - r.left) / r.width * 255, y: (1 - (e.clientY - r.top) / r.height) * 255, out: e.clientY < r.top - 30 || e.clientY > r.bottom + 30 || e.clientX < r.left - 30 || e.clientX > r.right + 30, scale: 255 / r.width };
    };
    var nearest = function (v) {
      var pts = curvePoints(), best = -1, bd = 1e9;
      pts.forEach(function (p, i) { var d = Math.hypot(p[0] - v.x, p[1] - v.y); if (d < bd) { bd = d; best = i; } });
      return bd <= 14 * v.scale ? best : -1;
    };
    cv.addEventListener('pointerdown', function (e) {
      if (!S.src) return;
      e.preventDefault();
      cv.setPointerCapture(e.pointerId);
      var v = toVal(e);
      drag = nearest(v);
      if (drag < 0) {
        var pts = curvePoints().map(function (p) { return p.slice(); });
        var x = Math.round(Math.min(254, Math.max(1, v.x)));
        if (pts.some(function (p) { return Math.abs(p[0] - x) < 3; })) return;
        pts.push([x, Math.round(Math.min(255, Math.max(0, v.y)))]);
        pts.sort(function (a, b) { return a[0] - b[0]; });
        drag = pts.findIndex(function (p) { return p[0] === x; });
        setCurve(pts);
      }
    });
    cv.addEventListener('pointermove', function (e) {
      if (drag < 0) return;
      var v = toVal(e), pts = curvePoints().map(function (p) { return p.slice(); });
      var last = pts.length - 1;
      if (v.out && drag > 0 && drag < last) { pts.splice(drag, 1); drag = -1; setCurve(pts); return; }
      var lo = drag === 0 ? 0 : pts[drag - 1][0] + 1, hi = drag === last ? 255 : pts[drag + 1][0] - 1;
      if (drag === 0) hi = Math.min(hi, 254);
      if (drag === last) lo = Math.max(lo, 1);
      pts[drag] = [Math.round(Math.min(hi, Math.max(lo, v.x))), Math.round(Math.min(255, Math.max(0, v.y)))];
      setCurve(pts);
    });
    var end = function () { drag = -1; };
    cv.addEventListener('pointerup', end);
    cv.addEventListener('pointercancel', end);
    cv.addEventListener('dblclick', function (e) {
      var i = nearest(toVal(e)), pts = curvePoints();
      if (i > 0 && i < pts.length - 1) { pts = pts.slice(); pts.splice(i, 1); setCurve(pts); }
    });
    document.querySelectorAll('[data-chan]').forEach(function (b) {
      b.addEventListener('click', function () {
        curveChan = b.dataset.chan;
        document.querySelectorAll('[data-chan]').forEach(function (x) { x.classList.toggle('on', x === b); });
        drawCurve();
      });
    });
    $('btnResetCurves').addEventListener('click', function () { S.adjust.curves = Ops.defaultAdjust().curves; drawCurve(); render(); });
    $('btnResetLevels').addEventListener('click', function () { S.adjust.levels = Ops.defaultAdjust().levels; syncAdvancedUI(); render(); });
    $('btnResetMixer').addEventListener('click', function () { S.adjust.mixer = Ops.defaultAdjust().mixer; syncAdvancedUI(); render(); });
    // tegn kurven paa nytt naar seksjonen aapnes (stoerrelsen er da kjent)
    $('secAdvanced').addEventListener('toggle', function () { drawHistogram(); drawCurve(); });
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
    Ops.NUM_KEYS.forEach(function (k) { setSlider(k, S.adjust[k]); });
    syncAdvancedUI();
    setSlider('brush', S.brush);
    setSlider('hardness', S.hardness);
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
    $('brushBar').hidden = !S.src || previewHidden || (t !== 'erase' && t !== 'restore');
    Popout.push();
  }

  function updateCursorSize() {
    var c = $('brushCursor'), v = $('view');
    if (!S.src) return;
    var rect = v.getBoundingClientRect();
    var d = S.brush / 1000 * Math.max(rect.width, rect.height) * 2;
    c.style.width = d + 'px'; c.style.height = d + 'px';
    var inner = c.firstElementChild || c.appendChild(document.createElement('i'));
    inner.style.width = inner.style.height = Math.max(2, S.hardness) + '%';
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
      drawing = { mode: S.tool, r: S.brush / 1000, h: S.hardness / 100, pts: [{ x: pos.x, y: pos.y }] };
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
    UI.bgPicker($('bgSelect'), $('stage'));

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
      S.adjust = Ops.defaultAdjust(); syncUI(); render();
    });

    $('btnInsert').addEventListener('click', doInsert);
    $('btnDownload').addEventListener('click', doDownload);
    $('btnResetAll').addEventListener('click', resetToOriginal);
    window.addEventListener('resize', updateCursorSize);
  }

  /* ---------------- Stort forhaandsvisningsvindu ---------------- */

  // Aapnes som et Office-dialogvindu (preview.html). Panelet eier tilstanden:
  // vinduet faar bildet og innstillingene, og sender klikk, penselstroek og
  // angring tilbake hit. Lange meldinger deles opp i biter.
  var Popout = (function () {
    var dlg = null, ready = false, sentFor = null, timer = null, inbox = {};
    var CHUNK = 30000;

    function supported() {
      try { return !!(Office.context.ui && Office.context.ui.displayDialogAsync) && Office.context.requirements.isSetSupported('DialogApi', '1.2'); }
      catch (e) { return false; }
    }

    function send(obj) {
      if (!dlg) return;
      var str = JSON.stringify(obj);
      try {
        if (str.length <= CHUNK) { dlg.messageChild(str); return; }
        var id = Math.random().toString(36).slice(2), n = Math.ceil(str.length / CHUNK);
        for (var i = 0; i < n; i++) dlg.messageChild(JSON.stringify({ t: 'chunk', id: id, i: i, n: n, d: str.slice(i * CHUNK, (i + 1) * CHUNK) }));
      } catch (e) { /* vinduet er lukket */ }
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
      handle(m);
    }

    function imageMessage() {
      var F = S.src.full, f = Math.min(1, 1600 / Math.max(F.width, F.height));
      var c = document.createElement('canvas');
      c.width = Math.max(1, Math.round(F.width * f)); c.height = Math.max(1, Math.round(F.height * f));
      var x = c.getContext('2d');
      x.imageSmoothingQuality = 'high';
      x.drawImage(toCanvas(F), 0, 0, c.width, c.height);
      var url = c.toDataURL('image/webp', 0.95);
      if (url.indexOf('data:image/webp') !== 0) url = c.toDataURL('image/png');
      return { t: 'image', url: url };
    }

    function stateMessage() {
      return {
        t: 'state', s: pipeState(), tool: S.tool, brush: S.brush, hardness: S.hardness,
        canUndo: !$('btnUndo').disabled, canRedo: !$('btnRedo').disabled, live: $('liveState').textContent
      };
    }

    function push() {
      if (!dlg || !ready || !S.src) return;
      if (sentFor !== S.src) { sentFor = S.src; send(imageMessage()); }
      if (timer) return;
      timer = setTimeout(function () { timer = null; if (dlg && ready && S.src) send(stateMessage()); }, 40);
    }

    function handle(m) {
      switch (m.t) {
        case 'ready': ready = true; sentFor = null; push(); break;
        case 'pick': if (S.src) pickAt({ x: m.x, y: m.y }, m.add); break;
        case 'stroke': if (S.src) { S.strokes.push(m.stroke); render(); setTimeout(commitNow, 0); } break;
        case 'tool': setTool(m.tool); break;
        case 'brush': S.brush = m.v; setSlider('brush', m.v); updateCursorSize(); push(); break;
        case 'hardness': S.hardness = m.v; setSlider('hardness', m.v); updateCursorSize(); push(); break;
        case 'undo': undo(); push(); break;
        case 'redo': redo(); push(); break;
        case 'close': close(); break;
      }
    }

    function updateButton() {
      var b = $('btnPopout');
      b.classList.toggle('on', !!dlg);
      b.dataset.tip = dlg ? 'Close the large preview window' : 'Open a large preview window';
    }

    function open() {
      if (dlg) { close(); return; }
      if (!supported()) { status('This version of PowerPoint cannot open a separate preview window.', 'err'); return; }
      var url = location.href.split('#')[0].replace(/taskpane\.html.*$/, 'preview.html');
      Office.context.ui.displayDialogAsync(url, { height: 85, width: 80, displayInIframe: false }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          status('Could not open the preview window' + (res.error ? ': ' + res.error.message : '.'), 'err');
          return;
        }
        dlg = res.value; ready = false;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (a) { receive(a.message); });
        dlg.addEventHandler(Office.EventType.DialogEventReceived, function () { dlg = null; ready = false; updateButton(); });
        updateButton();
      });
    }

    function close() {
      if (!dlg) return;
      try { dlg.close(); } catch (e) { /* allerede lukket */ }
      dlg = null; ready = false; updateButton();
    }

    return { open: open, close: close, push: push, supported: supported, isOpen: function () { return !!dlg; } };
  })();

  /* ---------------- Ikoner og verktoeytips ---------------- */

  var setIcon = UI.setIcon, initIcons = UI.initIcons, initTooltips = UI.initTooltips;

  async function start() {
    initIcons();
    initTooltips();
    buildSliders();
    bindControls();
    bindStage();
    bindHistory();
    bindPreviewToggle();
    bindCurves();
    syncAdvancedUI();
    $('btnPopout').addEventListener('click', function () { Popout.open(); });
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
    $('btnPopout').hidden = !Popout.supported();
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
