/*
 * PPTimagePoint bildemotor.
 * Rene funksjoner som jobber direkte paa pikseldata (RGBA, Uint8).
 * Brukes baade for rask forhaandsvisning (nedskalert kopi) og for
 * eksport i full opploesning.
 */
(function (global) {
  'use strict';

  var clamp = function (v, lo, hi) { return v < lo ? lo : (v > hi ? hi : v); };

  // sRGB -> lineaer, som oppslagstabell
  var LIN = new Float32Array(256);
  for (var i = 0; i < 256; i++) {
    var c = i / 255;
    LIN[i] = c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  }
  function linToSrgb8(v) {
    if (v <= 0) return 0;
    if (v >= 1) return 255;
    var s = v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
    return Math.round(s * 255);
  }

  function labF(t) { return t > 0.008856 ? Math.cbrt(t) : 7.787 * t + 16 / 116; }

  function rgbToLab(r, g, b) {
    var lr = LIN[r], lg = LIN[g], lb = LIN[b];
    var fx = labF((lr * 0.4124564 + lg * 0.3575761 + lb * 0.1804375) / 0.95047);
    var fy = labF(lr * 0.2126729 + lg * 0.7151522 + lb * 0.0721750);
    var fz = labF((lr * 0.0193339 + lg * 0.1191920 + lb * 0.9503041) / 1.08883);
    return { L: 116 * fy - 16, a: 500 * (fx - fy), b: 200 * (fy - fz) };
  }

  function makeImageData(w, h, data) {
    if (data) return new ImageData(data, w, h);
    return new ImageData(w, h);
  }

  function cloneImageData(img) {
    return makeImageData(img.width, img.height, new Uint8ClampedArray(img.data));
  }

  /** Gjennomsnittsfarge i et lite omraade rundt et normalisert punkt. */
  function sampleColor(img, nx, ny, radius) {
    var w = img.width, h = img.height, d = img.data;
    var cx = Math.round(clamp(nx, 0, 1) * (w - 1));
    var cy = Math.round(clamp(ny, 0, 1) * (h - 1));
    var r = radius == null ? 2 : radius;
    var sr = 0, sg = 0, sb = 0, n = 0;
    for (var y = Math.max(0, cy - r); y <= Math.min(h - 1, cy + r); y++) {
      for (var x = Math.max(0, cx - r); x <= Math.min(w - 1, cx + r); x++) {
        var p = (y * w + x) * 4;
        if (d[p + 3] < 16) continue;
        sr += d[p]; sg += d[p + 1]; sb += d[p + 2]; n++;
      }
    }
    if (!n) { var p0 = (cy * w + cx) * 4; sr = d[p0]; sg = d[p0 + 1]; sb = d[p0 + 2]; n = 1; }
    var col = { r: Math.round(sr / n), g: Math.round(sg / n), b: Math.round(sb / n) };
    var lab = rgbToLab(col.r, col.g, col.b);
    col.L = lab.L; col.A = lab.a; col.B = lab.b;
    return col;
  }

  /** Finner typisk bakgrunnsfarge ved aa se paa kantpikslene (median i Lab). */
  function detectBorderColor(img) {
    var w = img.width, h = img.height, d = img.data;
    var rs = [], gs = [], bs = [];
    var step = Math.max(1, Math.floor((w + h) / 400));
    function add(x, y) {
      var p = (y * w + x) * 4;
      if (d[p + 3] < 16) return;
      rs.push(d[p]); gs.push(d[p + 1]); bs.push(d[p + 2]);
    }
    for (var x = 0; x < w; x += step) { add(x, 0); add(x, h - 1); }
    for (var y = 0; y < h; y += step) { add(0, y); add(w - 1, y); }
    if (!rs.length) return null;
    var med = function (arr) { arr.sort(function (a, b) { return a - b; }); return arr[arr.length >> 1]; };
    var col = { r: med(rs), g: med(gs), b: med(bs) };
    var lab = rgbToLab(col.r, col.g, col.b);
    col.L = lab.L; col.A = lab.a; col.B = lab.b;
    return col;
  }

  /**
   * Avstandskart: for hver piksel minste fargeavstand (deltaE, Lab) til
   * en av noekkelfargene. Lagres som Uint8 med verdi = deltaE * 2.
   */
  function distanceMap(img, keys, lumWeight) {
    var w = img.width, h = img.height, d = img.data, n = w * h;
    var out = new Uint8Array(n);
    var wl = lumWeight == null ? 0.7 : lumWeight;
    var k = keys.length;
    var kL = new Float32Array(k), kA = new Float32Array(k), kB = new Float32Array(k);
    for (var j = 0; j < k; j++) { kL[j] = keys[j].L; kA[j] = keys[j].A; kB[j] = keys[j].B; }
    // Mellomlager for gjentatte farger (vanlig paa ensfargede bakgrunner)
    var lastKey = -1, lastVal = 0;
    for (var i = 0, p = 0; i < n; i++, p += 4) {
      var r = d[p], g = d[p + 1], b = d[p + 2];
      var key = (r << 16) | (g << 8) | b;
      if (key === lastKey) { out[i] = lastVal; continue; }
      var lr = LIN[r], lg = LIN[g], lb = LIN[b];
      var fx = labF((lr * 0.4124564 + lg * 0.3575761 + lb * 0.1804375) / 0.95047);
      var fy = labF(lr * 0.2126729 + lg * 0.7151522 + lb * 0.0721750);
      var fz = labF((lr * 0.0193339 + lg * 0.1191920 + lb * 0.9503041) / 1.08883);
      var L = 116 * fy - 16, A = 500 * (fx - fy), B = 200 * (fy - fz);
      var best = 1e9;
      for (j = 0; j < k; j++) {
        var dl = (L - kL[j]) * wl, da = A - kA[j], db = B - kB[j];
        var dd = dl * dl + da * da + db * db;
        if (dd < best) best = dd;
      }
      var v = Math.sqrt(best) * 2;
      v = v > 255 ? 255 : (v + 0.5) | 0;
      out[i] = v; lastKey = key; lastVal = v;
    }
    return out;
  }

  /**
   * Finner hvilke piksler som er "bakgrunn" og henger sammen med
   * startpunktene (klikk og/eller bildekantene). 1 = bakgrunnsomraade.
   */
  function computeRegion(dist, w, h, opts) {
    var n = w * h;
    var region = new Uint8Array(n);
    var limit = Math.min(255, Math.round((opts.tolerance + opts.softness) * 2));
    if (opts.mode === 'global') {
      for (var i = 0; i < n; i++) region[i] = dist[i] <= limit ? 1 : 0;
      return region;
    }
    var stack = new Int32Array(n);
    var sp = 0;
    function push(idx) {
      if (region[idx] || dist[idx] > limit) return;
      region[idx] = 1; stack[sp++] = idx;
    }
    if (opts.fromEdges) {
      for (var x = 0; x < w; x++) { push(x); push((h - 1) * w + x); }
      for (var y = 0; y < h; y++) { push(y * w); push(y * w + w - 1); }
    }
    var seeds = opts.seeds || [];
    var rr = Math.max(2, Math.round(Math.max(w, h) / 300));
    for (var s = 0; s < seeds.length; s++) {
      var sx = Math.round(clamp(seeds[s].x, 0, 1) * (w - 1));
      var sy = Math.round(clamp(seeds[s].y, 0, 1) * (h - 1));
      var found = -1, bestD = 256;
      for (var yy = Math.max(0, sy - rr); yy <= Math.min(h - 1, sy + rr); yy++) {
        for (var xx = Math.max(0, sx - rr); xx <= Math.min(w - 1, sx + rr); xx++) {
          var id = yy * w + xx;
          if (dist[id] <= limit && dist[id] < bestD) { bestD = dist[id]; found = id; }
        }
      }
      if (found >= 0) push(found);
    }
    while (sp > 0) {
      var idx = stack[--sp];
      var px = idx % w;
      if (px > 0) push(idx - 1);
      if (px < w - 1) push(idx + 1);
      if (idx >= w) push(idx - w);
      if (idx < n - w) push(idx + w);
    }
    return region;
  }

  /** Skalerer et omraadekart (fra forhaandsvisning) opp til full opploesning. */
  function upscaleRegion(region, pw, ph, W, H) {
    var grown = minMaxFilter(region, pw, ph, 1, false);
    var out = new Uint8Array(W * H);
    var xmap = new Int32Array(W);
    for (var x = 0; x < W; x++) xmap[x] = Math.min(pw - 1, Math.floor(x * pw / W));
    for (var y = 0; y < H; y++) {
      var row = Math.min(ph - 1, Math.floor(y * ph / H)) * pw;
      var o = y * W;
      for (x = 0; x < W; x++) out[o + x] = grown[row + xmap[x]];
    }
    return out;
  }

  /** Alfa (0..255) ut fra avstand: under toleranse = gjennomsiktig, myk overgang. */
  function alphaFromDist(dist, region, tolerance, softness) {
    var n = dist.length;
    var out = new Uint8Array(n);
    var lut = new Uint8Array(256);
    var soft = Math.max(0.0001, softness);
    for (var v = 0; v < 256; v++) {
      var t = clamp((v / 2 - tolerance) / soft, 0, 1);
      if (softness <= 0) t = v / 2 <= tolerance ? 0 : 1;
      t = t * t * (3 - 2 * t);
      lut[v] = Math.round(t * 255);
    }
    for (var i = 0; i < n; i++) out[i] = region[i] ? lut[dist[i]] : 255;
    return out;
  }

  /**
   * Min/maks-filter (erosjon/dilatasjon) med kvadratisk vindu, O(n) per
   * radius takket vaere van Herk / Gil-Werman.
   */
  function minMaxFilter(src, w, h, r, isMin) {
    if (r <= 0) return src;
    var k = 2 * r + 1;
    var tmp = new Uint8Array(w * h);
    var out = new Uint8Array(w * h);
    var maxLen = Math.max(w, h) + 2 * r;
    var L = Math.ceil(maxLen / k) * k;
    var line = new Uint8Array(L), g = new Uint8Array(L), hh = new Uint8Array(L);
    var pad = isMin ? 255 : 0;

    function run(input, output, len, count, stride, step) {
      for (var c = 0; c < count; c++) {
        var base = c * stride;
        line.fill(pad);
        for (var i = 0; i < len; i++) line[i + r] = input[base + i * step];
        var total = len + 2 * r;
        var blocks = Math.ceil(total / k) * k;
        var i2, v, bstart;
        for (bstart = 0; bstart < blocks; bstart += k) {
          v = line[bstart]; g[bstart] = v;
          for (i2 = bstart + 1; i2 < bstart + k; i2++) {
            var x = line[i2];
            v = isMin ? (x < v ? x : v) : (x > v ? x : v);
            g[i2] = v;
          }
          v = line[bstart + k - 1]; hh[bstart + k - 1] = v;
          for (i2 = bstart + k - 2; i2 >= bstart; i2--) {
            var y = line[i2];
            v = isMin ? (y < v ? y : v) : (y > v ? y : v);
            hh[i2] = v;
          }
        }
        for (i = 0; i < len; i++) {
          var a = hh[i], b = g[i + 2 * r];
          output[base + i * step] = isMin ? (a < b ? a : b) : (a > b ? a : b);
        }
      }
    }
    // Rader
    run(src, tmp, w, h, w, 1);
    // Kolonner (stride 1 mellom kolonner, step w innen kolonnen)
    run(tmp, out, h, w, 1, w);
    return out;
  }

  /** Boks-uskarphet paa en Uint8-kanal, med kantklemming. */
  function boxBlurU8(src, w, h, r) {
    if (r <= 0) return src;
    var tmp = new Uint8Array(w * h), out = new Uint8Array(w * h);
    var div = 2 * r + 1;
    var x, y, sum, o;
    for (y = 0; y < h; y++) {
      o = y * w; sum = 0;
      for (x = -r; x <= r; x++) sum += src[o + clamp(x, 0, w - 1)];
      for (x = 0; x < w; x++) {
        tmp[o + x] = (sum / div + 0.5) | 0;
        sum += src[o + Math.min(w - 1, x + r + 1)] - src[o + Math.max(0, x - r)];
      }
    }
    for (x = 0; x < w; x++) {
      sum = 0;
      for (y = -r; y <= r; y++) sum += tmp[clamp(y, 0, h - 1) * w + x];
      for (y = 0; y < h; y++) {
        out[y * w + x] = (sum / div + 0.5) | 0;
        sum += tmp[Math.min(h - 1, y + r + 1) * w + x] - tmp[Math.max(0, y - r) * w + x];
      }
    }
    return out;
  }

  function blurU8(src, w, h, radius, passes) {
    var r = Math.max(0, Math.round(radius));
    if (!r) return src;
    var out = src;
    for (var p = 0; p < (passes || 3); p++) out = boxBlurU8(out, w, h, r);
    return out;
  }

  /** Krymp (positiv) eller utvid (negativ) kanten, og myk den opp. */
  function refineAlpha(alpha, w, h, choke, feather) {
    var out = alpha;
    var c = Math.round(choke);
    if (c > 0) out = minMaxFilter(out, w, h, c, true);
    else if (c < 0) out = minMaxFilter(out, w, h, -c, false);
    if (feather > 0) {
      var fr = Math.max(1, Math.round(feather / 1.7));
      out = blurU8(out, w, h, fr, 3);
    }
    return out;
  }

  /**
   * Penselstroek: { mode: 'erase' | 'restore', r: radius som andel av
   * lengste side, pts: [{x,y}] normalisert }.
   */
  function applyStrokes(alpha, w, h, strokes) {
    if (!strokes || !strokes.length) return alpha;
    var out = new Uint8Array(alpha);
    var cover = new Uint8Array(w * h);
    var longSide = Math.max(w, h);
    for (var s = 0; s < strokes.length; s++) {
      var st = strokes[s];
      var R = Math.max(1, st.r * longSide);
      // Hardhet 0..1: andel av radiusen som er helt dekket. Resten toner mykt ut.
      var hard = st.h == null ? 0.65 : clamp(st.h, 0, 1);
      var inner = R * hard;
      var pts = densify(st.pts, w, h, Math.max(1, R * 0.25));
      var minX = w, minY = h, maxX = -1, maxY = -1;
      for (var p = 0; p < pts.length; p++) {
        var cx = pts[p][0], cy = pts[p][1];
        var x0 = Math.max(0, Math.floor(cx - R)), x1 = Math.min(w - 1, Math.ceil(cx + R));
        var y0 = Math.max(0, Math.floor(cy - R)), y1 = Math.min(h - 1, Math.ceil(cy + R));
        if (x0 < minX) minX = x0; if (y0 < minY) minY = y0;
        if (x1 > maxX) maxX = x1; if (y1 > maxY) maxY = y1;
        for (var y = y0; y <= y1; y++) {
          var dy = y - cy;
          for (var x = x0; x <= x1; x++) {
            var dx = x - cx;
            var dd = Math.sqrt(dx * dx + dy * dy);
            if (dd > R) continue;
            var t = dd <= inner ? 1 : 1 - (dd - inner) / Math.max(0.5, R - inner);
            if (t < 1) t = t * t * (3 - 2 * t);
            var v = Math.round(t * 255);
            var id = y * w + x;
            if (v > cover[id]) cover[id] = v;
          }
        }
      }
      if (maxX < 0) continue;
      for (y = minY; y <= maxY; y++) {
        for (x = minX; x <= maxX; x++) {
          id = y * w + x;
          var cv = cover[id];
          if (!cv) continue;
          cover[id] = 0;
          if (st.mode === 'erase') out[id] = Math.round(out[id] * (255 - cv) / 255);
          else out[id] = Math.round(out[id] + (255 - out[id]) * cv / 255);
        }
      }
    }
    return out;
  }

  function densify(pts, w, h, spacing) {
    var res = [];
    for (var i = 0; i < pts.length; i++) {
      var x = pts[i].x * (w - 1), y = pts[i].y * (h - 1);
      if (i === 0) { res.push([x, y]); continue; }
      var px = pts[i - 1].x * (w - 1), py = pts[i - 1].y * (h - 1);
      var dist = Math.hypot(x - px, y - py);
      var steps = Math.max(1, Math.ceil(dist / spacing));
      for (var s = 1; s <= steps; s++) res.push([px + (x - px) * s / steps, py + (y - py) * s / steps]);
    }
    return res;
  }

  /* ---------------- Fargejusteringer ---------------- */

  var DEFAULT_ADJUST = {
    exposure: 0, brightness: 0, contrast: 0, highlights: 0, shadows: 0,
    temperature: 0, tint: 0, hue: 0, saturation: 0, vibrance: 0,
    clarity: 0, sharpness: 0
  };

  function isNeutral(p) {
    for (var k in DEFAULT_ADJUST) if (p[k]) return false;
    return true;
  }

  function buildToneLUTs(p) {
    var luts = [new Uint8Array(256), new Uint8Array(256), new Uint8Array(256)];
    var expo = Math.pow(2, p.exposure || 0);
    var t = (p.temperature || 0) / 100, m = (p.tint || 0) / 100;
    var gains = [expo * (1 + 0.22 * t), expo * (1 - 0.18 * m), expo * (1 - 0.22 * t)];
    // normaliser slik at lysstyrken holdes omtrent lik ved fargetemperatur
    var lum = 0.2126 * gains[0] + 0.7152 * gains[1] + 0.0722 * gains[2];
    for (var c = 0; c < 3; c++) gains[c] = gains[c] * expo / lum;
    var gamma = Math.pow(2, -(p.brightness || 0) / 100);
    var con = (p.contrast || 0) / 100;
    var k = con >= 0 ? 1 + con * 1.2 : 1 + con * 0.8;
    var hi = (p.highlights || 0) / 100, sh = (p.shadows || 0) / 100;
    for (c = 0; c < 3; c++) {
      for (var i = 0; i < 256; i++) {
        var v = linToSrgb8(LIN[i] * gains[c]) / 255;
        v = Math.pow(v, gamma);
        if (sh) v += sh * 0.35 * (1 - v) * (1 - v) * v * 6.75;
        if (hi) v += hi * 0.35 * v * v * (1 - v) * 6.75;
        v = clamp(v, 0, 1);
        if (con) {
          v = 0.5 + (v - 0.5) * k;
          if (con > 0) {
            // litt myk S-kurve i tillegg for et mer naturlig resultat
            var s = clamp(v, 0, 1);
            v = s + con * 0.3 * (s * s * (3 - 2 * s) - s);
          }
        }
        luts[c][i] = Math.round(clamp(v, 0, 1) * 255);
      }
    }
    return luts;
  }

  function colorMatrix(p) {
    var a = (p.hue || 0) * Math.PI / 180;
    var cos = Math.cos(a), sin = Math.sin(a);
    var H = [
      0.213 + cos * 0.787 - sin * 0.213, 0.715 - cos * 0.715 - sin * 0.715, 0.072 - cos * 0.072 + sin * 0.928,
      0.213 - cos * 0.213 + sin * 0.143, 0.715 + cos * 0.285 + sin * 0.140, 0.072 - cos * 0.072 - sin * 0.283,
      0.213 - cos * 0.213 - sin * 0.787, 0.715 - cos * 0.715 + sin * 0.715, 0.072 + cos * 0.928 + sin * 0.072
    ];
    var s = 1 + (p.saturation || 0) / 100;
    var lr = 0.2126, lg = 0.7152, lb = 0.0722;
    var S = [
      lr * (1 - s) + s, lg * (1 - s), lb * (1 - s),
      lr * (1 - s), lg * (1 - s) + s, lb * (1 - s),
      lr * (1 - s), lg * (1 - s), lb * (1 - s) + s
    ];
    var M = new Float32Array(9);
    for (var r = 0; r < 3; r++) for (var c = 0; c < 3; c++) {
      M[r * 3 + c] = S[r * 3] * H[c] + S[r * 3 + 1] * H[3 + c] + S[r * 3 + 2] * H[6 + c];
    }
    return M;
  }

  /** Bruker alle fargejusteringer. scale = lengste side / 1000 (for radiuser). */
  function adjust(img, p, scale) {
    var out = cloneImageData(img);
    if (isNeutral(p)) return out;
    var w = img.width, h = img.height, n = w * h, d = out.data;
    var i, q;
    // 1. Tonekurver (eksponering, temperatur, lysstyrke, kontrast, hoeylys, skygger)
    if (p.exposure || p.temperature || p.tint || p.brightness || p.contrast || p.highlights || p.shadows) {
      var L = buildToneLUTs(p);
      var lr = L[0], lg = L[1], lb = L[2];
      for (q = 0; q < d.length; q += 4) { d[q] = lr[d[q]]; d[q + 1] = lg[d[q + 1]]; d[q + 2] = lb[d[q + 2]]; }
    }
    // 2. Klarhet (lokal kontrast i mellomtonene)
    if (p.clarity) {
      var Y = new Uint8Array(n);
      for (i = 0, q = 0; i < n; i++, q += 4) Y[i] = (d[q] * 54 + d[q + 1] * 183 + d[q + 2] * 19) >> 8;
      var blurred = blurU8(Y, w, h, Math.max(2, 14 * scale), 3);
      var amt = p.clarity / 100;
      for (i = 0, q = 0; i < n; i++, q += 4) {
        var yv = Y[i] / 255;
        var mid = 1 - (2 * yv - 1) * (2 * yv - 1);
        var delta = (Y[i] - blurred[i]) * amt * (0.35 + 0.9 * mid);
        d[q] += delta; d[q + 1] += delta; d[q + 2] += delta;
      }
    }
    // 3. Fargetone, metning og livfullhet
    if (p.hue || p.saturation || p.vibrance) {
      var M = colorMatrix(p);
      var vib = (p.vibrance || 0) / 100;
      for (q = 0; q < d.length; q += 4) {
        var r = d[q], g = d[q + 1], b = d[q + 2];
        var nr = M[0] * r + M[1] * g + M[2] * b;
        var ng = M[3] * r + M[4] * g + M[5] * b;
        var nb = M[6] * r + M[7] * g + M[8] * b;
        if (vib) {
          var mx = Math.max(nr, ng, nb), mn = Math.min(nr, ng, nb);
          var sat = mx > 0 ? (mx - mn) / Math.max(1, mx) : 0;
          var f = 1 + vib * (1 - clamp(sat, 0, 1)) * 1.2;
          var yy = 0.2126 * nr + 0.7152 * ng + 0.0722 * nb;
          nr = yy + (nr - yy) * f; ng = yy + (ng - yy) * f; nb = yy + (nb - yy) * f;
        }
        d[q] = nr; d[q + 1] = ng; d[q + 2] = nb;
      }
    }
    // 4. Skarphet (uskarp maske)
    if (p.sharpness > 0) {
      var rad = Math.max(1, Math.round(0.9 * scale));
      var amount = p.sharpness / 100 * 1.6;
      var ch = new Uint8Array(n);
      for (var c = 0; c < 3; c++) {
        for (i = 0, q = c; i < n; i++, q += 4) ch[i] = d[q];
        var bl = blurU8(ch, w, h, rad, 2);
        for (i = 0, q = c; i < n; i++, q += 4) d[q] = ch[i] + (ch[i] - bl[i]) * amount;
      }
    }
    return out;
  }

  /** Kjoerer en enkelt farge gjennom de globale justeringene (for kant-opprydding). */
  function adjustColor(col, p) {
    var img = makeImageData(1, 1);
    img.data[0] = col.r; img.data[1] = col.g; img.data[2] = col.b; img.data[3] = 255;
    var q = Object.assign({}, p, { clarity: 0, sharpness: 0 });
    var o = adjust(img, q, 1).data;
    return { r: o[0], g: o[1], b: o[2] };
  }

  /**
   * Setter sammen farge og alfa. Med defringe fjernes bakgrunnsfargen som
   * "lekker" inn i halvgjennomsiktige kantpiksler.
   */
  function compose(colorImg, alpha, keyRgb, defringe) {
    var out = cloneImageData(colorImg);
    if (!alpha) return out;
    var d = out.data, n = alpha.length;
    var k = keyRgb && defringe ? keyRgb.length : 0;
    for (var i = 0, q = 0; i < n; i++, q += 4) {
      var a = alpha[i];
      if (k && a > 0 && a < 250) {
        var r = d[q], g = d[q + 1], b = d[q + 2];
        var best = 0, bd = 1e9;
        for (var j = 0; j < k; j++) {
          var dr = r - keyRgb[j].r, dg = g - keyRgb[j].g, db = b - keyRgb[j].b;
          var dd = dr * dr + dg * dg + db * db;
          if (dd < bd) { bd = dd; best = j; }
        }
        var K = keyRgb[best];
        var af = Math.max(a / 255, 0.2);
        d[q] = (r - (1 - af) * K.r) / af;
        d[q + 1] = (g - (1 - af) * K.g) / af;
        d[q + 2] = (b - (1 - af) * K.b) / af;
      }
      d[q + 3] = (d[q + 3] * a + 127) / 255;
    }
    return out;
  }

  /* ---------------- Skalering (Lanczos3) ---------------- */

  function lanczos(x) {
    if (x === 0) return 1;
    if (x <= -3 || x >= 3) return 0;
    var px = Math.PI * x;
    return 3 * Math.sin(px) * Math.sin(px / 3) / (px * px);
  }

  function buildWeights(srcLen, dstLen) {
    var scale = dstLen / srcLen;
    var fs = Math.max(1, 1 / scale);
    var support = 3 * fs;
    var maxTaps = Math.ceil(support * 2) + 2;
    var starts = new Int32Array(dstLen), counts = new Int32Array(dstLen);
    var weights = new Float32Array(dstLen * maxTaps);
    for (var i = 0; i < dstLen; i++) {
      var center = (i + 0.5) / scale;
      var left = Math.max(0, Math.floor(center - support));
      var right = Math.min(srcLen - 1, Math.ceil(center + support));
      var cnt = Math.min(maxTaps, right - left + 1);
      var sum = 0;
      for (var j = 0; j < cnt; j++) {
        var wv = lanczos((left + j + 0.5 - center) / fs);
        weights[i * maxTaps + j] = wv; sum += wv;
      }
      if (sum) for (j = 0; j < cnt; j++) weights[i * maxTaps + j] /= sum;
      starts[i] = left; counts[i] = cnt;
    }
    return { starts: starts, counts: counts, weights: weights, taps: maxTaps };
  }

  function resize(img, nw, nh) {
    nw = Math.max(1, Math.round(nw)); nh = Math.max(1, Math.round(nh));
    var w = img.width, h = img.height;
    if (nw === w && nh === h) return cloneImageData(img);
    var s = img.data;
    var wx = buildWeights(w, nw), wy = buildWeights(h, nh);
    // Horisontal pass med forhaandsmultiplisert alfa
    var tmp = new Float32Array(nw * h * 4);
    var x, y, j, o, a;
    for (y = 0; y < h; y++) {
      var row = y * w * 4;
      for (x = 0; x < nw; x++) {
        var st = wx.starts[x], cnt = wx.counts[x], base = x * wx.taps;
        var r = 0, g = 0, b = 0, al = 0;
        for (j = 0; j < cnt; j++) {
          var wv = wx.weights[base + j];
          var p = row + (st + j) * 4;
          a = s[p + 3] * wv;
          r += s[p] * a; g += s[p + 1] * a; b += s[p + 2] * a; al += a;
        }
        o = (y * nw + x) * 4;
        tmp[o] = r; tmp[o + 1] = g; tmp[o + 2] = b; tmp[o + 3] = al;
      }
    }
    var out = makeImageData(nw, nh);
    var d = out.data;
    for (y = 0; y < nh; y++) {
      var sy = wy.starts[y], cy = wy.counts[y], by = y * wy.taps;
      for (x = 0; x < nw; x++) {
        var R = 0, G = 0, B = 0, A = 0;
        for (j = 0; j < cy; j++) {
          var w2 = wy.weights[by + j];
          var q = ((sy + j) * nw + x) * 4;
          R += tmp[q] * w2; G += tmp[q + 1] * w2; B += tmp[q + 2] * w2; A += tmp[q + 3] * w2;
        }
        o = (y * nw + x) * 4;
        if (A > 0.5) {
          d[o] = R / A; d[o + 1] = G / A; d[o + 2] = B / A; d[o + 3] = A;
        } else {
          d[o] = 0; d[o + 1] = 0; d[o + 2] = 0; d[o + 3] = 0;
        }
      }
    }
    return out;
  }

  /** Finner rammen rundt alt som ikke er gjennomsiktig. */
  function opaqueBounds(img, threshold) {
    var w = img.width, h = img.height, d = img.data, t = threshold == null ? 8 : threshold;
    var minX = w, minY = h, maxX = -1, maxY = -1;
    for (var y = 0; y < h; y++) {
      var o = y * w * 4;
      for (var x = 0; x < w; x++) {
        if (d[o + x * 4 + 3] > t) {
          if (x < minX) minX = x; if (x > maxX) maxX = x;
          if (y < minY) minY = y; if (y > maxY) maxY = y;
        }
      }
    }
    if (maxX < 0) return null;
    return { x: minX, y: minY, w: maxX - minX + 1, h: maxY - minY + 1 };
  }

  function crop(img, r) {
    var out = makeImageData(r.w, r.h);
    var s = img.data, d = out.data, w = img.width;
    for (var y = 0; y < r.h; y++) {
      var so = ((r.y + y) * w + r.x) * 4;
      d.set(s.subarray(so, so + r.w * 4), y * r.w * 4);
    }
    return out;
  }

  function hasTransparency(img) {
    var d = img.data;
    for (var q = 3; q < d.length; q += 4) if (d[q] < 255) return true;
    return false;
  }

  global.ImageOps = {
    DEFAULT_ADJUST: DEFAULT_ADJUST,
    rgbToLab: rgbToLab,
    sampleColor: sampleColor,
    detectBorderColor: detectBorderColor,
    distanceMap: distanceMap,
    computeRegion: computeRegion,
    upscaleRegion: upscaleRegion,
    alphaFromDist: alphaFromDist,
    minMaxFilter: minMaxFilter,
    blurU8: blurU8,
    refineAlpha: refineAlpha,
    applyStrokes: applyStrokes,
    adjust: adjust,
    adjustColor: adjustColor,
    isNeutral: isNeutral,
    compose: compose,
    resize: resize,
    opaqueBounds: opaqueBounds,
    crop: crop,
    hasTransparency: hasTransparency,
    cloneImageData: cloneImageData
  };
})(typeof window !== 'undefined' ? window : self);
