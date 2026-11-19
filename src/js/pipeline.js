/*
 * Felles behandlingskjede for forhaandsvisning og full opploesning.
 * Brukes av panelet (app.js) og av det store forhaandsvisningsvinduet (preview.js).
 *
 * state: { adjust, key, strokes, crop, fullW, fullH }
 */
(function (global) {
  'use strict';

  var Ops = global.ImageOps;

  function scaleOf(img) { return Math.max(img.width, img.height) / 1000; }

  function keyActive(state) { return !!(state.key && state.key.enabled && state.key.colors && state.key.colors.length); }

  function keyParamsKey(state) {
    var k = state.key;
    var rgb = function (c) { return [c.r, c.g, c.b]; };
    return JSON.stringify([k.colors.map(rgb), (k.keep || []).map(rgb), k.seeds, k.tolerance, k.softness, k.mode, k.fromEdges, k.choke, k.feather]);
  }

  /**
   * Avstand til bakgrunnsfargene. Piksler som ligner mer paa en "behold"-farge
   * (pipette minus) enn paa bakgrunnen, regnes ikke som bakgrunn.
   */
  function keyDistance(img, k) {
    var dist = Ops.distanceMap(img, k.colors);
    if (k.keep && k.keep.length) {
      var kd = Ops.distanceMap(img, k.keep);
      for (var i = 0; i < dist.length; i++) if (kd[i] <= dist[i]) dist[i] = 255;
    }
    return dist;
  }

  /**
   * Alfakart for bildet. Med opts.proxy og opts.proxyRegion brukes omraadet
   * fra forhaandsvisningen (slik at full opploesning blir lik det du saa).
   */
  function alphaFor(img, state, opts) {
    if (!keyActive(state)) return { alpha: null, region: null };
    opts = opts || {};
    var w = img.width, h = img.height, s = scaleOf(img), k = state.key;
    var dist = keyDistance(img, k);
    var region;
    if (!opts.proxy || k.mode === 'global') {
      region = Ops.computeRegion(dist, w, h, k);
    } else {
      var P = opts.proxy;
      var pr = opts.proxyRegion || Ops.computeRegion(keyDistance(P, k), P.width, P.height, k);
      region = Ops.upscaleRegion(pr, P.width, P.height, w, h);
    }
    var alpha = Ops.alphaFromDist(dist, region, k.tolerance, k.softness);
    alpha = Ops.refineAlpha(alpha, w, h, k.choke * s, k.feather * s);
    return { alpha: alpha, region: opts.proxy ? null : region };
  }

  function finalize(colorImg, baseAlpha, state) {
    var w = colorImg.width, h = colorImg.height;
    var alpha = baseAlpha;
    if (state.strokes && state.strokes.length) {
      if (!alpha) { alpha = new Uint8Array(w * h); alpha.fill(255); }
      alpha = Ops.applyStrokes(alpha, w, h, state.strokes);
    }
    var keyRgb = keyActive(state) ? state.key.colors.map(function (c) { return Ops.adjustColor(c, state.adjust); }) : null;
    return Ops.compose(colorImg, alpha, keyRgb, state.key && state.key.defringe);
  }

  /** Forhaandsvisning med mellomlager, slik at bare det som er endret regnes ut paa nytt. */
  function Previewer(proxy) {
    this.P = proxy;
    this.adjKey = null; this.adjusted = null;
    this.maskKey = null; this.alpha = null; this.region = null;
    this.last = null;
  }

  Previewer.prototype.render = function (state) {
    var P = this.P;
    var ak = JSON.stringify(state.adjust);
    if (this.adjKey !== ak) { this.adjusted = Ops.adjust(P, state.adjust, scaleOf(P)); this.adjKey = ak; }
    var kk = keyActive(state) ? keyParamsKey(state) : 'off';
    if (this.maskKey !== kk) {
      var r = alphaFor(P, state, null);
      this.alpha = r.alpha; this.region = r.region; this.maskKey = kk;
    }
    this.last = finalize(this.adjusted, this.alpha, state);
    return this.last;
  };

  /** Tegner til et lerret (som har samme stoerrelse som forhaandsvisningen). */
  Previewer.prototype.draw = function (ctx, state, compare) {
    var P = this.P;
    ctx.clearRect(0, 0, P.width, P.height);
    if (compare) { ctx.putImageData(P, 0, 0); return; }
    ctx.putImageData(this.render(state), 0, 0);
    drawCropOverlay(ctx, P, state);
  };

  /** Demper omraadet som er beskaaret bort i PowerPoint. */
  function drawCropOverlay(ctx, P, state) {
    var c = state.crop, W = state.fullW, H = state.fullH;
    if (!c || !W || (!c.x && !c.y && c.w === W && c.h === H)) return;
    var k = P.width / W;
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

  global.Pipeline = {
    scaleOf: scaleOf,
    keyActive: keyActive,
    keyParamsKey: keyParamsKey,
    alphaFor: alphaFor,
    finalize: finalize,
    Previewer: Previewer
  };
})(window);
