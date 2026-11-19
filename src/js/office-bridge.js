/*
 * Alt som snakker med PowerPoint. Hver funksjon sjekker at API-et finnes
 * foer det brukes, slik at tillegget faller pent tilbake paa eldre versjoner.
 */
(function (global) {
  'use strict';

  var Bridge = {
    available: false,
    host: null,

    init: function () {
      return new Promise(function (resolve) {
        if (typeof Office === 'undefined' || !Office.onReady) { resolve(false); return; }
        var done = false;
        var timer = setTimeout(function () { if (!done) { done = true; resolve(false); } }, 4000);
        Office.onReady(function (info) {
          if (done) return;
          done = true; clearTimeout(timer);
          Bridge.host = info && info.host;
          Bridge.available = !!(info && info.host && Office.HostType && info.host === Office.HostType.PowerPoint);
          resolve(Bridge.available);
        });
      });
    },

    supports: function (version) {
      try { return Office.context.requirements.isSetSupported('PowerPointApi', version); }
      catch (e) { return false; }
    },

    onSelectionChanged: function (cb) {
      try {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, cb);
      } catch (e) { /* ikke kritisk */ }
    },

    /** Rask sjekk av markeringen: { id } for markert bilde, ellers null. */
    getSelectedImageId: async function () {
      if (!Bridge.available || !global.PowerPoint || !Bridge.supports('1.5')) return null;
      return PowerPoint.run(async function (ctx) {
        var shapes = ctx.presentation.getSelectedShapes();
        shapes.load('items/id,items/type');
        await ctx.sync();
        var img = shapes.items.filter(function (s) { return s.type === 'Image'; })[0];
        if (!img) return null;
        return { id: img.id, sessionId: await readTag(ctx, img) };
      });
    },

    /** Plassering og stoerrelse (punkter) for en figur, eller null hvis den er slettet. */
    getShapeGeometry: async function (meta) {
      if (!Bridge.available || !global.PowerPoint || !Bridge.supports('1.3')) return undefined;
      return PowerPoint.run(async function (ctx) {
        var sh = ctx.presentation.slides.getItem(meta.slideId).shapes.getItemOrNullObject(meta.shapeId);
        sh.load('left,top,width,height');
        await ctx.sync();
        if (sh.isNullObject) return null;
        return { left: sh.left, top: sh.top, width: sh.width, height: sh.height };
      });
    },

    /** Beskjaeringen (andeler l/t/r/b) som er satt paa bildet i PowerPoint. */
    getShapeCrop: async function (meta) {
      if (!Bridge.available || !global.PowerPoint || !Bridge.supports('1.8')) return null;
      return PowerPoint.run(async function (ctx) {
        var exported = ctx.presentation.slides.getItem(meta.slideId).exportAsBase64();
        await ctx.sync();
        var info = await PptxExtract.extractPicture(exported.value, { id: meta.shapeId, name: meta.name, infoOnly: true });
        return info.crop;
      });
    },

    /**
     * Henter markert bilde. Returnerer { bytes, mime, crop, name, meta }.
     * meta beskriver plassering paa lysbildet (punkter) for aa kunne erstatte.
     */
    getSelectedImage: async function () {
      if (!Bridge.available) throw new Error('PowerPoint is not available.');
      var firstError = null;
      if (Bridge.supports('1.8') && global.PowerPoint) {
        try { return await Bridge._viaSlideExport(); }
        catch (e) { firstError = e; if (e && e.userFacing) throw e; }
      }
      try { return await Bridge._viaCoercion(); }
      catch (e2) {
        throw firstError || e2;
      }
    },

    _viaSlideExport: function () {
      return PowerPoint.run(async function (ctx) {
        var shapes = ctx.presentation.getSelectedShapes();
        shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
        var slides = ctx.presentation.getSelectedSlides();
        slides.load('items/id');
        await ctx.sync();
        if (!shapes.items.length) throw userError('Select a picture on the slide first.');
        var shape = shapes.items.filter(function (s) { return s.type === 'Image'; })[0];
        if (!shape) throw userError('The selected object is not a picture. Select a picture (not a group or a shape with a picture fill).');
        if (!slides.items.length) throw userError('Could not find the slide for this picture.');
        var slide = slides.items[0];
        var exported = slide.exportAsBase64();
        await ctx.sync();
        var pic = await PptxExtract.extractPicture(exported.value, { id: shape.id, name: shape.name });
        pic.meta = {
          slideId: slide.id, shapeId: shape.id, name: shape.name,
          left: shape.left, top: shape.top, width: shape.width, height: shape.height,
          sessionId: await readTag(ctx, shape)
        };
        return pic;
      });
    },

    _viaCoercion: function () {
      return new Promise(function (resolve, reject) {
        try {
          Office.context.document.getSelectedDataAsync(Office.CoercionType.Image, function (res) {
            if (res.status !== Office.AsyncResultStatus.Succeeded) {
              reject(userError('This version of PowerPoint does not let add-ins read the picture directly. Use "Open file" or paste the image (Ctrl+V) into the pane.'));
              return;
            }
            var bin = atob(res.value);
            var bytes = new Uint8Array(bin.length);
            for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            resolve({ bytes: bytes, mime: 'image/png', crop: { l: 0, t: 0, r: 0, b: 0 }, name: 'image', meta: null });
          });
        } catch (e) {
          reject(userError('Could not read the selection. Use "Open file" instead.'));
        }
      });
    },

    /**
     * Setter inn bildet (base64 uten prefiks). Med placement legges det paa
     * samme sted; med replaceMeta slettes originalen etterpaa.
     */
    insertImage: async function (base64, placement, replaceMeta, sessionId) {
      if (!Bridge.available) throw new Error('PowerPoint is not available.');
      var opts = { coercionType: Office.CoercionType.Image };
      if (placement) {
        opts.imageLeft = placement.left;
        opts.imageTop = placement.top;
        opts.imageWidth = placement.width;
        opts.imageHeight = placement.height;
      }
      var canRun = Bridge.supports('1.3') && global.PowerPoint;

      // Hvilke figurer ligger paa lysbildet foer innsetting? Da finner vi det nye bildet sikkert etterpaa.
      var before = null, slideId = replaceMeta && replaceMeta.slideId;
      if (canRun) {
        try {
          before = await PowerPoint.run(async function (ctx) {
            if (!slideId && Bridge.supports('1.5')) {
              var sl = ctx.presentation.getSelectedSlides();
              sl.load('items/id');
              await ctx.sync();
              slideId = sl.items.length ? sl.items[0].id : null;
            }
            if (!slideId) return null;
            var shapes = ctx.presentation.slides.getItem(slideId).shapes;
            shapes.load('items/id');
            await ctx.sync();
            var ids = {};
            shapes.items.forEach(function (x) { ids[x.id] = true; });
            return ids;
          });
        } catch (e) { before = null; }
      }

      await new Promise(function (resolve, reject) {
        Office.context.document.setSelectedDataAsync(base64, opts, function (res) {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error('PowerPoint rejected the image: ' + (res.error && res.error.message)));
        });
      });

      if (!canRun) return { deleted: false, meta: null };
      try {
        return await PowerPoint.run(async function (ctx) {
          var s = null, sid = slideId;
          // 1. Ny bildefigur paa lysbildet som ikke fantes foer
          if (before && sid) {
            var shapes = ctx.presentation.slides.getItem(sid).shapes;
            shapes.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
            await ctx.sync();
            var fresh = shapes.items.filter(function (x) { return x.type === 'Image' && !before[x.id]; });
            s = fresh.length ? fresh[fresh.length - 1] : null;
          }
          // 2. Ellers: det som er markert etter innsetting
          if (!s && Bridge.supports('1.5')) {
            var sel = ctx.presentation.getSelectedShapes();
            sel.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
            var slides = ctx.presentation.getSelectedSlides();
            slides.load('items/id');
            await ctx.sync();
            s = sel.items.filter(function (x) { return x.type === 'Image' && x.id !== (replaceMeta && replaceMeta.shapeId); })[0] || null;
            if (slides.items.length) sid = slides.items[0].id;
          }
          var deleted = false;
          if (replaceMeta && replaceMeta.slideId && replaceMeta.shapeId && (!s || s.id !== replaceMeta.shapeId)) {
            var old = ctx.presentation.slides.getItem(replaceMeta.slideId).shapes.getItemOrNullObject(replaceMeta.shapeId);
            await ctx.sync();
            if (!old.isNullObject) { old.delete(); await ctx.sync(); deleted = true; }
          }
          if (!s) return { deleted: deleted, meta: null };
          if (replaceMeta && replaceMeta.name && Bridge.supports('1.4')) {
            try { s.name = replaceMeta.name; await ctx.sync(); } catch (e) { /* navn er ikke viktig */ }
          }
          if (sessionId) await writeTag(ctx, s, sessionId);
          return {
            deleted: deleted,
            meta: { slideId: sid, shapeId: s.id, name: s.name, left: s.left, top: s.top, width: s.width, height: s.height }
          };
        });
      } catch (e) {
        return { deleted: false, meta: null, error: e };
      }
    }
  };

  // Usynlig merke paa bildet, slik at panelet kjenner det igjen senere
  var TAG = 'PPTIMAGEPOINT';

  async function readTag(ctx, shape) {
    if (!Bridge.supports('1.3') || !shape.tags) return null;
    try {
      var t = shape.tags.getItemOrNullObject(TAG);
      t.load('value');
      await ctx.sync();
      return t.isNullObject ? null : t.value;
    } catch (e) { return null; }
  }

  async function writeTag(ctx, shape, value) {
    if (!Bridge.supports('1.3') || !shape.tags) return;
    try { shape.tags.add(TAG, value); await ctx.sync(); } catch (e) { /* ikke kritisk */ }
  }

  function userError(msg) { var e = new Error(msg); e.userFacing = true; return e; }

  global.Bridge = Bridge;
})(window);
