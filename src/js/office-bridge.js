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

    /**
     * Henter markert bilde. Returnerer { bytes, mime, crop, name, meta }.
     * meta beskriver plassering paa lysbildet (punkter) for aa kunne erstatte.
     */
    getSelectedImage: async function () {
      if (!Bridge.available) throw new Error('PowerPoint er ikke tilgjengelig.');
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
        if (!shapes.items.length) throw userError('Marker et bilde på lysbildet først.');
        var shape = shapes.items.filter(function (s) { return s.type === 'Image'; })[0];
        if (!shape) throw userError('Det markerte objektet er ikke et bilde. Marker et bilde (ikke en gruppe eller figur med bildefyll).');
        if (!slides.items.length) throw userError('Fant ikke lysbildet til bildet.');
        var slide = slides.items[0];
        var exported = slide.exportAsBase64();
        await ctx.sync();
        var pic = await PptxExtract.extractPicture(exported.value, { id: shape.id, name: shape.name });
        pic.meta = {
          slideId: slide.id, shapeId: shape.id, name: shape.name,
          left: shape.left, top: shape.top, width: shape.width, height: shape.height
        };
        return pic;
      });
    },

    _viaCoercion: function () {
      return new Promise(function (resolve, reject) {
        try {
          Office.context.document.getSelectedDataAsync(Office.CoercionType.Image, function (res) {
            if (res.status !== Office.AsyncResultStatus.Succeeded) {
              reject(userError('Denne PowerPoint-versjonen lar ikke tillegget lese bildet direkte. Bruk "Åpne fil" eller lim inn bildet (Ctrl+V) i panelet.'));
              return;
            }
            var bin = atob(res.value);
            var bytes = new Uint8Array(bin.length);
            for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
            resolve({ bytes: bytes, mime: 'image/png', crop: { l: 0, t: 0, r: 0, b: 0 }, name: 'bilde', meta: null });
          });
        } catch (e) {
          reject(userError('Kunne ikke lese markeringen. Bruk "Åpne fil" i stedet.'));
        }
      });
    },

    /**
     * Setter inn bildet (base64 uten prefiks). Med placement legges det paa
     * samme sted; med replaceMeta slettes originalen etterpaa.
     */
    insertImage: async function (base64, placement, replaceMeta) {
      if (!Bridge.available) throw new Error('PowerPoint er ikke tilgjengelig.');
      var opts = { coercionType: Office.CoercionType.Image };
      if (placement) {
        opts.imageLeft = placement.left;
        opts.imageTop = placement.top;
        opts.imageWidth = placement.width;
        opts.imageHeight = placement.height;
      }
      await new Promise(function (resolve, reject) {
        Office.context.document.setSelectedDataAsync(base64, opts, function (res) {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error('PowerPoint avviste bildet: ' + (res.error && res.error.message)));
        });
      });

      var newMeta = null;
      if (Bridge.supports('1.5') && global.PowerPoint) {
        try {
          newMeta = await PowerPoint.run(async function (ctx) {
            var deleted = false;
            if (replaceMeta && replaceMeta.slideId && replaceMeta.shapeId && Bridge.supports('1.3')) {
              var slide = ctx.presentation.slides.getItem(replaceMeta.slideId);
              var old = slide.shapes.getItemOrNullObject(replaceMeta.shapeId);
              await ctx.sync();
              if (!old.isNullObject) { old.delete(); deleted = true; }
            }
            var sel = ctx.presentation.getSelectedShapes();
            sel.load('items/id,items/name,items/type,items/left,items/top,items/width,items/height');
            var slides = ctx.presentation.getSelectedSlides();
            slides.load('items/id');
            await ctx.sync();
            var s = sel.items.filter(function (x) { return x.type === 'Image' && x.id !== (replaceMeta && replaceMeta.shapeId); })[0];
            if (s && replaceMeta && replaceMeta.name && Bridge.supports('1.4')) {
              try { s.name = replaceMeta.name; await ctx.sync(); } catch (e) { /* navn er ikke viktig */ }
            }
            if (!s) return { deleted: deleted, meta: null };
            return {
              deleted: deleted,
              meta: {
                slideId: slides.items.length ? slides.items[0].id : null, shapeId: s.id, name: s.name,
                left: s.left, top: s.top, width: s.width, height: s.height
              }
            };
          });
        } catch (e) {
          newMeta = { deleted: false, meta: null, error: e };
        }
      }
      return newMeta || { deleted: false, meta: null };
    }
  };

  function userError(msg) { var e = new Error(msg); e.userFacing = true; return e; }

  global.Bridge = Bridge;
})(window);
