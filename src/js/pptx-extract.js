/*
 * Henter originalbildet (full opploesning) ut av et eksportert lysbilde.
 * PowerPoint kan eksportere et lysbilde som en liten .pptx (base64). Vi
 * pakker den ut, finner riktig <p:pic> og leser bildefilen fra ppt/media.
 */
(function (global) {
  'use strict';

  var NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  var NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

  var MIME = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg',
    gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', tif: 'image/tiff',
    tiff: 'image/tiff', svg: 'image/svg+xml'
  };

  function resolvePath(baseFile, target) {
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = baseFile.split('/');
    parts.pop();
    var segs = target.split('/');
    for (var i = 0; i < segs.length; i++) {
      if (segs[i] === '..') parts.pop();
      else if (segs[i] !== '.') parts.push(segs[i]);
    }
    return parts.join('/');
  }

  function firstNS(el, ns, name) {
    var list = el.getElementsByTagNameNS(ns, name);
    return list.length ? list[0] : null;
  }

  function idNumber(id) {
    if (id == null) return NaN;
    var m = String(id).match(/\d+/);
    return m ? parseInt(m[0], 10) : NaN;
  }

  function slideFiles(zip) {
    return Object.keys(zip.files)
      .filter(function (p) { return /^ppt\/slides\/slide\d+\.xml$/.test(p); })
      .sort(function (a, b) { return idNumber(a.replace(/.*slide/, '')) - idNumber(b.replace(/.*slide/, '')); });
  }

  /**
   * @param {string|ArrayBuffer|Uint8Array} data pptx (base64-streng eller binaer)
   * @param {{id?:string, name?:string, slideIndex?:number}} target
   * @returns {Promise<{bytes:Uint8Array, mime:string, crop:{l:number,t:number,r:number,b:number}, name:string}>}
   */
  async function extractPicture(data, target) {
    target = target || {};
    var zip = await JSZip.loadAsync(data, { base64: typeof data === 'string' });
    var slides = slideFiles(zip);
    if (!slides.length) throw new Error('Fant ingen lysbilder i eksporten.');
    var slidePath = slides[Math.min(slides.length - 1, target.slideIndex || 0)];
    var xml = await zip.file(slidePath).async('string');
    var doc = new DOMParser().parseFromString(xml, 'application/xml');

    var pics = Array.prototype.slice.call(doc.getElementsByTagNameNS(NS_P, 'pic'));
    var cands = [];
    pics.forEach(function (pic) {
      var cNvPr = firstNS(pic, NS_P, 'cNvPr');
      var blip = firstNS(pic, NS_A, 'blip');
      if (!blip) return;
      var embed = blip.getAttributeNS(NS_R, 'embed') || blip.getAttribute('r:embed');
      if (!embed) return;
      cands.push({
        id: cNvPr ? cNvPr.getAttribute('id') : null,
        name: cNvPr ? cNvPr.getAttribute('name') : '',
        embed: embed,
        srcRect: firstNS(pic, NS_A, 'srcRect')
      });
    });
    if (!cands.length) throw new Error('Fant ikke noe bilde på lysbildet.');

    var chosen = null;
    var wantId = idNumber(target.id);
    if (!isNaN(wantId)) chosen = cands.filter(function (c) { return idNumber(c.id) === wantId; })[0] || null;
    if (!chosen && target.name) {
      var byName = cands.filter(function (c) { return c.name === target.name; });
      if (byName.length) chosen = byName[0];
    }
    if (!chosen) chosen = cands.length === 1 ? cands[0] : null;
    if (!chosen) throw new Error('Klarte ikke å finne det markerte bildet i lysbildet.');

    var relsPath = slidePath.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels');
    var relsFile = zip.file(relsPath);
    if (!relsFile) throw new Error('Mangler relasjonsfil for lysbildet.');
    var relsDoc = new DOMParser().parseFromString(await relsFile.async('string'), 'application/xml');
    var rels = relsDoc.getElementsByTagNameNS(NS_REL, 'Relationship');
    if (!rels.length) rels = relsDoc.getElementsByTagName('Relationship');
    var targetPath = null;
    for (var i = 0; i < rels.length; i++) {
      if (rels[i].getAttribute('Id') === chosen.embed) {
        if (rels[i].getAttribute('TargetMode') === 'External') throw new Error('Bildet er lenket, ikke innebygd.');
        targetPath = resolvePath(slidePath, rels[i].getAttribute('Target'));
        break;
      }
    }
    if (!targetPath || !zip.file(targetPath)) throw new Error('Fant ikke bildefilen i presentasjonen.');

    var ext = (targetPath.split('.').pop() || '').toLowerCase();
    var mime = MIME[ext];
    if (!mime) throw new Error('Bildeformatet .' + ext + ' støttes ikke (prøv PNG eller JPG).');
    var bytes = await zip.file(targetPath).async('uint8array');

    var crop = { l: 0, t: 0, r: 0, b: 0 };
    if (chosen.srcRect) {
      ['l', 't', 'r', 'b'].forEach(function (k) {
        var v = parseInt(chosen.srcRect.getAttribute(k) || '0', 10);
        crop[k] = isNaN(v) ? 0 : v / 100000;
      });
    }
    return { bytes: bytes, mime: mime, crop: crop, name: chosen.name || 'bilde' };
  }

  global.PptxExtract = { extractPicture: extractPicture, resolvePath: resolvePath };
})(window);
