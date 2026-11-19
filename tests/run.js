// Ende-til-ende-tester av oppgaveruten i Chromium, med etterlignet PowerPoint-API.
// Kjoer: node tests/run.js   (krever playwright og python3 tests/make_fixtures.py foerst)
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = path.resolve(__dirname, '..');
const SHOTS = process.env.SHOTS || path.join(__dirname, 'screenshots');
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.jpg': 'image/jpeg', '.pptx': 'application/octet-stream' };

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const p = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]));
      if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end(); }
      res.writeHead(200, { 'Content-Type': TYPES[path.extname(p)] || 'application/octet-stream' });
      fs.createReadStream(p).pipe(res);
    });
    srv.listen(0, () => resolve(srv));
  });
}

let failures = 0;
function check(cond, msg) { console.log((cond ? '  ok   ' : '  FEIL ') + msg); if (!cond) failures++; }

const officeMock = (pptxB64) => {
  if (window !== window.top) {
    // Inne i dialogvinduet: meldinger til og fra panelet
    window.Office = {
      EventType: { DialogParentMessageReceived: 'pm' },
      onReady(cb) { setTimeout(() => cb({ host: 'PowerPoint' }), 0); },
      context: {
        requirements: { isSetSupported: () => true },
        ui: {
          messageParent(m) { window.top.__fromDialog(m); },
          addHandlerAsync(t, h, cb) { window.__onParentMsg = h; cb && cb({ status: 'succeeded' }); },
        },
      },
    };
    return;
  }
  // Med __persistMock overlever "presentasjonen" at panelet lastes paa nytt
  const saved = window.__persistMock && sessionStorage.getItem('mockState');
  const st = saved ? JSON.parse(saved) : null;
  window.__log = st ? st.log : [];
  const shapes = st ? st.shapes : [
    { id: '2', name: 'Picture 1', type: 'Image', left: 36, top: 36, width: 144, height: 144 },
    { id: '3', name: 'Testbilde', type: 'Image', left: 216, top: 72, width: 360, height: 225 },
  ];
  window.__shapes = shapes;
  let selected = st ? st.selected : (window.__initialSelection || '3');
  const persist = () => {
    if (window.__persistMock) sessionStorage.setItem('mockState', JSON.stringify({ shapes, selected, log: window.__log.map((l) => ({ op: l.op, id: l.id, opts: l.opts })) }));
  };
  const withTags = (s) => {
    if (!s.tagMap) s.tagMap = {};
    Object.defineProperty(s, 'tags', { configurable: true, enumerable: false, value: {
      getItemOrNullObject: (k) => ({ isNullObject: window.__noTags || !(k in s.tagMap), value: s.tagMap[k], load() {} }),
      add: (k, v) => { if (!window.__noTags) s.tagMap[k] = v; persist(); },
    } });
    return s;
  };
  // Etterligner Ctrl+Z / Ctrl+Y i PowerPoint: bytt synlig figur og marker den
  window.__pptSwap = (removeId, restoreId) => {
    const i = shapes.findIndex((x) => x.id === removeId);
    if (i >= 0) (window.__trash = window.__trash || []).push(shapes.splice(i, 1)[0]);
    const t = window.__trash.findIndex((x) => x.id === restoreId);
    shapes.push(window.__trash.splice(t, 1)[0]);
    window.__select(restoreId);
  };
  window.__select = (id) => { selected = id; persist(); window.__selHandler && window.__selHandler({}); };
  const coll = (items) => ({ items, load() {} });
  // Eksport av lysbildet: bygger en pptx der bildet er det som faktisk ligger paa lysbildet
  const tasks = [];
  const exportSlide = () => {
    const out = { value: null };
    tasks.push(async () => {
      const sel = shapes.find((x) => x.id === selected);
      if (!sel || !sel.data) { out.value = pptxB64; return; }
      const zip = await JSZip.loadAsync(pptxB64, { base64: true });
      const slide = Object.keys(zip.files).find((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
      const xml = await zip.file(slide).async('string');
      // finn bildefilen til figuren med samme navn, og bytt den ut
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const pic = [...doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'pic')]
        .find((p) => p.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr')[0].getAttribute('name') === sel.name);
      const blip = pic.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')[0];
      const rid = blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
      const rels = await zip.file(slide.replace('slides/', 'slides/_rels/') + '.rels').async('string');
      const target = new RegExp('Id="' + rid + '"[^>]*Target="\\.\\./media/([^"]+)"').exec(rels) || new RegExp('Target="\\.\\./media/([^"]+)"[^>]*Id="' + rid + '"').exec(rels);
      zip.file('ppt/media/' + target[1], sel.data, { base64: true });
      const sr = pic.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'srcRect')[0];
      if (sr) sr.parentNode.removeChild(sr);
      zip.file(slide, new XMLSerializer().serializeToString(doc));
      out.value = await zip.generateAsync({ type: 'base64' });
    });
    return out;
  };
  const ctx = {
    sync: async () => { while (tasks.length) await tasks.shift()(); },
    presentation: {
      getSelectedShapes: () => coll(shapes.filter((s) => s.id === selected).map(withTags)),
      getSelectedSlides: () => coll([{ id: '256#', exportAsBase64: exportSlide }]),
      slides: { getItem: () => ({
        exportAsBase64: exportSlide,
        shapes: { get items() { return shapes.map(withTags); }, load() {}, getItemOrNullObject: (id) => {
          const s = shapes.find((x) => x.id === id);
          return Object.assign({}, s || {}, { isNullObject: !s, load() {}, delete() { (window.__trash = window.__trash || []).push(s); shapes.splice(shapes.indexOf(s), 1); window.__log.push({ op: 'delete', id }); persist(); } });
        } },
      }) },
    },
  };
  window.PowerPoint = { run: (fn) => fn(ctx) };
  window.Office = {
    HostType: { PowerPoint: 'PowerPoint' }, CoercionType: { Image: 'image' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    EventType: { DocumentSelectionChanged: 'sel', DialogMessageReceived: 'dm', DialogEventReceived: 'de' },
    onReady(cb) { setTimeout(() => cb({ host: 'PowerPoint' }), 10); },
    context: {
      requirements: { isSetSupported: () => true },
      ui: {
        displayDialogAsync(url, opts, cb) {
          const f = document.createElement('iframe');
          f.id = 'dlg';
          f.style.cssText = 'position:fixed;left:0;top:0;width:100%;height:100%;border:0;background:#fff;z-index:99';
          const handlers = {};
          window.__fromDialog = (m) => handlers.dm && handlers.dm({ message: m });
          const dialog = {
            messageChild(m) { f.contentWindow.__onParentMsg && f.contentWindow.__onParentMsg({ message: m }); },
            addEventHandler(t, h) { handlers[t] = h; },
            close() { f.remove(); },
          };
          f.src = url;
          document.body.appendChild(f);
          window.__dialogOpts = opts;
          cb({ status: 'succeeded', value: dialog });
        },
      },
      document: {
        addHandlerAsync(t, cb) { window.__selHandler = cb; },
        getSelectedDataAsync(t, cb) { cb({ status: 'failed' }); },
        setSelectedDataAsync(data, opts, cb) {
          window.__log.push({ op: 'insert', data, opts });
          const id = String(100 + window.__log.length);
          shapes.push({ id, name: 'Picture ' + id, type: 'Image', left: opts.imageLeft, top: opts.imageTop, width: opts.imageWidth, height: opts.imageHeight, data });
          if (!window.__noSelectAfterInsert) selected = id;
          persist();
          cb({ status: 'succeeded' });
        },
      },
    },
  };
};

async function decodeB64Png(page, b64) {
  return page.evaluate(async (b64) => {
    const img = new Image();
    img.src = 'data:image/png;base64,' + b64;
    await img.decode();
    const c = document.createElement('canvas'); c.width = img.width; c.height = img.height;
    const x = c.getContext('2d'); x.drawImage(img, 0, 0);
    const d = x.getImageData(0, 0, c.width, c.height);
    const a = (fx, fy) => d.data[(Math.floor(fy * (c.height - 1)) * c.width + Math.floor(fx * (c.width - 1))) * 4 + 3];
    return { w: c.width, h: c.height, a };
  }, b64);
}

(async () => {
  fs.mkdirSync(SHOTS, { recursive: true });
  const srv = await serve();
  const base = `http://localhost:${srv.address().port}`;
  const browser = await chromium.launch();
  const pptxB64 = fs.readFileSync(path.join(__dirname, 'fixtures/test.pptx')).toString('base64');

  /* ---- 1. Frittstaaende: last fil, fjern groenn bakgrunn ---- */
  console.log('Test 1: frittstående modus og fargenøkkel');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.goto(base + '/src/taskpane.html');
    await page.setInputFiles('#fileInput', path.join(__dirname, 'fixtures/green.jpg'));
    await page.waitForSelector('#stage:not([hidden])');
    await page.click('#btnAuto');
    await page.waitForTimeout(700);
    // Angre og gjoer om med tastaturet
    await page.evaluate(() => { const i = document.getElementById('sl-saturation'); i.value = 50; i.dispatchEvent(new Event('input')); });
    await page.waitForTimeout(700);
    await page.keyboard.press('Control+z');
    const u1 = await page.evaluate(() => [window.__pip.S.adjust.saturation, window.__pip.S.key.enabled, document.getElementById('sl-saturation').value]);
    check(u1[0] === 0 && u1[1] === true && u1[2] === '0', 'Ctrl+Z angrer metning, glidebryteren følger med ' + u1);
    await page.keyboard.press('Control+z');
    check(await page.evaluate(() => window.__pip.S.key.enabled === false), 'Ctrl+Z en gang til angrer bakgrunnsfjerningen');
    await page.keyboard.press('Control+y');
    await page.keyboard.press('Control+y');
    const u2 = await page.evaluate(() => [window.__pip.S.adjust.saturation, window.__pip.S.key.enabled]);
    check(u2[0] === 50 && u2[1] === true, 'Ctrl+Y gjør om begge stegene ' + u2);
    // endring rett foer angring (foer den er lagret) skal ogsaa angres
    await page.evaluate(() => { const i = document.getElementById('sl-contrast'); i.value = 30; i.dispatchEvent(new Event('input')); });
    await page.keyboard.press('Control+z');
    check(await page.evaluate(() => window.__pip.S.adjust.contrast === 0 && window.__pip.S.adjust.saturation === 50), 'Ctrl+Z rett etter en endring angrer akkurat den');
    await page.click('#btnUndo');
    check(await page.evaluate(() => window.__pip.S.adjust.saturation === 0), 'Undo-knappen virker');
    await page.waitForTimeout(300);
    const r = await page.evaluate(() => {
      const res = window.__pip.renderFull();
      const d = res.image.data, w = res.image.width, h = res.image.height;
      const a = (x, y) => d[(Math.floor(y * (h - 1)) * w + Math.floor(x * (w - 1))) * 4 + 3];
      return { w, h, corner: a(0.02, 0.02), center: a(0.5, 0.3), hole: a(0.5, 0.5), blue: a(0.8, 0.5), right: a(0.97, 0.9) };
    });
    check(r.w === 1600 && r.h === 1000, 'full oppløsning beholdes (' + r.w + 'x' + r.h + ')');
    check(r.corner === 0 && r.right === 0, 'bakgrunnen er gjennomsiktig (hjørner: ' + r.corner + ', ' + r.right + ')');
    check(r.center === 255 && r.blue === 255, 'motivet er helt synlig (' + r.center + ', ' + r.blue + ')');
    check(r.hole === 255, 'lukket hull beholdes i sammenhengende modus (' + r.hole + ')');
    await page.selectOption('#keyMode', 'global');
    const hole = await page.evaluate(() => { const r = window.__pip.renderFull(); const d = r.image.data; return d[(500 * 1600 + 800) * 4 + 3]; });
    check(hole === 0, 'hullet fjernes i modus «fargen overalt» (' + hole + ')');
    await page.selectOption('#keyMode', 'connected');
    // Justeringer og skalering
    await page.evaluate(() => { const S = window.__pip.S; Object.assign(S.adjust, { saturation: 40, clarity: 50, hue: 30, exposure: 0.3, sharpness: 40, temperature: 20 }); S.out = { w: 400, h: 250, lock: true }; S.trim = true; window.__pip.syncUI(); window.__pip.renderNow(); });
    const t = await page.evaluate(() => { const r = window.__pip.renderFull(); return { w: r.image.width, h: r.image.height, rect: r.rect }; });
    check(t.rect.x > 400 && t.rect.w < 1000, 'beskjæring av tomme kanter (' + JSON.stringify(t.rect) + ')');
    check(Math.abs(t.w - t.rect.w / 4) <= 1, 'skalering til 25 % etter beskjæring (' + t.w + 'x' + t.h + ')');
    await page.waitForTimeout(200);
    await page.screenshot({ path: path.join(SHOTS, 'standalone.png') });
    const t0 = Date.now();
    await page.evaluate(() => { window.__pip.S.out = { w: 1600, h: 1000, lock: true }; window.__pip.renderFull(); });
    console.log('  info  full behandling 1600x1000 med alle justeringer: ' + (Date.now() - t0) + ' ms');
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 2. Med etterlignet PowerPoint ---- */
  console.log('Test 2: hent fra PowerPoint, erstatt bildet');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForFunction(() => window.Bridge && window.Bridge.available);
    await page.waitForSelector('#stage:not([hidden])');
    check(true, 'markert bilde lastes automatisk når panelet åpnes');
    await page.uncheck('#live');
    const src = await page.evaluate(() => ({ w: window.__pip.S.src.full.width, h: window.__pip.S.src.full.height, meta: window.__pip.S.src.meta }));
    check(src.w === 800 && src.h === 1000, 'riktig bilde hentet, med PowerPoint-beskjæring brukt (' + src.w + 'x' + src.h + ')');
    check(src.meta && src.meta.shapeId === '3', 'plassering lest fra figuren');
    // klikk paa bakgrunnen i forhaandsvisningen
    const box = await page.locator('#view').boundingBox();
    await page.mouse.click(box.x + 6, box.y + 6);
    await page.waitForTimeout(200);
    const enabled = await page.evaluate(() => window.__pip.S.key.enabled && window.__pip.S.key.colors.length === 1);
    check(enabled, 'klikk med pipetten velger bakgrunnsfarge');
    // et viskelaer-stroek
    await page.click('[data-tool="erase"]');
    await page.mouse.move(box.x + box.width * 0.85, box.y + box.height * 0.3);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width * 0.85, box.y + box.height * 0.7, { steps: 8 });
    await page.mouse.up();
    check(await page.evaluate(() => window.__pip.S.strokes.length === 1), 'penselstrøk registrert');
    await page.screenshot({ path: path.join(SHOTS, 'powerpoint-mock.png') });
    await page.click('#btnInsert');
    await page.waitForFunction(() => window.__log.some((l) => l.op === 'delete'), null, { timeout: 15000 });
    const log = await page.evaluate(() => window.__log.map((l) => ({ op: l.op, id: l.id, opts: l.opts, data: l.data })));
    const ins = log.find((l) => l.op === 'insert');
    check(ins && ins.opts.imageLeft === 216 && ins.opts.imageTop === 72 && Math.abs(ins.opts.imageWidth - 360) < 0.01 && Math.abs(ins.opts.imageHeight - 225) < 0.01,
      'satt inn på samme plass og i samme størrelse ' + JSON.stringify(ins && ins.opts));
    check(log.some((l) => l.op === 'delete' && l.id === '3'), 'originalen ble slettet');
    const png = await decodeB64Png(page, ins.data);
    check(png.w === 800 && png.h === 1000, 'innsatt PNG har full oppløsning (' + png.w + 'x' + png.h + ')');
    const st = await page.locator('#status').textContent();
    check(/inserted/.test(st), 'statusmelding: ' + st);
    // ny runde med trim: skal erstatte det nye bildet
    await page.check('#trim');
    await page.click('#btnInsert');
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 2, null, { timeout: 15000 });
    const log2 = await page.evaluate(() => window.__log.filter((l) => l.op !== 'insert' || true).map((l) => ({ op: l.op, id: l.id, opts: l.opts })));
    const ins2 = log2.filter((l) => l.op === 'insert')[1];
    check(ins2.opts.imageLeft > 216 && ins2.opts.imageWidth < 360, 'andre runde med beskjæring flytter og krymper riktig ' + JSON.stringify(ins2.opts));
    check(log2.filter((l) => l.op === 'delete')[1].id === '101', 'andre runde erstatter det forrige innsatte bildet');
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 3. Hvit bakgrunn, uten markering ---- */
  console.log('Test 3: hvit bakgrunn og mørk modus');
  {
    const page = await browser.newPage({ viewport: { width: 340, height: 1300 }, colorScheme: 'dark' });
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.goto(base + '/src/taskpane.html');
    await page.setInputFiles('#fileInput', path.join(__dirname, 'fixtures/white.png'));
    await page.waitForSelector('#stage:not([hidden])');
    const box = await page.locator('#view').boundingBox();
    await page.mouse.click(box.x + 4, box.y + box.height - 4);
    await page.selectOption('#bgSelect', 'black');
    await page.waitForTimeout(200);
    const r = await page.evaluate(() => { const res = window.__pip.renderFull(); const d = res.image.data; const w = res.image.width; const a = (x, y) => d[(y * w + x) * 4 + 3]; return [a(10, 10), a(450, 450), a(450, 160)]; });
    check(r[0] === 0 && r[1] === 255 && r[2] === 255, 'hvit bakgrunn fjernet, motiv beholdt ' + r);
    await page.screenshot({ path: path.join(SHOTS, 'dark.png'), fullPage: true });
    await page.close();
  }

  /* ---- 4. Bytte markering og live-oppdatering ---- */
  console.log('Test 4: følger markeringen og oppdaterer lysbildet live');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(() => { window.__initialSelection = 'none'; });
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForFunction(() => window.Bridge && window.Bridge.available);
    await page.waitForTimeout(500);
    check(await page.isVisible('#empty'), 'ingenting lastes når ingen bilde er markert');
    await page.evaluate(() => window.__select('3'));
    await page.waitForSelector('#stage:not([hidden])');
    check(await page.evaluate(() => window.__pip.S.src.meta.shapeId === '3'), 'bilde lastes når det markeres');
    await page.waitForTimeout(1200);
    check(await page.evaluate(() => window.__log.length === 0), 'ingen endring på lysbildet før du justerer noe');
    // juster metning: live skal oppdatere lysbildet etter en kort pause
    await page.evaluate(() => { const i = document.getElementById('sl-saturation'); i.value = 50; i.dispatchEvent(new Event('input')); });
    await page.waitForFunction(() => window.__log.some((l) => l.op === 'delete'), null, { timeout: 10000 });
    const log = await page.evaluate(() => window.__log.map((l) => ({ op: l.op, id: l.id })));
    check(log.length === 2 && log[0].op === 'insert' && log[1].id === '3', 'live: bildet erstattet automatisk ' + JSON.stringify(log));
    check(/updated/i.test(await page.locator('#liveState').textContent()), 'live-status vises');
    await page.waitForTimeout(2500);
    check(await page.evaluate(() => window.__log.length === 2), 'ingen ekstra oppdateringer eller ny innlasting etter live-oppdateringen');
    check(await page.evaluate(() => window.__pip.S.adjust.saturation === 50), 'justeringene beholdes etter live-oppdateringen');
    // ny justering erstatter den nye versjonen
    await page.evaluate(() => { const i = document.getElementById('sl-hue'); i.value = 40; i.dispatchEvent(new Event('input')); });
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 2, null, { timeout: 10000 });
    check(await page.evaluate(() => window.__log.filter((l) => l.op === 'delete')[1].id === '101'), 'andre live-oppdatering erstatter forrige versjon');
    // Ctrl+Z i panelet angrer fargetonen, og live oppdaterer lysbildet igjen
    await page.click('#info');
    await page.keyboard.press('Control+z');
    check(await page.evaluate(() => window.__pip.S.adjust.hue === 0 && window.__pip.S.adjust.saturation === 50), 'Ctrl+Z angrer siste justering i live-modus');
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 3, null, { timeout: 10000 });
    check(true, 'lysbildet oppdateres etter angring');
    // marker et annet bilde: det lastes inn
    await page.evaluate(() => window.__select('2'));
    await page.waitForFunction(() => window.__pip.S.src.meta.shapeId === '2', null, { timeout: 10000 });
    const dims = await page.evaluate(() => [window.__pip.S.src.full.width, window.__pip.S.src.full.height]);
    check(dims[0] === 900 && dims[1] === 900, 'bytter til nytt markert bilde (' + dims + ')');
    check(await page.evaluate(() => window.__pip.S.adjust.hue === 0 && window.__pip.S.adjust.saturation === 0), 'nytt bilde starter med nullstilte justeringer');
    await page.screenshot({ path: path.join(SHOTS, 'live.png') });
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 5. Beskjaering og flytting i PowerPoint, skjult forhaandsvisning ---- */
  console.log('Test 5: beskjæring i PowerPoint følges, forhåndsvisning kan skjules');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForSelector('#stage:not([hidden])');
    // Brukeren beskjaerer 25 % fra venstre med PowerPoints verktoey: figuren blir smalere og flytter seg
    await page.evaluate(() => {
      const sh = window.__shapes.find((x) => x.id === '3');
      sh.left += sh.width * 0.25; sh.width *= 0.75;
      window.Bridge.getShapeCrop = async () => ({ l: 0.25, t: 0, r: 0, b: 0 });
      const i = document.getElementById('sl-saturation'); i.value = 30; i.dispatchEvent(new Event('input'));
    });
    await page.waitForFunction(() => window.__log.some((l) => l.op === 'delete'), null, { timeout: 10000 });
    const ins = await page.evaluate(() => window.__log.find((l) => l.op === 'insert'));
    check(Math.abs(ins.opts.imageLeft - 306) < 0.01 && Math.abs(ins.opts.imageWidth - 270) < 0.01 && Math.abs(ins.opts.imageHeight - 225) < 0.01,
      'beskjæringen beholdes ved live-oppdatering ' + JSON.stringify(ins.opts));
    const png = await decodeB64Png(page, ins.data);
    check(png.w === 600 && png.h === 1000, 'bare det beskårne området legges inn (' + png.w + 'x' + png.h + ')');
    await page.waitForTimeout(400);
    await page.screenshot({ path: path.join(SHOTS, 'crop.png') });
    // Ny beskjaering av det nye bildet (10 % nederst) og flytting
    await page.evaluate(() => {
      const sh = window.__shapes.find((x) => x.id === '101');
      sh.height *= 0.9; sh.top += 50;
      window.Bridge.getShapeCrop = async () => ({ l: 0, t: 0, r: 0, b: 0.1 });
      const i = document.getElementById('sl-saturation'); i.value = 60; i.dispatchEvent(new Event('input'));
    });
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 2, null, { timeout: 10000 });
    const ins2 = await page.evaluate(() => window.__log.filter((l) => l.op === 'insert')[1]);
    const png2 = await decodeB64Png(page, ins2.data);
    check(png2.w === 600 && png2.h === 900 && Math.abs(ins2.opts.imageTop - 122) < 0.01 && Math.abs(ins2.opts.imageHeight - 202.5) < 0.01,
      'ny beskjæring og flytting legges oppå den forrige ' + png2.w + 'x' + png2.h + ' ' + JSON.stringify(ins2.opts));
    // Skjul forhaandsvisningen
    await page.click('#btnTogglePreview');
    check(await page.isHidden('#stage') && await page.isVisible('#previewHint') && await page.isVisible('#btnUndo'), 'forhåndsvisningen kan skjules');
    await page.screenshot({ path: path.join(SHOTS, 'hidden.png') });
    await page.reload();
    await page.waitForSelector('#infoRow:not([hidden])');
    check(await page.isHidden('#stage'), 'valget huskes til neste gang');
    await page.click('#btnTogglePreview');
    check(await page.isVisible('#stage'), 'forhåndsvisningen kan vises igjen');
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 6. Reset-knappen setter originalen tilbake ---- */
  console.log('Test 6: Reset legger tilbake originalen');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(() => { window.__initialSelection = '2'; });
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForSelector('#stage:not([hidden])');
    check(await page.isDisabled('#btnResetAll') === false, 'Reset-knappen er aktiv når et bilde er lastet');
    // fjern bakgrunnen og juster farger, vent paa live
    const box = await page.locator('#view').boundingBox();
    await page.mouse.click(box.x + 4, box.y + 4);
    await page.evaluate(() => { const i = document.getElementById('sl-hue'); i.value = 90; i.dispatchEvent(new Event('input')); });
    await page.waitForFunction(() => window.__log.some((l) => l.op === 'delete'), null, { timeout: 10000 });
    // flytt bildet litt i PowerPoint, saa Reset ma legge det tilbake paa opprinnelig plass
    await page.evaluate(() => { const sh = window.__shapes.find((x) => x.id === '101'); sh.left += 40; });
    await page.click('#btnResetAll');
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 2, null, { timeout: 10000 });
    const last = await page.evaluate(() => window.__log.filter((l) => l.op === 'insert')[1]);
    const orig = fs.readFileSync(path.join(__dirname, 'fixtures/white.png')).toString('base64');
    check(last.data === orig, 'den opprinnelige filen legges tilbake uendret');
    check(last.opts.imageLeft === 36 && last.opts.imageTop === 36 && last.opts.imageWidth === 144, 'på opprinnelig plass ' + JSON.stringify(last.opts));
    const st = await page.evaluate(() => [window.__pip.S.adjust.hue, window.__pip.S.key.enabled, document.getElementById('sl-hue').value]);
    check(st[0] === 0 && st[1] === false && st[2] === '0', 'alle innstillinger nullstilt ' + st);
    await page.waitForTimeout(1500);
    check(await page.evaluate(() => window.__log.length === 4), 'live legger ikke inn noe ekstra etter Reset');
    // verktoeytips
    await page.hover('#btnResetAll');
    await page.waitForTimeout(600);
    check(await page.isVisible('#tip') && /original/.test(await page.locator('#tip').textContent()), 'verktøytips vises ved hover');
    await page.screenshot({ path: path.join(SHOTS, 'icons.png') });
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 7. Panelet lukkes og aapnes igjen (f.eks. av Utforming) ---- */
  console.log('Test 7: redigeringene huskes når panelet lastes på nytt');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(() => { window.__persistMock = true; });
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForSelector('#stage:not([hidden])');
    await page.evaluate(() => { const i = document.getElementById('sl-hue'); i.value = 40; i.dispatchEvent(new Event('input')); });
    await page.waitForFunction(() => window.__log.some((l) => l.op === 'delete'), null, { timeout: 10000 });
    await page.waitForTimeout(800);
    check(await page.evaluate(() => window.__shapes.find((x) => x.id === '101').tagMap.PPTIMAGEPOINT === window.__pip.S.src.sessionId + '|1'), 'det nye bildet er merket med økten');
    await page.evaluate(() => { const i = document.getElementById('sl-saturation'); i.value = -50; i.dispatchEvent(new Event('input')); });
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 2, null, { timeout: 10000 });
    await page.waitForTimeout(500);
    // Panelet lukkes og aapnes (samme presentasjon, nytt bilde-id 102 er markert)
    await page.reload();
    await page.waitForSelector('#stage:not([hidden])');
    await page.waitForTimeout(300);
    const r = await page.evaluate(() => ({ hue: window.__pip.S.adjust.hue, sat: window.__pip.S.adjust.saturation, w: window.__pip.S.src.full.width, st: document.getElementById('status').textContent, hueSlider: document.getElementById('sl-hue').value }));
    check(r.hue === 40 && r.sat === -50 && r.hueSlider === '40', 'justeringene er hentet tilbake ' + JSON.stringify(r));
    check(/restored/.test(r.st), 'melding om at redigeringene er hentet tilbake');
    await page.waitForTimeout(1500);
    check(await page.evaluate(() => window.__log.length === 4), 'ingen ny oppdatering av lysbildet bare av å åpne panelet');
    await page.click('#info');
    await page.keyboard.press('Control+z');
    check(await page.evaluate(() => window.__pip.S.adjust.saturation === 0 && window.__pip.S.adjust.hue === 40), 'angrehistorikken er tatt vare på');
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 3, null, { timeout: 10000 });
    const dl = await page.evaluate(() => JSON.stringify(window.__log.map((l) => l.op + ':' + (l.id || ''))));
    check(await page.evaluate(() => window.__log.filter((l) => l.op === 'delete')[2].id === '103'), 'lysbildet oppdateres videre fra det redigerte bildet ' + dl);
    await page.click('#btnResetAll');
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 4, null, { timeout: 10000 });
    const last = await page.evaluate(() => window.__log.filter((l) => l.op === 'insert').pop());
    check(last.opts.imageLeft === 216 && last.opts.imageTop === 72 && last.opts.imageWidth === 360 && last.opts.imageHeight === 225, 'Reset legger originalen tilbake på opprinnelig plass ' + JSON.stringify(last.opts));
    check(await page.evaluate(() => window.__pip.S.adjust.hue === 0), 'Reset nullstiller etter gjenoppretting');
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 8. Ctrl+Z og Ctrl+Y i PowerPoint ---- */
  console.log('Test 8: panelet følger angring og gjør om i PowerPoint');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForSelector('#stage:not([hidden])');
    const setSlider = (k, v) => page.evaluate(([k, v]) => { const i = document.getElementById('sl-' + k); i.value = v; i.dispatchEvent(new Event('input')); i.dispatchEvent(new Event('change', { bubbles: true })); }, [k, v]);
    const adj = () => page.evaluate(() => [window.__pip.S.adjust.hue, window.__pip.S.adjust.saturation, document.getElementById('sl-saturation').value]);
    await setSlider('hue', 40);
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 1, null, { timeout: 10000 });
    await setSlider('saturation', -50);
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 2, null, { timeout: 10000 });
    await page.waitForTimeout(400);
    // Ctrl+Z i PowerPoint: versjon 1 (figur 101) kommer tilbake
    await page.evaluate(() => window.__pptSwap('102', '101'));
    await page.waitForTimeout(600);
    let a = await adj();
    check(a[0] === 40 && a[1] === 0 && a[2] === '0', 'Ctrl+Z i PowerPoint: panelet viser forrige versjon ' + a);
    // Ctrl+Z igjen: originalen
    await page.evaluate(() => window.__pptSwap('101', '3'));
    await page.waitForTimeout(600);
    a = await adj();
    check(a[0] === 0 && a[1] === 0, 'Ctrl+Z helt tilbake: panelet viser originalen ' + a);
    // Ctrl+Y i PowerPoint: versjon 1 igjen
    await page.evaluate(() => window.__pptSwap('3', '101'));
    await page.waitForTimeout(600);
    a = await adj();
    check(a[0] === 40 && a[1] === 0, 'Ctrl+Y i PowerPoint: panelet følger med ' + a);
    check(await page.evaluate(() => window.__log.filter((l) => l.op === 'insert').length === 2), 'ingen ekstra innsetting når panelet bare følger lysbildet');
    // Ny justering bygger videre paa versjonen som ligger paa lysbildet
    await setSlider('vibrance', 30);
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 3, null, { timeout: 10000 });
    check(await page.evaluate(() => window.__log.filter((l) => l.op === 'delete')[2].id === '101'), 'ny justering erstatter bildet som faktisk ligger på lysbildet');
    a = await page.evaluate(() => [window.__pip.S.adjust.hue, window.__pip.S.adjust.vibrance]);
    check(a[0] === 40 && a[1] === 30, 'og beholder innstillingene fra den versjonen ' + a);
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 9. Uten merker og uten markering etter innsetting (verste tilfelle) ---- */
  console.log('Test 9: kjenner igjen bildet på innholdet, og Reset virker');
  {
    const page = await browser.newPage({ viewport: { width: 360, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(() => { window.__persistMock = true; window.__noTags = true; window.__noSelectAfterInsert = true; window.__initialSelection = '2'; });
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForSelector('#stage:not([hidden])');
    const box = await page.locator('#view').boundingBox();
    await page.mouse.click(box.x + 4, box.y + 4);
    await page.waitForFunction(() => window.__log.some((l) => l.op === 'delete'), null, { timeout: 10000 });
    await page.waitForTimeout(600);
    check(await page.evaluate(() => window.__pip.S.src.meta && window.__pip.S.src.meta.shapeId === '101'), 'finner det nye bildet uten å lese markeringen');
    // Brukeren klikker paa det nye bildet, og panelet lastes paa nytt (f.eks. av Utforming)
    await page.evaluate(() => window.__select('101'));
    await page.reload();
    await page.waitForSelector('#stage:not([hidden])');
    await page.waitForTimeout(400);
    const r = await page.evaluate(() => ({ key: window.__pip.S.key.enabled, st: document.getElementById('status').textContent }));
    check(r.key === true && /restored/.test(r.st), 'kjenner igjen sitt eget bilde på innholdet ' + JSON.stringify(r));
    await page.click('#btnResetAll');
    await page.waitForFunction(() => window.__log.filter((l) => l.op === 'delete').length === 2, null, { timeout: 10000 });
    const last = await page.evaluate(() => window.__log.filter((l) => l.op === 'insert').pop());
    const orig = fs.readFileSync(path.join(__dirname, 'fixtures/white.png')).toString('base64');
    const sh = await page.evaluate(() => window.__shapes.map((x) => x.id));
    check(last.opts.imageLeft === 36 && last.opts.imageWidth === 144, 'Reset legger originalen på opprinnelig plass');
    check(await page.evaluate((o) => window.__shapes.find((x) => x.data === o) !== undefined, orig) && sh.length === 2, 'originalfilen ligger på lysbildet, og den redigerte er borte ' + sh);
    check(/back to the original/.test(await page.locator('#status').textContent()), 'melding om at originalen er tilbake');
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  /* ---- 10. Stort forhaandsvisningsvindu og myk pensel ---- */
  console.log('Test 10: stort forhåndsvisningsvindu og myk pensel');
  {
    const page = await browser.newPage({ viewport: { width: 1100, height: 760 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/appsforoffice.microsoft.com/**', (r) => r.fulfill({ body: '', contentType: 'text/javascript' }));
    await page.addInitScript(officeMock, pptxB64);
    await page.goto(base + '/src/taskpane.html');
    await page.waitForSelector('#stage:not([hidden])');
    await page.uncheck('#live');
    check(await page.isVisible('#btnPopout'), 'knappen for stort vindu vises');
    await page.click('#btnPopout');
    const frame = page.frameLocator('#dlg');
    await frame.locator('#pvView:not([hidden])').waitFor({ timeout: 10000 });
    const size = await frame.locator('#pvView').evaluate((c) => [c.width, c.height]);
    check(size[0] === 800 && size[1] === 1000, 'vinduet får bildet i høyere oppløsning (' + size + ')');
    // klikk paa bakgrunnen i vinduet
    const vb = await frame.locator('#pvView').boundingBox();
    await page.mouse.click(vb.x + 6, vb.y + 6);
    await page.waitForFunction(() => window.__pip.S.key.enabled);
    check(true, 'pipetten i vinduet fjerner bakgrunnen i panelet');
    // myk pensel i vinduet
    await frame.locator('[data-tool="erase"]').click();
    check(await page.evaluate(() => window.__pip.S.tool === 'erase') && await page.isVisible('#brushBar'), 'verktøyvalg følger med til panelet, pensellinjen vises');
    await frame.locator('#pvHard').evaluate((i) => { i.value = 0; i.dispatchEvent(new Event('input')); i.dispatchEvent(new Event('change')); });
    await frame.locator('#pvSize').evaluate((i) => { i.value = 80; i.dispatchEvent(new Event('input')); i.dispatchEvent(new Event('change')); });
    await page.waitForFunction(() => window.__pip.S.hardness === 0 && window.__pip.S.brush === 80);
    check(await page.evaluate(() => document.getElementById('sl-hardness').value === '0'), 'hardhet og størrelse synkroniseres');
    await page.mouse.move(vb.x + vb.width * 0.5, vb.y + vb.height * 0.3);
    await page.mouse.down();
    await page.mouse.move(vb.x + vb.width * 0.5, vb.y + vb.height * 0.7, { steps: 10 });
    await page.mouse.up();
    await page.waitForFunction(() => window.__pip.S.strokes.length === 1);
    const st = await page.evaluate(() => window.__pip.S.strokes[0]);
    check(st.h === 0 && st.mode === 'erase' && Math.abs(st.r - 0.08) < 1e-9, 'penselstrøk fra vinduet med myk kant ' + JSON.stringify({ h: st.h, r: st.r }));
    // myk kant: alfa avtar gradvis fra midten av stroeket
    const prof = await page.evaluate(() => {
      const r = window.__pip.renderFull(); const d = r.image.data, w = r.image.width, y = Math.round(r.image.height * 0.5);
      const cx = Math.round(w * 0.5), R = 0.08 * Math.max(w, r.image.height);
      return [0, 0.25, 0.5, 0.75, 0.95].map((f) => d[(y * w + Math.round(cx + f * R)) * 4 + 3]);
    });
    check(prof[0] === 0 && prof[1] > 0 && prof[1] < prof[2] && prof[2] < prof[3] && prof[3] < prof[4], 'myk pensel gir gradvis overgang ' + prof);
    await page.waitForTimeout(400);
    await page.screenshot({ path: path.join(SHOTS, 'popout.png') });
    // angre fra vinduet
    await frame.locator('#pvView').click({ position: { x: 5, y: 5 }, trial: true });
    await frame.locator('#pvUndo').click();
    await page.waitForFunction(() => window.__pip.S.strokes.length === 0);
    check(true, 'angring fra vinduet virker');
    await frame.locator('#pvClose').click();
    await page.waitForTimeout(200);
    check(await page.locator('#dlg').count() === 0, 'vinduet kan lukkes');
    check(errors.length === 0, 'ingen JS-feil ' + errors.join(' | '));
    await page.close();
  }

  await browser.close();
  srv.close();
  console.log(failures ? `\n${failures} test(er) feilet` : '\nAlle tester bestått');
  process.exit(failures ? 1 : 0);
})();
