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
  window.__log = [];
  const shapes = [
    { id: '2', name: 'Picture 1', type: 'Image', left: 36, top: 36, width: 144, height: 144 },
    { id: '3', name: 'Testbilde', type: 'Image', left: 216, top: 72, width: 360, height: 225 },
  ];
  let selected = '3';
  const coll = (items) => ({ items, load() {} });
  const ctx = {
    sync: async () => {},
    presentation: {
      getSelectedShapes: () => coll(shapes.filter((s) => s.id === selected)),
      getSelectedSlides: () => coll([{ id: '256#', exportAsBase64: () => ({ value: pptxB64 }) }]),
      slides: { getItem: () => ({ shapes: { getItemOrNullObject: (id) => {
        const s = shapes.find((x) => x.id === id);
        return { isNullObject: !s, delete() { shapes.splice(shapes.indexOf(s), 1); window.__log.push({ op: 'delete', id }); } };
      } } }) },
    },
  };
  window.PowerPoint = { run: (fn) => fn(ctx) };
  window.Office = {
    HostType: { PowerPoint: 'PowerPoint' }, CoercionType: { Image: 'image' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' }, EventType: { DocumentSelectionChanged: 'sel' },
    onReady(cb) { setTimeout(() => cb({ host: 'PowerPoint' }), 10); },
    context: {
      requirements: { isSetSupported: () => true },
      document: {
        addHandlerAsync() {},
        getSelectedDataAsync(t, cb) { cb({ status: 'failed' }); },
        setSelectedDataAsync(data, opts, cb) {
          window.__log.push({ op: 'insert', data, opts });
          const id = String(100 + window.__log.length);
          shapes.push({ id, name: 'Picture ' + id, type: 'Image', left: opts.imageLeft, top: opts.imageTop, width: opts.imageWidth, height: opts.imageHeight });
          selected = id;
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
    await page.click('#btnGet');
    await page.waitForSelector('#stage:not([hidden])');
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

  await browser.close();
  srv.close();
  console.log(failures ? `\n${failures} test(er) feilet` : '\nAlle tester bestått');
  process.exit(failures ? 1 : 0);
})();
