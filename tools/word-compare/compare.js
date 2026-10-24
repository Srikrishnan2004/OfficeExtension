// Renders Word's PDF export (pdf.js) and DocLens' rendering, then writes
// side-by-side PNGs: out/<doc>-<n>.png  (left = Word, right = DocLens)
const http = require('http');
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

const ROOT = process.env.WORK_DIR || __dirname; // holds docs/<name>.docx and out/<name>-word.pdf
const EXT = path.resolve(__dirname, '..', '..');
const DOCS = process.argv.slice(2).length ? process.argv.slice(2) : ['draxel', 'gnani'];
const MAX_PAGES = Number(process.env.MAX_PAGES || 60);

const types = { '.mjs': 'text/javascript', '.js': 'text/javascript', '.pdf': 'application/pdf', '.html': 'text/html', '.docx': 'application/octet-stream' };
const server = http.createServer((req, res) => {
  const f = path.join(ROOT, decodeURIComponent(new URL(req.url, 'http://x').pathname));
  if (!fs.existsSync(f)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { 'Content-Type': types[path.extname(f)] || 'application/octet-stream' });
  res.end(fs.readFileSync(f));
});
fs.writeFileSync(path.join(ROOT, 'pdf.html'), `<!doctype html><body style="margin:0;background:#888">
<script type="module">
import * as pdfjs from '/node_modules/pdfjs-dist/build/pdf.mjs';
pdfjs.GlobalWorkerOptions.workerSrc = '/node_modules/pdfjs-dist/build/pdf.worker.mjs';
window.renderPdf = async (url, max) => {
  const pdf = await pdfjs.getDocument(url).promise;
  const out = [];
  for (let i = 1; i <= Math.min(pdf.numPages, max); i++) {
    const page = await pdf.getPage(i);
    const vp = page.getViewport({ scale: 96 / 72 });
    const c = document.createElement('canvas');
    c.width = vp.width; c.height = vp.height;
    await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    out.push(c.toDataURL('image/png'));
  }
  return { pages: out, total: pdf.numPages };
};
window.ready = true;
</script>`);

(async () => {
  await new Promise((r) => server.listen(8799, r));
  const base = 'http://localhost:8799';
  const browser = await puppeteer.launch({ headless: true, pipe: true, enableExtensions: [EXT] });
  const sw = await browser.waitForTarget((t) => t.type() === 'service_worker');
  const viewer = `chrome-extension://${new URL(sw.url()).host}/viewer.html`;

  for (const name of DOCS) {
    // Word side
    const pp = await browser.newPage();
    await pp.goto(`${base}/pdf.html`);
    await pp.waitForFunction('window.ready');
    const word = await pp.evaluate((u, m) => renderPdf(u, m), `${base}/out/${name}-word.pdf`, MAX_PAGES);
    await pp.close();

    // DocLens side
    const p = await browser.newPage();
    const logs = [];
    p.on('console', (m) => { if (['error', 'warn'].includes(m.type())) logs.push(`${m.type()}: ${m.text()}`); });
    p.on('pageerror', (e) => logs.push(`pageerror: ${e.message}`));
    await p.setViewport({ width: 1000, height: 1200 });
    await p.goto(`${viewer}#url=${encodeURIComponent(`${base}/docs/${name}.docx`)}`);
    await p.waitForSelector('#docView:not([hidden]) section.docx, #errorState:not([hidden])', { timeout: 30000 });
    await new Promise((r) => setTimeout(r, 1500)); // images/fonts
    const info = await p.evaluate(() => {
      const secs = [...document.querySelectorAll('section.docx')];
      return {
        zoom: document.getElementById('zoomReset').textContent,
        sections: secs.length,
        rects: secs.map((s) => { const r = s.getBoundingClientRect(); return { y: r.top + document.querySelector('main').scrollTop, h: r.height, w: r.width, x: r.left }; }),
        err: document.getElementById('errorState').hidden ? null : document.getElementById('errorDetail').textContent,
      };
    });
    console.log(name, 'word pages', word.total, '| doclens sections', info.sections, 'zoom', info.zoom, info.err || '');
    // Slice each DocLens section into page-height chunks
    const shots = [];
    for (const r of info.rects) {
      const pageH = r.h < r.w * 1.6 ? Math.ceil(r.h) : Math.round(r.w * 1.414);
      for (let y = 0; y < r.h && shots.length < MAX_PAGES; y += pageH) {
        await p.evaluate((top) => document.querySelector('main').scrollTo(0, top), r.y + y - 52);
        await p.setViewport({ width: 1000, height: Math.min(pageH, Math.ceil(r.h - y)) + 52 });
        await p.evaluate((top) => document.querySelector('main').scrollTo(0, top), r.y + y - 52);
        const buf = await p.screenshot({ clip: { x: r.x, y: 52, width: r.w, height: Math.min(pageH, Math.ceil(r.h - y)) } });
        shots.push('data:image/png;base64,' + buf.toString('base64'));
      }
    }
    fs.writeFileSync(path.join(ROOT, 'out', `${name}-logs.txt`), logs.join('\n'));
    // also dump DocLens HTML for inspection
    fs.writeFileSync(path.join(ROOT, 'out', `${name}-doclens.html`), await p.$eval('#docView', (e) => e.innerHTML));
    await p.close();

    // Compose side by side
    const cp = await browser.newPage();
    const n = Math.max(word.pages.length, shots.length);
    for (let i = 0; i < n; i++) {
      const png = await cp.evaluate(async (a, b) => {
        const load = (s) => new Promise((r) => { if (!s) return r(null); const im = new Image(); im.onload = () => r(im); im.src = s; });
        const [A, B] = await Promise.all([load(a), load(b)]);
        const w = (A?.width || 800) + (B?.width || 800) + 30, h = Math.max(A?.height || 0, B?.height || 0) + 30;
        const c = document.createElement('canvas'); c.width = w; c.height = h;
        const g = c.getContext('2d'); g.fillStyle = '#666'; g.fillRect(0, 0, w, h);
        g.fillStyle = '#fff'; g.font = 'bold 16px sans-serif'; g.fillText('MS WORD', 10, 20); g.fillText('DOCLENS', (A?.width || 800) + 30, 20);
        if (A) g.drawImage(A, 10, 28); if (B) g.drawImage(B, (A?.width || 800) + 20, 28);
        return c.toDataURL('image/png');
      }, word.pages[i] || null, shots[i] || null);
      fs.writeFileSync(path.join(ROOT, 'out', `${name}-${i + 1}.png`), Buffer.from(png.split(',')[1], 'base64'));
    }
    await cp.close();
  }
  await browser.close();
  server.close();
})().catch((e) => { console.error(e); process.exit(1); });
