'use strict';

const $ = (id) => document.getElementById(id);

const KINDS = {
  word: {
    label: 'Word',
    color: 'var(--word)',
    exts: ['docx', 'docm', 'dotx', 'dotm'],
    libs: ['lib/jszip.min.js', 'lib/docx-preview.min.js'],
  },
  excel: {
    label: 'Excel',
    color: 'var(--excel)',
    exts: ['xlsx', 'xlsm', 'xlsb', 'xltx', 'xltm', 'xls', 'ods', 'fods', 'csv', 'tsv'],
    libs: ['lib/xlsx.full.min.js'],
  },
  powerpoint: {
    label: 'PowerPoint',
    color: 'var(--powerpoint)',
    exts: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'],
    libs: ['lib/jszip.min.js', 'lib/pptx-preview.umd.js'],
  },
};

const LEGACY_FORMATS = {
  doc: 'Word 97–2003 documents (.doc)',
  dot: 'Word 97–2003 templates (.dot)',
  ppt: 'PowerPoint 97–2003 presentations (.ppt)',
  pps: 'PowerPoint 97–2003 slide shows (.pps)',
  pot: 'PowerPoint 97–2003 templates (.pot)',
  odt: 'OpenDocument text files (.odt)',
  odp: 'OpenDocument presentations (.odp)',
  rtf: 'Rich Text files (.rtf)',
};

const ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];
const EXCEL_ROW_CHUNK = 1000;
const EXCEL_MAX_COLS = 500;

const state = {
  buf: null,
  name: '',
  url: '',
  kind: null,
  zoom: 1,
  cleanup: null,
};

const isExtension = typeof chrome !== 'undefined' && !!chrome.runtime?.id;

// ---------------------------------------------------------------- utilities

const scripts = new Map();
function loadScript(src) {
  if (!scripts.has(src)) {
    scripts.set(
      src,
      new Promise((resolve, reject) => {
        const el = document.createElement('script');
        el.src = src;
        el.onload = resolve;
        el.onerror = () => reject(new Error(`Failed to load ${src}`));
        document.head.append(el);
      })
    );
  }
  return scripts.get(src);
}
async function loadLibs(srcs) {
  for (const src of srcs) await loadScript(src);
}

const extOf = (name) => (/\.([a-z0-9]+)$/i.exec(name || '')?.[1] || '').toLowerCase();

function nameFromUrl(url) {
  try {
    const u = new URL(url);
    const last = u.pathname.split('/').filter(Boolean).pop();
    return last ? decodeURIComponent(last) : u.hostname;
  } catch {
    return 'document';
  }
}

function nameFromDisposition(header) {
  if (!header) return null;
  const star = /filename\*\s*=\s*(?:[\w-]*)'[^']*'([^;]+)/i.exec(header);
  if (star) {
    try {
      return decodeURIComponent(star[1].trim().replace(/^"|"$/g, ''));
    } catch {
      // fall through to plain filename
    }
  }
  const plain = /filename\s*=\s*(?:"([^"]*)"|([^;]+))/i.exec(header);
  return plain ? (plain[1] ?? plain[2]).trim() : null;
}

const formatBytes = (n) =>
  n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(0)} KB` : `${(n / 1048576).toFixed(1)} MB`;

const escapeHtml = (s) =>
  String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

const safeHref = (href) => (/^(https?:|mailto:)/i.test(href || '') ? href : null);

// ---------------------------------------------------------------- UI states

function show(which) {
  $('emptyState').hidden = which !== 'empty';
  $('loadingState').hidden = which !== 'loading';
  $('errorState').hidden = which !== 'error';
  $('docView').hidden = which !== 'doc';
  const docShown = which === 'doc';
  $('zoomGroup').hidden = !docShown;
  $('printBtn').hidden = !docShown;
  $('downloadBtn').hidden = !state.buf;
  if (!docShown) setSheetTabs(null);
}

function setLoading(text) {
  $('loadingText').textContent = text;
  show('loading');
}

function showError(title, detail, { fileAccess = false } = {}) {
  $('errorTitle').textContent = title;
  $('errorDetail').textContent = detail || '';
  $('errorFileAccessBtn').hidden = !(fileAccess && isExtension);
  $('errorDownloadBtn').hidden = !(state.url && /^https?:/i.test(state.url));
  show('error');
}

function setFileHeader(name, kind) {
  $('fileName').textContent = name || '';
  $('fileName').title = name || '';
  document.title = name ? `${name} – DocLens` : 'DocLens';
  const badge = $('kindBadge');
  if (kind) {
    badge.textContent = KINDS[kind].label;
    badge.style.setProperty('--kind-color', KINDS[kind].color);
    badge.hidden = false;
  } else {
    badge.hidden = true;
  }
}

// ---------------------------------------------------------------- loading

function fetchViaXhr(url) {
  // fetch() rejects file: URLs; XHR works when "Allow access to file URLs" is on.
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open('GET', url);
    xhr.responseType = 'arraybuffer';
    xhr.onload = () => {
      if ((xhr.status === 0 || xhr.status === 200) && xhr.response) {
        resolve({ buf: xhr.response, name: nameFromUrl(url) });
      } else {
        reject(Object.assign(new Error('The local file could not be read.'), { fileAccess: true }));
      }
    };
    xhr.onerror = () =>
      reject(
        Object.assign(
          new Error(
            'Chrome blocked access to this local file. Enable "Allow access to file URLs" for DocLens, then reload this page.'
          ),
          { fileAccess: true }
        )
      );
    xhr.send();
  });
}

async function fetchFile(url) {
  if (/^file:/i.test(url)) return fetchViaXhr(url);

  let res;
  try {
    res = await fetch(url, { credentials: 'include' });
  } catch (err) {
    throw new Error(`Network error while downloading the file (${err.message}).`);
  }
  if (!res.ok) throw new Error(`The server responded with ${res.status} ${res.statusText}.`.replace(' .', '.'));

  const name = nameFromDisposition(res.headers.get('content-disposition')) || nameFromUrl(res.url || url);
  const total = Number(res.headers.get('content-length')) || 0;

  if (!res.body) return { buf: await res.arrayBuffer(), name };

  const reader = res.body.getReader();
  const parts = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    parts.push(value);
    received += value.length;
    $('loadingText').textContent = total
      ? `Downloading… ${Math.min(100, Math.round((received / total) * 100))}%`
      : `Downloading… ${formatBytes(received)}`;
  }
  const bytes = new Uint8Array(received);
  let offset = 0;
  for (const p of parts) {
    bytes.set(p, offset);
    offset += p.length;
  }
  return { buf: bytes.buffer, name };
}

async function detectKind(buf, name) {
  const ext = extOf(name);
  for (const [kind, info] of Object.entries(KINDS)) if (info.exts.includes(ext)) return kind;

  const head = new Uint8Array(buf, 0, Math.min(8, buf.byteLength));
  if (head[0] === 0x50 && head[1] === 0x4b) {
    // ZIP container: look inside to tell OOXML / ODF formats apart.
    await loadScript('lib/jszip.min.js');
    const zip = await JSZip.loadAsync(buf);
    if (zip.file('word/document.xml')) return 'word';
    if (zip.file('ppt/presentation.xml')) return 'powerpoint';
    if (zip.file('xl/workbook.xml') || zip.file('xl/workbook.bin')) return 'excel';
    const types = zip.file('[Content_Types].xml');
    if (types) {
      const xml = await types.async('string');
      if (/wordprocessingml/.test(xml)) return 'word';
      if (/presentationml/.test(xml)) return 'powerpoint';
      if (/spreadsheetml/.test(xml)) return 'excel';
    }
    const mime = zip.file('mimetype');
    if (mime && /spreadsheet/.test(await mime.async('string'))) return 'excel';
    return null;
  }
  // OLE2 compound file: of the legacy binary formats only .xls is supported.
  if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) {
    return LEGACY_FORMATS[ext] ? null : 'excel';
  }
  return null;
}

async function openDocument({ buf, name, url = '' }) {
  state.cleanup?.();
  state.cleanup = null;
  Object.assign(state, { buf, name, url, kind: null, workbookTabs: null });
  setFileHeader(name, null);
  setLoading('Reading document…');

  try {
    const kind = await detectKind(buf, name);
    if (!kind) {
      const ext = extOf(name);
      showError(
        'This file type isn’t supported',
        LEGACY_FORMATS[ext]
          ? `${LEGACY_FORMATS[ext]} use an old binary format that can’t be displayed. Re-save the file as .docx, .xlsx or .pptx and open it again.`
          : `“${name}” doesn’t look like a Word, Excel or PowerPoint document.`
      );
      return;
    }
    state.kind = kind;
    setFileHeader(name, kind);
    setLoading(`Rendering ${KINDS[kind].label} document…`);
    await loadLibs(KINDS[kind].libs);

    const view = $('docView');
    view.replaceChildren();
    view.className = `doc doc-${kind}`;
    view.style.zoom = '';
    state.zoom = 1;
    updateZoomLabel();
    $('main').scrollTo(0, 0);

    // PowerPoint layout needs a visible container to measure.
    if (kind === 'powerpoint') show('doc');
    await RENDERERS[kind](buf, view);
    show('doc');
    if (kind === 'excel') setSheetTabs(state.workbookTabs);
    if (kind === 'word') {
      await document.fonts.ready; // late font loads (e.g. Devanagari fallback) reflow text
      paginateWord(view);
      // Some layout settles a frame later; a second pass splits any page that grew.
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      paginateWord(view);
      fitToWidth();
    }
  } catch (err) {
    console.error(err);
    showError('Couldn’t open this document', err?.message || String(err));
  }
}

async function openUrl(url, name) {
  state.url = url;
  state.buf = null;
  setFileHeader(name || nameFromUrl(url), null);
  setLoading('Downloading…');
  try {
    if (/^file:/i.test(url) && isExtension && !(await chrome.extension.isAllowedFileSchemeAccess())) {
      showError(
        'Allow DocLens to open files on your computer',
        'Click the button below, turn on “Allow access to file URLs” for DocLens, then reload this page.',
        { fileAccess: true }
      );
      return;
    }
    const file = await fetchFile(url);
    await openDocument({ buf: file.buf, name: name || file.name, url });
  } catch (err) {
    console.error(err);
    showError('Couldn’t load this document', err.message, { fileAccess: !!err.fileAccess });
  }
}

async function openLocalFile(file) {
  if (!file) return;
  history.replaceState(null, '', location.pathname);
  const buf = await file.arrayBuffer();
  await openDocument({ buf, name: file.name });
}

// ---------------------------------------------------------------- renderers

// Word's built-in paragraph defaults (space after 8pt, 1.16 line spacing),
// which Word applies when styles.xml has no <w:pPrDefault>.
const WORD_DEFAULT_PPR =
  '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault>';

// Fills in what Word assumes but docx-preview doesn't. Returns the (possibly
// rewritten) package and whether the document asks for auto-hyphenation.
async function prepareWordPackage(buf) {
  const zip = await JSZip.loadAsync(buf);
  let changed = false;

  const stylesFile = zip.file('word/styles.xml');
  if (stylesFile) {
    let xml = await stylesFile.async('string');
    if (!/<w:pPrDefault[\s>/]/.test(xml)) {
      if (/<w:docDefaults\s*\/>/.test(xml)) {
        xml = xml.replace(/<w:docDefaults\s*\/>/, `<w:docDefaults>${WORD_DEFAULT_PPR}</w:docDefaults>`);
      } else if (xml.includes('</w:docDefaults>')) {
        xml = xml.replace('</w:docDefaults>', `${WORD_DEFAULT_PPR}</w:docDefaults>`);
      } else {
        xml = xml.replace(/(<w:styles\b[^>]*>)/, `$1<w:docDefaults>${WORD_DEFAULT_PPR}</w:docDefaults>`);
      }
      zip.file('word/styles.xml', xml);
      changed = true;
    }
  }

  const settingsFile = zip.file('word/settings.xml');
  const settings = settingsFile ? await settingsFile.async('string') : '';
  const autoHyphenation = /<w:autoHyphenation(?![^>]*w:val="(?:false|0|off)")[^>]*\/?>/.test(settings);

  return { buf: changed ? await zip.generateAsync({ type: 'arraybuffer' }) : buf, autoHyphenation };
}

const SERIF_FONT_RE = /times|cambria|georgia|garamond|antiqua|palatino|bookman|schoolbook|constantia|baskerville|minion|mincho|song/i;
const MONO_FONT_RE = /courier|consolas|mono|menlo|lucida console|typewriter/i;
const SYMBOL_FONT_RE = /symbol|wingdings|webdings|dingbats/i;

// docx-preview emits bare font names ("Aptos", "Arial Rounded MT"). When the
// font isn't installed Chrome drops to Times, which changes text widths and
// breaks layouts that rely on spacing. Append a metric-similar fallback stack.
function withFontFallback(value) {
  if (!value || value.includes('var(') || /(^|,)\s*(serif|sans-serif|monospace)\s*$/i.test(value)) return value;
  const first = value.split(',')[0].trim().replace(/^["']|["']$/g, '');
  if (SYMBOL_FONT_RE.test(first)) return value;
  const stack = MONO_FONT_RE.test(first)
    ? ['Consolas', '"Courier New"', 'monospace']
    : SERIF_FONT_RE.test(first) && !/sans/i.test(first)
      ? ['Cambria', '"Times New Roman"', 'serif']
      : ['Calibri', 'Carlito', '"Segoe UI"', 'Arial', 'sans-serif'];
  return [value, ...stack].join(', ');
}

function dominantFont(view) {
  const weight = new Map();
  for (const span of [...view.querySelectorAll('section.docx span')].slice(0, 800)) {
    const len = span.childElementCount ? 0 : span.textContent.length;
    if (!len) continue;
    const family = getComputedStyle(span).fontFamily;
    weight.set(family, (weight.get(family) || 0) + len);
  }
  return [...weight].sort((a, b) => b[1] - a[1])[0]?.[0] || 'Calibri, sans-serif';
}

function naturalLineRatio(fontFamily) {
  const probe = document.createElement('div');
  probe.style.cssText = 'position:absolute;visibility:hidden;font-size:100px;line-height:normal;white-space:nowrap';
  probe.style.fontFamily = fontFamily;
  probe.textContent = 'Hg';
  document.body.append(probe);
  const ratio = probe.getBoundingClientRect().height / 100;
  probe.remove();
  return ratio;
}

function fixWordRendering(view, styleHost) {
  for (const style of styleHost.querySelectorAll('style')) {
    style.textContent = style.textContent.replace(
      /(font-family|--docx-[\w-]+-font)\s*:\s*([^;}\n]+)/g,
      (_, prop, value) => `${prop}: ${withFontFallback(value.trim())}`
    );
  }
  for (const el of view.querySelectorAll('[style*="font-family"]')) {
    el.style.fontFamily = withFontFallback(el.style.fontFamily);
  }

  // List markers end with a tab that should reach the hanging indent; HTML
  // collapses it to a space, so pad the marker to the hanging width instead.
  for (const p of view.querySelectorAll('p[class*="docx-num-"]')) {
    const hanging = -parseFloat(getComputedStyle(p).textIndent);
    const marker = getComputedStyle(p, '::before').content; // tab serializes as "\9 "
    if (hanging > 0 && /\t|\\9(\s|"|$)/.test(marker)) {
      p.style.setProperty('--doclens-hang', `${hanging}px`);
    }
  }

  // Word's "multiple" line spacing (e.g. 1.16) multiplies the font's natural
  // line height; CSS unitless line-height multiplies the font size. Scale by
  // the main font's natural ratio (Arial ~1.15, Calibri ~1.22).
  const ratio = naturalLineRatio(dominantFont(view));
  if (ratio > 1.01 && ratio < 1.6) {
    const scale = (v) => String(+(parseFloat(v) * ratio).toFixed(3));
    for (const style of styleHost.querySelectorAll('style')) {
      style.textContent = style.textContent.replace(/(line-height\s*:\s*)([\d.]+)(?=\s*[;}\n])/g, (_, p, v) => p + scale(v));
    }
    for (const el of view.querySelectorAll('[style*="line-height"]')) {
      if (/^[\d.]+$/.test(el.style.lineHeight)) el.style.lineHeight = scale(el.style.lineHeight);
    }
  }

  // Word sizes each line by its largest run, plus the paragraph mark (which
  // uses the paragraph style's font) on the last line only. CSS instead adds
  // the paragraph's own font (often the 16px browser-default serif) to every
  // line. So: append an invisible, unstyled run as the paragraph mark, and give
  // the paragraph its largest run's font so its strut adds nothing.
  // (All reads first, then writes, to avoid a style recalc per paragraph.)
  const paragraphs = [...view.querySelectorAll('section.docx p')];
  for (const p of paragraphs) {
    const mark = document.createElement('span');
    mark.className = 'doclens-mark'; // empty: still sets line height, never wraps
    p.append(mark);
  }
  const struts = paragraphs.map((p) => {
    let largest = null;
    for (const s of p.querySelectorAll('span:not(.doclens-mark)')) {
      if (!s.textContent) continue;
      const cs = getComputedStyle(s);
      const size = parseFloat(cs.fontSize);
      if (!largest || size > largest.size) largest = { size, family: cs.fontFamily };
    }
    if (largest) return largest;
    const mark = getComputedStyle(p.lastChild); // empty paragraph: just the mark
    return { size: parseFloat(mark.fontSize), family: mark.fontFamily };
  });
  paragraphs.forEach((p, i) => {
    p.style.fontSize = `${struts[i].size}px`;
    p.style.fontFamily = struts[i].family;
  });

  // Word lays tables out on their stored grid. Chrome's auto layout instead
  // redistributes widths by content (squeezing wrap-able columns), so pin
  // tables whose grid columns all have widths.
  for (const table of view.querySelectorAll('table')) {
    const cols = [...(table.querySelector(':scope > colgroup')?.children || [])];
    if (!cols.length || !cols.every((c) => /pt$/.test(c.style.width))) continue;
    table.style.tableLayout = 'fixed';
    if (!table.style.width || table.style.width === 'auto') {
      table.style.width = `${cols.reduce((sum, c) => sum + parseFloat(c.style.width), 0)}pt`;
    }
  }

  // VML straight connectors (horizontal rules) come out as an empty zero-height <svg>.
  const SVG_NS = 'http://www.w3.org/2000/svg';
  for (const svg of view.querySelectorAll('svg')) {
    const g = svg.firstElementChild;
    if (!g || g.localName !== 'g' || g.childElementCount || !g.getAttribute('stroke')) continue;
    if (parseFloat(svg.style.height) !== 0) continue;
    const color = g.getAttribute('stroke');
    const width = g.getAttribute('stroke-width') || '1pt';
    svg.removeAttribute('width');
    svg.removeAttribute('height');
    // Inline, like Word: the paragraph keeps its normal line height around the rule.
    Object.assign(svg.style, { height: width, display: 'inline-block', verticalAlign: 'middle', overflow: 'visible' });
    const line = document.createElementNS(SVG_NS, 'line');
    for (const [k, v] of Object.entries({ x1: '0', y1: '50%', x2: '100%', y2: '50%' })) line.setAttribute(k, v);
    line.setAttribute('stroke', /^[0-9a-f]{6}$/i.test(color) ? `#${color}` : color);
    line.setAttribute('stroke-width', width);
    g.append(line);
  }
}

// docx-preview only starts a new page at explicit breaks. Flow overflowing
// blocks onto new pages the way Word does, splitting tables between rows.
// Must run while the document is laid out (visible).
function paginateWord(view) {
  const sections = [...view.querySelectorAll('.docx-wrapper > section.docx')];
  for (const first of sections) {
    let section = first;
    for (let guard = 0; section && guard < 2000; guard++) section = splitOverflowingPage(section);
  }
}

function splitOverflowingPage(section) {
  const article = section.querySelector(':scope > article');
  const pageHeight = parseFloat(getComputedStyle(section).minHeight); // px; the Word page height
  if (!article || !pageHeight || !['auto', '1'].includes(getComputedStyle(article).columnCount)) return null;

  const sectionTop = section.getBoundingClientRect().top;
  const limit = sectionTop + pageHeight - parseFloat(getComputedStyle(section).paddingBottom);
  const blocks = [...article.children];
  const index = blocks.findIndex((b) => b.getBoundingClientRect().bottom > limit + 0.5);
  if (index < 0) return null;

  let moved = blocks.slice(index);
  const block = blocks[index];
  if (block.localName === 'table') {
    const tail = splitTable(block, limit);
    if (tail) moved = [tail, ...blocks.slice(index + 1)];
    else if (index === 0) moved = blocks.slice(1); // taller than a page and unsplittable: let it overflow
  } else if (index === 0) {
    moved = blocks.slice(1);
  }
  if (!moved.length) return null;

  const next = section.cloneNode(false);
  for (const part of section.children) {
    if (part === article) {
      const body = article.cloneNode(false);
      body.append(...moved);
      next.append(body);
    } else {
      next.append(part.cloneNode(true)); // header / footer
    }
  }
  section.after(next);
  return next;
}

// Moves everything below `limit` into a copy of the table and returns it, or
// null when there is no clean break. Like Word, a row may break across pages
// (between paragraphs inside its cells) unless merged cells are involved.
function splitTable(table, limit) {
  const body = table.tBodies[0] || table;
  const rows = [...body.rows];
  let cut = rows.findIndex((r) => r.getBoundingClientRect().bottom > limit + 0.5);
  if (cut < 0) return null;
  const spansInto = (k) =>
    rows.slice(0, k).some((r, i) => [...r.cells].some((c) => c.rowSpan > 1 && i + c.rowSpan > k));

  let tailRows = null;
  const row = rows[cut];
  if (!spansInto(cut) && ![...row.cells].some((c) => c.rowSpan > 1)) {
    const rest = splitRow(row, limit);
    if (rest) tailRows = [rest, ...rows.slice(cut + 1)];
  }
  if (!tailRows) {
    // Whole rows only: walk back until no rowspan reaches into the cut row.
    while (cut > 0 && spansInto(cut)) cut--;
    if (cut <= 0) return null;
    tailRows = rows.slice(cut);
  }

  const copy = table.cloneNode(false);
  for (const child of table.children) {
    if (child.localName === 'colgroup') copy.append(child.cloneNode(true));
  }
  const copyBody = body === table ? copy : copy.appendChild(body.cloneNode(false));
  copyBody.append(...tailRows);
  table.after(copy);
  return copy;
}

// Splits a row's cells between paragraphs at `limit`; returns the continuation
// row, or null if nothing would stay on this page or nothing would move.
function splitRow(row, limit) {
  const next = row.cloneNode(false);
  let stays = false;
  let moves = false;
  for (const cell of row.cells) {
    const copy = cell.cloneNode(false);
    const blocks = [...cell.children];
    const i = blocks.findIndex((b) => b.getBoundingClientRect().bottom > limit + 0.5);
    if (i !== 0 && blocks.length) stays = true;
    if (i >= 0) {
      copy.append(...blocks.slice(i));
      moves = true;
    }
    next.append(copy);
  }
  if (!(stays && moves)) return null;
  // A fixed row height would keep both halves at full height.
  row.style.height = next.style.height = '';
  return next;
}

async function renderWord(buf, view) {
  const styles = document.createElement('div');
  const body = document.createElement('div');
  view.append(styles, body);
  const prepared = await prepareWordPackage(buf);
  view.classList.toggle('docx-autohyphen', prepared.autoHyphenation);
  await docx.renderAsync(prepared.buf, body, styles, {
    className: 'docx',
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: true,
    experimental: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    renderComments: false,
    useBase64URL: true,
  });
  fixWordRendering(view, styles);
}

// [Content_Types].xml sometimes lists parts that aren't in the package
// (PowerPoint ignores them). pptx-preview aborts on those and silently renders
// nothing, so drop dangling entries first.
async function pruneMissingParts(buf) {
  const zip = await JSZip.loadAsync(buf);
  const types = zip.file('[Content_Types].xml');
  if (!types) return buf;
  let changed = false;
  const xml = (await types.async('string')).replace(/<Override\b[^>]*?PartName="\/([^"]+)"[^>]*\/>/g, (tag, part) => {
    if (zip.file(part) || zip.file(decodeURIComponent(part))) return tag;
    changed = true;
    return '';
  });
  if (!changed) return buf;
  zip.file('[Content_Types].xml', xml);
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function renderPowerPoint(buf, view) {
  const available = $('main').clientWidth - 48;
  const width = Math.max(320, Math.min(1280, available));
  const previewer = pptxPreview.init(view, { width, mode: 'list' });
  const pptx = await previewer.preview(await pruneMissingParts(buf));
  state.cleanup = () => previewer.destroy?.();
  if (!pptx?.slides?.length) throw new Error('No slides could be read from this presentation.');
}

async function renderExcel(buf, view) {
  const wb = XLSX.read(buf, { type: 'array', cellStyles: true, cellDates: true, cellNF: true });
  const sheets = wb.SheetNames.map((name, i) => ({ name, index: i, hidden: !!wb.Workbook?.Sheets?.[i]?.Hidden }))
    .filter((s) => !s.hidden);

  const showSheet = (sheetName) => {
    view.replaceChildren();
    $('main').scrollTo(0, 0);
    renderSheet(wb.Sheets[sheetName], view);
    for (const btn of $('sheetTabs').children) {
      btn.setAttribute('aria-selected', String(btn.dataset.sheet === sheetName));
    }
  };

  state.workbookTabs = sheets.length > 1 ? { sheets, onSelect: showSheet } : null;
  if (!sheets.length) {
    view.innerHTML = '<div class="xl-empty">This workbook has no visible sheets.</div>';
    return;
  }
  renderSheet(wb.Sheets[sheets[0].name], view);
}

function renderSheet(ws, view) {
  const ref = ws?.['!ref'];
  if (!ref) {
    view.innerHTML = '<div class="xl-empty">This sheet is empty.</div>';
    return;
  }
  const range = XLSX.utils.decode_range(ref);
  const lastCol = Math.min(range.e.c, range.s.c + EXCEL_MAX_COLS - 1);
  const colsMeta = ws['!cols'] || [];
  const rowsMeta = ws['!rows'] || [];

  const visibleCols = [];
  for (let c = range.s.c; c <= lastCol; c++) if (!colsMeta[c]?.hidden) visibleCols.push(c);

  // Merged ranges: top-left cell gets row/colspan, the rest are skipped.
  const mergeStart = new Map();
  const covered = new Set();
  for (const m of ws['!merges'] || []) {
    let colspan = 0;
    for (let c = m.s.c; c <= m.e.c; c++) if (!colsMeta[c]?.hidden) colspan++;
    let rowspan = 0;
    for (let r = m.s.r; r <= m.e.r; r++) if (!rowsMeta[r]?.hidden) rowspan++;
    mergeStart.set(`${m.s.r},${m.s.c}`, { rowspan, colspan });
    for (let r = m.s.r; r <= m.e.r; r++)
      for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) covered.add(`${r},${c}`);
  }

  const colWidth = (c) => {
    const meta = colsMeta[c];
    if (meta?.wpx) return Math.round(meta.wpx);
    if (meta?.wch) return Math.round(meta.wch * 7 + 5);
    if (meta?.width) return Math.round(meta.width * 7);
    return 72;
  };

  const table = document.createElement('table');
  table.className = 'xl-table';
  let html = '<colgroup><col style="width:48px">';
  for (const c of visibleCols) html += `<col style="width:${colWidth(c)}px">`;
  html += '</colgroup><thead><tr><th class="corner"></th>';
  for (const c of visibleCols) html += `<th>${XLSX.utils.encode_col(c)}</th>`;
  html += '</tr></thead>';
  table.innerHTML = html;
  const tbody = document.createElement('tbody');
  table.append(tbody);
  view.append(table);

  const cellHtml = (r, c) => {
    const key = `${r},${c}`;
    if (covered.has(key)) return '';
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    const merge = mergeStart.get(key);
    const attrs = [];
    const classes = [];
    const styles = [];
    if (merge) {
      if (merge.rowspan > 1) attrs.push(`rowspan="${merge.rowspan}"`);
      if (merge.colspan > 1) attrs.push(`colspan="${merge.colspan}"`);
    }
    let content = '';
    if (cell) {
      let text = cell.w;
      if (text == null && cell.v != null) text = cell.v instanceof Date ? cell.v.toLocaleString() : String(cell.v);
      text = text ?? '';
      if (cell.t === 'n' || cell.t === 'd') classes.push('num');
      else if (cell.t === 'b') classes.push('bool');
      else if (cell.t === 'e') classes.push('err');
      if (cell.t === 's' && /\n/.test(text)) classes.push('wrap');
      const fill = cell.s?.patternType === 'solid' ? cell.s.fgColor?.rgb : null;
      if (fill && /^[0-9a-f]{6,8}$/i.test(fill)) styles.push(`background:#${fill.slice(-6)}`);
      const href = safeHref(cell.l?.Target);
      content = href
        ? `<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(text)}</a>`
        : escapeHtml(text);
      if (cell.c?.length) attrs.push(`title="${escapeHtml(cell.c.map((n) => n.t).join('\n'))}"`);
    }
    if (classes.length) attrs.push(`class="${classes.join(' ')}"`);
    if (styles.length) attrs.push(`style="${styles.join(';')}"`);
    return `<td${attrs.length ? ' ' + attrs.join(' ') : ''}>${content}</td>`;
  };

  let nextRow = range.s.r;
  const more = document.createElement('div');
  more.className = 'xl-more';

  const appendRows = () => {
    let html = '';
    let rendered = 0;
    while (nextRow <= range.e.r && rendered < EXCEL_ROW_CHUNK) {
      const r = nextRow++;
      const meta = rowsMeta[r];
      if (meta?.hidden) continue;
      const height = meta?.hpx ? ` style="height:${Math.round(meta.hpx)}px"` : '';
      html += `<tr${height}><th>${r + 1}</th>`;
      for (const c of visibleCols) html += cellHtml(r, c);
      html += '</tr>';
      rendered++;
    }
    tbody.insertAdjacentHTML('beforeend', html);
    updateMore();
  };

  const updateMore = () => {
    const remaining = range.e.r - nextRow + 1;
    const notes = [];
    if (remaining > 0) notes.push(`Showing rows ${range.s.r + 1}–${nextRow} of ${range.e.r + 1}.`);
    if (range.e.c > lastCol) notes.push(`Only the first ${EXCEL_MAX_COLS} columns are shown.`);
    more.replaceChildren();
    if (!notes.length) {
      more.remove();
      return;
    }
    more.append(notes.join(' '));
    if (remaining > 0) {
      const btn = document.createElement('button');
      btn.className = 'btn';
      btn.textContent = `Load ${Math.min(EXCEL_ROW_CHUNK, remaining).toLocaleString()} more rows`;
      btn.onclick = appendRows;
      more.append(btn);
    }
    if (!more.isConnected) view.append(more);
  };

  appendRows();
}

const RENDERERS = { word: renderWord, excel: renderExcel, powerpoint: renderPowerPoint };

// ---------------------------------------------------------------- sheet tabs

function setSheetTabs(tabs) {
  const nav = $('sheetTabs');
  nav.replaceChildren();
  nav.hidden = !tabs;
  document.body.classList.toggle('has-tabs', !!tabs);
  if (!tabs) return;
  tabs.sheets.forEach((sheet, i) => {
    const btn = document.createElement('button');
    btn.textContent = sheet.name;
    btn.dataset.sheet = sheet.name;
    btn.setAttribute('aria-selected', String(i === 0));
    btn.onclick = () => tabs.onSelect(sheet.name);
    nav.append(btn);
  });
}

// ---------------------------------------------------------------- zoom

function updateZoomLabel() {
  $('zoomReset').textContent = `${Math.round(state.zoom * 100)}%`;
}

function setZoom(z) {
  state.zoom = Math.min(3, Math.max(0.25, z));
  $('docView').style.zoom = state.zoom === 1 ? '' : String(state.zoom);
  updateZoomLabel();
}

function stepZoom(dir) {
  const z = state.zoom;
  const next = dir > 0 ? ZOOM_STEPS.find((s) => s > z + 0.001) : [...ZOOM_STEPS].reverse().find((s) => s < z - 0.001);
  if (next) setZoom(next);
}

function fitToWidth() {
  const pages = $('docView').querySelectorAll('section.docx');
  const natural = Math.max(0, ...[...pages].map((p) => p.offsetWidth)) + 32;
  const available = $('main').clientWidth - 16;
  if (natural > available) setZoom(Math.max(0.25, available / natural));
}

// ---------------------------------------------------------------- actions

function downloadOriginal() {
  if (!state.buf) return;
  const url = URL.createObjectURL(new Blob([state.buf]));
  const a = document.createElement('a');
  a.href = url;
  a.download = state.name || 'document';
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function downloadFromUrl() {
  if (!state.url) return;
  if (isExtension && chrome.downloads) chrome.downloads.download({ url: state.url });
  else window.open(state.url, '_blank', 'noopener');
}

const pickFile = () => $('fileInput').click();

$('openBtn').onclick = pickFile;
$('chooseBtn').onclick = pickFile;
$('errorOpenBtn').onclick = pickFile;
$('downloadBtn').onclick = downloadOriginal;
$('errorDownloadBtn').onclick = downloadFromUrl;
$('errorFileAccessBtn').onclick = () => chrome.tabs.create({ url: `chrome://extensions/?id=${chrome.runtime.id}` });
$('printBtn').onclick = () => window.print();
$('zoomIn').onclick = () => stepZoom(1);
$('zoomOut').onclick = () => stepZoom(-1);
$('zoomReset').onclick = () => setZoom(1);
$('fileInput').onchange = (e) => {
  openLocalFile(e.target.files[0]);
  e.target.value = '';
};

document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') {
    e.preventDefault();
    pickFile();
  }
});

// Links inside rendered documents: keep in-document bookmarks, open web links
// in a new tab, block everything else (javascript:, file:, etc.).
$('docView').addEventListener('click', (e) => {
  const a = e.target.closest('a[href]');
  if (!a) return;
  const href = a.getAttribute('href');
  if (href.startsWith('#')) return;
  e.preventDefault();
  if (safeHref(a.href)) window.open(a.href, '_blank', 'noopener');
});

// Drag and drop anywhere on the page.
let dragDepth = 0;
const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
window.addEventListener('dragenter', (e) => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth++;
  $('dropOverlay').hidden = false;
});
window.addEventListener('dragover', (e) => {
  if (hasFiles(e)) e.preventDefault();
});
window.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) $('dropOverlay').hidden = true;
});
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  $('dropOverlay').hidden = true;
  openLocalFile(e.dataTransfer?.files?.[0]);
});

// ---------------------------------------------------------------- startup

// The background worker opens viewer.html#[name=<enc>&]url=<url>. The URL is
// either percent-encoded or (from the redirect rule) raw, so take everything
// after "url=" and decode only when it is encoded.
function parseHash() {
  const hash = location.hash.slice(1);
  const m = /(?:^|&)url=(.*)$/.exec(hash);
  if (!m) return null;
  let url = m[1];
  if (/^[a-z]+%3A/i.test(url)) {
    try {
      url = decodeURIComponent(url);
    } catch {
      // leave as-is
    }
  }
  const nameMatch = /(?:^|&)name=([^&]*)/.exec(hash.slice(0, m.index + 1));
  let name = '';
  if (nameMatch) {
    try {
      name = decodeURIComponent(nameMatch[1]);
    } catch {
      name = nameMatch[1];
    }
  }
  return { url, name };
}

function start() {
  const target = parseHash() || (() => {
    const url = new URLSearchParams(location.search).get('url');
    return url ? { url, name: '' } : null;
  })();
  if (target?.url) openUrl(target.url, target.name);
  else show('empty');
}

window.addEventListener('hashchange', start);
start();
