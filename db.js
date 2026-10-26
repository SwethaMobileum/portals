'use strict';
// Shared "database" = ONE Excel workbook used by both the Operator and the Enterprise server.
//  - Every change is first written to a small journal file (never blocked, even if Excel has the workbook open).
//  - Then it is applied IN PLACE to the same workbook (rows, status, clickable links, logo thumbnails).
//  - If the workbook is open in Excel (Windows locks it), the changes wait in the journal and are written
//    automatically as soon as it is closed. Both websites always show the merged, up-to-date view.
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const IS_VERCEL = !!process.env.VERCEL;
let cfg = {};
try { cfg = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8')); } catch (e) {}

const BASE_DIR = IS_VERCEL ? '/tmp' : __dirname;
const SHARED_DIR = path.resolve(BASE_DIR, process.env.SHARED_DIR || cfg.sharedDir || 'shared');
const USER_DB = !!(process.env.DATABASE_PATH || cfg.databasePath);
const DB_PATH = path.resolve(BASE_DIR, process.env.DATABASE_PATH || cfg.databasePath || path.join(SHARED_DIR, 'CallerID-Database.xlsx'));
const PREFIX = cfg.sheetPrefix || '';
const STORAGE_DIR = path.join(SHARED_DIR, 'storage', 'logos');
const JOURNAL = path.join(SHARED_DIR, '.pending-changes.jsonl');
const LOCK = path.join(SHARED_DIR, '.db.lock');

const memoryFiles = new Map();

function saveInMemoryFile(fileId, buffer, mimeType, originalName) {
  if (!fileId || !buffer) return;
  memoryFiles.set(fileId, {
    buffer,
    mimeType: mimeType || 'image/png',
    originalName: originalName || 'logo.png',
    dataUrl: `data:${mimeType || 'image/png'};base64,${buffer.toString('base64')}`
  });
}

function getInMemoryFile(fileId) {
  return memoryFiles.get(fileId) || null;
}

function initStorage() {
  try {
    fs.mkdirSync(STORAGE_DIR, { recursive: true });
    fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });

    // Pre-load bundled logos into memoryFiles so they are always in memory & never 404
    const logoDirs = [
      path.join(__dirname, 'shared', 'storage', 'logos'),
      STORAGE_DIR
    ];
    logoDirs.forEach(dir => {
      if (fs.existsSync(dir)) {
        try {
          const files = fs.readdirSync(dir);
          files.forEach(f => {
            const fileId = f.replace(/\.[^.]+$/, '');
            const fullPath = path.join(dir, f);
            try {
              if (fs.statSync(fullPath).isFile()) {
                const buf = fs.readFileSync(fullPath);
                const ext = path.extname(f).slice(1).toLowerCase();
                const mime = ext === 'png' ? 'image/png' : (ext === 'jpg' || ext === 'jpeg') ? 'image/jpeg' : 'image/png';
                saveInMemoryFile(fileId, buf, mime, f);
              }
            } catch (e) {}
          });
        } catch (e) {}
      }
    });

    if (IS_VERCEL) {
      const seedDb = path.join(__dirname, 'shared', 'CallerID-Database.xlsx');
      if (fs.existsSync(seedDb) && !fs.existsSync(DB_PATH)) {
        fs.copyFileSync(seedDb, DB_PATH);
      }
      const seedStorage = path.join(__dirname, 'shared', 'storage', 'logos');
      if (fs.existsSync(seedStorage)) {
        const files = fs.readdirSync(seedStorage);
        files.forEach(f => {
          const dest = path.join(STORAGE_DIR, f);
          if (!fs.existsSync(dest)) {
            try { fs.copyFileSync(path.join(seedStorage, f), dest); } catch (e) {}
          }
        });
      }
    }
  } catch (e) {
    console.error('Storage initialization notice:', e.message);
  }
}
initStorage();

const COLS = {
  Packages: ['packageId', 'name', 'capacity', 'pricePHP', 'brand', 'tagline', 'logoIncluded', 'status', 'createdAt'],
  Requests: ['requestId', 'packageId', 'brandName', 'tagline', 'status', 'logoFileId', 'createdAt', 'departments', 'numbersCount', 'monthlyFeePHP', 'vatPHP', 'totalPHP', 'logoLink', 'preview'],
  Numbers: ['numberId', 'requestId', 'brandName', 'department', 'phoneE164', 'phoneDisplay', 'numberType', 'status', 'createdAt'],
  Files: ['fileId', 'requestId', 'originalName', 'storedName', 'storagePath', 'url', 'mimeType', 'sizeBytes', 'sha256', 'uploadedBy', 'uploadedAt', 'preview'],
};
const MONEY = new Set(['pricePHP', 'monthlyFeePHP', 'vatPHP', 'totalPHP']);     // shown as ₱1,234.00 in Excel
const PESO_FMT = '"₱"#,##0.00';
const DISPLAY_ONLY = new Set(['logoLink', 'preview']);
const EXCEL_IMG = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif' };
const LINK_FONT = { color: { argb: 'FF0563C1' }, underline: true };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const newId = p => `${p}-${crypto.randomBytes(4).toString('hex')}`;

const plain = v => {
  if (v instanceof Date) return v.toISOString();
  if (v && typeof v === 'object') {
    if ('hyperlink' in v) return v.hyperlink;
    if ('result' in v) return v.result ?? '';
    if (v.richText) return v.richText.map(t => t.text).join('');
  }
  return v ?? '';
};

// ---------- cross-process lock (two websites, one workbook) ----------
async function acquire() {
  const t0 = Date.now();
  for (;;) {
    try { const fd = fs.openSync(LOCK, 'wx'); fs.writeSync(fd, String(process.pid)); fs.closeSync(fd); return; }
    catch (e) {
      if (e.code !== 'EEXIST') throw e;
      try { if (Date.now() - fs.statSync(LOCK).mtimeMs > 30000) fs.unlinkSync(LOCK); } catch (x) {}   // stale lock from a crash
      if (Date.now() - t0 > 20000) throw new Error('database is busy, try again');
      await sleep(25 + Math.random() * 40);
    }
  }
}
async function withDbLock(fn) { await acquire(); try { return await fn(); } finally { try { fs.unlinkSync(LOCK); } catch (e) {} } }

// ---------- journal ----------
function readJournal() {
  if (!fs.existsSync(JOURNAL)) return [];
  return fs.readFileSync(JOURNAL, 'utf8').split('\n').filter(Boolean).map(l => { try { return JSON.parse(l); } catch (e) { return null; } }).filter(Boolean);
}
async function enqueue(entry) {
  await withDbLock(async () => fs.appendFileSync(JOURNAL, JSON.stringify(entry) + '\n'));
  if (IS_VERCEL) {
    await flush();
  } else {
    scheduleFlush(150);
  }
}

// ---------- reading the workbook (cached until the file changes) ----------
const EMPTY = () => ({ packages: [], requests: [], numbers: [], files: [] });
let cache = { key: null, state: EMPTY() }, loading = null, readError = null;
function rowsOf(ws, name) {
  if (!ws) return [];
  const map = {};
  ws.getRow(1).eachCell({ includeEmpty: false }, (c, n) => { map[String(plain(c.value)).trim()] = n; });
  const out = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r), o = {};
    for (const k of COLS[name]) o[k] = DISPLAY_ONLY.has(k) || !map[k] ? '' : plain(row.getCell(map[k]).value);
    if (o[COLS[name][0]] !== '') out.push(o);
  }
  return out;
}
async function loadWorkbookState() {
  let st; try { st = fs.statSync(DB_PATH); } catch (e) { cache = { key: null, state: EMPTY() }; return cache.state; }
  const key = st.mtimeMs + ':' + st.size;
  if (cache.key === key) return cache.state;
  if (!loading) loading = (async () => {
    try {
      const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(DB_PATH);
      cache = { key, state: { packages: rowsOf(wb.getWorksheet(PREFIX + 'Packages'), 'Packages'), requests: rowsOf(wb.getWorksheet(PREFIX + 'Requests'), 'Requests'), numbers: rowsOf(wb.getWorksheet(PREFIX + 'Numbers'), 'Numbers'), files: rowsOf(wb.getWorksheet(PREFIX + 'Files'), 'Files') } };
      readError = null;
    } catch (e) { readError = e.message; }     // keep the last good copy
    finally { loading = null; }
  })();
  await loading; return cache.state;
}
function applyToState(s, e) {
  if (e.t === 'pkg' && !s.packages.some(p => p.packageId === e.row.packageId)) s.packages.push({ ...e.row });
  if (e.t === 'req') {
    if (!s.requests.some(r => r.requestId === e.row.requestId)) s.requests.push({ ...e.row });
    if (e.file && !s.files.some(f => f.fileId === e.file.fileId)) s.files.push({ ...e.file });
    for (const n of e.numbers || []) if (!s.numbers.some(x => x.numberId === n.numberId)) s.numbers.push({ ...n });
  }
  if (e.t === 'status') {
    const r = s.requests.find(x => x.requestId === e.requestId); if (r) r.status = e.status;
    s.numbers.forEach(n => { if (n.requestId === e.requestId) n.status = e.status; });
  }
}
async function readAll() {           // workbook + changes still waiting in the journal
  const s = JSON.parse(JSON.stringify(await loadWorkbookState()));
  for (const e of readJournal()) applyToState(s, e);
  s.files = (s.files || []).map(f => {
    const mem = memoryFiles.get(f.fileId);
    if (mem && mem.dataUrl) {
      return { ...f, dataUrl: mem.dataUrl };
    }
    // Attempt disk read to populate memory cache & dataUrl
    let full = path.join(SHARED_DIR, f.storagePath || '');
    if (!fs.existsSync(full)) full = path.join(__dirname, 'shared', f.storagePath || '');
    if (fs.existsSync(full)) {
      try {
        const buf = fs.readFileSync(full);
        const mime = f.mimeType || 'image/png';
        saveInMemoryFile(f.fileId, buf, mime, f.originalName);
        return { ...f, dataUrl: `data:${mime};base64,${buf.toString('base64')}` };
      } catch (e) {}
    }
    return f;
  });
  return s;
}

// ---------- writing into the SAME workbook ----------
function ensureSheet(wb, name) {
  let ws = wb.getWorksheet(PREFIX + name);
  if (!ws) {
    ws = wb.addWorksheet(PREFIX + name);
    ws.addRow(COLS[name]).font = { bold: true };
    COLS[name].forEach((h, i) => (ws.getColumn(i + 1).width = Math.max(14, h.length + 4)));
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }
  return ws;
}
function colIdx(ws, name) {         // finds the column by its header (so you may reorder columns); adds it if missing
  let found = 0, last = 0;
  ws.getRow(1).eachCell({ includeEmpty: false }, (c, n) => { last = Math.max(last, n); if (String(plain(c.value)).trim() === name) found = n; });
  if (found) return found;
  const n = last + 1, c = ws.getRow(1).getCell(n); c.value = name; c.font = { bold: true };
  ws.getColumn(n).width = Math.max(14, name.length + 4); return n;
}
function findRow(ws, col, value) {
  for (let r = 2; r <= ws.rowCount; r++) if (String(plain(ws.getRow(r).getCell(col).value)) === String(value)) return r;
  return 0;
}
function newRow(ws, name, obj) {
  const rowNo = (ws.lastRow ? ws.lastRow.number : 0) + 1, row = ws.getRow(rowNo);
  for (const k of COLS[name]) if (!DISPLAY_ONLY.has(k) && obj[k] !== undefined) {
    const c = row.getCell(colIdx(ws, k)); c.value = obj[k];
    if (MONEY.has(k)) c.numFmt = PESO_FMT;
  }
  return rowNo;
}
function imageIndex(wb, buffer, extension) {       // identical logo = stored once in the workbook
  const i = wb.media.findIndex(m => m.buffer && Buffer.compare(Buffer.from(m.buffer), buffer) === 0);
  return i >= 0 ? i : wb.addImage({ buffer, extension });
}
function preview(wb, ws, rowNo, col, file) {
  const ext = file && EXCEL_IMG[file.mimeType];
  const mem = file && memoryFiles.get(file.fileId);
  let full = file && path.join(SHARED_DIR, file.storagePath);
  if (full && !fs.existsSync(full)) {
    full = path.join(__dirname, 'shared', file.storagePath);
  }
  if (ext && mem && mem.buffer) {
    ws.addImage(imageIndex(wb, mem.buffer, ext), { tl: { col: col - 1 + 0.1, row: rowNo - 1 + 0.1 }, ext: { width: 90, height: 56 } });
  } else if (ext && fs.existsSync(full)) {
    ws.addImage(imageIndex(wb, fs.readFileSync(full), ext), { tl: { col: col - 1 + 0.1, row: rowNo - 1 + 0.1 }, ext: { width: 90, height: 56 } });
  } else if (file) {
    ws.getCell(rowNo, col).value = 'Click "Open file" to view';
  }
  ws.getRow(rowNo).height = 48;
}
const link = (cell, url) => { cell.value = { text: 'Open file', hyperlink: url }; cell.font = LINK_FONT; };
function applyToWorkbook(wb, e) {
  const P = ensureSheet(wb, 'Packages'), R = ensureSheet(wb, 'Requests'), N = ensureSheet(wb, 'Numbers'), F = ensureSheet(wb, 'Files');
  if (e.t === 'pkg' && !findRow(P, colIdx(P, 'packageId'), e.row.packageId)) newRow(P, 'Packages', e.row);
  if (e.t === 'req') {
    if (!findRow(R, colIdx(R, 'requestId'), e.row.requestId)) {
      const n = newRow(R, 'Requests', e.row);
      if (e.file) { link(R.getCell(n, colIdx(R, 'logoLink')), e.file.url); preview(wb, R, n, colIdx(R, 'preview'), e.file); }
    }
    for (const num of e.numbers || []) if (!findRow(N, colIdx(N, 'numberId'), num.numberId)) newRow(N, 'Numbers', num);
    if (e.file && !findRow(F, colIdx(F, 'fileId'), e.file.fileId)) {
      const n = newRow(F, 'Files', e.file);
      link(F.getCell(n, colIdx(F, 'url')), e.file.url); preview(wb, F, n, colIdx(F, 'preview'), e.file);
    }
  }
  if (e.t === 'status') {
    const n = findRow(R, colIdx(R, 'requestId'), e.requestId); if (n) R.getCell(n, colIdx(R, 'status')).value = e.status;
    const rc = colIdx(N, 'requestId'), sc = colIdx(N, 'status');
    for (let r = 2; r <= N.rowCount; r++) if (String(plain(N.getRow(r).getCell(rc).value)) === String(e.requestId)) N.getRow(r).getCell(sc).value = e.status;
  }
}
async function writeWorkbook(entries) {
  for (let attempt = 0; attempt < 3; attempt++) {
    const wb = new ExcelJS.Workbook(); let mtime = null;
    if (fs.existsSync(DB_PATH)) {
      mtime = fs.statSync(DB_PATH).mtimeMs;
      await wb.xlsx.readFile(DB_PATH);                       // unreadable file => throws => we never overwrite it
      const bak = DB_PATH.replace(/(\.xlsx)?$/i, '.backup-before-sync.xlsx');
      if (USER_DB && !fs.existsSync(bak)) fs.copyFileSync(DB_PATH, bak);   // your own workbook: original kept once
    }
    for (const e of entries) applyToWorkbook(wb, e);
    wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };   // your formulas recalculate on open
    const tmp = DB_PATH + '.tmp';
    await fs.promises.writeFile(tmp, Buffer.from(await wb.xlsx.writeBuffer()));
    if (mtime !== null && fs.statSync(DB_PATH).mtimeMs !== mtime) { await fs.promises.unlink(tmp); continue; }   // you saved meanwhile: redo on your version
    await fs.promises.rename(tmp, DB_PATH);                  // atomic; fails while Excel has the file open -> retried
    return;
  }
  throw new Error('the file kept changing, will try again');
}

// ---------- flushing the journal into the workbook ----------
let flushTimer = null, flushInfo = { lastFlush: null, error: null };
function scheduleFlush(ms = 150) { clearTimeout(flushTimer); flushTimer = setTimeout(flush, ms); }
async function flush() {
  try {
    await withDbLock(async () => {
      const entries = readJournal(); if (!entries.length) return;
      await writeWorkbook(entries);
      fs.writeFileSync(JOURNAL, '');                         // applied; entries are idempotent so a crash here is harmless
      flushInfo = { lastFlush: new Date().toISOString(), error: null };
    });
  } catch (e) {
    const locked = ['EBUSY', 'EPERM', 'EACCES'].includes(e.code);
    flushInfo.error = locked ? 'The Excel file is open in Excel (or locked).' : /zip|central directory/i.test(e.message) ? 'The database file is not a valid .xlsx workbook; it was left untouched.' : e.message;
    scheduleFlush(locked ? 3000 : 15000);
  }
}
function startBackground() {
  scheduleFlush(300);
  setInterval(() => { try { if (fs.existsSync(JOURNAL) && fs.statSync(JOURNAL).size > 0) scheduleFlush(100); } catch (e) {} }, 3000);
}
function getStatus() {
  return { database: DB_PATH, pending: readJournal().length, lastFlush: flushInfo.lastFlush, error: flushInfo.error, readError };
}

module.exports = { cfg, SHARED_DIR, STORAGE_DIR, DB_PATH, newId, enqueue, readAll, getStatus, startBackground, flush, saveInMemoryFile, getInMemoryFile };
