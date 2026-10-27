'use strict';
// MongoDB Atlas is used for Vercel or when MONGODB_URI is configured.
// Local and serverless runs use a dual-layer persistence engine:
// 1. An atomic, durable JSON store ('shared/database.json') for instant ACID-like persistence.
// 2. The shared Excel workbook ('shared/CallerID-Database.xlsx') with cross-process locking & automatic sync.
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const mongoStore = require('./mongo-store');

const IS_VERCEL = !!process.env.VERCEL;
const USE_MONGO = mongoStore.configured();
let mongoInitialization;
let cfg = {};
try { cfg = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8')); } catch (e) {}

const BASE_DIR = IS_VERCEL ? '/tmp' : __dirname;
const SHARED_DIR = path.resolve(BASE_DIR, process.env.SHARED_DIR || cfg.sharedDir || 'shared');
const USER_DB = !!(process.env.DATABASE_PATH || cfg.databasePath);
const DB_PATH = path.resolve(BASE_DIR, process.env.DATABASE_PATH || cfg.databasePath || path.join(SHARED_DIR, 'CallerID-Database.xlsx'));
const JSON_DB = path.join(SHARED_DIR, 'database.json');
const PREFIX = cfg.sheetPrefix || '';
const STORAGE_DIR = path.join(SHARED_DIR, 'storage', 'logos');
const JOURNAL = path.join(SHARED_DIR, '.pending-changes.jsonl');
const LOCK = path.join(SHARED_DIR, '.db.lock');

const memoryFiles = new Map();

function saveInMemoryFile(fileId, buffer, mimeType, originalName) {
  if (!fileId || !buffer) return;
  const mime = mimeType || 'image/png';
  const name = originalName || 'logo.png';
  memoryFiles.set(fileId, {
    buffer,
    mimeType: mime,
    originalName: name,
    dataUrl: `data:${mime};base64,${buffer.toString('base64')}`
  });
  // Also write to STORAGE_DIR immediately so all processes (ports 4001, 4002, 4003) can access the file
  try {
    fs.mkdirSync(STORAGE_DIR, { recursive: true });
    let ext = 'png';
    if (mime === 'image/jpeg') ext = 'jpg';
    else if (mime === 'image/svg+xml') ext = 'svg';
    else if (mime === 'image/webp') ext = 'webp';
    else if (name.includes('.')) ext = name.split('.').pop().toLowerCase().replace(/[^a-z0-9]/g, '') || 'png';
    const filePath = path.join(STORAGE_DIR, `${fileId}.${ext}`);
    if (!fs.existsSync(filePath)) {
      fs.writeFileSync(filePath, buffer);
    }
  } catch (e) {
    console.warn('Could not persist in-memory file to disk:', e.message);
  }
}

function getInMemoryFile(fileId) {
  return memoryFiles.get(fileId) || null;
}

const COLS = {
  Packages: ['packageId', 'name', 'capacity', 'duration', 'pricePHP', 'brandedCalls', 'brand', 'callReason', 'tagline', 'logoIncluded', 'status', 'createdAt'],
  Requests: ['requestId', 'packageId', 'brandName', 'callReason', 'tagline', 'status', 'remarks', 'logoFileId', 'createdAt', 'departments', 'numbersCount', 'monthlyFeePHP', 'vatPHP', 'totalPHP', 'logoLink', 'preview'],
  Numbers: ['numberId', 'requestId', 'brandName', 'department', 'phoneE164', 'phoneDisplay', 'numberType', 'status', 'remarks', 'decidedAt', 'createdAt'],
  Files: ['fileId', 'requestId', 'originalName', 'storedName', 'storagePath', 'url', 'mimeType', 'sizeBytes', 'sha256', 'uploadedBy', 'uploadedAt', 'preview'],
  DepartmentLines: ['lineId', 'department', 'phoneE164', 'phoneDisplay', 'numberType', 'brandName', 'updatedAt']
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

// In-memory state and cache control
const EMPTY = () => ({ packages: [], requests: [], numbers: [], files: [], departmentLines: [] });
let memoryState = null;
let lastJsonMtime = 0;
let readError = null;

function saveStateToJson(state) {
  if (!state) return;
  try {
    fs.mkdirSync(path.dirname(JSON_DB), { recursive: true });
    const payload = JSON.stringify(state, null, 2);
    const tmp = `${JSON_DB}.${process.pid}.${Date.now()}.tmp`;
    fs.writeFileSync(tmp, payload, 'utf8');
    try {
      fs.renameSync(tmp, JSON_DB);
    } catch (e) {
      // Windows file replacement fallback
      fs.copyFileSync(tmp, JSON_DB);
      try { fs.unlinkSync(tmp); } catch (x) {}
    }
    try {
      lastJsonMtime = fs.statSync(JSON_DB).mtimeMs;
    } catch (e) {}
  } catch (err) {
    console.error('Failed to write database.json:', err);
  }
}

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

async function loadWorkbookRaw() {
  if (!fs.existsSync(DB_PATH)) {
    const seed = path.join(__dirname, 'shared', 'CallerID-Database.xlsx');
    if (fs.existsSync(seed)) {
      try {
        fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
        fs.copyFileSync(seed, DB_PATH);
      } catch (e) {}
    }
  }
  if (!fs.existsSync(DB_PATH)) return EMPTY();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(DB_PATH);
  return {
    packages: rowsOf(wb.getWorksheet(PREFIX + 'Packages'), 'Packages'),
    requests: rowsOf(wb.getWorksheet(PREFIX + 'Requests'), 'Requests'),
    numbers: rowsOf(wb.getWorksheet(PREFIX + 'Numbers'), 'Numbers'),
    files: rowsOf(wb.getWorksheet(PREFIX + 'Files'), 'Files'),
    departmentLines: rowsOf(wb.getWorksheet(PREFIX + 'DepartmentLines'), 'DepartmentLines')
  };
}

function initStorage() {
  try {
    fs.mkdirSync(STORAGE_DIR, { recursive: true });
    fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
    fs.mkdirSync(SHARED_DIR, { recursive: true });

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

    // On Vercel or isolated temp directories, seed database.json and CallerID-Database.xlsx from bundled shared folder
    if (IS_VERCEL) {
      const seedJson = path.join(__dirname, 'shared', 'database.json');
      if (fs.existsSync(seedJson) && !fs.existsSync(JSON_DB)) {
        try { fs.copyFileSync(seedJson, JSON_DB); } catch (e) {}
      }
      const seedDb = path.join(__dirname, 'shared', 'CallerID-Database.xlsx');
      if (fs.existsSync(seedDb) && !fs.existsSync(DB_PATH)) {
        try { fs.copyFileSync(seedDb, DB_PATH); } catch (e) {}
      }
      const seedStorage = path.join(__dirname, 'shared', 'storage', 'logos');
      if (fs.existsSync(seedStorage)) {
        try {
          const files = fs.readdirSync(seedStorage);
          files.forEach(f => {
            const dest = path.join(STORAGE_DIR, f);
            if (!fs.existsSync(dest)) {
              try { fs.copyFileSync(path.join(seedStorage, f), dest); } catch (e) {}
            }
          });
        } catch (e) {}
      }
    }

    // Initialize or load state from JSON_DB
    if (fs.existsSync(JSON_DB)) {
      try {
        const raw = fs.readFileSync(JSON_DB, 'utf8');
        memoryState = JSON.parse(raw);
        lastJsonMtime = fs.statSync(JSON_DB).mtimeMs;
      } catch (e) {
        console.warn('database.json parse error, will recover from workbook:', e.message);
        memoryState = null;
      }
    }

    if (!memoryState || !memoryState.packages) {
      if (fs.existsSync(DB_PATH)) {
        memoryState = EMPTY();
      } else {
        const seedDb = path.join(__dirname, 'shared', 'CallerID-Database.xlsx');
        if (fs.existsSync(seedDb) && !fs.existsSync(DB_PATH)) {
          try { fs.copyFileSync(seedDb, DB_PATH); } catch (e) {}
        }
        memoryState = EMPTY();
      }
    }
  } catch (e) {
    console.error('Storage initialization failed:', e);
  }
}
initStorage();

// ---------- cross-process lock (multiple websites, one workbook) ----------
async function acquire() {
  const t0 = Date.now();
  for (;;) {
    try {
      const fd = fs.openSync(LOCK, 'wx');
      fs.writeSync(fd, String(process.pid));
      fs.closeSync(fd);
      return;
    } catch (e) {
      if (e.code !== 'EEXIST') throw e;
      try {
        if (Date.now() - fs.statSync(LOCK).mtimeMs > 30000) fs.unlinkSync(LOCK);
      } catch (x) {} // stale lock from a crash
      if (Date.now() - t0 > 20000) throw new Error('database is busy, try again');
      await sleep(25 + Math.random() * 40);
    }
  }
}

async function withDbLock(fn) {
  await acquire();
  try {
    return await fn();
  } finally {
    try { fs.unlinkSync(LOCK); } catch (e) {}
  }
}

// ---------- journal ----------
function readJournal() {
  if (!fs.existsSync(JOURNAL)) return [];
  return fs.readFileSync(JOURNAL, 'utf8').split('\n').filter(Boolean).map((line, index) => {
    try { return JSON.parse(line); }
    catch (e) { console.warn(`Skipping invalid JSON in journal at line ${index + 1}: ${e.message}`); return null; }
  }).filter(Boolean);
}

function deriveRequestStatus(numbers) {
  if (!numbers || !numbers.length) return 'pending';
  const statuses = numbers.map(n => String(n.status || 'pending').toLowerCase().trim());
  const hasApproved = statuses.includes('approved');
  const hasRejected = statuses.includes('rejected');
  const hasPending = statuses.includes('pending');
  if (hasApproved && !hasRejected && !hasPending) return 'approved';
  if (hasRejected && !hasApproved && !hasPending) return 'rejected';
  if (!hasApproved && !hasRejected && hasPending) return 'pending';
  return 'partially approved';
}

function applyToState(s, e) {
  if (!s || !e) return;
  if (!s.packages) s.packages = [];
  if (!s.requests) s.requests = [];
  if (!s.numbers) s.numbers = [];
  if (!s.files) s.files = [];
  if (!s.departmentLines) s.departmentLines = [];

  if (e.t === 'pkg') {
    const row = {
      ...e.row,
      duration: e.row.duration || '30 Days',
      brandedCalls: Number(e.row.brandedCalls) || (Number(e.row.capacity || 5) * 10000) || 50000
    };
    const existingIdx = s.packages.findIndex(p => p.packageId === row.packageId);
    if (existingIdx >= 0) {
      s.packages[existingIdx] = { ...s.packages[existingIdx], ...row };
    } else {
      s.packages.push({ ...row });
    }
  }
  if (e.t === 'pkg-update') {
    const p = s.packages.find(x => x.packageId === e.row.packageId);
    if (p) Object.assign(p, e.row);
  }
  if (e.t === 'pkg-remove') {
    const p = s.packages.find(x => x.packageId === e.packageId);
    if (p) {
      p.status = 'removed';
    } else {
      // Record removal placeholder so it stays removed
      s.packages.push({ packageId: e.packageId, status: 'removed' });
    }
  }
  if (e.t === 'req') {
    const callReason = e.row.callReason || e.row.tagline || '';
    const row = {
      ...e.row,
      callReason,
      tagline: callReason,
      remarks: e.row.remarks || ''
    };
    const existingReq = s.requests.find(r => r.requestId === row.requestId);
    if (!existingReq) {
      s.requests.push({ ...row });
    } else {
      Object.assign(existingReq, row);
    }
    if (e.file && !s.files.some(f => f.fileId === e.file.fileId)) s.files.push({ ...e.file });
    for (const n of e.numbers || []) {
      const existingNum = s.numbers.find(x => x.numberId === n.numberId);
      if (!existingNum) {
        s.numbers.push({ ...n });
      } else {
        Object.assign(existingNum, n);
      }
    }
  }
  if (e.t === 'number-decisions') {
    const r = s.requests.find(x => x.requestId === e.requestId);
    const decisionTime = e.at || new Date().toISOString();
    for (const d of e.decisions || []) {
      const n = s.numbers.find(x => (d.numberId && x.numberId === d.numberId) || (x.requestId === e.requestId && x.phoneE164 === d.phoneE164));
      if (n) {
        n.status = d.status;
        n.remarks = d.remarks || '';
        n.decidedAt = decisionTime;
      }
    }
    if (r) {
      const reqNumbers = s.numbers.filter(n => n.requestId === e.requestId);
      r.status = deriveRequestStatus(reqNumbers);
      if (e.remarks !== undefined && e.remarks !== '') r.remarks = e.remarks;
    }
  }
  if (e.t === 'resubmit-numbers') {
    const r = s.requests.find(x => x.requestId === e.requestId);
    for (const item of e.numbers || []) {
      const n = s.numbers.find(x => (item.numberId && x.numberId === item.numberId) || (x.requestId === e.requestId && x.phoneE164 === item.phoneE164));
      if (n) {
        if (item.department) n.department = item.department;
        if (item.phoneE164) n.phoneE164 = item.phoneE164;
        if (item.phoneDisplay) n.phoneDisplay = item.phoneDisplay;
        if (item.numberType) n.numberType = item.numberType;
        n.status = 'pending';
        n.remarks = '';
        n.decidedAt = null;
      }
    }
    if (r) {
      const reqNumbers = s.numbers.filter(n => n.requestId === e.requestId);
      r.status = deriveRequestStatus(reqNumbers);
    }
  }
  if (e.t === 'status') {
    const r = s.requests.find(x => x.requestId === e.requestId);
    const decisionTime = e.at || new Date().toISOString();
    if (e.decisions && Array.isArray(e.decisions)) {
      for (const d of e.decisions) {
        const n = s.numbers.find(x => (d.numberId && x.numberId === d.numberId) || (x.requestId === e.requestId && x.phoneE164 === d.phoneE164));
        if (n) {
          n.status = d.status;
          n.remarks = d.remarks || '';
          n.decidedAt = decisionTime;
        }
      }
      if (r) {
        const reqNumbers = s.numbers.filter(n => n.requestId === e.requestId);
        r.status = deriveRequestStatus(reqNumbers);
        if (e.remarks !== undefined) r.remarks = e.remarks;
      }
    } else {
      if (r) {
        r.status = e.status;
        if (e.remarks !== undefined) r.remarks = e.remarks;
      }
      s.numbers.forEach(n => {
        if (n.requestId === e.requestId) {
          n.status = e.status;
          if (e.status === 'rejected' && e.remarks) n.remarks = e.remarks;
          n.decidedAt = decisionTime;
        }
      });
    }
  }
  if (e.t === 'dept-line') {
    const idx = s.departmentLines.findIndex(l => l.lineId === e.row.lineId);
    if (idx >= 0) s.departmentLines[idx] = { ...s.departmentLines[idx], ...e.row };
    else s.departmentLines.push({ ...e.row });
  }
  if (e.t === 'dept-line-remove') {
    s.departmentLines = s.departmentLines.filter(l => l.lineId !== e.lineId);
  }
}

async function ensureMemoryLoaded() {
  if (memoryState && memoryState.packages && memoryState.packages.length > 0) {
    // Check if JSON_DB was touched by another process
    if (fs.existsSync(JSON_DB)) {
      try {
        const st = fs.statSync(JSON_DB);
        if (st.mtimeMs > lastJsonMtime) {
          const raw = fs.readFileSync(JSON_DB, 'utf8');
          memoryState = JSON.parse(raw);
          lastJsonMtime = st.mtimeMs;
        }
      } catch (e) {}
    }
    return memoryState;
  }

  // Load from JSON_DB if available
  if (fs.existsSync(JSON_DB)) {
    try {
      const raw = fs.readFileSync(JSON_DB, 'utf8');
      memoryState = JSON.parse(raw);
      lastJsonMtime = fs.statSync(JSON_DB).mtimeMs;
      if (memoryState && memoryState.packages) return memoryState;
    } catch (e) {}
  }

  // Load from Excel workbook
  try {
    memoryState = await loadWorkbookRaw();
    readError = null;
  } catch (e) {
    readError = e.message;
    memoryState = EMPTY();
  }

  // Apply any pending journal entries
  const pending = readJournal();
  if (pending.length > 0) {
    for (const e of pending) applyToState(memoryState, e);
  }

  // Persist the combined authoritative state to JSON_DB
  saveStateToJson(memoryState);
  return memoryState;
}

async function enqueue(entry) {
  if (USE_MONGO) {
    await ensureMongo();
    await mongoStore.enqueue(entry);
    return;
  }

  // 1. Immediately apply to authoritative in-memory state
  const state = await ensureMemoryLoaded();
  applyToState(state, entry);

  // 2. Immediately write to durable JSON database
  saveStateToJson(state);

  // 3. Append to journal file with lock
  await withDbLock(async () => {
    const fd = fs.openSync(JOURNAL, 'a');
    try {
      fs.writeSync(fd, JSON.stringify(entry) + '\n');
      fs.fsyncSync(fd);
    } finally {
      fs.closeSync(fd);
    }
  });

  // 4. Flush to Excel workbook
  if (IS_VERCEL) {
    try { await flush(true); } catch (e) {}
  } else {
    scheduleFlush(50);
  }
}

async function ensureMongo() {
  if (!mongoInitialization) {
    mongoInitialization = (async () => {
      const seed = await ensureMemoryLoaded();
      const fileBuffers = new Map();
      for (const file of seed.files) {
        const fullPath = path.join(SHARED_DIR, file.storagePath || '');
        if (file.fileId && file.storagePath && fs.existsSync(fullPath)) {
          fileBuffers.set(file.fileId, fs.readFileSync(fullPath));
        }
      }
      await mongoStore.initialize(seed, fileBuffers);
    })().catch(error => {
      mongoInitialization = null;
      throw error;
    });
  }
  return mongoInitialization;
}

async function readAll() {
  if (USE_MONGO) {
    await ensureMongo();
    return mongoStore.readAll();
  }

  const rawState = await ensureMemoryLoaded();
  const s = JSON.parse(JSON.stringify(rawState));

  // Also apply any changes that might be pending in the journal from other processes
  for (const e of readJournal()) applyToState(s, e);

  // Ensure file URLs and dataUrls are populated
  s.files = (s.files || []).map(f => {
    const mem = memoryFiles.get(f.fileId);
    if (mem && mem.dataUrl) {
      return { ...f, dataUrl: mem.dataUrl };
    }
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

async function getStoredFile(fileId) {
  if (!USE_MONGO) return null;
  await ensureMongo();
  const file = await mongoStore.getFile(fileId);
  if (file) saveInMemoryFile(fileId, file.buffer, file.mimeType, file.originalName);
  return file;
}

// ---------- writing into the workbook ----------
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

function colIdx(ws, name) {
  let found = 0, last = 0;
  ws.getRow(1).eachCell({ includeEmpty: false }, (c, n) => {
    last = Math.max(last, n);
    if (String(plain(c.value)).trim() === name) found = n;
  });
  if (found) return found;
  const n = last + 1, c = ws.getRow(1).getCell(n);
  c.value = name;
  c.font = { bold: true };
  ws.getColumn(n).width = Math.max(14, name.length + 4);
  return n;
}

function findRow(ws, col, value) {
  for (let r = 2; r <= ws.rowCount; r++) {
    if (String(plain(ws.getRow(r).getCell(col).value)) === String(value)) return r;
  }
  return 0;
}

function newRow(ws, name, obj) {
  const rowNo = (ws.lastRow ? ws.lastRow.number : 0) + 1, row = ws.getRow(rowNo);
  for (const k of COLS[name]) {
    if (!DISPLAY_ONLY.has(k) && obj[k] !== undefined) {
      const c = row.getCell(colIdx(ws, k));
      c.value = obj[k];
      if (MONEY.has(k)) c.numFmt = PESO_FMT;
    }
  }
  return rowNo;
}

function imageIndex(wb, buffer, extension) {
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
  if (e.t === 'pkg') {
    const existing = findRow(P, colIdx(P, 'packageId'), e.row.packageId);
    if (!existing) {
      newRow(P, 'Packages', e.row);
    } else {
      for (const k of COLS.Packages) {
        if (k !== 'packageId' && e.row[k] !== undefined) {
          const cell = P.getRow(existing).getCell(colIdx(P, k));
          cell.value = e.row[k];
          if (MONEY.has(k)) cell.numFmt = PESO_FMT;
        }
      }
    }
  }
  if (e.t === 'pkg-update') {
    const rowNo = findRow(P, colIdx(P, 'packageId'), e.row.packageId);
    if (rowNo) {
      for (const k of COLS.Packages) {
        if (k !== 'packageId' && e.row[k] !== undefined) {
          const cell = P.getRow(rowNo).getCell(colIdx(P, k));
          cell.value = e.row[k];
          if (MONEY.has(k)) cell.numFmt = PESO_FMT;
        }
      }
    }
  }
  if (e.t === 'pkg-remove') {
    const rowNo = findRow(P, colIdx(P, 'packageId'), e.packageId);
    if (rowNo) {
      P.getRow(rowNo).getCell(colIdx(P, 'status')).value = 'removed';
    }
  }
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
  if (e.t === 'number-decisions' || (e.t === 'status' && e.decisions)) {
    const decisions = e.decisions || [];
    const decisionTime = e.at || new Date().toISOString();
    const nc = colIdx(N, 'numberId'), rc = colIdx(N, 'requestId'), pc = colIdx(N, 'phoneE164'), sc = colIdx(N, 'status'), rmc = colIdx(N, 'remarks'), dc = colIdx(N, 'decidedAt');
    for (const d of decisions) {
      for (let r = 2; r <= N.rowCount; r++) {
        const row = N.getRow(r);
        const matchId = d.numberId && String(plain(row.getCell(nc).value)) === String(d.numberId);
        const matchReqAndPhone = String(plain(row.getCell(rc).value)) === String(e.requestId) && String(plain(row.getCell(pc).value)) === String(d.phoneE164);
        if (matchId || matchReqAndPhone) {
          row.getCell(sc).value = d.status;
          row.getCell(rmc).value = d.remarks || '';
          row.getCell(dc).value = decisionTime;
        }
      }
    }
    const n = findRow(R, colIdx(R, 'requestId'), e.requestId);
    if (n) {
      if (e.requestStatus) {
        R.getCell(n, colIdx(R, 'status')).value = e.requestStatus;
      }
      if (e.remarks !== undefined) {
        R.getCell(n, colIdx(R, 'remarks')).value = e.remarks;
      }
    }
  } else if (e.t === 'resubmit-numbers') {
    const nc = colIdx(N, 'numberId'), rc = colIdx(N, 'requestId'), pc = colIdx(N, 'phoneE164'), sc = colIdx(N, 'status'), rmc = colIdx(N, 'remarks'), dc = colIdx(N, 'decidedAt'), deptc = colIdx(N, 'department');
    for (const item of e.numbers || []) {
      for (let r = 2; r <= N.rowCount; r++) {
        const row = N.getRow(r);
        const matchId = item.numberId && String(plain(row.getCell(nc).value)) === String(item.numberId);
        const matchReqAndPhone = String(plain(row.getCell(rc).value)) === String(e.requestId) && String(plain(row.getCell(pc).value)) === String(item.phoneE164);
        if (matchId || matchReqAndPhone) {
          row.getCell(sc).value = 'pending';
          row.getCell(rmc).value = '';
          row.getCell(dc).value = '';
          if (item.department) row.getCell(deptc).value = item.department;
          if (item.phoneDisplay) row.getCell(colIdx(N, 'phoneDisplay')).value = item.phoneDisplay;
        }
      }
    }
    const n = findRow(R, colIdx(R, 'requestId'), e.requestId);
    if (n && e.requestStatus) {
      R.getCell(n, colIdx(R, 'status')).value = e.requestStatus;
    }
  } else if (e.t === 'status') {
    const n = findRow(R, colIdx(R, 'requestId'), e.requestId);
    if (n) {
      R.getCell(n, colIdx(R, 'status')).value = e.status;
      if (e.remarks !== undefined) {
        R.getCell(n, colIdx(R, 'remarks')).value = e.remarks;
      }
    }
    const rc = colIdx(N, 'requestId'), sc = colIdx(N, 'status'), rmc = colIdx(N, 'remarks'), dc = colIdx(N, 'decidedAt');
    const decisionTime = e.at || new Date().toISOString();
    for (let r = 2; r <= N.rowCount; r++) {
      if (String(plain(N.getRow(r).getCell(rc).value)) === String(e.requestId)) {
        N.getRow(r).getCell(sc).value = e.status;
        if (e.status === 'rejected' && e.remarks) N.getRow(r).getCell(rmc).value = e.remarks;
        N.getRow(r).getCell(dc).value = decisionTime;
      }
    }
  }
}

async function writeWorkbook(entries) {
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      const wb = new ExcelJS.Workbook();
      let mtime = null;
      if (fs.existsSync(DB_PATH)) {
        mtime = fs.statSync(DB_PATH).mtimeMs;
        await wb.xlsx.readFile(DB_PATH);
        const bak = DB_PATH.replace(/(\.xlsx)?$/i, '.backup-before-sync.xlsx');
        if (USER_DB && !fs.existsSync(bak)) {
          try { fs.copyFileSync(DB_PATH, bak); } catch (e) {}
        }
      }
      for (const e of entries) applyToWorkbook(wb, e);
      wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };

      const tmp = `${DB_PATH}.${process.pid}.${Date.now()}.tmp`;
      const buf = Buffer.from(await wb.xlsx.writeBuffer());
      await fs.promises.writeFile(tmp, buf);

      try {
        await fs.promises.rename(tmp, DB_PATH);
      } catch (renameErr) {
        // Windows atomic replacement fallback
        try {
          fs.copyFileSync(tmp, DB_PATH);
          await fs.promises.unlink(tmp).catch(() => {});
        } catch (copyErr) {
          await fs.promises.unlink(tmp).catch(() => {});
          throw renameErr;
        }
      }
      return;
    } catch (err) {
      if (attempt === 4) throw err;
      await sleep(100 + attempt * 150);
    }
  }
}

// ---------- flushing the journal into the workbook ----------
let flushTimer = null;
let flushInfo = { lastFlush: null, error: null };

function scheduleFlush(ms = 100) {
  clearTimeout(flushTimer);
  flushTimer = setTimeout(flush, ms);
}

async function flush(required = false) {
  if (USE_MONGO) return;
  try {
    await withDbLock(async () => {
      const entries = readJournal();
      if (!entries.length) return;
      await writeWorkbook(entries);
      try { fs.writeFileSync(JOURNAL, ''); } catch (e) {}
      flushInfo = { lastFlush: new Date().toISOString(), error: null };
    });
  } catch (e) {
    const locked = ['EBUSY', 'EPERM', 'EACCES'].includes(e.code);
    flushInfo.error = locked
      ? 'The Excel file is currently open in Excel. Changes are saved in database.json and will sync to Excel once closed.'
      : /zip|central directory/i.test(e.message)
      ? 'The database file is not a valid .xlsx workbook; it was left untouched.'
      : e.message;
    console.warn('Workbook sync notice:', flushInfo.error);
    scheduleFlush(locked ? 2000 : 5000);
    if (required) throw e;
  }
}

function startBackground() {
  if (USE_MONGO) return;
  scheduleFlush(200);
  setInterval(() => {
    try {
      if (fs.existsSync(JOURNAL) && fs.statSync(JOURNAL).size > 0) scheduleFlush(50);
    } catch (e) {}
  }, 2000);
}

function getStatus() {
  if (USE_MONGO) {
    return mongoStore.status();
  }
  return {
    database: DB_PATH,
    jsonStore: JSON_DB,
    pending: readJournal().length,
    lastFlush: flushInfo.lastFlush,
    error: flushInfo.error,
    readError
  };
}

module.exports = {
  cfg,
  COLS,
  SHARED_DIR,
  STORAGE_DIR,
  DB_PATH,
  JSON_DB,
  newId,
  enqueue,
  readAll,
  readWorkbookSnapshot: ensureMemoryLoaded,
  getStatus,
  startBackground,
  flush,
  saveInMemoryFile,
  getInMemoryFile,
  getStoredFile,
  saveStateToJson,
  deriveRequestStatus
};
