'use strict';
// One code base, two websites:  node server.js operator   (port 4001)
//                               node server.js enterprise (port 4002)
// Both read and write the SAME Excel workbook through db.js.
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const db = require('./db');
const PH = require('./public/shared/ph.js');

const RAW_ROLE = (process.argv[2] || process.env.ROLE || '').toLowerCase();
const IS_UNIFIED = !RAW_ROLE || RAW_ROLE === 'all' || RAW_ROLE === 'unified' || !!process.env.VERCEL;
const ROLE = IS_UNIFIED ? 'unified' : RAW_ROLE;
if (!IS_UNIFIED && !['operator', 'enterprise', 'customer'].includes(ROLE)) {
  console.error('Usage: node server.js operator|enterprise|customer (or run node server.js for unified mode)');
  process.exit(1);
}
const cfg = db.cfg;
const PORTS = { operator: Number(cfg.operatorPort) || 4001, enterprise: Number(cfg.enterprisePort) || 4002, customer: Number(cfg.customerPort) || 4003 };
const PORT = Number(process.env.PORT) || (IS_UNIFIED ? 3000 : PORTS[ROLE]);
const OPERATOR_URL = (cfg.operatorUrl || `http://localhost:${PORTS.operator}`).replace(/\/$/, '');
const ENTERPRISE_URL = (cfg.enterpriseUrl || `http://localhost:${PORTS.enterprise}`).replace(/\/$/, '');
const CUSTOMER_URL = (cfg.customerUrl || `http://localhost:${PORTS.customer}`).replace(/\/$/, '');
const MAX_BYTES = (Number(process.env.MAX_MB) || Number(cfg.maxUploadMB) || 20) * 1024 * 1024;

// ---------- file types: detected from file CONTENT, not from the name ----------
const ascii = (b, from, to) => b.slice(from, to).toString('latin1');
const TYPES = [
  { ext: 'png',  mime: 'image/png',       test: b => b.slice(0, 8).equals(Buffer.from([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a])) },
  { ext: 'jpg',  mime: 'image/jpeg',      test: b => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff },
  { ext: 'gif',  mime: 'image/gif',       test: b => ['GIF87a','GIF89a'].includes(ascii(b, 0, 6)) },
  { ext: 'webp', mime: 'image/webp',      test: b => ascii(b, 0, 4) === 'RIFF' && ascii(b, 8, 12) === 'WEBP' },
  { ext: 'bmp',  mime: 'image/bmp',       test: b => ascii(b, 0, 2) === 'BM' },
  { ext: 'tif',  mime: 'image/tiff',      test: b => b.slice(0, 4).equals(Buffer.from([0x49,0x49,0x2a,0x00])) || b.slice(0, 4).equals(Buffer.from([0x4d,0x4d,0x00,0x2a])) },
  { ext: 'ico',  mime: 'image/x-icon',    test: b => b.slice(0, 4).equals(Buffer.from([0x00,0x00,0x01,0x00])) },
  { ext: 'avif', mime: 'image/avif',      test: b => ascii(b, 4, 8) === 'ftyp' && ['avif','avis'].includes(ascii(b, 8, 12)) },
  { ext: 'heic', mime: 'image/heic',      test: b => ascii(b, 4, 8) === 'ftyp' && ['heic','heix','hevc','hevx','mif1','msf1'].includes(ascii(b, 8, 12)) },
  { ext: 'jp2',  mime: 'image/jp2',       test: b => b.slice(0, 12).equals(Buffer.from([0,0,0,0x0c,0x6a,0x50,0x20,0x20,0x0d,0x0a,0x87,0x0a])) },
  { ext: 'psd',  mime: 'image/vnd.adobe.photoshop', test: b => ascii(b, 0, 4) === '8BPS' },
  { ext: 'svg',  mime: 'image/svg+xml',   test: b => /<svg[\s>]/i.test(b.slice(0, 2048).toString('utf8')) },
  { ext: 'pdf',  mime: 'application/pdf', test: b => ascii(b, 0, 5) === '%PDF-' },
];
const PREVIEWABLE = new Set(['image/png','image/jpeg','image/gif','image/webp','image/bmp','image/x-icon','image/avif','image/svg+xml','application/pdf']);
const BLOCKED_EXT = new Set(['exe','dll','bat','cmd','com','msi','scr','sh','js','mjs','jar','vbs','ps1','html','htm','php']);
function classify(buf, originalName) {
  const known = TYPES.find(t => t.test(buf));
  if (known) return { ext: known.ext, mime: known.mime };
  const raw = path.extname(originalName || '').slice(1).toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 8);
  if (BLOCKED_EXT.has(raw)) return null;
  return { ext: raw || 'bin', mime: 'application/octet-stream' };
}
const truthy = v => v === true || v === 'true' || v === 1 || v === 'TRUE';

const app = express();
app.use(express.json());
app.use('/shared', express.static(path.join(__dirname, 'public', 'shared')));

if (IS_UNIFIED) {
  app.use('/operator', express.static(path.join(__dirname, 'public', 'operator')));
  app.use('/enterprise', express.static(path.join(__dirname, 'public', 'enterprise')));
  app.use('/customer', express.static(path.join(__dirname, 'public', 'customer')));
  app.use('/sheet.html', express.static(path.join(__dirname, 'public', 'operator', 'sheet.html')));
  app.use(express.static(path.join(__dirname, 'public')));
} else {
  app.use(express.static(path.join(__dirname, 'public', ROLE)));
}

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: MAX_BYTES, files: 1 } });
const only = (...roles) => (req, res, next) => {
  if (IS_UNIFIED) return next();
  return roles.includes(ROLE) ? next() : res.status(404).json({ error: `Not available on the ${ROLE} site` });
};
const wrap = fn => (req, res, next) => fn(req, res, next).catch(next);

const activeStatus = st => st === 'pending' || st === 'approved';      // rejected registrations free their numbers
function decorate(s) {
  const files = Object.fromEntries(s.files.map(f => [f.fileId, f])), pk = Object.fromEntries(s.packages.map(p => [p.packageId, p]));
  return s.requests.map(r => {
    const f = files[r.logoFileId], p = pk[r.packageId];
    const numbers = s.numbers.filter(n => n.requestId === r.requestId).map(n => ({ numberId: n.numberId, department: n.department, phoneE164: n.phoneE164, phoneDisplay: n.phoneDisplay, numberType: n.numberType, status: n.status }));
    const logoUrl = f ? (f.dataUrl || `/api/files/${f.fileId}`) : null;
    return { ...r, packageName: p ? p.name : '', numbers, fileName: f ? f.originalName : null, mimeType: f ? f.mimeType : null,
      sizeBytes: f ? f.sizeBytes : null, previewable: f ? PREVIEWABLE.has(f.mimeType) && f.mimeType.startsWith('image/') : false, logoUrl };
  });
}
// what a customer's phone would show for a number (the network looks the caller up in the operator's approved registry)
function lookup(s, input) {
  const r = PH.normalize(input);
  if (!r.ok) return { registered: false, reason: 'invalid', message: r.error };
  const reqOf = id => s.requests.find(x => x.requestId === id);
  const hit = s.numbers.find(n => n.phoneE164 === r.e164 && reqOf(n.requestId) && activeStatus(reqOf(n.requestId).status));
  const base = { e164: r.e164, display: r.display, numberType: r.type };
  if (!hit) return { ...base, registered: false, reason: 'unregistered' };
  const q = reqOf(hit.requestId);
  if (q.status !== 'approved') return { ...base, registered: false, reason: 'pending', brandName: q.brandName };
  const f = s.files.find(x => x.fileId === q.logoFileId);
  const logoUrl = f && f.mimeType.startsWith('image/') && PREVIEWABLE.has(f.mimeType) ? (f.dataUrl || `/api/files/${f.fileId}`) : null;
  return { ...base, registered: true, verified: true, brandName: q.brandName, tagline: q.tagline || '', department: hit.department, logoUrl };
}

app.get('/api/info', (req, res) => {
  const host = req.get('host') || `localhost:${PORT}`;
  const proto = (req.headers['x-forwarded-proto'] || req.protocol || 'http').split(',')[0].trim();
  const base = `${proto}://${host}`;
  const opUrl = IS_UNIFIED ? `${base}/operator` : OPERATOR_URL;
  const enUrl = IS_UNIFIED ? `${base}/enterprise` : ENTERPRISE_URL;
  const cuUrl = IS_UNIFIED ? `${base}/customer` : CUSTOMER_URL;
  res.json({ role: ROLE, operatorUrl: opUrl, enterpriseUrl: enUrl, customerUrl: cuUrl, ...db.getStatus() });
});

// packages: operator sees all, enterprise sees only published ones
app.get('/api/packages', wrap(async (req, res) => {
  const s = await db.readAll();
  const isEnterprise = ROLE === 'enterprise' || req.query.role === 'enterprise' || (req.headers.referer && req.headers.referer.includes('/enterprise'));
  res.set('Cache-Control', 'no-store');
  res.json(s.packages.filter(p => p.status !== 'removed' && (!isEnterprise || p.status === 'published')));
}));
app.get('/api/requests', only('operator', 'enterprise'), wrap(async (_req, res) => res.json(decorate(await db.readAll()))));
// registered numbers (the customer phone uses this to offer demo callers) and the caller-ID lookup
app.get('/api/numbers', wrap(async (_req, res) => {
  const s = await db.readAll(); const st = id => (s.requests.find(r => r.requestId === id) || {}).status;
  res.json(s.numbers.filter(n => activeStatus(st(n.requestId))).map(n => ({ phoneE164: n.phoneE164, phoneDisplay: n.phoneDisplay, department: n.department, brandName: n.brandName, status: st(n.requestId) })));
}));
app.get('/api/lookup', wrap(async (req, res) => res.json(lookup(await db.readAll(), req.query.number))));

// OPERATOR: publish a package
app.post('/api/packages', only('operator'), wrap(async (req, res) => {
  const { name, capacity, pricePHP = 0, brand = true, tagline = false, logoIncluded = false } = req.body || {};
  if (typeof name !== 'string' || !name.trim() || !Number.isInteger(capacity) || capacity <= 0) return res.status(400).json({ error: 'name and a positive whole-number capacity are required' });
  if (typeof pricePHP !== 'number' || !isFinite(pricePHP) || pricePHP < 0) return res.status(400).json({ error: 'price must be an amount in Philippine pesos (₱), 0 or more' });
  const row = { packageId: db.newId('PKG'), name: name.trim().slice(0, 80), capacity, pricePHP: PH.round2(pricePHP), brand: !!brand, tagline: !!tagline, logoIncluded: !!logoIncluded, status: 'published', createdAt: new Date().toISOString() };
  await db.enqueue({ t: 'pkg', row });
  res.status(201).json(row);
}));

async function updatePackage(req, res) {
  const s = await db.readAll();
  const current = s.packages.find(p => p.packageId === req.params.id && p.status !== 'removed');
  if (!current) return res.status(404).json({ error: 'package not found' });

  const body = req.body || {};
  const name = body.name === undefined ? current.name : (typeof body.name === 'string' ? body.name.trim() : '');
  const capacity = body.capacity === undefined ? Number(current.capacity) : body.capacity;
  const pricePHP = body.pricePHP === undefined ? Number(current.pricePHP) : body.pricePHP;
  if (typeof name !== 'string' || !name || !Number.isInteger(capacity) || capacity <= 0) {
    return res.status(400).json({ error: 'name and a positive whole-number capacity are required' });
  }
  if (typeof pricePHP !== 'number' || !isFinite(pricePHP) || pricePHP < 0) {
    return res.status(400).json({ error: 'price must be an amount in Philippine pesos (₱), 0 or more' });
  }

  const row = {
    ...current,
    name: name.slice(0, 80),
    capacity,
    pricePHP: PH.round2(pricePHP),
    brand: body.brand === undefined ? truthy(current.brand) : !!body.brand,
    tagline: body.tagline === undefined ? truthy(current.tagline) : !!body.tagline,
    logoIncluded: body.logoIncluded === undefined ? truthy(current.logoIncluded) : !!body.logoIncluded
  };
  await db.enqueue({ t: 'pkg-update', row });
  res.json(row);
}

app.put('/api/packages/:id', only('operator'), wrap(updatePackage));
app.patch('/api/packages/:id', only('operator'), wrap(updatePackage));

// Keep historical references valid while persistently removing plans from both portals.
app.delete('/api/packages/:id', only('operator', 'enterprise'), wrap(async (req, res) => {
  const s = await db.readAll();
  const pkg = s.packages.find(p => p.packageId === req.params.id);
  if (!pkg) return res.status(404).json({ error: 'package not found' });
  if (pkg.status === 'removed') return res.json({ packageId: pkg.packageId, status: 'removed' });
  await db.enqueue({ t: 'pkg-remove', packageId: pkg.packageId });
  res.json({ packageId: pkg.packageId, status: 'removed' });
}));

// OPERATOR: approve / reject
app.patch('/api/requests/:id/status', only('operator'), wrap(async (req, res) => {
  const { status } = req.body || {};
  if (!['approved', 'rejected'].includes(status)) return res.status(400).json({ error: 'status must be approved or rejected' });
  await db.enqueue({ t: 'status', requestId: req.params.id, status, at: new Date().toISOString() });
  res.json({ requestId: req.params.id, status });
}));

// SERVERLESS SYNC: Re-hydrates requests and status overrides across cold/ephemeral Vercel lambda containers
app.post('/api/sync', wrap(async (req, res) => {
  const { requests = [], statusOverrides = {} } = req.body || {};
  const s = await db.readAll();
  const existingReqIds = new Set(s.requests.map(r => r.requestId));
  let addedCount = 0;

  for (const r of requests) {
    if (!r || !r.requestId) continue;
    if (!existingReqIds.has(r.requestId)) {
      const finalStatus = statusOverrides[r.requestId] || r.status || 'pending';
      const row = {
        requestId: r.requestId,
        packageId: r.packageId || '',
        brandName: r.brandName || '',
        tagline: r.tagline || '',
        status: finalStatus,
        logoFileId: r.logoFileId || '',
        createdAt: r.createdAt || new Date().toISOString(),
        departments: r.departments || '',
        numbersCount: r.numbersCount || (r.numbers ? r.numbers.length : 0),
        monthlyFeePHP: r.monthlyFeePHP || 0,
        vatPHP: r.vatPHP || 0,
        totalPHP: r.totalPHP || 0
      };
      const numbers = (r.numbers || []).map(n => ({
        numberId: n.numberId || db.newId('NUM'),
        requestId: r.requestId,
        brandName: r.brandName,
        department: n.department,
        phoneE164: n.phoneE164,
        phoneDisplay: n.phoneDisplay,
        numberType: n.numberType,
        status: finalStatus,
        createdAt: row.createdAt
      }));
      if (r.logoFileId && r.logoUrl && r.logoUrl.startsWith('data:')) {
        try {
          const match = r.logoUrl.match(/^data:([^;]+);base64,(.+)$/);
          if (match) {
            const buf = Buffer.from(match[2], 'base64');
            db.saveInMemoryFile(r.logoFileId, buf, match[1], r.fileName || 'logo.png');
          }
        } catch (e) {}
      }
      await db.enqueue({ t: 'req', row, file: null, numbers });
      existingReqIds.add(r.requestId);
      addedCount++;
    }
  }

  for (const [reqId, status] of Object.entries(statusOverrides)) {
    if (['approved', 'rejected'].includes(status)) {
      await db.enqueue({ t: 'status', requestId: reqId, status, at: new Date().toISOString() });
    }
  }

  res.json({ ok: true, synced: addedCount });
}));

// ENTERPRISE: register a bank (official name + logo) with phone numbers per department (multipart; field "logo" + "departments" JSON)
app.post('/api/requests', only('enterprise'), upload.single('logo'), wrap(async (req, res) => {
  const { packageId, tagline = '', uploadedBy = 'enterprise-user' } = req.body;
  const brandName = String(req.body.brandName || '').trim().slice(0, 80);
  if (!packageId || !brandName) return res.status(400).json({ error: 'The package and the official bank name are required' });
  const s = await db.readAll();
  const pkg = s.packages.find(p => p.packageId === packageId && p.status === 'published');
  if (!pkg) return res.status(404).json({ error: 'package not found or not published' });

  let deps; try { deps = JSON.parse(req.body.departments || '[]'); } catch (e) { deps = null; }
  if (!Array.isArray(deps) || !deps.length) return res.status(400).json({ error: 'Add at least one department with a phone number' });
  if (deps.length > 10) return res.status(400).json({ error: 'A registration can have up to 10 departments' });
  const reqStatus = id => (s.requests.find(r => r.requestId === id) || {}).status;
  const errors = [], seen = new Set(), numbers = [], requestId = db.newId('REQ'), now = new Date().toISOString();
  for (const d of deps) {
    const dept = String(d.name || '').trim().slice(0, 60);
    if (!dept) { errors.push({ department: '(no name)', error: 'Every department needs a name' }); continue; }
    const list = (Array.isArray(d.numbers) ? d.numbers : String(d.numbers || '').split(/[\n,;]+/)).map(x => String(x).trim()).filter(Boolean);
    if (!list.length) errors.push({ department: dept, error: 'Add at least one phone number' });
    for (const raw of list) {
      const r = PH.normalize(raw);
      if (!r.ok) { errors.push({ department: dept, input: raw, error: r.error }); continue; }
      if (seen.has(r.e164)) { errors.push({ department: dept, input: raw, error: 'This number is listed twice' }); continue; }
      const clash = s.numbers.find(n => n.phoneE164 === r.e164 && activeStatus(reqStatus(n.requestId)));
      if (clash) { errors.push({ department: dept, input: raw, error: `Already registered by ${clash.brandName} (${reqStatus(clash.requestId)})` }); continue; }
      seen.add(r.e164);
      numbers.push({ numberId: db.newId('NUM'), requestId, brandName, department: dept, phoneE164: r.e164, phoneDisplay: r.display, numberType: r.type, status: 'pending', createdAt: now });
    }
  }
  if (errors.length) return res.status(400).json({ error: 'Please fix the phone numbers', errors });
  if (numbers.length > pkg.capacity) return res.status(400).json({ error: `The ${pkg.name} package allows up to ${pkg.capacity} numbers; you entered ${numbers.length}` });

  let file = null;
  if (req.file) {
    if (!req.file.size) return res.status(400).json({ error: 'The uploaded file is empty' });
    const type = classify(req.file.buffer, req.file.originalname);                 // VALIDATE
    if (!type) return res.status(415).json({ error: 'This file type is blocked for security (programs and scripts)' });
    const fileId = db.newId('FILE'), storedName = `${fileId}.${type.ext}`;         // STORE (server-made name)
    try {
      await fs.promises.writeFile(path.join(db.STORAGE_DIR, storedName), req.file.buffer);
    } catch (e) {
      console.warn('Could not write uploaded file to storage disk:', e.message);
    }
    const mime = type.mime || req.file.mimetype || 'image/png';
    const dataUrl = `data:${mime};base64,${req.file.buffer.toString('base64')}`;
    db.saveInMemoryFile(fileId, req.file.buffer, mime, req.file.originalname);

    file = {
      fileId, requestId, originalName: path.basename(req.file.originalname), storedName, storagePath: `storage/logos/${storedName}`,
      url: `${OPERATOR_URL}/api/files/${fileId}`, mimeType: mime, sizeBytes: req.file.size,
      sha256: crypto.createHash('sha256').update(req.file.buffer).digest('hex'), uploadedBy, uploadedAt: now,
      dataUrl
    };
  }
  const f = PH.fee(pkg.pricePHP || 0);
  const row = { requestId, packageId, brandName, tagline, status: 'pending', logoFileId: file ? file.fileId : '', createdAt: now,
    departments: [...new Set(numbers.map(n => n.department))].join(', '), numbersCount: numbers.length, monthlyFeePHP: f.fee, vatPHP: f.vat, totalPHP: f.total };
  await db.enqueue({ t: 'req', row, file, numbers });                                 // RECORD in the Excel workbook
  res.status(201).json({ ...row, numbers, logoUrl: file ? (file.dataUrl || `/api/files/${file.fileId}`) : null });
}));

// both sites: retrieve the ORIGINAL file (inline if the browser can show it, otherwise download)
app.get('/api/files/:id', wrap(async (req, res) => {
  const fileId = req.params.id;

  // 1. Check in-memory store
  const mem = db.getInMemoryFile(fileId);
  if (mem && mem.buffer) {
    const dispo = req.query.download || !PREVIEWABLE.has(mem.mimeType) ? 'attachment' : 'inline';
    res.set({
      'Content-Security-Policy': "default-src 'none'; style-src 'unsafe-inline'; img-src data:; sandbox",
      'Content-Type': mem.mimeType,
      'Content-Length': mem.buffer.length,
      'Cache-Control': 'public, max-age=86400',
      'Content-Disposition': `${dispo}; filename*=UTF-8''${encodeURIComponent(mem.originalName)}`,
      'X-Content-Type-Options': 'nosniff'
    });
    return res.send(mem.buffer);
  }

  // 2. Check database files array & physical disk
  const s = await db.readAll();
  const rec = (s.files || []).find(f => f.fileId === fileId);
  if (rec) {
    let full = path.join(db.SHARED_DIR, rec.storagePath);
    if (!fs.existsSync(full)) {
      full = path.join(__dirname, 'shared', rec.storagePath);
    }
    if (fs.existsSync(full)) {
      const bytes = await fs.promises.readFile(full);
      const dispo = req.query.download || !PREVIEWABLE.has(rec.mimeType) ? 'attachment' : 'inline';
      res.set({
        'Content-Security-Policy': "default-src 'none'; style-src 'unsafe-inline'; img-src data:; sandbox",
        'Content-Type': rec.mimeType,
        'Content-Length': bytes.length,
        'Cache-Control': 'public, max-age=86400',
        'Content-Disposition': `${dispo}; filename*=UTF-8''${encodeURIComponent(rec.originalName)}`,
        'X-Content-Type-Options': 'nosniff'
      });
      return res.send(bytes);
    }
  }

  // 3. Fallback: On serverless/ephemeral environments where an uploaded file is not on this instance,
  // serve the official Mobileum logo instead of returning a 404 error
  const fallbackPaths = [
    path.join(__dirname, 'public', 'shared', 'assets', 'mobileum-logo.png'),
    path.join(__dirname, 'shared', 'storage', 'logos', 'FILE-07c47447.png')
  ];
  for (const fp of fallbackPaths) {
    if (fs.existsSync(fp)) {
      const bytes = await fs.promises.readFile(fp);
      res.set({
        'Content-Type': 'image/png',
        'Content-Length': bytes.length,
        'Cache-Control': 'public, max-age=300',
        'Content-Disposition': 'inline; filename="logo.png"',
        'X-Content-Type-Options': 'nosniff'
      });
      return res.send(bytes);
    }
  }

  res.status(404).json({ error: 'file not found' });
}));

// OPERATOR: download a copy of the shared Excel database, or a live CSV feed for a second workbook
app.get('/api/database', only('operator'), (_req, res) => fs.existsSync(db.DB_PATH) ? res.download(db.DB_PATH, 'CallerID-Database-copy.xlsx') : res.status(404).json({ error: 'database file not created yet' }));
const csvCell = v => { let t = String(v ?? ''); if (/^[=+\-@\t\r]/.test(t)) t = "'" + t; return /[",\n\r]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; };
const csvRaw = v => String(v ?? '');                       // phone numbers: already validated, keep the leading +63
const sendCsv = (res, header, rows) => { res.set({ 'Content-Type': 'text/csv; charset=utf-8', 'Cache-Control': 'no-store' }); res.send('\ufeff' + [header.join(','), ...rows].join('\r\n') + '\r\n'); };
app.get('/api/requests.csv', only('operator'), wrap(async (_req, res) => {
  const H = ['requestId','packageId','packageName','brandName','tagline','status','createdAt','departments','numbersCount','monthlyFeePHP','vatPHP','totalPHP','logoFileName','logoType','logoSizeKB','logoUrl'];
  sendCsv(res, H, decorate(await db.readAll()).map(r => [r.requestId, r.packageId, r.packageName, r.brandName, r.tagline, r.status, r.createdAt, r.departments, r.numbersCount,
    r.monthlyFeePHP, r.vatPHP, r.totalPHP, r.fileName, r.mimeType, r.sizeBytes ? Math.max(1, Math.round(r.sizeBytes / 1024)) : '', r.logoUrl ? OPERATOR_URL + r.logoUrl : ''].map(csvCell).join(',')));
}));
app.get('/api/numbers.csv', only('operator'), wrap(async (_req, res) => {
  const H = ['numberId','requestId','brandName','department','phone','numberType','status','createdAt'];
  sendCsv(res, H, (await db.readAll()).numbers.map(n => [csvCell(n.numberId), csvCell(n.requestId), csvCell(n.brandName), csvCell(n.department), csvRaw(n.phoneDisplay), csvCell(n.numberType), csvCell(n.status), csvCell(n.createdAt)].join(',')));
}));
app.get('/api/packages.csv', only('operator'), wrap(async (_req, res) => {
  const H = ['packageId','name','capacity','pricePHP','brand','tagline','logoIncluded','status','createdAt'];
  sendCsv(res, H, (await db.readAll()).packages.map(r => H.map(k => csvCell(r[k])).join(',')));
}));
// the whole workbook as JSON, for the live sheet page
app.get('/api/sheet', only('operator'), wrap(async (_req, res) => {
  const s = await db.readAll(); res.set('Cache-Control', 'no-store');
  res.json({ packages: s.packages, requests: decorate(s), numbers: s.numbers, files: s.files, status: db.getStatus() });
}));

app.use((err, _req, res, _next) => {
  if (err.code === 'LIMIT_FILE_SIZE') return res.status(413).json({ error: `File too large. Max ${Math.round(MAX_BYTES / 1024 / 1024)} MB` });
  console.error(`[${ROLE}]`, err); res.status(500).json({ error: err.message || 'server error' });
});

if (require.main === module) {
  app.listen(PORT, () => {
    const desc = IS_UNIFIED ? 'Mobileum Unified Portal' : { operator: 'Operator portal', enterprise: 'Enterprise portal', customer: 'Customer phone (user portal)' }[ROLE];
    console.log(`[${ROLE}] ${desc} running on http://localhost:${PORT}`);
    console.log(`[${ROLE}] shared Excel database: ${db.DB_PATH}`);
    db.startBackground();
  });
}

module.exports = app;
