'use strict';
// One code base, two websites:  node server.js operator   (port 4001)
//                               node server.js enterprise (port 4002)
// Both read and write the SAME Excel workbook through db.js.
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
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
app.use(express.json({ limit: MAX_BYTES }));
app.use(express.urlencoded({ extended: true, limit: MAX_BYTES }));
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
  const files = Object.fromEntries((s.files || []).map(f => [f.fileId, f])), pk = Object.fromEntries((s.packages || []).map(p => [p.packageId, p]));
  return (s.requests || []).map(r => {
    const f = files[r.logoFileId], p = pk[r.packageId];
    const numbers = (s.numbers || []).filter(n => n.requestId === r.requestId).map(n => ({
      numberId: n.numberId,
      department: n.department,
      phoneE164: n.phoneE164,
      phoneDisplay: n.phoneDisplay,
      numberType: n.numberType,
      status: n.status || 'pending',
      remarks: n.remarks || '',
      decidedAt: n.decidedAt || null,
      createdAt: n.createdAt
    }));
    const derivedStatus = numbers.length ? db.deriveRequestStatus(numbers) : (r.status || 'pending');
    const logoUrl = f ? `/api/files/${f.fileId}` : (r.logoUrl || null);
    const callReason = r.callReason || r.tagline || '';
    return {
      ...r,
      status: derivedStatus,
      callReason,
      tagline: callReason,
      remarks: r.remarks || '',
      packageName: p ? p.name : '',
      numbers,
      numbersCount: numbers.length,
      fileName: f ? f.originalName : null,
      mimeType: f ? f.mimeType : null,
      sizeBytes: f ? f.sizeBytes : null,
      previewable: f ? PREVIEWABLE.has(f.mimeType) && f.mimeType.startsWith('image/') : false,
      logoUrl
    };
  });
}
// what a customer's phone would show for a number (the network looks the caller up in the operator's approved registry)
function lookup(s, input) {
  const r = PH.normalize(input);
  if (!r.ok) return { registered: false, reason: 'invalid', message: r.error };
  const reqOf = id => s.requests.find(x => x.requestId === id);
  const hit = s.numbers.find(n => n.phoneE164 === r.e164);
  const base = { e164: r.e164, display: r.display, numberType: r.type };
  if (!hit) return { ...base, registered: false, reason: 'unregistered' };
  const q = reqOf(hit.requestId);
  if (!q) return { ...base, registered: false, reason: 'unregistered' };
  
  // Only numbers with status 'approved' can show the brand and verified state
  if (hit.status !== 'approved') {
    return {
      ...base,
      registered: false,
      reason: hit.status === 'rejected' ? 'rejected' : 'pending',
      brandName: q.brandName,
      message: hit.status === 'rejected' ? 'Number registration was rejected' : 'Number registration is pending approval'
    };
  }

  const f = (s.files || []).find(x => x.fileId === q.logoFileId);
  const logoUrl = f && f.mimeType && f.mimeType.startsWith('image/') && PREVIEWABLE.has(f.mimeType) ? `/api/files/${f.fileId}` : (q.logoUrl || null);
  const callReason = q.callReason || q.tagline || '';
  return { ...base, registered: true, verified: true, brandName: q.brandName, callReason, tagline: callReason, department: hit.department, logoUrl };
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
  res.json(s.packages.filter(p => p.status !== 'removed' && (!isEnterprise || p.status === 'published')).map(p => ({
    ...p,
    duration: p.duration || '30 Days',
    brandedCalls: Number(p.brandedCalls) || (Number(p.capacity || 5) * 10000) || 50000
  })));
}));
app.get('/api/requests', only('operator', 'enterprise'), wrap(async (_req, res) => res.json(decorate(await db.readAll()))));
// registered numbers (the customer phone uses this to offer demo callers) and the caller-ID lookup
app.get('/api/numbers', wrap(async (req, res) => {
  const s = await db.readAll();
  const includeAll = req.query.all === '1' || req.query.all === 'true';
  const list = includeAll ? s.numbers : s.numbers.filter(n => n.status === 'approved');
  res.json(list.map(n => ({
    numberId: n.numberId,
    requestId: n.requestId,
    phoneE164: n.phoneE164,
    phoneDisplay: n.phoneDisplay,
    department: n.department,
    brandName: n.brandName,
    enterprise: n.brandName,
    numberType: n.numberType,
    status: n.status || 'pending',
    remarks: n.remarks || '',
    decidedAt: n.decidedAt || null
  })));
}));
app.get('/api/lookup', wrap(async (req, res) => res.json(lookup(await db.readAll(), req.query.number))));

// Packages: operator sees all active, enterprise sees only published ones
app.post('/api/packages', only('operator', 'enterprise'), wrap(async (req, res) => {
  const { name, capacity, duration = '30 Days', pricePHP = 0, brandedCalls, brand = true, callReason = true, tagline = true, logoIncluded = false } = req.body || {};
  const cap = Number(capacity);
  if (typeof name !== 'string' || !name.trim() || !Number.isInteger(cap) || cap <= 0) {
    return res.status(400).json({ error: 'name and a positive whole-number capacity are required' });
  }
  const price = typeof pricePHP === 'number' ? pricePHP : parseFloat(pricePHP || '0');
  if (isNaN(price) || price < 0) {
    return res.status(400).json({ error: 'price must be an amount in Philippine pesos (₱), 0 or more' });
  }
  const calls = Number(brandedCalls) || (cap * 10000) || 50000;
  const row = {
    packageId: db.newId('PKG'),
    name: name.trim().slice(0, 80),
    capacity: cap,
    duration: String(duration || '30 Days').trim().slice(0, 30),
    pricePHP: PH.round2(price),
    brandedCalls: calls,
    brand: truthy(brand),
    callReason: truthy(callReason !== undefined ? callReason : tagline),
    tagline: truthy(callReason !== undefined ? callReason : tagline),
    logoIncluded: truthy(logoIncluded),
    status: 'published',
    createdAt: new Date().toISOString()
  };
  await db.enqueue({ t: 'pkg', row });
  res.status(201).json(row);
}));

async function updatePackage(req, res) {
  const s = await db.readAll();
  const current = s.packages.find(p => p.packageId === req.params.id && p.status !== 'removed');
  if (!current) return res.status(404).json({ error: 'package not found' });

  const body = req.body || {};
  const name = body.name === undefined ? current.name : (typeof body.name === 'string' ? body.name.trim() : '');
  const capacity = body.capacity === undefined ? Number(current.capacity) : Number(body.capacity);
  const pricePHP = body.pricePHP === undefined ? Number(current.pricePHP) : Number(body.pricePHP);
  if (typeof name !== 'string' || !name || !Number.isInteger(capacity) || capacity <= 0) {
    return res.status(400).json({ error: 'name and a positive whole-number capacity are required' });
  }
  if (typeof pricePHP !== 'number' || !isFinite(pricePHP) || pricePHP < 0) {
    return res.status(400).json({ error: 'price must be an amount in Philippine pesos (₱), 0 or more' });
  }

  const calls = body.brandedCalls !== undefined ? Number(body.brandedCalls) : (Number(current.brandedCalls) || (capacity * 10000) || 50000);
  const row = {
    ...current,
    name: name.slice(0, 80),
    capacity,
    duration: body.duration !== undefined ? String(body.duration).trim().slice(0, 30) : (current.duration || '30 Days'),
    pricePHP: PH.round2(pricePHP),
    brandedCalls: calls,
    brand: body.brand === undefined ? truthy(current.brand) : truthy(body.brand),
    callReason: body.callReason !== undefined ? truthy(body.callReason) : (body.tagline !== undefined ? truthy(body.tagline) : truthy(current.callReason || current.tagline)),
    tagline: body.callReason !== undefined ? truthy(body.callReason) : (body.tagline !== undefined ? truthy(body.tagline) : truthy(current.callReason || current.tagline)),
    logoIncluded: body.logoIncluded === undefined ? truthy(current.logoIncluded) : truthy(body.logoIncluded)
  };
  await db.enqueue({ t: 'pkg-update', row });
  res.json(row);
}

app.put('/api/packages/:id', only('operator', 'enterprise'), wrap(updatePackage));
app.patch('/api/packages/:id', only('operator', 'enterprise'), wrap(updatePackage));

// Keep historical references valid while persistently removing plans from both portals.
app.delete('/api/packages/:id', only('operator', 'enterprise'), wrap(async (req, res) => {
  const s = await db.readAll();
  const pkg = s.packages.find(p => p.packageId === req.params.id);
  if (!pkg) return res.status(404).json({ error: 'package not found' });
  if (pkg.status === 'removed') return res.json({ packageId: pkg.packageId, status: 'removed' });
  await db.enqueue({ t: 'pkg-remove', packageId: pkg.packageId });
  res.json({ packageId: pkg.packageId, status: 'removed' });
}));

// OPERATOR: approve / reject individual numbers or whole request
async function handleOperatorDecision(req, res) {
  const { decisions, status, remarks = '', requestRemarks = '' } = req.body || {};
  const s = await db.readAll();
  const r = s.requests.find(x => x.requestId === req.params.id);
  if (!r) return res.status(404).json({ error: 'Brand request not found' });
  const existingNumbers = s.numbers.filter(n => n.requestId === req.params.id);

  if (Array.isArray(decisions) && decisions.length > 0) {
    const cleanDecisions = [];
    for (const d of decisions) {
      const st = String(d.status || '').toLowerCase().trim();
      if (!['approved', 'rejected'].includes(st)) {
        return res.status(400).json({ error: `Invalid status "${d.status}". Status must be approved or rejected.` });
      }
      const rem = String(d.remarks || '').trim();
      if (st === 'rejected' && !rem) {
        return res.status(400).json({ error: 'Rejection remarks are required for all unticked/rejected numbers.' });
      }
      cleanDecisions.push({
        numberId: d.numberId,
        phoneE164: d.phoneE164,
        status: st,
        remarks: st === 'rejected' ? rem.slice(0, 500) : ''
      });
    }

    const updatedNumbers = existingNumbers.map(n => {
      const d = cleanDecisions.find(item => (item.numberId && item.numberId === n.numberId) || (item.phoneE164 && item.phoneE164 === n.phoneE164));
      return d ? { ...n, status: d.status, remarks: d.remarks } : n;
    });

    const finalReqStatus = db.deriveRequestStatus(updatedNumbers);
    const overallRemarks = String(requestRemarks || remarks || '').trim().slice(0, 500);

    await db.enqueue({
      t: 'number-decisions',
      requestId: req.params.id,
      decisions: cleanDecisions,
      requestStatus: finalReqStatus,
      remarks: overallRemarks,
      at: new Date().toISOString()
    });

    return res.json({
      requestId: req.params.id,
      status: finalReqStatus,
      remarks: overallRemarks,
      decisions: cleanDecisions
    });
  }

  // Fallback for single request status update
  if (!['approved', 'rejected', 'partially approved'].includes(status)) {
    return res.status(400).json({ error: 'status must be approved or rejected' });
  }
  if (status === 'rejected' && (!remarks || !String(remarks).trim())) {
    return res.status(400).json({ error: 'Rejection remarks are required when rejecting a brand request' });
  }
  const cleanRemarks = remarks ? String(remarks).trim().slice(0, 500) : '';
  await db.enqueue({ t: 'status', requestId: req.params.id, status, remarks: cleanRemarks, at: new Date().toISOString() });
  res.json({ requestId: req.params.id, status, remarks: cleanRemarks });
}

app.patch('/api/requests/:id/status', only('operator'), wrap(handleOperatorDecision));
app.post('/api/requests/:id/decision', only('operator'), wrap(handleOperatorDecision));

// ENTERPRISE: fix and resubmit rejected numbers
app.post('/api/requests/:id/resubmit-numbers', only('enterprise', 'operator'), wrap(async (req, res) => {
  const { numbers = [] } = req.body || {};
  if (!Array.isArray(numbers) || !numbers.length) {
    return res.status(400).json({ error: 'At least one number is required to resubmit' });
  }
  const s = await db.readAll();
  const r = s.requests.find(x => x.requestId === req.params.id);
  if (!r) return res.status(404).json({ error: 'Brand request not found' });

  const existingNumbers = s.numbers.filter(n => n.requestId === req.params.id);
  const validated = [];
  const errors = [];
  for (const item of numbers) {
    const raw = item.phone || item.phoneE164 || item.phoneDisplay || '';
    const norm = PH.normalize(raw);
    if (!norm.ok) {
      errors.push({ number: raw, error: norm.error });
      continue;
    }
    const dept = String(item.department || '').trim().slice(0, 60);
    if (!dept) {
      errors.push({ number: raw, error: 'Department is required' });
      continue;
    }
    const clash = s.numbers.find(n => n.phoneE164 === norm.e164 && n.requestId !== req.params.id && n.status === 'approved');
    if (clash) {
      errors.push({ number: raw, error: `Already registered by ${clash.brandName || 'another brand'}` });
      continue;
    }
    validated.push({
      numberId: item.numberId,
      phoneE164: norm.e164,
      phoneDisplay: norm.display,
      department: dept,
      numberType: norm.type
    });
  }
  if (errors.length) {
    return res.status(400).json({ error: 'Please fix phone number errors', errors });
  }

  const updatedReqNumbers = existingNumbers.map(n => {
    const match = validated.find(v => (v.numberId && v.numberId === n.numberId) || v.phoneE164 === n.phoneE164);
    if (match) return { ...n, ...match, status: 'pending', remarks: '', decidedAt: null };
    return n;
  });
  const newStatus = db.deriveRequestStatus(updatedReqNumbers);

  await db.enqueue({
    t: 'resubmit-numbers',
    requestId: req.params.id,
    numbers: validated,
    requestStatus: newStatus,
    at: new Date().toISOString()
  });

  res.json({ ok: true, requestId: req.params.id, status: newStatus, resubmittedCount: validated.length });
}));

// SERVERLESS & CLIENT SYNC: Re-hydrates requests, status overrides, packages, and removals across instances
app.post('/api/sync', wrap(async (req, res) => {
  const { requests = [], statusOverrides = {}, packages = [], deletedPackageIds = [] } = req.body || {};
  const s = await db.readAll();
  const existingReqIds = new Set(s.requests.map(r => r.requestId));
  let addedCount = 0;
  let addedPkgCount = 0;
  let deletedPkgCount = 0;

  // Process package deletions first
  const existingPkgs = new Map(s.packages.map(p => [p.packageId, p]));
  for (const delId of deletedPackageIds) {
    if (delId && existingPkgs.has(delId) && existingPkgs.get(delId).status !== 'removed') {
      await db.enqueue({ t: 'pkg-remove', packageId: delId });
      deletedPkgCount++;
    }
  }

  // Process packages
  for (const p of packages) {
    if (!p || !p.packageId || deletedPackageIds.includes(p.packageId)) continue;
    if (!existingPkgs.has(p.packageId)) {
      const row = {
        packageId: p.packageId,
        name: String(p.name || '').trim().slice(0, 80),
        capacity: Number(p.capacity) || 5,
        pricePHP: PH.round2(Number(p.pricePHP) || 0),
        brand: truthy(p.brand),
        tagline: truthy(p.tagline),
        logoIncluded: truthy(p.logoIncluded),
        status: p.status || 'published',
        createdAt: p.createdAt || new Date().toISOString()
      };
      await db.enqueue({ t: 'pkg', row });
      existingPkgs.set(p.packageId, row);
      addedPkgCount++;
    }
  }

  // Process requests
  for (const r of requests) {
    if (!r || !r.requestId) continue;
    if (!existingReqIds.has(r.requestId)) {
      const override = statusOverrides[r.requestId];
      const finalStatus = typeof override === 'string' ? override : (override?.status || r.status || 'pending');
      const finalRemarks = typeof override === 'object' ? (override?.remarks || '') : (r.remarks || '');
      const callReason = r.callReason || r.tagline || '';
      const row = {
        requestId: r.requestId,
        packageId: r.packageId || '',
        brandName: r.brandName || '',
        callReason,
        tagline: callReason,
        status: finalStatus,
        remarks: finalRemarks,
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
        status: n.status || finalStatus,
        remarks: n.remarks || '',
        decidedAt: n.decidedAt || null,
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

  for (const [reqId, override] of Object.entries(statusOverrides)) {
    if (typeof override === 'object' && Array.isArray(override?.decisions) && override.decisions.length > 0) {
      await db.enqueue({
        t: 'number-decisions',
        requestId: reqId,
        decisions: override.decisions,
        remarks: override.remarks || '',
        at: override.at || new Date().toISOString()
      });
    } else {
      const status = typeof override === 'string' ? override : override?.status;
      const remarks = typeof override === 'object' ? (override?.remarks || '') : '';
      if (['approved', 'rejected', 'partially approved'].includes(status)) {
        await db.enqueue({ t: 'status', requestId: reqId, status, remarks, at: new Date().toISOString() });
      }
    }
  }

  res.json({ ok: true, syncedRequests: addedCount, syncedPackages: addedPkgCount, deletedPackages: deletedPkgCount });
}));

// DEPARTMENT LINES ENDPOINTS
app.get('/api/department-lines', wrap(async (_req, res) => {
  const s = await db.readAll();
  res.json(s.departmentLines || []);
}));

app.post('/api/department-lines', only('enterprise', 'operator'), wrap(async (req, res) => {
  const { department, phoneE164, brandName = '' } = req.body || {};
  if (!department || !phoneE164) {
    return res.status(400).json({ error: 'Department and phone number are required' });
  }
  const r = PH.normalize(phoneE164);
  if (!r.ok) return res.status(400).json({ error: r.error });
  const row = {
    lineId: req.body.lineId || db.newId('DLINE'),
    department: String(department).trim().slice(0, 60),
    phoneE164: r.e164,
    phoneDisplay: r.display,
    numberType: r.type,
    brandName: String(brandName || '').trim().slice(0, 80),
    updatedAt: new Date().toISOString()
  };
  await db.enqueue({ t: 'dept-line', row });
  res.json(row);
}));

app.delete('/api/department-lines/:id', only('enterprise', 'operator'), wrap(async (req, res) => {
  await db.enqueue({ t: 'dept-line-remove', lineId: req.params.id });
  res.json({ ok: true, lineId: req.params.id });
}));

// ENTERPRISE: register a brand request with logo, call reason, departments and individual numbers
app.post('/api/requests', only('enterprise'), upload.single('logo'), wrap(async (req, res) => {
  const { packageId, callReason = '', tagline = '', uploadedBy = 'enterprise-user' } = req.body;
  const brandName = String(req.body.brandName || '').trim().slice(0, 80);
  const finalCallReason = String(callReason || tagline || '').trim().slice(0, 120);
  if (!brandName) return res.status(400).json({ error: 'Official brand name is required' });

  const s = await db.readAll();
  // If packageId is provided, validate it; otherwise use active published package if enterprise subscribed
  let pkg = null;
  if (packageId) {
    pkg = s.packages.find(p => p.packageId === packageId && p.status === 'published');
  }
  if (!pkg) {
    pkg = s.packages.find(p => p.status === 'published') || { name: 'Standard Pack', capacity: 500, pricePHP: 15000 };
  }

  let deps = [];
  try { deps = JSON.parse(req.body.departments || '[]'); } catch (e) { deps = []; }

  // If numbers array was passed directly (e.g. [{ phone, department }])
  if ((!Array.isArray(deps) || !deps.length) && req.body.numbers) {
    try {
      const rawList = JSON.parse(req.body.numbers || '[]');
      if (Array.isArray(rawList) && rawList.length) {
        const byDept = {};
        for (const item of rawList) {
          const dept = String(item.department || '').trim();
          const p = item.phone || item.phoneE164 || item.phoneDisplay || '';
          if (!byDept[dept]) byDept[dept] = [];
          if (p) byDept[dept].push(p);
        }
        deps = Object.entries(byDept).map(([name, numbers]) => ({ name, numbers }));
      }
    } catch (e) {}
  }

  if (!Array.isArray(deps) || !deps.length) return res.status(400).json({ error: 'Add at least one calling number with an assigned department' });
  if (deps.length > 50) return res.status(400).json({ error: 'A registration can have up to 50 departments' });

  const reqStatus = id => (s.requests.find(r => r.requestId === id) || {}).status;
  const errors = [], seen = new Set(), numbers = [], requestId = db.newId('REQ'), now = new Date().toISOString();
  for (const d of deps) {
    const dept = String(d.name || '').trim().slice(0, 60);
    if (!dept) { errors.push({ department: '(no name)', error: 'Every calling number must belong to a department' }); continue; }
    const list = (Array.isArray(d.numbers) ? d.numbers : String(d.numbers || '').split(/[\n,;]+/)).map(x => String(x).trim()).filter(Boolean);
    if (!list.length) errors.push({ department: dept, error: `Department "${dept}" has no phone numbers assigned` });
    for (const raw of list) {
      const r = PH.normalize(raw);
      if (!r.ok) { errors.push({ department: dept, input: raw, error: r.error }); continue; }
      if (seen.has(r.e164)) { errors.push({ department: dept, input: raw, error: 'This number is listed twice in this request' }); continue; }
      // Check against active registered numbers across other enterprises
      const clash = s.numbers.find(n => n.phoneE164 === r.e164);
      if (clash && clash.status !== 'rejected') {
        const st = reqStatus(clash.requestId) || clash.status || 'saved';
        errors.push({ department: dept, input: raw, error: `Already registered by ${clash.brandName || 'another brand'} (${st})` });
        continue;
      }
      seen.add(r.e164);
      numbers.push({
        numberId: db.newId('NUM'),
        requestId,
        brandName,
        enterprise: brandName,
        department: dept,
        phoneE164: r.e164,
        phoneDisplay: r.display,
        numberType: r.type,
        status: 'pending',
        remarks: '',
        decidedAt: null,
        createdAt: now
      });
    }
  }
  if (errors.length) return res.status(400).json({ error: 'Please fix the phone numbers and department assignments', errors });
  if (pkg && numbers.length > pkg.capacity) return res.status(400).json({ error: `The ${pkg.name} package allows up to ${pkg.capacity} numbers; you entered ${numbers.length}` });

  let file = null;
  if (req.file) {
    if (!req.file.size) return res.status(400).json({ error: 'The uploaded file is empty' });
    if (req.file.size > 5 * 1024 * 1024) return res.status(400).json({ error: 'Logo file size exceeds 5MB limit' });

    const type = classify(req.file.buffer, req.file.originalname);
    if (!type || !['png', 'jpg', 'svg', 'webp'].includes(type.ext)) {
      return res.status(400).json({ error: 'Only PNG, JPG, and SVG image files are allowed for brand logos' });
    }
    const fileId = db.newId('FILE'), storedName = `${fileId}.${type.ext}`;
    try {
      fs.mkdirSync(db.STORAGE_DIR, { recursive: true });
      await fs.promises.writeFile(path.join(db.STORAGE_DIR, storedName), req.file.buffer);
    } catch (e) {
      console.warn('Could not write uploaded file to storage disk:', e.message);
    }
    const mime = type.mime || req.file.mimetype || 'image/png';
    const dataUrl = `data:${mime};base64,${req.file.buffer.toString('base64')}`;
    db.saveInMemoryFile(fileId, req.file.buffer, mime, req.file.originalname);

    file = {
      fileId, requestId, originalName: path.basename(req.file.originalname), storedName, storagePath: `storage/logos/${storedName}`,
      url: `/api/files/${fileId}`, mimeType: mime, sizeBytes: req.file.size,
      sha256: crypto.createHash('sha256').update(req.file.buffer).digest('hex'), uploadedBy, uploadedAt: now,
      dataUrl
    };
  }
  const f = PH.fee(pkg ? (pkg.pricePHP || 0) : 0);
  const row = {
    requestId, packageId: pkg ? pkg.packageId : '', brandName,
    callReason: finalCallReason, tagline: finalCallReason, status: 'pending', remarks: '',
    logoFileId: file ? file.fileId : '', createdAt: now,
    departments: [...new Set(numbers.map(n => n.department))].join(', '), numbersCount: numbers.length, monthlyFeePHP: f.fee, vatPHP: f.vat, totalPHP: f.total
  };
  await db.enqueue({ t: 'req', row, file, numbers });
  res.status(201).json({ ...row, numbers, logoUrl: file ? `/api/files/${file.fileId}` : null });
}));

// both sites: retrieve the ORIGINAL file (inline if previewable, with no-cache to avoid stale logos)
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
      'Cache-Control': 'no-cache, must-revalidate',
      'Content-Disposition': `${dispo}; filename*=UTF-8''${encodeURIComponent(mem.originalName)}`,
      'X-Content-Type-Options': 'nosniff'
    });
    return res.send(mem.buffer);
  }

  const storedFile = await db.getStoredFile(fileId);
  if (storedFile && storedFile.buffer) {
    const dispo = req.query.download || !PREVIEWABLE.has(storedFile.mimeType) ? 'attachment' : 'inline';
    res.set({
      'Content-Security-Policy': "default-src 'none'; style-src 'unsafe-inline'; img-src data:; sandbox",
      'Content-Type': storedFile.mimeType,
      'Content-Length': storedFile.buffer.length,
      'Cache-Control': 'no-cache, must-revalidate',
      'Content-Disposition': `${dispo}; filename*=UTF-8''${encodeURIComponent(storedFile.originalName)}`,
      'X-Content-Type-Options': 'nosniff'
    });
    return res.send(storedFile.buffer);
  }

  // 2. Directly inspect STORAGE_DIR on disk
  const searchDirs = [
    db.STORAGE_DIR,
    path.join(db.SHARED_DIR, 'storage', 'logos'),
    path.join(__dirname, 'shared', 'storage', 'logos')
  ];

  for (const sDir of searchDirs) {
    if (fs.existsSync(sDir)) {
      try {
        const entries = fs.readdirSync(sDir);
        const match = entries.find(fn => fn === fileId || fn.startsWith(fileId + '.'));
        if (match) {
          const fullPath = path.join(sDir, match);
          const bytes = await fs.promises.readFile(fullPath);
          const ext = path.extname(match).slice(1).toLowerCase();
          const mime = ext === 'svg' ? 'image/svg+xml' : (ext === 'jpg' || ext === 'jpeg') ? 'image/jpeg' : ext === 'webp' ? 'image/webp' : 'image/png';
          db.saveInMemoryFile(fileId, bytes, mime, match);
          res.set({
            'Content-Type': mime,
            'Content-Length': bytes.length,
            'Cache-Control': 'no-cache, must-revalidate',
            'Content-Disposition': `inline; filename="${match}"`,
            'X-Content-Type-Options': 'nosniff'
          });
          return res.send(bytes);
        }
      } catch (e) {}
    }
  }

  // 3. Check database files array & physical disk
  const s = await db.readAll();
  const rec = (s.files || []).find(f => f.fileId === fileId);
  if (rec && rec.storagePath) {
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
        'Cache-Control': 'no-cache, must-revalidate',
        'Content-Disposition': `${dispo}; filename*=UTF-8''${encodeURIComponent(rec.originalName)}`,
        'X-Content-Type-Options': 'nosniff'
      });
      return res.send(bytes);
    }
  }

  res.status(404).json({ error: 'file not found' });
}));

// OPERATOR: download a copy of the configured database
app.get('/api/database', only('operator'), wrap(async (_req, res) => {
  if (db.getStatus().backend !== 'mongodb') {
    return fs.existsSync(db.DB_PATH)
      ? res.download(db.DB_PATH, 'CallerID-Database-copy.xlsx')
      : res.status(404).json({ error: 'database file not created yet' });
  }

  const state = await db.readAll();
  const workbook = new ExcelJS.Workbook();
  for (const [sheetName, key] of [
    ['Packages', 'packages'],
    ['Requests', 'requests'],
    ['Numbers', 'numbers'],
    ['Files', 'files']
  ]) {
    const worksheet = workbook.addWorksheet(sheetName);
    const columns = db.COLS[sheetName];
    worksheet.addRow(columns).font = { bold: true };
    for (const row of state[key]) worksheet.addRow(columns.map(column => row[column] ?? ''));
    worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  }
  const content = Buffer.from(await workbook.xlsx.writeBuffer());
  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': 'attachment; filename="CallerID-Database-copy.xlsx"',
    'Cache-Control': 'no-store'
  });
  res.send(content);
}));
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
