/* ==========================================================================
   Mobileum Caller ID Platform - Shared Common Utilities
   ========================================================================== */

const $ = id => document.getElementById(id);

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;'
}[c]));

const peso = n => '₱' + Number(n || 0).toLocaleString('en-PH', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
});

const initials = n => String(n || '?')
  .split(/\s+/)
  .filter(Boolean)
  .slice(0, 2)
  .map(w => w[0].toUpperCase())
  .join('');

async function api(url, opt) {
  const r = await fetch(url, opt);
  let j = null;
  try { j = await r.json(); } catch (e) {}
  if (!r.ok) {
    const e = new Error((j && j.error) || r.statusText);
    e.data = j;
    e.status = r.status;
    throw e;
  }
  return j;
}

// Mobileum signature dots loader HTML
function dotsLoader() {
  return '<div class="mobileum-dots" title="Loading..."><span></span><span></span><span></span></div>';
}

// Toast notification helper
function showToast(title, message, type = 'info', duration = 4000) {
  let container = document.getElementById('toastContainer');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toastContainer';
    container.className = 'toast-container';
    document.body.appendChild(container);
  }

  const icons = {
    success: '✓',
    error: '✕',
    info: 'ℹ',
    warning: '⚠'
  };

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `
    <div class="toast-icon">${icons[type] || 'ℹ'}</div>
    <div class="toast-content">
      <div class="toast-title">${esc(title)}</div>
      ${message ? `<div class="toast-message">${esc(message)}</div>` : ''}
    </div>
    <button class="toast-close" aria-label="Close">&times;</button>
  `;

  toast.querySelector('.toast-close').onclick = () => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 200);
  };

  container.appendChild(toast);
  // Trigger animation frame
  requestAnimationFrame(() => {
    toast.classList.add('show');
  });

  if (duration > 0) {
    setTimeout(() => {
      if (toast.parentElement) {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 250);
      }
    }, duration);
  }
  return toast;
}

// Modern enterprise "what the customer will see" card
function callerCard(o) {
  const lg = o.logoUrl
    ? `<img src="${esc(o.logoUrl)}" alt="${esc(o.brandName || 'Logo')}" onerror="this.onerror=null;this.src='/shared/assets/mobileum-logo.png';">`
    : esc(initials(o.brandName));
  const isVerified = o.verified !== false;
  return `
    <div class="cid">
      <div class="lg">${lg}</div>
      <div style="flex:1;min-width:0;">
        <div class="nm">${esc(o.brandName)}</div>
        <div class="sub">${esc(o.department || '')}${o.department && o.display ? ' &middot; ' : ''}${esc(o.display || '')}</div>
        ${o.tagline ? `<div class="sub" style="font-style:italic;opacity:0.85;">"${esc(o.tagline)}"</div>` : ''}
        ${isVerified ? '<div class="vf"><svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg> Verified business</div>' : '<div style="color:#F59E0B;font-size:11.5px;margin-top:3px;">⚠ Unverified preview</div>'}
      </div>
    </div>
  `;
}

// Shows where the data is saved (the shared Excel workbook)
function dbStatus(el, s) {
  if (!el) return;
  if (s.error) {
    el.innerHTML = `<span class="err">Database unavailable: ${esc(s.error)}</span>`;
    return;
  }
  const name = (s.database || '').split(/[\\/]/).pop() || 'CallerID-Database.xlsx';
  if (s.backend === 'mongodb') {
    el.innerHTML = `<span style="color:var(--status-approved)">●</span> Shared MongoDB database: <b>${esc(name.replace(/^MongoDB:\s*/, ''))}</b>` +
      (s.lastFlush ? ` &middot; last saved ${new Date(s.lastFlush).toLocaleTimeString('en-PH', { timeZone: 'Asia/Manila' })}` : '');
    return;
  }
  el.innerHTML = s.pending > 0
    ? `<span class="err"><strong>${s.pending}</strong> change(s) waiting to be written into <b>${esc(name)}</b> (close it in Excel if open)</span>`
    : `<span style="color:var(--status-approved)">●</span> Shared Excel database: <b>${esc(name)}</b>` +
      (s.lastFlush ? ` &middot; last saved ${new Date(s.lastFlush).toLocaleTimeString('en-PH', { timeZone: 'Asia/Manila' })}` : '');
}

// In-app non-blocking confirmation modal dialog
function showConfirmModal(opts) {
  const {
    title = 'Confirm Removal',
    message = 'Are you sure you want to proceed?',
    confirmText = 'Remove',
    confirmClass = 'no',
    onConfirm
  } = opts || {};

  let modal = document.getElementById('mobileumConfirmModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'mobileumConfirmModal';
    modal.className = 'confirm-modal-backdrop';
    document.body.appendChild(modal);
  }

  modal.innerHTML = `
    <div class="confirm-modal-box">
      <div class="confirm-modal-header">
        <h3 class="confirm-modal-title">${esc(title)}</h3>
        <button type="button" class="confirm-modal-close" id="btnConfirmClose" title="Cancel">&times;</button>
      </div>
      <div class="confirm-modal-body">
        <p style="margin:0;font-size:13.5px;color:var(--deep-navy);line-height:1.5;">${esc(message)}</p>
      </div>
      <div class="confirm-modal-footer">
        <button type="button" class="btn sm" id="btnConfirmCancel">Cancel</button>
        <button type="button" class="btn sm ${esc(confirmClass)}" id="btnConfirmOk">${esc(confirmText)}</button>
      </div>
    </div>
  `;

  modal.classList.add('open');

  const close = () => {
    modal.classList.remove('open');
  };

  modal.querySelector('#btnConfirmClose').onclick = close;
  modal.querySelector('#btnConfirmCancel').onclick = close;
  modal.onclick = e => {
    if (e.target === modal) close();
  };

  modal.querySelector('#btnConfirmOk').onclick = () => {
    close();
    if (typeof onConfirm === 'function') {
      try { onConfirm(); } catch (err) { console.error(err); }
    }
  };
}

// Cross-port shared removal state for packages, numbers, and requests
function getSharedList(key) {
  let list = [];
  try {
    const match = document.cookie.match(new RegExp('(?:^|;\\s*)' + key + '=([^;]*)'));
    if (match && match[1]) {
      list = JSON.parse(decodeURIComponent(match[1]));
    }
  } catch (e) {}

  if (!Array.isArray(list) || !list.length) {
    try {
      const stored = localStorage.getItem(key);
      if (stored) list = JSON.parse(stored);
    } catch (e) {}
  }
  return Array.isArray(list) ? list : [];
}

function addToSharedList(key, item) {
  const current = getSharedList(key);
  if (!current.includes(item)) {
    current.push(item);
  }
  const str = encodeURIComponent(JSON.stringify(current));
  document.cookie = `${key}=${str}; path=/; max-age=2592000; SameSite=Lax`;
  try { document.cookie = `${key}=${str}; path=/; domain=localhost; max-age=2592000; SameSite=Lax`; } catch (e) {}
  try { localStorage.setItem(key, JSON.stringify(current)); } catch (e) {}
  return current;
}

function removeFromSharedList(key, item) {
  let current = getSharedList(key);
  current = current.filter(x => x !== item);
  const str = encodeURIComponent(JSON.stringify(current));
  document.cookie = `${key}=${str}; path=/; max-age=2592000; SameSite=Lax`;
  try { document.cookie = `${key}=${str}; path=/; domain=localhost; max-age=2592000; SameSite=Lax`; } catch (e) {}
  try { localStorage.setItem(key, JSON.stringify(current)); } catch (e) {}
  return current;
}

async function migrateLegacyPackageRemovals() {
  const ids = getSharedList('deleted_packages');
  for (const id of ids) {
    try {
      await api('/api/packages/' + encodeURIComponent(id), { method: 'DELETE' });
    } catch (e) {
      if (e.status !== 404) throw e;
    }
    removeFromSharedList('deleted_packages', id);
  }
}

// ==========================================================================
// Cross-Portal Synchronization & State Persistence (Vercel Serverless Resilient)
// Prevents requests and packages from disappearing across instances & tabs
// ==========================================================================
const MOBILEUM_SYNC_KEY = 'mobileum_persisted_requests';
const MOBILEUM_STATUS_KEY = 'mobileum_status_overrides';
const MOBILEUM_PACKAGES_KEY = 'mobileum_persisted_packages';
const MOBILEUM_DELETED_PACKAGES_KEY = 'mobileum_deleted_packages';

const syncChannel = typeof BroadcastChannel !== 'undefined' ? new BroadcastChannel('mobileum_sync_channel') : null;

// ---------- Package Persistence & Sync ----------
function getClientPackages() {
  try {
    const raw = localStorage.getItem(MOBILEUM_PACKAGES_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr)) return arr;
    }
  } catch (e) {}
  return [];
}

function getDeletedPackageIds() {
  const ids = new Set();
  try {
    const raw = localStorage.getItem(MOBILEUM_DELETED_PACKAGES_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr)) arr.forEach(id => ids.add(id));
    }
  } catch (e) {}
  const shared = getSharedList('deleted_packages');
  shared.forEach(id => ids.add(id));
  return ids;
}

function saveClientPackage(pkg) {
  if (!pkg || !pkg.packageId) return;
  const list = getClientPackages();
  const idx = list.findIndex(p => p.packageId === pkg.packageId);
  if (idx >= 0) {
    list[idx] = { ...list[idx], ...pkg, status: 'published' };
  } else {
    list.push({ ...pkg, status: 'published' });
  }
  try {
    localStorage.setItem(MOBILEUM_PACKAGES_KEY, JSON.stringify(list));
  } catch (e) {}

  // Remove from deleted list if re-added
  const deleted = Array.from(getDeletedPackageIds()).filter(id => id !== pkg.packageId);
  try {
    localStorage.setItem(MOBILEUM_DELETED_PACKAGES_KEY, JSON.stringify(deleted));
  } catch (e) {}
  removeFromSharedList('deleted_packages', pkg.packageId);

  if (syncChannel) {
    try { syncChannel.postMessage({ type: 'package_update', packageId: pkg.packageId, package: pkg }); } catch (e) {}
  }
  syncWithServer();
}

function removeClientPackage(packageId) {
  if (!packageId) return;
  // Remove from persisted active list
  const list = getClientPackages().filter(p => p.packageId !== packageId);
  try {
    localStorage.setItem(MOBILEUM_PACKAGES_KEY, JSON.stringify(list));
  } catch (e) {}

  // Add to permanently deleted set
  const deleted = getDeletedPackageIds();
  deleted.add(packageId);
  const delArr = Array.from(deleted);
  try {
    localStorage.setItem(MOBILEUM_DELETED_PACKAGES_KEY, JSON.stringify(delArr));
  } catch (e) {}
  addToSharedList('deleted_packages', packageId);

  if (syncChannel) {
    try { syncChannel.postMessage({ type: 'package_delete', packageId }); } catch (e) {}
  }
  syncWithServer();
}

function mergePackages(serverPackages) {
  const serverList = Array.isArray(serverPackages) ? serverPackages : [];
  const clientList = getClientPackages();
  const deletedIds = getDeletedPackageIds();

  const map = new Map();

  // 1. Ingest server packages (skip deleted/removed)
  serverList.forEach(p => {
    if (p && p.packageId && !deletedIds.has(p.packageId) && p.status !== 'removed') {
      map.set(p.packageId, { ...p });
    }
  });

  // 2. Ingest client-persisted packages (ensures plans survive cold starts & refreshes)
  let missingFromServer = false;
  clientList.forEach(p => {
    if (!p || !p.packageId || deletedIds.has(p.packageId) || p.status === 'removed') return;
    if (!map.has(p.packageId)) {
      map.set(p.packageId, { ...p });
      missingFromServer = true;
    } else {
      map.set(p.packageId, { ...map.get(p.packageId), ...p });
    }
  });

  // Sync client packages to server if server was missing any
  if (missingFromServer) {
    syncWithServer();
  }

  return Array.from(map.values());
}

// ---------- Request Persistence & Sync ----------
function getClientRequests() {
  try {
    const raw = localStorage.getItem(MOBILEUM_SYNC_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr)) return arr;
    }
  } catch (e) {}
  return [];
}

function saveClientRequest(req) {
  if (!req || !req.requestId) return;
  const list = getClientRequests();
  const idx = list.findIndex(r => r.requestId === req.requestId);
  if (idx >= 0) {
    list[idx] = { ...list[idx], ...req };
  } else {
    list.push(req);
  }
  try {
    localStorage.setItem(MOBILEUM_SYNC_KEY, JSON.stringify(list));
  } catch (e) {}

  if (syncChannel) {
    try { syncChannel.postMessage({ type: 'sync_update', requestId: req.requestId }); } catch (e) {}
  }
  syncWithServer();
}

function getStatusOverrides() {
  try {
    const raw = localStorage.getItem(MOBILEUM_STATUS_KEY);
    if (raw) {
      const obj = JSON.parse(raw);
      if (obj && typeof obj === 'object') return obj;
    }
  } catch (e) {}
  return {};
}

function saveStatusOverride(requestId, status) {
  if (!requestId || !status) return;
  const overrides = getStatusOverrides();
  overrides[requestId] = status;
  try {
    localStorage.setItem(MOBILEUM_STATUS_KEY, JSON.stringify(overrides));
  } catch (e) {}

  const list = getClientRequests();
  const req = list.find(r => r.requestId === requestId);
  if (req) {
    req.status = status;
    if (req.numbers) req.numbers.forEach(n => { n.status = status; });
    try { localStorage.setItem(MOBILEUM_SYNC_KEY, JSON.stringify(list)); } catch (e) {}
  }

  if (syncChannel) {
    try { syncChannel.postMessage({ type: 'status_override', requestId, status }); } catch (e) {}
  }
  syncWithServer();
}

function mergeRequests(serverRequests) {
  const serverList = Array.isArray(serverRequests) ? serverRequests : [];
  const clientList = getClientRequests();
  const overrides = getStatusOverrides();

  const map = new Map();

  // 1. Ingest server requests
  serverList.forEach(r => {
    if (r && r.requestId) {
      const copy = { ...r };
      if (overrides[copy.requestId]) {
        copy.status = overrides[copy.requestId];
        if (copy.numbers) copy.numbers.forEach(n => { n.status = overrides[copy.requestId]; });
      }
      map.set(copy.requestId, copy);
    }
  });

  // 2. Ingest client requests (preserves newly submitted requests across instances)
  let missingFromServer = false;
  clientList.forEach(clientReq => {
    if (!clientReq || !clientReq.requestId) return;
    if (!map.has(clientReq.requestId)) {
      const copy = { ...clientReq };
      if (overrides[copy.requestId]) {
        copy.status = overrides[copy.requestId];
        if (copy.numbers) copy.numbers.forEach(n => { n.status = overrides[copy.requestId]; });
      }
      map.set(copy.requestId, copy);
      missingFromServer = true;
    } else {
      const existing = map.get(clientReq.requestId);
      if (!existing.logoUrl && clientReq.logoUrl) {
        existing.logoUrl = clientReq.logoUrl;
      }
      if (overrides[clientReq.requestId]) {
        existing.status = overrides[clientReq.requestId];
        if (existing.numbers) existing.numbers.forEach(n => { n.status = overrides[clientReq.requestId]; });
      }
    }
  });

  if (missingFromServer) {
    syncWithServer();
  }

  return Array.from(map.values());
}

function mergeNumbers(serverNumbers) {
  const list = Array.isArray(serverNumbers) ? serverNumbers.slice() : [];
  const clientRequests = getClientRequests();
  const overrides = getStatusOverrides();
  const deletedNums = getSharedList('deleted_numbers');

  clientRequests.forEach(r => {
    const status = overrides[r.requestId] || r.status;
    if (status === 'approved' && Array.isArray(r.numbers)) {
      r.numbers.forEach(n => {
        if (!list.some(x => x.phoneE164 === n.phoneE164) && !deletedNums.includes(n.phoneE164)) {
          list.push({
            phoneE164: n.phoneE164,
            phoneDisplay: n.phoneDisplay,
            department: n.department,
            brandName: r.brandName,
            status: 'approved'
          });
        }
      });
    }
  });

  return list;
}

function lookupLocalNumber(phoneInput) {
  const clientRequests = getClientRequests();
  const overrides = getStatusOverrides();
  const norm = typeof PH !== 'undefined' ? PH.normalize(phoneInput) : null;
  const e164 = norm && norm.ok ? norm.e164 : String(phoneInput || '').replace(/[^0-9+]/g, '');

  for (const r of clientRequests) {
    const status = overrides[r.requestId] || r.status;
    if (status === 'approved' && Array.isArray(r.numbers)) {
      const numHit = r.numbers.find(n => n.phoneE164 === e164);
      if (numHit) {
        return {
          registered: true,
          verified: true,
          brandName: r.brandName,
          tagline: r.tagline || '',
          department: numHit.department,
          display: numHit.phoneDisplay || (norm && norm.display) || e164,
          logoUrl: r.logoUrl || null,
          numberType: numHit.numberType || (norm && norm.type) || 'Phone'
        };
      }
    }
  }
  return null;
}

let syncTimer = null;
function syncWithServer() {
  clearTimeout(syncTimer);
  syncTimer = setTimeout(async () => {
    try {
      const requests = getClientRequests();
      const statusOverrides = getStatusOverrides();
      const packages = getClientPackages();
      const deletedPackageIds = Array.from(getDeletedPackageIds());
      if (!requests.length && !Object.keys(statusOverrides).length && !packages.length && !deletedPackageIds.length) return;
      await api('/api/sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ requests, statusOverrides, packages, deletedPackageIds })
      });
    } catch (e) {
      // Non-blocking sync attempt
    }
  }, 350);
}
