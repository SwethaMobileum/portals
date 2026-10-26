// Philippines rules shared by the servers and the web pages: +63 phone numbers and Philippine peso (PHP).
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(); else root.PH = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  const VAT_RATE = 0.12;   // Philippine VAT
  const DEPARTMENTS = ['Customer Service', 'Collections', 'Credit Cards', 'Fraud & Security', 'Loans', 'Sales & Marketing', 'Branch Operations'];
  const EXAMPLE = 'Examples: +63 917 123 4567 (mobile), +63 2 8123 4567 (Metro Manila landline), 1800 10 123 4567 (toll-free)';

  // Accepts +63 917 123 4567, 0917-123-4567, 63917..., 917 123 4567, (02) 8123 4567, 02 8123 4567, 1800-10-123-4567
  function normalize(input) {
    let s = String(input == null ? '' : input).trim();
    if (!s) return { ok: false, error: 'Empty number' };
    if (/[^\d\s()+\-.]/.test(s)) return { ok: false, error: 'Use digits only (spaces, dashes and brackets are fine)' };
    const plus = s.startsWith('+');
    s = s.replace(/\D/g, '');
    let n;
    if (plus) { if (!s.startsWith('63')) return { ok: false, error: 'Philippine numbers start with +63' }; n = s.slice(2); }
    else if (s.startsWith('63') && s.length >= 11) n = s.slice(2);
    else if (s.startsWith('0')) n = s.slice(1);
    else n = s;
    if (n.startsWith('0')) return { ok: false, error: 'Remove the extra 0 after +63' };
    let type, display;
    if (/^9\d{9}$/.test(n)) { type = 'Mobile'; display = `+63 ${n.slice(0, 3)} ${n.slice(3, 6)} ${n.slice(6)}`; }
    else if (/^2\d{8}$/.test(n)) { type = 'Landline (Metro Manila)'; display = `+63 2 ${n.slice(1, 5)} ${n.slice(5)}`; }
    else if (/^[3-8]\d{8}$/.test(n)) { type = 'Landline'; display = `+63 ${n.slice(0, 2)} ${n.slice(2, 5)} ${n.slice(5)}`; }
    else if (/^1800\d{7,10}$/.test(n)) { type = 'Toll-free'; display = `+63 1800 ${n.slice(4).replace(/^(\d+)(\d{4})$/, '$1 $2')}`; }
    else if (n[0] === '9') return { ok: false, error: 'Mobile numbers have 10 digits after +63 (e.g. +63 917 123 4567)' };
    else if (n[0] === '2') return { ok: false, error: 'Metro Manila landlines are +63 2 followed by 8 digits' };
    else return { ok: false, error: 'Not a valid Philippine number. ' + EXAMPLE };
    return { ok: true, e164: '+63' + n, display, type };
  }
  const round2 = x => Math.round((Number(x) + Number.EPSILON) * 100) / 100;
  function php(n) { return '₱' + round2(n || 0).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function fee(price) { const f = round2(price || 0), vat = round2(f * VAT_RATE); return { fee: f, vat, total: round2(f + vat) }; }
  return { normalize, php, fee, round2, VAT_RATE, DEPARTMENTS, EXAMPLE };
});
