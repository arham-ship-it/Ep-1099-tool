// Pure 1099 logic — no browser APIs, so it is unit-tested in tests/logic.test.mjs.
// To change thresholds when the law changes, edit RULES (add a new tax year block).

export const RULES_VERIFIED = 'November 8, 2026';

// th = threshold in USD per vendor per year; corp = true means corporations are exempt
export const RULES = {
  2025: {
    'Services': { form: '1099-NEC', box: '1', th: 600, corp: true },
    'Legal Services': { form: '1099-NEC', box: '1', th: 600, corp: false },
    'Rent': { form: '1099-MISC', box: '1', th: 600, corp: true },
    'Other Income': { form: '1099-MISC', box: '3', th: 600, corp: true },
    'Medical': { form: '1099-MISC', box: '6', th: 600, corp: false },
    'Royalties': { form: '1099-MISC', box: '2', th: 10, corp: true },
    'Attorney Gross Proceeds': { form: '1099-MISC', box: '10', th: 600, corp: false }
  },
  // One Big Beautiful Bill Act (P.L. 119-21) §70433 — payments made in 2026.
  // From 2027 the $2,000 is indexed for inflation: add a 2027 block when the IRS publishes it.
  2026: {
    'Services': { form: '1099-NEC', box: '1a', th: 2000, corp: true },
    'Legal Services': { form: '1099-NEC', box: '1a', th: 2000, corp: false },
    'Rent': { form: '1099-MISC', box: '1', th: 2000, corp: true },
    'Other Income': { form: '1099-MISC', box: '3', th: 2000, corp: true },
    'Medical': { form: '1099-MISC', box: '6', th: 2000, corp: false },
    'Royalties': { form: '1099-MISC', box: '2', th: 10, corp: true },
    'Attorney Gross Proceeds': { form: '1099-MISC', box: '10', th: 600, corp: false }
  }
};

export const CATS = ['Services', 'Legal Services', 'Rent', 'Other Income', 'Medical', 'Royalties', 'Attorney Gross Proceeds', 'Not Reportable'];
export const ENTITIES = ['Unknown', 'Individual / Sole proprietor', 'Single-member LLC', 'Partnership', 'LLC - Partnership', 'LLC - C Corp',
  'LLC - S Corp', 'C Corporation', 'S Corporation', 'Trust / Estate', 'Exempt / Government', 'Foreign (W-8)'];
export const S = { ELIG: 'Eligible', W9: 'Needs W-9', REV: 'Needs review', NOT: 'Not eligible' };

export const latestYear = () => Math.max(...Object.keys(RULES).map(Number));
export const rulesFor = y => RULES[y] || RULES[latestYear()];

export function normalizeVendor(name) {
  let s = String(name || '').toLowerCase().replace(/&/g, ' and ');
  s = s.replace(/l\.l\.c\.?/g, ' llc ').replace(/[^a-z0-9 ]+/g, ' ');
  s = s.replace(/\b(llc|l l c|inc|incorporated|corp|corporation|co|company|ltd|limited|pllc|plc|lp|llp|pc|pa|the|dba)\b/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s || String(name || '').toLowerCase().trim();
}
export const fieldKey = k => 'v_' + k.replace(/[^a-z0-9]+/g, '_').slice(0, 80);
export const slug = s => String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60) || 'company';
export const cidFor = co => 'c-' + slug(co);
export const r2 = n => Math.round(n * 100) / 100;
export const money = n => (n < 0 ? '-$' : '$') + Math.abs(Math.round(n)).toLocaleString('en-US');

export function isCard(m) {
  return /\b(credit|debit|card|cc|visa|master ?card|amex|american express|discover|paypal|venmo|stripe|square|p-?card)\b/i.test(String(m || ''));
}
export function toNum(v) {
  if (typeof v === 'number') return v;
  let s = String(v ?? '').trim(); if (!s) return 0;
  const neg = /^\(.*\)$/.test(s) || /^-/.test(s);
  s = s.replace(/[^0-9.]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : (neg ? -n : n);
}
export function yearOf(v) {
  if (v == null || v === '') return 0;
  if (v instanceof Date) return isNaN(v) ? 0 : v.getFullYear();
  if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000)).getUTCFullYear(); // Excel serial
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-\/]/); if (m) return +m[1];
  m = s.match(/(\d{4})\s*$/); if (m) return +m[1];
  m = s.match(/^\d{1,2}[\/-]\d{1,2}[\/-](\d{2})$/); if (m) return 2000 + +m[1];
  return 0;
}
export function suggestCategory(name) {
  const n = String(name || '').toLowerCase(); if (!n) return '';
  if (/settlement/.test(n)) return '';
  if (/legal|attorney|law firm|lawyer|litigation/.test(n)) return 'Legal Services';
  if (/royalt/.test(n)) return 'Royalties';
  if (/\brent|lease/.test(n)) return 'Rent';
  if (/medical|health ?care|clinic|physician|doctor/.test(n)) return 'Medical';
  if (/prize|award/.test(n)) return 'Other Income';
  if (/payable|receivable|checking|savings|\bbank\b|undeposited|clearing|sales\b|revenue|income|equity|retained|capital|loan|note payable|accrued|deferred|prepaid|credit card|owner|draw|distribution|opening balance|fixed asset|accumulated|^\d*\s*(cash|petty cash)\b/.test(n)) return 'Not Reportable';
  if (/cogs|cost of goods|inventory|freight|shipping|postage|utilit|electric|telephone|phone|internet|payroll|wage|salar|benefit|travel|meal|airfare|hotel|software|subscription|saas|bank|merchant|interest|insurance|tax|depreciation|amortization|supplies|materials|equipment purchase|reimburse/.test(n)) return 'Not Reportable';
  if (/consult|contract|professional|freelanc|service|labor|commission|subcontract|design|accounting|bookkeep|audit fee|repair|maintenance|cleaning|janitor|security guard|landscap|it support|marketing|advertising agency|staffing|temp/.test(n)) return 'Services';
  return '';
}

export const ALIASES = {
  company: ['company', 'entity', 'company name', 'subsidiary', 'legal entity', 'client'],
  date: ['payment date', 'date', 'txn date', 'transaction date', 'posting date', 'check date'],
  vendor: ['vendor', 'vendor name', 'payee', 'name', 'supplier', 'supplier name'],
  account: ['gl account', 'account', 'account number', 'account #', 'acct', 'account no', 'gl code'],
  accountName: ['account name', 'account description', 'gl account name', 'account title'],
  amount: ['amount', 'net amount', 'paid amount', 'payment amount'],
  debit: ['debit'], credit: ['credit'],
  method: ['payment method', 'method', 'pay method', 'payment type', 'payment mode', 'split', 'payment account']
};
export const FIELD_LABELS = { company: 'Company', date: 'Payment date', vendor: 'Vendor', account: 'GL account', accountName: 'Account name',
  amount: 'Amount', debit: 'Debit', credit: 'Credit', method: 'Payment method' };

/**
 * Turns a raw sheet (array of rows) into {headers, rows, companyHint}.
 * Handles report-style exports (QuickBooks / Xero "General Ledger", "Transaction Detail by Account")
 * where the account is a heading row above its transactions instead of a column:
 *   "6100 Consulting Fees"            <- account heading row
 *        01/15/2026  Expense  ABC LLC  ...  6,000.00
 *   "Total for 6100 Consulting Fees"  <- skipped
 * The account is copied down into a new "Account (from report)" column.
 */
export function prepareSheet(aoa) {
  const filled = r => r.filter(x => String(x ?? '').trim() !== '').length;
  let h = 0; while (h < aoa.length && filled(aoa[h]) < 3) h++;
  if (h >= aoa.length) return null;
  const titleRows = aoa.slice(0, h).map(r => r.find(x => String(x ?? '').trim())).filter(Boolean).map(String);
  const companyHint = titleRows.find(t => !/general ledger|transaction|report|detail|basis|^\w+ \d{1,2}(,|\s)|\d{4}\s*-|through/i.test(t)) || '';
  let headers = aoa[h].map(x => String(x ?? '').trim());
  let rows = aoa.slice(h + 1).filter(r => filled(r));
  const idx = mapHeaders(headers);
  if (idx.account !== undefined || idx.accountName !== undefined) return { headers, rows, companyHint, grouped: false };

  const dateCol = idx.date, amtCol = idx.amount ?? idx.debit, vendCol = idx.vendor;
  const firstText = r => { const i = r.findIndex(x => String(x ?? '').trim() !== ''); return i < 0 ? null : { i, t: String(r[i]).trim() }; };
  const isHeading = r => {
    const f = firstText(r); if (!f) return false;
    const hasDate = dateCol !== undefined && String(r[dateCol] ?? '').trim() !== '';
    const hasVendor = vendCol !== undefined && vendCol !== f.i && String(r[vendCol] ?? '').trim() !== '';
    return !hasDate && !hasVendor && f.i <= Math.max(0, (dateCol ?? 1) - 0) && !/^total\b/i.test(f.t);
  };
  if (!rows.some(isHeading)) return { headers, rows, companyHint, grouped: false };

  let acct = ''; const out = [];
  for (const r of rows) {
    const f = firstText(r); if (!f) continue;
    if (/^total\b/i.test(f.t) || /^(beginning|opening|ending) balance/i.test(f.t)) continue;
    if (isHeading(r)) { acct = f.t.replace(/\s*\(\d+\)\s*$/, ''); continue; }
    if (amtCol !== undefined && String(r[amtCol] ?? '').trim() === '' && (idx.credit === undefined || String(r[idx.credit] ?? '').trim() === '')) continue;
    out.push([...r, acct]);
  }
  headers = [...headers, 'Account Name'];
  return { headers, rows: out, companyHint, grouped: true };
}

export function mapHeaders(headers) {
  const h = headers.map(x => String(x).trim().toLowerCase()); const idx = {}; const used = new Set();
  for (const k in ALIASES) for (const a of ALIASES[k]) { const i = h.indexOf(a); if (i >= 0 && !used.has(i)) { idx[k] = i; used.add(i); break; } }
  return idx;
}

const lineAmount = (r, idx) => idx.amount !== undefined ? toNum(r[idx.amount]) : toNum(r[idx.debit]) - toNum(r[idx.credit]);
const lineCompany = (r, idx, def) => idx.company !== undefined && String(r[idx.company] ?? '').trim() ? String(r[idx.company]).trim() : def;
const inYear = (r, idx, y) => { if (idx.date === undefined) return true; const ly = yearOf(r[idx.date]); return !ly || ly === y; };

/** Distinct GL accounts (per company) for the mapping step. */
export function listAccounts(rows, idx, taxYear, defaultCo) {
  const m = {};
  for (const r of rows) {
    if (!inYear(r, idx, taxYear)) continue;
    if (idx.method !== undefined && isCard(r[idx.method])) continue;
    const co = lineCompany(r, idx, defaultCo);
    const acct = idx.account !== undefined ? String(r[idx.account] ?? '').trim() : '';
    const an = idx.accountName !== undefined ? String(r[idx.accountName] ?? '').trim() : '';
    const k = co + '|' + acct + '|' + an;
    m[k] ??= { co, acct, an, total: 0 }; m[k].total += lineAmount(r, idx);
  }
  return Object.values(m).sort((a, b) => a.co.localeCompare(b.co) || String(a.acct).localeCompare(String(b.acct), undefined, { numeric: true }));
}

/** Roll GL lines up to company -> vendor. map keys are "company|account|accountName" -> category. */
export function aggregate(rows, idx, taxYear, defaultCo, map) {
  const cos = {}; let skipped = 0;
  for (const r of rows) {
    const vendor = String(r[idx.vendor] ?? '').trim(); if (!vendor) continue;
    if (!inYear(r, idx, taxYear)) { skipped++; continue; }
    const amt = lineAmount(r, idx); if (!amt) continue;
    const co = lineCompany(r, idx, defaultCo);
    const key = normalizeVendor(vendor);
    const acct = idx.account !== undefined ? String(r[idx.account] ?? '').trim() : '';
    const an = idx.accountName !== undefined ? String(r[idx.accountName] ?? '').trim() : '';
    const label = (acct + (an ? ' ' + an : '')).trim() || '(no account)';
    cos[co] ??= {};
    const v = cos[co][key] ??= { n: vendor, k: key, t: 0, x: 0, c: {}, u: 0, ua: [], nr: [] };
    v.t += amt;
    if (idx.method !== undefined && isCard(r[idx.method])) { v.x += amt; continue; }
    const cat = map[co + '|' + acct + '|' + an] || '';
    if (!cat) { v.u += amt; if (!v.ua.includes(label)) v.ua.push(label); continue; }
    if (cat === 'Not Reportable' && !v.nr.includes(label)) v.nr.push(label);
    v.c[cat] = (v.c[cat] || 0) + amt;
  }
  return { cos, skipped };
}

/** 1099 decision + plain-English reason. v.e = entity type from the W-9. */
export function evaluate(v, rules) {
  const e = v.e || 'Unknown';
  const unknown = /unknown/i.test(e), corp = /corporation|c corp|s corp/i.test(e), exempt = /exempt|government/i.test(e);
  const card = v.x > 0 ? ` ${money(v.x)} paid by card/PayPal excluded (the processor reports it on 1099-K).` : '';
  const unm = v.u > 0 ? ` ${money(v.u)} in unmapped accounts (${v.ua.join(', ')}) not counted yet.` : '';
  if (/foreign/i.test(e)) return { s: S.NOT, fb: '', rep: 0, why: 'Foreign vendor with W-8 on file — no 1099; review for Form 1042-S.' };
  const q = [], below = [], corpOut = []; let rep = 0;
  for (const cat in v.c) {
    const amt = v.c[cat], r = rules[cat]; if (!r) continue; rep += amt;
    const it = { cat, amt, r };
    if (amt >= r.th) (corp && r.corp ? corpOut : q).push(it); else below.push(it);
  }
  const d = it => `${it.cat} ${money(it.amt)} ≥ ${money(it.r.th)} threshold (${it.r.form} box ${it.r.box})`;
  if (q.length) {
    const fb = [...new Set(q.map(x => x.r.form + ' ' + x.r.box))].join('; ');
    const det = q.map(d).join('; ');
    if (exempt) return { s: S.NOT, fb: '', rep, why: `${det}, but the W-9 shows an exempt payee.${card}` };
    if (unknown) return { s: S.W9, fb, rep, why: `${det} paid by check/ACH/wire. Entity type unknown — get the W-9 to confirm (corporations are exempt for services and rent).${card}${unm}` };
    return { s: S.ELIG, fb, rep, why: `${det} paid by check/ACH/wire; W-9 entity: ${e}.${card}${unm}` };
  }
  if (v.u > 0) return { s: S.REV, fb: '', rep, why: `${money(v.u)} paid through unmapped GL accounts (${v.ua.join(', ')}). Map them and re-import.${card}` };
  if (corpOut.length) return { s: S.NOT, fb: '', rep, why: `${corpOut.map(d).join('; ')}, but the W-9 shows ${e} — corporations are exempt for these payments.${card}` };
  if (below.length) {
    const near = below.filter(b => b.amt >= 0.9 * b.r.th);
    return { s: S.NOT, fb: '', rep, why: `Below threshold: ${below.map(b => `${b.cat} ${money(b.amt)} < ${money(b.r.th)}`).join('; ')}.${near.length ? ' Within 10% of the threshold — check for missing payments.' : ''}${card}` };
  }
  if (v.x > 0 && !Object.keys(v.c).length) return { s: S.NOT, fb: '', rep: 0, why: `All ${money(v.x)} paid by card/PayPal — the processor reports it on 1099-K, not you.` };
  return { s: S.NOT, fb: '', rep: 0, why: `Payments are for non-reportable items (${v.nr.join(', ') || 'no reportable accounts'}).${card}` };
}

/** W-9 tracking state: rec = received, trk = requested via email tracker, req = to request, na = not needed */
export function w9State(v, ev) {
  if (v.w9 && v.w9.on) return 'rec';
  if (ev.s !== S.ELIG && ev.s !== S.W9) return 'na';
  if (v.rq && v.rq.s === 'tracker') return 'trk';
  return 'req';
}
export function stats(vendors, rules) {
  const st = { paid: 0, elig: 0, needs: 0, rev: 0, rec: 0, trk: 0, req: 0, scope: 0 };
  for (const v of vendors) {
    const ev = evaluate(v, rules); st.paid++;
    if (ev.s === S.ELIG) st.elig++; if (ev.s === S.W9) st.needs++; if (ev.s === S.REV) st.rev++;
    if (ev.s === S.ELIG || ev.s === S.W9) { st.scope++; st[w9State(v, ev)]++; }
  }
  return st;
}
/** Match a file name / name on a W-9 to a vendor. Returns the vendor's field key or null. */
export function matchByName(text, vendors) {
  const t = ' ' + normalizeVendor(String(text).replace(/w-?\s?9/ig, ' ').replace(/_/g, ' ').replace(/\b(19|20)\d{2}\b/g, ' ')) + ' ';
  let best = null, bl = 0;
  for (const v of vendors) { const k = v.k || ''; if (k.length >= 3 && t.includes(' ' + k + ' ') && k.length > bl) { best = v.fk; bl = k.length; } }
  return best;
}

/**
 * Read a fillable IRS W-9 (Rev. March 2024) from its PDF form fields.
 * fields = [{name, value}] from pdf.js getAnnotations(). Best effort — the user confirms.
 */
export function parseW9Fields(fields) {
  const get = re => (fields.find(f => re.test(f.name) && f.value && f.value !== 'Off') || {}).value || '';
  const on = re => fields.some(f => re.test(f.name) && f.value && f.value !== 'Off');
  const legal = get(/f1_01\b|f1_1\[/);
  const biz = get(/f1_02\b|f1_2\[/);
  const llcCls = String(get(/f1_03\b|f1_3\[/)).trim().toUpperCase();
  let entity = 'Unknown';
  if (on(/c1_1\[0\]/)) entity = 'Individual / Sole proprietor';
  else if (on(/c1_1\[1\]/)) entity = 'C Corporation';
  else if (on(/c1_1\[2\]/)) entity = 'S Corporation';
  else if (on(/c1_1\[3\]/)) entity = 'Partnership';
  else if (on(/c1_1\[4\]/)) entity = 'Trust / Estate';
  else if (on(/c1_1\[5\]/)) entity = llcCls === 'C' ? 'LLC - C Corp' : llcCls === 'S' ? 'LLC - S Corp' : llcCls === 'P' ? 'LLC - Partnership' : 'Single-member LLC';
  const digits = ['f1_11', 'f1_12', 'f1_13', 'f1_14', 'f1_15'].map(n => String(get(new RegExp(n + '\\b'))).replace(/\D/g, '')).join('');
  const tin = digits.slice(-4);
  return { legal_name: legal, business_name: biz, entity_type: entity, tin_last4: tin, from: 'form fields' };
}

/** Text-layer fallback: detect a W-8 and pull the line-1 name when fields are flattened. */
export function parseW9Text(text) {
  const t = String(text || '');
  if (/\bW-8\s?BEN/i.test(t)) return { entity_type: 'Foreign (W-8)', is_w9: false };
  return { is_w9: /Form\s*W-9|Request for Taxpayer/i.test(t) };
}
