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

/** 1099 filing due date: Jan 31 after the tax year, moved to Monday when it falls on a weekend. */
export function filingDue(taxYear) {
  const d = new Date(Number(taxYear) + 1, 0, 31);
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
  return d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
}
const taxYearOf = rules => Object.keys(RULES).find(y => RULES[y] === rules) || '';
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
  return /\b(credit|debit|card|cc|visa|master ?card|amex|american express|discover|paypal|venmo|stripe|square|p-?card|divvy|brex|ramp)\b/i.test(String(m || ''));
}
/** Customer-side transactions (revenue, A/R) — never payments to vendors, so they are skipped. */
export const CUSTOMER_TXN = /^(invoice|payment|receive payment|sales receipt|credit memo|deposit|refund receipt|estimate|statement charge|delayed charge|delayed credit|inventory starting value|inventory qty adjust)$/i;
const isCustomerTxn = (r, idx) => idx.txnType !== undefined && CUSTOMER_TXN.test(String(r[idx.txnType] ?? '').trim());
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
  // Balance-sheet and revenue accounts (1xxx-4xxx) are not vendor services, with two exceptions that are
  // usually paid to service providers: capitalised patent work (attorneys) and leasehold improvements (contractors).
  const num = n.match(/^\s*(\d)\d{3,4}\b/);
  if (num && '1234'.includes(num[1])) {
    if (/accumulated|depreciation|amortization/.test(n)) return 'Not Reportable';
    if (/patent|trademark/.test(n)) return 'Legal Services';
    if (/leasehold improvement/.test(n)) return 'Services';
    return 'Not Reportable';
  }
  if (/settlement/.test(n)) return '';
  if (/legal|attorney|law firm|lawyer|litigation|patent/.test(n)) return 'Legal Services';
  // payroll, benefits and wages are never 1099 (they go on W-2)
  if (/salar|wage|payroll|benefit|401k|401\(k\)|pto|vacation|sick|bonus|workers comp|stock equity|equity comp|\bpfl\b|\bsui\b|\bsuta\b|\bfuta\b|\bfica\b|medicare|social security|unemployment|withholding|employer tax/.test(n)) return 'Not Reportable';
  if (/contractor|subcontract|consult|freelanc|professional fee|accounting|bookkeep|taxation service|tax service|tax prep|audit/.test(n)) return 'Services';
  if (/install|repair|maintenance|testing|certification|engineering|recruit|staffing|affiliate|commission|referral|website|web dev|photograph|design|cleaning|janitor|landscap|security guard|it support|labor/.test(n)) return 'Services';
  if (/advertis/.test(n) && !/paid media/.test(n)) return 'Services';
  if (/sponsor|conference|event|trade ?show|paid media|hubspot|shopify|affirm|selling fee|merchant|platform fee|customer compensation|charit|donation/.test(n)) return 'Not Reportable';
  if (/equipment rental|\brent|lease/.test(n)) return 'Rent';
  if (/equipment|hardware|meter|controller|batter|thermostat|\baws\b|hosting|deposit|discount|refund|bad debt|exchange|gain or loss|incentive|cost of sales|\bdivvy\b|\bbrex\b|\bramp\b/.test(n)) return 'Not Reportable';
  if (/royalt/.test(n)) return 'Royalties';
  if (/medical|health ?care|clinic|physician|doctor/.test(n)) return 'Medical';
  if (/prize|award/.test(n)) return 'Other Income';
  if (/payable|receivable|checking|savings|\bbank\b|undeposited|clearing|sales\b|revenue|income|equity|retained|capital|loan|note payable|accrued|deferred|prepaid|credit card|owner|draw|distribution|opening balance|fixed asset|accumulated|^\d*\s*(cash|petty cash)\b/.test(n)) return 'Not Reportable';
  if (/cogs|cost of goods|inventory|freight|shipping|postage|utilit|electric|telephone|phone|internet|travel|meal|airfare|hotel|software|subscription|dues|licen[sc]e|saas|interest|insurance|tax|depreciation|amortization|supplies|materials|collateral|tooling|scrap|landed cost|reimburse/.test(n)) return 'Not Reportable';
  if (/service|marketing/.test(n)) return 'Services';
  return '';
}

/* ---------------- double-entry GL helpers (cash basis) ---------------- */
/** 'bank' = cash accounts payments leave from, 'card' = credit-card accounts, 'ap' = accounts payable, else 'other'. */
export function acctKind(name) {
  const n = String(name || '').toLowerCase();
  if (!n) return 'other';
  if (/accounts payable|\(a\/p\)|\ba\/p\b/.test(n)) return 'ap';
  if (/credit card|\bamex\b|american express|\bvisa\b|master ?card|\bdivvy\b|\bcard\b|ramp card|brex card/.test(n)) return 'card';
  if (/charges|fees|interest|in transit|supplier deposit|security deposit/.test(n)) return 'other';
  if (/checking|savings|money market|\bbank\b|petty cash|^\d*\s*cash\b|treasury|vault|money out clearing|bill\.com.*clearing|operating account/.test(n)) return 'bank';
  return 'other';
}
export const PRIOR_BILLS = 'bills dated before this GL (not in the file)';
const SKIP_TXN = /journal|transfer|^invoice$|^payment$|receive payment|sales receipt|credit memo|inventory|credit card payment|estimate|statement|purchase order/i;
/** True when the file is a full double-entry ledger (has transaction type + split and bank-account lines). */
export function isLedger(rows, idx) {
  if (idx.txnType === undefined || idx.split === undefined || idx.account === undefined) return false;
  let n = 0;
  for (const r of rows) { if (acctKind(r[idx.account]) === 'bank' && /bill payment|expense|check/i.test(String(r[idx.txnType] ?? ''))) { if (++n >= 3) return true; } }
  return false;
}

/**
 * IRS "eyeball test" (Treas. Reg. §1.6049-4(c)(1)(ii)(A), applied to 1099-NEC/MISC by §1.6041-3(p)):
 * a payee whose name contains an unambiguous expression of corporate status (Incorporated, Inc.,
 * Corporation, Corp., P.C. — but NOT "Company" or "Co.") may be treated as a corporation.
 * Also spots likely government / university payees and likely foreign companies.
 */
export function guessEntity(name) {
  const n = ' ' + String(name || '').toLowerCase().replace(/[.,]/g, ' ').replace(/\s+/g, ' ') + ' ';
  if (/ (inc|incorporated|corp|corporation|p c|pc) /.test(n.replace(/-/g, ' '))) return 'corp';
  if (/ (university|college|city of|county of|state of|department of|dept of|authority|commission|school district|u s |united states|internal revenue|franchise tax board|secretary of state|comptroller|public utilit) /.test(n) || / (cdtfa|irs|uspto|dmv) /.test(n)) return 'gov';
  if (/ (ltd|limited|gmbh|s r o|pte|b v|s a|ag|oy|ab|s p a|srl|sarl|kk|pty|co ltd|shenzhen|guangzhou|dongguan|shanghai|hong kong|taiwan) /.test(n)) return 'foreign';
  return '';
}

export const ALIASES = {
  company: ['company', 'entity', 'company name', 'subsidiary', 'legal entity', 'client'],
  date: ['payment date', 'date', 'txn date', 'transaction date', 'posting date', 'check date'],
  vendor: ['vendor', 'vendor name', 'payee', 'name', 'supplier', 'supplier name'],
  txnType: ['transaction type', 'type', 'txn type', 'trans type'],
  account: ['gl account', 'account', 'distribution account', 'account number', 'account #', 'acct', 'account no', 'gl code'],
  accountName: ['account name', 'account description', 'gl account name', 'account title'],
  amount: ['amount', 'net amount', 'paid amount', 'payment amount'],
  debit: ['debit'], credit: ['credit'],
  method: ['payment method', 'method', 'pay method', 'payment type', 'payment mode', 'payment account'],
  split: ['split', 'split account', 'offset account']
};
export const FIELD_LABELS = { company: 'Company', date: 'Payment date', txnType: 'Transaction type', vendor: 'Vendor', account: 'GL account', accountName: 'Account name',
  amount: 'Amount', debit: 'Debit', credit: 'Credit', method: 'Payment method', split: 'Split (paid from / to)' };

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
  if (idx.account !== undefined || idx.accountName !== undefined) {
    // Account is already a column (e.g. QuickBooks "Distribution account"): just drop report heading / total rows.
    const amt = idx.amount ?? idx.debit;
    rows = rows.filter(r => {
      const t = String(r.find(x => String(x ?? '').trim() !== '') ?? '').trim();
      if (/^total\b/i.test(t) || /^(beginning|opening|ending) balance$/i.test(t)) return false;
      const noDate = idx.date !== undefined && String(r[idx.date] ?? '').trim() === '';
      const noAmt = amt !== undefined && String(r[amt] ?? '').trim() === '';
      return !(noDate && noAmt);
    });
    return { headers, rows, companyHint, grouped: false };
  }

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

const txt = (r, i) => i === undefined ? '' : String(r[i] ?? '').trim();
const methodOf = (r, idx) => idx.method !== undefined ? r[idx.method] : (idx.split !== undefined ? r[idx.split] : '');

/** Distinct GL accounts (per company) for the mapping step. */
export function listAccounts(rows, idx, taxYear, defaultCo) {
  const ledger = isLedger(rows, idx);
  const m = {};
  for (const r of rows) {
    if (!txt(r, idx.vendor)) continue; // lines with no payee can't create a 1099
    const acct = txt(r, idx.account), an = txt(r, idx.accountName);
    const type0 = txt(r, idx.txnType);
    if (!inYear(r, idx, taxYear)) {
      // in a full ledger, earlier-year bills are listed too: they say what this year's payments were for
      const ly = idx.date !== undefined ? yearOf(r[idx.date]) : 0;
      if (!(ledger && /^bill$|vendor credit/i.test(type0) && ly && ly < taxYear)) continue;
    }
    if (ledger) {
      const type = type0;
      if (SKIP_TXN.test(type) || /bill payment|deposit/i.test(type)) continue;
      if (acctKind(acct) !== 'other') continue;                         // bank, card and A/P sides
      if (!/^bill$|vendor credit/i.test(type) && acctKind(txt(r, idx.split)) === 'card') continue; // paid by card
    } else if (isCustomerTxn(r, idx) || isCard(methodOf(r, idx))) continue;
    const co = lineCompany(r, idx, defaultCo);
    const k = co + '|' + acct + '|' + an;
    m[k] ??= { co, acct, an, total: 0 }; m[k].total += lineAmount(r, idx);
  }
  return Object.values(m).sort((a, b) => a.co.localeCompare(b.co) || String(a.acct).localeCompare(String(b.acct), undefined, { numeric: true }));
}

/**
 * Roll GL lines up to company -> vendor. map keys are "company|account|accountName" -> category.
 * For a full double-entry ledger this is CASH BASIS: only money that actually left a bank account counts.
 *   - Bill Payment lines on bank accounts  -> paid; split across categories by the vendor's bill mix
 *   - Expense / Check lines paid from bank -> paid, category of the expense account
 *   - anything paid from a credit-card account -> excluded (1099-K)
 *   - journal entries (payroll, accruals), transfers, invoices, customer payments -> ignored
 *   - unpaid bills -> ignored (shown as a note)
 */
export function aggregate(rows, idx, taxYear, defaultCo, map) {
  if (isLedger(rows, idx)) return aggregateLedger(rows, idx, taxYear, defaultCo, map);
  const cos = {}; let skipped = 0;
  for (const r of rows) {
    const vendor = txt(r, idx.vendor); if (!vendor) continue;
    if (!inYear(r, idx, taxYear)) { skipped++; continue; }
    if (isCustomerTxn(r, idx)) continue;
    const amt = lineAmount(r, idx); if (!amt) continue;
    const co = lineCompany(r, idx, defaultCo);
    const key = normalizeVendor(vendor);
    const acct = txt(r, idx.account), an = txt(r, idx.accountName);
    const label = (acct + (an ? ' ' + an : '')).trim() || '(no account)';
    cos[co] ??= {};
    const v = cos[co][key] ??= { n: vendor, k: key, t: 0, x: 0, c: {}, u: 0, ua: [], nr: [] };
    v.t += amt;
    if (isCard(methodOf(r, idx))) { v.x += amt; continue; }
    const cat = map[co + '|' + acct + '|' + an] || '';
    if (!cat) { v.u += amt; if (!v.ua.includes(label)) v.ua.push(label); continue; }
    if (cat === 'Not Reportable' && !v.nr.includes(label)) v.nr.push(label);
    v.c[cat] = (v.c[cat] || 0) + amt;
  }
  return { cos, skipped };
}

function aggregateLedger(rows, idx, taxYear, defaultCo, map) {
  const W = {}; let skipped = 0;
  for (const r of rows) {
    const vendor = txt(r, idx.vendor); if (!vendor) continue;
    const type = txt(r, idx.txnType);
    if (SKIP_TXN.test(type)) continue;
    const isBill = /^bill$|vendor credit/i.test(type);
    const ly = idx.date !== undefined ? yearOf(r[idx.date]) : 0;
    const thisYear = !ly || ly === taxYear;
    // Bills from an earlier year are kept ONLY to learn what the vendor was paid for (a Dec 2025 bill paid
    // in Jan 2026 counts in 2026). Bills from a later year and all other out-of-year lines are skipped.
    if (!thisYear && !(isBill && ly < taxYear)) { skipped++; continue; }
    const amt = lineAmount(r, idx); if (!amt) continue;
    const co = lineCompany(r, idx, defaultCo);
    const key = normalizeVendor(vendor);
    if (key === normalizeVendor(co)) continue;                          // the company itself
    const acct = txt(r, idx.account), an = txt(r, idx.accountName);
    const kind = acctKind(acct);
    const w = (W[co] ??= {})[key] ??= { n: vendor, k: key, bank: 0, cardBP: 0, bill: {}, billLbl: {}, billIn: 0, billPrior: 0, dir: {}, dirLbl: {}, x: 0 };
    if (/bill payment/i.test(type)) {
      if (kind === 'bank') w.bank += -amt; else if (kind === 'card') w.cardBP += Math.abs(amt);
      continue;
    }
    if (kind !== 'other') continue;                                     // bank / card / A-P side of the entry
    const cat = map[co + '|' + acct + '|' + an] || '';
    const label = (acct + (an ? ' ' + an : '')).trim() || '(no account)';
    if (isBill) {
      w.bill[cat] = (w.bill[cat] || 0) + amt; (w.billLbl[cat] ??= new Set()).add(label);
      if (thisYear) w.billIn += amt; else w.billPrior += amt;
      continue;
    }
    if (acctKind(txt(r, idx.split)) === 'card') { w.x += amt; continue; }   // paid on a credit card
    if (/deposit/i.test(type) && (amt > 0 || !cat)) continue;              // money in; only vendor refunds on mapped accounts count
    w.dir[cat] = (w.dir[cat] || 0) + amt; (w.dirLbl[cat] ??= new Set()).add(label);
  }
  const cos = {};
  for (const co in W) for (const key in W[co]) {
    const w = W[co][key];
    const v = { n: w.n, k: w.k, t: 0, x: Math.max(0, w.x + w.cardBP), c: {}, u: 0, ua: [], nr: [], up: 0, cash: 1, pb: 0 };
    const add = (cat, amt, labels) => {
      if (!amt) return;
      if (!cat) { v.u += amt; labels.forEach(l => { if (!v.ua.includes(l)) v.ua.push(l); }); return; }
      if (cat === 'Not Reportable') labels.forEach(l => { if (!v.nr.includes(l)) v.nr.push(l); });
      v.c[cat] = (v.c[cat] || 0) + amt;
    };
    for (const cat in w.dir) add(cat, w.dir[cat], [...w.dirLbl[cat]]);
    // bank bill payments, spread over what the vendor's bills were for
    const billTot = Object.values(w.bill).filter(x => x > 0).reduce((a, b) => a + b, 0);
    if (w.bank > 0.005) {
      if (billTot > 0.005) {
        // payments on bills dated before this file are assumed to be for the same kind of work
        for (const cat in w.bill) if (w.bill[cat] > 0) add(cat, w.bank * w.bill[cat] / billTot, [...w.billLbl[cat]]);
        // payments in the tax year settle the oldest (prior-year) bills first
        v.pb = Math.min(w.bank, Math.max(0, w.billPrior));
        v.up = Math.max(0, w.billIn - Math.max(0, w.bank + w.cardBP - Math.max(0, w.billPrior)));
      } else {
        const dirTot = Object.values(w.dir).reduce((a, b) => a + Math.abs(b), 0);
        if (dirTot > 0.005) { for (const cat in w.dir) add(cat, w.bank * Math.abs(w.dir[cat]) / dirTot, [...w.dirLbl[cat]]); }
        else { add('', w.bank, [PRIOR_BILLS]); v.pb = w.bank; }
      }
    } else if (w.billIn > 0.005) v.up = Math.max(0, w.billIn - w.cardBP);
    for (const cat in v.c) { v.c[cat] = Math.max(0, v.c[cat]); if (!v.c[cat]) delete v.c[cat]; }
    v.u = Math.max(0, v.u);
    v.t = Object.values(v.c).reduce((a, b) => a + b, 0) + v.u + v.x;
    if (v.t < 0.005 && v.up < 0.005) continue;                          // nothing paid, nothing owed (incl. prior-year bills paid last year)
    (cos[co] ??= {})[key] = v;
  }
  return { cos, skipped, cash: true };
}

/** 1099 decision + plain-English reason. v.e = entity type from the W-9. */
export function evaluate(v, rules) {
  // A category picked by hand for money the GL can't explain (e.g. a Dec bill paid in Jan) — v.oc
  if (v.oc && v.u > 0) {
    const c = { ...(v.c || {}) }; c[v.oc] = (c[v.oc] || 0) + v.u;
    v = { ...v, c, u: 0, ua: [], nr: v.oc === 'Not Reportable' ? [...(v.nr || []), 'set by hand'] : v.nr };
  }
  const e = v.e || 'Unknown';
  const unknown = /unknown/i.test(e), exempt = /exempt|government/i.test(e);
  const rsG = v.rs && v.rs.c !== 'low' ? ({ us_corporation: 'corp', foreign: 'foreign', government: 'gov', nonprofit: 'gov' })[v.rs.k] || '' : '';
  const guess = unknown ? (guessEntity(v.n) || rsG) : '';
  const byWeb = unknown && !guessEntity(v.n) && !!rsG;
  const corp = /corporation|c corp|s corp/i.test(e) || guess === 'corp';
  const how = v.cash ? 'paid from the bank' : 'paid by check/ACH/wire';
  const card = v.x > 0 ? ` ${money(v.x)} paid by card excluded (reported on 1099-K).` : '';
  const unpaid = v.up > 0.5 ? ` ${money(v.up)} billed but not yet paid is not counted.` : '';
  const prior = (v.ua || []).includes(PRIOR_BILLS);
  const unm = v.u > 0 ? (prior ? ` ${money(v.u)} paid this year for bills dated before this GL — pick what it was for to include it.` : ` ${money(v.u)} in unmapped accounts (${v.ua.join(', ')}) not counted yet.`) : '';
  const pbNote = v.pb > 0.5 && !prior ? ` Includes ${money(v.pb)} paid this year for bills from last year.` : '';
  if (/foreign/i.test(e)) return { s: S.NOT, fb: '', rep: 0, why: 'Foreign vendor with W-8 on file — no 1099; review for Form 1042-S.' };
  const q = [], below = [], corpOut = []; let rep = 0;
  for (const cat in v.c) {
    const amt = v.c[cat], r = rules[cat]; if (!r) continue; rep += amt;
    const it = { cat, amt, r };
    if (amt >= r.th) (corp && r.corp ? corpOut : q).push(it); else below.push(it);
  }
  const d = it => `${it.cat} ${money(it.amt)} ≥ ${money(it.r.th)} (${it.r.form} box ${it.r.box})`;
  const corpWhy = guess === 'corp' ? (byWeb ? 'online research shows a corporation (confirm with W-9 if in doubt)' : 'the name shows a corporation (Inc./Corp. — IRS name rule)') : `the W-9 shows ${e}`;
  if (q.length) {
    const fb = [...new Set(q.map(x => x.r.form + ' ' + x.r.box))].join('; ');
    const det = q.map(d).join('; ');
    if (exempt || guess === 'gov') return { s: S.NOT, fb: '', rep, why: `${det}, but ${exempt ? 'the W-9 shows an exempt payee' : 'this looks like a government, public or tax-exempt body (exempt)'}.${card}` };
    if (unknown) {
      const hint = guess === 'corp' ? ' Corporation by name, but attorneys and medical providers are reported even when incorporated — get the W-9.'
        : guess === 'foreign' ? ' Looks like a foreign company — ask for Form W-8BEN-E instead of a W-9 (no 1099 if foreign).'
        : ' Entity type unknown — get the W-9 (corporations are exempt for services and rent).';
      return { s: S.W9, fb, rep, why: `${det} ${how}.${hint}${pbNote}${card}${unpaid}${unm}`, guess };
    }
    return { s: S.ELIG, fb, rep, why: `${det} ${how}; W-9 entity: ${e}.${pbNote}${card}${unpaid}${unm}` };
  }
  const minTh = (rules.Services || {}).th || Math.min(...Object.values(rules).filter(r => r.th > 10).map(r => r.th));
  if (guess === 'gov' && !below.length && !corpOut.length) return { s: S.NOT, fb: '', rep, why: `Government, public or tax-exempt body — exempt from 1099 reporting.${card}` };
  const biggest = below.reduce((m, b) => Math.max(m, b.amt), 0);
  if (v.u > 0 && v.u + biggest >= minTh && !corpOut.length) return { s: S.REV, fb: '', rep: rep + v.u, why: prior
      ? `${money(v.u)} ${how} in ${'' + taxYearOf(rules)} for bills dated before this GL (e.g. December bills paid in January), so the GL doesn’t say what they were for. Pick “Paid for” on this row, or upload a GL that starts a few months earlier.${card}`
      : `${money(v.u)} ${how} through GL accounts with no 1099 category (${v.ua.join(', ')}). Set the category and re-import.${card}` };
  if (corpOut.length) return { s: S.NOT, fb: '', rep, why: `${corpOut.map(d).join('; ')}, but ${corpWhy} — corporations are exempt for these payments.${card}` };
  if (below.length) {
    const near = below.filter(b => b.amt >= 0.9 * b.r.th);
    return { s: S.NOT, fb: '', rep, why: `Below threshold: ${below.map(b => `${b.cat} ${money(b.amt)} < ${money(b.r.th)}`).join('; ')}.${near.length ? ' Within 10% of the threshold — check for missing payments.' : ''}${card}${unpaid}` };
  }
  if (v.u > 0) return { s: S.NOT, fb: '', rep: rep + v.u, why: prior
      ? `${money(v.u)} ${how} for bills dated before this GL — below the ${money(minTh)} threshold whatever it was for.${card}`
      : `Only ${money(v.u)} ${how} in uncategorised accounts — below the threshold either way.${card}` };
  if (v.t < 0.005 && v.up > 0) return { s: S.NOT, fb: '', rep: 0, why: `Nothing paid yet this year — ${money(v.up)} billed but unpaid.` };
  if (v.x > 0 && !Object.keys(v.c).length) return { s: S.NOT, fb: '', rep: 0, why: `All ${money(v.x)} paid by card — the processor reports it on 1099-K, not you.` };
  return { s: S.NOT, fb: '', rep: 0, why: `Payments are for non-reportable items (${v.nr.slice(0, 4).join(', ') || 'no reportable accounts'}).${card}` };
}

/** True when a vendor is worth listing: paid in a 1099-type category, paid for something the GL can't
 *  explain (e.g. prior-year bills), or needs review. Payroll, card-only and goods-only vendors are left out. */
export function isRelevant(v, ev) {
  if (ev.s !== S.NOT) return true;
  if ((v.u || 0) > 0.5) return true;
  return Object.keys(v.c || {}).some(c => c !== 'Not Reportable' && v.c[c] > 0);
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
