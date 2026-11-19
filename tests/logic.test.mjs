// Run: npm test
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { mapHeaders, listAccounts, suggestCategory, aggregate, evaluate, rulesFor, S, parseW9Fields, matchByName, fieldKey, isRelevant } from '../src/logic.js';

const wb = XLSX.read(readFileSync(new URL('../sample/sample_GL_2026.csv', import.meta.url), 'utf8'), { type: 'string' });
const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
const idx = mapHeaders(aoa[0]);
const rows = aoa.slice(1);
const map = {};
listAccounts(rows, idx, 2026, 'X').forEach(a => { const c = suggestCategory(a.an); if (c) map[a.co + '|' + a.acct + '|' + a.an] = c; });
const { cos, skipped } = aggregate(rows, idx, 2026, 'X', map);
const R = rulesFor(2026);
const acme = cos['Acme Holdings LLC'], beta = cos['Beta Retail Inc'];
const st = (v, e) => evaluate({ ...v, e }, R).s;

test('splits companies and ignores prior-year lines', () => {
  assert.deepEqual(Object.keys(cos).sort(), ['Acme Holdings LLC', 'Beta Retail Inc']);
  assert.equal(skipped, 1);
});
test('merges vendor name spellings', () => { assert.equal(acme['abc consulting'].c.Services, 14200); });
test('2026 $2,000 threshold', () => {
  assert.equal(st(acme['jane doe design'], 'Individual / Sole proprietor'), S.NOT); // $1,999
  assert.equal(st(acme['abc consulting'], 'Single-member LLC'), S.ELIG);
  assert.equal(st(acme['abc consulting'], 'Unknown'), S.W9);
});
test('corporations exempt except legal and medical', () => {
  assert.equal(st(acme['brightline'], 'C Corporation'), S.NOT);
  assert.equal(st(acme['smith and partners law'], 'LLC - S Corp'), S.ELIG);
  assert.equal(st(acme['dr lee clinic'], 'C Corporation'), S.ELIG);
});
test('card payments, goods, unmapped, refunds, royalties, foreign', () => {
  assert.equal(st(acme['cloud freelancer'], 'Individual / Sole proprietor'), S.NOT);
  assert.equal(st(acme['quickprint'], 'Unknown'), S.NOT);
  assert.equal(st(acme['northwind services'], 'Unknown'), S.REV);
  assert.equal(beta['mike s repair'].c.Services, 1800);
  assert.equal(st(beta['sunrise royalties'], 'Single-member LLC'), S.ELIG); // $50 >= $10
  assert.equal(st(acme['global dev'], 'Foreign (W-8)'), S.NOT);
});
test('2025 uses $600', () => { assert.equal(evaluate({ ...acme['jane doe design'], e: 'Partnership' }, rulesFor(2025)).s, S.ELIG); });
test('W-9 form fields and name matching', () => {
  const x = parseW9Fields([
    { name: 'topmostSubform[0].Page1[0].f1_01[0]', value: 'ABC Consulting LLC' },
    { name: 'topmostSubform[0].Page1[0].Boxes3a-b_ReadOrder[0].c1_1[5]', value: '6' },
    { name: 'topmostSubform[0].Page1[0].Boxes3a-b_ReadOrder[0].f1_03[0]', value: 'P' },
    { name: 'topmostSubform[0].Page1[0].f1_14[0]', value: '12' }, { name: 'topmostSubform[0].Page1[0].f1_15[0]', value: '3456789' }
  ]);
  assert.equal(x.entity_type, 'LLC - Partnership'); assert.equal(x.tin_last4, '6789');
  const vs = Object.values(acme).map(v => ({ ...v, fk: fieldKey(v.k) }));
  assert.equal(matchByName('ABC Consulting_W9_2026', vs), fieldKey('abc consulting'));
});

test('QuickBooks report-style General Ledger (account headings, totals, split column)', async () => {
  const { prepareSheet } = await import('../src/logic.js');
  const aoa = [
    ['Elexity, Inc.'], ['General Ledger'], ['January - December 2026'], [],
    ['', 'Date', 'Transaction Type', 'Num', 'Name', 'Memo/Description', 'Split', 'Amount', 'Balance'],
    ['4000 Sales', '', '', '', '', '', '', '', ''],
    ['', '03/01/2026', 'Invoice', '1001', 'Honda of Surprise (C):Honda of Surprise', '', 'Accounts Receivable (A/R)', '8100', '8100'],
    ['Total for 4000 Sales', '', '', '', '', '', '', '8100', ''],
    ['6100 Consulting Fees', '', '', '', '', '', '', '', ''],
    ['', '', 'Beginning Balance', '', '', '', '', '', '0'],
    ['', '01/15/2026', 'Expense', '', 'ABC Consulting LLC', '', 'Chase Checking', '6000', '6000'],
    ['', '04/15/2026', 'Bill', '', 'ABC Consulting LLC', '', 'Accounts Payable (A/P)', '8200', '14200'],
    ['', '07/20/2026', 'Expense', '', 'Cloud Freelancer', '', 'Amex Card', '5000', '19200'],
    ['Total for 6100 Consulting Fees', '', '', '', '', '', '', '19200', '']
  ];
  const p = prepareSheet(aoa);
  assert.equal(p.grouped, true);
  assert.equal(p.companyHint, 'Elexity, Inc.');
  assert.equal(p.rows.length, 4);
  const idx = mapHeaders(p.headers);
  const map = {};
  listAccounts(p.rows, idx, 2026, p.companyHint).forEach(a => { const c = suggestCategory(a.an); if (c) map[a.co + '|' + a.acct + '|' + a.an] = c; });
  const { cos } = aggregate(p.rows, idx, 2026, p.companyHint, map);
  const v = cos['Elexity, Inc.'];
  assert.equal(v['abc consulting'].c.Services, 14200);
  assert.equal(evaluate({ ...v['abc consulting'], e: 'Unknown' }, rulesFor(2026)).s, S.W9);
  assert.equal(v['cloud freelancer'].x, 5000);   // paid by Amex -> 1099-K
  assert.equal(v['honda of surprise c honda of surprise'], undefined); // customer invoice skipped
});

test('QuickBooks Online GL with "Distribution account" column and revenue accounts', async () => {
  const { prepareSheet } = await import('../src/logic.js');
  const H = ['', 'Distribution account', 'Transaction date', 'Transaction type', 'Num', 'Name', 'Description', 'Split', 'Amount', 'Balance'];
  const aoa = [
    ['Elexity, Inc.'], ['General Ledger'], ['January, 2025-December, 2026'], [], H,
    ['1003 Chase Checking', '', '', '', '', '', '', '', '', ''],
    ['', 'Beginning Balance', '', '', '', '', '', '', '', 1000],
    ['', '1003 Chase Checking', '01/06/2026', 'Payment', '', 'Longroad Energy', '', '1100 Accounts Receivable', 16200, 0],
    ['', '1003 Chase Checking', '02/06/2026', 'Expense', '', 'Sandra Murray', '', '5218 R&D - Contractor', -27500, 0],
    ['Total for 1003 Chase Checking', '', '', '', '', '', '', '', 0, ''],
    ['', '4043 Professional Services', '03/01/2026', 'Invoice', '', 'Longroad Energy', '', '1100 Accounts Receivable', 585000, 0],
    ['', '5218 R&D - Contractor', '02/06/2026', 'Expense', '', 'Sandra Murray', '', '1003 Chase Checking', 27500, 0],
    ['', '5218 R&D - Contractor', '03/06/2026', 'Expense', '', 'Ramp Dev', '', '2160 Divvy', 9000, 0],
    ['', '7170 G&A - Legal Fees', '04/06/2025', 'Bill', '', 'Brix Law LLP', '', '2000 Accounts Payable', 5000, 0]
  ];
  const p = prepareSheet(aoa);
  const idx = mapHeaders(p.headers);
  assert.equal(p.rows.length, 6);
  const accts = listAccounts(p.rows, idx, 2026, p.companyHint);
  const map = {}; accts.forEach(a => { const c = suggestCategory(a.an || a.acct); if (c) map[a.co + '|' + a.acct + '|' + a.an] = c; });
  assert.equal(suggestCategory('4043 Professional Services'), 'Not Reportable');
  assert.equal(suggestCategory('5124 Operations - Electrical Contractor'), 'Services');
  const v = aggregate(p.rows, idx, 2026, p.companyHint, map).cos['Elexity, Inc.'];
  assert.equal(v['longroad energy'], undefined);              // customer never listed
  assert.equal(evaluate({ ...v['sandra murray'], e: 'Unknown' }, rulesFor(2026)).s, S.W9);
  assert.equal(v['sandra murray'].c.Services, 27500);
  assert.equal(v['ramp dev'].x, 9000);                         // paid on Divvy card
  assert.equal(v['brix law'], undefined);                      // 2025 line ignored for 2026
});

test('cash basis: only bank payments count; unpaid bills, payroll journals and card payments do not', async () => {
  const { prepareSheet, isLedger, guessEntity, acctKind } = await import('../src/logic.js');
  const H = ['', 'Distribution account', 'Transaction date', 'Transaction type', 'Num', 'Name', 'Description', 'Split', 'Amount', 'Balance'];
  const AP = '21000 Accounts Payable (A/P)', BANK = '10010 Chase Checking', CARD = '20300 Brex Credit Card';
  const L = (acct, date, type, name, split, amt) => ['', acct, date, type, '', name, '', split, amt, 0];
  const aoa = [['Test Co'], ['General Ledger'], ['January-October, 2026'], [], H,
    // Law firm: billed 300k, only 176k paid from the bank
    L('61255 G&A - Legal Services', '02/01/2026', 'Bill', 'Gish PLLC', AP, 300000), L(AP, '02/01/2026', 'Bill', 'Gish PLLC', '61255 G&A - Legal Services', 300000),
    L(BANK, '03/01/2026', 'Bill Payment (Check)', 'Gish PLLC', AP, -176000), L(AP, '03/01/2026', 'Bill Payment (Check)', 'Gish PLLC', BANK, -176000),
    // Contractor billed but never paid
    L('61160 G&A - Contractors', '09/01/2026', 'Bill', 'Unpaid Person', AP, 9000), L(AP, '09/01/2026', 'Bill', 'Unpaid Person', '61160 G&A - Contractors', 9000),
    // Contractor paid directly from the bank
    L('61160 G&A - Contractors', '04/01/2026', 'Expense', 'Jane Doe', BANK, 5000), L(BANK, '04/01/2026', 'Expense', 'Jane Doe', '61160 G&A - Contractors', -5000),
    // Contractor paid on the credit card
    L('61160 G&A - Contractors', '04/02/2026', 'Expense', 'Card Guy', CARD, 7000), L(CARD, '04/02/2026', 'Expense', 'Card Guy', '61160 G&A - Contractors', 7000),
    // Payroll journal entry and payroll provider
    L('61110 G&A - Salaries & Wages', '04/15/2026', 'Journal Entry', 'Emp Loyee', '', 80000),
    L('29000 Salaries Payable', '04/15/2026', 'Expense', 'Gusto', BANK, -80000), L(BANK, '04/15/2026', 'Expense', 'Gusto', '29000 Salaries Payable', -80000),
    // Corporation by name, paid for services
    L('63121 R&D - Non Recurring Engineering', '05/01/2026', 'Bill', 'E.G.O. North America, Inc.', AP, 30000),
    L(BANK, '05/20/2026', 'Bill Payment (Check)', 'E.G.O. North America, Inc.', AP, -30000)
  ];
  const p = prepareSheet(aoa); const idx = mapHeaders(p.headers);
  assert.equal(isLedger(p.rows, idx), true);
  assert.equal(acctKind(BANK), 'bank'); assert.equal(acctKind(CARD), 'card'); assert.equal(acctKind('61265 G&A - Bank Charges & Fees'), 'other');
  const map = {}; listAccounts(p.rows, idx, 2026, 'Test Co').forEach(a => { const c = suggestCategory(a.an || a.acct); if (c) map[a.co + '|' + a.acct + '|' + a.an] = c; });
  const v = aggregate(p.rows, idx, 2026, 'Test Co', map).cos['Test Co']; const R = rulesFor(2026);
  assert.equal(Math.round(v['gish'].c['Legal Services']), 176000);          // paid, not billed
  assert.equal(Math.round(v['gish'].up), 124000);
  assert.equal(evaluate(v['unpaid person'], R).s, S.NOT);                        // accrual only
  assert.equal(evaluate(v['jane doe'], R).s, S.W9);
  assert.equal(v['card guy'].x, 7000); assert.equal(evaluate(v['card guy'], R).s, S.NOT);
  assert.equal(v['emp loyee'], undefined);                                       // wages never listed
  assert.equal(v['gusto'], undefined);                                           // payroll provider never listed
  assert.equal(guessEntity('E.G.O. North America, Inc.'), 'corp'); assert.equal(guessEntity('Copper Company'), '');
  assert.equal(evaluate(v['e g o north america'], R).s, S.NOT);                  // IRS name rule
  assert.equal(guessEntity('Shenzhen Ace Battery Co., Ltd'), 'foreign'); assert.equal(guessEntity('Oregon State University'), 'gov');
});

test('bills dated last year but paid this year count in this year (cash basis)', async () => {
  const { prepareSheet } = await import('../src/logic.js');
  const H = ['', 'Distribution account', 'Transaction date', 'Transaction type', 'Num', 'Name', 'Description', 'Split', 'Amount', 'Balance'];
  const AP = '2100 Accounts Payable', BANK = '1103 CitiBank Checking-9813';
  const L = (acct, date, type, name, split, amt) => ['', acct, date, type, '', name, '', split, amt, 0];
  const R = rulesFor(2026);
  const ledger = rows => { const p = prepareSheet([['PCT'], ['General Ledger'], ['2025-2026'], [], H, ...rows]); return { p, idx: mapHeaders(p.headers) }; };
  const run = rows => { const { p, idx } = ledger(rows); const map = {};
    listAccounts(p.rows, idx, 2026, 'PCT').forEach(a => { const c = suggestCategory(a.an || a.acct); if (c) map[a.co + '|' + a.acct + '|' + a.an] = c; });
    return aggregate(p.rows, idx, 2026, 'PCT', map).cos['PCT']; };
  const pay = (name, amt, d) => [L(BANK, d || '01/14/2026', 'Bill Payment (Check)', name, AP, -amt), L(AP, d || '01/14/2026', 'Bill Payment (Check)', name, BANK, -amt)];
  const bill = (name, acct, amt, d) => [L(acct, d, 'Bill', name, AP, amt), L(AP, d, 'Bill', name, acct, amt)];
  // A) GL starts Jan 2026: Dec bill is missing — paid amount is shown and flagged, never hidden
  let v = run([...pay('NYMAN PAINTING LLC', 2735), ...pay('Small Painter', 900), L('6131 DC PFL', '01/15/2026', 'Expense', 'Jonathan Kalin', BANK, 300), L(BANK, '01/15/2026', 'Expense', 'Jonathan Kalin', '6131 DC PFL', -300)]);
  let ev = evaluate(v['nyman painting'], R);
  assert.equal(ev.s, S.REV); assert.equal(Math.round(ev.rep), 2735); assert.match(ev.why, /before this GL/);
  const sp = evaluate(v['small painter'], R);
  assert.equal(sp.s, S.NOT); assert.equal(isRelevant(v['small painter'], sp), true);  // listed as Not eligible
  assert.equal(isRelevant(v['jonathan kalin'], evaluate(v['jonathan kalin'], R)), false); // payroll tax, not listed
  // picking "Paid for" by hand classifies it
  assert.equal(evaluate({ ...v['nyman painting'], oc: 'Services' }, R).s, S.W9);
  // B) GL includes Dec 2025: the bill gives the category, the 2026 payment is counted
  v = run([...bill('NYMAN PAINTING LLC', '6120 Repairs & Maintenance', 2735, '12/18/2025'), ...pay('NYMAN PAINTING LLC', 2735), ...pay('Filler A', 10), ...pay('Filler B', 10),
           ...bill('Old Vendor', '6120 Repairs & Maintenance', 5000, '11/01/2025'), ...pay('Old Vendor', 5000, '11/20/2025')]);
  ev = evaluate(v['nyman painting'], R);
  assert.equal(ev.s, S.W9); assert.equal(Math.round(v['nyman painting'].c.Services), 2735); assert.equal(Math.round(v['nyman painting'].up), 0);
  assert.match(ev.why, /bills from last year/);
  assert.equal(v['old vendor'], undefined);   // billed and paid in 2025 — not a 2026 vendor
});
