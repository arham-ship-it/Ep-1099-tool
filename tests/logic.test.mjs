// Run: npm test
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { mapHeaders, listAccounts, suggestCategory, aggregate, evaluate, rulesFor, S, parseW9Fields, matchByName, fieldKey } from '../src/logic.js';

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
