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
