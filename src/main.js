import './styles.css';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import {
  RULES, RULES_VERIFIED, CATS, ENTITIES, S, FIELD_LABELS, latestYear, rulesFor, normalizeVendor, fieldKey, slug, cidFor, r2, money,
  mapHeaders, listAccounts, aggregate, evaluate, w9State, stats, matchByName, suggestCategory
} from './logic.js';
import { Store } from './store.js';
import { readW9, listDriveFiles, downloadDriveFile, sendToSheet, sheetLinks } from './integrations.js';
import SAMPLE_GL from '../sample/sample_GL_2026.csv?raw';

const SCLASS = { 'Eligible': 's-elig', 'Needs W-9': 's-w9', 'Needs review': 's-rev', 'Not eligible': 's-not' };
const state = { cid: null, filter: 'all', q: '', clientView: false };
const $ = id => document.getElementById(id);

/* ------------------------------- helpers ------------------------------- */
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtDate = iso => { if (!iso) return ''; const d = new Date(iso); return isNaN(d) ? '' : d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }); };
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
function toast(t) { const el = $('toast'); el.textContent = t; el.hidden = false; clearTimeout(toast._t); toast._t = setTimeout(() => el.hidden = true, 4500); }
function closeModal() { $('modal').innerHTML = ''; }
function modal(html, label) { $('modal').innerHTML = `<div class="scrim"><div class="panel" role="dialog" aria-modal="true" aria-label="${esc(label)}">${html}</div></div>`; return $('modal'); }
function download(name, text, type = 'text/plain') {
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type })); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function dropZone(el, cb) {
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('over'); });
  el.addEventListener('dragleave', () => el.classList.remove('over'));
  el.addEventListener('drop', e => { e.preventDefault(); el.classList.remove('over'); cb([...e.dataTransfer.files]); });
}
const vendorList = cid => { const o = Store.vendors(cid); return Object.keys(o).filter(k => k.startsWith('v_')).map(k => ({ ...o[k], fk: k })); };
const sortedClientIds = () => Object.keys(Store.clients()).sort((a, b) => Store.clients()[a].name.localeCompare(Store.clients()[b].name));
const segBar = st => { const sc = Math.max(st.scope, 1);
  return `<div class="bar" role="img" aria-label="${st.rec} of ${st.scope} W-9s received"><i style="width:${st.rec / sc * 100}%;background:var(--ok)"></i><i style="width:${st.trk / sc * 100}%;background:var(--info)"></i><i style="width:${st.req / sc * 100}%;background:var(--bad)"></i></div>`; };

/* --------------------------------- rail --------------------------------- */
function renderRail() {
  const ids = sortedClientIds();
  $('clist').innerHTML = ids.length ? ids.map(id => {
    const c = Store.clients()[id];
    const s = stats(vendorList(id), rulesFor(c.taxYear));
    const pct = s.scope ? Math.round(s.rec / s.scope * 100) : 0;
    return `<button class="citem" type="button" data-cid="${id}" aria-current="${id === state.cid}">
      <span class="cn">${esc(c.name)}</span>
      <span class="cm"><span>${s.scope} need 1099 · ${s.req} to request</span><span class="num">${pct}%</span></span>${segBar(s)}</button>`;
  }).join('') : '<p class="note">No company trackers yet. Upload a GL to create one per company.</p>';
  $('clist').querySelectorAll('.citem').forEach(b => b.onclick = () => select(b.dataset.cid));
}
function select(cid) { state.cid = cid; state.filter = 'all'; state.q = ''; try { localStorage.setItem('ep1099.cid', cid); } catch { } render(); }
function render() {
  if (!state.cid || !Store.clients()[state.cid]) {
    let saved = null; try { saved = localStorage.getItem('ep1099.cid'); } catch { }
    state.cid = (saved && Store.clients()[saved]) ? saved : (sortedClientIds()[0] || null);
  }
  renderRail(); renderMain();
}

/* --------------------------------- main --------------------------------- */
function renderMain() {
  const m = $('main'); const c = state.cid && Store.clients()[state.cid];
  if (!c) { m.innerHTML = emptyHtml(); bindEmpty(); return; }
  const rules = rulesFor(c.taxYear);
  const vs = vendorList(state.cid).map(v => ({ ...v, ev: evaluate(v, rules) }));
  vs.forEach(v => v.w = w9State(v, v.ev));
  const st = stats(vs, rules);
  const pct = st.scope ? Math.round(st.rec / st.scope * 100) : 0;
  const order = { 'Eligible': 0, 'Needs W-9': 1, 'Needs review': 2, 'Not eligible': 3 };
  let rows = [...vs].sort((a, b) => (order[a.ev.s] - order[b.ev.s]) || (b.ev.rep - a.ev.rep));
  const f = state.filter;
  if (f === 'need') rows = rows.filter(v => v.ev.s === S.ELIG || v.ev.s === S.W9);
  if (f === 'req') rows = rows.filter(v => v.w === 'req');
  if (f === 'rec') rows = rows.filter(v => v.w === 'rec');
  if (f === 'rev') rows = rows.filter(v => v.ev.s === S.REV);
  if (f === 'not') rows = rows.filter(v => v.ev.s === S.NOT);
  if (state.q) { const q = state.q.toLowerCase(); rows = rows.filter(v => v.n.toLowerCase().includes(q)); }
  const cv = state.clientView; document.body.classList.toggle('clientview', cv);
  const chip = (k, l, n) => `<button type="button" class="chip" data-f="${k}" aria-pressed="${f === k}">${l} <span class="num">${n}</span></button>`;
  m.innerHTML = `
  <section class="card"><div class="chead">
    <div style="min-width:0"><h1>${esc(c.name)}</h1>
      <div class="meta"><span>Tax year <b>${c.taxYear}</b></span><span>${c.glFile ? 'GL: ' + esc(c.glFile) + ' · imported ' + esc(fmtDate(c.importedAt)) : ''}</span>
        ${c.driveUrl ? `<a href="${esc(c.driveUrl)}" target="_blank" rel="noopener">Open client W-9 folder</a>` : ''}</div>
      <div class="meta internal">
        <label>W-9 Drive folder <input id="cDrive" value="${esc(c.driveUrl || '')}" placeholder="Paste Drive folder link (view-only is fine)"></label>
        <label>Sender name <input id="cSName" value="${esc(c.senderName || '')}" placeholder="Who sends the emails"></label>
        <label>Sender email <input id="cSEmail" value="${esc(c.senderEmail || '')}" placeholder="name@company.com"></label>
      </div></div>
    <div class="actions">
      <button class="ghost internal" id="aW9" type="button">Upload W-9s</button>
      <button class="ghost internal" id="aDrive" type="button">Read W-9s from Drive</button>
      <button class="primary internal" id="aSend" type="button">Send ${st.req} to email tracker</button>
      <button class="ghost" id="aView" type="button">${cv ? 'Back to full tracker' : 'Client view'}</button>
      <button class="ghost internal" id="aCsv" type="button">Export CSV</button>
      <button class="ghost internal" id="aDel" type="button">Delete</button>
    </div></div></section>
  <section class="kpis">
    <div class="kpi"><span class="v num">${st.paid}</span><span class="l">Vendors paid in ${c.taxYear}</span></div>
    <div class="kpi k-elig"><span class="v num">${st.scope}</span><span class="l">Need a 1099 (${st.elig} confirmed, ${st.needs} pending W-9)</span></div>
    <div class="kpi k-rec"><span class="v num">${st.rec}</span><span class="l">W-9s received</span></div>
    <div class="kpi k-req"><span class="v num">${st.req}</span><span class="l">W-9s to request</span></div>
    <div class="kpi k-trk"><span class="v num">${st.trk}</span><span class="l">Requested via email tracker</span></div>
  </section>
  <section class="card progress">
    <div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap"><h3 style="font-size:15px">W-9 collection: ${pct}% complete</h3><span class="note">1099s due to vendors and the IRS by January 31, ${c.taxYear + 1}</span></div>
    ${segBar(st)}
    <div class="legend"><span><i style="background:var(--ok)"></i>Received ${st.rec}</span><span><i style="background:var(--info)"></i>Requested ${st.trk}</span><span><i style="background:var(--bad)"></i>To request ${st.req}</span>${st.rev ? `<span><i style="background:var(--risk)"></i>${st.rev} need review</span>` : ''}</div>
  </section>
  <section style="display:flex;flex-direction:column;gap:10px">
    <div class="filters">${chip('all', 'All', vs.length)}${chip('need', 'Need 1099', st.scope)}${chip('req', 'To request', st.req)}${chip('rec', 'W-9 received', st.rec)}${chip('rev', 'Needs review', st.rev)}${chip('not', 'Not eligible', st.paid - st.scope - st.rev)}
      <input class="search" id="q" placeholder="Search vendor" value="${esc(state.q)}" aria-label="Search vendor"></div>
    <div class="tablewrap"><table>
      <thead><tr><th>Vendor</th><th style="text-align:right">Reportable</th><th>1099 status</th><th>Form / box</th><th>Why</th><th>W-9</th><th class="internal">Entity type (from W-9)</th><th class="internal">Contact</th></tr></thead>
      <tbody>${rows.map(rowHtml).join('') || '<tr><td colspan="8" class="note" style="padding:24px">No vendors match this filter.</td></tr>'}</tbody>
    </table></div>
    <p class="note">Status colors: <span class="pill s-elig">Eligible</span> <span class="pill s-w9">Needs W-9</span> <span class="pill s-rev">Needs review</span> <span class="pill s-not">Not eligible</span> · Rules: tax year ${c.taxYear}${RULES[c.taxYear] ? '' : ' (latest rules — verify the inflation-adjusted threshold)'} · verified ${RULES_VERIFIED}</p>
  </section>`;
  bindMain(c, vs);
}
function rowHtml(v) {
  const ev = v.ev;
  let w;
  if (v.w === 'rec') w = `<span class="pill w-rec">✓ Received</span>${v.w9 && v.w9.f ? `<br><span class="sub">${esc(v.w9.f)}</span>` : ''}`;
  else if (v.w === 'trk') w = `<span class="pill w-trk">Requested</span><br><span class="sub">${esc(fmtDate(v.rq && v.rq.d))}</span>`;
  else if (v.w === 'req') w = '<span class="pill w-req">To request</span>';
  else w = '<span class="pill w-na">Not needed</span>';
  const toggle = (ev.s === S.ELIG || ev.s === S.W9 || v.w === 'rec') ? `<br><button class="link internal" type="button" data-w9toggle="${v.fk}">${v.w === 'rec' ? 'Undo received' : 'Mark received'}</button>` : '';
  return `<tr>
    <td><div class="vname">${esc(v.n)}</div>${v.w9 && v.w9.ln && v.w9.ln.toLowerCase() !== v.n.toLowerCase() ? `<div class="sub">W-9: ${esc(v.w9.ln)}</div>` : ''}${v.w9 && v.w9.l4 ? `<div class="sub">TIN •••${esc(v.w9.l4)}</div>` : ''}</td>
    <td class="amt num">${money(ev.rep)}${v.x ? `<div class="sub">${money(v.x)} card</div>` : ''}</td>
    <td><span class="pill ${SCLASS[ev.s]}">${ev.s}</span></td>
    <td class="num">${esc(ev.fb) || '—'}</td>
    <td class="reason">${esc(ev.why)}</td>
    <td>${w}${toggle}</td>
    <td class="internal"><select data-ent="${v.fk}" aria-label="Entity type for ${esc(v.n)}">${ENTITIES.map(e => `<option ${((v.e || 'Unknown') === e) ? 'selected' : ''}>${e}</option>`).join('')}</select></td>
    <td class="internal" style="min-width:190px"><input data-cn="${v.fk}" placeholder="Contact name" value="${esc(v.cn || '')}" aria-label="Contact name"><input data-ce="${v.fk}" placeholder="email@vendor.com" value="${esc(v.ce || '')}" style="margin-top:4px" aria-label="Contact email"></td>
  </tr>`;
}
function bindMain(c, vs) {
  const m = $('main'); const cid = state.cid;
  m.querySelectorAll('.chip').forEach(b => b.onclick = () => { state.filter = b.dataset.f; renderMain(); });
  const q = m.querySelector('#q'); q.oninput = debounce(() => { state.q = q.value; renderMain(); const nq = $('q'); nq.focus(); nq.setSelectionRange(nq.value.length, nq.value.length); }, 250);
  m.querySelectorAll('[data-ent]').forEach(s => s.onchange = () => Store.patchVendor(cid, s.dataset.ent, { e: s.value }));
  m.querySelectorAll('[data-cn]').forEach(i => i.onchange = () => Store.patchVendor(cid, i.dataset.cn, { cn: i.value.trim() }));
  m.querySelectorAll('[data-ce]').forEach(i => i.onchange = () => Store.patchVendor(cid, i.dataset.ce, { ce: i.value.trim() }));
  m.querySelectorAll('[data-w9toggle]').forEach(b => b.onclick = () => { const v = Store.vendors(cid)[b.dataset.w9toggle] || {};
    Store.patchVendor(cid, b.dataset.w9toggle, { w9: { ...(v.w9 || {}), on: !(v.w9 && v.w9.on), d: new Date().toISOString(), src: (v.w9 && v.w9.src) || 'manual' } }); });
  [['cDrive', 'driveUrl'], ['cSName', 'senderName'], ['cSEmail', 'senderEmail']].forEach(([id, f]) => { const el = m.querySelector('#' + id); el.onchange = () => Store.putClient(cid, { [f]: el.value.trim() }); });
  m.querySelector('#aW9').onclick = () => $('w9File').click();
  m.querySelector('#aDrive').onclick = () => readFromDrive(c);
  m.querySelector('#aSend').onclick = () => openSend(c, vs);
  m.querySelector('#aView').onclick = () => { state.clientView = !state.clientView; renderMain(); };
  m.querySelector('#aCsv').onclick = () => exportCsv(c, vs);
  m.querySelector('#aDel').onclick = () => {
    const md = modal(`<h2>Delete the ${esc(c.name)} tracker?</h2><p class="note">This removes its vendors, W-9 status and contacts from this browser.</p><div class="pfoot"><button class="ghost" id="cNo" type="button">Cancel</button><button class="primary" id="cYes" type="button" style="background:var(--bad);color:#fff">Delete tracker</button></div>`, 'Delete tracker');
    md.querySelector('#cNo').onclick = closeModal; md.querySelector('#cYes').onclick = () => { closeModal(); Store.deleteClient(cid); state.cid = null; toast('Tracker deleted'); };
  };
}
function emptyHtml() {
  return `<section class="card empty"><h1 style="font-size:26px">Find your 1099 vendors in minutes</h1>
    <p>Upload a general ledger or AP payment export (.xlsx or .csv). The tool totals payments per vendor, applies the IRS ${latestYear()} rules ($2,000 for 1099-NEC and most 1099-MISC), and builds one color-coded tracker for each company in the file.</p>
    <div class="drop" id="dropGl" role="button" tabindex="0">Drop a GL file here, or click to choose</div>
    <div class="actions" style="justify-content:center"><button class="ghost" id="trySample" type="button">Try it with a sample GL</button></div>
    <p class="note">Next: read W-9s from the client's Drive folder (or upload them), then send vendors that still need a W-9 to your Google Sheet email tracker.</p></section>`;
}
function bindEmpty() {
  const d = $('dropGl'); d.onclick = () => $('glFile').click();
  d.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); d.click(); } };
  dropZone(d, files => files[0] && readGl(files[0]));
  $('trySample').onclick = () => { const wb = XLSX.read(SAMPLE_GL, { type: 'string', cellDates: true }); startImport(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' }), 'sample_GL_2026.csv'); };
}

/* ------------------------------ GL import ------------------------------ */
let imp = null;
async function readGl(file) {
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
    startImport(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' }), file.name);
  } catch (err) { toast('Couldn’t read that file: ' + err.message); }
}
function startImport(aoa, name) {
  let h = 0; while (h < aoa.length && aoa[h].filter(x => String(x).trim()).length < 3) h++;
  if (h >= aoa.length) { toast('No header row found in that file.'); return; }
  const headers = aoa[h].map(x => String(x).trim());
  imp = { name, headers, rows: aoa.slice(h + 1).filter(r => r.some(x => String(x).trim())), idx: mapHeaders(headers), taxYear: latestYear(), defaultCo: '', step: 1 };
  renderImport();
}
function renderImport() {
  const steps = `<div class="steps"><span class="${imp.step === 1 ? 'on' : ''}">1 · Columns</span><span class="${imp.step === 2 ? 'on' : ''}">2 · GL accounts</span><span class="${imp.step === 3 ? 'on' : ''}">3 · Results</span></div>`;
  let body = '';
  if (imp.step === 1) {
    const opt = k => `<select id="m_${k}"><option value="">— not in file —</option>${imp.headers.map((h, i) => `<option value="${i}" ${imp.idx[k] === i ? 'selected' : ''}>${esc(h)}</option>`).join('')}</select>`;
    body = `<p class="note"><b>${esc(imp.name)}</b> · ${imp.rows.length.toLocaleString()} lines. Check that each field points at the right column.</p>
      <div class="grid2">${Object.keys(FIELD_LABELS).map(k => `<label class="field">${FIELD_LABELS[k]}${k === 'vendor' ? ' *' : ''}${opt(k)}</label>`).join('')}
      <label class="field">Tax year<select id="m_year">${Object.keys(RULES).map(y => `<option ${+y === imp.taxYear ? 'selected' : ''}>${y}</option>`).join('')}<option>${latestYear() + 1}</option></select></label>
      <label class="field">Company name (if no Company column)<input id="m_co" value="${esc(imp.defaultCo)}" placeholder="e.g. Acme Holdings LLC"></label></div>
      ${imp.idx.method === undefined ? '<p class="warnbox">No payment method column found. Card and PayPal payments can’t be excluded — use the AP payment register if you can.</p>' : ''}`;
  } else if (imp.step === 2) {
    body = `<p class="note">Pick the 1099 category for each GL account. Suggestions come from the account name — review every row, especially blanks. Saved per company for next time.</p>
      <div class="tablewrap" style="max-height:55vh;overflow:auto"><table style="min-width:640px"><thead><tr><th>Company</th><th>Account</th><th style="text-align:right">Total</th><th>1099 category</th></tr></thead><tbody>
      ${imp.accounts.map((a, i) => `<tr><td>${esc(a.co)}</td><td>${esc(a.acct)} ${esc(a.an)}</td><td class="amt num">${money(a.total)}</td><td><select data-ai="${i}" aria-label="Category"><option value="">— needs mapping —</option>${CATS.map(c => `<option ${a.cat === c ? 'selected' : ''}>${c}</option>`).join('')}</select>${a.src ? `<div class="sub">${a.src}</div>` : ''}</td></tr>`).join('')}
      </tbody></table></div>`;
  } else body = `<div class="okbox">${imp.summary}</div><p class="note">Next: open each company, read or upload W-9s, add vendor contacts, then send the rest to the email tracker.</p>`;
  const md = modal(`<div class="phead"><h2>Import general ledger</h2><button class="ghost" id="mClose" type="button">Close</button></div>${steps}${body}
    <div class="pfoot">${imp.step === 2 ? '<button class="ghost" id="mBack" type="button">Back</button>' : ''}${imp.step < 3 ? `<button class="primary" id="mNext" type="button">${imp.step === 1 ? 'Next: GL accounts' : 'Run 1099 check'}</button>` : '<button class="primary" id="mDone" type="button">Open trackers</button>'}</div>`, 'Import GL');
  md.querySelector('#mClose').onclick = closeModal;
  if (imp.step === 1) md.querySelector('#mNext').onclick = () => {
    const idx = {}; Object.keys(FIELD_LABELS).forEach(k => { const v = md.querySelector('#m_' + k).value; if (v !== '') idx[k] = +v; });
    if (idx.vendor === undefined) return toast('Choose the Vendor column.');
    if (idx.amount === undefined && idx.debit === undefined) return toast('Choose the Amount (or Debit) column.');
    imp.idx = idx; imp.taxYear = +md.querySelector('#m_year').value; imp.defaultCo = md.querySelector('#m_co').value.trim() || imp.name.replace(/\.[^.]+$/, '');
    imp.accounts = listAccounts(imp.rows, imp.idx, imp.taxYear, imp.defaultCo);
    for (const a of imp.accounts) {
      const saved = (Store.clients()[cidFor(a.co)] || {}).accountMap || {};
      const sk = a.acct + '|' + a.an;
      if (saved[sk]) { a.cat = saved[sk]; a.src = 'Saved mapping'; } else { a.cat = suggestCategory(a.an || a.acct); a.src = a.cat ? 'Suggested — check' : ''; }
    }
    imp.step = 2; renderImport();
  };
  if (imp.step === 2) {
    md.querySelectorAll('[data-ai]').forEach(s => s.onchange = () => { const a = imp.accounts[+s.dataset.ai]; a.cat = s.value; a.src = 'Set by you'; });
    md.querySelector('#mBack').onclick = () => { imp.step = 1; renderImport(); };
    md.querySelector('#mNext').onclick = runImport;
  }
  if (imp.step === 3) md.querySelector('#mDone').onclick = () => { closeModal(); if (imp.firstCid) select(imp.firstCid); };
}
function runImport() {
  const map = {}; imp.accounts.forEach(a => { if (a.cat) map[a.co + '|' + a.acct + '|' + a.an] = a.cat; });
  const { cos, skipped } = aggregate(imp.rows, imp.idx, imp.taxYear, imp.defaultCo, map);
  const lines = []; imp.firstCid = null;
  for (const co of Object.keys(cos).sort()) {
    const cid = cidFor(co); imp.firstCid ??= cid;
    const old = Store.vendors(cid); const out = {};
    for (const key in cos[co]) {
      const a = cos[co][key]; const fk = fieldKey(key); const o = old[fk] || {};
      out[fk] = { n: a.n, k: a.k, t: r2(a.t), x: r2(a.x), c: Object.fromEntries(Object.entries(a.c).map(([k, v]) => [k, r2(v)])), u: r2(a.u), ua: a.ua, nr: a.nr,
        e: o.e || 'Unknown', cn: o.cn || '', ce: o.ce || '', w9: o.w9 || null, rq: o.rq || null };
    }
    for (const fk in old) if (fk.startsWith('v_') && !out[fk] && old[fk].w9 && old[fk].w9.on) out[fk] = { ...old[fk], t: 0, x: 0, c: {}, u: 0, ua: [], nr: [] };
    const prev = Store.clients()[cid] || {};
    const accountMap = { ...(prev.accountMap || {}) };
    imp.accounts.filter(a => a.co === co && a.cat).forEach(a => accountMap[a.acct + '|' + a.an] = a.cat);
    Store.putVendors(cid, out);
    Store.putClient(cid, { name: co, taxYear: imp.taxYear, glFile: imp.name, importedAt: new Date().toISOString(), accountMap,
      driveUrl: prev.driveUrl || '', senderName: prev.senderName || '', senderEmail: prev.senderEmail || '' });
    const st = stats(Object.values(out), rulesFor(imp.taxYear));
    lines.push(`<b>${esc(co)}</b>: ${st.paid} vendors · ${st.scope} need a 1099 · ${st.rev} need review`);
  }
  imp.summary = lines.join('<br>') + (skipped ? `<br>${skipped} lines outside ${imp.taxYear} were ignored.` : '');
  imp.step = 3; renderImport();
}

/* ------------------------------- W-9s ------------------------------- */
let w9q = [];
async function expandFiles(files) {
  const list = [];
  for (const f of files) {
    if (/\.zip$/i.test(f.name)) {
      const z = await JSZip.loadAsync(await f.arrayBuffer());
      for (const name of Object.keys(z.files)) {
        const zf = z.files[name]; if (zf.dir || !/\.(pdf|png|jpe?g)$/i.test(name)) continue;
        const base = name.split('/').pop();
        list.push(new File([await zf.async('blob')], base, { type: /pdf$/i.test(base) ? 'application/pdf' : (/png$/i.test(base) ? 'image/png' : 'image/jpeg') }));
      }
    } else list.push(f);
  }
  return list;
}
async function processW9s(files) {
  if (!files.length) return toast('No PDF or image files found.');
  const vs = vendorList(state.cid);
  w9q = files.map(f => ({ f, match: matchByName(f.name.replace(/\.[^.]+$/, ''), vs), x: null, st: 'Waiting' }));
  renderW9();
  for (const it of w9q) {
    it.st = 'Reading…'; renderW9();
    const r = await readW9(it.f, Store.settings());
    it.x = r.x; it.note = r.note;
    if (!it.match && it.x) it.match = matchByName(it.x.legal_name || '', vs) || matchByName(it.x.business_name || '', vs);
    it.st = 'Ready'; renderW9();
  }
}
async function readFromDrive(c) {
  if (!c.driveUrl) return toast('Paste the client’s W-9 Drive folder link first.');
  const s = Store.settings();
  if (!s.googleClientId) return toast('Add your Google OAuth Client ID in Settings first (see README).');
  try {
    toast('Signing in to Google…');
    const files = await listDriveFiles(c.driveUrl, s.googleClientId);
    if (!files.length) return toast('No PDF or image files in that folder.');
    toast(`Downloading ${files.length} files…`);
    const out = [];
    for (const f of files) out.push(await downloadDriveFile(f, s.googleClientId));
    processW9s(out);
  } catch (e) { toast(e.message); }
}
function renderW9() {
  const vs = vendorList(state.cid).sort((a, b) => a.n.localeCompare(b.n));
  const ready = w9q.every(i => i.st === 'Ready');
  const md = modal(`<div class="phead"><h2>W-9s for ${esc(Store.clients()[state.cid].name)}</h2><button class="ghost" id="wClose" type="button">Close</button></div>
    <p class="note">Each file is matched to a vendor by file name, then by the name on the form. Check the match and entity type, then save. Only the last 4 digits of the TIN are kept.</p>
    <div class="tablewrap" style="max-height:60vh;overflow:auto"><table style="min-width:760px"><thead><tr><th>File</th><th>Read from form</th><th>Vendor</th><th>Entity type</th></tr></thead><tbody>
    ${w9q.map((it, i) => `<tr><td>${esc(it.f.name)}<div class="sub">${it.st}</div></td>
      <td class="reason">${it.x ? `${esc(it.x.legal_name || '')}${it.x.business_name ? ' / ' + esc(it.x.business_name) : ''}${it.x.tin_last4 ? '<br>TIN •••' + esc(it.x.tin_last4) : ''}${it.x.from ? `<br><span class="sub">via ${esc(it.x.from)}</span>` : ''}${it.x.is_w9 === false ? '<br><b>This may not be a W-9.</b>' : ''}` : esc(it.note || '—')}</td>
      <td><select data-wm="${i}" aria-label="Vendor"><option value="">— skip this file —</option>${vs.map(v => `<option value="${v.fk}" ${it.match === v.fk ? 'selected' : ''}>${esc(v.n)}</option>`).join('')}</select></td>
      <td><select data-we="${i}" aria-label="Entity">${ENTITIES.map(e => `<option ${((it.x && it.x.entity_type) || 'Unknown') === e ? 'selected' : ''}>${e}</option>`).join('')}</select></td></tr>`).join('')}
    </tbody></table></div>
    <div class="pfoot"><button class="primary" id="wSave" type="button" ${ready ? '' : 'disabled'}>${ready ? 'Save W-9s' : 'Reading files…'}</button></div>`, 'Upload W-9s');
  md.querySelector('#wClose').onclick = closeModal;
  md.querySelectorAll('[data-wm]').forEach(s => s.onchange = () => w9q[+s.dataset.wm].match = s.value);
  md.querySelectorAll('[data-we]').forEach(s => s.onchange = () => { const it = w9q[+s.dataset.we]; it.x = { ...(it.x || {}), entity_type: s.value }; });
  md.querySelector('#wSave').onclick = () => {
    let n = 0;
    for (const it of w9q) {
      if (!it.match) continue; const x = it.x || {};
      Store.patchVendor(state.cid, it.match, { e: x.entity_type || 'Unknown', w9: { on: true, f: it.f.name, d: new Date().toISOString(), ln: x.legal_name || '', l4: String(x.tin_last4 || '').slice(-4), src: 'upload' } });
      n++;
    }
    closeModal(); toast(`${n} W-9${n === 1 ? '' : 's'} saved`);
  };
}

/* --------------------------- email tracker --------------------------- */
function openSend(c, vs) {
  const rows = vs.filter(v => v.w === 'req');
  const s = Store.settings();
  const payload = rows.map(v => [v.n, v.cn || '', v.ce || '', r2(v.ev.rep), v.ev.fb, v.ev.s]);
  const sender = [c.senderName || '', c.senderEmail || ''];
  const missing = rows.filter(v => !v.ce).length;
  let action = '';
  if (!s.sheetUrl) action = '<p class="warnbox">Connect your Google Sheet email tracker first (Settings, top right).</p>';
  else if (!rows.length) action = '<p class="okbox">Nothing to send — every vendor that needs a 1099 has a W-9 or is already in the tracker.</p>';
  else action = `<button class="primary" id="sPost" type="button">Send ${rows.length} vendors to the email tracker</button>`;
  const md = modal(`<div class="phead"><h2>Send to email tracker</h2><button class="ghost" id="sClose" type="button">Close</button></div>
    <p class="note">${rows.length} vendor${rows.length === 1 ? '' : 's'} from <b>${esc(c.name)}</b> still need a W-9. They go to the <b>${esc(c.name)}</b> tab of your Sheet. Then the sender clicks <b>1099 Emails → Create my W-9 request drafts</b> in the Sheet.</p>
    <p class="note">Sender: <b>${esc(c.senderName || 'not set')}</b> &lt;${esc(c.senderEmail || 'not set')}&gt;${!c.senderEmail ? ' — set it first so drafts land in the right Gmail.' : ''}${missing ? ` · ${missing} without a contact email (add it here or in the Sheet).` : ''}</p>
    <div class="actions">${action}</div>
    <details><summary class="note" style="cursor:pointer">Didn’t arrive? Other ways to send</summary>
      <div class="actions" style="margin-top:8px" id="sAlt"></div>
      <p class="note">Or copy rows and paste them into the Sheet’s <b>Paste_Inbox</b> tab at A2, then run <b>1099 Emails → Import pasted rows</b>.</p>
      <button class="ghost" id="sCopy" type="button">Copy ${rows.length} rows</button></details>`, 'Send to email tracker');
  md.querySelector('#sClose').onclick = closeModal;
  if (s.sheetUrl && rows.length) {
    md.querySelector('#sAlt').innerHTML = sheetLinks(s, c.name, c.taxYear, sender, payload)
      .map((u, i, all) => `<a class="ghost" style="text-decoration:none" href="${esc(u.url)}" target="_blank" rel="noopener">${all.length > 1 ? `Open link ${i + 1} of ${all.length}` : 'Send by link'} (${u.n})</a>`).join(' ');
    md.querySelector('#sPost').onclick = async () => {
      const b = md.querySelector('#sPost'); b.disabled = true; b.textContent = 'Sending…';
      try {
        await sendToSheet(s, c.name, c.taxYear, sender, payload);
        const d = new Date().toISOString(); rows.forEach(v => Store.patchVendor(state.cid, v.fk, { rq: { s: 'tracker', d } }));
        closeModal(); toast(`${rows.length} vendors sent — check the ${c.name} tab in your Sheet`);
      } catch (e) { b.disabled = false; b.textContent = 'Try again'; toast('Couldn’t reach the Sheet: ' + e.message + '. Use the link option below.'); }
    };
  }
  md.querySelector('#sCopy').onclick = async () => {
    const tsv = rows.map(v => [c.name, v.n, v.cn || '', v.ce || '', r2(v.ev.rep), v.ev.fb, v.ev.s, c.senderName || '', c.senderEmail || '', c.taxYear].join('\t')).join('\n');
    try { await navigator.clipboard.writeText(tsv); toast('Copied — paste into Paste_Inbox!A2'); } catch { toast('Copy failed — your browser blocked clipboard access'); }
  };
}

/* ------------------------------ settings ------------------------------ */
$('btnSettings').onclick = () => {
  const s = Store.settings();
  const md = modal(`<div class="phead"><h2>Settings</h2><button class="ghost" id="eClose" type="button">Close</button></div>
    <h3 style="font-size:15px">Google Sheet email tracker</h3>
    <div class="grid2"><label class="field">Web app URL<input id="eUrl" value="${esc(s.sheetUrl || '')}" placeholder="https://script.google.com/macros/s/…/exec"></label>
    <label class="field">App key (Sheet → Settings tab)<input id="eKey" value="${esc(s.sheetKey || '')}"></label></div>
    ${s.sheetUrl ? `<p class="note"><a href="${esc(s.sheetUrl.split('?')[0] + '?k=' + encodeURIComponent(s.sheetKey || ''))}" target="_blank" rel="noopener">Test the connection</a> — should say “Connected”.</p>` : ''}
    <h3 style="font-size:15px">Google Drive (read W-9 folders)</h3>
    <div class="grid2"><label class="field">Google OAuth Client ID<input id="eCid" value="${esc(s.googleClientId || '')}" placeholder="…apps.googleusercontent.com"></label></div>
    <h3 style="font-size:15px">Optional: AI reading of scanned W-9s (Gemini)</h3>
    <div class="grid2"><label class="field">Gemini API key<input id="eGk" type="password" value="${esc(s.geminiKey || '')}" placeholder="Leave blank to skip"></label>
    <label class="field">Gemini model<input id="eGm" value="${esc(s.geminiModel || 'gemini-2.5-flash')}"></label></div>
    <p class="note">Keys are stored only in this browser. W-9 images go to Google’s Gemini API only when a key is set.</p>
    <div class="pfoot"><button class="primary" id="eSave" type="button">Save</button></div>`, 'Settings');
  md.querySelector('#eClose').onclick = closeModal;
  md.querySelector('#eSave').onclick = () => {
    const u = md.querySelector('#eUrl').value.trim();
    if (u && !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/.test(u)) return toast('The web app URL should end in /exec.');
    Store.putSettings({ sheetUrl: u, sheetKey: md.querySelector('#eKey').value.trim(), googleClientId: md.querySelector('#eCid').value.trim(),
      geminiKey: md.querySelector('#eGk').value.trim(), geminiModel: md.querySelector('#eGm').value.trim() });
    closeModal(); toast('Settings saved');
  };
};

/* ------------------------------ rules panel ------------------------------ */
$('btnRules').onclick = () => {
  const y = latestYear(), R = RULES[y], P = RULES[y - 1];
  const md = modal(`<div class="phead"><h2>IRS 1099 rules used</h2><button class="ghost" id="rClose" type="button">Close</button></div>
    <p class="note">For payments made in ${y}, the One Big Beautiful Bill Act raised the 1099-NEC and most 1099-MISC thresholds to $2,000 per vendor per year. From 2027 it is indexed for inflation. Last verified ${RULES_VERIFIED}. The Sheet’s weekly IRS check emails you when IRS pages change; then update <code>src/logic.js</code>.</p>
    <div class="tablewrap"><table style="min-width:620px"><thead><tr><th>Payment type</th><th>Form / box (${y})</th><th style="text-align:right">${y - 1}</th><th style="text-align:right">${y}</th><th>Corporations</th></tr></thead><tbody>
    ${Object.keys(R).map(k => `<tr><td>${k}</td><td>${R[k].form} ${R[k].box}</td><td class="amt num">${money(P[k].th)}</td><td class="amt num"><b>${money(R[k].th)}</b></td><td>${R[k].corp ? 'Exempt' : 'Included'}</td></tr>`).join('')}
    </tbody></table></div>
    <p class="note">Always excluded: card / PayPal payments (1099-K), goods and inventory, freight, utilities, payroll. Foreign vendors with a W-8 go to 1042-S review. Sources: <a href="https://www.irs.gov/instructions/i1099mec" target="_blank" rel="noopener">IRS Instructions for Forms 1099-MISC and 1099-NEC</a>.</p>`, 'IRS rules');
  md.querySelector('#rClose').onclick = closeModal;
};

/* ------------------------------ export / backup ------------------------------ */
function exportCsv(c, vs) {
  const head = ['Company', 'Vendor', 'Total paid', 'Card excluded', 'Reportable', '1099 status', 'Form/box', 'Reason', 'Entity type', 'W-9', 'Contact name', 'Contact email'];
  const lines = [head, ...vs.map(v => [c.name, v.n, v.t, v.x, r2(v.ev.rep), v.ev.s, v.ev.fb, v.ev.why, v.e || 'Unknown', { rec: 'Received', trk: 'Requested', req: 'To request', na: 'Not needed' }[v.w], v.cn || '', v.ce || ''])];
  download(`${slug(c.name)}-1099-tracker-${c.taxYear}.csv`, lines.map(r => r.map(x => { const s = String(x ?? ''); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(',')).join('\n'), 'text/csv');
}
$('btnBackup').onclick = () => download(`ep1099-backup-${new Date().toISOString().slice(0, 10)}.json`, Store.exportJson(), 'application/json');
$('btnRestore').onclick = () => $('restoreFile').click();
$('restoreFile').onchange = async e => { const f = e.target.files[0]; e.target.value = ''; if (!f) return;
  try { Store.importJson(await f.text()); state.cid = null; toast('Backup restored'); } catch (err) { toast('Restore failed: ' + err.message); } };

/* ------------------------------ boot ------------------------------ */
$('btnUploadGl').onclick = () => $('glFile').click();
$('glFile').onchange = e => { const f = e.target.files[0]; e.target.value = ''; if (f) readGl(f); };
$('w9File').onchange = async e => { const fs = [...e.target.files]; e.target.value = ''; if (fs.length) processW9s(await expandFiles(fs)); };
Store.onChange(render);
render();
