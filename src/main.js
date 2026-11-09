import './styles.css';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import {
  RULES, RULES_VERIFIED, CATS, ENTITIES, S, FIELD_LABELS, latestYear, rulesFor, normalizeVendor, fieldKey, slug, cidFor, r2, money,
  mapHeaders, prepareSheet, listAccounts, yearOf, aggregate, evaluate, w9State, stats, matchByName, suggestCategory
} from './logic.js';
import { Store } from './store.js';
import { loginWithGoogle, logoutUser } from './firebase.js';
import { readW9, listDriveFiles, downloadDriveFile, sendToSheet, sheetLinks } from './integrations.js';
import SAMPLE_GL from '../sample/sample_GL_2026.csv?raw';

const SCLASS = { 'Eligible': 's-elig', 'Needs W-9': 's-w9', 'Needs review': 's-rev', 'Not eligible': 's-not' };
const state = { cid: null, filter: 'all', q: '', clientView: false, teamScope: 'all' };
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

/* --------------------------------- auth UI --------------------------------- */
function renderAuth(user) {
  const uSec = $('userSection');
  const authOverlay = $('authOverlay');

  if (user) {
    if (authOverlay) authOverlay.hidden = true;
    if (uSec) {
      uSec.innerHTML = `
        <div class="user-pill">
          ${user.photoURL ? `<img class="user-avatar" src="${esc(user.photoURL)}" alt="" onerror="this.style.display='none'">` : ''}
          <span class="user-name" title="${esc(user.email)}">${esc(user.name || user.email.split('@')[0])}</span>
          <button class="tbtn tbtn-signout" id="btnSignOut" type="button">Sign out</button>
        </div>`;
      const btnOut = $('btnSignOut');
      if (btnOut) btnOut.onclick = () => logoutUser();
    }
  } else if (Store.isConfigured()) {
    if (uSec) uSec.innerHTML = `<button class="primary" id="btnHeaderSignIn" type="button" style="padding:4px 10px;font-size:12px">Sign in</button>`;
    const btnIn = $('btnHeaderSignIn');
    if (btnIn) btnIn.onclick = () => showSignInModal();

    if (authOverlay) {
      authOverlay.hidden = false;
      showSignInModal();
    }
  } else {
    // Local fallback mode when Firebase keys aren't added to .env yet
    if (uSec) uSec.innerHTML = `<span class="badge-tool" title="Add VITE_FIREBASE_* in .env to enable team Firebase">Local mode</span>`;
    if (authOverlay) authOverlay.hidden = true;
  }
}

function showSignInModal(errorMsg = '') {
  const authOverlay = $('authOverlay');
  if (!authOverlay) return;
  authOverlay.hidden = false;
  authOverlay.innerHTML = `
    <div class="auth-card">
      <svg class="logo-svg" viewBox="0 0 39 40" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path fill-rule="evenodd" clip-rule="evenodd" d="M14.8191 3.21039C14.009 10.5704 14.0637 18.6298 14.2302 26.7065C14.265 28.4329 14.3954 30.1594 14.525 31.8738C14.5421 32.1004 14.5592 32.3267 14.5761 32.5528C14.7031 34.2902 14.3572 34.6182 12.7237 34.2511C5.99079 32.7418 1.16498 27.1604 0.0636213 21.3358C-0.188179 20.0002 0.300095 19.5115 1.65982 19.6657C4.25227 19.9553 6.84254 20.2622 9.43061 20.5865C11.0772 20.7928 11.2217 20.6473 11.1648 18.9642L11.1581 18.7681C11.1033 17.1534 11.0497 15.5749 11.1275 13.9822C11.2107 12.2926 11.0465 12.082 9.32551 12.399C6.86589 12.8522 4.40846 13.309 1.95322 13.7694C1.77507 13.8022 1.59526 13.8254 1.41459 13.8389C0.400815 13.9019 0.216892 13.6847 0.525621 12.7313C1.29416 10.3641 2.41741 8.18151 4.16031 6.37896C6.6192 3.83802 9.47221 1.90734 12.9339 0.966975C14.8783 0.441412 15.0096 0.543481 14.8191 3.21039ZM36.8506 10.6226C38.1841 13.27 38.6548 16.1041 38.5191 19.0121C38.1928 26.0268 35.2741 31.7711 29.6732 36.0516C27.2077 37.9345 24.4182 39.2832 21.2937 39.6806C20.0675 39.837 19.7238 39.4852 19.761 38.219C19.8559 34.3866 19.9851 30.5542 20.1485 26.7218C20.2208 25.0995 21.609 24.1157 23.7219 23.9246C28.7229 23.4598 32.5722 21.1274 35.088 16.7578C36.1675 14.8576 36.6667 12.7944 36.8506 10.6226ZM22.8089 1.31443C20.6609 0.443556 20.5033 0.606431 20.639 2.86288C20.7156 4.81376 20.7987 6.76513 20.8819 8.71674C20.9772 10.9538 21.0726 13.1912 21.158 15.4286C21.1755 16.1974 21.5017 16.61 22.2528 16.5753C23.4086 16.5534 24.5616 16.4525 25.7035 16.2734C27.2472 15.9824 28.2084 14.9812 28.3945 13.4045C28.8456 9.5714 27.9435 6.0662 25.4802 3.08005C24.8277 2.28953 23.7855 1.70968 22.8089 1.31443Z" fill="url(#ep_auth_grad)"/>
        <defs>
          <linearGradient id="ep_auth_grad" x1="-4.05" y1="22.2" x2="35.39" y2="7.67" gradientUnits="userSpaceOnUse">
            <stop stop-color="#02CFA1"/>
            <stop offset="1" stop-color="#00D9A6"/>
          </linearGradient>
        </defs>
      </svg>
      <h2>Team Sign-in</h2>
      <p>Sign in with your <b>@enduringplanet.com</b> Google Workspace account to access and collaborate on company 1099 trackers.</p>
      ${errorMsg ? `<div class="auth-error">${esc(errorMsg)}</div>` : ''}
      <button type="button" class="auth-google-btn" id="btnGoogleLogin">
        <svg width="18" height="18" viewBox="0 0 18 18"><path fill="#4285F4" d="M17.64 9.2c0-.637-.057-1.251-.164-1.84H9v3.481h4.844c-.209 1.125-.843 2.078-1.796 2.717v2.258h2.908c1.702-1.567 2.684-3.874 2.684-6.616z"/><path fill="#34A853" d="M9 18c2.43 0 4.467-.806 5.956-2.184l-2.908-2.258c-.806.54-1.837.86-3.048.86-2.344 0-4.328-1.584-5.036-3.711H.957v2.332C2.438 15.983 5.482 18 9 18z"/><path fill="#FBBC05" d="M3.964 10.707c-.18-.54-.282-1.117-.282-1.707s.102-1.167.282-1.707V4.961H.957C.347 6.173 0 7.548 0 9s.347 2.827.957 4.039l3.007-2.332z"/><path fill="#EA4335" d="M9 3.58c1.321 0 2.508.454 3.44 1.345l2.582-2.58C13.463.891 11.426 0 9 0 5.482 0 2.438 2.017.957 4.961L3.964 7.293C4.672 5.166 6.656 3.58 9 3.58z"/></svg>
        Sign in with Google
      </button>
      <p class="domain-note">Access is strictly restricted to @enduringplanet.com domains.</p>
    </div>`;

  const btn = $('btnGoogleLogin');
  if (btn) {
    btn.onclick = async () => {
      btn.disabled = true;
      btn.textContent = 'Signing in…';
      try {
        await loginWithGoogle();
      } catch (err) {
        showSignInModal(err.message || 'Failed to sign in with Google');
      }
    };
  }
}

/* --------------------------------- rail --------------------------------- */
function renderRail() {
  let ids = sortedClientIds();
  if (state.teamScope === 'my') {
    const myEmail = (Store.currentUser()?.email || '').toLowerCase();
    ids = ids.filter(id => {
      const c = Store.clients()[id];
      return c && c.assignedTo && c.assignedTo.toLowerCase() === myEmail;
    });
  }

  $('clist').innerHTML = ids.length ? ids.map(id => {
    const c = Store.clients()[id];
    const s = stats(vendorList(id), rulesFor(c.taxYear));
    const pct = s.scope ? Math.round(s.rec / s.scope * 100) : 0;
    const assigneeText = c.assignedToName || c.assignedTo;
    const assigneeBadge = assigneeText
      ? `<span class="c-assignee">👤 ${esc(assigneeText)}</span>`
      : `<span class="c-assignee unassigned">Unassigned</span>`;
    return `<button class="citem" type="button" data-cid="${id}" aria-current="${id === state.cid}">
      <span class="cn">${esc(c.name)}</span>
      <span class="cm"><span>${s.scope} need 1099 · ${s.req} to request</span><span class="num">${pct}%</span></span>
      ${assigneeBadge}
      ${segBar(s)}</button>`;
  }).join('') : (state.teamScope === 'my'
    ? '<p class="note">No clients assigned to you yet. Switch to "All clients" to pick a company and assign it.</p>'
    : '<p class="note">No company trackers yet. Upload a GL to create one per company.</p>');

  $('clist').querySelectorAll('.citem').forEach(b => b.onclick = () => select(b.dataset.cid));

  const teamTabs = $('teamTabs');
  if (teamTabs) {
    teamTabs.querySelectorAll('.team-tab').forEach(b => {
      b.classList.toggle('active', b.dataset.scope === state.teamScope);
      b.onclick = () => {
        state.teamScope = b.dataset.scope;
        renderRail();
      };
    });
  }
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
        <label>Assigned to
          <select id="cAssigned">
            <option value="">— Unassigned —</option>
            ${Store.teamUsers().map(u => `<option value="${esc(u.email)}" data-name="${esc(u.name || u.email)}" ${c.assignedTo === u.email ? 'selected' : ''}>${esc(u.name ? `${u.name} (${u.email})` : u.email)}</option>`).join('')}
          </select>
        </label>
        <label>W-9 Drive folder <input id="cDrive" value="${esc(c.driveUrl || '')}" placeholder="Paste Drive folder link (view-only is fine)"></label>
        <label>Sender name <input id="cSName" value="${esc(c.senderName || '')}" placeholder="Who sends the emails"></label>
        <label>Sender email <input id="cSEmail" value="${esc(c.senderEmail || '')}" placeholder="name@company.com"></label>
      </div></div>
    <div class="actions">
      <button class="ghost internal" id="aW9" type="button">Upload W-9s</button>
      <button class="ghost internal" id="aDrive" type="button">Read W-9s from Drive</button>
      <button class="primary internal" id="aSend" type="button">Send ${st.req} to email tracker</button>
      <button class="ghost" id="aView" type="button">${cv ? 'Back to full tracker' : 'Client view'}</button>
      <button class="ghost" id="aXlsx" type="button">Download tracker (Excel)</button>
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

  const cAssigned = m.querySelector('#cAssigned');
  if (cAssigned) {
    cAssigned.onchange = () => {
      const opt = cAssigned.selectedOptions[0];
      const email = cAssigned.value;
      const name = opt ? (opt.dataset.name || email) : '';
      Store.putClient(cid, { assignedTo: email, assignedToName: name });
      toast(email ? `Assigned to ${name || email}` : 'Client unassigned');
    };
  }

  [['cDrive', 'driveUrl'], ['cSName', 'senderName'], ['cSEmail', 'senderEmail']].forEach(([id, f]) => { const el = m.querySelector('#' + id); el.onchange = () => Store.putClient(cid, { [f]: el.value.trim() }); });
  m.querySelector('#aW9').onclick = () => $('w9File').click();
  m.querySelector('#aDrive').onclick = () => readFromDrive(c);
  m.querySelector('#aSend').onclick = () => openSend(c, vs);
  m.querySelector('#aView').onclick = () => { state.clientView = !state.clientView; renderMain(); };
  m.querySelector('#aCsv').onclick = () => exportCsv(c, vs);
  m.querySelector('#aXlsx').onclick = async e => {
    const b = e.currentTarget; b.disabled = true; b.textContent = 'Building…';
    try {
      const { buildTrackerXlsx } = await import('./trackerXlsx.js');
      const blob = await buildTrackerXlsx(c, vs, { assignee: c.assignedToName || c.assignedTo || '' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
      a.download = `${slug(c.name)}-1099-tracker-${c.taxYear}.xlsx`; a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 3000);
      toast('Tracker downloaded');
    } catch (err) { toast('Couldn’t build the tracker: ' + err.message); }
    finally { b.disabled = false; b.textContent = 'Download tracker (Excel)'; }
  };
  m.querySelector('#aDel').onclick = () => {
    const md = modal(`<h2>Delete the ${esc(c.name)} tracker?</h2><p class="note">This removes its vendors, W-9 status and contacts from this browser.</p><div class="pfoot"><button class="ghost" id="cNo" type="button">Cancel</button><button class="primary" id="cYes" type="button" style="background:var(--bad);color:#fff">Delete tracker</button></div>`, 'Delete tracker');
    md.querySelector('#cNo').onclick = closeModal; md.querySelector('#cYes').onclick = () => { closeModal(); Store.deleteClient(cid); state.cid = null; toast('Tracker deleted'); };
  };
}

function emptyHtml() {
  return `<section class="card empty"><h1 style="font-size:26px">Find your 1099 vendors in minutes</h1>
    <p class="note" style="max-width:520px;margin:12px auto 24px">Upload a general ledger or AP payment export (.xlsx or .csv). The tool totals payments per vendor, applies the IRS ${latestYear()} rules (${money(RULES[latestYear()]['Nonemployee compensation'].th)} for 1099-NEC and most 1099-MISC), and builds one color-coded tracker for each company in the file.</p>
    <div class="dropzone" id="emptyDrop" tabindex="0" role="button" aria-label="Drop general ledger file here">Drop a GL file here, or click to choose</div>
    <div style="margin-top:20px"><button class="ghost" id="trySample" type="button">Try it with a sample GL</button></div>
    <p class="note" style="margin-top:28px">Next: read W-9s from the client’s Drive folder (or upload them), then send vendors that still need a W-9 to your Google Sheet email tracker.</p></section>`;
}

function bindEmpty() {
  const d = $('emptyDrop'); if (!d) return;
  d.onclick = () => $('glFile').click();
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
  const p = prepareSheet(aoa);
  if (!p) { toast('No header row found in that file.'); return; }
  imp = { name, headers: p.headers, rows: p.rows, idx: mapHeaders(p.headers), taxYear: latestYear(), defaultCo: p.companyHint, grouped: p.grouped, step: 1 };
  renderImport();
}

function renderImport() {
  const steps = `<div class="steps"><span class="${imp.step === 1 ? 'on' : ''}">1 · Columns</span><span class="${imp.step === 2 ? 'on' : ''}">2 · GL accounts</span><span class="${imp.step === 3 ? 'on' : ''}">3 · Results</span></div>`;
  let body = '';
  if (imp.step === 1) {
    const opt = k => `<select id="m_${k}"><option value="">— not in file —</option>${imp.headers.map((h, i) => `<option value="${i}" ${imp.idx[k] === i ? 'selected' : ''}>${esc(h)}</option>`).join('')}</select>`;
    body = `<p class="note"><b>${esc(imp.name)}</b> · ${imp.rows.length.toLocaleString()} transaction lines. Check that each field points at the right column.</p>
      ${imp.grouped ? '<p class="okbox">Report-style ledger detected (QuickBooks/Xero): account headings were copied onto each transaction as “Account Name”, and Total / Beginning Balance rows were skipped.</p>' : ''}
      <div class="grid2">${Object.keys(FIELD_LABELS).map(k => `<label class="field">${FIELD_LABELS[k]}${k === 'vendor' ? ' *' : ''}${opt(k)}</label>`).join('')}
      <label class="field">Tax year<select id="m_year">${Object.keys(RULES).map(y => `<option ${+y === imp.taxYear ? 'selected' : ''}>${y}</option>`).join('')}<option>${latestYear() + 1}</option></select></label>
      <label class="field">Company name (if no Company column)<input id="m_co" value="${esc(imp.defaultCo)}" placeholder="e.g. Acme Holdings LLC"></label></div>
      ${imp.idx.method === undefined ? '<p class="warnbox">No payment method column found. Card and PayPal payments can’t be excluded — use the AP payment register if you can.</p>' : ''}`;
  } else if (imp.step === 2 && !imp.accounts.length) {
    const yrs = {}; imp.rows.forEach(r => { const y = imp.idx.date !== undefined ? yearOf(r[imp.idx.date]) : 0; if (y) yrs[y] = (yrs[y] || 0) + 1; });
    body = `<p class="warnbox">No vendor payment lines found for tax year <b>${imp.taxYear}</b>. ${imp.rows.length.toLocaleString()} lines were read; years found in the file: ${Object.keys(yrs).map(y => `${y} (${yrs[y].toLocaleString()} lines)`).join(', ') || 'none — check the Payment date column'}. Go back and pick the right tax year, or check the Vendor, Amount and Payment date columns.</p>`;
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
      driveUrl: prev.driveUrl || '', senderName: prev.senderName || '', senderEmail: prev.senderEmail || '', assignedTo: prev.assignedTo || '', assignedToName: prev.assignedToName || '' });
    const st = stats(Object.values(out), rulesFor(imp.taxYear));
    lines.push(`<b>${esc(co)}</b>: ${st.paid} vendors · ${st.scope} need a 1099 · ${st.rev} need review`);
  }
  imp.summary = lines.join('<br>') + (skipped ? `<br>${skipped} lines outside ${imp.taxYear} were ignored.` : '');
  imp.step = 3; renderImport();
}

/* ------------------------------- W-9s ------------------------------- */
let w9q = [];
async function expandFiles(files) {
  const out = [];
  for (const f of files) {
    if (/\.zip$/i.test(f.name)) {
      try {
        const zip = await JSZip.loadAsync(await f.arrayBuffer());
        for (const [p, zf] of Object.entries(zip.files)) {
          if (zf.dir || !/\.(pdf|png|jpe?g)$/i.test(p)) continue;
          out.push(new File([await zf.async('blob')], p.split('/').pop(), { type: p.endsWith('.pdf') ? 'application/pdf' : 'image/jpeg' }));
        }
      } catch (e) { toast('Couldn’t unzip ' + f.name + ': ' + e.message); }
    } else out.push(f);
  }
  return out;
}

function processW9s(files) {
  const vs = vendorList(state.cid);
  w9q = files.map(f => {
    const raw = f.name.replace(/\.[^.]+$/, '').replace(/[-_]+/g, ' ').replace(/\bw9\b/gi, '').trim();
    const match = matchByName(raw, vs);
    return { f, name: f.name, st: 'Queued', match, x: null, note: '' };
  });
  renderW9(); readQueue();
}

function renderW9() {
  const vs = vendorList(state.cid);
  const rows = w9q.map((it, i) => `<tr>
    <td><b>${esc(it.name)}</b>${it.note ? `<div class="sub">${esc(it.note)}</div>` : ''}</td>
    <td><select data-wi="${i}"><option value="">— choose vendor —</option>${vs.map(v => `<option value="${v.fk}" ${it.match === v.fk ? 'selected' : ''}>${esc(v.n)}</option>`).join('')}</select></td>
    <td>${it.x ? `<span class="pill s-elig">${esc(it.x.tax_classification || 'Read')}</span> ${it.x.tin_last4 ? `TIN •••${esc(it.x.tin_last4)}` : ''}` : esc(it.st)}</td>
  </tr>`).join('');
  const md = modal(`<div class="phead"><h2>Match W-9 files (${w9q.length})</h2><button class="ghost" id="wClose" type="button">Close</button></div>
    <p class="note">Each file is matched to a vendor by name. Change the dropdown if the match looks wrong. Never stores full SSN/EIN — only last 4 digits.</p>
    <div class="tablewrap" style="max-height:55vh;overflow:auto"><table style="min-width:580px"><thead><tr><th>File</th><th>Vendor</th><th>Extracted</th></tr></thead><tbody>${rows}</tbody></table></div>
    <div class="pfoot"><button class="ghost" id="wCancel" type="button">Cancel</button><button class="primary" id="wSave" type="button">Apply to tracker</button></div>`, 'Match W-9s');
  md.querySelector('#wClose').onclick = closeModal;
  md.querySelector('#wCancel').onclick = closeModal;
  md.querySelectorAll('[data-wi]').forEach(s => s.onchange = () => w9q[+s.dataset.wi].match = s.value);
  md.querySelector('#wSave').onclick = () => {
    const cid = state.cid; let applied = 0;
    for (const it of w9q) {
      if (!it.match) continue;
      const patch = { w9: { on: true, f: it.name, d: new Date().toISOString(), src: it.x ? 'ai' : 'filename', l4: it.x && it.x.tin_last4, ln: it.x && it.x.legal_name } };
      if (it.x && it.x.tax_classification && ENTITIES.includes(it.x.tax_classification)) patch.e = it.x.tax_classification;
      Store.patchVendor(cid, it.match, patch); applied++;
    }
    closeModal(); toast(`Saved W-9 on ${applied} vendor(s)`);
  };
}

async function readQueue() {
  const vs = vendorList(state.cid);
  for (const it of w9q) {
    it.st = 'Reading…'; renderW9();
    const r = await readW9(it.f, Store.settings());
    it.x = r.x; it.note = r.note;
    if (!it.match && it.x) it.match = matchByName(it.x.legal_name || '', vs) || matchByName(it.x.business_name || '', vs);
    it.st = it.x ? 'Done' : (it.note || 'Done');
    renderW9();
  }
}

async function readFromDrive(c) {
  if (!c.driveUrl) return toast('Paste the client’s W-9 Drive folder link first.');
  const s = Store.settings();
  if (!s.googleClientId) return toast('Add your Google OAuth Client ID in Settings first (see README).');
  try {
    toast('Signing in to Google…');
    const files = await listDriveFiles(c.driveUrl, s.googleClientId);
    if (!files.length) return toast('No PDF or image files found in that Drive folder.');
    toast(`Found ${files.length} file(s). Downloading…`);
    const blobs = [];
    for (const df of files) {
      const b = await downloadDriveFile(df.id);
      blobs.push(new File([b], df.name, { type: df.mimeType }));
    }
    processW9s(blobs);
  } catch (err) { toast('Drive error: ' + err.message); }
}

/* ------------------------------- email tracker ------------------------------- */
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
  const md = modal(`<div class="phead"><h2>Send to email tracker (${rows.length})</h2><button class="ghost" id="sClose" type="button">Close</button></div>
    <p class="note">Appends these vendors to your Google Sheet’s <b>Vendors</b> tab. The Sheet’s trigger runs every morning at 9am, drafts customized emails in the sender’s Gmail (never sends automatically), tracks replies, and sends follow-ups 5 and 10 days later.</p>
    ${missing ? `<p class="warnbox">${missing} vendor(s) have no contact email. You can add them in the table before sending, or directly in the Sheet.</p>` : ''}
    <div class="pfoot"><button class="ghost" id="sCancel" type="button">Cancel</button>${action}</div>`, 'Send to email tracker');
  md.querySelector('#sClose').onclick = closeModal;
  md.querySelector('#sCancel').onclick = closeModal;
  const p = md.querySelector('#sPost');
  if (p) p.onclick = async () => {
    p.disabled = true; p.textContent = 'Sending…';
    try {
      const res = await sendToSheet(s, c.name, payload, sender);
      const now = new Date().toISOString();
      rows.forEach(v => Store.patchVendor(state.cid, v.fk, { rq: { on: true, d: now } }));
      const links = sheetLinks(s.sheetUrl);
      closeModal();
      modal(`<h2>${res.added || rows.length} vendors sent</h2><p class="note">They’re now in your Google Sheet. Tomorrow at 9am (or when you run "Create Gmail drafts now" from the Sheet menu), drafts will be waiting in Gmail.</p>
        <p><a href="${esc(links.vendors)}" target="_blank" rel="noopener">Open Vendors tab</a> · <a href="${esc(links.drafts)}" target="_blank" rel="noopener">Open Drafts log</a></p>
        <div class="pfoot"><button class="primary" id="sDone" type="button">Done</button></div>`, 'Sent');
      $('modal').querySelector('#sDone').onclick = closeModal;
    } catch (err) { toast('Send failed: ' + err.message); p.disabled = false; p.textContent = 'Retry'; }
  };
}

/* ------------------------------ settings ------------------------------ */
$('btnSettings').onclick = () => {
  const s = Store.settings();
  const md = modal(`<div class="phead"><h2>Settings</h2><button class="ghost" id="eClose" type="button">Close</button></div>
    <div style="background:var(--brand-tint);border:1px solid rgba(1,213,164,0.3);padding:10px 12px;border-radius:8px;margin-bottom:14px;font-size:12.5px;color:var(--brand-ink)">
      <b>Team-wide settings:</b> Synced to Firebase Firestore (<code>settings/main</code>) for all team members.
    </div>
    <h3 style="font-size:15px">Google Sheet email tracker</h3>
    <div class="grid2"><label class="field">Web app URL<input id="eUrl" value="${esc(s.sheetUrl || '')}" placeholder="https://script.google.com/macros/s/…/exec"></label>
    <label class="field">App key (Sheet → Settings tab)<input id="eKey" value="${esc(s.sheetKey || '')}"></label></div>
    ${s.sheetUrl ? `<p class="note"><a href="${esc(s.sheetUrl.split('?')[0] + '?k=' + encodeURIComponent(s.sheetKey || ''))}" target="_blank" rel="noopener">Test the connection</a> — should say “Connected”.</p>` : ''}
    <h3 style="font-size:15px">Google Drive (read W-9 folders)</h3>
    <div class="grid2"><label class="field">Google OAuth Client ID<input id="eCid" value="${esc(s.googleClientId || '')}" placeholder="…apps.googleusercontent.com"></label></div>

    <div style="background:var(--surface-2);border:1px solid var(--line);padding:10px 12px;border-radius:8px;margin:16px 0 10px;font-size:12.5px;color:var(--ink-2)">
      <b>Personal settings:</b> Saved only in this browser (not shared across the team).
    </div>
    <h3 style="font-size:15px">Optional: AI reading of scanned W-9s (Gemini)</h3>
    <div class="grid2"><label class="field">Gemini API key<input id="eGk" type="password" value="${esc(s.geminiKey || '')}" placeholder="Leave blank to skip"></label>
    <label class="field">Gemini model<input id="eGm" value="${esc(s.geminiModel || 'gemini-2.5-flash')}"></label></div>
    <p class="note">Keys are stored only in this browser. W-9 images go to Google’s Gemini API only when a key is set.</p>
    <div class="pfoot"><button class="primary" id="eSave" type="button">Save</button></div>`, 'Settings');
  md.querySelector('#eClose').onclick = closeModal;
  md.querySelector('#eSave').onclick = () => {
    const u = md.querySelector('#eUrl').value.trim();
    if (u && !/^https:\/\/script\.google\.com\/(a\/macros\/[^/]+|macros)\/s\/[^/]+\/exec/.test(u)) return toast('The web app URL should end in /exec.');
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
  try { await Store.importJson(await f.text()); state.cid = null; toast('Backup restored'); } catch (err) { toast('Restore failed: ' + err.message); } };

/* ------------------------------ boot ------------------------------ */
$('btnUploadGl').onclick = () => $('glFile').click();
$('glFile').onchange = e => { const f = e.target.files[0]; e.target.value = ''; if (f) readGl(f); };
$('w9File').onchange = async e => { const fs = [...e.target.files]; e.target.value = ''; if (fs.length) processW9s(await expandFiles(fs)); };

const btnImpLocal = $('btnImportLocal');
if (btnImpLocal) {
  btnImpLocal.onclick = async () => {
    btnImpLocal.disabled = true;
    toast('Importing local browser data to Firebase…');
    try {
      const res = await Store.importFromLocalStorage();
      toast(res.message);
      render();
    } catch (err) {
      toast('Import failed: ' + err.message);
    } finally {
      btnImpLocal.disabled = false;
    }
  };
}

Store.onAuthChange(user => {
  renderAuth(user);
  render();
});

Store.onChange(render);
render();
