/******************************************************************************
 * EP 1099 — EMAIL TRACKER (Google Sheets + Apps Script)  v2.0, Nov 2026
 *
 * The EP 1099 Tool web app does the 1099 work and sends vendors here (doPost / doGet). This Sheet does the email side:
 *   • Receives vendors from the app (one click → a tab per company)
 *   • Creates W-9 request DRAFTS in the sender's own Gmail (nothing auto-sends)
 *   • One-click follow-ups as replies in the same email thread
 *   • Syncs sent / replied status and saves W-9 attachments to Drive
 *   • Summary tab across all companies, colour-coded
 *   • Weekly IRS 1099 page check that emails the admin when rule text changes
 *
 * Install: see "Setup guide" in the implementation plan. Needs the Gmail
 * advanced service (appsscript.json enables it).
 ******************************************************************************/

const SYS = { SETTINGS: 'Settings', TEMPLATES: 'Templates', SUMMARY: 'Summary', INBOX: 'Paste_Inbox', LOG: 'Log' };
const CT = ['Vendor', 'Reportable', 'Form / Box', '1099 Status', 'Contact Name', 'Contact Email',
  'Sender Name', 'Sender Email', 'Request Status', 'Sent Count', 'Last Contact', 'Vendor Replied',
  'W-9 Received', 'Thread ID', 'Added On', 'Notes'];
const RQ = { NOT: 'Not sent', DRAFT: 'Draft created', SENT: 'Sent', FU_DRAFT: 'Follow-up draft created',
  FU1: 'Follow-up 1 sent', FU2: 'Follow-up 2 sent', FINAL: 'Final notice sent', RECEIVED: 'W-9 received', DONE: 'Verified' };
const RQ_LIST = Object.keys(RQ).map(function (k) { return RQ[k]; });
const IRS_W9 = 'https://www.irs.gov/pub/irs-pdf/fw9.pdf';
const COLORS = { head: '#1c2a4d', headInk: '#ffffff', ok: '#dcf4e8', warn: '#fff0cc', info: '#e1eaff', bad: '#fde2e6', risk: '#ffe3d1', muted: '#eceff3' };

/* ================================ MENU ==================================== */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('1099 Emails')
    .addItem('Create my W-9 request drafts (this tab)', 'draftsThisTab')
    .addItem('Create my W-9 request drafts (all companies)', 'draftsAllTabs')
    .addSeparator()
    .addItem('Follow up on selected row(s)', 'followUpSelected')
    .addItem('Create all my due follow-ups', 'followUpAllDue')
    .addSeparator()
    .addItem('Sync sent / reply status (mine)', 'syncMine')
    .addItem('Refresh summary', 'refreshSummary')
    .addItem('Import pasted rows (Paste_Inbox)', 'importPasted')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('1. Set up this sheet (first time)', 'setupSheet')
      .addItem('Show app key + web app steps', 'showKey')
      .addItem('Turn on daily email sync (each sender)', 'installDailySync')
      .addItem('Turn on weekly IRS rule check (admin)', 'installIrsWatcher')
      .addItem('Check IRS for rule changes now', 'checkIrsUpdates'))
    .addToUi();
}

/* ================================ SETUP =================================== */

function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('APP_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 20); props.setProperty('APP_KEY', key); }
  props.setProperty('SHEET_ID', ss.getId());
  const me = Session.getActiveUser().getEmail();

  ensure_(SYS.SETTINGS, ['Setting', 'Value', 'Help'], [
    ['App Key', key, 'Paste into the EP 1099 Tool → Email tracker setup. Keep private.'],
    ['Admin Email', me, 'Gets IRS rule-change alerts'],
    ['Default Tax Year', 2026, 'Used when the app does not send one'],
    ['W-9 Due Date', 'December 15, 2026', 'Shown in emails as {Due Date}'],
    ['Follow-up After (days)', 7, 'Days after the last email before a follow-up is due'],
    ['Sender Title', 'Accounts Payable', 'Shown in emails as {Sender Title}'],
    ['Secure Upload Link', '', 'Optional: a Google Form / file-request link for vendors'],
    ['Attach W-9 PDF', 'Yes', 'Yes = attach the blank IRS W-9 to request drafts'],
    ['W-9 Save Folder URL', '', 'Drive folder where W-9s that vendors email back are saved (a subfolder per company is created)'],
    ['IRS Watch URLs', 'https://www.irs.gov/instructions/i1099mec, https://www.irs.gov/forms-pubs/about-form-1099-nec, https://www.irs.gov/forms-pubs/about-form-1099-misc', 'Pages checked weekly']
  ]);
  ensure_(SYS.TEMPLATES, ['Template', 'Subject', 'Body'], templates_());
  ensure_(SYS.SUMMARY, ['Company'], []);
  ensure_(SYS.INBOX, ['Company', 'Vendor', 'Contact Name', 'Contact Email', 'Reportable', 'Form / Box', '1099 Status', 'Sender Name', 'Sender Email', 'Tax Year'], []);
  ensure_(SYS.LOG, ['Timestamp', 'User', 'Action', 'Details'], []);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
  ss.setActiveSheet(ss.getSheetByName(SYS.SETTINGS));
  log_('Setup', 'Sheet set up');
  showKey();
}

function showKey() {
  const key = PropertiesService.getScriptProperties().getProperty('APP_KEY') || '(run Setup first)';
  SpreadsheetApp.getUi().alert('Connect the EP 1099 Tool',
    'App key: ' + key + '\n\n' +
    '1) Extensions → Apps Script → Deploy → New deployment → type "Web app".\n' +
    '2) Execute as: Me.   Who has access: Anyone with a Google account (or your domain).\n' +
    '3) Deploy, copy the Web app URL (ends in /exec).\n' +
    '4) In the EP 1099 Tool: Email tracker setup → paste the URL and this key → Save.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function templates_() {
  return [
    ['Request', 'W-9 request – {Company} vendor records ({Tax Year})',
      'Hi {Contact Name},\n\nAs part of our year-end 1099 reporting for {Tax Year}, we need a current Form W-9 for {Vendor}.\n\n' +
      'Please complete and sign the attached Form W-9 and return it by {Due Date}{Upload Line}.\n\n' +
      'If you have already sent us a W-9 in the last three years and nothing has changed, simply reply to confirm.\n\n' +
      'For your security, please do not type your Taxpayer Identification Number in the body of an email.\n\n' +
      'Thank you,\n{Sender Name}\n{Sender Title}\n{Company}'],
    ['Follow-up 1', '', 'Hi {Contact Name},\n\nA quick reminder about the Form W-9 for {Vendor}. We need it by {Due Date} so we can issue your 1099 on time.\n\n' +
      'Blank form: {W-9 Link}{Upload Line}.\n\nThank you,\n{Sender Name}\n{Sender Title}, {Company}'],
    ['Follow-up 2', '', 'Hi {Contact Name},\n\nFollowing up again on the Form W-9 for {Vendor}. We still need it by {Due Date} to complete our {Tax Year} 1099 filings.\n\n' +
      'Blank form: {W-9 Link}{Upload Line}. If someone else handles this at your company, please let me know who to contact.\n\nThank you,\n{Sender Name}\n{Sender Title}, {Company}'],
    ['Final notice', '', 'Hi {Contact Name},\n\nThis is our final request for a Form W-9 for {Vendor}. Our deadline is {Due Date}.\n\n' +
      'Without a W-9 showing a valid Taxpayer Identification Number, IRS rules may require us to apply backup withholding (currently 24%) to future payments.\n\n' +
      'Blank form: {W-9 Link}{Upload Line}.\n\nThank you,\n{Sender Name}\n{Sender Title}, {Company}']
  ];
}

/* ===================== RECEIVE FROM THE WEB APP (doGet) ==================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const props = PropertiesService.getScriptProperties();
  if (!p.k || p.k !== props.getProperty('APP_KEY')) return page_('Not connected', 'The app key is missing or wrong. Copy it from the Sheet’s Settings tab into the EP 1099 Tool.', false);
  if (!p.d) return page_('Connected', 'The EP 1099 Tool can send vendors to this email tracker.', true);
  let data;
  try { data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p.d)).getDataAsString()); }
  catch (err) { return page_('Couldn’t read the data', 'The link looks incomplete. Send again from the app, or use “Copy rows”.', false); }
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const res = upsert_(ss, data.c, data.y, data.s || ['', ''], (data.r || []).map(function (r) {
      return { vendor: r[0], cn: r[1], ce: r[2], amt: r[3], fb: r[4], st: r[5] };
    }));
    refreshSummary_(ss);
    log_('From app', data.c + ': ' + res.added + ' added, ' + res.updated + ' updated', ss);
    return page_('Added to the email tracker',
      res.added + ' new and ' + res.updated + ' updated vendors on the “' + res.tab + '” tab. ' +
      'The sender can now open the Sheet and click 1099 Emails → Create my W-9 request drafts. You can close this tab.', true, ss.getUrl());
  } finally { lock.releaseLock(); }
}

/** Receives a POST from the EP 1099 Tool (fetch, no-cors). Body = JSON {k, c, y, s, r}. */
function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return ContentService.createTextOutput('bad json'); }
  if (!data.k || data.k !== props.getProperty('APP_KEY')) return ContentService.createTextOutput('bad key');
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const res = upsert_(ss, data.c, data.y, data.s || ['', ''], (data.r || []).map(function (r) {
      return { vendor: r[0], cn: r[1], ce: r[2], amt: r[3], fb: r[4], st: r[5] };
    }));
    refreshSummary_(ss);
    log_('From app (POST)', data.c + ': ' + res.added + ' added, ' + res.updated + ' updated', ss);
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

function page_(title, msg, ok, url) {
  const c = ok ? '#11845b' : '#b4233b';
  return HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;max-width:560px;margin:60px auto;padding:24px;border:1px solid #dbe3ea;border-radius:12px">' +
    '<div style="font-weight:700;font-size:20px;color:' + c + '">' + title + '</div><p style="color:#333;line-height:1.5">' + msg + '</p>' +
    (url ? '<p><a href="' + url + '" target="_blank">Open the email tracker sheet</a></p>' : '') + '</div>').setTitle('EP 1099 email tracker');
}

function upsert_(ss, company, taxYear, sender, rows) {
  const sh = clientTab_(ss, company, taxYear);
  const t = read_(sh);
  const byV = {};
  t.rows.forEach(function (r, i) { byV[String(r['Vendor']).trim().toLowerCase()] = i; });
  let added = 0, updated = 0;
  rows.forEach(function (x) {
    const k = String(x.vendor).trim().toLowerCase();
    if (byV[k] !== undefined) {
      const r = t.rows[byV[k]];
      r['Reportable'] = x.amt; r['Form / Box'] = x.fb; r['1099 Status'] = x.st;
      if (x.cn && !r['Contact Name']) r['Contact Name'] = x.cn;
      if (x.ce && !r['Contact Email']) r['Contact Email'] = x.ce;
      if (sender[0] && !r['Sender Name']) r['Sender Name'] = sender[0];
      if (sender[1] && !r['Sender Email']) r['Sender Email'] = sender[1];
      updated++;
    } else {
      const r = {};
      CT.forEach(function (h) { r[h] = ''; });
      r['Vendor'] = x.vendor; r['Reportable'] = x.amt; r['Form / Box'] = x.fb; r['1099 Status'] = x.st;
      r['Contact Name'] = x.cn || ''; r['Contact Email'] = x.ce || '';
      r['Sender Name'] = sender[0] || ''; r['Sender Email'] = sender[1] || '';
      r['Request Status'] = RQ.NOT; r['Sent Count'] = 0; r['W-9 Received'] = 'No'; r['Added On'] = new Date();
      t.rows.push(r); added++;
    }
  });
  write_(sh, t.rows);
  return { added: added, updated: updated, tab: sh.getName() };
}

function importPasted() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SYS.INBOX);
  const vals = sh.getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); });
  if (!vals.length) { SpreadsheetApp.getUi().alert('Paste rows from the app into Paste_Inbox at cell A2 first.'); return; }
  const groups = {};
  vals.forEach(function (r) {
    const k = r[0] + '|' + (r[9] || '');
    (groups[k] = groups[k] || { c: r[0], y: r[9], s: [r[7], r[8]], rows: [] })
      .rows.push({ vendor: r[1], cn: r[2], ce: r[3], amt: r[4], fb: r[5], st: r[6] });
  });
  const out = [];
  Object.keys(groups).forEach(function (k) { const g = groups[k]; const res = upsert_(ss, g.c, g.y, g.s, g.rows); out.push(res.tab + ': ' + res.added + ' added, ' + res.updated + ' updated'); });
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), sh.getLastColumn()).clearContent();
  refreshSummary_(ss);
  log_('Import pasted', out.join(' | '));
  SpreadsheetApp.getUi().alert(out.join('\n'));
}

/* ============================= CLIENT TABS ================================ */

function clientTab_(ss, company, taxYear) {
  const name = String(company || 'Company').replace(/[\[\]\*\?\:\/\\]/g, ' ').trim().slice(0, 90);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, CT.length).setValues([CT]);
    sh.setTabColor('#00c49a');
  }
  const year = taxYear || setting_('Default Tax Year', ss) || new Date().getFullYear() - 1;
  sh.getRange('A1').setNote('Tax year: ' + year + '\nCompany: ' + company);
  return sh;
}

function isClientTab_(sh) {
  if (Object.keys(SYS).map(function (k) { return SYS[k]; }).indexOf(sh.getName()) >= 0) return false;
  const h = sh.getRange(1, 1, 1, 2).getValues()[0];
  return h[0] === 'Vendor' && h[1] === 'Reportable';
}

function clientTabs_(ss) { return ss.getSheets().filter(isClientTab_); }

function tabYear_(sh) {
  const m = String(sh.getRange('A1').getNote() || '').match(/Tax year:\s*(\d{4})/);
  return m ? m[1] : String(setting_('Default Tax Year') || '');
}

function read_(sh) {
  const v = sh.getDataRange().getValues();
  const head = v[0].map(String);
  const rows = v.slice(1).filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    const o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o;
  });
  return { head: head, rows: rows };
}

function write_(sh, rows) {
  const n = Math.max(sh.getLastRow() - 1, 0);
  if (n) sh.getRange(2, 1, n, CT.length).clearContent();
  sh.getRange(1, 1, 1, CT.length).setValues([CT]);
  if (rows.length) sh.getRange(2, 1, rows.length, CT.length).setValues(rows.map(function (o) { return CT.map(function (h) { return o[h] === undefined ? '' : o[h]; }); }));
  format_(sh, rows.length);
}

function format_(sh, n) {
  const H = function (h) { return CT.indexOf(h) + 1; };
  sh.getRange(1, 1, 1, CT.length).setFontWeight('bold').setBackground(COLORS.head).setFontColor(COLORS.headInk).setWrap(true);
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 220);
  sh.hideColumns(H('Thread ID'));
  const rows = Math.max(n, 1);
  sh.getRange(2, H('Reportable'), rows, 1).setNumberFormat('$#,##0');
  sh.getRange(2, H('Last Contact'), rows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, H('Added On'), rows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, H('Request Status'), Math.max(sh.getMaxRows() - 1, 1), 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RQ_LIST, true).setAllowInvalid(true).build());
  const rs = sh.getRange(2, H('Request Status'), Math.max(sh.getMaxRows() - 1, 1), 1);
  const st = sh.getRange(2, H('1099 Status'), Math.max(sh.getMaxRows() - 1, 1), 1);
  const w9 = sh.getRange(2, H('W-9 Received'), Math.max(sh.getMaxRows() - 1, 1), 1);
  const cf = function (range, text, bg, contains) {
    const b = SpreadsheetApp.newConditionalFormatRule();
    (contains ? b.whenTextContains(text) : b.whenTextEqualTo(text));
    return b.setBackground(bg).setRanges([range]).build();
  };
  sh.setConditionalFormatRules([
    cf(rs, RQ.NOT, COLORS.bad), cf(rs, RQ.DRAFT, COLORS.warn), cf(rs, RQ.FU_DRAFT, COLORS.warn),
    cf(rs, 'sent', COLORS.info, true), cf(rs, RQ.RECEIVED, COLORS.ok), cf(rs, RQ.DONE, COLORS.ok),
    cf(st, 'Eligible', COLORS.ok), cf(st, 'Needs W-9', COLORS.warn), cf(st, 'Needs review', COLORS.risk),
    cf(w9, 'Yes', COLORS.ok, true), cf(w9, 'No', COLORS.muted)
  ]);
}

/* ================================ DRAFTS ================================== */

function draftsThisTab() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isClientTab_(sh)) { SpreadsheetApp.getUi().alert('Open a company tab first (or use “all companies”).'); return; }
  createDrafts_([sh]);
}
function draftsAllTabs() { createDrafts_(clientTabs_(SpreadsheetApp.getActive())); }

function createDrafts_(tabs) {
  const ui = SpreadsheetApp.getUi();
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const todo = []; let noContact = 0, others = 0;
  tabs.forEach(function (sh) {
    const t = read_(sh);
    t.rows.forEach(function (r, i) {
      if (r['W-9 Received'] && /^y/i.test(r['W-9 Received'])) return;
      if (r['Request Status'] && r['Request Status'] !== RQ.NOT) return;
      if (String(r['Sender Email']).trim().toLowerCase() !== me) { others++; return; }
      if (!String(r['Contact Email']).trim()) { noContact++; return; }
      todo.push({ sh: sh, t: t, i: i });
    });
  });
  if (!todo.length) {
    ui.alert('No request drafts to create for ' + me + '.\n\nA row needs Request Status "Not sent", Sender Email = your email, and a Contact Email.' +
      (noContact ? '\n' + noContact + ' of your rows have no Contact Email.' : '') + (others ? '\n' + others + ' rows belong to other senders.' : ''));
    return;
  }
  if (ui.alert('Create ' + todo.length + ' W-9 request drafts in your Gmail (' + me + ')?\nNothing is sent — you review and send from Gmail Drafts.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const tpl = templatesMap_()['Request'];
  let w9 = null;
  if (/^y/i.test(String(setting_('Attach W-9 PDF')))) { try { w9 = UrlFetchApp.fetch(IRS_W9).getBlob().setName('Form W-9.pdf'); } catch (e) { log_('W-9 download failed', e.message); } }
  const touched = [];
  todo.forEach(function (x) {
    const r = x.t.rows[x.i];
    const data = tplData_(r, x.sh);
    const opts = { name: r['Sender Name'] || undefined };
    if (w9) opts.attachments = [w9];
    const d = GmailApp.createDraft(String(r['Contact Email']).trim(), fill_(tpl.subject, data), fill_(tpl.body, data), opts);
    r['Thread ID'] = d.getMessage().getThread().getId();
    r['Request Status'] = RQ.DRAFT;
    if (touched.indexOf(x) < 0) touched.push(x);
  });
  const done = [];
  todo.forEach(function (x) { if (done.indexOf(x.sh) < 0) { write_(x.sh, x.t.rows); done.push(x.sh); } });
  refreshSummary_(SpreadsheetApp.getActive());
  log_('Request drafts', todo.length + ' for ' + me);
  ui.alert(todo.length + ' drafts created. Open Gmail → Drafts, review each one and click Send.\nThen run 1099 Emails → Sync sent / reply status (or turn on the daily sync).');
}

/* ============================== FOLLOW-UPS ================================ */

function followUpSelected() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isClientTab_(sh)) { ui.alert('Select vendor rows on a company tab first.'); return; }
  const sel = sh.getActiveRange(); const t = read_(sh);
  const idx = [];
  for (let row = Math.max(2, sel.getRow()); row < sel.getRow() + sel.getNumRows(); row++) idx.push(row - 2);
  const res = followUp_(sh, t, idx);
  write_(sh, t.rows); refreshSummary_(SpreadsheetApp.getActive());
  ui.alert(res.join('\n') || 'Nothing selected.');
}

function followUpAllDue() {
  const ui = SpreadsheetApp.getUi();
  syncMine(true);
  const days = Number(setting_('Follow-up After (days)')) || 7;
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const now = new Date(); let total = 0; const plans = [];
  clientTabs_(SpreadsheetApp.getActive()).forEach(function (sh) {
    const t = read_(sh); const idx = [];
    t.rows.forEach(function (r, i) {
      if (String(r['Sender Email']).trim().toLowerCase() !== me) return;
      if (/^y/i.test(String(r['W-9 Received']))) return;
      if ((Number(r['Sent Count']) || 0) < 1 || r['Request Status'] === RQ.FU_DRAFT) return;
      if (r['Vendor Replied']) return; // read the reply first
      const last = r['Last Contact'] ? new Date(r['Last Contact']) : null;
      if (last && (now - last) / 86400000 >= days) idx.push(i);
    });
    if (idx.length) { plans.push({ sh: sh, t: t, idx: idx }); total += idx.length; }
  });
  if (!total) { ui.alert('No follow-ups are due for ' + me + '.'); return; }
  if (ui.alert('Create ' + total + ' follow-up drafts (replies in the same email threads)?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  let made = 0;
  plans.forEach(function (pl) { made += followUp_(pl.sh, pl.t, pl.idx).filter(function (x) { return /created/.test(x); }).length; write_(pl.sh, pl.t.rows); });
  refreshSummary_(SpreadsheetApp.getActive());
  ui.alert(made + ' follow-up drafts created. Review them in Gmail → Drafts and click Send.');
}

function followUp_(sh, t, idx) {
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const tpls = templatesMap_(); const out = [];
  idx.forEach(function (i) {
    const r = t.rows[i]; if (!r) return;
    const label = r['Vendor'] + ': ';
    if (/^y/i.test(String(r['W-9 Received']))) { out.push(label + 'W-9 already received — skipped'); return; }
    if (String(r['Sender Email']).trim().toLowerCase() !== me) { out.push(label + 'sender is ' + (r['Sender Email'] || 'blank') + ', not you — skipped'); return; }
    if (!r['Thread ID']) { out.push(label + 'no request yet — create the request draft first'); return; }
    let th = null; try { th = GmailApp.getThreadById(r['Thread ID']); } catch (e) { th = null; }
    if (!th) { out.push(label + 'email thread not found in your mailbox'); return; }
    const msgs = th.getMessages();
    if (msgs.some(function (m) { return m.isDraft(); })) { out.push(label + 'a draft is already waiting in this thread'); return; }
    const sent = msgs.filter(function (m) { return !m.isDraft() && from_(m, me); });
    if (!sent.length) { out.push(label + 'request not sent yet — send it from Drafts first'); return; }
    const n = sent.length;
    const tpl = n === 1 ? tpls['Follow-up 1'] : n === 2 ? tpls['Follow-up 2'] : tpls['Final notice'];
    replyDraft_(th.getId(), sent[sent.length - 1], String(r['Contact Email']).trim(), fill_(tpl.body, tplData_(r, sh)));
    r['Request Status'] = RQ.FU_DRAFT;
    out.push(label + (n === 1 ? 'Follow-up 1' : n === 2 ? 'Follow-up 2' : 'Final notice') + ' draft created');
  });
  log_('Follow-ups', out.join(' | '));
  return out;
}

function replyDraft_(threadId, last, to, body) {
  const s0 = last.getSubject() || '';
  const subject = /^re:/i.test(s0) ? s0 : 'Re: ' + s0;
  const mid = last.getHeader('Message-ID') || last.getHeader('Message-Id') || '';
  const lines = ['To: ' + to, 'Subject: =?UTF-8?B?' + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + '?=',
    'MIME-Version: 1.0', 'Content-Type: text/plain; charset=UTF-8', 'Content-Transfer-Encoding: base64'];
  if (mid) { lines.push('In-Reply-To: ' + mid); lines.push('References: ' + mid); }
  const raw = lines.join('\r\n') + '\r\n\r\n' + Utilities.base64Encode(body, Utilities.Charset.UTF_8);
  Gmail.Users.Drafts.create({ message: { raw: Utilities.base64EncodeWebSafe(raw), threadId: threadId } }, 'me');
}

/* ================================ SYNC ==================================== */

function syncMine(silent) {
  const ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const folderUrl = String(setting_('W-9 Save Folder URL', ss) || '').trim();
  const root = folderUrl ? DriveApp.getFolderById(folderId_(folderUrl)) : null;
  let synced = 0, saved = 0;
  clientTabs_(ss).forEach(function (sh) {
    const t = read_(sh); let changed = false;
    t.rows.forEach(function (r) {
      if (!r['Thread ID'] || String(r['Sender Email']).trim().toLowerCase() !== me) return;
      let th = null; try { th = GmailApp.getThreadById(r['Thread ID']); } catch (e) { th = null; }
      if (!th) return;
      const msgs = th.getMessages();
      const sent = msgs.filter(function (m) { return !m.isDraft() && from_(m, me); });
      const replies = msgs.filter(function (m) { return !m.isDraft() && !from_(m, me); });
      const draft = msgs.some(function (m) { return m.isDraft(); });
      r['Sent Count'] = sent.length;
      if (sent.length) r['Last Contact'] = sent[sent.length - 1].getDate();
      if (replies.length) r['Vendor Replied'] = 'Yes — ' + Utilities.formatDate(replies[replies.length - 1].getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (root && !/^y/i.test(String(r['W-9 Received']))) {
        replies.forEach(function (m) {
          m.getAttachments().forEach(function (a) {
            if (/^y/i.test(String(r['W-9 Received'])) || !/pdf|image/i.test(a.getContentType())) return;
            const sub = subfolder_(root, sh.getName());
            const ext = /pdf/i.test(a.getContentType()) ? '.pdf' : '.' + (a.getName().split('.').pop() || 'jpg');
            const f = sub.createFile(a.copyBlob()).setName(r['Vendor'] + '_W9_' + tabYear_(sh) + ext);
            r['W-9 Received'] = 'Yes (email) — check file'; r['Notes'] = (r['Notes'] ? r['Notes'] + ' ' : '') + f.getUrl(); saved++;
          });
        });
      }
      if (/^y/i.test(String(r['W-9 Received']))) r['Request Status'] = r['Request Status'] === RQ.DONE ? RQ.DONE : RQ.RECEIVED;
      else if (draft) r['Request Status'] = sent.length ? RQ.FU_DRAFT : RQ.DRAFT;
      else r['Request Status'] = sent.length >= 4 ? RQ.FINAL : sent.length === 3 ? RQ.FU2 : sent.length === 2 ? RQ.FU1 : sent.length === 1 ? RQ.SENT : r['Request Status'];
      synced++; changed = true;
    });
    if (changed) write_(sh, t.rows);
  });
  refreshSummary_(ss);
  const msg = synced + ' of your vendor threads synced; ' + saved + ' W-9 attachments saved to Drive. Upload saved W-9s into the EP 1099 Tool to update the client tracker.';
  log_('Sync', msg + ' (' + me + ')', ss);
  if (silent !== true) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* trigger */ } }
}

function dailySync_() { syncMine(true); }

/* ================================ SUMMARY ================================= */

function refreshSummary() { refreshSummary_(SpreadsheetApp.getActive()); SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(SYS.SUMMARY)); }

function refreshSummary_(ss) {
  const days = Number(setting_('Follow-up After (days)', ss)) || 7;
  const head = ['Company', 'Tax Year', 'Vendors in Tracker', 'Not Sent', 'Draft Waiting', 'Sent / Following Up', 'Follow-up Due', 'Vendor Replied', 'W-9 Received', '% Received', 'Open Tab'];
  const now = new Date(); const out = [];
  clientTabs_(ss).forEach(function (sh) {
    const t = read_(sh); const c = { n: 0, not: 0, draft: 0, sent: 0, due: 0, rep: 0, rec: 0 };
    t.rows.forEach(function (r) {
      c.n++;
      const s = r['Request Status']; const rec = /^y/i.test(String(r['W-9 Received']));
      if (rec) { c.rec++; return; }
      if (!s || s === RQ.NOT) c.not++;
      else if (s === RQ.DRAFT || s === RQ.FU_DRAFT) c.draft++;
      else c.sent++;
      if (r['Vendor Replied']) c.rep++;
      const last = r['Last Contact'] ? new Date(r['Last Contact']) : null;
      if (last && (Number(r['Sent Count']) || 0) > 0 && (now - last) / 86400000 >= days) c.due++;
    });
    out.push([sh.getName(), tabYear_(sh), c.n, c.not, c.draft, c.sent, c.due, c.rep, c.rec, c.n ? c.rec / c.n : 0,
      '=HYPERLINK("#gid=' + sh.getSheetId() + '","Open")']);
  });
  const sh = ss.getSheetByName(SYS.SUMMARY) || ss.insertSheet(SYS.SUMMARY);
  sh.clear(); sh.clearConditionalFormatRules();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(COLORS.head).setFontColor(COLORS.headInk).setWrap(true);
  if (out.length) {
    sh.getRange(2, 1, out.length, head.length).setValues(out);
    sh.getRange(2, 10, out.length, 1).setNumberFormat('0%');
    const pct = sh.getRange(2, 10, out.length, 1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().setGradientMinpointWithValue(COLORS.bad, SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue(COLORS.warn, SpreadsheetApp.InterpolationType.NUMBER, '0.5')
        .setGradientMaxpointWithValue(COLORS.ok, SpreadsheetApp.InterpolationType.NUMBER, '1').setRanges([pct]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(COLORS.bad).setRanges([sh.getRange(2, 7, out.length, 1)]).build()
    ]);
  }
  sh.setFrozenRows(1); sh.setColumnWidth(1, 240);
  sh.getRange(out.length + 3, 1, 2, 1).setValues([
    ['Updated ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')],
    ['Follow-up Due = sent, no W-9, last email ≥ ' + days + ' days ago. 1099s are due to vendors and the IRS by January 31.']]).setFontColor('#55627d');
}

/* ============================ IRS WATCHER ================================= */

function checkIrsUpdates() {
  const ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const urls = String(setting_('IRS Watch URLs', ss) || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  const props = PropertiesService.getScriptProperties(); const changes = []; let base = 0;
  urls.forEach(function (url) {
    let html;
    try { const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); if (r.getResponseCode() !== 200) return; html = r.getContentText(); }
    catch (e) { log_('IRS check failed', url + ': ' + e.message, ss); return; }
    const lines = ruleLines_(html);
    const id = Utilities.base64EncodeWebSafe(url).slice(0, 60);
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, lines.join('\n')));
    const old = props.getProperty('h_' + id); const oldLines = (props.getProperty('t_' + id) || '').split('\n');
    if (!old) base++;
    else if (old !== hash) changes.push({ url: url, added: lines.filter(function (l) { return oldLines.indexOf(l) < 0; }).slice(0, 25) });
    props.setProperty('h_' + id, hash); props.setProperty('t_' + id, lines.join('\n').slice(0, 8500));
  });
  if (changes.length) {
    let body = 'Changed rule text was found on IRS 1099 pages. Review with your CPA before changing the EP 1099 Tool rules.\n\n';
    changes.forEach(function (c) { body += c.url + '\n- ' + c.added.join('\n- ') + '\n\n'; });
    MailApp.sendEmail(String(setting_('Admin Email', ss)), '[EP 1099] IRS 1099 rule text changed — review needed', body);
  }
  log_('IRS check', changes.length ? changes.length + ' page(s) changed; admin emailed' : (base ? 'Baseline saved' : 'No changes'), ss);
}

function ruleLines_(html) {
  return html.replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, '\n')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').split(/\n/).map(function (s) { return s.replace(/\s+/g, ' ').trim(); })
    .filter(function (s) { return s.length > 20 && s.length < 400 && /(\$\s?\d[\d,]*|threshold|box \d|backup withholding|january 31|inflation)/i.test(s) && !/last reviewed|page last/i.test(s); })
    .filter(function (s, i, a) { return a.indexOf(s) === i; });
}

function installIrsWatcher() {
  removeTriggers_('checkIrsUpdates');
  ScriptApp.newTrigger('checkIrsUpdates').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  checkIrsUpdates();
  SpreadsheetApp.getUi().alert('Weekly IRS check is on (Mondays ~8am). You will be emailed when IRS 1099 rule text changes.');
}
function installDailySync() {
  removeTriggers_('dailySync_');
  ScriptApp.newTrigger('dailySync_').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('Daily email sync is on for ' + Session.getActiveUser().getEmail() + ' (~7am).');
}
function removeTriggers_(fn) { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); }); }

/* ================================ HELPERS ================================= */

function tplData_(r, sh) {
  const up = String(setting_('Secure Upload Link') || '').trim();
  return {
    'Contact Name': String(r['Contact Name'] || '').trim() || 'there', 'Vendor': r['Vendor'], 'Company': sh.getName(),
    'Tax Year': tabYear_(sh), 'Due Date': fmt_(setting_('W-9 Due Date')), 'Upload Line': up ? ', or upload it securely here: ' + up : '',
    'W-9 Link': IRS_W9, 'Sender Name': r['Sender Name'] || '', 'Sender Title': String(setting_('Sender Title') || '')
  };
}
function fill_(text, data) { return String(text || '').replace(/\{([^}]+)\}/g, function (m, k) { return Object.prototype.hasOwnProperty.call(data, k) ? data[k] : m; }); }
function templatesMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SYS.TEMPLATES); const out = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { out[String(r[0]).trim()] = { subject: r[1], body: r[2] }; });
  ['Request', 'Follow-up 1', 'Follow-up 2', 'Final notice'].forEach(function (k) { if (!out[k]) throw new Error('Templates tab is missing "' + k + '".'); });
  return out;
}
function setting_(name, ss) {
  ss = ss || SpreadsheetApp.getActive();
  const sh = ss && ss.getSheetByName(SYS.SETTINGS); if (!sh) return '';
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === name) return v[i][1];
  return '';
}
function fmt_(v) { return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'MMMM d, yyyy') : String(v || ''); }
function from_(m, me) { return String(m.getFrom() || '').toLowerCase().indexOf(me) >= 0; }
function folderId_(url) { const m = String(url).match(/folders\/([a-zA-Z0-9_-]+)/) || String(url).match(/id=([a-zA-Z0-9_-]+)/); return m ? m[1] : String(url).trim(); }
function subfolder_(root, name) { const it = root.getFoldersByName(name); return it.hasNext() ? it.next() : root.createFolder(name); }
function ensure_(name, head, seed) {
  const ss = SpreadsheetApp.getActive(); let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, head.length).setValues([head]); if (seed && seed.length) sh.getRange(2, 1, seed.length, seed[0].length).setValues(seed); }
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground(COLORS.head).setFontColor(COLORS.headInk);
  sh.setFrozenRows(1); return sh;
}
function log_(action, details, ss) {
  try { const sh = (ss || SpreadsheetApp.getActive()).getSheetByName(SYS.LOG); if (sh) sh.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]); } catch (e) { }
}
