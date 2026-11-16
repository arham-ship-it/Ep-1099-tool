/******************************************************************************
 * EP 1099 — EMAIL TRACKER (Google Sheets + Apps Script)  v4.1 (timeline + management dashboard), Nov 2026
 *
 * The EP 1099 Tool web app finds the 1099 vendors and sends them here.
 * This Sheet is the team's working tracker:
 *   • Dashboard tab — every company: to email, draft ready, emailed, follow-up due, W-9 received
 *   • One formatted, colour-coded tab per company with live counts at the top
 *   • Clickable buttons: Create W-9 drafts · Follow up selected · All due follow-ups · Sync status
 *     (drafts land in the Gmail of whoever clicks — nothing is sent automatically)
 *   • Team can change Request Status / W-9 Received by dropdown; counts update instantly
 *   • Weekly IRS 1099 page check that emails the admin when rule text changes
 ******************************************************************************/

const SYS = { SETTINGS: 'Settings', TEMPLATES: 'Templates', SUMMARY: 'Dashboard', INBOX: 'Paste_Inbox', LOG: 'Log' };
const OLD_SYS = ['Summary'];
const HR = 8;                       // header row of each company tab (rows 1-7 = title, buttons, counts, legend)
const CT = ['Vendor', 'Reportable', 'Form / Box', '1099 Status', 'Contact Name', 'Contact Email',
  'Sender Name', 'Sender Email', 'Request Status', 'Sent Count', 'Last Contact', 'Vendor Replied',
  'W-9 Received', 'Thread ID', 'Added On', 'Notes'];
const COL = {}; CT.forEach(function (h, i) { COL[h] = i + 1; });
const RQ = { NOT: 'Not sent', DRAFT: 'Draft created', SENT: 'Sent', FU_DRAFT: 'Follow-up draft created',
  FU1: 'Follow-up 1 sent', FU2: 'Follow-up 2 sent', FINAL: 'Final notice sent', RECEIVED: 'W-9 received', DONE: 'Verified' };
const RQ_LIST = Object.keys(RQ).map(function (k) { return RQ[k]; });
const W9_LIST = ['No', 'Yes', 'Yes (on file)', 'Yes (Drive)', 'Yes (email) — check file'];
const IRS_W9 = 'https://www.irs.gov/pub/irs-pdf/fw9.pdf';
const C = {
  navy: '#1c2a4d', teal: '#00c49a', mint: '#e6f9f4', white: '#ffffff', ink: '#1c2a4d', muted: '#55627d', line: '#d5e6e1', band: '#ffffff',
  red: '#fde2e6', redInk: '#b4233b', yellow: '#fff0cc', yellowInk: '#8a5a00', blue: '#e1eaff', blueInk: '#2357c6',
  green: '#dcf4e8', greenInk: '#11845b', orange: '#ffe3d1', orangeInk: '#b3470e', grey: '#eceff3', greyInk: '#5f6b80'
};
const FONT = 'Poppins';
const VERSION = 'v4.1';
// Timeline: key, Settings row name, default date. Dates falling on a weekend/holiday move to the next working day.
const TL = [
  ['start', 'Timeline: Project start', '2026-11-17'],
  ['inv', 'Timeline: W-9 inventory target', '2026-11-25'],
  ['invMax', 'Timeline: W-9 inventory latest', '2026-11-30'],
  ['email', 'Timeline: Email requests start', '2026-12-01'],
  ['fu1', 'Timeline: Follow-up 1', '2026-12-04'],
  ['fu2', 'Timeline: Follow-up 2', '2026-12-08'],
  ['fu3', 'Timeline: Follow-up 3 (final notice)', '2026-12-11'],
  ['due', 'Timeline: W-9 collection deadline', '2026-12-15'],
  ['file', 'Timeline: File 1099s (IRS + vendors)', '2027-01-31']
];
const HOLIDAYS_DEFAULT = '2026-11-26, 2026-12-25, 2027-01-01, 2027-01-18';

/* ================================ MENU ==================================== */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('1099 Emails')
    .addItem('Create my W-9 request drafts (this tab)', 'draftsThisTab')
    .addItem('Create drafts for the company selected on Dashboard', 'draftsSelectedCompany')
    .addItem('Create my W-9 request drafts (all companies)', 'draftsAllTabs')
    .addSeparator()
    .addItem('Follow up on selected row(s)', 'followUpSelected')
    .addItem('Create all my due follow-ups', 'followUpAllDue')
    .addSeparator()
    .addItem('Check Drive for W-9s (this tab)', 'checkDriveThisTab')
    .addItem('Check Drive for W-9s (all companies)', 'checkDriveAll')
    .addItem('Run today’s tasks now (sync, Drive check, due drafts)', 'runTodayTasks')
    .addSeparator()
    .addItem('Undo: reset selected rows to Not sent (deletes their drafts)', 'resetSelected')
    .addItem('Sync sent / reply status (mine)', 'syncMine')
    .addItem('Refresh dashboard', 'refreshSummary')
    .addItem('Re-format all company tabs', 'reformatAll')
    .addItem('Repair this tab (fix broken layout)', 'repairThisTab')
    .addItem('Import pasted rows (Paste_Inbox)', 'importPasted')
    .addSeparator()
    .addSubMenu(ui.createMenu('Setup')
      .addItem('1. Set up this sheet (first time)', 'setupSheet')
      .addItem('Show app key + web app steps', 'showKey')
      .addItem('Turn on daily automation (each sender)', 'installDailyAuto')
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
    ['App Key', key, 'Paste into the EP 1099 Tool → Settings. Keep private.'],
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
  ensureSettings_(ss);
  ensure_(SYS.TEMPLATES, ['Template', 'Subject', 'Body'], templates_());
  ensure_(SYS.INBOX, ['Company', 'Vendor', 'Contact Name', 'Contact Email', 'Reportable', 'Form / Box', '1099 Status', 'Sender Name', 'Sender Email', 'Tax Year'], []);
  ensure_(SYS.LOG, ['Timestamp', 'User', 'Action', 'Details'], []);
  OLD_SYS.forEach(function (n) { const s = ss.getSheetByName(n); if (s && ss.getSheets().length > 1) ss.deleteSheet(s); });
  refreshSummary_(ss);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
  ss.setActiveSheet(ss.getSheetByName(SYS.SUMMARY));
  ss.moveActiveSheet(1);
  log_('Setup', 'Sheet set up (' + VERSION + ')');
  showKey();
}

function showKey() {
  const key = PropertiesService.getScriptProperties().getProperty('APP_KEY') || '(run Setup first)';
  SpreadsheetApp.getUi().alert('Connect the EP 1099 Tool',
    'App key: ' + key + '\n\n1) Extensions → Apps Script → Deploy → Manage deployments (or New deployment) → Web app.\n' +
    '2) Execute as: Me.   Who has access: Anyone (or Anyone within your company).\n' +
    '3) Copy the Web app URL (ends in /exec).\n4) In the EP 1099 Tool: Settings → paste the URL and this key → Save.',
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

/* ===================== RECEIVE FROM THE WEB APP ==================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const props = PropertiesService.getScriptProperties();
  if (!p.k || p.k !== props.getProperty('APP_KEY')) return page_('Not connected', 'The app key is missing or wrong. Copy it from the Sheet’s Settings tab into the EP 1099 Tool.', false);
  if (p.action === 'w9') return w9Page_(p);
  if (!p.d) return page_('Connected', 'The EP 1099 Tool can send vendors to this email tracker.', true);
  let data;
  try { data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p.d)).getDataAsString()); }
  catch (err) { return page_('Couldn’t read the data', 'The link looks incomplete. Send again from the app, or use “Copy rows”.', false); }
  const res = receive_(data);
  return page_('Added to the email tracker',
    res.added + ' new and ' + res.updated + ' updated vendors on the “' + res.tab + '” tab. ' +
    'Open the Sheet and click the green “Create W-9 drafts” button on that tab. You can close this tab.', true, res.url);
}

function doPost(e) {
  const props = PropertiesService.getScriptProperties();
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return ContentService.createTextOutput('bad json'); }
  if (!data.k || data.k !== props.getProperty('APP_KEY')) return ContentService.createTextOutput('bad key');
  return ContentService.createTextOutput(JSON.stringify(receive_(data))).setMimeType(ContentService.MimeType.JSON);
}

function receive_(data) {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const res = upsert_(ss, data.c, data.y, data.s || ['', ''], (data.r || []).map(function (r) {
      return { vendor: r[0], cn: r[1], ce: r[2], amt: r[3], fb: r[4], st: r[5], w9: r[6] };
    }), data.m);
    refreshSummary_(ss);
    log_('From app', data.c + ': ' + res.added + ' added, ' + res.updated + ' updated', ss);
    res.url = ss.getUrl() + '#gid=' + ss.getSheetByName(res.tab).getSheetId();
    return res;
  } finally { lock.releaseLock(); }
}

function page_(title, msg, ok, url) {
  const c = ok ? C.greenInk : C.redInk;
  return HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;max-width:560px;margin:60px auto;padding:24px;border:1px solid #dbe3ea;border-radius:12px">' +
    '<div style="font-weight:700;font-size:20px;color:' + c + '">' + title + '</div><p style="color:#333;line-height:1.5">' + msg + '</p>' +
    (url ? '<p><a href="' + url + '" target="_blank">Open the company tab in the email tracker</a></p>' : '') + '</div>').setTitle('EP 1099 email tracker');
}

function upsert_(ss, company, taxYear, sender, rows, m) {
  const sh = clientTab_(ss, company, taxYear);
  if (m) { const mt = meta_(sh); if (m.paid !== undefined) mt.paid = m.paid; if (m.folder) mt.folder = m.folder; setMeta_(sh, mt); }
  const t = read_(sh);
  const byV = {};
  t.rows.forEach(function (r, i) { byV[String(r['Vendor']).trim().toLowerCase()] = i; });
  let added = 0, updated = 0;
  rows.forEach(function (x) {
    const k = String(x.vendor).trim().toLowerCase();
    let r;
    if (byV[k] !== undefined) { r = t.rows[byV[k]]; updated++; }
    else {
      r = {}; CT.forEach(function (h) { r[h] = ''; });
      r['Vendor'] = x.vendor; r['Request Status'] = RQ.NOT; r['Sent Count'] = 0; r['W-9 Received'] = 'No'; r['Added On'] = new Date();
      t.rows.push(r); added++;
    }
    r['Reportable'] = x.amt; r['Form / Box'] = x.fb; r['1099 Status'] = x.st;
    if (x.cn && !r['Contact Name']) r['Contact Name'] = x.cn;
    if (x.ce && !r['Contact Email']) r['Contact Email'] = x.ce;
    if (sender[0] && !r['Sender Name']) r['Sender Name'] = sender[0];
    if (sender[1] && !r['Sender Email']) r['Sender Email'] = sender[1];
    if (x.w9 && !/^y/i.test(String(r['W-9 Received']))) {            // the app already has this vendor's W-9
      r['W-9 Received'] = 'Yes (on file)';
      if (!r['Request Status'] || r['Request Status'] === RQ.NOT) r['Request Status'] = RQ.RECEIVED;
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
    (groups[k] = groups[k] || { c: r[0], y: r[9], s: [r[7], r[8]], rows: [] }).rows.push({ vendor: r[1], cn: r[2], ce: r[3], amt: r[4], fb: r[5], st: r[6] });
  });
  const out = [];
  Object.keys(groups).forEach(function (k) { const g = groups[k]; const res = upsert_(ss, g.c, g.y, g.s, g.rows); out.push(res.tab + ': ' + res.added + ' added, ' + res.updated + ' updated'); });
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), sh.getLastColumn()).clearContent();
  refreshSummary_(ss);
  log_('Import pasted', out.join(' | '));
  SpreadsheetApp.getUi().alert(out.join('\n'));
}

/* ======================= COMPANY TABS (formatted) ======================== */

function clientTab_(ss, company, taxYear) {
  const name = tabName_(company);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const meta = meta_(sh);
  meta.company = company || meta.company || name;
  meta.year = taxYear || meta.year || setting_('Default Tax Year', ss) || new Date().getFullYear() - 1;
  setMeta_(sh, meta);
  return sh;
}
function tabName_(company) { return String(company || 'Company').replace(/[\[\]\*\?\:\/\\']/g, ' ').trim().slice(0, 90); }
function meta_(sh) {
  const n = String(sh.getRange('A1').getNote() || '');
  const g = function (re) { const m = n.match(re); return m ? m[1].trim() : ''; };
  return { year: g(/Tax year:\s*(\d{4})/), company: g(/Company:\s*(.+)/), paid: g(/Vendors paid:\s*(\d+)/), folder: g(/W-9 folder:\s*(\S+)/), checked: g(/Drive checked:\s*(.+)/) };
}
function setMeta_(sh, m) {
  sh.getRange('A1').setNote('Tax year: ' + m.year + '\nCompany: ' + m.company + (m.paid ? '\nVendors paid: ' + m.paid : '') +
    (m.folder ? '\nW-9 folder: ' + m.folder : '') + (m.checked ? '\nDrive checked: ' + m.checked : ''));
}

function headerRow_(sh) {
  const last = Math.min(sh.getLastRow(), 60);
  if (last < 1) return 0;
  const a = sh.getRange(1, 1, last, COL['Request Status']).getValues();
  if (a.length >= HR && a[HR - 1][0] === 'Vendor' && a[HR - 1][COL['Request Status'] - 1] === 'Request Status') return HR;
  for (let i = 0; i < a.length; i++) if (a[i][0] === 'Vendor' && a[i][COL['Request Status'] - 1] === 'Request Status') return i + 1;
  return 0;
}
function isClientTab_(sh) {
  const n = sh.getName();
  if (Object.keys(SYS).map(function (k) { return SYS[k]; }).concat(OLD_SYS).indexOf(n) >= 0) return false;
  return headerRow_(sh) > 0 || /Company:/.test(String(sh.getRange('A1').getNote() || ''));
}
function clientTabs_(ss) { return ss.getSheets().filter(isClientTab_); }
function tabYear_(sh) { return String(meta_(sh).year || setting_('Default Tax Year') || ''); }
function tabCompany_(sh) { return meta_(sh).company || sh.getName(); }

/**
 * Reads vendor rows by column position (the column order never changes), wherever they sit on the tab.
 * This also recovers a tab whose layout was broken (e.g. by Ctrl+Z after a script run).
 */
function read_(sh) {
  const hr = headerRow_(sh) || HR;
  const last = sh.getLastRow();
  if (last < 1) return { hr: hr, rows: [] };
  const v = sh.getRange(1, 1, last, CT.length).getValues();
  const seen = {};
  const rows = v.filter(function (r) {
    const name = String(r[0]).trim();
    if (!name || name === 'Vendor') return false;
    const ok = RQ_LIST.indexOf(String(r[COL['Request Status'] - 1]).trim()) >= 0 || W9_LIST.indexOf(String(r[COL['W-9 Received'] - 1]).trim()) >= 0;
    if (!ok) return false;
    const k = name.toLowerCase(); if (seen[k]) return false; seen[k] = 1;  // drop duplicates left by partial undo
    return true;
  }).map(function (r) { const o = {}; CT.forEach(function (h, i) { o[h] = r[i]; }); return o; });
  return { hr: hr, rows: rows };
}

/** Rewrites a company tab in the formatted tracker layout. */
function write_(sh, rows) {
  const meta = meta_(sh);
  const tl = timeline_();
  const D = function (d) { return 'DATE(' + d.getFullYear() + ',' + (d.getMonth() + 1) + ',' + d.getDate() + ')'; };
  if (sh.getMaxColumns() < CT.length) sh.insertColumnsAfter(sh.getMaxColumns(), CT.length - sh.getMaxColumns());
  const need = HR + rows.length + 5;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart(); all.clearContent(); all.clearFormat(); all.clearDataValidations();
  sh.clearConditionalFormatRules();
  all.setFontFamily(FONT).setFontSize(10).setFontColor(C.ink).setVerticalAlignment('middle');
  sh.getRange(1, 1, 3, CT.length).setWrap(false);
  setMeta_(sh, meta);

  // Row 1-2: title band
  sh.getRange(1, 1, 2, CT.length).setBackground(C.white);
  sh.getRange(2, 1, 1, CT.length).setBorder(null, null, true, null, null, null, C.teal, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange(1, 1).setValue((meta.company || sh.getName()) + '  —  1099 W-9 Tracker')
    .setFontSize(18).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange(2, 1).setValue('Tax year ' + meta.year + '   ·   Emails start ' + short_(tl.email) + '   ·   Follow-ups ' + short_(tl.fu1) + ', ' + short_(tl.fu2) + ', ' + short_(tl.fu3) +
    '   ·   W-9s due ' + short_(tl.due) + '   ·   1099s due ' + short_(tl.file) + ', ' + tl.file.getFullYear()).setFontSize(10).setFontColor(C.muted);
  sh.setRowHeight(1, 40); sh.setRowHeight(2, 24);
  // Row 3: buttons (images); row 4-5: live counts
  sh.setRowHeight(3, 46);
  sh.getRange(3, 12).setValue('Select vendor rows first for “Follow up selected” and “Undo selected”.').setFontSize(9).setFontColor(C.muted).setFontStyle('italic');
  const kpi = [
    ['Vendors', '=COUNTA(A' + (HR + 1) + ':A)', C.band, C.ink],
    ['To email', '=COUNTIF(I' + (HR + 1) + ':I,"' + RQ.NOT + '")', C.red, C.redInk],
    ['Draft ready', '=COUNTIF(I' + (HR + 1) + ':I,"*draft*")', C.yellow, C.yellowInk],
    ['Emailed – waiting', '=COUNTIF(I' + (HR + 1) + ':I,"' + RQ.SENT + '")+COUNTIF(I' + (HR + 1) + ':I,"Follow-up ? sent")+COUNTIF(I' + (HR + 1) + ':I,"' + RQ.FINAL + '")', C.blue, C.blueInk],
    ['Follow-up due', '=' + [tl.fu1, tl.fu2, tl.fu3].map(function (d, i) { return 'IF(TODAY()>=' + D(d) + ',COUNTIFS(J' + (HR + 1) + ':J,' + (i + 1) + ',M' + (HR + 1) + ':M,"No"),0)'; }).join('+'), C.orange, C.orangeInk],
    ['Vendor replied', '=COUNTIF(L' + (HR + 1) + ':L,"Yes*")', C.band, C.ink],
    ['W-9 received', '=COUNTIF(M' + (HR + 1) + ':M,"Yes*")', C.green, C.greenInk],
    ['% complete', '=IF(A5=0,0,G5/A5)', C.green, C.greenInk]
  ];
  kpi.forEach(function (k, i) {
    sh.getRange(4, i + 1).setValue(k[0]).setFontSize(9).setFontWeight('bold').setFontColor(C.navy).setBackground(C.mint).setHorizontalAlignment('center');
    sh.getRange(5, i + 1).setFormula(k[1]).setFontSize(20).setFontWeight('bold').setFontColor(k[3]).setBackground(C.white).setHorizontalAlignment('center');
  });
  sh.getRange(5, 8).setNumberFormat('0%');
  sh.getRange(4, 1, 2, kpi.length).setBorder(true, true, true, true, true, false, C.teal, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(4, 22); sh.setRowHeight(5, 40);
  // Row 6: colour legend
  const legend = [['Colour key:', C.white, C.muted], ['Not sent', C.red, C.redInk], ['Draft ready', C.yellow, C.yellowInk], ['Emailed', C.blue, C.blueInk],
    ['W-9 received', C.green, C.greenInk], ['Follow-up due', C.orange, C.orangeInk]];
  legend.forEach(function (l, i) { sh.getRange(6, i + 1).setValue(l[0]).setBackground(l[1]).setFontColor(l[2]).setFontSize(9).setFontWeight(i ? 'bold' : 'normal').setHorizontalAlignment('center'); });
  sh.setRowHeight(7, 8);

  // Row 8: headers, row 9+: data
  sh.getRange(HR, 1, 1, CT.length).setValues([CT]).setFontWeight('bold').setFontColor(C.navy).setBackground(C.teal).setWrap(true).setVerticalAlignment('middle');
  sh.setRowHeight(HR, 32);
  sh.setRowHeights(HR + 1, sh.getMaxRows() - HR, 21);
  if (rows.length) {
    sh.getRange(HR + 1, 1, rows.length, CT.length).setValues(rows.map(function (o) { return CT.map(function (h) { return o[h] === undefined ? '' : o[h]; }); }));
    sh.getRange(HR + 1, 1, rows.length, CT.length).setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(HR + 1, COL['Vendor'], rows.length, 1).setFontWeight('bold');
    sh.getRange(HR + 1, COL['Reportable'], rows.length, 1).setNumberFormat('$#,##0');
    sh.getRange(HR + 1, COL['Last Contact'], rows.length, 1).setNumberFormat('mmm d, yyyy');
    sh.getRange(HR + 1, COL['Added On'], rows.length, 1).setNumberFormat('mmm d, yyyy');
  }
  // Dropdowns + colours on the whole column (so team-added rows get them too)
  const n = sh.getMaxRows() - HR;
  const rq = sh.getRange(HR + 1, COL['Request Status'], n, 1);
  const w9 = sh.getRange(HR + 1, COL['W-9 Received'], n, 1);
  const st = sh.getRange(HR + 1, COL['1099 Status'], n, 1);
  rq.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RQ_LIST, true).setAllowInvalid(true).build()).setFontWeight('bold');
  w9.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(W9_LIST, true).setAllowInvalid(true).build()).setFontWeight('bold');
  const due = sh.getRange(HR + 1, 1, n, 1);
  const cf = function (range, text, bg, ink, how) {
    const b = SpreadsheetApp.newConditionalFormatRule();
    if (how === 'contains') b.whenTextContains(text); else if (how === 'formula') b.whenFormulaSatisfied(text); else b.whenTextEqualTo(text);
    return b.setBackground(bg).setFontColor(ink).setRanges([range]).build();
  };
  const r0 = HR + 1;
  sh.setConditionalFormatRules([
    cf(due, '=AND($A' + r0 + '<>"",$M' + r0 + '="No",OR(' + [tl.fu1, tl.fu2, tl.fu3].map(function (d, i) { return 'AND($J' + r0 + '=' + (i + 1) + ',TODAY()>=' + D(d) + ')'; }).join(',') + '))', C.orange, C.orangeInk, 'formula'),
    cf(rq, RQ.NOT, C.red, C.redInk), cf(rq, 'draft', C.yellow, C.yellowInk, 'contains'), cf(rq, RQ.RECEIVED, C.green, C.greenInk), cf(rq, RQ.DONE, C.green, C.greenInk),
    cf(rq, 'sent', C.blue, C.blueInk, 'contains'),
    cf(w9, 'Yes', C.green, C.greenInk, 'contains'), cf(w9, 'No', C.grey, C.greyInk),
    cf(st, 'Eligible', C.green, C.greenInk), cf(st, 'Needs W-9', C.yellow, C.yellowInk), cf(st, 'Needs review', C.orange, C.orangeInk)
  ]);
  // Layout
  const widths = [230, 125, 110, 110, 140, 210, 120, 200, 160, 70, 110, 120, 140, 60, 100, 220];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(COL['Thread ID']);
  sh.setFrozenRows(HR); sh.setFrozenColumns(1);
  sh.setTabColor(C.teal);
  if (sh.getFilter()) sh.getFilter().remove();
  if (rows.length) sh.getRange(HR, 1, rows.length + 1, CT.length).createFilter();
  addButtons_(sh, [['drafts', 'draftsThisTab', 1, 6], ['followup', 'followUpSelected', 2, 0], ['allfu', 'followUpAllDue', 4, 0], ['sync', 'syncMine', 6, 0], ['undo', 'resetSelected', 7, 20], ['drive', 'checkDriveThisTab', 9, 0]]);
}

/** Adds clickable image buttons once (they stay put when the tab is rewritten). Click runs as the clicker. */
function addButtons_(sh, list) {
  const ver = 'btn-v4';
  const imgs = sh.getImages();
  if (imgs.length && sh.getRange('A3').getNote() === ver) return;
  imgs.forEach(function (im) { im.remove(); });
  sh.getRange('A3').setNote(ver);
  list.forEach(function (b) {
    const img = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(BTN[b[0]]), 'image/png', b[0] + '.png'), b[2], 3, b[3], 6);
    img.assignScript(b[1]);
  });
}

function repairThisTab() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isClientTab_(sh)) { SpreadsheetApp.getUi().alert('Open the company tab you want to repair first.'); return; }
  const t = read_(sh);
  write_(sh, t.rows); refreshSummary_(SpreadsheetApp.getActive());
  SpreadsheetApp.getUi().alert(sh.getName() + ' rebuilt ✓  ' + t.rows.length + ' vendors kept.\n\nTip: don’t use Ctrl+Z after a button or menu action — use “Undo selected” instead.');
}

function reformatAll() {
  const ss = SpreadsheetApp.getActive();
  clientTabs_(ss).forEach(function (sh) { write_(sh, read_(sh).rows); });
  refreshSummary_(ss);
  toast_('All company tabs re-formatted.');
}

/* ================================ DRAFTS ================================== */

function draftsThisTab() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isClientTab_(sh)) { SpreadsheetApp.getUi().alert('Open a company tab first, or select a company on the Dashboard and use “Drafts for selected company”.'); return; }
  createDrafts_([sh]);
}
function draftsAllTabs() { createDrafts_(clientTabs_(SpreadsheetApp.getActive())); }
function draftsSelectedCompany() {
  const ss = SpreadsheetApp.getActive(); const sh = ss.getActiveSheet();
  if (sh.getName() !== SYS.SUMMARY) { draftsThisTab(); return; }
  const row = sh.getActiveRange().getRow();
  const name = String(sh.getRange(row, 1).getValue()).trim();
  const tab = name && ss.getSheetByName(name);
  if (!tab) { SpreadsheetApp.getUi().alert('Click a company name on the Dashboard first, then press the button.'); return; }
  createDrafts_([tab]);
}

function createDrafts_(tabs, silent) {
  const ui = silent ? null : SpreadsheetApp.getUi();
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const todo = []; let noContact = 0, others = 0;
  const data = tabs.map(function (sh) { return { sh: sh, t: read_(sh) }; });
  data.forEach(function (d) {
    d.t.rows.forEach(function (r, i) {
      if (/^y/i.test(String(r['W-9 Received']))) return;
      if (r['Request Status'] && r['Request Status'] !== RQ.NOT) return;
      if (String(r['Sender Email']).trim().toLowerCase() !== me) { others++; return; }
      if (!String(r['Contact Email']).trim()) { noContact++; return; }
      todo.push({ d: d, i: i });
    });
  });
  if (!todo.length) {
    if (ui) ui.alert('No W-9 request drafts to create for ' + me + '.\n\nA row needs: Request Status "Not sent", W-9 Received "No", Sender Email = your email, and a Contact Email.' +
      (noContact ? '\n• ' + noContact + ' of your rows have no Contact Email.' : '') + (others ? '\n• ' + others + ' rows belong to other senders (or have no Sender Email).' : ''));
    return 0;
  }
  const tl = timeline_();
  if (ui) {
    const early = today_() < tl.email ? '\n\nNote: the plan starts emails on ' + short_(tl.email) + '. You can still create them now.' : '';
    if (ui.alert('Create ' + todo.length + ' W-9 request drafts in your Gmail (' + me + ')?\n\nNothing is sent — you review and send them from Gmail → Drafts.' + early, ui.ButtonSet.YES_NO) !== ui.Button.YES) return 0;
  }
  const tpl = templatesMap_()['Request'];
  let w9 = null;
  if (/^y/i.test(String(setting_('Attach W-9 PDF')))) { try { w9 = UrlFetchApp.fetch(IRS_W9).getBlob().setName('Form W-9.pdf'); } catch (e) { log_('W-9 download failed', e.message); } }
  const touched = [];
  todo.forEach(function (x) {
    const r = x.d.t.rows[x.i];
    const td = tplData_(r, x.d.sh);
    const opts = { name: r['Sender Name'] || undefined };
    if (w9) opts.attachments = [w9];
    const dr = GmailApp.createDraft(String(r['Contact Email']).trim(), fill_(tpl.subject, td), fill_(tpl.body, td), opts);
    r['Thread ID'] = dr.getMessage().getThread().getId();
    r['Request Status'] = RQ.DRAFT;
    if (touched.indexOf(x.d) < 0) touched.push(x.d);
  });
  touched.forEach(function (d) { write_(d.sh, d.t.rows); });
  refreshSummary_(SpreadsheetApp.getActive());
  log_('Request drafts', todo.length + ' for ' + me);
  if (ui) ui.alert(todo.length + ' drafts created ✓\n\nOpen Gmail → Drafts, review each one and click Send.\nThen click “Sync status” here so the tracker turns blue.');
  return todo.length;
}

/* ================================ UNDO ==================================== */

/** Puts the selected rows back to "Not sent" and deletes their unsent Gmail drafts. Rows already emailed are left alone. */
function resetSelected() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isClientTab_(sh)) { ui.alert('Open a company tab and select the vendor rows to reset (or click the column letter to select all).'); return; }
  const t = read_(sh); const sel = sh.getActiveRange(); const idx = [];
  for (let row = Math.max(t.hr + 1, sel.getRow()); row < sel.getRow() + sel.getNumRows(); row++) if (t.rows[row - t.hr - 1]) idx.push(row - t.hr - 1);
  if (!idx.length) { ui.alert('Select the vendor rows to reset first (row ' + (t.hr + 1) + ' or below). Tip: drag down the Vendor column.'); return; }
  if (ui.alert('Reset ' + idx.length + ' row(s) to "Not sent" and delete their unsent drafts from YOUR Gmail?\nRows that were already emailed are skipped.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const draftsByThread = {};
  GmailApp.getDrafts().forEach(function (d) { try { const id = d.getMessage().getThread().getId(); (draftsByThread[id] = draftsByThread[id] || []).push(d); } catch (e) { } });
  let reset = 0, removed = 0, skipped = 0;
  idx.forEach(function (i) {
    const r = t.rows[i];
    if ((Number(r['Sent Count']) || 0) > 0) { skipped++; return; }
    const id = String(r['Thread ID'] || '');
    if (id && draftsByThread[id] && String(r['Sender Email']).trim().toLowerCase() === me) {
      draftsByThread[id].forEach(function (d) { d.deleteDraft(); removed++; });
    }
    r['Request Status'] = RQ.NOT; r['Thread ID'] = ''; r['Sent Count'] = 0; r['Last Contact'] = '';
    reset++;
  });
  write_(sh, t.rows); refreshSummary_(SpreadsheetApp.getActive());
  log_('Reset rows', reset + ' reset, ' + removed + ' drafts deleted, ' + skipped + ' already emailed (skipped)');
  ui.alert(reset + ' row(s) set back to "Not sent" ✓\n' + removed + ' draft(s) deleted from your Gmail.' + (skipped ? '\n' + skipped + ' row(s) skipped because they were already emailed.' : ''));
}

/* ============================== FOLLOW-UPS ================================ */

function followUpSelected() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isClientTab_(sh)) { ui.alert('Select vendor rows on a company tab first.'); return; }
  const t = read_(sh); const sel = sh.getActiveRange(); const idx = [];
  for (let row = Math.max(t.hr + 1, sel.getRow()); row < sel.getRow() + sel.getNumRows(); row++) idx.push(row - t.hr - 1);
  if (!idx.length) { ui.alert('Click one or more vendor rows (row ' + (t.hr + 1) + ' or below), then press “Follow up selected”.'); return; }
  const res = followUp_(sh, t, idx);
  write_(sh, t.rows); refreshSummary_(SpreadsheetApp.getActive());
  ui.alert(res.join('\n') || 'Nothing selected.');
}

/** A follow-up is due when its timeline date has arrived: 1 email sent → Follow-up 1 date, 2 → Follow-up 2, 3 → Follow-up 3 (final). */
function fuDue_(r, tl, today) {
  if (/^y/i.test(String(r['W-9 Received']))) return false;
  const sent = Number(r['Sent Count']) || 0;
  if (sent < 1 || sent > 3) return false;
  return today >= [tl.fu1, tl.fu2, tl.fu3][sent - 1];
}

function followUpAllDue(silent) {
  silent = silent === true;
  const ui = silent ? null : SpreadsheetApp.getUi();
  if (!silent) syncMine(true);
  const tl = timeline_(), today = today_();
  const me = Session.getActiveUser().getEmail().toLowerCase();
  let total = 0, replied = 0; const plans = [];
  clientTabs_(SpreadsheetApp.getActive()).forEach(function (sh) {
    const t = read_(sh); const idx = [];
    t.rows.forEach(function (r, i) {
      if (String(r['Sender Email']).trim().toLowerCase() !== me) return;
      if (!fuDue_(r, tl, today) || r['Request Status'] === RQ.FU_DRAFT) return;
      if (r['Vendor Replied']) { replied++; return; } // read the reply first
      idx.push(i);
    });
    if (idx.length) { plans.push({ sh: sh, t: t, idx: idx }); total += idx.length; }
  });
  if (!total) {
    if (ui) ui.alert('No follow-ups are due today for ' + me + '.\n\nSchedule: Follow-up 1 on ' + short_(tl.fu1) + ', Follow-up 2 on ' + short_(tl.fu2) + ', final notice on ' + short_(tl.fu3) + '.' +
      (replied ? '\n' + replied + ' vendor(s) replied without a W-9 — read their reply, then use “Follow up selected” if needed.' : ''));
    return 0;
  }
  if (ui && ui.alert('Create ' + total + ' follow-up drafts (replies in the same email threads)?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return 0;
  let made = 0;
  plans.forEach(function (pl) { made += followUp_(pl.sh, pl.t, pl.idx).filter(function (x) { return /created/.test(x); }).length; write_(pl.sh, pl.t.rows); });
  refreshSummary_(SpreadsheetApp.getActive());
  if (ui) ui.alert(made + ' follow-up drafts created ✓ Review them in Gmail → Drafts and click Send.');
  return made;
}

function followUp_(sh, t, idx) {
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const tpls = templatesMap_(); const out = [];
  idx.forEach(function (i) {
    const r = t.rows[i]; if (!r) return;
    const label = r['Vendor'] + ': ';
    if (/^y/i.test(String(r['W-9 Received']))) { out.push(label + 'W-9 already received — skipped'); return; }
    if (String(r['Sender Email']).trim().toLowerCase() !== me) { out.push(label + 'sender is ' + (r['Sender Email'] || 'blank') + ', not you — skipped'); return; }
    if (!r['Thread ID']) { out.push(label + 'no request email yet — click “Create W-9 drafts” first'); return; }
    let th = null; try { th = GmailApp.getThreadById(r['Thread ID']); } catch (e) { th = null; }
    if (!th) { out.push(label + 'email thread not found in your mailbox'); return; }
    const msgs = th.getMessages();
    if (msgs.some(function (m) { return m.isDraft(); })) { out.push(label + 'a draft is already waiting in this thread'); return; }
    const sent = msgs.filter(function (m) { return !m.isDraft() && from_(m, me); });
    if (!sent.length) { out.push(label + 'request not sent yet — send it from Gmail Drafts first'); return; }
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
  const msg = synced + ' of your vendor emails synced; ' + saved + ' W-9 attachments saved to Drive.';
  log_('Sync', msg + ' (' + me + ')', ss);
  if (silent !== true) { try { SpreadsheetApp.getUi().alert(msg + '\n\nUpload saved W-9s into the EP 1099 Tool so the client tracker updates too.'); } catch (e) { /* trigger */ } }
}
function dailySync_() { syncMine(true); }

/* =============================== TIMELINE ================================= */

function today_() { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function toDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(String(v)); return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function ymd_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
function short_(d) { return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getMonth()] + ' ' + d.getDate(); }
function holidays_(ss) {
  const raw = String(setting_('Timeline: Holidays (not working days)', ss) || HOLIDAYS_DEFAULT);
  const out = {}; (raw.match(/\d{4}-\d{1,2}-\d{1,2}/g) || []).forEach(function (x) { out[ymd_(toDate_(x))] = 1; }); return out;
}
let HOL_CACHE = null;
function isWorkday_(d, ss) { HOL_CACHE = HOL_CACHE || holidays_(ss); const w = d.getDay(); return w !== 0 && w !== 6 && !HOL_CACHE[ymd_(d)]; }
/** Weekend or holiday → next working day. */
function workday_(d, ss) { const x = new Date(d); while (!isWorkday_(x, ss)) x.setDate(x.getDate() + 1); return x; }
/** Timeline dates from the Settings tab (defaults if missing), each moved to a working day. */
function timeline_(ss) {
  ss = ss || book_();
  const out = {};
  TL.forEach(function (t) { out[t[0]] = workday_(toDate_(setting_(t[1], ss)) || toDate_(t[2]), ss); });
  return out;
}
/** Adds any Settings rows this version needs without touching existing ones. */
function ensureSettings_(ss) {
  const sh = ss.getSheetByName(SYS.SETTINGS); if (!sh) return;
  const have = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const want = TL.map(function (t) { return [t[1], t[2], 'Date (YYYY-MM-DD). Weekends and holidays move to the next working day.']; });
  want.push(['Timeline: Holidays (not working days)', HOLIDAYS_DEFAULT, 'Comma-separated dates that are not working days']);
  want.push(['W-9 Drive Folder URL (all clients)', '', 'Optional: one Drive folder that holds W-9s for every client (subfolders are searched). A company tab can also have its own folder.']);
  const add = want.filter(function (w) { return have.indexOf(w[0]) < 0; });
  if (add.length) { const r = sh.getRange(sh.getLastRow() + 1, 1, add.length, 3); r.setNumberFormat('@'); r.setValues(add); }
}

/* ============================ DRIVE W-9 CHECK ============================== */

function normV_(name) {
  let s = String(name || '').toLowerCase().replace(/&/g, ' and ');
  s = s.replace(/l\.l\.c\.?/g, ' llc ').replace(/[^a-z0-9 ]+/g, ' ');
  s = s.replace(/\b(llc|l l c|inc|incorporated|corp|corporation|co|company|ltd|limited|pllc|plc|lp|llp|pc|pa|the|dba)\b/g, ' ');
  return s.replace(/\s+/g, ' ').trim();
}
function fileNorm_(name) {
  return ' ' + normV_(String(name).replace(/\.[a-z0-9]{2,5}$/i, '').replace(/w-?\s?9|w-?\s?8\w*|form|signed|completed|final|copy/ig, ' ').replace(/[_\-,.]/g, ' ').replace(/\b(19|20)\d{2}\b/g, ' ')) + ' ';
}
function driveFiles_(folderUrl) {
  const out = [];
  // Each file is matched on its own name AND the names of the folders it sits in,
  // so "Gish PLLC / W9.pdf" and "W-9s / 2026 / Gish PLLC W-9.pdf" both work.
  const walk = function (folder, path, depth) {
    const it = folder.getFiles();
    while (it.hasNext() && out.length < 5000) {
      const f = it.next(); const nm = f.getName();
      if (!/\.(pdf|png|jpe?g|tiff?|heic|docx?)$/i.test(nm) && !/pdf|image/i.test(f.getMimeType())) continue;
      out.push({ name: (path.length ? path.join(' / ') + ' / ' : '') + nm, url: f.getUrl(), norm: fileNorm_(nm), pnorm: path.map(fileNorm_), isW9: /w-?\s?9/i.test(nm) || path.some(function (p) { return /w-?\s?9/i.test(p); }) });
    }
    if (depth >= 5) return;
    const sub = folder.getFolders(); while (sub.hasNext()) { const d = sub.next(); walk(d, path.concat(d.getName()), depth + 1); }
  };
  walk(DriveApp.getFolderById(folderId_(folderUrl)), [], 0);
  return out;
}
function nameIn_(key, words, norm) {
  if (norm.indexOf(' ' + key + ' ') >= 0) return true;
  return words.length >= 2 && words.every(function (w) { return norm.indexOf(' ' + w + ' ') >= 0; });
}
/** Best W-9 file for a vendor: vendor name in the file name, or in one of its folder names. W-9-named files win. */
function findW9_(vendor, files) {
  const key = normV_(vendor); if (key.length < 3) return null;
  const words = key.split(' ').filter(function (w) { return w.length > 1; });
  let best = null, score = 0;
  files.forEach(function (f) {
    let sc = 0;
    if (nameIn_(key, words, f.norm)) sc = 3;
    else if (f.pnorm.some(function (p) { return nameIn_(key, words, p); })) sc = 2;
    if (!sc) return;
    if (f.isW9) sc += 2;
    if (sc > score) { best = f; score = sc; }
  });
  return best;
}
function tabFolder_(sh, ss) { return meta_(sh).folder || String(setting_('W-9 Drive Folder URL (all clients)', ss) || '').trim(); }

function checkDriveTab_(sh, ss, files) {
  const folder = tabFolder_(sh, ss); if (!folder) return -1;
  files = files || driveFiles_(folder);
  const t = read_(sh); let found = 0;
  t.rows.forEach(function (r) {
    if (/^y/i.test(String(r['W-9 Received']))) return;
    const f = findW9_(r['Vendor'], files); if (!f) return;
    r['W-9 Received'] = 'Yes (Drive)'; r['Request Status'] = RQ.RECEIVED;
    if (String(r['Notes']).indexOf(f.url) < 0) r['Notes'] = (r['Notes'] ? r['Notes'] + ' ' : '') + f.url;
    found++;
  });
  const m = meta_(sh); m.checked = ymd_(today_()) + ' (' + files.length + ' files)'; setMeta_(sh, m);
  write_(sh, t.rows);
  return found;
}
function checkDriveThisTab() {
  const ui = SpreadsheetApp.getUi(); const ss = SpreadsheetApp.getActive(); const sh = ss.getActiveSheet();
  if (!isClientTab_(sh)) { ui.alert('Open a company tab first.'); return; }
  if (!tabFolder_(sh, ss)) {
    const r = ui.prompt('W-9 folder for ' + sh.getName(), 'Paste the Google Drive folder link that holds this client’s W-9s (view access is enough; subfolders are searched):', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK || !/folders\/|id=/.test(r.getResponseText())) return;
    const m = meta_(sh); m.folder = r.getResponseText().trim(); setMeta_(sh, m);
  }
  let found;
  try { found = checkDriveTab_(sh, ss); } catch (e) { ui.alert('Couldn’t open that Drive folder.\n\nMake sure it is shared with ' + Session.getActiveUser().getEmail() + ' and the link is a folder link.\n\n(' + e.message + ')'); return; }
  refreshSummary_(ss);
  ui.alert(found + ' new W-9' + (found === 1 ? '' : 's') + ' found in Drive ✓\n\nMatched vendors are marked “Yes (Drive)” with the file link in Notes, and will not be emailed.\nFiles are matched by file name or folder name (e.g. “Vendor Name W-9.pdf” or “Vendor Name / W9.pdf”).');
}
function checkDriveAll_(ss) {
  let total = 0; const cache = {};
  clientTabs_(ss).forEach(function (sh) {
    const folder = tabFolder_(sh, ss); if (!folder) return;
    try { cache[folder] = cache[folder] || driveFiles_(folder); total += Math.max(0, checkDriveTab_(sh, ss, cache[folder])); }
    catch (e) { log_('Drive check failed', sh.getName() + ': ' + e.message, ss); }
  });
  return total;
}
function checkDriveAll() {
  const ss = SpreadsheetApp.getActive(); const n = checkDriveAll_(ss); refreshSummary_(ss);
  const none = clientTabs_(ss).filter(function (sh) { return !tabFolder_(sh, ss); }).map(function (sh) { return sh.getName(); });
  SpreadsheetApp.getUi().alert(n + ' new W-9s found in Drive ✓' + (none.length ? '\n\nNo W-9 folder set for: ' + none.join(', ') + '.\nOpen that tab and click “Check Drive for W-9s” to paste its folder link.' : ''));
}

/** Called from the EP 1099 Tool ("Get W-9s from Drive"): scans the folder as the signed-in user and hands the matches back to the app. */
function w9Page_(p) {
  let names = [];
  try { names = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p.d)).getDataAsString()); } catch (e) { return page_('Couldn’t read the vendor list', 'Try again from the app.', false); }
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const folder = p.f || String(setting_('W-9 Drive Folder URL (all clients)', ss) || '');
  if (!folder) return page_('No W-9 folder', 'Paste the client’s W-9 Drive folder link on the tracker in the app first.', false);
  let files;
  try { files = driveFiles_(folder); } catch (e) { return page_('Couldn’t open the Drive folder', 'Make sure the folder is shared with ' + Session.getActiveUser().getEmail() + ' and the link is a folder link. (' + e.message + ')', false); }
  const hits = [];
  names.forEach(function (n) { const f = findW9_(n, files); if (f) hits.push([n, f.name, f.url]); });
  const sh = p.c ? ss.getSheetByName(tabName_(p.c)) : null;
  if (sh && isClientTab_(sh)) { const m = meta_(sh); m.folder = folder; setMeta_(sh, m); try { checkDriveTab_(sh, ss, files); refreshSummary_(ss); } catch (e) { } }
  const back = /^https:\/\//.test(String(p.r || '')) ? String(p.r).split('#')[0] : '';
  const payload = Utilities.base64EncodeWebSafe(JSON.stringify({ c: p.c || '', m: hits.slice(0, 60) }));
  const rows = hits.map(function (h) { return '<tr><td style="padding:4px 12px 4px 0"><b>' + esc_(h[0]) + '</b></td><td style="padding:4px 0;color:#55627d">' + esc_(h[1]) + '</td></tr>'; }).join('');
  return HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,sans-serif;max-width:680px;margin:40px auto;padding:24px;border:1px solid #d5e6e1;border-radius:12px">' +
    '<div style="font-weight:700;font-size:20px;color:#11845b">' + hits.length + ' of ' + names.length + ' vendors have a W-9 in Drive</div>' +
    '<p style="color:#55627d">' + files.length + ' files checked in the folder and its sub folders. Matched on file and folder names.</p>' +
    (rows ? '<table style="font-size:13px;border-collapse:collapse">' + rows + '</table>' : '<p>No matches. Put each W-9 in a folder named after the vendor, or name the file like “Vendor Name W-9.pdf”.</p>') +
    (back && hits.length ? '<p style="margin-top:20px"><a target="_top" style="background:#00c49a;color:#1c2a4d;font-weight:700;padding:10px 16px;border-radius:8px;text-decoration:none" href="' + back + '#w9=' + payload + '">Save these W-9s in the EP 1099 Tool</a></p>' : '') +
    '</div>').setTitle('EP 1099 – W-9s in Drive');
}
function esc_(s) { return String(s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); }

/* ============================== DASHBOARD ================================= */

function refreshSummary() { const ss = SpreadsheetApp.getActive(); refreshSummary_(ss); ss.setActiveSheet(ss.getSheetByName(SYS.SUMMARY)); }

/** Per-company numbers for the management dashboard. */
function tabStats_(sh, tl, today) {
  const t = read_(sh); const m = meta_(sh);
  const c = { name: sh.getName(), year: m.year, paid: m.paid, checked: m.checked, gid: sh.getSheetId(), sender: '',
    n: 0, onFile: 0, toEmail: 0, draft: 0, waiting: 0, due: 0, fu: [0, 0, 0], noResp: 0, replied: 0, emailed: 0 };
  t.rows.forEach(function (r) {
    c.n++; if (!c.sender && r['Sender Email']) c.sender = r['Sender Email'];
    const sent = Number(r['Sent Count']) || 0;
    if (sent > 0) c.emailed++;
    if (/^y/i.test(String(r['W-9 Received']))) { c.onFile++; return; }
    const s = String(r['Request Status'] || RQ.NOT);
    if (/draft/i.test(s)) c.draft++; else if (sent === 0) c.toEmail++;
    if (sent > 0) c.waiting++;
    if (fuDue_(r, tl, today)) { c.due++; c.fu[sent - 1]++; }
    if (sent >= 4) c.noResp++;
    if (r['Vendor Replied']) c.replied++;
  });
  return c;
}

function refreshSummary_(ss) {
  let sh = ss.getSheetByName(SYS.SUMMARY);
  if (!sh) sh = ss.insertSheet(SYS.SUMMARY, 0);
  HOL_CACHE = null;
  const tl = timeline_(ss), today = today_();
  const tabs = clientTabs_(ss).map(function (t) { return tabStats_(t, tl, today); });
  const T = { n: 0, onFile: 0, toEmail: 0, draft: 0, waiting: 0, due: 0, noResp: 0, replied: 0, emailed: 0, paid: 0, checked: 0, fu: [0, 0, 0] };
  tabs.forEach(function (c) { ['n', 'onFile', 'toEmail', 'draft', 'waiting', 'due', 'noResp', 'replied', 'emailed'].forEach(function (k) { T[k] += c[k]; });
    T.paid += Number(c.paid) || 0; if (c.checked) T.checked++; c.fu.forEach(function (x, i) { T.fu[i] += x; }); });
  const W = 15;
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  const need = 24 + tabs.length; if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart(); all.clearContent(); all.clearFormat(); sh.clearConditionalFormatRules();
  all.setFontFamily(FONT).setFontSize(10).setFontColor(C.ink).setVerticalAlignment('middle');
  sh.setFrozenRows(0);

  // Title
  sh.getRange(1, 1, 3, W).setWrap(false);
  sh.getRange(2, 1, 1, W).setBorder(null, null, true, null, null, null, C.teal, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange(1, 1).setValue('Enduring Planet  ·  1099 Management Dashboard').setFontSize(18).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange(2, 1).setValue('Updated ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE MMM d, yyyy h:mm a') + '   ·   click “Refresh dashboard” for the latest numbers   ·   tracker ' + VERSION).setFontColor(C.muted);
  sh.setRowHeight(1, 40); sh.setRowHeight(2, 24); sh.setRowHeight(3, 46); sh.setRowHeight(4, 8);

  // Management numbers
  const pct = T.n ? T.onFile / T.n : 0;
  const kpi = [['Clients', tabs.length, C.ink], ['Vendors paid', T.paid || '—', C.ink], ['1099-eligible vendors', T.n, C.ink], ['W-9 on file / received', T.onFile, C.greenInk],
    ['To email', T.toEmail + T.draft, C.redInk], ['Emailed – waiting', T.waiting, C.blueInk], ['Follow-up due today', T.due, C.orangeInk], ['No reply after 3 follow-ups', T.noResp, C.redInk], ['W-9s collected', pct, C.greenInk]];
  kpi.forEach(function (k, i) {
    sh.getRange(5, i + 1).setValue(k[0]).setFontSize(9).setFontWeight('bold').setFontColor(C.navy).setBackground(C.mint).setHorizontalAlignment('center').setWrap(true);
    sh.getRange(6, i + 1).setValue(k[1]).setFontSize(20).setFontWeight('bold').setFontColor(k[2]).setBackground(C.white).setHorizontalAlignment('center');
  });
  sh.getRange(6, kpi.length).setNumberFormat('0%');
  sh.getRange(5, 1, 2, kpi.length).setBorder(true, true, true, true, true, false, C.teal, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(5, 32); sh.setRowHeight(6, 40); sh.setRowHeight(7, 12);

  // Timeline
  const st = function (done, date, overdueAfter, active) {
    if (done) return 'Done';
    if (today > (overdueAfter || date)) return 'Overdue';
    if (today >= date || active) return 'In progress';
    return 'Upcoming';
  };
  const toSend = T.toEmail + T.draft;
  const fuRow = function (i, d, label) {
    const dueNow = today >= d ? T.fu[i] : 0;
    return [label, short_(d), today < d ? 'Upcoming' : (dueNow ? 'Due now' : 'Done'), today < d ? 'Vendors with no W-9 get a reminder in the same email thread' : (dueNow ? dueNow + ' follow-up drafts to create and send' : 'Nothing outstanding')];
  };
  const phases = [
    ['1. Kick-off — upload GLs, confirm 1099 vendors', short_(tl.start), today >= tl.start ? (tabs.length ? 'Done' : 'In progress') : 'Upcoming', tabs.length + ' client' + (tabs.length === 1 ? '' : 's') + ' in the tracker, ' + T.n + ' vendors need a 1099'],
    ['2. W-9 inventory — check Drive for W-9s on file', short_(tl.inv) + ' (latest ' + short_(tl.invMax) + ')', st(tabs.length > 0 && T.checked === tabs.length, tl.start, tl.invMax), T.checked + ' of ' + tabs.length + ' clients checked · ' + T.onFile + ' W-9s on file, ' + (T.n - T.onFile) + ' missing'],
    ['3. Send W-9 request emails', short_(tl.email), st(T.n > 0 && toSend === 0, tl.email, tl.fu1), T.emailed + ' emailed · ' + toSend + ' still to send'],
    fuRow(0, tl.fu1, '4. Follow-up 1'), fuRow(1, tl.fu2, '5. Follow-up 2'), fuRow(2, tl.fu3, '6. Follow-up 3 (final notice)'),
    ['7. W-9 collection deadline', short_(tl.due), st(T.n > 0 && T.onFile === T.n, tl.email, tl.due), T.onFile + ' of ' + T.n + ' W-9s collected (' + Math.round(pct * 100) + '%)' + (T.noResp ? ' · ' + T.noResp + ' no reply after 3 follow-ups' : '')],
    ['8. File 1099s with the IRS and send to vendors', short_(tl.file) + ', ' + tl.file.getFullYear(), today > tl.file ? 'Overdue' : 'Upcoming', 'Jan 31 falls on a weekend in 2027, so the due date is the next working day']
  ];
  const TR = 8;
  sh.getRange(TR, 1).setValue('Timeline — where we are').setFontSize(13).setFontWeight('bold');
  sh.getRange(TR + 1, 1, 1, 9).setBackground(C.teal).setFontColor(C.navy).setFontWeight('bold');
  sh.getRange(TR + 1, 1).setValue('Phase'); sh.getRange(TR + 1, 4).setValue('Target date'); sh.getRange(TR + 1, 5).setValue('Status'); sh.getRange(TR + 1, 6).setValue('Progress');
  let here = phases.findIndex(function (p) { return p[2] !== 'Done'; }); if (here < 0) here = phases.length - 1;
  phases.forEach(function (p, i) {
    const r = TR + 2 + i;
    sh.getRange(r, 1).setValue(p[0] + (i === here ? '   ◀ we are here' : '')).setFontWeight(i === here ? 'bold' : 'normal');
    sh.getRange(r, 4).setValue(p[1]); sh.getRange(r, 6).setValue(p[3]).setFontColor(C.muted);
    const col = { 'Done': [C.green, C.greenInk], 'In progress': [C.blue, C.blueInk], 'Due now': [C.orange, C.orangeInk], 'Overdue': [C.red, C.redInk], 'Upcoming': [C.grey, C.greyInk] }[p[2]];
    sh.getRange(r, 5).setValue(p[2]).setBackground(col[0]).setFontColor(col[1]).setFontWeight('bold').setHorizontalAlignment('center');
    if (i === here) sh.getRange(r, 1, 1, 4).setBackground(C.mint);
    sh.getRange(r, 1, 1, 9).setBorder(null, null, true, null, null, null, C.line, SpreadsheetApp.BorderStyle.SOLID);
  });
  sh.getRange(TR + 1, 1, phases.length + 1, 9).setWrap(false);

  // Clients
  const CR = TR + phases.length + 3;
  const head = ['Client', 'Tax Year', 'Sender', 'Vendors paid', '1099-eligible', 'W-9 on file', 'To email', 'Draft ready', 'Emailed – waiting', 'Follow-up due', 'No reply (3 follow-ups)', 'Vendor replied', '% W-9 collected', 'Drive checked', 'Open'];
  sh.getRange(CR, 1).setValue('Clients').setFontSize(13).setFontWeight('bold');
  sh.getRange(CR + 1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor(C.navy).setBackground(C.teal).setWrap(true).setHorizontalAlignment('center');
  sh.getRange(CR + 1, 1).setHorizontalAlignment('left'); sh.setRowHeight(CR + 1, 34);
  if (tabs.length) {
    const rows = tabs.map(function (c) { return [c.name, c.year, c.sender, c.paid || '', c.n, c.onFile, c.toEmail, c.draft, c.waiting, c.due, c.noResp, c.replied, c.n ? c.onFile / c.n : 0, c.checked || 'Not yet', '']; });
    const r = sh.getRange(CR + 2, 1, rows.length, head.length);
    r.setValues(rows).setHorizontalAlignment('center').setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
    tabs.forEach(function (c, i) { sh.getRange(CR + 2 + i, head.length).setFormula('=HYPERLINK("#gid=' + c.gid + '","Open →")'); });
    sh.getRange(CR + 2, 1, rows.length, 1).setFontWeight('bold').setHorizontalAlignment('left');
    sh.getRange(CR + 2, 4, rows.length, 9).setFontSize(12).setFontWeight('bold');
    sh.getRange(CR + 2, 13, rows.length, 1).setNumberFormat('0%');
    sh.setRowHeights(CR + 2, rows.length, 30);
    const col = function (c) { return sh.getRange(CR + 2, c, rows.length, 1); };
    const gt0 = function (c, bg, ink) { return SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(bg).setFontColor(ink).setRanges([col(c)]).build(); };
    sh.setConditionalFormatRules([gt0(6, C.green, C.greenInk), gt0(7, C.red, C.redInk), gt0(8, C.yellow, C.yellowInk), gt0(9, C.blue, C.blueInk), gt0(10, C.orange, C.orangeInk), gt0(11, C.red, C.redInk),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Not yet').setBackground(C.yellow).setFontColor(C.yellowInk).setRanges([col(14)]).build(),
      SpreadsheetApp.newConditionalFormatRule().setGradientMinpointWithValue('#fde2e6', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue('#fff0cc', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
        .setGradientMaxpointWithValue('#dcf4e8', SpreadsheetApp.InterpolationType.NUMBER, '1').setRanges([col(13)]).build()]);
    const tr = CR + 2 + rows.length;
    sh.getRange(tr, 1, 1, head.length).setValues([['ALL CLIENTS', '', '', T.paid || '', T.n, T.onFile, T.toEmail, T.draft, T.waiting, T.due, T.noResp, T.replied, pct, T.checked + ' of ' + tabs.length, '']])
      .setBackground(C.mint).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    sh.getRange(tr, 13).setNumberFormat('0%'); sh.getRange(tr, 1).setHorizontalAlignment('left');
  } else {
    sh.getRange(CR + 2, 1).setValue('No clients yet — in the EP 1099 Tool, open a company and click “Send to email tracker”.').setFontColor(C.muted);
  }
  [250, 80, 210, 100, 110, 110, 100, 100, 110, 110, 120, 100, 110, 150, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setTabColor(C.teal);
  addButtons_(sh, [['today', 'runTodayTasks', 1, 6], ['company', 'draftsSelectedCompany', 3, 0], ['allfu', 'followUpAllDue', 5, 40], ['drive', 'checkDriveAll', 7, 30], ['refresh', 'refreshSummary', 9, 40]]);
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
function installDailyAuto() {
  removeTriggers_('dailySync_'); removeTriggers_('dailyAuto_');
  ScriptApp.newTrigger('dailyAuto_').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('Daily automation is on for ' + Session.getActiveUser().getEmail() + ' (~7am, working days).\n\n' +
    'Each morning it: syncs your sent emails and replies, checks Drive for new W-9s, and — from the email start date — creates the W-9 request and follow-up DRAFTS that are due.\n' +
    'Nothing is ever sent automatically. You get an email when drafts are waiting.');
}
function installDailySync() { installDailyAuto(); }
function dailyAuto_() { if (isWorkday_(today_())) todayTasks_(true); }
function runTodayTasks() { todayTasks_(false); }

function todayTasks_(silent) {
  const ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const tl = timeline_(ss), today = today_();
  syncMine(true);
  let w9 = 0; try { w9 = checkDriveAll_(ss); } catch (e) { log_('Drive check failed', e.message, ss); }
  let req = 0, fu = 0;
  if (today >= tl.email) { req = createDrafts_(clientTabs_(ss), true); fu = followUpAllDue(true); }
  refreshSummary_(ss);
  const msg = 'W-9s found in Drive: ' + w9 + '\nNew W-9 request drafts: ' + req + '\nFollow-up drafts: ' + fu +
    (today < tl.email ? '\n\nEmail requests start on ' + short_(tl.email) + ' — no drafts created before then.' : '');
  log_('Today’s tasks', msg.replace(/\n/g, ' | '), ss);
  if (silent) {
    if (req + fu > 0) MailApp.sendEmail(Session.getActiveUser().getEmail(), '[EP 1099] ' + (req + fu) + ' W-9 email drafts waiting in your Gmail',
      'Review and send them from Gmail → Drafts.\n\n' + msg + '\n\nDashboard: ' + ss.getUrl());
  } else SpreadsheetApp.getUi().alert('Today’s tasks done ✓\n\n' + msg + (req + fu ? '\n\nReview the drafts in Gmail → Drafts and click Send.' : ''));
}
function removeTriggers_(fn) { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); }); }

/* ================================ HELPERS ================================= */

function tplData_(r, sh) {
  const up = String(setting_('Secure Upload Link') || '').trim();
  return {
    'Contact Name': String(r['Contact Name'] || '').trim() || 'there', 'Vendor': r['Vendor'], 'Company': tabCompany_(sh),
    'Tax Year': tabYear_(sh), 'Due Date': fmt_(timeline_().due), 'Upload Line': up ? ', or upload it securely here: ' + up : '',
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
function book_() { return SpreadsheetApp.getActive() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')); }
function setting_(name, ss) {
  ss = ss || book_();
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
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground(C.teal).setFontColor(C.navy).setFontFamily(FONT);
  sh.setFrozenRows(1); return sh;
}
function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg, 'EP 1099', 5); } catch (e) { } }
function log_(action, details, ss) {
  try { const sh = (ss || SpreadsheetApp.getActive()).getSheetByName(SYS.LOG); if (sh) sh.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]); } catch (e) { }
}

/* ======================= BUTTON IMAGES (do not edit) ======================= */
const BTN = {
  drafts: 'iVBORw0KGgoAAAANSUhEUgAAALMAAAAiCAMAAAAj8BgnAAAASFBMVEUAzqAOdnMKiHwAAAAAx5wgFUIAxJgAf38AsJ0A/38AzZ8Kj4AhDz8AxJocJ0sBuZQZN1MTV2MeG0UWR1sImIQFp4sA1KIQZmub/XSbAAAAGHRSTlP5//8Anf8nAg8CQ////v/////////////2j1zfAAADCklEQVR42u1Ya5MiIQwExtF9HCTh+f//6XWYh467e6V7n6gyZSEMkDSdhlHMNL1P08fJj2Cnj47WTG/T28mYITAbcwLaybxPZ1P9KFbNeXo3l/MgJG9Uny/mcqp+JKunizkbP5YB8Wk4zMMh7qhfmM2X8sm45j8wZ9uaLc9Czq7clWrkWnPHgU7sN9PJirjfYiYJcZ5npp8nUvvGffqT0BNiL+cVFqUY8Ojgq+nALy4xMFoSKb/BLDGIy23FfFtqZanxbOttW80GpppD4KJlXvEBb+HYrth8bTHtjc2DKSE4X3MMpd75fQAz4on2EQiX3IKt1Diw8pYTKukTHRySCPmsFXedmGsLCaWNW5ZSbAbAtyZcse2YrTjLyVvGIqWYnBC2WS3F9hYnehiz3TnC0mHWc2ThKN5YOIS3snQEysg6hLSCJo62JnaxVYmyCSZKrYI51xwCZdKOoN+JBb7ZO21ivbAoFHqg8jBmkFJ2zFyIQBp40JXQortmFB5aSaFBEruoEC3hg1VaQ2rqzYGFNT48Im3cMUMLVLrLnqBFTlvJZH6Qx/eYd1IKKKxeWVKDggtEApPaKQWxvSNu421MGSQnXlXCEBk29Bx5c/nZK10bmoBVbjBXy4a59IxBMO5xPTvlYmNFuRVIQ6RJ9mlOLiu5nWfFnLSn0SZoFqyyBdXCihleHO3yPmBu8A1aWnbKzQ1m7KSGpUT3MGZlzxGVRosbrxFr9ZkgM6eyWDDXVayV8jZTGSvVdaH6RRuEqcsK77WhmI1TFbgjz7VP8tc98cBZlzkiNTECc8xdwTO2BM8FcbAZg/rSWkMEfKG6n9BBtY3o17MtY2qY9yPgZg/Obd0ymLXzjDJJLpq/EO0T70HNDOMIK7ycIP2sSxaZxHMnAcFQ08B61t0ID1tBgaBvf6QDuV2Pd7iyTTXT3UCJGglx6hrMMpJCKvJkn3y/L3uW7l8lus1p66OvZ/+h626qP3qgg8vbEPT8O+X1W/SF+YX59Y/wX4hHvCsY8U5mxLuvEe8YR7zLHfDO/C+Xcjq81br6+AAAAABJRU5ErkJggg==',
  company: 'iVBORw0KGgoAAAANSUhEUgAAAQgAAAAiCAMAAACUaLcpAAAASFBMVEUIl4MAzaALhnsAAAAgFUMAx5wAxJgAf38AsJ0A/38AzZ8hDj8WT2AAxJocJ0sBupUZN1MeG0UTV2MWRlsNd3MA1KIQZ2sFp4tPoWy1AAAAGHRSTlP/+v8A/50nAg8CQ////v////////////8nb1+UAAAEGUlEQVR42u1a23LjIAwF7KTb7iLu8P9/ukfC+NImu8lDpw8xMyW2kAQ6SEK1rabpfZo+LvZl2+VDIFDTr+nXRevXBULrCyCY1Pt01cW+dCv6Or2rt+sru8NwiuubersU+/KtXN7UVZ84wCeu6nICwUBc1AlCbycQ3wTE9wSa/ikgmlFKNXdHpMbY9N2RRyeO1T1qs4n5h4CIM4VAvt4EiQJRu7W2PvLgxHn27lGb6+/0U0CQj9FTiMu9LNlJX3KABYXXpjfqNnKk2fVqo/WLwXykb3pXSZ4s7IAYAvtJ3Mqv95LOHmn6H3F2Bwie2UUKprQYVaVsgQsl7Fb2AKlW7qOSu+E4fcTZnPCrWL5LirNEZmQiS+YVCMO87ESu8pXLXW9Z2XDh8w4IkzCgOLI8JcN6osneG1u544BT3idoVMyYVOcwWCQWJt52J87uAoFVOk+xNOKWHaUYaa4l8q333AdlZvacNNyIR1wNFFMgtUjWoRDiubgUfPTsaQJE5/XNgk7Q5JLoNdYPtgg6b8EIkoB1kCmKoJAClEOQmwi2kmdiQe9Khj4Q1eAItZiZHIyanwYCKhICn3yDkzlZBhbPFrAJSfq4c0WxDfy5HCStQGpQyTuOnqYV1sdXzGEYg4o7kOyqd7A10QbpNLIQAwsG1g9bkZHgSE1mArthEnbABwjxtEmgAs2xAKQyILyZm/4DBLE5gW1gh5cNFwtG2BqAnWrb5YhiJF9WMbNLLtnXp+y0eA3RnEWN6Xchlbg4/6J3sBkVurYFCLMkYwdTBRYF42KBA9XCtCVNMY2DCy3KnWa9TmfoqXN9Nkewb0toKOw5fC4j4jsQbl2bgvuFcVLwiN4DQZv2LLkXBiOWkGLUACLybbYDCNHr7MIWlaL/A1E19jvbIxA2zckovrDiFqICrgWSegoIuL/y1CNdCSUW5IcDEJozOjuiuRcaq/YGRkNeWEqxyq2hUaChraGhZcFlZXsgNDoQwyNGaDio1s4fgeD9ofTU8YkigisJgV0ghEek1ENDCoDMyc1l+sMZqe1Lg12y3MBPnFR5s9PMSXFuZSTLFNMsSVSSZel6N7ZbydIfkuUMIPy8AdGTpZWsLAVA4mCoc+pAhvxkHYFDiM9A27zsK4dczQIEyYnCy2vbGSVAyMh6fC6S3QDP56Jbjk/OFp3ZMLicZPj4JDB0vXqwCX+uFO8en5HYIyhLBxzTcny2hPM0sgNJlzuW8W7FdxsIx+1GAdP/RtXkvtQ1x1rHOfu1Dvo8+rnQ2ulfyW6v6cskbtdJJv8kuHUcc/E1/vtci9ubzfyhoF4EiB6ddxpOPfMizyMOQXRj+Hwwcz6hehiI8+Ht8vD2fJy/PM4/X/AsL3jOV37LK7/zJfDyEvj8LGD5LOD8UKR/KPIXoyvVNz0CoK8AAAAASUVORK5CYII=',
  followup: 'iVBORw0KGgoAAAANSUhEUgAAALsAAAAiCAMAAAAwJ1jTAAAASFBMVEUfLVDV2t8EEzlNWHSSmapncIg3RGO1usUAAAB5gpYFyZ8CxJgCzKAAf38AsJ0A/39W1rlz3ca+wswHzaP9/v4Aw5gQyKAYJ0pjI3LEAAAAGHRSTlP//v///////wD/nSfpAg8C+vz/Q//+/v+KIvhSAAACfUlEQVR42u2Z23bjIAxFj7j5kqRpUoj+/09HErZxZ9JZ9bQvdI2ejKqIjTgoDsU4XsfxbS492fxm1Bhfxhclf/RjSi/UI67j6VweN8r9GN0e5Xwar7gI+mvuzV4F/oLLfO4PXeDP8wWn8sg92qOcMJdbl+y3MqMU6pKdSkGnklHR/Dx2Gsx2aqIYKQ9++q6Zj6QKPhxg9w5ObPcRghtycPgu9g9SPcWMLh1gj0CMWukdOws70rexP081Of9J58fsW3TwgJc1EEPrnnYuq5EuUPS0E1aOftCnISKFTXFgpEEfEpu3pbIK0ZQkaQ6JET21MK1iOMoO1bs+sPeO7zv2zaUpRUqh1YWco+oiQKKwqs67FD2LO8kDXFzYJ8dR5Cmf8ZoU5FkMIW9hOlmSVIc0w6p3Ik0kM/uVHfmdKwfmKIMFkZiliFIplVisUWaJNYKUWBfHtKRiSzXpSDdl3fAlDFQjEvxxvedBp1mRK3tzZTjRglQLuL9j17pruEbRso8MH0jzamUloKaqI980WrdwC7Nl7hT8Oc2k5UT9jV3qnWLgwAthZacn7CLsZIy1JnFNVUfhCbv9gYYvsO8EYn1G3M2lQw4y3uRoWzyxaQZ7zegKgkpMxSBDWlJJzmW0aGZqmsn5nv9JM2tH/f2sYnNZ79E6g9v3QHJ61qzuIpJ2VmX2mLQCyZm76l1IzR+2sypu9pFamJ1VPsbOf/RIyAKCuVeX1FWbdNwkI1Ee/m4NRQq265HSAdmaofRCVX5uqTgtfu2R1kxVd2tY7ZGH+szX3/IA6vVdrJ7VTtmtQfx/B/6h7B3/5uv5t3bPdxw93y31fKfX811qz3fYHf/v4BeC1b/J9Hc7UwAAAABJRU5ErkJggg==',
  allfu: 'iVBORw0KGgoAAAANSUhEUgAAALUAAAAiCAMAAAAu7mhgAAAASFBMVEVNWHMEEjkeLE7W3OCUmqtqc4o4RGO0ucQAAAB6g5cFyZ8CxJgCzKC9wswAf38AsJ0A/39W1rlz3cYHzaP9/v4Aw5gQyKAYJ0qnCQPQAAAAGHRSTlP////+/////wD/nSfp/wIPAvr8Q//+/v8Bp5//AAACcUlEQVR42u2YDXPbIAyGhcSH7Xy0TU34//90r4RjnK31krvuNrbprg4GJB6EpCbQMLwOw/uU+5Dp3XhpeBlelPnagyg3eAd6HQ7HfL3I3IPI5ZqPh+GVzoB+m/uRN2Cf6Twde4IG9nE60yFf577kmg805Utn1Jc8Uc7SGbXkTN0FiIbI30UtKVnkSAgyJxq3Y5783gH6sB23yfsaX0dNzrExsEuzr+2bBBd3bJIrzm9fw/L49dRSIMmo8eH5DnN0tOPqqvdbqL2LdSUprL5eqQOzN+qxnbwn5rBUIqFSCC8n9Nlgo5YQS8SQBZ1o0LXAadbw8DCHvkRcmOQJanIeYSE/UAdXKHIBNbtRX5XflUA2WXfAOCSXEFhEFikrtUTHgVUhot+7Itao0qxFVyI0g65MgcoT1NA44c9/T43jR19ksn3VdWofuXEbIaQA2IE0avMCduOxzTDjSBKmSvPSQk0ajR4jwlG99kSEWL7ZyvfUSTluPlnWwXlqEtxi3VAMruqu1KZlreSiwO1jslWg7pu1JZhUk+H14J+gZszH6ej6j1AHFb9Sn35CLcgNTkTBhj6h9rOMGHP+YWrAOYipWg35KEKQKCeN8KVoiGx8vR8hCA+M+MK8Iq3WWoRYcHxYfD6hRm4lCAF28TVvs7FwzcGCda2FtFkTS1MC+y472aj24gzIlmrNGnZ0y0ZNxsd9LbHYXHgDOYED94XuKp++Sm2FWvmspFVltmrSKl8Zl8et8tkc3QVvyv5qDRuslU9ggaPv43vIA/+Q/lN/FXUZ/7Vvqn86dZe/wPr8tdvnzUKftzh93pj1eTvZ501wl7fu3wBZZLLzTvzT/wAAAABJRU5ErkJggg==',
  sync: 'iVBORw0KGgoAAAANSUhEUgAAAIMAAAAiCAMAAABLA5ofAAAASFBMVEXX3eEcKk1PWnWWna0FEzpmcIc4RGIAAAC2usUFyZ8CxJgCzKC9wct5gpYAf38AsJ0A/39W1rlz3cYHzaP9/v4Aw5gQyKAZJ0qL7C8eAAAAGHRSTlP+////////AP+dJ+n//wIPAvr8Q//+/v8enQNwAAAB80lEQVR42uWY7W7CMAxFbSdOU2iBsbR5/zfddVLKNk1Tux800iw1wmDsE3/QBuq6W9e999MR0r+X6NRdu6sRzK8Xo0D0jm7d6TzNd0qvF7rP0/nU3egChLd0lLwB4kKX/nwcAiDO/YVO05yOlHk6UT/dD2W4Tz1NEx3KYAAHl8KK0TBDHFUdHcqg3iTvgCBV+k3fzUA5a3RO9jB4pt/03QyRHw60cGDFvgbl4Mz9GJjlYSqcWShKziIuuQAVGXjoI67kbFkM9+Qh1HZQj2ARVSHmDAEbiYf7BZLgV6E4bx9qElZlH5LjqrMfq5NqmLczJJetHbAf0BB8aCLOYiDYqueI4NVw4EBFibn4t2uwl4su3lWG1XD7XJBTzvZ9c8IISrakAK2kZrVjbylLFRbxkfFspIseHgyr4b7fh4BcYtcSOZSk/8SA5hBmvFs2b1bqnDEM3xhWw839oBFlCCgBvOTiB7VYGL7UItWE6xLTWTO7/GQQtOGAMj4NNzOgGVCK4sW6jT7n4VtPWqfZu+g5iahYEF51zEXpVZan4WYGJwHOYq14mUPkYcBGsvs6m4QE14nFVHKZTYnBUlZ1Uqxj1k+Gf7lfyE/wL71nqWDaD2YIWegf37tfzNDAs1wLz7QtPNu3cMZp4azXwpm3hbN/A/+BfADC1jynRnHmZQAAAABJRU5ErkJggg==',
  undo: 'iVBORw0KGgoAAAANSUhEUgAAAJgAAAAiCAMAAACUTeAAAAAASFBMVEXy2t7bmKOrCiTBS1/ReIe3KkIAAADovMO+QlatHja6KkF/AAC1Jz60KUG+KUCqVVX/AAD/AH/+/f2yHDW5MUi0IzvGWGrLZ3hd14pLAAAAGHRSTlP+/v////cA//8cmgKmPkMDAgL//v78//6eNdC0AAACMklEQVR42u2Yf2/cMAiG8a/Ed9d2bQDz/b/pAOfSTXOr5rQ/vKmWTnJIDI+BvLkEluVtWV6ujJMMvr44Eiw/loth0STD0C4KBa/LpSAF2CYZEAjLZXmFm3LJNFiOJkp2g9sTyjbZEHy6wTMTzAYGxM9wxbBNNwJegRHmAwNkQNomHIT/GNjqugYDdQvtCy25Sj4hDiOx+gAMOJpniu2PUznWL/RuLEPzaFNrxBNg5LdqGihcHsAObqphf+SYxn1+Biw7WPNtBiIvTaYSOlhITO3wF4gLhW4uVpgOBq1wMjPkyqXZZSS6CrJOQl9YJDwKViMjY1RHErnVgsm2ji3hvQRaOJFSzcySYoEOtiI2QU09JF3YCBoysx7XSEJRjB+Twj8I1rAA2AzMqa5I+yzdi5otyAa7uWl7OljqsUl/vNr5vZQh6qMGim7L/eeTGQsHmPk3lz3n77OjkXTjJWXLE3tyk8XW0GxD83q0Vp/oI3qvgUdZz4BtPWNkzW8Z+BxM1SFpKFCz6MhhB8NkRxnkHaw5GJld1kfAGmq6w1Eax/mwlNBzvAJ6ngF+KaUGHpbSLztfSmtcLOgeDrDe/Dxofm3yZO7NLDX2jKmPWEV3eDS/mrkJQNWbpZUI9+Y/oWPqtRHVDJ687LF3ueiz3+QiVJWLtrq5cE0WMbmmqyyYpu5y4QbuclFcR1RAJBSC/+BZ+Q32DfYXwab9az3ty8i0r2/TvvBO+4lg2o8q036GmvXD3U9icVfbPOzyygAAAABJRU5ErkJggg==',
  drive: 'iVBORw0KGgoAAAANSUhEUgAAAM0AAAAiCAMAAADLcK0MAAAASFBMVEUeLE7X3OEEEjlMV3Jpc4qVm6w5RWO0ucQAAAB6gpcFyZ8CxJgCzKC9wcsAf38AsJ0A/39W1rlz3cYHzaP9/v4Aw5gQyKAYJkovVBtsAAAAGHRSTlP//v///////wD/nSfp/wIPAvr8Q//+/v8FDva1AAADCElEQVR42u2Z0XLjIAxFJUBgJ3Hapnb4/z/dK4Ft3KnbZmf2Id4yU5wxltCRLpA61HWvXfd+mp67nd6Ng7qX7kVZ7s/clAccHb1258t0v/H4zI1v9+ly7l7pCpi38fnbG3CudD1djgADnMvpSufpPh6j3acznabbQWhu04mmiQ9Cw9NEhxGaSu3/oGEJQaBB8bJrvR3jlNKwUS2H8IWKMfr5MBzx3I+lrxH1/Hc0kp22NEYXdq39Zsw7ci5T34SlDnZhssuZd0aQJXGlp/pMwvNE8jc0QjlKkigfIv6KJhAF/DnfxMX7yRQX+Su/IWd46hd3hJuyw/8NTcwlJayeBWEWYeGD3uY+Evk6q/hajeCiyicjfxBRCllGSIm9mpjmOETkaIbBFCEUp748YSbaevWEKahJGBsI2YwwbSXwHQ2vgvGZcrbS9y4Hr4Vn77L3pKQ9GGYpGM3IyokORmJKM9uAFHN0MdAcndcnSJ16uBtmkyIqyswUescIXxaaVJxDfNBA/DlNyotAQTOw1whtLgWAu6Rl08SFVS+FZrRnIVSoDJdUtBLhz6SFW9wojW0i7xaTEjm5xE6Sk/VxLGCfTH3BEsIP1UZaDWtEYsnLzluiZ1JaV0lLU+xxSSaRQQsYqgNZaGx9si31dkqdVMQxQ+RO1QZD0bo4my859OERGsr+E5qgTTY0mXhLg6iCdsNMo6FZOkN1kPZohiYx3mOAolat0ozjkFilMabgAcY/3wUC1gh0IGmlMeHqNvVBadQojTlFXa3IxlCSMljYNWYTJXNDs1Ha0Gyplk7wrwWzw8ecl9WZHjhvvB4dqoqFBh0KHHGr2QXC6BccrGvSE0fmotQL4jQ9YhcgmLmWRtNWd4HcxIfMaRHENQULpCXRIyh7fHqkNnXnxNnrc28Oyq0cbb+dd2gd87nuA9ibZ0HXxVsvoWYYZuBatugi5mWH3sTnqeSB1jNJ1Lsap6jpSf/qexr/fuv8pfmleW6aQ/0nfay3HMd6A3Wst4PHenN7rLfqx/rF41C/Rv0Bd1HquzyXDMgAAAAASUVORK5CYII=',
  refresh: 'iVBORw0KGgoAAAANSUhEUgAAAL0AAAAiCAMAAAA9OSiUAAAASFBMVEUfLU/W3OCSmapOWXQFFDtocYg5RWO1ucUAAAB7g5gFyZ8CxJgCzKC9wswAf38AsJ0A/39W1rlz3cYHzaP9/v4Aw5gQyKAYJkpJp7RNAAAAGHRSTlP//v///////wD/nSfp/wIPAvr8Q//+/v8FDva1AAACyElEQVR42u2Z21LjMBBEZ3S1cwOCHP3/n273yE4AO0uZp9UWqqKQZWl03NOjECPD8DoM74epr3Z4N24ZXoYXst96auQH9yCvw/E03a5aemp6vU2n4/AqF8C/lf7aG/AvcjmceoQH/ulwkeN0K32223SUw3TtlP46HWSatFN6nSbp1ji0zv9In9m+OkpjjE9dplv3Rp+3gvtx1fvOJJs7P6HXKiGEmvTzYKj1Ob2ENakEtzHVBVn1vmk5yA56qSn6GvznTdNf6htL1neTbNJLWvW+077uoqdoXkivXsTjyqUqMRbnneO4wzCzqVGqpIz41UVJd1osSzkhDDq1jWf2vFKGs2eHvZysV87zNtgHsxgZtsIK5UB0m1l/Si+xaAoRMbBhrKDA1hUK+FCRmDKGGj3T6QOz5EjfJtyfX8AjvOGZxpH6sQcc4cyQqD2QhWEyEm2zsE+MYjdhVEDDNJ5773IO4lYq4PkIluHmHA+NVak0741F2ClKeq8qi1VcQ1ouRwRQ4XolPUKYmujllmhPlzq6zxJR67lgRlRtd6Lspcfi1qNQYgkvpQWb5WNhROTCO2W6ULVpKdNo08yAsBxiCClRTW5+MjV6q1oKZDMxls1flSpgD11inPfRI6CjX0DvY7SjcqGPjd5GyeITnwK+2KYHocuRJDrCSzaSyhN6OeMnorJknvYzejEHykJbWond6U0lswGDOiprI3f6h3PmzIvFsOUf6NfOMcVdfdD/xDmsn0wFUE4oy1ZFD3pULUqQrMl+r7RvVVtN6Yp6tOez8nVNcTthWZhfqhZx0rJQZh32Vm2yWvTVl3bi8TPT8Qpj7eMRJyZcrGWcDzjQnHnXfTwx7RK6jw7HFI9eOzotkDaNPbRenZg+J+Sk7VfaiSk7zvvfv9J+6X/p/2n6rr8Z9v2tvO83In2/jer7TWDfb2H7fgPe9X8f/gBS+cHokq6toAAAAABJRU5ErkJggg==',
  today: 'iVBORw0KGgoAAAANSUhEUgAAALMAAAAiCAMAAAAj8BgnAAAASFBMVEUQZ2sAzaAAAAAKin0Ax5wgFUIAxJgAf38AsJ0A/38AzZ8JjoAhDz8AxJocJ0sBuZQZNlMUVmMeG0UFposWRlsA06IIl4MNd3Oz3nzbAAAAGHRSTlP/+AD/nf8nAg8CQ////v////////////8eMVyjAAAC0ElEQVR42u1Y23LjIAwVxk6atgiJ6///6UpgJ3ab6XjzsDvMlEmIjSVxJI5kAkzTbZreZzdCm98bWpjeprfZmCEwGzML2glu08VkN0rL5jLd4HoZJMhbqC9XuM7ZjdTyfIWLcWM1c4F5OMwzuPHaL+b/iNkXbfEVc1yeDZ+25Znji5gxhCVg+kk9PrPuaanP7AU2J7R1HJfyImbCBAkDa5o+Elaf3MtkCRh7ZffboPxSqLshs/Ye4SC4034o+44ZFfPx0UnMAbJLmJyzDBoV8K4wlERb7EtCZBaANSE1QsREVFKo2TMhpup8I0rr1XvVoOQP2lUkSSMeRYfoIyvmLItgD+JnMbtIi8wEn+SNXSQqdUFpIXWRqjfIGQJyChIc0SEFW2XexIwLZP5UqBJR9xE1gswJ40E7ETMFQcZBdaBh9jrxQfwkN1A+LF6CWixIghllhQE3M6WtYlSQLoWUaxAZGax9TSGQPoziFK/y4vx9vZu2b1Rrl4Q2K7nUXpKA5aP4ScyIjYQrZo1zIK+gHpjlyrZeAYKugG98FrpIIy++VGF42dZOaGDdXlsI0GaymRchQm2Y+7xH8bPcqKjzK5Rsw4rZncCsJKoFBLPoJavYu1EQgMHutSMGKFYGs6vCq8AdM38T/4sclGmr0MInfI45P+NG5sA5s2KV9VoLiVaBnIW2vNPO6pARtz9cdO2uc0Pz6Ch+Ns4SQwlawd4emMtWSSXV7Lcc7HFOqXHD1XD3sWfm5kHTLlFLA2o0qSejRr70CnAQP8dnFZZi56UcEdieg2s2rkLyIFTztdZhNR4IoTbMYojvb7heAffa2RIlS+IxELWiF+UmF9H/In5iv+F7wrZeu/71Xwr9er1/p/iDiri9Z+SxCqxT3C92xr5P9c/2SJJGgfxY+zrdAkQ32F7Ue/e7fx4Y84j/YUc8KxjxTGbEs68RzxhHPMsd8Mz8D5QgPIeehgwqAAAAAElFTkSuQmCC'
};
