/******************************************************************************
 * EP 1099 — EMAIL TRACKER (Google Sheets + Apps Script)  v3.1 (EP colours), Nov 2026
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
const W9_LIST = ['No', 'Yes', 'Yes (email) — check file'];
const IRS_W9 = 'https://www.irs.gov/pub/irs-pdf/fw9.pdf';
const C = {
  navy: '#1c2a4d', teal: '#00c49a', mint: '#e6f9f4', white: '#ffffff', ink: '#1c2a4d', muted: '#55627d', line: '#d5e6e1', band: '#ffffff',
  red: '#fde2e6', redInk: '#b4233b', yellow: '#fff0cc', yellowInk: '#8a5a00', blue: '#e1eaff', blueInk: '#2357c6',
  green: '#dcf4e8', greenInk: '#11845b', orange: '#ffe3d1', orangeInk: '#b3470e', grey: '#eceff3', greyInk: '#5f6b80'
};
const FONT = 'Poppins';

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
    .addItem('Sync sent / reply status (mine)', 'syncMine')
    .addItem('Refresh dashboard', 'refreshSummary')
    .addItem('Re-format all company tabs', 'reformatAll')
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
  ensure_(SYS.TEMPLATES, ['Template', 'Subject', 'Body'], templates_());
  ensure_(SYS.INBOX, ['Company', 'Vendor', 'Contact Name', 'Contact Email', 'Reportable', 'Form / Box', '1099 Status', 'Sender Name', 'Sender Email', 'Tax Year'], []);
  ensure_(SYS.LOG, ['Timestamp', 'User', 'Action', 'Details'], []);
  OLD_SYS.forEach(function (n) { const s = ss.getSheetByName(n); if (s && ss.getSheets().length > 1) ss.deleteSheet(s); });
  refreshSummary_(ss);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
  ss.setActiveSheet(ss.getSheetByName(SYS.SUMMARY));
  ss.moveActiveSheet(1);
  log_('Setup', 'Sheet set up (v3)');
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
      return { vendor: r[0], cn: r[1], ce: r[2], amt: r[3], fb: r[4], st: r[5] };
    }));
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

function upsert_(ss, company, taxYear, sender, rows) {
  const sh = clientTab_(ss, company, taxYear);
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
  const y = n.match(/Tax year:\s*(\d{4})/), c = n.match(/Company:\s*(.+)/);
  return { year: y ? y[1] : '', company: c ? c[1].trim() : '' };
}
function setMeta_(sh, m) { sh.getRange('A1').setNote('Tax year: ' + m.year + '\nCompany: ' + m.company); }

function headerRow_(sh) {
  if (sh.getLastRow() >= HR && sh.getRange(HR, 1).getValue() === 'Vendor' && sh.getRange(HR, 2).getValue() === 'Reportable') return HR;
  if (sh.getLastRow() >= 1 && sh.getRange(1, 1).getValue() === 'Vendor' && sh.getRange(1, 2).getValue() === 'Reportable') return 1; // v2 layout
  return 0;
}
function isClientTab_(sh) {
  const n = sh.getName();
  if (Object.keys(SYS).map(function (k) { return SYS[k]; }).concat(OLD_SYS).indexOf(n) >= 0) return false;
  return headerRow_(sh) > 0;
}
function clientTabs_(ss) { return ss.getSheets().filter(isClientTab_); }
function tabYear_(sh) { return String(meta_(sh).year || setting_('Default Tax Year') || ''); }
function tabCompany_(sh) { return meta_(sh).company || sh.getName(); }

function read_(sh) {
  const hr = headerRow_(sh) || HR;
  if (sh.getLastRow() <= hr) return { hr: hr, rows: [] };
  const v = sh.getRange(hr, 1, sh.getLastRow() - hr + 1, CT.length).getValues();
  const head = v[0].map(String);
  const rows = v.slice(1).filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    const o = {}; head.forEach(function (h, i) { if (h) o[h] = r[i]; }); return o;
  });
  return { hr: hr, rows: rows };
}

/** Rewrites a company tab in the formatted tracker layout. */
function write_(sh, rows) {
  const meta = meta_(sh);
  const days = Number(setting_('Follow-up After (days)')) || 7;
  if (sh.getMaxColumns() < CT.length) sh.insertColumnsAfter(sh.getMaxColumns(), CT.length - sh.getMaxColumns());
  const need = HR + rows.length + 5;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart(); all.clearContent(); all.clearFormat(); all.clearDataValidations();
  sh.clearConditionalFormatRules();
  all.setFontFamily(FONT).setFontSize(10).setFontColor(C.ink).setVerticalAlignment('middle');
  setMeta_(sh, meta);

  // Row 1-2: title band
  sh.getRange(1, 1, 2, CT.length).setBackground(C.white);
  sh.getRange(2, 1, 1, CT.length).setBorder(null, null, true, null, null, null, C.teal, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange(1, 1, 1, 9).merge().setValue((meta.company || sh.getName()) + '  —  1099 W-9 Tracker')
    .setFontSize(18).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange(2, 1, 1, 9).merge().setValue('Tax year ' + meta.year + '   ·   W-9s due ' + fmt_(setting_('W-9 Due Date')) +
    '   ·   1099s due to vendors and IRS by January 31, ' + (Number(meta.year) + 1)).setFontSize(10).setFontColor(C.muted);
  sh.setRowHeight(1, 40); sh.setRowHeight(2, 24);
  // Row 3: buttons (images); row 4-5: live counts
  sh.setRowHeight(3, 46);
  sh.getRange(3, 10, 1, 6).merge().setValue('Buttons: select vendor rows first for “Follow up selected”.').setFontSize(9).setFontColor(C.muted).setFontStyle('italic');
  const kpi = [
    ['Vendors', '=COUNTA(A' + (HR + 1) + ':A)', C.band, C.ink],
    ['To email', '=COUNTIF(I' + (HR + 1) + ':I,"' + RQ.NOT + '")', C.red, C.redInk],
    ['Draft ready', '=COUNTIF(I' + (HR + 1) + ':I,"*draft*")', C.yellow, C.yellowInk],
    ['Emailed – waiting', '=COUNTIF(I' + (HR + 1) + ':I,"' + RQ.SENT + '")+COUNTIF(I' + (HR + 1) + ':I,"Follow-up ? sent")+COUNTIF(I' + (HR + 1) + ':I,"' + RQ.FINAL + '")', C.blue, C.blueInk],
    ['Follow-up due', '=COUNTIFS(K' + (HR + 1) + ':K,"<="&(TODAY()-' + days + '),M' + (HR + 1) + ':M,"No",I' + (HR + 1) + ':I,"<>' + RQ.RECEIVED + '",I' + (HR + 1) + ':I,"<>' + RQ.DONE + '")', C.orange, C.orangeInk],
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
    cf(due, '=AND($A' + r0 + '<>"",$K' + r0 + '<>"",$K' + r0 + '<=TODAY()-' + days + ',$M' + r0 + '="No")', C.orange, C.orangeInk, 'formula'),
    cf(rq, RQ.NOT, C.red, C.redInk), cf(rq, 'draft', C.yellow, C.yellowInk, 'contains'), cf(rq, RQ.RECEIVED, C.green, C.greenInk), cf(rq, RQ.DONE, C.green, C.greenInk),
    cf(rq, 'sent', C.blue, C.blueInk, 'contains'),
    cf(w9, 'Yes', C.green, C.greenInk, 'contains'), cf(w9, 'No', C.grey, C.greyInk),
    cf(st, 'Eligible', C.green, C.greenInk), cf(st, 'Needs W-9', C.yellow, C.yellowInk), cf(st, 'Needs review', C.orange, C.orangeInk)
  ]);
  // Layout
  const widths = [230, 100, 110, 110, 140, 210, 120, 200, 160, 70, 110, 120, 140, 60, 100, 220];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(COL['Thread ID']);
  sh.setFrozenRows(HR); sh.setFrozenColumns(1);
  sh.setTabColor(C.teal);
  if (sh.getFilter()) sh.getFilter().remove();
  if (rows.length) sh.getRange(HR, 1, rows.length + 1, CT.length).createFilter();
  addButtons_(sh, [['drafts', 'draftsThisTab', 1, 6], ['followup', 'followUpSelected', 2, 0], ['allfu', 'followUpAllDue', 4, 0], ['sync', 'syncMine', 6, 0]]);
}

/** Adds clickable image buttons once (they stay put when the tab is rewritten). Click runs as the clicker. */
function addButtons_(sh, list) {
  const ver = 'btn-v2';
  const imgs = sh.getImages();
  if (imgs.length && sh.getRange('A3').getNote() === ver) return;
  imgs.forEach(function (im) { im.remove(); });
  sh.getRange('A3').setNote(ver);
  list.forEach(function (b) {
    const img = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(BTN[b[0]]), 'image/png', b[0] + '.png'), b[2], 3, b[3], 6);
    img.assignScript(b[1]);
  });
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
  const name = row >= 6 ? String(sh.getRange(row, 1).getValue()).trim() : '';
  const tab = name && ss.getSheetByName(name);
  if (!tab) { SpreadsheetApp.getUi().alert('Click a company name on the Dashboard first, then press the button.'); return; }
  createDrafts_([tab]);
}

function createDrafts_(tabs) {
  const ui = SpreadsheetApp.getUi();
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
    ui.alert('No W-9 request drafts to create for ' + me + '.\n\nA row needs: Request Status "Not sent", Sender Email = your email, and a Contact Email.' +
      (noContact ? '\n• ' + noContact + ' of your rows have no Contact Email.' : '') + (others ? '\n• ' + others + ' rows belong to other senders (or have no Sender Email).' : ''));
    return;
  }
  if (ui.alert('Create ' + todo.length + ' W-9 request drafts in your Gmail (' + me + ')?\n\nNothing is sent — you review and send them from Gmail → Drafts.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const tpl = templatesMap_()['Request'];
  let w9 = null;
  if (/^y/i.test(String(setting_('Attach W-9 PDF')))) { try { w9 = UrlFetchApp.fetch(IRS_W9).getBlob().setName('Form W-9.pdf'); } catch (e) { log_('W-9 download failed', e.message); } }
  todo.forEach(function (x) {
    const r = x.d.t.rows[x.i];
    const data = tplData_(r, x.d.sh);
    const opts = { name: r['Sender Name'] || undefined };
    if (w9) opts.attachments = [w9];
    const dr = GmailApp.createDraft(String(r['Contact Email']).trim(), fill_(tpl.subject, data), fill_(tpl.body, data), opts);
    r['Thread ID'] = dr.getMessage().getThread().getId();
    r['Request Status'] = RQ.DRAFT;
  });
  data.forEach(function (d) { write_(d.sh, d.t.rows); });
  refreshSummary_(SpreadsheetApp.getActive());
  log_('Request drafts', todo.length + ' for ' + me);
  ui.alert(todo.length + ' drafts created ✓\n\nOpen Gmail → Drafts, review each one and click Send.\nThen click “Sync status” here so the tracker turns blue.');
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
  if (!total) { ui.alert('No follow-ups are due for ' + me + ' (emailed ' + days + '+ days ago, no reply, no W-9).'); return; }
  if (ui.alert('Create ' + total + ' follow-up drafts (replies in the same email threads)?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  let made = 0;
  plans.forEach(function (pl) { made += followUp_(pl.sh, pl.t, pl.idx).filter(function (x) { return /created/.test(x); }).length; write_(pl.sh, pl.t.rows); });
  refreshSummary_(SpreadsheetApp.getActive());
  ui.alert(made + ' follow-up drafts created ✓ Review them in Gmail → Drafts and click Send.');
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

/* ============================== DASHBOARD ================================= */

function refreshSummary() { const ss = SpreadsheetApp.getActive(); refreshSummary_(ss); ss.setActiveSheet(ss.getSheetByName(SYS.SUMMARY)); }

function refreshSummary_(ss) {
  let sh = ss.getSheetByName(SYS.SUMMARY);
  if (!sh) { sh = ss.insertSheet(SYS.SUMMARY, 0); }
  const head = ['Company', 'Tax Year', 'Sender', 'Vendors', 'To Email', 'Draft Ready', 'Emailed – Waiting', 'Follow-up Due', 'Vendor Replied', 'W-9 Received', '% Complete', 'Open'];
  const tabs = clientTabs_(ss);
  const all = sh.getRange(1, 1, sh.getMaxRows(), Math.max(sh.getMaxColumns(), head.length));
  all.breakApart(); all.clearContent(); all.clearFormat(); sh.clearConditionalFormatRules();
  all.setFontFamily(FONT).setFontSize(10).setFontColor(C.ink).setVerticalAlignment('middle');
  sh.getRange(1, 1, 2, head.length).setBackground(C.white);
  sh.getRange(2, 1, 1, head.length).setBorder(null, null, true, null, null, null, C.teal, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.getRange(1, 1, 1, 8).merge().setValue('Enduring Planet  ·  1099 W-9 Collection Dashboard').setFontSize(18).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange(2, 1, 1, 8).merge().setValue(tabs.length + ' companies   ·   counts update live as the team changes statuses   ·   1099s due January 31')
    .setFontColor(C.muted);
  sh.setRowHeight(1, 40); sh.setRowHeight(2, 24); sh.setRowHeight(3, 46);
  sh.getRange(3, 6, 1, 6).merge().setValue('Click a company name below, then “Drafts for selected company”.').setFontSize(9).setFontStyle('italic').setFontColor(C.muted);
  sh.getRange(5, 1, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor(C.navy).setBackground(C.teal).setWrap(true).setHorizontalAlignment('center');
  sh.getRange(5, 1).setHorizontalAlignment('left');
  sh.setRowHeight(5, 32);
  const rows = tabs.map(function (t) {
    const q = "'" + t.getName().replace(/'/g, "''") + "'!";
    return [t.getName(), tabYear_(t), '=IFERROR(' + q + 'H' + (HR + 1) + ',"")', '=' + q + 'A5', '=' + q + 'B5', '=' + q + 'C5', '=' + q + 'D5', '=' + q + 'E5', '=' + q + 'F5', '=' + q + 'G5', '=' + q + 'H5',
      '=HYPERLINK("#gid=' + t.getSheetId() + '","Open →")'];
  });
  if (rows.length) {
    const r = sh.getRange(6, 1, rows.length, head.length);
    r.setValues(rows.map(function (x) { return x.map(function (v) { return /^=/.test(String(v)) ? '' : v; }); }));
    rows.forEach(function (x, i) { x.forEach(function (v, j) { if (/^=/.test(String(v))) sh.getRange(6 + i, j + 1).setFormula(v); }); });
    r.setBorder(null, null, true, null, null, true, C.line, SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment('center');
    sh.getRange(6, 1, rows.length, 1).setFontWeight('bold').setHorizontalAlignment('left');
    sh.getRange(6, 4, rows.length, 7).setFontSize(12).setFontWeight('bold');
    sh.getRange(6, 11, rows.length, 1).setNumberFormat('0%').setFontWeight('bold');
    sh.setRowHeights(6, rows.length, 30);
    const col = function (c) { return sh.getRange(6, c, rows.length, 1); };
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(C.red).setFontColor(C.redInk).setRanges([col(5)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(C.yellow).setFontColor(C.yellowInk).setRanges([col(6)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(C.blue).setFontColor(C.blueInk).setRanges([col(7)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(C.orange).setFontColor(C.orangeInk).setRanges([col(8)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(C.green).setFontColor(C.greenInk).setRanges([col(10)]).build(),
      SpreadsheetApp.newConditionalFormatRule().setGradientMinpointWithValue('#fde2e6', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue('#fff0cc', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
        .setGradientMaxpointWithValue('#dcf4e8', SpreadsheetApp.InterpolationType.NUMBER, '1').setRanges([col(11)]).build()
    ]);
    // Totals row
    const tr = 6 + rows.length;
    sh.getRange(tr, 1).setValue('ALL COMPANIES').setFontWeight('bold');
    for (let c = 4; c <= 10; c++) { const L = String.fromCharCode(64 + c); sh.getRange(tr, c).setFormula('=SUM(' + L + '6:' + L + (tr - 1) + ')'); }
    sh.getRange(tr, 11).setFormula('=IF(D' + tr + '=0,0,J' + tr + '/D' + tr + ')').setNumberFormat('0%');
    sh.getRange(tr, 1, 1, head.length).setBackground(C.mint).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    sh.getRange(tr, 1).setHorizontalAlignment('left');
  } else {
    sh.getRange(6, 1).setValue('No companies yet — in the EP 1099 Tool, open a company and click “Send to email tracker”.').setFontColor(C.muted);
  }
  [240, 80, 220, 80, 90, 90, 110, 100, 100, 100, 100, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(5);
  sh.setTabColor(C.teal);
  addButtons_(sh, [['company', 'draftsSelectedCompany', 1, 6], ['allfu', 'followUpAllDue', 3, 30], ['sync', 'syncMine', 5, 0]]);
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
    'Contact Name': String(r['Contact Name'] || '').trim() || 'there', 'Vendor': r['Vendor'], 'Company': tabCompany_(sh),
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
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground(C.teal).setFontColor(C.navy).setFontFamily(FONT);
  sh.setFrozenRows(1); return sh;
}
function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg, 'EP 1099', 5); } catch (e) { } }
function log_(action, details, ss) {
  try { const sh = (ss || SpreadsheetApp.getActive()).getSheetByName(SYS.LOG); if (sh) sh.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]); } catch (e) { }
}

/* ======================= BUTTON IMAGES (do not edit) ======================= */
const BTN = {
  drafts: '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',
  company: 'iVBORw0KGgoAAAANSUhEUgAAAQgAAAAiCAYAAACjtkcbAAAZPklEQVR42u2deZDdV3XnP/f+tre/7lYvkqytJdvaLMs2Nra8YhwTFgeDbSgSIJBMMTUzzDAFk4GaMBNITdVMQibDwNRMTVIJkGCWGBMTJ4YM4C2WRzK2LFuLtdqtvdV79+u3/bZ754/7e/1et14Lm5IAm3equqT6vfe7Z7n3nnvO95zf+wkapD8nEX+oANj11TuI9X1o3o1Sy4hiEELQoQ516I1BWmtsC6QcRvAwlniQN/3Oowt9gdn0Dzxg8f73x3z/SyvoL/4X4MMIAUEIkWp8q0Md6tAbykkAtgTXAa0Bvs7ozO/zzn97quETxJy3ePLLg6Szj5FJr2Fq1kQSCIHouIcOdegN7CR0wzvQnZdUa8eoVd7KbZ8YQn9OCj73OcnbV/RhWTuw7EEqtRApnI7lOtShXzFSOiSbdoijIeJ4G/94akwCEKkvkkl3nEOHOvSrTFI4VGohmfQgkfqiwSCe/otb8ewn8YMIjd2xUoc69CtOggjPtfGj2yTwAYTQaN3BGjrUoQ6B1gIhNPABiVbvIwgFGtmxTIc61CE00vgE9T4bS/aaUuYF7HNQepHQRVy8kulCnheL18+Lz4U7DUw5y+SYr4fT6/Ul7xsyxRCCSIEte22iWF/wJqi0A7ZoTrRONlYYJ30VF3hTaSBtm5quTro76hGEihYA5iLwEVALF3eIvwx1btcyfwoj6y/zZtaAZ4NjQazMHHboF4VDQBRr+4I6Bw1YAm//KNZkDW1LkAJtS1TWIRrIoZZkwI8g1os7Cc25n+kWwc/ZBBJvzwiZn5xBVgKIFaV7NhJeusQsYqUXX2yLjcur4BMpSvduIlxVhCAyDuPVjreYnovJ99PGXDiW1pCySe88RfbJ48R9Gabfvxnt2sYe4mfQv/V7r2Uz61chu9aQtsn+6BUyu84QrO5i5p6Nr37dnU+H1zonC7+rL8B8vtp55GfgexEjCfuCnyiOTeF7B8nuOIXK2IhqCFKicg5xd4bq9Zcwc+9GtJssLCHmn8ACsGRz0ejkdJItPGJt7kODa+MdGKP/j55G+JEZF83suy4n89gQ2e3HCdZ0MfOejWbsRvSiNFiiefJpbU5ZsYhe7fgIKL/9UjNGrE3UZCUDxKope7v0ZKGe7XiSfKf1fr1gk2ptvmMlnzdsk3TJuSdK5B4dIhwsMn3fpua9ShvZWuVtpEwLN26rneKkr8azyP442cyriufatzF37WRbKL8tITa28o5MkPvxK1SvWf7THcRCHRbap+EIGzZUunnPYiljIzpsrE0S+zbWXet60cyPUO1kHbTyaMhjtfBs2PB8cjRsHieLUpwnYmvoLRfM3wWgi1LW1GnHRAxL0vgb+yBSuK9M4YxUKDx0AOfMLOOfvAEtEiNknHmGE7M+WNJsQkuCHyEis+G0a0HOhVpk0pWCh/vyJLIWEqwqMvKfbkV1paHokfnyTnKPDlG9djkzH9pqePnJfVkXKgEijEGCti1w5fy0ZG4CaM9nScbIEURQ8KAWImoRCNA510xWPTT/nk9PR85fbFqbk9mSiLJv+FvC6C6FkV9rsE36IMommtGONHr5EYR6zl4q76Ky7nydEhuKWmT0z7mGTxAZHioZ37OgHCDCyMiask2qWEjhnJgm96NXzrVvEJu0xllEtjhR1hZgW4iSj7bM5zplo/IeOvMq2nFyLtSjps0922ysOLFPxoFYIyqBsUUqWTvVllQr45h7IuNMRMkHIdA5x9jZthAzPjiWkcmPWlI3u+lYI4WcDVBZx3zmJxGlLc249WbUrDOOuZ7I3ZRDkjz3BNUQEUTogmd4RKp9xKYbKX1ybz2+oFGHfdGApjBG5TwmPv5msCVyuk73V18g/exp0s+dJvvYEOX3bMB+eZLiQwchVlRvWkX62dNkdp5m5n2bKN+5liX/fQfOiRmsagBCEvVmqF6zlNnfWI9O2xTuf5Hsk8dRWRdZj+j+xl5UykanHdL7xgiX57FmA5Z8cQeyFlF612X4m/sp3r+HzHOnsSbraEei0g61Ny1j+re2zPfmWoNrteWjPZup39qCdh3yDx0ks/MkztkyyrUJB7soveMy/C39ECus0Qpd3z0wX89nTjPz3g3M3rMJKoHZYBpwLOzhMl0P7Mc9OoGshKiMg8q6zNyzgdq2lRBrrKkahe8dJLVvFGumTtyVpvrm5ZTu3gC2NEFII8Vq6CTM6ZN7+BDZ7SdwzpbRnkX98iWU3ruRcE2XWbgpG2uqRv7hQ6ReGsOeqKItSbCum/F/dR2Fr+4mvXeU8JIW+9YjSm+/FP/a5VinShQePtRWNhPlgayGdH1zL6m9o4ggpvTu9U286ny4TkOHR46Q2XESZ7iMiBVxwWPyd6/Gv6IfNKSeHyb/j0dxjk8jQ0W4NEf1xpXMvm3d3Glc/OZe3GPT1K5djqiF5H9wFG1LZu+6nOotq+j6s12k9o2iUzal37icym1rQArSO0+Re/IY4SUF/MuXUHxgP9asT7gsz8wHrsBf3wtovD2jdH9rL9aMj6iF6LRNuDRP+c61VLetNM5UQvHb+3BfmaJ63XKE0uQfOYIIFfXNfUx/+EpUwSP/94dJ7RvF39BrbBXEkLLJPXKY9J4R/PW9lO5ebw458cvsIFo2lygHaFeiCi7TH9yCd2AUGUSkdw9Tfvd6ZCUk+0/HQQjSu4aTSEEjqyGiFpLaP0awukiwfgmyEuIdGKfn/j1IP2b6n78J7+A4zulZVNZB1CPSz55G5VxU1sMeraAdiawEpHefRZZ8yrcPknnqON33v0i4skh96wAojXt4gtS+0fYL0xK4hyeafPyI9E9OozMO0x/cQvH+Fyl+9wAq71G7agB7sk5m52lSe0YY+70bqd+wAqs0da6egFUO5k9m8v/uv3qBzI6T+Jv78Tf0IWd9UntH8Iamqb1lEGt4lr4vPI13aJxgbQ+1a5fjHRyn65v7cIbLjP/rNydhacvQCrRj0f2V3RQeOkDck6G2dQB7qkb2qRN4hycY/eytRCsLyKm6Gf/gOHHBM44u0rjHZ5D1CO/gOPbZMjrtNO0761PZthJRDsy9Rybay/aJ68GSdH9lN7nHjhF3pfA39ZJ94hjWdB2VcRCLOQilIe3Q9VcvUHzoINqR+JcvIRrIktozgnNmFn/bStI/fpneLz2D8GP8jb2EBY/UC2eNw5qoMf3bW6EWkn5+GO/QON6BcUSs0JZE+hHFb+8j9+NXcI5PozMOcrJGz1d2E6zuIryiH/fYNNl/Ok7clSbz9AnCwW5EEBsZzpYZ+YPbiNb14Jwq4ZwqUb9qKcqzsYdnSe0fxTswRpx38a9cCvWQ1PPDpPaP4R2aMJGDbSFiRf7/HiUuesz8y+tQadvM06FxqtevIFqWQ874FL5/BPfIJPWNvSZ6CeILlmrYP4f2TZAS6hFxziXuz2KPV7FKPtQiA2DmXUSoiAZyTHz8OhMORwrt2Zz58jtQRc+EZ5ag669fpPD3h0i9OIIYrzL+mZspfGsfXQ/ux1/Tzejv34x2bMg6dH31eQqPHKF2RT9j//4mUBrdk6b7z3ehHYvalgEmP36dCe2UwjpbaWITuqWMWY+Y+Hc3Ev7N/iaf/3AzuujhHC+Re2wIlXGYuW8TpQ9diZio0vdft5PaO0L+Hw5Rv3Y52haovIcI43l6ijA24WIjerAEohpij1ZQOZfSOy+jeuc6E1aWA6zpOmhN5sljZgMOdjP2qRuI1vbg7R6m90s7ST97Bm//GP4tq5pYhgaVsnCGpsg+OoTKeUx9aAuVX78UOVWj94s7Se8eJv/Dl5n65A3kHjyAe2SCaEmayX9xHbWbV0EYY41UUJ7F2GdvpesvdlH4/rn2zf/tAbyjk21ly/zkNO7RSVTeJbNrGJWxKd29ntJHrkKeKTHw+SfM2lgMC0rZ2K9Mknt0CJWyKb9tHVP/7Gpjt8kaMlAwUzcncD2ids0yxj5zMxQ8il/bTfFvD5B7fIjyW9YQrSig0zYq56LTNqOf2oaWgv4/3o6sRTBTZ/i/vQ1r2qfvf+xEzvp4hyYIr1qKdi3iQgptC6Y+ejW1t6/D2TVM/xe2Y41XyTx9ktKaLqrbVlC+Y9CkB2jkjE/fnzxN6qVx0s8N41+zzKiWcVAFD20Lxv/NjYTruun90x14B8bwjk7CVI3qzauoPTpE6uA46WdOMfvbW/F2nsIerxJc1kPl5tXgxxcUh7B/LqU23QIeJeGPtmRzU2gQtYjZd1xKuKkPZupgS0QQk318iOz/O4k9VjFOIkkHZDVA+BGqL9PM4YXBP7RjQdo2uEJSjtRpJ2Ek8Df0wg+OkHvyGOndw0T9WcLVXcz+2lri/mzigReokbLNuHN8bHTewz5TQtYjVMahtqUfSnV01iG4fAmpfaM4IxWYDeaArnl6lvwmptCIHmKFzrkEl/bgvTzFkv/zHF1/s49oaQ5/Yx+zv7bWYLND02jXxpoNWPoHjyMijbYlsh4ip328lyfx37KmpdRswEDnZAnpR6i8S/c39tL9tRfBFohIQ6hwj0xALcI5OYNQEC0vULt6qZFVa+KeVJIPWy32aLVvIpu3iGxTddxj04RLc4haiMq61K/oh+kaqiuFv74X9+Upg08t8niyezKxedqmcvNKo9tMgPYc4pxETlSxx6sgBf7GXkhZUAmobRkg/8gRZDXEOVMiGuwyWGElpHrDCoKtAzBZJ+5KYU9MUb59DdGGXuKRCnHORU7XkPVwThYRxER9GeqbemGkQriqSLgsjz1SwTldMlNaj+j59j5SB8aRM/UEzDS4kzWb4EuNs7QSUH3zcurXrzCyr+0mtWcEESpEPUZ3pahuW0HqpTEyz5xm9p6NpPaOIMsB9dsHUf3ZZqr6unEQVgLUZB3c54dxhmdBCKL+LKQdhFJz34uLngGdghidssk/dIiev9xFuCxP6a7LidZ0k3ruDNnHXoG8m6DNC6oAsQapm0jxXE1XIdDoSkD1ltWMpmwyT5/APlvGGSnjHZogs/MUZ//z7USXFJply4XIc4OiBo/XULfS8/UkVGYyxYIEO1ZMfuwagtVFUvvGsMcqeIcnSf/kDO7hCcY+/5ak+qNQWQd//RIIYlOc8cyUhisKBkw8j2j1TX1o10JECu1YaEsS9WcW+Q0QPb9c186+Ss1VdxaVTWmiFQUD2v28foPoVbSpGPBVIxpzokGlHSN7ELevNJ23PGkqGr3/+1nSu89Su2qA0ns3EGcdCn93iNREte19KpOAo7TYtlEN8yOqN6yg8MgR3OPTpJ86gXd0CpV2qN5wyWuT8RfuIJJSm6wEUBW4z52h64H9iFChpaB82+pEoWSDiBbDJve6L0+gUg7B2m5KH7nKgENPn0C0qymLNn9JiiNiE5rqBF22zpSob+6jdstqEGAfHKf/j7Zjj1Vxj08TDXaBT1tsYG5saU77cEUBlbIQtciAdluXwkQV9/A4aE04kDXO7LRKylCterbxLdJEWdZEjdm7NzD7vs0GkHvwJbr/+kXc49NQDgjWdZN74hjaEkz95hbUuh7jEMoB3gtnjZNrlOUa5ceoIa+NqEdUr7+E6t0bzH2hwnlpzGyQlE2wqoiWYJ8pkdp9lnqSYsjRCtqz0Dlxrn0dExUGq4vknlRtZUvtPkuwpgtrvIpKGzwntX+U4E3LkWdKeIcn0K6NaLfYE0cUrDQ6SD8m+/RJgvW9UPQQU3XEbIzqSRP1ZrDPlvEOjJuIsOCR3jtioqesS7i8kESz4ty5pc16PGeuTERmT9VIvTRuUoznh3GGy2AJwlUFxEwd+2yZOO9SuWkV5Y9sRe4ZxSrvNT1C7fSb+1fMW8taAGFMvCxP9bpLyP/gCF0P7scer+Gv68bf0GcqJ/J1UOZslLHs8QpLP/soRBprpo7QGuXZzPzmFvytBpzRMikxRefWlcNVXcinTpA6OE7/Z36EqAQ4IxWUazXD8gYY2jjRW8XozUKkcE6V6Pvj7aiUzfRHryL31AkKDx3Av2KAOO9iJSGpStuElxQgUIvmwHN8EmwiXNdN5bbVFB4+TPG7L+EencCaqpM6OI7KOMzetd6UKxt19IV6tjt9pKD3y88gSz7BYDfas3CHppCVEH9TH9iS6q2ryW43tln26R/hb+hFRAp3aBpneJYzf/I24tVdxk6ROe1lPSZc203l9kEK3ztI7/96lvrjx1BZF3u0QvqFYSY/ehXB1gGqt60h+8xpvIPj9P7PZ/CfGEIrjXuixOh/vIWoL0vcm5lv37TDzPs3U7lzHdkdpxaX7U9/nWBzH7Wrl5J7/BiFvzuEd3gSa6yCNVlLUi3d/tCpR0SDPZTvGKT40EHy3z+Ce3SSaCBHau8IM/dtonzvJmbfdRne0UnSu88y8PknUAlIKfyY8l2DRKuKprO0MaetoKhKri2MTFuvJamz8my6v7ab3JPHcE6XsEYrRP1ZKjeuRKcd4t4M9kSV4sOHSO8dwTlRQtZCk1rE6lyeCypo89e1cfKVW1eRe+IY1lTd4Cw3rDBl2dng9eEgdNrUslXGMUbI2PjLcoTL8lRuXY2/dcDUolM2SIHKuUkqYiIHLQXUImbvugw5Uye9ZwTn+AzVm1dRvWk1xe/uN/ckXl67FnHebdbOpVlIlVtX4Q5N4R2ewDs0AbFi5t5NBGu7CVcWsM+UcKsh2rWob+432MDqrnPTCxbhIwSEiqkPbyXqy5HZedI4Bteiev0KSu9Mypy1EG1bBnyVYn4b+iJlvPrmPtK7h/GOTiKCCFXwKL9tLTP3bgINccFj7NM3UXj4EN5LY6T2j6EdSTSQo3zHINHSLIRx0lfgojIOWppTaOpDWwiX5sjuOGmwhlgT9aQp3b2B6g0roRIQF1OMffqmuTKnd3AcLU2ZU2VNL0HlltW4Q9NN+2rN7DsvJVpVZOz3bqTwD4fbyzaQhXrE1O9ejU7ZpPaN4h6doHTXeqyST/4HR1A5Z3HQO4yY/uAWooHcXJnTOVsmznuEy/Mw61N78wrGPm3PlTmd4TLB2u5mmTPp99AZh7hg+i/msqiMY+bZtecwJ5V1UHnXRKFJZCHCmHhJmpn3bKD43YOIekT9ygFmPnAFcX8OtGbyY9fQdf9e7OFZnBMzlO7ZiLd/jMyOk2Yd6Wbv0JwcjSXn2QvWtWmXD9b1ULt6KZmfnCbuSlG9bnlb3OzCdFxv//MLnrgIP0bECRiJaRLSrjTNII1nAmSz+00mzSfKs5ubJ8nXcSxEqQ5SooselAODMEuBStkgQQQKEcZoKZoG1kkjjiVN80uy8bSTPJugNcKP5zx22+aVBZt2UT4iaXSphmbM1kapWpg07yyi52LkWBDGpsqhTDefzpnqjmnmajZKmYYvkzvrtG2u1aLkBFImh27Yq7VBqB4ZLICkAS3jzKUbrY1Ycw1ljUapRrdgO/u6lvmPLReXrR6Zg8OWBowu+QZQzruIGd9gUJaYt1naUkOH1kYp2dIolW7TKJV25jVKiXqEiDXakWjPbH5RjxAquZY8xyLrxp7akuieFIXvvETX118kXFVk+At3gmMhp33j2GzT3AfCNJrFClkOzIGZchDTNQPctvL0o+a1pE9EBLHBhxrrLWn3J+9R/MvddH1jD+W3DjLxqW1Gp4vwcNtFiyD0wl5/raESntsm2oggWltG5xB9DVGYIORAOQBLoApus6EmOdl1ympem8tXNUSR2fyt7behKQU1JsKgeippWhGLhv6L8mnI1tgE0Oy4a1RqFtNzMWqccF6LI6qGTTwhCTcJY9NladtzHX2temjHmluE88LXhrxzjk6ba6Il7543vtO0kziPfRvh8E+TTYjm/enkJC0H7W28GC20udJzHZGNbkQETR2VnrtnblozjlmrLfzmrqnELkkEMaezZl5qIqohOiuMc2g899PgkXRUqpxrPqsEptK2kGe7a56NTtNs/U7Z5H70CplnTuO+PIn2bMpvHbyo+K590R691u2739pSpNs/xCKaaPCcY9Ftvq81RG0ehGmAPbE+d0xY/Pr5OkSjRR64kQvGbPfE6mJ6ngfknbdJFtpPtPneQjsvJvM58r6G8X+afV+tbK33Nz47n435aTb/GXSM2+h1vmtaz0s3VcZpjhlpkwbM07Hlswb/hc+MLMaz1RbJcy3OiRlSL54lvCTP9Ps24185YKLFi/RovOCpP1Odd150qEOv7VHoRrp5Tup2sVlHyqQdtkQ3nm25iO/OsLEt0Xn3RYc69Nr6Kl5TKnQhWTdSxoVY3kV5Z4YlJLEaN4+4at2Z+Q516DU8kNimtP5z47vwcfEL/tYtCbEalwj5HVxHI1CdWe9Qh17jry6JXxDfizu+Mj5BfkcC3275FdsOdahDnd+kbPzK/bclPzy5nZr/LYp5G6XDjnU61CF+td+uVczb1Pxv8cOT22XyM1ufpFobIpt2Ok6iQx36FX/1XrU2hC0/CSD5PLDtYyPUyncQh8foyTtorcwfnbSjQx3iDf7y3sZ+78k7xOExauU72PaxET4P5s3eDzxgcdsnhhgv30I9+DppT5JJSWwpOi6iQx1645ZrsaUgk5KkPUk9+Drj5Vu47RNDPPCAhfhD1dJz+jnjLAB2ffUOYn0fmnej1LLkhzQ7nRId6tAbp0yrsS2QchjBw1jiQd70O48u9AX/H7lFi5Wlcp8GAAAAAElFTkSuQmCC',
  followup: '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',
  allfu: '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',
  sync: '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'
};
