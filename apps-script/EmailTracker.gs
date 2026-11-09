/******************************************************************************
 * EP 1099 — EMAIL TRACKER (Google Sheets + Apps Script)  v3.0, Nov 2026
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
  navy: '#1c2a4d', teal: '#00b38c', white: '#ffffff', ink: '#1c2a4d', muted: '#55627d', line: '#dbe3ea', band: '#f5f8fb',
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
  if (sh.getFilter()) sh.getFilter().remove();
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.breakApart(); all.clearContent(); all.clearFormat(); all.clearDataValidations();
  sh.clearConditionalFormatRules();
  all.setFontFamily(FONT).setFontSize(10).setFontColor(C.ink).setVerticalAlignment('middle');
  setMeta_(sh, meta);

  // Row 1-2: title band
  sh.getRange(1, 1, 2, CT.length).setBackground(C.navy);
  sh.getRange(1, 1, 1, 9).merge().setValue((meta.company || sh.getName()) + '  —  1099 W-9 Tracker')
    .setFontSize(18).setFontWeight('bold').setFontColor(C.white);
  sh.getRange(2, 1, 1, 9).merge().setValue('Tax year ' + meta.year + '   ·   W-9s due ' + fmt_(setting_('W-9 Due Date')) +
    '   ·   1099s due to vendors and IRS by January 31, ' + (Number(meta.year) + 1)).setFontSize(10).setFontColor('#c9d3ea');
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
    sh.getRange(4, i + 1).setValue(k[0]).setFontSize(9).setFontColor(C.muted).setBackground(k[2]).setHorizontalAlignment('center');
    sh.getRange(5, i + 1).setFormula(k[1]).setFontSize(20).setFontWeight('bold').setFontColor(k[3]).setBackground(k[2]).setHorizontalAlignment('center');
  });
  sh.getRange(5, 8).setNumberFormat('0%');
  sh.getRange(4, 1, 2, kpi.length).setBorder(true, true, true, true, true, false, C.line, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(4, 22); sh.setRowHeight(5, 40);
  // Row 6: colour legend
  const legend = [['Colour key:', C.white, C.muted], ['Not sent', C.red, C.redInk], ['Draft ready', C.yellow, C.yellowInk], ['Emailed', C.blue, C.blueInk],
    ['W-9 received', C.green, C.greenInk], ['Follow-up due', C.orange, C.orangeInk]];
  legend.forEach(function (l, i) { sh.getRange(6, i + 1).setValue(l[0]).setBackground(l[1]).setFontColor(l[2]).setFontSize(9).setFontWeight(i ? 'bold' : 'normal').setHorizontalAlignment('center'); });
  sh.setRowHeight(7, 8);

  // Row 8: headers, row 9+: data
  sh.getRange(HR, 1, 1, CT.length).setValues([CT]).setFontWeight('bold').setFontColor(C.white).setBackground(C.navy).setWrap(true).setVerticalAlignment('middle');
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
  if (rows.length) sh.getRange(HR, 1, rows.length + 1, CT.length).createFilter();
  addButtons_(sh, [['drafts', 'draftsThisTab', 1, 6], ['followup', 'followUpSelected', 2, 0], ['allfu', 'followUpAllDue', 4, 0], ['sync', 'syncMine', 6, 0]]);
}

/** Adds clickable image buttons once (they stay put when the tab is rewritten). Click runs as the clicker. */
function addButtons_(sh, list) {
  if (sh.getImages().length) return;
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
    const td = tplData_(r, x.d.sh);
    const opts = { name: r['Sender Name'] || undefined };
    if (w9) opts.attachments = [w9];
    const dr = GmailApp.createDraft(String(r['Contact Email']).trim(), fill_(tpl.subject, td), fill_(tpl.body, td), opts);
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
  sh.getRange(1, 1, 2, head.length).setBackground(C.navy);
  sh.getRange(1, 1, 1, 8).merge().setValue('EP 1099 — W-9 Collection Dashboard').setFontSize(18).setFontWeight('bold').setFontColor(C.white);
  sh.getRange(2, 1, 1, 8).merge().setValue(tabs.length + ' companies   ·   counts update live as the team changes statuses   ·   1099s due January 31')
    .setFontColor('#c9d3ea');
  sh.setRowHeight(1, 40); sh.setRowHeight(2, 24); sh.setRowHeight(3, 46);
  sh.getRange(3, 7, 1, 6).merge().setValue('Click a company name below, then “Drafts for selected company”.').setFontSize(9).setFontStyle('italic').setFontColor(C.muted);
  sh.getRange(5, 1, 1, head.length).setValues([head]).setFontWeight('bold').setFontColor(C.white).setBackground(C.navy).setWrap(true).setHorizontalAlignment('center');
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
    sh.getRange(tr, 1, 1, head.length).setBackground('#e8eef7').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    sh.getRange(tr, 1).setHorizontalAlignment('left');
  } else {
    sh.getRange(6, 1).setValue('No companies yet — in the EP 1099 Tool, open a company and click “Send to email tracker”.').setFontColor(C.muted);
  }
  [240, 80, 220, 80, 90, 90, 110, 100, 100, 100, 100, 80].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(5);
  sh.setTabColor(C.navy);
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
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground(C.navy).setFontColor(C.white);
  sh.setFrozenRows(1); return sh;
}
function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg, 'EP 1099', 5); } catch (e) { } }
function log_(action, details, ss) {
  try { const sh = (ss || SpreadsheetApp.getActive()).getSheetByName(SYS.LOG); if (sh) sh.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]); } catch (e) { }
}

/* ======================= BUTTON IMAGES (do not edit) ======================= */
const BTN = {
  drafts: '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',
  followup: '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',
  sync: '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',
  company: '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',
  allfu: 'iVBORw0KGgoAAAANSUhEUgAAAL8AAAAiCAYAAAAOEgibAAATsUlEQVR42u2deZRU9ZXHP7/fe6/2ql5plgZpGlp2VFRAQIkYNQqoOCoRJ3uMMRrNMSZzYv4gyeTMJDnJRE00xpOJicMkQU2IYhLNqKgIooiIoCCrCM3SNL1U1/7e+/3mj1dVVC9s2qgxdc9pzunHq/t+9/7u797vXeq1oEhagNCN85+sE77YVxBytpDmudrJKISQlKlM/0iktRJmQGrlrECrZ3Qufu+OpRe3FOwcQIIWs2YtN0HoxuvWzZGB2g3SCi0S0jhXuxnKhl+mf0gSQmo3g5DGudIKLZKB2g2N162bA0J79q6FYJGWfFeoxgUvPGZYsXnazaCU4witZdnwy/RRiABaCCWlaQojgGvHl+1YMvMyFmkpAEYsWDnXtKLLlJ10wZXgXS9TmT46pDQYSlphw7G75u1cMuNx0fivK+uELbcJaYW1ygJlb1+mj2wYUEL60cpOakuNktjyc9IXjiiVVWXDL9NHPBGQSmWV9IUj2PJzEq0v064t0JQNv0z/BM4f6dm7vsyUpn+6djKIDyi5FcV/QOvSZL33tffK8/0vOJTmXe/us/2lkzIVdCikdjJIMzDdPJl1fClFt9TZdTU9981V2ttMAUb+Zq29ews8xAmm333x/CCM3nE1BYFPRI6+5H+vOilT9w3STkaZJ8vwhYBUxiWVcYvuuDJsInsYYyhgYBoCpb37NWAagmDIACBrK3K2PqHNLvDUJTzfT8PP2hq0piJiovP2b9sKpY8nKvctv2kKgr53r5My9e4DmCfF4wtIZRXTJ1Ywc2IFWoNS8OCT++hIuJjS22TH1fz4piZmTqpge3OKhd97k7ZOm4+fXc0vvj4agO898Da/eWIvVVGr6PmO9txkRvH96xuZO72Wti6bud98na6UUzwMJzmkkrVdRgwK8u1PN3DWmCjVMQuAK+54nVc2dxHyyyMeAikFyZTD+WdWc9/tnvzf/+3b/PxPu7lqVh0/v+3EdVKmI5N5cnIKzwN+fcEwpo6rKF5vbs3w4JP7qY5aHiTIe+lI0CAaMrt5vkjQ83I+69hGa+RhgMhD/QJPx9UF6I8ogUBFWFTy2SJMKckbSuGF1qC0PupaZP4Zd3y6gYumVJPNKX731AE64jYHO2xMQ/R6Zk++R5LfNI+sEylASFGUVWlQJSdMCooRtyB7X/o4ko6OpnOtvXt76lgpjdK9r2lKfu+xzgIPWSKL1qC1Pq6o+YEbv5SQzirGDQ9zRlO0m0HNOaeWJU+3oNRhXFxQcul9nsDHTu4Kyo8nHbK2wjIFUgi07s5TCMjmFF0pD4JFggZ+n0Rr6Ew65GyFlIKKiOkZvQBXQSLl/Z/WYJmScEBiGhLVx6K8aOfiOJqR9UFcpXn2tXauW7QRpOCUQQEClsR2NemMQ87x+JqGIOj31nM0+Y+kEykgnVNksgpXaYQQ+H2CkN8o6jiVVR78FFARNjENQc7RdCXtIhw1DUHW1nSlbACiIfOojqe9y8ZxNaYhiIVNhADb0RxK2kWnFgoYeb0f5ikEtKUcQBPwGYQDhyOhFJCxNems40U0AaYU+CyJ3ycRH3bjF0KQySkunlqNz5J0dDls3ZPi7LExzplQQdOwENuaUwR8BuB5ZiHohV+PhWeFyCfQGubNqGVcQ5iWjhx/eLqFbE4XeQoBmZzijKYoF0+pBuCvqw/x5s4kpim4ZnYdo+qDdCZcFv99P67WpHOKkN/gkqk1jB0exrIE2/akee61dtq7HEIBiVI94I6jGDE4yHUXDmRAhYUhBaPqQ9xz+xiytuJXy/bSkXCoCJvMnlzFmFNCWKZgz8EsL27sZPveNBVh86jy93UtkVFMbAxz9tgYdZU+khmXDdsTvPhGZ7HIMGVsjI+dUYWrNEueOcDulixNQ4Nced4AXAX/+/f97GvLMW54mHkzagDBspUH2bI7nXcS3R2TAK6fV8/Aaot9rTkeea6FnK0YPijAlbOGIQU8v76DFes7mNAY4fKZtXmerXSlHL58+QBMQ/DiG52sfiNOyC+9HDGraBoaYvrECuoqfTiuYn9bjk1vJ9myO92n0/lQGb/jaCojJp+YWoMG3tyV5O5HdvO7ReMJ+CQXT6lm4x8Snmd6D+VRL0wL7rq1iXkzaouQ4YpzB5C11eEQKgTZnGZCY4RbrhoGwK4DWV7bksAwBAtm1zFjYiXxpMPDz7ZwqNNm/IgwP7l5FBMbI92eu2V3ilvv3sKG7UlCfqO4GQWvN6zOzw2X1xfvH1kf5Cvz6+lKufxiaTPjG0LceeupjGsI9/CiDv+5eBe/f2o/4jizWIHAcRTf/tRwrp83BJ/VvW7x/PoObr9nGzv3eYfqlquGArC3NctdD+/m5ivri/rY1pziV8v2MffamuK1ZStbUUrjuqp40AXeOJhlwBfnDWHE4ABb96T50/MHydmaUwYG+NrV3udtR/P4ylaahgaLPMc2hBhVH2LE4EAeBsF//M/b/PLRZjRw0dnV3HlLE7Fwb7O84o4NvLI53k3v7xml9KfhG1KQzrqcPSbGmFNCCGDNpjgrN3Sy50AWgEum1hALmzhK826tXxqCeNJh4ccHFg3/3qV7GL1wNUufP8g54yvywolusMdxNY6ryWRV0Yt2Jh0cV3Mo7oXxUEBy161NTGyM8M6BDJd8Yz2TPvsyL2zo5NRhIe786qlEQ4aHc/M8lNKEA5KX3ogz+QtreOdABoCn1rYzeuFqpt2wBsOAn902mnENYXa3ZJn7b68z6bMv8cRLh6iKmvzwyyOZOq6CVMbFkMeSH7pSDtfMHshNVw7FZ0nue7SZUxeu5vofbSaZdjnvtEq+89kRWIZgzeY4O/dlcFzNhBFhgj7JxMZoEZdPaowQ9AsmNEZwXM2qDZ1s25PK504msfDhH38eChVgT1vcLuZ4tqPJ2p6eU1kPSuZsT+e2q5kwIsLt92xl8udf5tl17UgJd3xqOKeNipC1FddeOJBY2OTZde00Xfsioxeu5pPf2chjL7QiRf/3N+TJGB+6ZFoNUgpSGZeHl7ewZ1+Gh5a3ADC+McwZp0ZJZ1zku6zVKaUJ+A1mn1WF0tAWt7nv0b0k0i6//ds+1myKeyG/RFtCePjaNEoy4/yBNQ2BZUrSWZcpY2OMHR5GKc0jzx5k3ZYu2uI2Dy8/QM5RNA0LMWNCJak+1u8q7ZVW9eG+Rjrr0hp3OHtMjNHDQigNDz9zgBXrO2hpt7nvz3uxHQ+mzTmnhpyjjukUtAK/TzJ3ei1KQ2unzT1/2kM64/LoioM891oHSmlmTa6kcUiQlvYcr29PYBqCiY0R6uv8TBkbZef+DImUw1mjYwyrCzChMYxpCF7fnqQ1bnP6qDCr7z+Ll+4/i1W/OJPV95/F1efXEU86+CyJaQgvec1ny6U6LlS0hQDDEFiG4PdP7eepV9rYczDLA3/dh1IaKQXnT64ikXJJ5HOyUwYGWDB7IFPHxdjenOGGH2/m1S2FSpn+8MEeAeQcxaAaHxfmsXUyrbj6/DrmTK9lZH2wmCBdOq2G51/r6GaEJ/IcpSFgCSojJlJAR8LFdTXRoJdg7W/LHbvHLbrPrgrhJcgDq3z5BpngtgXDuG3BsF6fHjM8xNIVHg/VA5MXKjqF3y3TSy7r8nylgNa4TThgEA5K2hNeUm0aBtUxE58lj+nhlNYEfJKqqCd/W9zGVZpw0EBpONCeQ+YP9YBKH+u3JVj7VpzLZ9bSMDjAzEmV1FX5eOTZZgbX+JhzTi2zTq+itsIry67Z3IkhOK7+iNYakW9mFipufXbcgUOdNuGAScAv6Ui4OAosAZURE58puHdpM4Nr/EwbH+P71zcWP79mcxff+uU2trzj5Yr9dQD6zfilIUilHObNqGVQfqMHVFncenVv47l4ajU/fWg3HQnnhJGPLqkKdHQ5KA2VEQPDELQnHNCaQdW+I3ZOCxWiZFphmZL6Wv9hZRiCA+05hPCiy9OvtrP6jTixiInjaLqSDqGgwStvdRH09e2Fel5S2quAtRT4aqiJWfkI4eVHPstL+Nri3kE4VkCU+aJCe17+6qiFIQRdaQfbUQys8qGUxnE0BztyhIMGL2+KFw/Ijfm8ZOWGDkYMDjL/vAF8cd4QDClojzus25qgMmqxfnuSaV96pZsRp23l6SNfbDBND9rEUw61MV+vJmax9A3UVFgkMw6OK6mMGJiSotx+n2Tb3hRzvvkao+pDNAwO0DQ0xNeuGcbZY6LceEU9N/90CwGh6a+upey/7wx4m3LxlBpsxwv39z+2lx///h3ufGg3P1i8i+fWeeF4UJWPqeNiJNOeATiu7tasKZQpHbfv+q6UgkzW5ZlX25ECqmMWN1w+hEhQ8plLBjN5dLS4OUprDFli1BpmT66irtrHly4bwujh4eLzg36DNZvibN6VwpBeR3XZqkP8cPEufrWsmW3NaYYN8LO9OY1pHrkMWFi7qzRaQchvsGZznLd2p5ACrj6/jhmTKqittPjyFUOw8rz+8uIhfKZEq97yl+oEvBzm8VUeFq6ttPjK/KEE/ZLLZg7gvNMrkVKwfF0Hb+/PEAka7NybYctuT67TmqKks4pNu5Js3JHAzecCjqtZuyVOa4eNZQocRxNPOsWfzqRXgszkNO1dNkJAw6AgZ42OMqkxwo3z64sYv3TfCsdh4YWDuODMauoH+PncpYORUqCU9mCa1tx61TC+94VGAj7Jy5viPLO2ja6UA0DQb9Dfkypmf3V001mXSaMizDmnBoB1W5N8495tGDLfuUy7TJ9QwdN3nQHA5y8dzOOrWolFvBpzdcwqVnF8lijCh6CvNwxQriYWNvndUweYMjbGvBm13DR/KDfNH8qazXHe2Jlk8qlRavLd1WBA8vKbcZa/2s75k6u46mN1/MusOla83sGmt5NMGhmhKmJhGoLOhOKWu7bwk5tGMX1iBavvO7OXvD/7454jGn7Bw5mGoCJsovJ5RSLtcutdW/mvm0cxriHMX350Wrdqzw8W7+KlNzsJBbwoVpTfL1FKd9OJz5JEggYPPXOAoQP8XD+vnhvnez8FWrG+g+/8eqeXzxiCQ10Or21LMHZ4OF/xynCo0yFnp0lnFZH8OMWrW7tIZV1CAQutvbGKnlHHcV0efGIf08ZXUBU1efjfJ3Cww+aFDR3FSlbI333ftIa1m+P89KtNDB9UWu3ZxfrtXfhMycAqH9d+3EviS6mlPceDT+wvmVjsH9cvRi5YpfvF+HOK00ZGuHhqDVprXnwjzqoNnR4OzbsBrWHhhQOpipokM4rf/HUfl06rYcSQAIfiNn94uoV0xmXU0BBXnOtVcZ5e287at7oI9lBmsc6fL5GNbQhzsN3mj8+1MHNSBZNGRkhlFQ/+bT9ZWxVx8iem1tAwKMA7LRn+/HwrF02pZuzwEJ0Jh8V/P4DSmky+zj9tfIxxDWGCfkki7bK3Ncebbyd5pyXTp/611piG5LqLBlIZMdnenGbZylZvwE9COqOIhExmTKwo1vmbD2ZZtbGTHXvTREMmmZyicUiQ+ed58j/zajsvbuwsqZd7Onl1SxcBnySZr/NPydf5EyV1fqXBlzfedFZx5ugoF5xZBcDGHUmefLkN0xBccW5t0SCXrmhlR3MKn2V0q+/37DdksorTm6Kce1olUsD/vdLGwY4cn/nEYISAFa938vQrbVw+cwAP3DEWgK/dvZXn17fzyQsGIkVJnT/fN4mGDcY3hGkaGqIyaqK1pvlgjpUbOtndkullAx8K4y8tJybSbjFMhQK9F9uZcHDzWX5FxCSRcrEdhZHvFEoBOVsXw104aBDw9S10Ab8n0i452+MRDZlksi6ZnEKKfNc2n4wp5d1rO6rYmUxmXLK5vju8qaxLztZeUic8z+u3RK+aek98G8/LWPDQpU7CVd6wXdZR0KPDq5RX9ekmf8AgWOyU9tbJsTq8Bb0Vmn3J/P74fd7atIau/B4cT2e3VPfprCKdVWg04YCBZQg6Et4aIyET11VccGYNv/7WGAAW/fdO7lm6B58lcV3Vq8PrKk02p8g5uugwTdOTxWeKfh9xMPsT8/t9klDAOOLcBnlIUIqNK/PGqTncsfVZgroqX7cZkSM9E7yWfWGjldJEQyaxcPd5nYKhFJ+Xn0mJhUxEj3tV/t5oyOyGM49nzkSUyFg691LKtzAOUDrbUxode8mfhz196URpDxqGAkav2Z6esCPok4Tz+1O6toJOjmemp5RfyH/4cBfmdgprRAg6u1Q3uOa3JJYpGFDpDeT1tBHTEFj5EYjS2R71jzDb03NG50jJYM/a+LvhczQe7hE0pU/wXq10t1LmcXe5j7J2fZRnHk3+o+lEaeA49KW0ly8dS38n0tPpye/wPJV3YLc3p7n7kd0ArN0Sx2/J4rxUX3Lr9/GbOmLkNSvd8itKynSyvttwvBCWD+CVJqYwA1I7GcrfjChT/9vX8UPY93thwgxIqZzsKmEG0Fqr8naV6WQcgKP1bN7/9WglzADKya6SCPGYMCyNoGz8ZfonwGIoz97FYxJLPaByyYSUfumNTJWpTB/dl1ZJ6Zcql0xgqQfkjsUzWhRqoTB8EqTmw4HKylQm+n/eWGph+KRCLdyxeEaLZJGWO5fMeNy148ukFTaE9AkNDuUcoEwflRfVgiOkT0grbLh2fNnOJTMeZ5GWku+iZ81abu5YMvMyV7tzwWgxrLApTL/kfX3pR5nK1P+vUhCmXxpW2ASjxdXu3B1LZl42a9Zyk++iS/uX5T9OUaZ/qj9O8f8f7FxFq0dW7QAAAABJRU5ErkJggg=='
};
