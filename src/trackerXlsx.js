// Builds the formatted, colour-coded 1099 tracker workbook (.xlsx) for one company.
// Same layout as the company tabs in the EP 1099 Email Tracker Google Sheet.
import ExcelJS from 'exceljs';
import { S, money } from './logic.js';

const C = {
  navy: '1C2A4D', teal: '00B38C', white: 'FFFFFF', ink: '1C2A4D', muted: '55627D', line: 'DBE3EA', band: 'F5F8FB', sub: 'C9D3EA',
  red: 'FDE2E6', redInk: 'B4233B', yellow: 'FFF0CC', yellowInk: '8A5A00', blue: 'E1EAFF', blueInk: '2357C6',
  green: 'DCF4E8', greenInk: '11845B', orange: 'FFE3D1', orangeInk: 'B3470E', grey: 'ECEFF3', greyInk: '5F6B80'
};
const FONT = 'Calibri';
const argb = h => ({ argb: 'FF' + h });
const fill = h => ({ type: 'pattern', pattern: 'solid', fgColor: argb(h) });
const cfFill = h => ({ type: 'pattern', pattern: 'solid', bgColor: argb(h) });
const thin = { style: 'thin', color: argb(C.line) };

export const W9_STATUSES = ['To request', 'Emailed', 'Follow-up 1', 'Follow-up 2', 'Final notice', 'Received', 'Verified', 'Check first', 'Not needed'];
const fromApp = { req: 'To request', trk: 'Emailed', rec: 'Received', na: 'Not needed' };

const COLS = [
  { h: 'Vendor', w: 34 }, { h: 'Reportable ($)', w: 14 }, { h: '1099 Status', w: 14 }, { h: 'Form / Box', w: 13 },
  { h: 'W-9 Status', w: 14 }, { h: 'Entity Type (W-9)', w: 22 }, { h: 'Contact Name', w: 18 }, { h: 'Contact Email', w: 28 },
  { h: 'Last Emailed', w: 13 }, { h: 'Next Follow-up', w: 14 }, { h: 'Assigned To', w: 16 }, { h: 'Why (1099 reason)', w: 60 }, { h: 'Notes', w: 30 }
];
const HR = 8, FIRST = HR + 1, LAST = 2000, FU_DAYS = 7;

/** c = client record, vs = vendors with .ev (evaluate) and .w (w9State). Returns a Blob. */
export async function buildTrackerXlsx(c, vs, opts = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'EP 1099 Tool'; wb.created = new Date();
  const order = { 'Eligible': 0, 'Needs W-9': 1, 'Needs review': 2, 'Not eligible': 3 };
  const sorted = [...vs].sort((a, b) => (order[a.ev.s] - order[b.ev.s]) || (b.ev.rep - a.ev.rep));
  const tracked = sorted.filter(v => v.ev.s !== S.NOT);

  /* ---------------- Sheet 1: 1099 Tracker ---------------- */
  const ws = wb.addWorksheet('1099 Tracker', { views: [{ state: 'frozen', xSplit: 1, ySplit: HR, showGridLines: false }], properties: { tabColor: argb(C.teal) } });
  ws.columns = COLS.map(c => ({ width: c.w }));
  const L = i => ws.getColumn(i).letter;
  const col = h => COLS.findIndex(c => c.h === h) + 1;
  const cST = L(col('1099 Status')), cW9 = L(col('W-9 Status')), cLE = L(col('Last Emailed'));
  const rng = l => `${l}${FIRST}:${l}${LAST}`;

  // Title band
  for (let r = 1; r <= 2; r++) for (let i = 1; i <= COLS.length; i++) ws.getCell(r, i).fill = fill(C.navy);
  ws.mergeCells(1, 1, 1, 8); ws.mergeCells(2, 1, 2, 8);
  ws.getCell(1, 1).value = `${c.name}  —  1099 W-9 Tracker`;
  ws.getCell(1, 1).font = { name: FONT, size: 18, bold: true, color: argb(C.white) };
  ws.getCell(2, 1).value = `Tax year ${c.taxYear}   ·   Sender: ${c.senderName || '—'} ${c.senderEmail ? '<' + c.senderEmail + '>' : ''}   ·   1099s due to vendors and IRS by January 31, ${c.taxYear + 1}   ·   Exported ${new Date().toLocaleDateString('en-US')}`;
  ws.getCell(2, 1).font = { name: FONT, size: 10, color: argb(C.sub) };
  ws.getRow(1).height = 32; ws.getRow(2).height = 20; ws.getRow(3).height = 8;

  // Live counts (formulas update as the team edits W-9 Status)
  const kpis = [
    ['Need a 1099', `COUNTIF(${rng(cST)},"Eligible")+COUNTIF(${rng(cST)},"Needs W-9")`, C.band, C.ink],
    ['To request', `COUNTIF(${rng(cW9)},"To request")`, C.red, C.redInk],
    ['Emailed – waiting', `COUNTIF(${rng(cW9)},"Emailed")+COUNTIF(${rng(cW9)},"Follow-up*")+COUNTIF(${rng(cW9)},"Final notice")`, C.blue, C.blueInk],
    ['Follow-up due', `COUNTIFS(${rng(cLE)},"<="&(TODAY()-${FU_DAYS}),${rng(cW9)},"<>Received",${rng(cW9)},"<>Verified",${rng(cW9)},"<>Not needed")`, C.orange, C.orangeInk],
    ['W-9 received', `COUNTIF(${rng(cW9)},"Received")+COUNTIF(${rng(cW9)},"Verified")`, C.green, C.greenInk],
    ['% complete', 'IF(A5=0,0,E5/A5)', C.green, C.greenInk],
    ['Needs review', `COUNTIF(${rng(cST)},"Needs review")`, C.orange, C.orangeInk],
    ['Vendors paid', null, C.band, C.ink]
  ];
  const st = { need: 0, req: 0, trk: 0, due: 0, rec: 0, rev: 0 };
  tracked.forEach(v => { if (v.ev.s === S.ELIG || v.ev.s === S.W9) st.need++; if (v.w === 'req') st.req++; if (v.w === 'trk') st.trk++; if (v.w === 'rec') st.rec++; if (v.ev.s === S.REV) st.rev++; });
  const results = [st.need, st.req, st.trk, 0, st.rec, st.need ? st.rec / st.need : 0, st.rev, vs.length];
  kpis.forEach((k, i) => {
    const lab = ws.getCell(4, i + 1), val = ws.getCell(5, i + 1);
    lab.value = k[0]; lab.font = { name: FONT, size: 9, color: argb(C.muted) };
    val.value = k[1] ? { formula: k[1], result: results[i] } : vs.length;
    val.font = { name: FONT, size: 20, bold: true, color: argb(k[3]) };
    [lab, val].forEach(x => { x.fill = fill(k[2]); x.alignment = { horizontal: 'center', vertical: 'middle' }; x.border = { top: thin, left: thin, bottom: thin, right: thin }; });
  });

  ws.getRow(4).height = 18; ws.getRow(5).height = 34;

  // Legend
  [['Colour key:', C.white, C.muted], ['To request', C.red, C.redInk], ['Emailed', C.blue, C.blueInk], ['Follow-up due', C.orange, C.orangeInk],
    ['Received', C.green, C.greenInk], ['Needs W-9', C.yellow, C.yellowInk], ['Not needed', C.grey, C.greyInk]].forEach((l, i) => {
    const x = ws.getCell(6, i + 1); x.value = l[0]; x.fill = fill(l[1]);
    x.font = { name: FONT, size: 9, bold: i > 0, color: argb(l[2]) }; x.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.getRow(7).height = 6;

  // Header + rows
  const hdr = ws.getRow(HR);
  COLS.forEach((c, i) => { const x = hdr.getCell(i + 1); x.value = c.h; x.font = { name: FONT, size: 10, bold: true, color: argb(C.white) }; x.fill = fill(C.navy); x.alignment = { vertical: 'middle', wrapText: true }; });
  hdr.height = 30;
  tracked.forEach((v, i) => {
    const r = ws.getRow(FIRST + i);
    const w9 = v.ev.s === S.REV ? 'Check first' : fromApp[v.w];
    const last = v.rq && v.rq.d ? new Date(v.rq.d) : null;
    const vals = [v.n, Math.round(v.ev.rep * 100) / 100, v.ev.s, v.ev.fb || '—', w9, v.e || 'Unknown', v.cn || '', v.ce || '', last, null, opts.assignee || '', v.ev.why, ''];
    vals.forEach((x, j) => { r.getCell(j + 1).value = x; });
    const n = FIRST + i;
    r.getCell(col('Next Follow-up')).value = { formula: `IF(OR(${cLE}${n}="",${cW9}${n}="Received",${cW9}${n}="Verified",${cW9}${n}="Not needed"),"",${cLE}${n}+${FU_DAYS})` };
    r.eachCell({ includeEmpty: true }, (cell, j) => {
      cell.font = { name: FONT, size: 10, color: argb(C.ink), bold: j === 1 || j === col('W-9 Status') };
      cell.alignment = { vertical: 'top', wrapText: j === col('Why (1099 reason)') || j === 1 };
      cell.border = { bottom: thin };
    });
  });
  const last = Math.max(FIRST, FIRST + tracked.length - 1);
  ws.getColumn(col('Reportable ($)')).numFmt = '$#,##0';
  ws.getColumn(col('Last Emailed')).numFmt = 'mmm d, yyyy';
  ws.getColumn(col('Next Follow-up')).numFmt = 'mmm d, yyyy';
  ws.autoFilter = { from: { row: HR, column: 1 }, to: { row: last, column: COLS.length } };
  for (let i = 1; i <= kpis.length; i++) ws.getCell(5, i).numFmt = i === 6 ? '0%' : '0';  // after column formats

  // Dropdowns (team updates these)
  ws.dataValidations.add(rng(cW9), { type: 'list', allowBlank: true, formulae: ['"' + W9_STATUSES.join(',') + '"'], showErrorMessage: false });
  ws.dataValidations.add(rng(L(col('Entity Type (W-9)'))), { type: 'list', allowBlank: true,
    formulae: ['"Unknown,Individual / Sole proprietor,Single-member LLC,Partnership,LLC - Partnership,LLC - C Corp,LLC - S Corp,C Corporation,S Corporation,Trust / Estate,Exempt / Government,Foreign (W-8)"'], showErrorMessage: false });

  // Colour coding (updates when the team changes a status)
  // expression rules (work in Excel, Google Sheets and LibreOffice)
  const txt = (ref, text, bg, ink) => { const top = ref.split(':')[0]; return { ref, rules: [{ type: 'expression', priority: 1,
    formulae: [`ISNUMBER(SEARCH("${text}",$${top}))`], style: { fill: cfFill(bg), font: { color: argb(ink), bold: true } } }] }; };
  [
    { ref: `A${FIRST}:A${LAST}`, rules: [{ type: 'expression', priority: 1, formulae: [`AND($${cLE}${FIRST}<>"",$${cLE}${FIRST}<=TODAY()-${FU_DAYS},$${cW9}${FIRST}<>"Received",$${cW9}${FIRST}<>"Verified",$${cW9}${FIRST}<>"Not needed")`], style: { fill: cfFill(C.orange), font: { color: argb(C.orangeInk), bold: true } } }] },
    txt(rng(cW9), 'To request', C.red, C.redInk), txt(rng(cW9), 'Emailed', C.blue, C.blueInk), txt(rng(cW9), 'Follow-up', C.blue, C.blueInk),
    txt(rng(cW9), 'Final notice', C.orange, C.orangeInk), txt(rng(cW9), 'Received', C.green, C.greenInk), txt(rng(cW9), 'Verified', C.green, C.greenInk),
    txt(rng(cW9), 'Check first', C.orange, C.orangeInk), txt(rng(cW9), 'Not needed', C.grey, C.greyInk),
    txt(rng(cST), 'Eligible', C.green, C.greenInk), txt(rng(cST), 'Needs W-9', C.yellow, C.yellowInk), txt(rng(cST), 'Needs review', C.orange, C.orangeInk)
  ].forEach((cf, i) => { cf.rules[0].priority = i + 1; ws.addConditionalFormatting(cf); });

  /* ---------------- Sheet 2: All vendors ---------------- */
  const all = wb.addWorksheet('All vendors', { views: [{ state: 'frozen', ySplit: 1, showGridLines: false }], properties: { tabColor: argb(C.navy) } });
  all.columns = [{ header: 'Vendor', width: 34 }, { header: 'Total paid ($)', width: 14 }, { header: 'Card / PayPal ($)', width: 15 }, { header: 'Reportable ($)', width: 14 },
    { header: '1099 Status', width: 14 }, { header: 'Form / Box', width: 13 }, { header: 'Entity Type', width: 22 }, { header: 'Why', width: 80 }];
  all.getRow(1).eachCell(x => { x.font = { name: FONT, bold: true, color: argb(C.white) }; x.fill = fill(C.navy); });
  sorted.forEach(v => all.addRow([v.n, v.t, v.x, Math.round(v.ev.rep * 100) / 100, v.ev.s, v.ev.fb || '—', v.e || 'Unknown', v.ev.why]));
  [2, 3, 4].forEach(i => { all.getColumn(i).numFmt = '$#,##0'; });
  all.getColumn(8).alignment = { wrapText: true, vertical: 'top' };
  all.autoFilter = { from: 'A1', to: { row: Math.max(1, sorted.length + 1), column: 8 } };
  ['Eligible', 'Needs W-9', 'Needs review', 'Not eligible'].forEach((t, i) => all.addConditionalFormatting(
    { ref: `E2:E${sorted.length + 1}`, rules: [{ type: 'expression', priority: i + 1, formulae: [`ISNUMBER(SEARCH("${t}",$E2))`],
      style: { fill: cfFill([C.green, C.yellow, C.orange, C.grey][i]), font: { color: argb([C.greenInk, C.yellowInk, C.orangeInk, C.greyInk][i]), bold: true } } }] }));

  /* ---------------- Sheet 3: How to use ---------------- */
  const h = wb.addWorksheet('How to use', { views: [{ showGridLines: false }] });
  h.getColumn(1).width = 110;
  [
    [`${c.name} — how to use this tracker`, 16, true],
    ['1. “1099 Tracker” lists every vendor that needs a 1099 (or needs review). “All vendors” lists everyone paid, with the reason.', 11],
    ['2. Update the W-9 Status dropdown as you go: To request → Emailed → Follow-up 1/2 → Received → Verified. The counts at the top update by themselves.', 11],
    ['3. Put the date you emailed in “Last Emailed”. “Next Follow-up” fills in, and the vendor turns orange when a follow-up is due (7 days).', 11],
    ['4. To email vendors: in the EP 1099 Tool click “Send to email tracker”, then in the EP 1099 Email Tracker Google Sheet click the green “Create W-9 drafts” button. Drafts land in the sender’s Gmail.', 11],
    ['5. Upload returned W-9s into the EP 1099 Tool so entity types and statuses update for everyone.', 11],
    ['Rules: tax year ' + c.taxYear + ' thresholds — $2,000 for 1099-NEC and most 1099-MISC (royalties $10). Card/PayPal payments are excluded (1099-K). Corporations are exempt except attorneys and medical. Not tax advice — confirm with your CPA.', 10]
  ].forEach(([t, size, b], i) => { const x = h.getCell(i + 1, 1); x.value = t; x.font = { name: FONT, size, bold: !!b, color: argb(C.ink) }; x.alignment = { wrapText: true, vertical: 'top' }; h.getRow(i + 1).height = i ? 34 : 26; });

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
