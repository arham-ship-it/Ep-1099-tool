// Outside services: Google Drive (read W-9s from a shared folder), Gemini (optional AI W-9
// reading), and the Google Sheet email tracker (Apps Script web app).
import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.js?url';
import { ENTITIES, parseW9Fields, parseW9Text } from './logic.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

/* ---------------------------- Google Drive ---------------------------- */
let token = null, tokenExp = 0;

/** Ask the user to sign in with Google (read-only Drive). Needs an OAuth Client ID in Settings. */
export function driveToken(clientId) {
  return new Promise((resolve, reject) => {
    if (token && Date.now() < tokenExp) return resolve(token);
    if (!window.google?.accounts?.oauth2) return reject(new Error('Google sign-in did not load. Check your connection.'));
    if (!clientId) return reject(new Error('Add your Google OAuth Client ID in Settings first.'));
    const tc = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: 'https://www.googleapis.com/auth/drive.readonly',
      callback: r => {
        if (r.error) return reject(new Error(r.error_description || r.error));
        token = r.access_token; tokenExp = Date.now() + (Number(r.expires_in) - 60) * 1000; resolve(token);
      }
    });
    tc.requestAccessToken();
  });
}
export function folderIdFromUrl(url) {
  const m = String(url).match(/folders\/([a-zA-Z0-9_-]+)/) || String(url).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : String(url).trim();
}
/** List PDF/image files in a Drive folder (and one level of subfolders). Works with view-only shares. */
export async function listDriveFiles(folderUrl, clientId) {
  const t = await driveToken(clientId);
  const list = async id => {
    const q = encodeURIComponent(`'${id}' in parents and trashed = false`);
    const out = []; let page = '';
    do {
      const r = await fetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=nextPageToken,files(id,name,mimeType,modifiedTime)&pageSize=200&supportsAllDrives=true&includeItemsFromAllDrives=true${page ? '&pageToken=' + page : ''}`,
        { headers: { Authorization: 'Bearer ' + t } });
      if (!r.ok) throw new Error(r.status === 404 ? 'Folder not found, or it isn’t shared with the Google account you signed in with.' : 'Drive error ' + r.status);
      const j = await r.json(); out.push(...j.files); page = j.nextPageToken || '';
    } while (page);
    return out;
  };
  const root = await list(folderIdFromUrl(folderUrl));
  const files = root.filter(f => /pdf|image\//.test(f.mimeType));
  for (const sub of root.filter(f => f.mimeType === 'application/vnd.google-apps.folder')) {
    files.push(...(await list(sub.id)).filter(f => /pdf|image\//.test(f.mimeType)));
  }
  return files;
}
export async function downloadDriveFile(f, clientId) {
  const t = await driveToken(clientId);
  const r = await fetch(`https://www.googleapis.com/drive/v3/files/${f.id}?alt=media&supportsAllDrives=true`, { headers: { Authorization: 'Bearer ' + t } });
  if (!r.ok) throw new Error('Couldn’t download ' + f.name + ' (' + r.status + ')');
  return new File([await r.blob()], f.name, { type: f.mimeType });
}

/* ------------------------------ W-9 reading ----------------------------- */
/** Reads a W-9: PDF form fields first (fillable W-9s), then Gemini if a key is set. */
export async function readW9(file, settings) {
  let fields = [], text = '', img = null, note = '';
  try {
    if (/pdf/i.test(file.type) || /\.pdf$/i.test(file.name)) {
      const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
      const page = await pdf.getPage(1);
      text = (await page.getTextContent()).items.map(i => i.str).join(' ');
      fields = (await page.getAnnotations()).filter(a => a.fieldName).map(a => ({ name: a.fieldName, value: a.fieldValue }));
      const vp = page.getViewport({ scale: 1.6 }); const cv = document.createElement('canvas');
      cv.width = vp.width; cv.height = vp.height;
      await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
      img = await new Promise(r => cv.toBlob(r, 'image/png'));
    } else img = file;
  } catch (e) { note = 'Couldn’t open the file (' + e.message + ').'; }

  let x = { ...parseW9Text(text) };
  if (fields.length) {
    const f = parseW9Fields(fields);
    if (f.legal_name || f.entity_type !== 'Unknown') x = { ...x, ...f };
  }
  const needsAi = !x.legal_name || !x.entity_type || x.entity_type === 'Unknown';
  if (needsAi && settings.geminiKey && img) {
    try { x = { ...x, ...(await geminiReadW9(img, text, settings)), from: 'Gemini' }; }
    catch (e) { note = 'AI reading failed: ' + e.message; }
  } else if (needsAi && !note) {
    note = settings.geminiKey ? '' : 'Scanned or flattened form — pick the vendor and entity yourself (or add a Gemini key in Settings).';
  }
  return { x: (x.legal_name || x.entity_type) ? x : null, note };
}

async function geminiReadW9(img, text, settings) {
  const b64 = await new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1]); fr.onerror = rej; fr.readAsDataURL(img); });
  const prompt = `You are reading an IRS Form W-9 (or W-8) for a US accounts-payable team. Return ONLY JSON:
{"is_w9":bool,"legal_name":string,"business_name":string,"entity_type":one of ${JSON.stringify(ENTITIES)},"tin_last4":string,"signed":bool,"signature_date":"YYYY-MM-DD" or ""}.
Line 3a mapping: individual/sole proprietor -> "Individual / Sole proprietor"; C corporation -> "C Corporation"; S corporation -> "S Corporation"; partnership -> "Partnership"; trust/estate -> "Trust / Estate"; LLC taxed as C/S/P -> "LLC - C Corp"/"LLC - S Corp"/"LLC - Partnership"; disregarded single-member LLC -> "Single-member LLC"; a W-8 form -> "Foreign (W-8)". Use "Unknown" if unclear. Never return the full TIN.
PDF text layer (may be empty): ${String(text).slice(0, 6000)}`;
  const model = settings.geminiModel || 'gemini-2.5-flash';
  const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(settings.geminiKey)}`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: img.type || 'image/png', data: b64 } }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0 }
    })
  });
  if (!r.ok) throw new Error('Gemini ' + r.status);
  const j = await r.json();
  const out = JSON.parse(j.candidates?.[0]?.content?.parts?.[0]?.text || '{}');
  out.tin_last4 = String(out.tin_last4 || '').replace(/\D/g, '').slice(-4);
  return out;
}

/* ------------------------- Google Sheet email tracker ------------------------- */
/** POSTs vendor rows to the Sheet's Apps Script web app (doPost). Response can't be read (no-cors). */
export async function sendToSheet(settings, company, taxYear, sender, rows) {
  const body = JSON.stringify({ k: settings.sheetKey || '', c: company, y: taxYear, s: sender, r: rows });
  await fetch(settings.sheetUrl, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body });
}
/** GET fallback link (opens a confirmation page in a new tab). Split when long. */
export function sheetLinks(settings, company, taxYear, sender, rows) {
  const base = settings.sheetUrl.split('?')[0];
  const enc = s => { const bytes = new TextEncoder().encode(s); let bin = ''; bytes.forEach(b => bin += String.fromCharCode(b)); return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_'); };
  const make = rs => `${base}?k=${encodeURIComponent(settings.sheetKey || '')}&d=${encodeURIComponent(enc(JSON.stringify({ c: company, y: taxYear, s: sender, r: rs })))}`;
  const out = []; let cur = [];
  for (const row of rows) { cur.push(row); if (make(cur).length > 6500 && cur.length > 1) { cur.pop(); out.push({ url: make(cur), n: cur.length }); cur = [row]; } }
  if (cur.length) out.push({ url: make(cur), n: cur.length });
  return out;
}
