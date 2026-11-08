// Storage: browser localStorage. Swap this file for Firebase/Supabase to share trackers
// across a team (same method names) — see ANTIGRAVITY_PROMPT.md, task 1.
const KEY = 'ep1099.data.v1';
const listeners = new Set();

function load() {
  try { return JSON.parse(localStorage.getItem(KEY)) || { clients: {}, vendors: {}, settings: {} }; }
  catch { return { clients: {}, vendors: {}, settings: {} }; }
}
const data = load();

function persist() {
  try { localStorage.setItem(KEY, JSON.stringify(data)); }
  catch (e) { alertQuota(e); }
  listeners.forEach(fn => fn());
}
function alertQuota(e) { console.error(e); document.getElementById('saveState').textContent = 'Not saved — browser storage is full. Back up and delete old trackers.'; }

export const Store = {
  onChange(fn) { listeners.add(fn); },
  clients() { return data.clients; },
  vendors(cid) { return data.vendors[cid] || {}; },
  settings() { return data.settings; },
  putClient(cid, patch) { data.clients[cid] = { ...(data.clients[cid] || {}), ...patch, id: cid }; persist(); },
  putVendors(cid, obj) { data.vendors[cid] = obj; persist(); },
  patchVendor(cid, fk, patch) {
    const cur = data.vendors[cid] || (data.vendors[cid] = {});
    cur[fk] = { ...cur[fk], ...patch }; persist();
  },
  putSettings(patch) { data.settings = { ...data.settings, ...patch }; persist(); },
  deleteClient(cid) { delete data.clients[cid]; delete data.vendors[cid]; persist(); },
  exportJson() { return JSON.stringify(data, null, 2); },
  importJson(text) {
    const d = JSON.parse(text);
    if (!d || typeof d !== 'object' || !d.clients || !d.vendors) throw new Error('Not an EP 1099 backup file');
    data.clients = d.clients; data.vendors = d.vendors; data.settings = d.settings || {}; persist();
  }
};
