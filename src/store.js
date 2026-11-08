import { db, auth, isFirebaseConfigured, logoutUser } from './firebase.js';
import {
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot
} from 'firebase/firestore';
import { onAuthStateChanged } from 'firebase/auth';

const LOCAL_KEY = 'ep1099.data.v1';
const PERSONAL_KEY = 'ep1099.personalSettings';

const listeners = new Set();
const authListeners = new Set();

let currentUser = null;
let teamUsersList = [];
let unsubs = [];

function loadFallback() {
  try { return JSON.parse(localStorage.getItem(LOCAL_KEY)) || { clients: {}, vendors: {}, settings: {} }; }
  catch { return { clients: {}, vendors: {}, settings: {} }; }
}

function loadPersonalSettings() {
  try { return JSON.parse(localStorage.getItem(PERSONAL_KEY)) || {}; }
  catch { return {}; }
}

function savePersonalSettings(patch) {
  try {
    const cur = loadPersonalSettings();
    localStorage.setItem(PERSONAL_KEY, JSON.stringify({ ...cur, ...patch }));
  } catch (err) {
    console.error('Failed to save personal settings', err);
  }
}

const data = {
  clients: {},
  vendors: {},
  settings: {}
};

// Initialize with fallback local cache so UI displays instantly while connecting
const fallback = loadFallback();
data.clients = { ...(fallback.clients || {}) };
data.vendors = { ...(fallback.vendors || {}) };
data.settings = { ...(fallback.settings || {}) };

function updateSaveStateIndicator(text) {
  const el = document.getElementById('saveState');
  if (el) el.textContent = text;
}

function notify() {
  listeners.forEach(fn => {
    try { fn(); } catch (err) { console.error('Store listener error', err); }
  });
}

function notifyAuth() {
  authListeners.forEach(fn => {
    try { fn(currentUser); } catch (err) { console.error('Auth listener error', err); }
  });
}

// Attach real-time Firestore listeners once an authorized user signs in
function setupFirestoreListeners() {
  unsubs.forEach(u => { try { u(); } catch {} });
  unsubs = [];

  if (!db || !currentUser) return;

  updateSaveStateIndicator('Connecting to Firebase…');

  // 1. Clients collection
  const unsubClients = onSnapshot(collection(db, 'clients'), snapshot => {
    const clients = {};
    snapshot.forEach(docSnap => {
      clients[docSnap.id] = { ...docSnap.data(), id: docSnap.id };
    });
    data.clients = clients;
    updateSaveStateIndicator('☁ Synced with Firebase');
    notify();
  }, err => {
    console.error('Firestore clients error:', err);
    updateSaveStateIndicator('Firebase sync paused');
  });
  unsubs.push(unsubClients);

  // 2. Vendors collection
  const unsubVendors = onSnapshot(collection(db, 'vendors'), snapshot => {
    const vendors = {};
    snapshot.forEach(docSnap => {
      vendors[docSnap.id] = docSnap.data();
    });
    data.vendors = vendors;
    updateSaveStateIndicator('☁ Synced with Firebase');
    notify();
  }, err => {
    console.error('Firestore vendors error:', err);
  });
  unsubs.push(unsubVendors);

  // 3. Settings/main doc
  const unsubSettings = onSnapshot(doc(db, 'settings', 'main'), docSnap => {
    data.settings = docSnap.exists() ? docSnap.data() : {};
    notify();
  }, err => {
    console.error('Firestore settings error:', err);
  });
  unsubs.push(unsubSettings);

  // 4. Team users directory
  const unsubUsers = onSnapshot(collection(db, 'users'), snapshot => {
    const list = [];
    snapshot.forEach(docSnap => {
      list.push(docSnap.data());
    });
    teamUsersList = list.sort((a, b) => (a.name || a.email).localeCompare(b.name || b.email));
    notify();
  }, err => {
    console.error('Firestore users error:', err);
  });
  unsubs.push(unsubUsers);
}

// Track Authentication state
if (auth) {
  onAuthStateChanged(auth, async user => {
    if (user) {
      const email = user.email || '';
      const isDomainUser = email.toLowerCase().endsWith('@enduringplanet.com');

      if (!isDomainUser) {
        console.warn('Access denied: Unauthorized email domain', email);
        await logoutUser();
        currentUser = null;
        unsubs.forEach(u => { try { u(); } catch {} });
        unsubs = [];
        notifyAuth();
        alert('Access denied: Only @enduringplanet.com Google accounts can access this tool.');
        return;
      }

      currentUser = {
        uid: user.uid,
        email: user.email,
        name: user.displayName || user.email.split('@')[0],
        photoURL: user.photoURL || ''
      };

      // Register or update user record in team directory
      if (db) {
        try {
          await setDoc(doc(db, 'users', user.uid), {
            uid: user.uid,
            email: user.email,
            name: currentUser.name,
            photoURL: currentUser.photoURL,
            lastSeen: new Date().toISOString()
          }, { merge: true });
        } catch (err) {
          console.warn('Could not update user directory:', err);
        }
      }

      setupFirestoreListeners();
    } else {
      currentUser = null;
      unsubs.forEach(u => { try { u(); } catch {} });
      unsubs = [];
      updateSaveStateIndicator('Signed out');
    }
    notifyAuth();
  });
}

export const Store = {
  onChange(fn) { listeners.add(fn); },
  onAuthChange(fn) { authListeners.add(fn); fn(currentUser); },
  isConfigured() { return isFirebaseConfigured; },
  currentUser() { return currentUser; },
  teamUsers() { return teamUsersList; },

  clients() { return data.clients; },
  vendors(cid) { return data.vendors[cid] || {}; },

  settings() {
    const personal = loadPersonalSettings();
    return {
      ...(data.settings || {}),
      ...personal
    };
  },

  async putClient(cid, patch) {
    const docData = { ...(data.clients[cid] || {}), ...patch, id: cid };
    data.clients[cid] = docData;
    notify();

    if (db && currentUser) {
      try {
        await setDoc(doc(db, 'clients', cid), docData, { merge: true });
      } catch (err) {
        console.error('Failed to save client to Firestore', err);
        updateSaveStateIndicator('Error saving to Firebase');
      }
    } else {
      // Local fallback
      try { localStorage.setItem(LOCAL_KEY, JSON.stringify(data)); } catch {}
    }
  },

  async putVendors(cid, obj) {
    data.vendors[cid] = obj;
    notify();

    if (db && currentUser) {
      try {
        await setDoc(doc(db, 'vendors', cid), obj);
      } catch (err) {
        console.error('Failed to save vendors to Firestore', err);
        updateSaveStateIndicator('Error saving to Firebase');
      }
    } else {
      try { localStorage.setItem(LOCAL_KEY, JSON.stringify(data)); } catch {}
    }
  },

  async patchVendor(cid, fk, patch) {
    const cur = data.vendors[cid] || (data.vendors[cid] = {});
    cur[fk] = { ...(cur[fk] || {}), ...patch };
    notify();

    if (db && currentUser) {
      try {
        await setDoc(doc(db, 'vendors', cid), { [fk]: cur[fk] }, { merge: true });
      } catch (err) {
        console.error('Failed to patch vendor in Firestore', err);
      }
    } else {
      try { localStorage.setItem(LOCAL_KEY, JSON.stringify(data)); } catch {}
    }
  },

  async putSettings(patch) {
    const { geminiKey, geminiModel, ...teamSettings } = patch;
    if (geminiKey !== undefined || geminiModel !== undefined) {
      savePersonalSettings({ geminiKey, geminiModel });
    }

    if (Object.keys(teamSettings).length) {
      data.settings = { ...(data.settings || {}), ...teamSettings };
      notify();
      if (db && currentUser) {
        try {
          await setDoc(doc(db, 'settings', 'main'), teamSettings, { merge: true });
        } catch (err) {
          console.error('Failed to save team settings to Firestore', err);
        }
      }
    } else {
      notify();
    }
  },

  async deleteClient(cid) {
    delete data.clients[cid];
    delete data.vendors[cid];
    notify();

    if (db && currentUser) {
      try {
        await deleteDoc(doc(db, 'clients', cid));
        await deleteDoc(doc(db, 'vendors', cid));
      } catch (err) {
        console.error('Failed to delete client from Firestore', err);
      }
    } else {
      try { localStorage.setItem(LOCAL_KEY, JSON.stringify(data)); } catch {}
    }
  },

  exportJson() {
    return JSON.stringify({
      clients: data.clients,
      vendors: data.vendors,
      settings: data.settings
    }, null, 2);
  },

  async importJson(text) {
    const d = JSON.parse(text);
    if (!d || typeof d !== 'object' || !d.clients || !d.vendors) {
      throw new Error('Not an EP 1099 backup file');
    }

    data.clients = d.clients;
    data.vendors = d.vendors;
    data.settings = d.settings || {};
    notify();

    if (db && currentUser) {
      for (const [cid, cl] of Object.entries(d.clients)) {
        await setDoc(doc(db, 'clients', cid), { ...cl, id: cid }, { merge: true });
      }
      for (const [cid, vs] of Object.entries(d.vendors)) {
        await setDoc(doc(db, 'vendors', cid), vs);
      }
      if (d.settings) {
        await setDoc(doc(db, 'settings', 'main'), d.settings, { merge: true });
      }
    } else {
      try { localStorage.setItem(LOCAL_KEY, JSON.stringify(data)); } catch {}
    }
  },

  async importFromLocalStorage() {
    const raw = localStorage.getItem(LOCAL_KEY);
    if (!raw) return { count: 0, message: 'No local browser trackers found to import.' };
    let parsed;
    try { parsed = JSON.parse(raw); } catch {
      return { count: 0, message: 'Failed to read local browser data.' };
    }
    if (!parsed || (!parsed.clients && !parsed.vendors)) {
      return { count: 0, message: 'No local trackers found in this browser.' };
    }

    let count = 0;
    if (parsed.clients) {
      for (const [cid, cl] of Object.entries(parsed.clients)) {
        await this.putClient(cid, cl);
        count++;
      }
    }
    if (parsed.vendors) {
      for (const [cid, vs] of Object.entries(parsed.vendors)) {
        await this.putVendors(cid, vs);
      }
    }
    if (parsed.settings) {
      await this.putSettings(parsed.settings);
    }
    return {
      count,
      message: `Successfully imported ${count} company tracker(s) from this browser into team Firebase!`
    };
  }
};
