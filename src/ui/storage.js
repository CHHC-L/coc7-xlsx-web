// Local persistence of character cards: IndexedDB, falling back to localStorage.

const DB = 'coc7card';
const STORE = 'cards';
const LS_PREFIX = 'coc7card:';

function b64(bytes) {
  let s = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return btoa(s);
}

function unb64(s) {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function idb() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) { reject(new Error('no indexedDB')); return; }
    let req;
    try { req = indexedDB.open(DB, 1); } catch (e) { reject(e); return; }
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'id' });
      if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
    setTimeout(() => reject(new Error('indexedDB timeout')), 4000);
  });
}

function tx(db, store, mode, fn) {
  return new Promise((resolve, reject) => {
    const t = db.transaction(store, mode);
    const s = t.objectStore(store);
    let out;
    const r = fn(s);
    if (r) r.onsuccess = () => { out = r.result; };
    t.oncomplete = () => resolve(out);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  });
}

class IdbStore {
  constructor(db) { this.db = db; this.kind = 'indexedDB'; }
  async list() {
    const all = await tx(this.db, STORE, 'readonly', (s) => s.getAll());
    return (all || []).map(({ base, ...rest }) => ({ ...rest, hasBase: !!base })).sort((a, b) => b.updated - a.updated);
  }
  get(id) { return tx(this.db, STORE, 'readonly', (s) => s.get(id)); }
  put(rec) { return tx(this.db, STORE, 'readwrite', (s) => s.put(rec)); }
  del(id) { return tx(this.db, STORE, 'readwrite', (s) => s.delete(id)); }
  getMeta(k) { return tx(this.db, 'meta', 'readonly', (s) => s.get(k)); }
  setMeta(k, v) { return tx(this.db, 'meta', 'readwrite', (s) => s.put(v, k)); }
}

class LsStore {
  constructor() { this.kind = 'localStorage'; }
  _ids() { try { return JSON.parse(localStorage.getItem(LS_PREFIX + 'ids') || '[]'); } catch (e) { return []; } }
  async list() {
    const out = [];
    for (const id of this._ids()) {
      try {
        const r = JSON.parse(localStorage.getItem(LS_PREFIX + 'rec:' + id));
        if (r) { const { base, ...rest } = r; out.push({ ...rest, hasBase: !!base }); }
      } catch (e) { /* skip */ }
    }
    return out.sort((a, b) => b.updated - a.updated);
  }
  async get(id) {
    const r = JSON.parse(localStorage.getItem(LS_PREFIX + 'rec:' + id) || 'null');
    if (r && r.base) r.base = unb64(r.base);
    return r;
  }
  async put(rec) {
    const r = { ...rec, base: rec.base ? b64(rec.base) : null };
    localStorage.setItem(LS_PREFIX + 'rec:' + rec.id, JSON.stringify(r));
    const ids = this._ids();
    if (!ids.includes(rec.id)) { ids.push(rec.id); localStorage.setItem(LS_PREFIX + 'ids', JSON.stringify(ids)); }
  }
  async del(id) {
    localStorage.removeItem(LS_PREFIX + 'rec:' + id);
    localStorage.setItem(LS_PREFIX + 'ids', JSON.stringify(this._ids().filter((x) => x !== id)));
  }
  async getMeta(k) { try { return JSON.parse(localStorage.getItem(LS_PREFIX + 'meta:' + k)); } catch (e) { return null; } }
  async setMeta(k, v) { localStorage.setItem(LS_PREFIX + 'meta:' + k, JSON.stringify(v)); }
}

class MemStore {
  constructor() { this.kind = 'memory'; this.m = new Map(); this.meta = new Map(); }
  async list() { return [...this.m.values()].map(({ base, ...r }) => ({ ...r, hasBase: !!base })).sort((a, b) => b.updated - a.updated); }
  async get(id) { return this.m.get(id) || null; }
  async put(rec) { this.m.set(rec.id, rec); }
  async del(id) { this.m.delete(id); }
  async getMeta(k) { return this.meta.get(k); }
  async setMeta(k, v) { this.meta.set(k, v); }
}

export async function openStore() {
  try {
    const db = await idb();
    const s = new IdbStore(db);
    await s.getMeta('probe');
    return s;
  } catch (e) {
    try {
      localStorage.setItem(LS_PREFIX + 'probe', '1');
      localStorage.removeItem(LS_PREFIX + 'probe');
      return new LsStore();
    } catch (e2) {
      return new MemStore();
    }
  }
}

export { b64, unb64 };
