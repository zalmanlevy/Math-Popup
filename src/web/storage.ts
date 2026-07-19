// Local-only settings + notes persistence for the Math Popup web/PWA build.
// Backs window.mathPopup.getSettings/setSettings (see shim.ts) with IndexedDB plus
// an in-memory cache, so the unchanged desktop renderer runs in the browser. No
// network, no login — everything lives in the user's browser, mirroring the desktop
// app's local settings.json.
import { DEFAULT_SETTINGS, type Settings, type Mode } from '../shared/types';

const DB_NAME = 'math-popup';
const STORE = 'kv';
const KEY = 'settings.v1';
const LOCAL_KEY = 'math-popup.settings.v1';
const CURRENT_SCHEMA = 1;

let cache: Settings | null = null;
let loadPromise: Promise<Settings> | null = null;
let writeChain: Promise<void> = Promise.resolve();
let saveTimer: ReturnType<typeof setTimeout> | null = null;
let pending: Settings | null = null;
// False until we've genuinely read the store (a record OR a confirmed-empty store).
// A transient read failure leaves this false so we never overwrite data we couldn't
// read — persistence resumes once a read succeeds.
let storageReadable = false;

const bus = typeof BroadcastChannel !== 'undefined' ? new BroadcastChannel('math-popup') : null;
let changedCb: ((s: Settings) => void) | null = null;

// ---------- promise-wrapped IndexedDB (no dependency) ----------
let dbPromise: Promise<IDBDatabase> | null = null;
function getDB(): Promise<IDBDatabase> {
  return (dbPromise ??= new Promise<IDBDatabase>((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }).catch((err) => { dbPromise = null; throw err; })); // allow reopen after a failed open
}
async function idbGet<T>(key: string): Promise<T | undefined> {
  const db = await getDB();
  return new Promise<T | undefined>((resolve, reject) => {
    const req = db.transaction(STORE, 'readonly').objectStore(STORE).get(key);
    req.onsuccess = () => resolve(req.result as T | undefined);
    req.onerror = () => reject(req.error);
  });
}
async function idbSet(key: string, value: unknown): Promise<void> {
  const db = await getDB();
  return new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

interface Envelope { schemaVersion: number; app: string; settings: Settings; savedAt?: number; }
function envelope(settings: Settings): Envelope {
  return { schemaVersion: CURRENT_SCHEMA, app: __APP_VERSION__, settings, savedAt: Date.now() };
}

// iPadOS can suspend a PWA before an IndexedDB transaction completes. Keep a
// synchronous localStorage mirror and choose the newest valid copy on startup.
// IndexedDB remains the primary large-capacity store; the mirror is a recovery
// path for settings/navigation writes interrupted by suspension.
function readLocalEnvelope(): Envelope | null {
  try {
    const raw = localStorage.getItem(LOCAL_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Envelope;
    return parsed?.settings ? parsed : null;
  } catch { return null; }
}

function writeLocalEnvelope(env: Envelope): void {
  try { localStorage.setItem(LOCAL_KEY, JSON.stringify(env)); }
  catch (err) { console.warn('[math-popup] local settings mirror failed', err); }
}
// Forward-only migrations for future breaking schema changes (v1 -> v2 lands here).
const migrations: Record<number, (s: Settings) => Settings> = {};

// Mirror the intent of main/store.ts migrateSettings: guarantee pages + lineModes.
function normalize(s: Settings): Settings {
  if (!Array.isArray(s.pages) || s.pages.length === 0) {
    s.pages = [{
      id: Date.now().toString(),
      title: 'Page 1',
      content: (s as { noteContent?: string }).noteContent || '',
      mode: s.mode || 'math',
    }];
    s.activePageId = s.pages[0].id;
  }
  if (!Array.isArray(s.closedPages)) s.closedPages = [];
  if (!Array.isArray(s.archivedPages)) s.archivedPages = [];
  if (!Array.isArray(s.obsidianRecentNotes)) s.obsidianRecentNotes = [];
  for (const p of s.pages) {
    if (!Array.isArray(p.lineModes)) {
      const n = Math.max(1, (p.content ?? '').split('\n').length);
      const seed: Mode = p.mode === 'math' ? 'math' : 'text';
      p.lineModes = Array.from({ length: n }, () => seed);
    }
  }
  return s;
}

// iOS WKWebView's IndexedDB can wedge (open/get never settles) on the second
// page loaded under the app's custom scheme — settings.html, or index.html
// after navigating back. An unbounded await there hangs startup forever even
// though the synchronous localStorage mirror already holds the data.
function withTimeout<T>(p: Promise<T>, ms: number): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(`IndexedDB read timed out after ${ms}ms`)), ms);
    p.then(
      (v) => { clearTimeout(t); resolve(v); },
      (e) => { clearTimeout(t); reject(e); }
    );
  });
}

// Returns the stored Settings, or null ONLY when the store is genuinely empty.
// Throws on a real read error (so the caller can avoid treating it as first-run).
async function readStored(): Promise<Settings | null> {
  const localEnv = readLocalEnvelope();
  let idbEnv: Envelope | undefined;
  try {
    // The mirror is written synchronously BEFORE every debounced IndexedDB
    // write and refreshed after every successful read, so it is never older
    // than the IndexedDB copy — when it exists, a hung IndexedDB read may be
    // abandoned in its favor. With no mirror (true first run) wait fully.
    idbEnv = localEnv
      ? await withTimeout(idbGet<Envelope>(KEY), 1200)
      : await idbGet<Envelope>(KEY);
  } catch (err) {
    if (!localEnv) throw err;
    console.warn('[math-popup] IndexedDB read failed or hung; recovered settings from local mirror', err);
  }
  const env = localEnv && (!idbEnv || (localEnv.savedAt ?? 0) > (idbEnv.savedAt ?? 0))
    ? localEnv
    : idbEnv;
  if (!env) return null;
  if (env.schemaVersion > CURRENT_SCHEMA) return env.settings; // newer build wrote it; read as-is, don't migrate
  let s = env.settings;
  for (let v = env.schemaVersion; v < CURRENT_SCHEMA; v++) s = migrations[v + 1]?.(s) ?? s;
  return s;
}

export function getSettings(): Promise<Settings> {
  if (cache) return Promise.resolve(cache);
  if (loadPromise) return loadPromise;
  loadPromise = (async () => {
    let stored: Settings | null = null;
    let readFailed = false;
    try {
      stored = await readStored();        // null => genuinely empty
      storageReadable = true;
    } catch (err) {
      // A transient read failure must NOT be mistaken for first-run (that path would
      // overwrite the user's real data with defaults). Render with defaults, persist
      // nothing, and allow a later getSettings() to retry the read.
      readFailed = true;
      storageReadable = false;
      console.error('[math-popup] settings read failed; using defaults without overwriting stored data', err);
    }
    const seeded: Settings = { ...DEFAULT_SETTINGS, ...(stored ?? {}) };
    if (!Array.isArray(seeded.suffixes) || seeded.suffixes.length === 0) {
      seeded.suffixes = DEFAULT_SETTINGS.suffixes;
    }
    const result = normalize(seeded);
    if (!readFailed) {
      cache = result;
      writeLocalEnvelope(envelope(result));
      // Genuine first run only — seed via the write chain so it's ordered with later writes.
      if (stored === null) writeChain = writeChain.then(() => idbSet(KEY, envelope(result))).catch(() => {});
    } else {
      loadPromise = null;                 // don't pin a defaults snapshot from a transient failure
    }
    return result;
  })();
  return loadPromise;
}

export async function setSettings(partial: Partial<Settings>): Promise<Settings> {
  // Always merge onto a fully-valid base (never a bare DEFAULT_SETTINGS missing
  // pages/activePageId) — await the initial load if it hasn't resolved yet.
  const base: Settings = cache ?? await getSettings();
  const merged: Settings = { ...base, ...partial }; // shallow merge — matches main/store.ts (pages replaced wholesale)
  cache = merged;
  writeLocalEnvelope(envelope(merged));
  // Don't write while storage is known-unreadable; we'd risk clobbering data we
  // couldn't read. The change stays in memory; persistence resumes on a good read.
  if (!storageReadable) return merged;
  pending = merged;
  if (saveTimer !== null) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    saveTimer = null;
    const snap = pending!;
    pending = null;
    writeChain = writeChain
      .then(() => idbSet(KEY, envelope(snap)))
      .then(() => { bus?.postMessage({ type: 'settings', value: snap }); })
      .catch((err) => console.error('[math-popup] persist failed', err));
  }, 150);
  return merged;
}

// Force-write the latest state immediately (called on pagehide / tab-hidden so a
// backgrounded PWA doesn't lose recent keystrokes before iOS kills it). Flushes the
// authoritative cache even if no debounce was pending (e.g. a prior write rejected).
export async function flushSettings(): Promise<void> {
  if (saveTimer !== null) { clearTimeout(saveTimer); saveTimer = null; }
  const snap = pending ?? cache;
  pending = null;
  if (snap && storageReadable) {
    writeChain = writeChain.then(() => idbSet(KEY, envelope(snap))).catch(() => {});
  }
  await writeChain;
}

export function setOnSettingsChanged(cb: ((s: Settings) => void) | null): void {
  changedCb = cb;
}

// Cross-tab live sync: another tab's committed write arrives here (a tab never
// receives its own BroadcastChannel message, so this never echoes same-tab saves).
// Ignore the remote snapshot while this tab has un-flushed local edits, so we don't
// clobber unsaved work with a stale wholesale copy.
bus?.addEventListener('message', (e: MessageEvent) => {
  if (e.data?.type !== 'settings') return;
  if (pending !== null || saveTimer !== null) return;
  cache = e.data.value as Settings;
  changedCb?.(cache);
});
