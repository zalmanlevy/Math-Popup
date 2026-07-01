// Browser implementation of the `window.mathPopup` bridge that the Electron preload
// normally injects. Loaded BEFORE the renderer bundle on every web page, so the
// unchanged desktop renderer runs as a local-only web app / PWA. Desktop-only
// capabilities (tray, always-on-top, auto-update, Obsidian file sync) degrade to
// safe no-ops; everything else is backed by real browser APIs.
import { getSettings, setSettings, flushSettings, setOnSettingsChanged } from './storage';
import type { Settings } from '../shared/types';

type Resolved = 'light' | 'dark';

// ---- theme: OS scheme changes drive onThemeChanged (renderer re-renders syntax) ----
const themeMq = window.matchMedia('(prefers-color-scheme: dark)');
const themeCbs = new Set<(r: Resolved) => void>();
themeMq.addEventListener('change', () => {
  const r: Resolved = themeMq.matches ? 'dark' : 'light';
  themeCbs.forEach((cb) => cb(r));
});

// ---- settings change fan-out (fed by cross-tab writes; see storage.ts) ----
const settingsCbs = new Set<(s: Settings) => void>();
setOnSettingsChanged((s) => settingsCbs.forEach((cb) => cb(s)));

// ---- update channel: web is always the latest deploy, so this is inert ----
const updateCbs = new Set<(state: any) => void>();

// Navigate away only AFTER the pending write is committed. iOS freezes the page
// during unload, so the async IndexedDB write in the pagehide handler frequently
// doesn't land — which is why settings toggled right before tapping Back weren't
// saved. Awaiting the flush inside the click/gesture (page still foregrounded) is
// reliable; a short timeout guards against a hung write so navigation never stalls.
function flushThenNavigate(href: string): void {
  const guard = new Promise<void>((resolve) => setTimeout(resolve, 500));
  void Promise.race([flushSettings(), guard]).finally(() => { location.assign(href); });
}

function clipboardFallback(text: string): void {
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    ta.remove();
  } catch { /* ignore */ }
}

window.mathPopup = {
  getSettings,
  setSettings,
  hidePopup: async () => { /* no tray/window on web */ },
  setAlwaysOnTop: async () => { /* no equivalent on web */ },
  openSettings: async () => { flushThenNavigate('settings.html'); },
  openHelp: async () => { flushThenNavigate('help.html'); },
  copyText: (text: string) => {
    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(text).catch(() => clipboardFallback(text));
    } else {
      clipboardFallback(text);
    }
  },
  setZoomFactor: (factor: number) => {
    (document.documentElement.style as unknown as { zoom: string }).zoom = String(factor);
  },
  onThemeChanged: (cb) => {
    themeCbs.add(cb);
    return () => { themeCbs.delete(cb); };
  },
  onSettingsChanged: (cb) => {
    settingsCbs.add(cb);
    return () => { settingsCbs.delete(cb); };
  },
  getAppVersion: async () => __APP_VERSION__,
  getUpdateState: async () => ({ phase: 'idle' as const }),
  checkForUpdates: async () => { updateCbs.forEach((cb) => cb({ phase: 'not-available' })); },
  installUpdate: async () => { /* unreachable on web */ },
  onUpdateState: (cb) => {
    updateCbs.add(cb);
    return () => { updateCbs.delete(cb); };
  },
  // Obsidian linking / live Markdown sync needs a filesystem — unavailable in the
  // browser (and absent on iOS Safari). Advanced Mode is off by default and every
  // call site degrades gracefully on these stubs.
  chooseObsidianNote: async () => null,
  readObsidianNote: async () => { throw new Error('Obsidian linking is unavailable in the web version'); },
  writeObsidianNote: async () => { throw new Error('Obsidian linking is unavailable in the web version'); },
  watchObsidianNotes: async (paths: string[]) => paths,
  onObsidianFileChanged: () => () => {},
};

// ---- PWA plumbing ----

// Register the offline/service worker (root scope controls all three pages).
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').catch(() => { /* offline still works from cache */ });
  });
}

// Ask the browser to keep our IndexedDB data (best defense against iOS evicting an
// installed PWA's storage). Harmless if unsupported or already persisted.
if (navigator.storage?.persist) {
  navigator.storage.persisted().then((p) => { if (!p) navigator.storage.persist(); }).catch(() => {});
}

// Flush the debounced write before the page is hidden/closed (beforeunload is
// unreliable on iOS; pagehide + visibilitychange=hidden are the dependable hooks).
const flush = () => { void flushSettings(); };
window.addEventListener('pagehide', flush);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });

// The in-app Back link (settings/help → notes) navigates via a plain <a>; route it
// through flushThenNavigate so a setting toggled just before tapping Back is committed
// first (the pagehide flush alone is unreliable on iOS). Capture phase so we intercept
// before the browser follows the href.
document.addEventListener('click', (e) => {
  const back = (e.target as HTMLElement | null)?.closest?.('a.web-back') as HTMLAnchorElement | null;
  if (!back || e.defaultPrevented || e.button !== 0 || e.metaKey || e.ctrlKey) return;
  e.preventDefault();
  flushThenNavigate(back.getAttribute('href') || 'index.html');
}, true);

// Keep the focused editor line above the on-screen keyboard on iOS.
if (window.visualViewport) {
  window.visualViewport.addEventListener('resize', () => {
    const el = document.activeElement as HTMLElement | null;
    if (el && el.isContentEditable) el.scrollIntoView({ block: 'nearest' });
  });
}
