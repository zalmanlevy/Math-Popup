// Web / PWA build for Math Popup. Produces a static site in web-dist/ that runs the
// UNCHANGED desktop renderer in the browser, backed by the browser shim (src/web).
// Completely separate from the Electron build (build/build.mjs) — it reuses the same
// renderer entry files but never touches dist/, release/, or the desktop scripts.
//
//   npm run build:web   ->   node build/build-web.mjs   ->   web-dist/
import { build } from 'esbuild';
import { mkdir, copyFile, writeFile, readFile, rm, readdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const out = join(root, 'web-dist');
const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
const version = pkg.version;
// Unique per deploy so the service worker cache busts on EVERY deploy, not only on
// version bumps (web-only changes keep the same package version). Vercel sets the
// commit SHA; locally we fall back to a build timestamp.
const buildId = (process.env.VERCEL_GIT_COMMIT_SHA || String(Date.now())).slice(0, 12);

const common = {
  bundle: true,
  platform: 'browser',
  format: 'iife',
  target: ['safari16', 'chrome120'],
  minify: true,
  sourcemap: true,
  define: { __APP_VERSION__: JSON.stringify(version) },
  logLevel: 'info',
};

// ---- clean ----
await rm(out, { recursive: true, force: true });
await mkdir(join(out, 'icons'), { recursive: true });

// ---- 1. bundle the shim + the three renderer entries (unchanged renderer code) ----
await Promise.all([
  build({ ...common, entryPoints: [join(root, 'src/web/shim.ts')],         outfile: join(out, 'shim.js') }),
  build({ ...common, entryPoints: [join(root, 'src/renderer/popup.ts')],    outfile: join(out, 'popup.js') }),
  build({ ...common, entryPoints: [join(root, 'src/renderer/settings.ts')], outfile: join(out, 'settings.js') }),
  build({ ...common, entryPoints: [join(root, 'src/renderer/help.ts')],     outfile: join(out, 'help.js') }),
]);

// ---- 2. copy the renderer CSS verbatim (desktop CSS is untouched) ----
for (const f of ['popup.css', 'settings.css', 'help.css']) {
  await copyFile(join(root, 'src/renderer', f), join(out, f));
}

// ---- 3. web-only CSS tail (mobile shell; loaded AFTER the app CSS) ----
const WEB_CSS = `/* Web/PWA-only adjustments — desktop CSS files are untouched. */
:root { color-scheme: light dark; }
* { -webkit-tap-highlight-color: transparent; }
body { min-height: 100svh; }
@supports (height: 100dvh) { body { min-height: 100dvh; } }
button, a, .icon-btn, .tab-chip, .seg-btn, .btn, .tab-bar-add, .tab-overflow-btn { touch-action: manipulation; }
.editor { overscroll-behavior: contain; }
.status-bar { padding-bottom: env(safe-area-inset-bottom); }
/* Never leave the actual editor text invisible if startup fails before the
   syntax overlay is initialized. The inline web bootstrap guard adds this
   class on an exception or a startup timeout. */
:root.web-startup-failed .editor,
:root.web-startup-failed .editor .ed-line { color: var(--text) !important; }
:root.web-startup-failed .syntax-overlay { display: none; }
/* Window-chrome controls that have no meaning in a browser/PWA (no tray to hide to,
   no always-on-top). Their shim methods are no-ops, so hide them rather than show
   dead buttons. */
#close-window, #toggle-pin { display: none; }
/* In-app Back control for the settings/help pages (standalone PWAs have no browser back button). */
.web-back {
  position: fixed; top: calc(env(safe-area-inset-top) + 8px); left: 10px; z-index: 100;
  width: 34px; height: 34px; display: inline-flex; align-items: center; justify-content: center;
  border-radius: 8px; font-size: 20px; line-height: 1; text-decoration: none; color: inherit;
  background: rgba(127, 127, 127, 0.16);
}
.web-back:active { background: rgba(127, 127, 127, 0.30); }
/* Keep the settings/help heading clear of the fixed Back control. */
body:has(> .web-back) header { padding-left: 52px; }

/* ---- Native app shell (Capacitor iOS/iPadOS only; html gets .cap-native) ---- */
/* Flatten the desktop floating-panel chrome: a border+radius+shadow around the
   whole screen reads as a website inside a native window. The native webview is
   edge-to-edge (viewport-fit=cover + contentInset 'never'), so CSS owns every
   safe-area inset here. */
:root.cap-native body {
  border: none;
  border-radius: 0;
  box-shadow: none;
  padding-left: env(safe-area-inset-left);
  padding-right: env(safe-area-inset-right);
}
/* Notes page: the title bar extends up behind the status bar. */
:root.cap-native .title-bar {
  border-radius: 0;
  height: calc(34px + env(safe-area-inset-top));
  padding-top: env(safe-area-inset-top);
}
:root.cap-native .drag-handle { display: none; }  /* desktop window-drag affordance */
/* Settings/help pages (no .title-bar): pad the page itself clear of the status
   bar and the home indicator. */
:root.cap-native body:not(:has(> .title-bar)) {
  padding-top: env(safe-area-inset-top);
  padding-bottom: env(safe-area-inset-bottom);
}
/* The footer is height:24px with border-box sizing, so the safe-area bottom
   padding (home indicator) would squeeze its text instead of extending it —
   grow the bar by the inset. */
:root.cap-native .status-bar {
  height: calc(24px + env(safe-area-inset-bottom));
}

/* iPhone only in practice: WKWebView "text autosizing" inflates font sizes on
   phone-width screens as if this were an unoptimized desktop page, pushing the
   title bar's right-side buttons off screen. Render at designed sizes. */
:root.cap-native {
  -webkit-text-size-adjust: 100%;
  text-size-adjust: 100%;
}

/* ---- Native: app chrome is UI, not selectable text ---- */
/* iOS long-press starts the text-selection loupe on ANY selectable text, which
   fights the synthesized context-menu long-press on tabs / the new-tab button
   (half the time it selected the tab's words instead of opening its menu).
   The notes page outside the editor is all chrome — kill selection there;
   the editor and real inputs stay selectable. (-webkit-user-select inherits.) */
:root.cap-native body:has(> .title-bar) {
  -webkit-user-select: none;
  user-select: none;
  -webkit-touch-callout: none;
}
:root.cap-native .editor,
:root.cap-native input,
:root.cap-native textarea {
  -webkit-user-select: text;
  user-select: text;
  -webkit-touch-callout: default;
}

/* ---- Native editor scrolling: one scroller, all layers move as one ---- */
/* Desktop scrolls the contenteditable and JS mirrors its scrollTop onto the
   overlay/gutter layers per scroll event. On iOS the finger-driven scroll runs
   on the compositor while that mirroring runs on the main thread, so the
   visible text lags the finger and momentum feels off. Natively the SHELL is
   the one scroller: the editor becomes an in-flow, full-content-height child
   (it sizes .editor-stack; the absolutely-positioned layers stretch with it),
   the gutter column stretches alongside, and scrolling is a single composited
   move — native momentum, rubber-band, scrollbar and caret-follow included.
   popup.js routes every scroll read/write through the shell (scrollHost). */
:root.cap-native .editor-shell {
  overflow-y: auto;
  overscroll-behavior: contain;
  -webkit-overflow-scrolling: touch;
}
:root.cap-native .editor-stack {
  overflow: visible;
  /* Size to the note, not the shell viewport: the stack (and the stretched
     gutter column beside it) must span the full scrollable content so their
     backgrounds and absolutely-positioned layers cover every scrolled line. */
  height: max-content;
  min-height: 100%;
  display: flex;
  flex-direction: column;
}
:root.cap-native .editor {
  position: relative;
  inset: auto;
  height: auto;
  flex: 1 0 auto;
  overflow: visible;
  /* Overscroll room past the last line — the note can always be scrolled a few
     extra lines up, so the line being written never hugs the keyboard/footer. */
  padding-bottom: clamp(120px, 25vh, 260px);
}
:root.cap-native .syntax-overlay,
:root.cap-native .result-overlay,
:root.cap-native .find-layer {
  overflow: visible;
}
:root.cap-native .line-gutter { overflow: visible; }

/* ---- Native settings/help pages: the page body scrolls itself ---- */
/* The webview's own scrolling is disabled app-wide (that's what pins the notes
   header), which also froze these document-scrolling pages — the help page
   could not scroll at all. Make body an internal scroller. html must be
   overflow:hidden or the body's overflow would propagate to the (disabled)
   viewport scroller instead of scrolling the body element itself. */
:root.cap-native { overflow: hidden; }
:root.cap-native body:not(:has(> .title-bar)) {
  height: 100dvh;
  overflow-y: auto;
  overscroll-behavior: contain;
  -webkit-overflow-scrolling: touch;
}

/* ---- Native help page: sidebar TOC becomes a slide-in Contents drawer ---- */
:root.cap-native body:has(> nav.toc) { display: block; }  /* drop the 220px grid column */
:root.cap-native nav.toc {
  position: fixed;
  top: 0;
  left: 0;
  bottom: 0;
  width: min(300px, 82vw);
  max-height: none;
  padding-top: calc(env(safe-area-inset-top) + 56px);
  padding-bottom: calc(env(safe-area-inset-bottom) + 16px);
  background: var(--bg-elev);
  border-right: 1px solid var(--border);
  box-shadow: 0 0 28px rgba(0, 0, 0, 0.28);
  z-index: 95;
  transform: translateX(-105%);
  transition: transform 200ms ease;
  overflow-y: auto;
  -webkit-overflow-scrolling: touch;
}
:root.cap-native.toc-open nav.toc { transform: translateX(0); }
:root.cap-native .toc-scrim { display: none; }
:root.cap-native.toc-open .toc-scrim {
  display: block;
  position: fixed;
  inset: 0;
  z-index: 94;
  background: rgba(0, 0, 0, 0.32);
}
/* The drawer toggle sits beside the Back control (web-back is at left 10px). */
:root.cap-native .toc-toggle {
  position: fixed;
  top: calc(env(safe-area-inset-top) + 8px);
  left: 54px;
  z-index: 100;
  width: 34px;
  height: 34px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 0;
  border-radius: 8px;
  font-size: 17px;
  line-height: 1;
  color: inherit;
  background: rgba(127, 127, 127, 0.16);
}
:root.cap-native .toc-toggle:active { background: rgba(127, 127, 127, 0.30); }
/* Windowed mode shifts the Back control right of the window controls; follow it. */
:root.cap-native.cap-windowed .toc-toggle { left: 120px; }
/* Room for both fixed controls above the heading. */
:root.cap-native body:has(> nav.toc) header { padding-left: 96px; }
:root.cap-native.cap-windowed body:has(> nav.toc) header { padding-left: 162px; }

/* ---- iPadOS windowed mode (resizable windows / Stage Manager) ---- */
/* .cap-windowed is toggled by the native WindowState plugin. Per Apple's
   iPadOS 26 guidance, toolbar content shares the top row with the
   traffic-light window controls and shifts right past them (the controls
   occupy the top-leading ~64px; safe-area reporting for them is unreliable,
   so the clearance is a constant). The forced height/padding-top also
   overrides the fullscreen status-bar inset rule — there is no status bar
   inside a window. */
:root.cap-native.cap-windowed .title-bar {
  height: 38px;
  padding-top: 4px;
  padding-left: max(env(safe-area-inset-left), 72px);
  padding-right: 14px;  /* keep ⚙ off the rounded top-right corner */
}
/* Settings/help pages: their fixed Back control lives exactly where the
   window controls are — shift it (and the heading) right of them. */
:root.cap-native.cap-windowed .web-back { left: 76px; }
:root.cap-native.cap-windowed body:has(> .web-back) header { padding-left: 122px; }
/* Settings that only mean something on the Windows desktop app (startup/tray,
   taskbar pinning, Obsidian file sync, in-app updater — iOS updates ship via
   the App Store/TestFlight). Hidden in the native app only. */
:root.cap-native label.row:has(#launch-at-startup),
:root.cap-native label.row:has(#taskbar-icon),
:root.cap-native label.row:has(#advanced-mode),
:root.cap-native #update-banner,
:root.cap-native div.row:has(> #check-updates) { display: none; }
/* Rounded window corners clip flush content and the resize grip overlays the
   bottom-right — give the footer breathing room on all three sides. */
:root.cap-native.cap-windowed .status-bar {
  height: calc(24px + max(env(safe-area-inset-bottom), 8px));
  padding-bottom: max(env(safe-area-inset-bottom), 8px);
  padding-left: 16px;
  padding-right: 34px;
}

/* ---- Native on-screen-keyboard bar (math keys + 123/ABC toggle) ---- */
/* Built by the popup renderer only inside the native app; shown only while
   the on-screen keyboard is up (.kb-open tracks the native keyboard events,
   so a hardware keyboard never shows it). The webview resizes with the
   keyboard, so this bottom bar sits directly on top of it. */
.kb-bar { display: none; }
:root.cap-native.kb-open .kb-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 8px;
  background: var(--bg-elev);
  border-top: 1px solid var(--border);
  overflow-x: auto;
}
.kb-bar button { touch-action: manipulation; -webkit-tap-highlight-color: transparent; }
.kb-toggle {
  display: inline-flex;
  flex: 0 0 auto;
  background: var(--bg-soft);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 2px;
}
.kb-seg {
  border: 0;
  background: transparent;
  padding: 5px 10px;
  font-size: 12px;
  font-weight: 600;
  color: var(--text-soft);
  border-radius: 6px;
}
.kb-seg.active {
  background: var(--bg-elev);
  color: var(--text);
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.18);
}
.kb-keys { display: flex; gap: 5px; flex: 1; }
.kb-key {
  flex: 1 0 30px;
  min-width: 30px;
  height: 34px;
  border: 1px solid var(--border);
  border-radius: 7px;
  background: var(--bg-elev);
  color: var(--text);
  font-size: 15px;
  font-family: var(--font-mono);
}
.kb-key:active { background: var(--bg-soft); }
/* ABC-mode markdown keys: styled like what they produce. */
.kb-key.kb-b { font-family: var(--font-ui); font-weight: 800; }
.kb-key.kb-i { font-family: var(--font-ui); font-style: italic; }
.kb-key.kb-u { font-family: var(--font-ui); text-decoration: underline; text-underline-offset: 2px; }
/* Invisible 1px input that briefly holds focus during a 123/ABC switch so the
   iOS keyboard morphs in place instead of dismissing and re-presenting.
   opacity:0 (not display:none / visibility:hidden) — iOS refuses to focus
   fully hidden fields. */
.kb-focus-keeper {
  position: fixed;
  bottom: 0;
  left: 0;
  width: 1px;
  height: 1px;
  padding: 0;
  border: 0;
  margin: 0;
  opacity: 0;
  background: transparent;
  color: transparent;
  caret-color: transparent;
  pointer-events: none;
}
.kb-left { display: flex; align-items: center; gap: 8px; flex: 0 0 auto; }
.kb-key-space {
  flex: 0 0 56px;
  min-width: 56px;
  display: inline-flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 3px;
  padding: 3px 0 5px;
}
.kb-space-word {
  font-size: 9px;
  line-height: 1;
  font-family: var(--font-ui);
  color: var(--text-soft);
  letter-spacing: 0.05em;
}
/* Drawn open-box space glyph (⎵) — the Unicode character renders miniature,
   so draw it: full-size, as wide as the key allows. */
.kb-space-sym {
  display: block;
  width: 62%;
  min-width: 24px;
  max-width: 46px;
  height: 9px;
  border: 2px solid var(--text);
  border-top: 0;
  border-radius: 0 0 4px 4px;
}
/* Narrow (iPhone portrait, skinny iPad windows): two fixed rows of six keys —
   every key visible and in a stable position, no scrolling. The left cluster
   stacks: 123/ABC on the first row, an equally wide space key on the second. */
@media (max-width: 519px) {
  :root.cap-native.kb-open .kb-bar { overflow-x: visible; align-items: stretch; }
  .kb-left { flex-direction: column; align-items: stretch; gap: 5px; }
  .kb-toggle { flex: 1; }
  .kb-toggle .kb-seg { flex: 1; }
  .kb-key-space { flex: 1; min-width: 0; height: auto; }
  .kb-keys { display: grid; grid-template-columns: repeat(6, 1fr); gap: 5px; }
  .kb-key { min-width: 0; }
}
`;
await writeFile(join(out, 'web.css'), WEB_CSS, 'utf8');

// ---- 4. generate the three HTML pages from the desktop HTML (string transform) ----
// maximum-scale=1 disables iOS's automatic page-zoom when focusing a text
// field with a sub-16px font (the editor is 14px) — the app has its own zoom.
const PWA_HEAD = `  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover" />
  <link rel="manifest" href="/manifest.webmanifest" />
  <meta name="theme-color" media="(prefers-color-scheme: light)" content="#fafafa" />
  <meta name="theme-color" media="(prefers-color-scheme: dark)" content="#0f1115" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <meta name="mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-status-bar-style" content="default" />
  <meta name="apple-mobile-web-app-title" content="Math Popup" />
  <link rel="apple-touch-icon" href="/icons/apple-touch-icon-180.png" />
  <link rel="stylesheet" href="/web.css" />
`;

function transformHtml(html, appScript, withBack) {
  // Drop any existing viewport meta (popup has one, settings/help don't), then add
  // our PWA head block (which includes the cover viewport).
  html = html.replace(/[ \t]*<meta name="viewport"[^>]*>\s*\n?/i, '');
  html = html.replace('</head>', PWA_HEAD + '</head>');
  // Fingerprint every executable/style URL. An older controlling iOS service
  // worker may still be cache-first; the unique query guarantees it cannot
  // satisfy a new deploy with a stale bundle under the old URL.
  html = html.replace(/href="([^"]+\.css)"/g, (_m, href) => `href="${href}?v=${buildId}"`);
  // Load the shim BEFORE the app bundle so window.mathPopup exists first.
  html = html.replace(`<script src="${appScript}"></script>`,
    `<script>\n` +
    `    (() => {\n` +
    `      let done = false;\n` +
    `      const fail = (reason) => {\n` +
    `        if (done) return;\n` +
    `        document.documentElement.classList.add('web-startup-failed');\n` +
    `        const status = document.getElementById('status-msg');\n` +
    `        if (status) status.textContent = 'Startup issue: ' + reason;\n` +
    `      };\n` +
    `      const timer = setTimeout(() => fail('loading timed out'), 5000);\n` +
    `      window.mathPopupWebReady = () => { done = true; clearTimeout(timer); };\n` +
    `      window.addEventListener('error', (e) => fail(e.message || 'script error'));\n` +
    `      window.addEventListener('unhandledrejection', (e) => fail(e.reason?.message || String(e.reason || 'startup failed')));\n` +
    `    })();\n` +
    `  </script>\n` +
    `  <script src="shim.js?v=${buildId}"></script>\n` +
    `  <script src="${appScript}?v=${buildId}"></script>`);
  if (withBack) {
    html = html.replace(/<body([^>]*)>/i,
      `<body$1>\n  <a class="web-back" href="index.html" aria-label="Back to notes">←</a>`);
  }
  return html;
}

const popupHtml = await readFile(join(root, 'src/renderer/popup.html'), 'utf8');
const settingsHtml = await readFile(join(root, 'src/renderer/settings.html'), 'utf8');
const helpHtml = await readFile(join(root, 'src/renderer/help.html'), 'utf8');
await writeFile(join(out, 'index.html'),    transformHtml(popupHtml, 'popup.js', false), 'utf8');
await writeFile(join(out, 'settings.html'), transformHtml(settingsHtml, 'settings.js', true), 'utf8');
await writeFile(join(out, 'help.html'),     transformHtml(helpHtml, 'help.js', true), 'utf8');

// ---- 5. PWA manifest ----
const manifest = {
  name: 'Math Popup',
  short_name: 'Math Popup',
  description: pkg.description || 'A notepad that does the math while you type.',
  start_url: '/',
  scope: '/',
  display: 'standalone',
  orientation: 'any',
  background_color: '#fafafa',
  theme_color: '#fafafa',
  icons: [
    { src: '/icons/icon-192.png', sizes: '192x192', type: 'image/png', purpose: 'any' },
    { src: '/icons/icon-512.png', sizes: '512x512', type: 'image/png', purpose: 'any' },
    { src: '/icons/maskable-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
  ],
};
await writeFile(join(out, 'manifest.webmanifest'), JSON.stringify(manifest, null, 2), 'utf8');

// ---- 6. service worker (cache-first, version-stamped so each deploy busts the cache) ----
const SW = `// Auto-generated by build/build-web.mjs — do not edit.
const CACHE = 'mathpopup-${version}-${buildId}';
const ASSETS = [
  '/', '/index.html', '/settings.html', '/help.html',
  '/shim.js?v=${buildId}', '/popup.js?v=${buildId}', '/settings.js?v=${buildId}', '/help.js?v=${buildId}',
  '/popup.css?v=${buildId}', '/settings.css?v=${buildId}', '/help.css?v=${buildId}', '/web.css?v=${buildId}',
  '/manifest.webmanifest',
  '/icons/icon-192.png', '/icons/icon-512.png',
  '/icons/maskable-512.png', '/icons/apple-touch-icon-180.png',
];
self.addEventListener('install', (e) => {
  // Resilient precache: one missing asset must not abort the whole cache.
  e.waitUntil(caches.open(CACHE).then((c) => Promise.allSettled(ASSETS.map((u) => c.add(u)))));
  self.skipWaiting();
});
self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  // Page navigations AND executable shell assets are network-first. Keeping JS/CSS
  // cache-first allowed an activating iOS service worker to pair a new popup.js
  // with an old shim.js, crashing startup. Offline still falls back to the atomic
  // cache populated during install.
  const url = new URL(e.request.url);
  const shellAsset = url.origin === self.location.origin &&
    (url.pathname.endsWith('.js') || url.pathname.endsWith('.css'));
  if (e.request.mode === 'navigate' || shellAsset) {
    e.respondWith(
      fetch(e.request).catch(() =>
        caches.match(e.request).then((r) => r ||
          (e.request.mode === 'navigate' ? caches.match('/index.html') : Response.error()))
      )
    );
    return;
  }
  // Static assets: cache-first.
  e.respondWith(caches.match(e.request).then((r) => r || fetch(e.request)));
});
`;
await writeFile(join(out, 'sw.js'), SW, 'utf8');

// ---- 7. copy committed icons (generated by build/gen-icons.mjs) ----
const iconSrc = join(root, 'src/web/icons');
for (const f of await readdir(iconSrc)) {
  await copyFile(join(iconSrc, f), join(out, 'icons', f));
}

console.log(`[build:web] done -> ${out} (v${version})`);
