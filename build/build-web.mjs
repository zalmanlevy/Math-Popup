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
`;
await writeFile(join(out, 'web.css'), WEB_CSS, 'utf8');

// ---- 4. generate the three HTML pages from the desktop HTML (string transform) ----
const PWA_HEAD = `  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover" />
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
  // Load the shim BEFORE the app bundle so window.mathPopup exists first.
  html = html.replace(`<script src="${appScript}"></script>`,
    `<script src="shim.js"></script>\n  <script src="${appScript}"></script>`);
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
  '/shim.js', '/popup.js', '/settings.js', '/help.js',
  '/popup.css', '/settings.css', '/help.css', '/web.css',
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
