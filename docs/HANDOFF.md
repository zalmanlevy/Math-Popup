# Math Popup — Developer Handoff

Orientation for anyone (human or agent) picking up this project mid-stream.
For the TestFlight/App Store release mechanics see **[docs/IOS_RELEASE.md](./IOS_RELEASE.md)**;
this doc covers the architecture, the rules, the non-obvious gotchas, and where
things currently stand.

## One app, three targets

The product is fundamentally a **web app**. The three "versions" are the same
web UI wrapped in different shells:

| Target | Shell | Build | Distribution |
| --- | --- | --- | --- |
| **Windows** | Electron (`src/main`, `dist/`) | `npm run build` → `npm run dist` (electron-builder) | GitHub Release + electron-updater auto-update (manual) |
| **Web / PWA** | none (static site) | `npm run build:web` → `web-dist/` | Vercel (`vercel.json`) |
| **iOS (iPhone/iPad)** | Capacitor 8 (`ios/`, `capacitor.config.ts`) | `web-dist/` synced via `cap sync ios` | TestFlight via GitHub Actions — **no Mac needed** |

- Shared core: `src/renderer/` (UI — `popup.ts`, `help.ts`) and `src/web/`
  (storage, native glue). Built by `build/build-web.mjs` (web/iOS) and
  `build/build.mjs` (Electron).
- iOS app name is **"Math Pad X"**; Windows/web keep the **"Math Popup"** brand.
  Bundle id `com.zalmanlevy.mathpopup` is permanent and internal-only.

## The golden rule: do not change the Windows/desktop app

Every mobile change must leave Electron **byte-identical**. This is enforced by
gating, not by discipline alone:

- **CSS** — native/mobile styles live in `WEB_CSS` inside `build/build-web.mjs`
  (emitted to `web-dist/web.css`, used by web + iOS **only**). Electron uses
  `src/renderer/popup.css`, which mobile work never touches. Native-only rules
  are scoped under `:root.cap-native`; iPad-only under `:root.cap-native.cap-ipad`.
- **JS** — behavior branches are guarded by:
  - `IS_CAP_NATIVE` — `documentElement` has class `cap-native` (added by
    `src/web/native.ts` `initNativeShell()` when running inside Capacitor).
  - `IS_WEB_SHELL` — web **and** iOS, but not Electron.
  - `isNativeApp` (`src/web/native.ts`) — Capacitor platform check.
- The on-screen **operator bar** and all its keys exist only on native
  (`initMathKeyboardBar()` runs only when `IS_CAP_NATIVE`).

When a fix must touch genuinely shared logic, gate it so desktop is unaffected,
and say so in the commit. (Two shared fixes so far — `/clear` and the `/`+space
bug — are already desktop-safe by construction; see below.)

## Build & release

- **iOS:** commit → push `main` → Actions → **iOS TestFlight**
  (`workflow_dispatch`, ref `main`). Build number = `run_number × 100 +
  run_attempt`. Runbook + one-time Apple setup: **docs/IOS_RELEASE.md**.
  - The working branch and `main` are kept in sync (fast-forward); CI dispatches
    on `main`.
- **Windows:** manual and currently **deferred**. `npm run dist` →
  electron-builder → publish a GitHub Release → electron-updater rolls it out.
  No CI builds Windows. Shared bug fixes already on `main` ride along the next
  manual Windows release automatically — nothing to backport.
- **Always green before shipping:** `npx tsc --noEmit`, `npm run build:web`,
  `npm run build`.

## Key subsystems & non-obvious gotchas

**Editor.** `#editor` is a `contenteditable` with per-line `.ed-line` blocks that
expose a textarea-shaped API (`editor.value`, `selectionStart/End`,
`setSelectionRange`, via `Object.defineProperties`). The editor's own text is
transparent; the visible text is painted by `.syntax-overlay`, results by
`.result-overlay`, gutters kept in sync in JS. Assigning `editor.value` (with
`\n`) then `render()` rebuilds the line blocks.

**Scrolling (`scrollHost`).** Desktop/web: the editor element scrolls and
`syncScroll()` mirrors the overlay/gutter layers. Native iOS: the
`.editor-shell` scrolls **all** layers as one composited block (kills the
finger-lag from main-thread mirroring); `syncScroll()` early-returns on native.
`.editor` carries an overscroll `padding-bottom` so the last line never hugs the
keyboard.

**Persistence (`src/web/storage.ts`).** iOS WKWebView IndexedDB wedges on the
2nd page load under `capacitor://`, so a synchronous **localStorage mirror** is
authoritative on native and IDB is skipped when the mirror exists (this is what
makes settings/back-nav instant instead of waiting on a ~1.2s IDB timeout).
Lightweight renderer UI prefs use localStorage directly (e.g. operator-bar
collapse, key `mp-kb-collapsed`).

**iOS keyboard/operator bar** (native only, `popup.ts`):
- Built by `initMathKeyboardBar()`. Two key sets: `KB_MATH_KEYS` (123 mode) and
  `KB_TEXT_KEYS` (ABC mode). Left cluster = the 123/ABC toggle over a split
  `space | return` (both modes). `renderKbKeys()` rebuilds `.kb-keys` per mode.
- Visibility is driven by `.kb-open` on `documentElement` (added on editor focus
  and on native `keyboardWillShow`; a hardware keyboard never fires those).
- **Mode switch** (`setKeyboardMode`): sets `inputmode`/`autocorrect`/
  `autocapitalize`, then hops focus through a hidden 1px **focus-keeper** input
  so iOS keeps the keyboard **presented** and morphs it in place (one webview
  resize instead of dismiss + re-present).
- **FLIP slide** (`armKbFlip` / `onKbFlipResize`): the bar rides the webview's
  bottom under `resize:'native'`, and a switch resizes the webview twice (the
  keyboard morph, then iOS's QuickType predictive bar a beat later), which would
  jump the bar. Each resize is turned into a slide (invert the bar's transform
  back to where it was, animate to zero). This replaced an earlier opacity fade
  that left a visible blank gap. The editor layout is untouched — only the bar's
  own transform moves, so worst case is a slightly out-of-sync slide, never a
  wrong final position.
- **iPad collapse toggle**: `cap-ipad` (added when the physical screen's short
  side ≥ 640px, i.e. iPad even in a small window) puts a chevron in the footer to
  hide/show the bar, persisted. iPhone never gets it (the number pad has no
  operators, so the bar is essential there).

**Slash menu vs. division.** Typing `/` at line-start or after whitespace opens
the command menu. On web+iOS (`IS_WEB_SHELL`) **space** confirms the highlighted
command — but only once a query follows the `/` (`caret − triggerStart > 1`); a
bare `/` + space is division (`10 / 2`) and types a literal space. Desktop always
types a literal space here, so it's unaffected.

**`/clear` fix.** The slash-action path commits the trigger-fragment removal
(`scheduleSave()` + `render()`) itself, so an action that bails without further
editing can't strand stale overlay text.

**iOS long-press tabs.** App chrome is `user-select:none` +
`-webkit-touch-callout:none` (iOS text selection was hijacking the hold); the
synthesized `contextmenu` re-resolves its target via `elementFromPoint` if a
re-render detached the pressed node.

## Current state (shipped to TestFlight)

`main` is at **build 1301** (`4055c5e`). Recent builds, newest first:

| Build | Commit | What |
| --- | --- | --- |
| 1301 | `4055c5e` | Operator-bar keys: split `space \| return`, `=` replaces italic in ABC, bold drawn checkbox; `/`+space division fix |
| 1201 | `a979e7c` | FLIP slide for the 123/ABC bar transition (replaced the fade) |
| 1101 | `029306d` | iPad collapse toggle **(current)** + a morph fade **(superseded by 1201)** |
| 1001 | `9e095d7` | Keyboard bar lands instantly on switch (focus-keeper hop) |
| 901 | `bd86516` | Native editor scrolling, ABC markdown row, instant settings, help drawer, `/clear` fix |

All bugs from the last on-device testing round are addressed.

## Pending / needs on-device confirmation

- **FLIP slide feel** — timing is `230ms`; real iOS keyboard timing can't be
  reproduced headlessly, so tune to taste on device.
- **Return key** currently inserts a plain newline. Open question offered to the
  user: should it continue a bullet/numbered list when inside one (like the
  hardware Return)? Not yet wired.
- Autocorrect suggestions surviving the per-keystroke DOM rebuild in ABC mode.

## Tests

Smoke tests are written as small Playwright scripts placed under
`node_modules/.test-*.mjs` — **not committed** (node_modules is regenerated), so
treat them as disposable and recreate as needed. The pattern:

- A tiny HTTP server serves `web-dist/` and injects
  `<script>window.webkit={messageHandlers:{bridge:{postMessage:()=>{}}}}</script>`
  into `<head>` so the bundled `@capacitor/core` reports **native** (sets
  `cap-native`). A plain `window.Capacitor` stub does **not** work — the bundle
  re-derives platform from that `webkit.messageHandlers.bridge` marker.
- Launch Chromium from the preinstalled path under `/opt/pw-browsers/…`.
- Set the browser context `screen` to iPad vs iPhone dimensions to exercise
  `cap-ipad` (short side ≥ 640 ⇒ iPad).
- Coverage exercised so far: keyboard-switch focus continuity, the FLIP slide,
  the iPad collapse toggle + persistence, the ABC/123 key sets, the `/`+space
  fix (both the bar and physical-key paths), single-scroll-host layering,
  `/clear`, settings/back-nav speed under wedged IDB, and the help drawer.

Always pair these with `tsc` + both builds.

## Map

- Release runbook & Apple setup: `docs/IOS_RELEASE.md`
- Native shell glue: `src/web/native.ts`
- Native/mobile CSS: `WEB_CSS` in `build/build-web.mjs`
- Shared renderer: `src/renderer/popup.ts`
- Desktop-only CSS (**do not touch for mobile**): `src/renderer/popup.css`
- Capacitor config: `capacitor.config.ts`
