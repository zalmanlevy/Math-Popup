# Shipping Math Popup to iPhone/iPad (TestFlight → App Store)

The iOS app is the web build (`web-dist/`, same one deployed to Vercel) wrapped
in a Capacitor native shell (`ios/`). Everything builds and uploads from GitHub
Actions on a macOS runner — **no Mac is needed anywhere**. The Electron/Windows
app (`src/main`, `dist/`, electron-builder) is a completely separate pipeline
and is untouched by any of this.

## How the pieces fit

```
build/build-web.mjs  ->  web-dist/          (same output the PWA deploys)
npx cap sync ios     ->  ios/App/App/public (webview assets, gitignored)
ios/App/…            ->  native Xcode project (Capacitor 8, SPM — no CocoaPods)
.github/workflows/ios-testflight.yml -> archive, sign, upload to TestFlight
```

Native behavior lives in:

- `capacitor.config.ts` — app id/name, `webDir: web-dist`, `contentInset:
  'never'` (CSS owns safe areas), keyboard resize mode.
- `src/web/native.ts` — status-bar + keyboard theme sync, copy haptics,
  background settings flush. No-ops on web/PWA and desktop.
- `build/build-web.mjs` (`WEB_CSS`, `:root.cap-native` block) — edge-to-edge
  shell + safe-area padding, applied only inside the native app.
- Local dev loop: `npm run ios:sync` (builds web, syncs to `ios/`, then
  normalizes any Windows backslash paths in `CapApp-SPM/Package.swift` —
  required when syncing from a Windows machine).

## One-time Apple setup (human, in a browser)

1. **Register the bundle ID** `com.zalmanlevy.mathpopup`:
   developer.apple.com → Certificates, Identifiers & Profiles → Identifiers →
   “+” → App IDs → type *App*. No capability checkboxes needed.
2. **App Store Connect API key** (reuse the existing team key if you have one):
   App Store Connect → Users and Access → Integrations → App Store Connect API
   → Team Keys. **Role must be Admin** — cloud-managed signing cannot mint the
   distribution certificate with an App Manager key (fails with a cloud-signing
   permission error). Record Issuer ID + Key ID; the `.p8` downloads once.
3. **Create the app record**: App Store Connect → Apps → “+” → New App → iOS,
   bundle ID `com.zalmanlevy.mathpopup`, any SKU. App names are globally
   unique — have a fallback ready if “Math Popup” is taken.
4. **Repo secrets** (Settings → Secrets and variables → Actions → *Secrets*):
   - `APPLE_TEAM_ID` — the team ID from the developer portal
   - `ASC_KEY_ID` — API key ID
   - `ASC_ISSUER_ID` — API issuer ID
   - `ASC_API_PRIVATE_KEY` — entire `.p8` contents including BEGIN/END lines

## Releasing a build

1. Merge to `main` (the workflow file must exist on the default branch before
   `workflow_dispatch` works — a dispatch against a branch 404s until then).
2. Actions → **iOS TestFlight** → Run workflow (ref: `main`). ~35 billed macOS
   minutes per run.
3. First-ever build can take 30–60+ min to process in App Store Connect (extra
   malware scans); later builds process in 5–15 min. Processing builds hide in
   the collapsed **“Build Uploads”** section of the TestFlight tab.
4. TestFlight tab → create an **INTERNAL** testing group, add yourself, and
   turn ON “automatic distribution”: every future CI build reaches your
   devices with zero clicks. (External groups trigger a ~24 h Beta App Review
   before invites go out — don’t use one for self-testing.)
5. Invite emails often land in Gmail Spam/Promotions — search “TestFlight”.
   “Ready to Submit” status means processed and testable; nothing is waiting.
6. Builds auto-update on device and expire after 90 days.

Version numbers: the App Store *version* comes from `package.json` `version`;
the *build number* is `run_number × 100 + run_attempt` (always unique, safe
across re-runs). Nothing to bump by hand for TestFlight builds.

## App Store submission (after TestFlight verification)

- Screenshots: 6.9″ **1320×2868** PNG (3–10) and 13″ iPad **2064×2752** PNG —
  Chrome DevTools device emulation at exactly those sizes works fine.
- Privacy policy URL: the app stores everything locally (no accounts, no
  network I/O), so a short static privacy page on the existing Vercel deploy is
  enough.
- App Privacy labels: “Data Not Collected” (nothing leaves the device).
- Complete the age-rating questionnaire (4+).
- Organization accounts: complete the **EU DSA trader status** declaration
  before the app can go live in the EU.
- First review: 1–5 days.

## Native project notes

- Capacitor 8 + Swift Package Manager: there is **no** `.xcworkspace` and no
  Podfile — CI builds `-project ios/App/App.xcodeproj -scheme App`. The shared
  scheme is committed at `ios/App/App.xcodeproj/xcshareddata/xcschemes/` (the
  Capacitor template doesn’t ship one; without it `xcodebuild -scheme App`
  fails on a clean checkout).
- `ios/App/App/public/` and `ios/App/App/capacitor.config.json` are generated
  by `cap sync` and gitignored — CI regenerates them every build.
- `Info.plist` already sets `ITSAppUsesNonExemptEncryption=false` (HTTPS-only /
  local app ⇒ no per-build export-compliance prompt) and
  `UIRequiredDeviceCapabilities=arm64`.
- The app icon is generated into the asset catalog by `node build/gen-icons.mjs`
  (needs `npm i -D sharp` once, locally) from `assets/icon.png` — 1024×1024,
  alpha stripped (the App Store rejects transparent icons).
- PR touching `ios/**` → the **iOS Simulator Check** workflow compiles the
  native project unsigned on the generic simulator; that’s the no-Mac way to
  verify Swift/project changes.
