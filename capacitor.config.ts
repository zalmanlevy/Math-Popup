import type { CapacitorConfig } from '@capacitor/cli';

// iOS/iPad native shell configuration. The native app wraps the SAME static
// site the web/PWA build produces (web-dist/, see build/build-web.mjs) — the
// Electron desktop build (dist/) is completely separate and unaffected.
const config: CapacitorConfig = {
  appId: 'com.zalmanlevy.mathpopup',
  appName: 'Math Popup',
  webDir: 'web-dist',
  // Matches the app's light page background (--bg). Prevents a white flash on
  // the native webview surface before first paint.
  backgroundColor: '#fafafa',
  ios: {
    // CSS owns the safe areas (env(safe-area-inset-*) + viewport-fit=cover in
    // the web build's head). 'automatic' would add a native inset ON TOP of the
    // CSS padding — double top gap + phantom scroll.
    contentInset: 'never',
  },
  plugins: {
    Keyboard: {
      // Resize the whole webview when the on-screen keyboard shows, so the
      // app's flex column (title bar / editor / status bar) reflows and the
      // renderer's resize path keeps the caret line visible. Pinned explicitly
      // so a future plugin default change can't alter behavior.
      resize: 'native',
    },
  },
};

export default config;
