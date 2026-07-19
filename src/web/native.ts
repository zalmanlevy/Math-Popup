// Capacitor (iOS/iPad) native-shell glue. Bundled into shim.js for every web
// page, but everything here no-ops unless the page is actually running inside
// the Capacitor native app — the browser/PWA build and the Electron desktop app
// are untouched. Detection uses the bridge object the native runtime injects
// into the webview before any page script runs.
import { registerPlugin } from '@capacitor/core';
import { StatusBar, Style } from '@capacitor/status-bar';
import { Keyboard, KeyboardStyle } from '@capacitor/keyboard';
import { Haptics, ImpactStyle } from '@capacitor/haptics';
import { App as CapApp } from '@capacitor/app';

// In-app native plugin (see ios/App/App/AppViewController.swift): reports
// whether the app runs as a resizable iPadOS window instead of fullscreen.
interface WindowStatePlugin {
  get(): Promise<{ windowed: boolean }>;
  addListener(
    eventName: 'windowedchange',
    listener: (state: { windowed: boolean }) => void
  ): Promise<unknown>;
}
const WindowState = registerPlugin<WindowStatePlugin>('WindowState');

interface CapacitorGlobal {
  isNativePlatform?: () => boolean;
}

export const isNativeApp: boolean = (() => {
  try {
    return (window as unknown as { Capacitor?: CapacitorGlobal }).Capacitor?.isNativePlatform?.() === true;
  } catch {
    return false;
  }
})();

// The app's three theme prefs resolve to a concrete light/dark here (mirrors the
// CSS: data-theme="system" defers to prefers-color-scheme).
function resolvedTheme(): 'light' | 'dark' {
  const pref = document.documentElement.getAttribute('data-theme');
  if (pref === 'dark') return 'dark';
  if (pref === 'light') return 'light';
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

// Native chrome follows the RESOLVED theme: status-bar icons flip so they stay
// visible (Style.Light = dark icons on the light app background), and the iOS
// keyboard matches the app instead of popping up glaring white in dark mode.
function syncNativeChrome(): void {
  const dark = resolvedTheme() === 'dark';
  StatusBar.setStyle({ style: dark ? Style.Dark : Style.Light }).catch(() => {});
  Keyboard.setStyle({ style: dark ? KeyboardStyle.Dark : KeyboardStyle.Light }).catch(() => {});
}

// Light tap confirmation for copy actions (result chips, copy shortcuts). Fire
// and forget; never blocks the copy itself.
export function nativeCopyHaptic(): void {
  if (!isNativeApp) return;
  Haptics.impact({ style: ImpactStyle.Light }).catch(() => {});
}

// Wire the native shell. `onHidden` is the storage flush — iOS can suspend or
// kill a backgrounded webview at any point, and the native appStateChange event
// is more dependable there than pagehide/visibilitychange alone.
export function initNativeShell(onHidden: () => void): void {
  if (!isNativeApp) return;
  // Native-only CSS hooks (edge-to-edge shell, safe-area padding — see web.css).
  document.documentElement.classList.add('cap-native');
  syncNativeChrome();
  // data-theme is set per page load and whenever the user changes the theme
  // setting; the media query covers "system" flipping while the app is open.
  new MutationObserver(syncNativeChrome)
    .observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', syncNativeChrome);
  CapApp.addListener('appStateChange', ({ isActive }) => { if (!isActive) onHidden(); }).catch(() => {});
  // .kb-open marks the ON-SCREEN keyboard being up (native events — a hardware
  // keyboard never fires these). The renderer's math key bar shows only then.
  Keyboard.addListener('keyboardWillShow', () => {
    document.documentElement.classList.add('kb-open');
  }).catch(() => {});
  Keyboard.addListener('keyboardWillHide', () => {
    document.documentElement.classList.remove('kb-open');
  }).catch(() => {});
  // iPadOS windowed mode: .cap-windowed drives the clearance for the window's
  // traffic-light controls (title bar) and rounded corners / resize grip
  // (footer). Pull on load — each page is a fresh document — then live-update.
  const applyWindowed = (windowed: boolean) =>
    document.documentElement.classList.toggle('cap-windowed', windowed);
  WindowState.get().then(s => applyWindowed(s.windowed)).catch(() => {});
  // catch: never surface a bridge rejection — the inline bootstrap guard treats
  // an unhandled one as a failed startup and unhides the fallback editor text.
  WindowState.addListener('windowedchange', s => applyWindowed(s.windowed))
    .then(() => {}, () => {});
}
