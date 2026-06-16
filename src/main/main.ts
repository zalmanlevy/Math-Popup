import { app, BrowserWindow, Tray, Menu, nativeImage, ipcMain, screen, globalShortcut, nativeTheme } from 'electron';
import { join } from 'node:path';
import { appendFileSync } from 'node:fs';
import { loadSettings, saveSettings, flushSettings } from './store';
import { Settings } from '../shared/types';
import { autoUpdater } from 'electron-updater';

const isDev = !app.isPackaged;
const startedHidden = process.argv.includes('--hidden');

let tray: Tray | null = null;
let popupWindow: BrowserWindow | null = null;
let settingsWindow: BrowserWindow | null = null;
let helpWindow: BrowserWindow | null = null;

type UpdatePhase =
  | 'idle'
  | 'checking'
  | 'available'
  | 'downloading'
  | 'downloaded'
  | 'not-available'
  | 'error';

interface UpdateState {
  phase: UpdatePhase;
  percent?: number;
  error?: string;
  version?: string;
}

let updateState: UpdateState = { phase: 'idle' };

const UPDATE_CHECK_INTERVAL_MS = 4 * 60 * 60 * 1000; // 4 hours
const INITIAL_UPDATE_DELAY_MS = 30 * 1000; // 30 seconds after launch

const ICON_PATH = join(__dirname, '..', 'assets', 'icon.png');
const PRELOAD_PATH = join(__dirname, 'preload.js');
const POPUP_HTML = join(__dirname, '..', 'renderer', 'popup.html');
const SETTINGS_HTML = join(__dirname, '..', 'renderer', 'settings.html');
const HELP_HTML = join(__dirname, '..', 'renderer', 'help.html');

const LIGHT_BG = '#fafafa';
const DARK_BG = '#0f1115';

function currentBg(): string {
  return nativeTheme.shouldUseDarkColors ? DARK_BG : LIGHT_BG;
}

function applyThemeSource(theme: Settings['theme']) {
  nativeTheme.themeSource = theme;
}

function broadcastTheme() {
  for (const w of [popupWindow, settingsWindow, helpWindow]) {
    if (w && !w.isDestroyed()) {
      w.webContents.send('theme:changed', nativeTheme.shouldUseDarkColors ? 'dark' : 'light');
      w.setBackgroundColor(currentBg());
    }
  }
}

function createPopup() {
  if (popupWindow && !popupWindow.isDestroyed()) {
    showPopup();
    return;
  }

  const settings = loadSettings();
  const display = screen.getPrimaryDisplay();
  const defaultBounds = {
    width: 460,
    height: 560,
    x: display.workArea.x + display.workArea.width - 480,
    y: display.workArea.y + display.workArea.height - 600
  };
  const bounds = settings.windowBounds ?? defaultBounds;

  popupWindow = new BrowserWindow({
    width: bounds.width,
    height: bounds.height,
    x: bounds.x,
    y: bounds.y,
    minWidth: 280,
    minHeight: 200,
    frame: false,
    transparent: false,
    backgroundColor: currentBg(),
    show: false,
    resizable: true,
    skipTaskbar: !settings.showTaskbarIcon,
    alwaysOnTop: settings.alwaysOnTop,
    icon: ICON_PATH,
    webPreferences: {
      preload: PRELOAD_PATH,
      contextIsolation: true,
      sandbox: false,
      nodeIntegration: false
    }
  });

  popupWindow.loadFile(POPUP_HTML);

  popupWindow.once('ready-to-show', () => {
    showPopup();
  });

  const persistBounds = () => {
    if (!popupWindow || popupWindow.isDestroyed()) return;
    const b = popupWindow.getBounds();
    saveSettings({ windowBounds: b });
  };

  popupWindow.on('move', persistBounds);
  popupWindow.on('resize', persistBounds);

  popupWindow.on('closed', () => {
    popupWindow = null;
  });
}

// If the saved position lands off every connected display — e.g. a secondary
// monitor (the bounds can be negative) was unplugged or went to sleep — the
// window would open invisibly and look like it "disappeared". Re-center it on the
// primary display in that case. Called on every show, so a monitor change between
// summons can't strand it off-screen. No-op while the saved monitor is present.
function placeOnVisibleScreen(win: BrowserWindow) {
  const b = win.getBounds();
  const onScreen = screen.getAllDisplays().some(d => {
    const a = d.workArea;
    return b.x < a.x + a.width && b.x + b.width > a.x &&
           b.y < a.y + a.height && b.y + b.height > a.y;
  });
  if (onScreen) return;
  const wa = screen.getPrimaryDisplay().workArea;
  const width = Math.min(b.width, wa.width);
  const height = Math.min(b.height, wa.height);
  win.setBounds({
    x: Math.round(wa.x + (wa.width - width) / 2),
    y: Math.round(wa.y + (wa.height - height) / 2),
    width,
    height,
  });
}

// Bring the popup on-screen and re-apply the pin before showing. Re-asserting
// always-on-top on every show guards against Electron dropping the flag across a
// hide/show cycle on Windows.
function showPopup() {
  if (!popupWindow || popupWindow.isDestroyed()) return;
  placeOnVisibleScreen(popupWindow);
  popupWindow.setAlwaysOnTop(loadSettings().alwaysOnTop);
  popupWindow.show();
  popupWindow.focus();
}

function togglePopup() {
  if (popupWindow && !popupWindow.isDestroyed()) {
    if (popupWindow.isVisible() && popupWindow.isFocused()) {
      popupWindow.hide();
    } else {
      showPopup();
    }
  } else {
    createPopup();
  }
}

function openSettings() {
  if (settingsWindow && !settingsWindow.isDestroyed()) {
    settingsWindow.show();
    settingsWindow.focus();
    return;
  }
  settingsWindow = new BrowserWindow({
    width: 540,
    height: 680,
    title: 'Math Popup — Settings',
    backgroundColor: currentBg(),
    autoHideMenuBar: true,
    icon: ICON_PATH,
    webPreferences: {
      preload: PRELOAD_PATH,
      contextIsolation: true,
      sandbox: false,
      nodeIntegration: false
    }
  });
  settingsWindow.loadFile(SETTINGS_HTML);
  settingsWindow.on('closed', () => { settingsWindow = null; });
}

function openHelp() {
  if (helpWindow && !helpWindow.isDestroyed()) {
    helpWindow.show();
    helpWindow.focus();
    return;
  }
  helpWindow = new BrowserWindow({
    width: 720,
    height: 760,
    title: 'Math Popup — Help',
    backgroundColor: currentBg(),
    autoHideMenuBar: true,
    icon: ICON_PATH,
    webPreferences: {
      preload: PRELOAD_PATH,
      contextIsolation: true,
      sandbox: false,
      nodeIntegration: false
    }
  });
  helpWindow.loadFile(HELP_HTML);
  helpWindow.on('closed', () => { helpWindow = null; });
}

function buildTray() {
  const image = nativeImage.createFromPath(ICON_PATH);
  // On Windows, the system tray expects a small icon; resize if necessary.
  const trayImg = image.isEmpty() ? nativeImage.createEmpty() : image.resize({ width: 16, height: 16 });
  tray = new Tray(trayImg);
  tray.setToolTip('Math Popup');
  tray.on('click', togglePopup);
  tray.on('double-click', togglePopup);
  tray.setContextMenu(Menu.buildFromTemplate([
    { label: 'Open', click: togglePopup },
    { label: 'Settings…', click: openSettings },
    { type: 'separator' },
    { label: 'Quit', click: () => { app.quit(); } }
  ]));
}

function applyStartup(enabled: boolean) {
  // Skip on Linux: setLoginItemSettings is a no-op there in stock Electron.
  if (process.platform === 'linux') return;
  const opts: Electron.Settings = { openAtLogin: enabled };
  if (process.platform === 'win32') {
    opts.path = process.execPath;
    if (app.isPackaged) {
      // Packaged: process.execPath is the app's exe; Electron resolves the
      // bundled app from there, so only --hidden is needed.
      opts.args = ['--hidden'];
    } else {
      // Dev: pass the app path so the electron binary knows what to load.
      // Windows' Run registry key joins args with spaces *without* quoting
      // them, so any space in the project path (e.g. "OneDrive\General
      // Projects") would chop the path at the first space and produce a
      // "Cannot find module" error on boot. Wrap it in quotes ourselves.
      opts.args = [`"${app.getAppPath()}"`, '--hidden'];
    }
  }
  app.setLoginItemSettings(opts);
}

function registerIPC() {
  ipcMain.handle('settings:get', () => loadSettings());
  ipcMain.handle('settings:set', (_e, partial: Partial<Settings>) => {
    const updated = saveSettings(partial);
    if (Object.prototype.hasOwnProperty.call(partial, 'launchAtStartup')) {
      applyStartup(updated.launchAtStartup);
    }
    if (Object.prototype.hasOwnProperty.call(partial, 'theme')) {
      applyThemeSource(updated.theme);
      broadcastTheme();
    }
    if (Object.prototype.hasOwnProperty.call(partial, 'showTaskbarIcon')) {
      // Live-apply so the taskbar button appears/disappears without a restart.
      if (popupWindow && !popupWindow.isDestroyed()) {
        popupWindow.setSkipTaskbar(!updated.showTaskbarIcon);
      }
    }
    return updated;
  });
  ipcMain.handle('window:hide', () => {
    if (popupWindow && !popupWindow.isDestroyed()) popupWindow.hide();
  });
  ipcMain.handle('window:setAlwaysOnTop', (_e, on: boolean) => {
    if (popupWindow && !popupWindow.isDestroyed()) popupWindow.setAlwaysOnTop(on);
    saveSettings({ alwaysOnTop: on });
  });
  ipcMain.handle('settings:open', () => openSettings());
  ipcMain.handle('help:open', () => openHelp());

  ipcMain.handle('app:getVersion', () => app.getVersion());
  ipcMain.handle('update:getState', () => updateState);
  ipcMain.handle('update:check', () => triggerUpdateCheck());
  ipcMain.handle('update:install', () => {
    if (app.isPackaged) {
      // isSilent = true, isForceRunAfter = true
      autoUpdater.quitAndInstall(true, true);
    }
  });

  // Route electron-updater diagnostics to a file (userData/update.log) so an
  // update failure — e.g. a locked file during the old-version uninstall — can
  // be inspected after the fact instead of guessing.
  const updateLogPath = join(app.getPath('userData'), 'update.log');
  const logUpdate = (level: string, message: unknown) => {
    try {
      const text = typeof message === 'string' ? message : JSON.stringify(message);
      appendFileSync(updateLogPath, `[${new Date().toISOString()}] ${level} ${text}\n`);
    } catch { /* logging must never break an update */ }
  };
  autoUpdater.logger = {
    info: (m: unknown) => logUpdate('INFO', m),
    warn: (m: unknown) => logUpdate('WARN', m),
    error: (m: unknown) => logUpdate('ERROR', m),
    debug: (m: unknown) => logUpdate('DEBUG', m)
  };

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = false;

  autoUpdater.on('checking-for-update', () => setUpdateState({ phase: 'checking' }));
  autoUpdater.on('update-available', (info) => setUpdateState({ phase: 'available', version: info?.version }));
  autoUpdater.on('update-not-available', () => setUpdateState({ phase: 'not-available' }));
  autoUpdater.on('error', (err) => setUpdateState({ phase: 'error', error: err.message }));
  autoUpdater.on('download-progress', (progressObj) => {
    setUpdateState({ phase: 'downloading', percent: Math.round(progressObj.percent) });
  });
  autoUpdater.on('update-downloaded', (info) => setUpdateState({ phase: 'downloaded', version: info?.version }));
}

function setUpdateState(next: UpdateState) {
  updateState = next;
  for (const w of [popupWindow, settingsWindow]) {
    if (w && !w.isDestroyed()) {
      w.webContents.send('update:state', updateState);
    }
  }
}

function triggerUpdateCheck() {
  if (!app.isPackaged) {
    setUpdateState({ phase: 'not-available' });
    return;
  }
  autoUpdater.checkForUpdates().catch(err => {
    setUpdateState({ phase: 'error', error: err.message });
  });
}

app.whenReady().then(() => {
  // Match the installer's appId so a taskbar button (when the user enables the
  // taskbar icon) groups under the app's identity and pins to the right target.
  if (process.platform === 'win32') app.setAppUserModelId('com.zalmanlevy.mathpopup');
  registerIPC();
  const initial = loadSettings();
  applyThemeSource(initial.theme);
  // Push theme to all open windows whenever the OS or themeSource changes.
  nativeTheme.on('updated', broadcastTheme);
  buildTray();
  // Sync the OS auto-launch entry with the saved setting on every boot, so
  // toggling the setting takes effect from then on (including any path
  // changes if the project was moved).
  applyStartup(initial.launchAtStartup);
  // When launched at OS login, start hidden — tray icon only, no popup.
  if (!startedHidden) createPopup();

  // Optional global shortcut to toggle the popup. Ctrl+Alt+M.
  globalShortcut.register('Control+Alt+M', togglePopup);

  // Update checks: a delayed first check (so launch isn't slowed) plus a
  // recurring 4h interval. Gated to packaged builds — autoUpdater can't pull
  // releases from the dev tree.
  if (app.isPackaged) {
    setTimeout(triggerUpdateCheck, INITIAL_UPDATE_DELAY_MS);
    setInterval(triggerUpdateCheck, UPDATE_CHECK_INTERVAL_MS);
  }

  // Ensure any target="_blank" links open in the user's default web browser
  app.on('web-contents-created', (_, contents) => {
    contents.setWindowOpenHandler(({ url }) => {
      if (url.startsWith('http')) {
        require('electron').shell.openExternal(url);
      }
      return { action: 'deny' };
    });
  });
});

app.on('window-all-closed', (e: Electron.Event) => {
  // Keep the app alive in the tray even when all windows are closed.
  e.preventDefault();
});

app.on('before-quit', () => {
  flushSettings();
  // Guard: when a second instance loses the single-instance lock it calls
  // app.quit() before whenReady fires, and globalShortcut throws if used
  // before the app is ready.
  if (app.isReady()) globalShortcut.unregisterAll();
});

// Single-instance lock so a second launch focuses the existing popup.
const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    togglePopup();
  });
}
