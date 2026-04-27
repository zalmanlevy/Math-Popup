import { app, BrowserWindow, Tray, Menu, nativeImage, ipcMain, screen, globalShortcut, nativeTheme } from 'electron';
import { join } from 'node:path';
import { loadSettings, saveSettings, flushSettings } from './store';
import { Settings } from '../shared/types';

const isDev = !app.isPackaged;
const startedHidden = process.argv.includes('--hidden');

let tray: Tray | null = null;
let popupWindow: BrowserWindow | null = null;
let settingsWindow: BrowserWindow | null = null;
let helpWindow: BrowserWindow | null = null;

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
    popupWindow.show();
    popupWindow.focus();
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
    minWidth: 220,
    minHeight: 200,
    frame: false,
    transparent: false,
    backgroundColor: currentBg(),
    show: false,
    resizable: true,
    skipTaskbar: true,
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
    popupWindow?.show();
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

function togglePopup() {
  if (popupWindow && !popupWindow.isDestroyed()) {
    if (popupWindow.isVisible() && popupWindow.isFocused()) {
      popupWindow.hide();
    } else {
      popupWindow.show();
      popupWindow.focus();
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
    // Use the Electron binary + our app dir + --hidden so the OS-launched
    // instance only puts the tray icon up (no popup until the user clicks).
    opts.path = process.execPath;
    opts.args = [app.getAppPath(), '--hidden'];
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
}

app.whenReady().then(() => {
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
