import { contextBridge, ipcRenderer, clipboard } from 'electron';
import type { Settings } from '../shared/types';

contextBridge.exposeInMainWorld('mathPopup', {
  getSettings: (): Promise<Settings> => ipcRenderer.invoke('settings:get'),
  setSettings: (partial: Partial<Settings>): Promise<Settings> =>
    ipcRenderer.invoke('settings:set', partial),
  hidePopup: (): Promise<void> => ipcRenderer.invoke('window:hide'),
  setAlwaysOnTop: (on: boolean): Promise<void> =>
    ipcRenderer.invoke('window:setAlwaysOnTop', on),
  openSettings: (): Promise<void> => ipcRenderer.invoke('settings:open'),
  openHelp: (): Promise<void> => ipcRenderer.invoke('help:open'),
  copyText: (text: string) => clipboard.writeText(text),
  onThemeChanged: (cb: (resolved: 'light' | 'dark') => void) => {
    const listener = (_e: unknown, resolved: 'light' | 'dark') => cb(resolved);
    ipcRenderer.on('theme:changed', listener);
    return () => ipcRenderer.removeListener('theme:changed', listener);
  },
  getAppVersion: (): Promise<string> => ipcRenderer.invoke('app:getVersion'),
  checkForUpdates: (): Promise<void> => ipcRenderer.invoke('update:check'),
  installUpdate: (): Promise<void> => ipcRenderer.invoke('update:install'),
  onUpdateStatus: (cb: (status: string) => void) => {
    const listener = (_e: unknown, status: string) => cb(status);
    ipcRenderer.on('update:status', listener);
    return () => ipcRenderer.removeListener('update:status', listener);
  }
});

declare global {
  interface Window {
    mathPopup: {
      getSettings(): Promise<Settings>;
      setSettings(partial: Partial<Settings>): Promise<Settings>;
      hidePopup(): Promise<void>;
      setAlwaysOnTop(on: boolean): Promise<void>;
      openSettings(): Promise<void>;
      openHelp(): Promise<void>;
      copyText(text: string): void;
      onThemeChanged(cb: (resolved: 'light' | 'dark') => void): () => void;
      getAppVersion(): Promise<string>;
      checkForUpdates(): Promise<void>;
      installUpdate(): Promise<void>;
      onUpdateStatus(cb: (status: string) => void): () => void;
    };
  }
}
