import { contextBridge, ipcRenderer, clipboard } from 'electron';
import type { Settings } from '../shared/types';

contextBridge.exposeInMainWorld('mathPopup', {
  getSettings: (): Promise<Settings> => ipcRenderer.invoke('settings:get'),
  setSettings: (partial: Partial<Settings>): Promise<Settings> =>
    ipcRenderer.invoke('settings:set', partial),
  hidePopup: (): Promise<void> => ipcRenderer.invoke('window:hide'),
  openSettings: (): Promise<void> => ipcRenderer.invoke('settings:open'),
  copyText: (text: string) => clipboard.writeText(text)
});

declare global {
  interface Window {
    mathPopup: {
      getSettings(): Promise<Settings>;
      setSettings(partial: Partial<Settings>): Promise<Settings>;
      hidePopup(): Promise<void>;
      openSettings(): Promise<void>;
      copyText(text: string): void;
    };
  }
}
