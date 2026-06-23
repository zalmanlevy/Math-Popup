import { contextBridge, ipcRenderer, clipboard, webFrame } from 'electron';
import type { Settings } from '../shared/types';

export type UpdatePhase =
  | 'idle'
  | 'checking'
  | 'available'
  | 'downloading'
  | 'downloaded'
  | 'not-available'
  | 'error';

export interface UpdateState {
  phase: UpdatePhase;
  percent?: number;
  error?: string;
  version?: string;
}

export interface ObsidianNotePayload {
  path: string;
  title: string;
  content: string;
  mtimeMs: number;
}

export interface ObsidianFileChange extends Partial<ObsidianNotePayload> {
  path: string;
  error?: string;
}

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
  setZoomFactor: (factor: number) => webFrame.setZoomFactor(factor),
  onThemeChanged: (cb: (resolved: 'light' | 'dark') => void) => {
    const listener = (_e: unknown, resolved: 'light' | 'dark') => cb(resolved);
    ipcRenderer.on('theme:changed', listener);
    return () => ipcRenderer.removeListener('theme:changed', listener);
  },
  onSettingsChanged: (cb: (settings: Settings) => void) => {
    const listener = (_e: unknown, settings: Settings) => cb(settings);
    ipcRenderer.on('settings:changed', listener);
    return () => ipcRenderer.removeListener('settings:changed', listener);
  },
  getAppVersion: (): Promise<string> => ipcRenderer.invoke('app:getVersion'),
  getUpdateState: (): Promise<UpdateState> => ipcRenderer.invoke('update:getState'),
  checkForUpdates: (): Promise<void> => ipcRenderer.invoke('update:check'),
  installUpdate: (): Promise<void> => ipcRenderer.invoke('update:install'),
  chooseObsidianNote: (): Promise<ObsidianNotePayload | null> =>
    ipcRenderer.invoke('obsidian:chooseNote'),
  readObsidianNote: (path: string): Promise<ObsidianNotePayload> =>
    ipcRenderer.invoke('obsidian:readNote', path),
  writeObsidianNote: (path: string, content: string): Promise<ObsidianNotePayload> =>
    ipcRenderer.invoke('obsidian:writeNote', path, content),
  watchObsidianNotes: (paths: string[]): Promise<string[]> =>
    ipcRenderer.invoke('obsidian:watchNotes', paths),
  onObsidianFileChanged: (cb: (note: ObsidianFileChange) => void) => {
    const listener = (_e: unknown, note: ObsidianFileChange) => cb(note);
    ipcRenderer.on('obsidian:fileChanged', listener);
    return () => ipcRenderer.removeListener('obsidian:fileChanged', listener);
  },
  onUpdateState: (cb: (state: UpdateState) => void) => {
    const listener = (_e: unknown, state: UpdateState) => cb(state);
    ipcRenderer.on('update:state', listener);
    return () => ipcRenderer.removeListener('update:state', listener);
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
      setZoomFactor(factor: number): void;
      onThemeChanged(cb: (resolved: 'light' | 'dark') => void): () => void;
      onSettingsChanged(cb: (settings: Settings) => void): () => void;
      getAppVersion(): Promise<string>;
      getUpdateState(): Promise<UpdateState>;
      checkForUpdates(): Promise<void>;
      installUpdate(): Promise<void>;
      chooseObsidianNote(): Promise<ObsidianNotePayload | null>;
      readObsidianNote(path: string): Promise<ObsidianNotePayload>;
      writeObsidianNote(path: string, content: string): Promise<ObsidianNotePayload>;
      watchObsidianNotes(paths: string[]): Promise<string[]>;
      onObsidianFileChanged(cb: (note: ObsidianFileChange) => void): () => void;
      onUpdateState(cb: (state: UpdateState) => void): () => void;
    };
  }
}
