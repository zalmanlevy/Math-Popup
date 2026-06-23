export interface Suffix {
  symbol: string;       // e.g. "m", "M", "k", "B"
  multiplier: number;   // e.g. 1_000_000
  caseSensitive: boolean;
}

export type Mode = 'math' | 'text';
export type ThemePref = 'system' | 'light' | 'dark';

export interface Page {
  id: string;
  title: string;
  content: string;
  mode: Mode;              // legacy per-page mode; kept for migration + as the seed for lineModes
  lineModes?: Mode[];      // per-line mode, parallel to content's lines ('text' by default)
  obsidianPath?: string;   // absolute path to a linked Obsidian/Markdown note
}

export interface ObsidianRecentNote {
  path: string;
  title: string;
  lastOpenedAt: number;
}

export interface WindowBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface Settings {
  mode: Mode;                     // 'math' = full evaluator + gutters; 'text' = plain notepad
  launchAtStartup: boolean;       // start with Windows; opens hidden in the tray
  showTaskbarIcon: boolean;       // show the window in the Windows taskbar (so it can be pinned); off = tray-only
  advancedMode: boolean;          // enables optional power-user integrations such as Obsidian note linking
  tabBarOpen: boolean;            // whether the tab bar is expanded (remembered across launches)
  suffixes: Suffix[];
  autoFormatNumbers: boolean;     // insert commas on space/operator
  expandSuffixesInEditor: boolean; // 1m -> 1,000,000 in textarea
  decimals: number;               // fixed number of decimal places shown in results
  noteContent: string;            // last saved note text
  alwaysOnTop: boolean;           // pin window above other apps
  theme: ThemePref;               // 'system' follows OS, otherwise forced
  zoom: number;                   // current popup zoom factor (1.0 = 100%)
  zoomDefault: number;            // user's preferred default; Ctrl+0 returns here
  windowBounds?: WindowBounds;
  pages?: Page[];
  activePageId?: string;
  closedPages?: Page[];
  obsidianRecentNotes?: ObsidianRecentNote[];
}

export const ZOOM_MIN = 0.5;
export const ZOOM_MAX = 2.0;
export const ZOOM_STEP = 0.1;

export const DEFAULT_SETTINGS: Settings = {
  mode: 'math',
  launchAtStartup: true,
  showTaskbarIcon: false,
  advancedMode: false,
  tabBarOpen: false,
  suffixes: [
    { symbol: 'k', multiplier: 1_000, caseSensitive: false },
    { symbol: 'm', multiplier: 1_000_000, caseSensitive: false },
    { symbol: 'b', multiplier: 1_000_000_000, caseSensitive: false },
    { symbol: 't', multiplier: 1_000_000_000_000, caseSensitive: false }
  ],
  autoFormatNumbers: true,
  expandSuffixesInEditor: true,
  decimals: 2,
  noteContent: '',
  alwaysOnTop: false,
  theme: 'system',
  zoom: 1.0,
  zoomDefault: 1.0,
  obsidianRecentNotes: []
};
