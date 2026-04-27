export interface Suffix {
  symbol: string;       // e.g. "m", "M", "k", "B"
  multiplier: number;   // e.g. 1_000_000
  caseSensitive: boolean;
}

export type Mode = 'math' | 'text';

export interface Settings {
  mode: Mode;                     // 'math' = full evaluator + gutters; 'text' = plain notepad
  launchAtStartup: boolean;       // start with Windows; opens hidden in the tray
  suffixes: Suffix[];
  autoFormatNumbers: boolean;     // insert commas on space/operator
  expandSuffixesInEditor: boolean; // 1m -> 1,000,000 in textarea
  decimals: number;               // fixed number of decimal places shown in results
  noteContent: string;            // last saved note text
  windowBounds?: { x: number; y: number; width: number; height: number };
}

export const DEFAULT_SETTINGS: Settings = {
  mode: 'math',
  launchAtStartup: true,
  suffixes: [
    { symbol: 'k', multiplier: 1_000, caseSensitive: false },
    { symbol: 'm', multiplier: 1_000_000, caseSensitive: false },
    { symbol: 'b', multiplier: 1_000_000_000, caseSensitive: false },
    { symbol: 't', multiplier: 1_000_000_000_000, caseSensitive: false }
  ],
  autoFormatNumbers: true,
  expandSuffixesInEditor: true,
  decimals: 2,
  noteContent: ''
};
