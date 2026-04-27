import { app } from 'electron';
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { DEFAULT_SETTINGS, Settings } from '../shared/types';

let cache: Settings | null = null;
let saveTimer: NodeJS.Timeout | null = null;

function filePath(): string {
  return join(app.getPath('userData'), 'settings.json');
}

export function loadSettings(): Settings {
  if (cache) return cache;
  const path = filePath();
  if (existsSync(path)) {
    try {
      const raw = JSON.parse(readFileSync(path, 'utf8'));
      cache = { ...DEFAULT_SETTINGS, ...raw };
      // Merge suffix arrays: prefer user value if present
      if (!Array.isArray(cache!.suffixes) || cache!.suffixes.length === 0) {
        cache!.suffixes = DEFAULT_SETTINGS.suffixes;
      }
      return cache!;
    } catch {
      cache = { ...DEFAULT_SETTINGS };
      return cache;
    }
  }
  cache = { ...DEFAULT_SETTINGS };
  return cache;
}

export function saveSettings(next: Partial<Settings>): Settings {
  const merged: Settings = { ...loadSettings(), ...next };
  cache = merged;
  if (saveTimer) clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    const path = filePath();
    if (!existsSync(dirname(path))) mkdirSync(dirname(path), { recursive: true });
    writeFileSync(path, JSON.stringify(merged, null, 2), 'utf8');
  }, 150);
  return merged;
}

export function flushSettings(): void {
  if (saveTimer) {
    clearTimeout(saveTimer);
    saveTimer = null;
  }
  if (!cache) return;
  const path = filePath();
  if (!existsSync(dirname(path))) mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, JSON.stringify(cache, null, 2), 'utf8');
}
