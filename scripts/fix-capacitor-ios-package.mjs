// Normalizes path separators in ios/App/CapApp-SPM/Package.swift after `cap sync`.
//
// Running `npx cap sync ios` on Windows can rewrite the local-package paths with
// backslashes (e.g. `path: "..\\..\\..\\node_modules\\@capacitor\\app"`), which
// Swift Package Manager on macOS then fails to resolve. This script rewrites any
// backslash separators inside `path: "..."` strings back to forward slashes.
// It is a no-op when the file is already clean, so it's safe to run everywhere
// (macOS CI included) — wired into `npm run ios:sync`.
import { readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const file = join(root, 'ios', 'App', 'CapApp-SPM', 'Package.swift');

let text;
try {
  text = await readFile(file, 'utf8');
} catch {
  console.log('[fix-ios-package] Package.swift not found (run `npx cap sync ios` first) — nothing to do');
  process.exit(0);
}

const fixed = text.replace(/path:\s*"([^"]*)"/g, (m, p) =>
  p.includes('\\') ? `path: "${p.replace(/\\+/g, '/')}"` : m
);

if (fixed !== text) {
  await writeFile(file, fixed, 'utf8');
  console.log('[fix-ios-package] normalized backslash paths in CapApp-SPM/Package.swift');
} else {
  console.log('[fix-ios-package] Package.swift already clean');
}
