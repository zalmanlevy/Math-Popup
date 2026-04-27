import { build, context } from 'esbuild';
import { mkdir, copyFile, writeFile, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');
const watch = process.argv.includes('--watch');

const mainCfg = {
  entryPoints: [join(root, 'src/main/main.ts')],
  bundle: true,
  platform: 'node',
  target: 'node18',
  format: 'cjs',
  external: ['electron'],
  outfile: join(root, 'dist/main/main.js'),
  sourcemap: true,
  logLevel: 'info'
};

const preloadCfg = {
  entryPoints: [join(root, 'src/main/preload.ts')],
  bundle: true,
  platform: 'node',
  target: 'node18',
  format: 'cjs',
  external: ['electron'],
  outfile: join(root, 'dist/main/preload.js'),
  sourcemap: true,
  logLevel: 'info'
};

const rendererPopupCfg = {
  entryPoints: [join(root, 'src/renderer/popup.ts')],
  bundle: true,
  platform: 'browser',
  target: 'chrome120',
  format: 'iife',
  outfile: join(root, 'dist/renderer/popup.js'),
  sourcemap: true,
  logLevel: 'info'
};

const rendererSettingsCfg = {
  entryPoints: [join(root, 'src/renderer/settings.ts')],
  bundle: true,
  platform: 'browser',
  target: 'chrome120',
  format: 'iife',
  outfile: join(root, 'dist/renderer/settings.js'),
  sourcemap: true,
  logLevel: 'info'
};

async function ensureDir(p) {
  if (!existsSync(p)) await mkdir(p, { recursive: true });
}

async function copyAssets() {
  await ensureDir(join(root, 'dist/renderer'));
  await copyFile(join(root, 'src/renderer/popup.html'), join(root, 'dist/renderer/popup.html'));
  await copyFile(join(root, 'src/renderer/settings.html'), join(root, 'dist/renderer/settings.html'));
  await copyFile(join(root, 'src/renderer/popup.css'), join(root, 'dist/renderer/popup.css'));
  await copyFile(join(root, 'src/renderer/settings.css'), join(root, 'dist/renderer/settings.css'));

  await ensureDir(join(root, 'dist/assets'));
  await copyFile(join(root, 'assets/icon.png'), join(root, 'dist/assets/icon.png'));
}

async function run() {
  await ensureDir(join(root, 'dist/main'));
  await ensureDir(join(root, 'dist/renderer'));

  if (watch) {
    const ctxs = await Promise.all([
      context(mainCfg),
      context(preloadCfg),
      context(rendererPopupCfg),
      context(rendererSettingsCfg)
    ]);
    await copyAssets();
    await Promise.all(ctxs.map(c => c.watch()));
    console.log('[build] watching...');
  } else {
    await Promise.all([
      build(mainCfg),
      build(preloadCfg),
      build(rendererPopupCfg),
      build(rendererSettingsCfg)
    ]);
    await copyAssets();
    console.log('[build] done');
  }
}

run().catch(err => {
  console.error(err);
  process.exit(1);
});
