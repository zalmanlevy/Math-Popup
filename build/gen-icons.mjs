// One-time / occasional icon generator for the web/PWA build. Produces square PWA
// icons from assets/icon.png (which is NON-square, 1616x1404 — must be padded to a
// square canvas or it would distort). Output is committed under src/web/icons/ and
// copied verbatim by build/build-web.mjs, so the Vercel build needs NO native deps.
//
// Run only when the source icon changes:  npm i -D sharp && node build/gen-icons.mjs
import sharp from 'sharp';
import { mkdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const src = join(root, 'assets', 'icon.png');
const outDir = join(root, 'src', 'web', 'icons');
const bg = { r: 250, g: 250, b: 250, alpha: 1 };          // #fafafa — matches the manifest background
const transparent = { r: 0, g: 0, b: 0, alpha: 0 };

await mkdir(outDir, { recursive: true });

const fit = async (size, background) =>
  sharp(src).resize(size, size, { fit: 'contain', background }).png().toBuffer();

// "any" icons — artwork fit on transparent, full size.
for (const size of [192, 512]) {
  await sharp(await fit(size, transparent)).toFile(join(outDir, `icon-${size}.png`));
}

// Maskable — artwork inside the ~80% safe zone, on an opaque square (no transparency).
await sharp({ create: { width: 512, height: 512, channels: 4, background: bg } })
  .composite([{ input: await fit(410, transparent), gravity: 'center' }])
  .png().toFile(join(outDir, 'maskable-512.png'));

// apple-touch — full-bleed opaque square (iOS rounds the corners itself).
await sharp({ create: { width: 180, height: 180, channels: 4, background: bg } })
  .composite([{ input: await fit(160, transparent), gravity: 'center' }])
  .png().toFile(join(outDir, 'apple-touch-icon-180.png'));

console.log('[gen-icons] wrote icon-192, icon-512, maskable-512, apple-touch-icon-180 to', outDir);
