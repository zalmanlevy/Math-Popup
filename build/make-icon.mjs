// Generates a simple 32x32 PNG tray icon: a rounded square with a bold "=" glyph.
// No external deps — writes raw PNG with deflate from zlib.
import { writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync, crc32 } from 'node:zlib';

const __dirname = dirname(fileURLToPath(import.meta.url));
const out = join(__dirname, '..', 'assets', 'icon.png');

const W = 32, H = 32;

// Build pixel buffer (RGBA).
const pixels = Buffer.alloc(W * H * 4);

function setPixel(x, y, r, g, b, a) {
  if (x < 0 || y < 0 || x >= W || y >= H) return;
  const i = (y * W + x) * 4;
  pixels[i] = r; pixels[i + 1] = g; pixels[i + 2] = b; pixels[i + 3] = a;
}

function fillRoundedRect(x0, y0, w, h, r, color) {
  for (let y = y0; y < y0 + h; y++) {
    for (let x = x0; x < x0 + w; x++) {
      // Distance from corner to determine rounding
      let inside = true;
      const cx = x < x0 + r ? x0 + r : (x >= x0 + w - r ? x0 + w - r - 1 : x);
      const cy = y < y0 + r ? y0 + r : (y >= y0 + h - r ? y0 + h - r - 1 : y);
      const dx = x - cx;
      const dy = y - cy;
      if (dx * dx + dy * dy > r * r) inside = false;
      if (inside) setPixel(x, y, color[0], color[1], color[2], color[3]);
    }
  }
}

function fillRect(x0, y0, w, h, color) {
  for (let y = y0; y < y0 + h; y++) {
    for (let x = x0; x < x0 + w; x++) {
      setPixel(x, y, color[0], color[1], color[2], color[3]);
    }
  }
}

// Background: rounded rect, dark slate blue
fillRoundedRect(1, 1, 30, 30, 6, [37, 99, 235, 255]);
// Highlight pass for a softer top edge
for (let x = 4; x < 28; x++) setPixel(x, 2, 96, 165, 250, 200);

// Two horizontal bars to form a stylized "="
fillRect(8, 12, 16, 3, [255, 255, 255, 255]);
fillRect(8, 18, 16, 3, [255, 255, 255, 255]);

// PNG encoding ----------------------------------------------------------
function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const typeBuf = Buffer.from(type, 'ascii');
  const crcInput = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(crcInput) >>> 0, 0);
  return Buffer.concat([len, typeBuf, data, crc]);
}

const SIG = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);

const ihdr = Buffer.alloc(13);
ihdr.writeUInt32BE(W, 0);
ihdr.writeUInt32BE(H, 4);
ihdr[8] = 8;   // bit depth
ihdr[9] = 6;   // color type RGBA
ihdr[10] = 0;  // compression
ihdr[11] = 0;  // filter
ihdr[12] = 0;  // interlace

// Filter byte = 0 per scanline.
const stride = W * 4;
const rawScanlines = Buffer.alloc((stride + 1) * H);
for (let y = 0; y < H; y++) {
  rawScanlines[y * (stride + 1)] = 0;
  pixels.copy(rawScanlines, y * (stride + 1) + 1, y * stride, y * stride + stride);
}
const idat = deflateSync(rawScanlines);

const png = Buffer.concat([
  SIG,
  chunk('IHDR', ihdr),
  chunk('IDAT', idat),
  chunk('IEND', Buffer.alloc(0))
]);

if (!existsSync(dirname(out))) mkdirSync(dirname(out), { recursive: true });
writeFileSync(out, png);
console.log('[make-icon] wrote', out, 'bytes:', png.length);
