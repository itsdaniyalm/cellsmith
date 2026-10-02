// Generates the add-in's PNG icons (a green rounded square with a pixel-art "fx") with no dependencies.
// Run: node scripts/make-icons.mjs
import { mkdirSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { deflateSync } from 'node:zlib';

const OUT = path.resolve(import.meta.dirname, '..', 'public', 'assets');
const GREEN = [0x21, 0x73, 0x46];
const WHITE = [0xff, 0xff, 0xff];

// 5x7 glyphs
const F = ['..###', '.#...', '.#...', '####.', '.#...', '.#...', '.#...'];
const X = ['.....', '#...#', '.#.#.', '..#..', '.#.#.', '#...#', '.....'];

function crc32(buf) {
  let c;
  let crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function png(size) {
  const radius = size * 0.2;
  const scale = Math.max(1, Math.floor((size * 0.62) / 11));
  const gw = 11 * scale;
  const gh = 7 * scale;
  const ox = Math.floor((size - gw) / 2);
  const oy = Math.floor((size - gh) / 2);

  const lit = (x, y) => {
    const gx = x - ox;
    const gy = y - oy;
    if (gx < 0 || gy < 0 || gx >= gw || gy >= gh) return false;
    const col = Math.floor(gx / scale);
    const row = Math.floor(gy / scale);
    if (col < 5) return F[row][col] === '#';
    if (col === 5) return false;
    return X[row][col - 6] === '#';
  };
  const inside = (x, y) => {
    const cx = Math.min(Math.max(x + 0.5, radius), size - radius);
    const cy = Math.min(Math.max(y + 0.5, radius), size - radius);
    return (x + 0.5 - cx) ** 2 + (y + 0.5 - cy) ** 2 <= radius ** 2;
  };

  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) {
    const row = y * (size * 4 + 1);
    raw[row] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const o = row + 1 + x * 4;
      if (!inside(x, y)) continue; // transparent
      const [r, g, b] = lit(x, y) ? WHITE : GREEN;
      raw[o] = r;
      raw[o + 1] = g;
      raw[o + 2] = b;
      raw[o + 3] = 255;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

mkdirSync(OUT, { recursive: true });
for (const size of [16, 32, 64, 80]) {
  writeFileSync(path.join(OUT, `icon-${size}.png`), png(size));
  console.log(`wrote public/assets/icon-${size}.png`);
}
