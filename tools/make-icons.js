// Generates icons/icon{16,32,48,128}.png without external dependencies.
// Usage: node tools/make-icons.js
'use strict';
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const CRC_TABLE = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
};
const encodePng = (size, rgba) => {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
  const raw = Buffer.alloc((size * 4 + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0;
    rgba.copy(raw, y * (size * 4 + 1) + 1, y * size * 4, (y + 1) * size * 4);
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
};

const hex = (h) => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];

// Shapes in a 128-unit coordinate space; later entries paint over earlier ones.
const inRoundRect = (x, y, x0, y0, x1, y1, r) => {
  if (x < x0 || x > x1 || y < y0 || y > y1) return false;
  const cx = Math.min(Math.max(x, x0 + r), x1 - r);
  const cy = Math.min(Math.max(y, y0 + r), y1 - r);
  return (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
};
const shapes = [
  // background tile
  { color: '#1f2a44', test: (x, y) => inRoundRect(x, y, 4, 4, 124, 124, 26) },
  // page with folded corner
  { color: '#ffffff', test: (x, y) => inRoundRect(x, y, 30, 18, 98, 110, 7) && !(x > 76 && y < 18 + (x - 76)) },
  { color: '#c9d2e3', test: (x, y) => x >= 76 && x <= 98 && y >= 18 && y <= 40 && y >= 18 + (x - 76) - 0.5 },
  // Word / Excel / PowerPoint bars
  { color: '#2b6fd6', test: (x, y) => inRoundRect(x, y, 42, 50, 86, 61, 4) },
  { color: '#1e9e5a', test: (x, y) => inRoundRect(x, y, 42, 68, 86, 79, 4) },
  { color: '#e3622f', test: (x, y) => inRoundRect(x, y, 42, 86, 74, 97, 4) },
];

const render = (size) => {
  const ss = 4; // supersampling per axis
  const out = Buffer.alloc(size * size * 4);
  for (let py = 0; py < size; py++) {
    for (let px = 0; px < size; px++) {
      let r = 0, g = 0, b = 0, a = 0;
      for (let sy = 0; sy < ss; sy++) {
        for (let sx = 0; sx < ss; sx++) {
          const x = ((px + (sx + 0.5) / ss) / size) * 128;
          const y = ((py + (sy + 0.5) / ss) / size) * 128;
          let hit = null;
          for (const s of shapes) if (s.test(x, y)) hit = s.color;
          if (hit) {
            const [cr, cg, cb] = hex(hit);
            r += cr; g += cg; b += cb; a += 1;
          }
        }
      }
      const i = (py * size + px) * 4;
      if (a) {
        out[i] = Math.round(r / a);
        out[i + 1] = Math.round(g / a);
        out[i + 2] = Math.round(b / a);
      }
      out[i + 3] = Math.round((a / (ss * ss)) * 255);
    }
  }
  return encodePng(size, out);
};

const dir = path.join(__dirname, '..', 'icons');
fs.mkdirSync(dir, { recursive: true });
for (const size of [16, 32, 48, 128]) {
  fs.writeFileSync(path.join(dir, `icon${size}.png`), render(size));
}
console.log('icons written to', dir);
