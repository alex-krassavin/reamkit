// Build a minimal valid RGBA PNG of the given dimensions for tests. Each
// pixel takes a 32-bit ABGR sample (well, 4 bytes R G B A); the PNG is
// uncompressed-friendly so the test doesn't need a real PNG asset on disk.

import { zlibSync } from 'fflate';

const PNG_SIGNATURE = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

export function buildTinyPng(
  width: number,
  height: number,
  pixel: readonly [number, number, number, number],
): Uint8Array {
  // IHDR — 13 bytes
  const ihdrData = new Uint8Array(13);
  writeU32BE(ihdrData, 0, width);
  writeU32BE(ihdrData, 4, height);
  ihdrData[8] = 8; // bit depth
  ihdrData[9] = 6; // color type RGBA
  ihdrData[10] = 0; // compression
  ihdrData[11] = 0; // filter
  ihdrData[12] = 0; // interlace
  const ihdrChunk = makeChunk('IHDR', ihdrData);

  // IDAT — scanlines = (1 byte filter + width*4 RGBA) per row
  const raw = new Uint8Array(height * (1 + width * 4));
  for (let y = 0; y < height; y++) {
    const off = y * (1 + width * 4);
    raw[off] = 0; // filter type None
    for (let x = 0; x < width; x++) {
      const p = off + 1 + x * 4;
      raw[p] = pixel[0];
      raw[p + 1] = pixel[1];
      raw[p + 2] = pixel[2];
      raw[p + 3] = pixel[3];
    }
  }
  const idatData = zlibSync(raw);
  const idatChunk = makeChunk('IDAT', idatData);

  const iendChunk = makeChunk('IEND', new Uint8Array(0));

  return concat(PNG_SIGNATURE, ihdrChunk, idatChunk, iendChunk);
}

function makeChunk(type: string, data: Uint8Array): Uint8Array {
  const lenBytes = new Uint8Array(4);
  writeU32BE(lenBytes, 0, data.length);
  const typeBytes = new Uint8Array(4);
  for (let i = 0; i < 4; i++) typeBytes[i] = type.charCodeAt(i);
  const crcInput = concat(typeBytes, data);
  const crc = crc32(crcInput);
  const crcBytes = new Uint8Array(4);
  writeU32BE(crcBytes, 0, crc);
  return concat(lenBytes, typeBytes, data, crcBytes);
}

function writeU32BE(buf: Uint8Array, offset: number, value: number): void {
  buf[offset] = (value >>> 24) & 0xff;
  buf[offset + 1] = (value >>> 16) & 0xff;
  buf[offset + 2] = (value >>> 8) & 0xff;
  buf[offset + 3] = value & 0xff;
}

function concat(...parts: Array<Uint8Array>): Uint8Array {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

let crcTable: Uint32Array | null = null;
function crc32(data: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      crcTable[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const b of data) {
    crc = (crcTable[(crc ^ b) & 0xff]! ^ (crc >>> 8)) >>> 0;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
