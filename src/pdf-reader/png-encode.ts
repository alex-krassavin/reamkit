// E-PDF EP6 — a minimal PNG encoder (RFC 2083). The PDF reader decodes image
// XObjects to raw 8-bit samples (image-decode.ts); this wraps them back into a
// PNG file so the FlowDoc resource store carries a format every writer already
// embeds — the HTML data-URI path and the docx media part both go through
// `detectImageFormat`, which recognises this output. Only what the reader needs
// is implemented: 8-bit depth, colour types 0/2/4/6, no interlacing, a single
// IDAT and filter-none scanlines (fflate's deflate still compresses them well).

import { zlibSync } from 'fflate';

/** A PNG colour type: `gray` = 1ch, `rgb` = 3ch, `gray-alpha` = 2ch, `rgba` = 4ch. */
export type PngColor = 'gray' | 'rgb' | 'gray-alpha' | 'rgba';

const COLOR_TYPE: Record<PngColor, number> = { gray: 0, rgb: 2, 'gray-alpha': 4, rgba: 6 };
const CHANNELS: Record<PngColor, number> = { gray: 1, rgb: 3, 'gray-alpha': 2, rgba: 4 };

const SIGNATURE = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/**
 * Wrap raw 8-bit samples into a minimal PNG file (RFC 2083): a single `IDAT`
 * with filter-none scanlines, no interlacing, colour types 0/2/4/6. Produces a
 * format every writer already embeds — `detectImageFormat` recognises this
 * output (the HTML data-URI and docx media paths both rely on it).
 *
 * @param samples Row-major, 8-bit, `CHANNELS[color]` interleaved values per pixel.
 */
export function encodePng(
  width: number,
  height: number,
  color: PngColor,
  samples: Uint8Array,
): Uint8Array {
  const channels = CHANNELS[color];
  const stride = width * channels;
  // Prepend a filter byte (0 = None) to each scanline — the PNG image data layout.
  const raw = new Uint8Array(height * (stride + 1));
  for (let y = 0; y < height; y++) {
    const dst = y * (stride + 1);
    raw[dst] = 0;
    raw.set(samples.subarray(y * stride, y * stride + stride), dst + 1);
  }
  const idat = zlibSync(raw);

  const ihdr = new Uint8Array(13);
  writeU32(ihdr, 0, width);
  writeU32(ihdr, 4, height);
  ihdr[8] = 8; // bit depth
  ihdr[9] = COLOR_TYPE[color]; // colour type
  ihdr[10] = 0; // compression method (deflate)
  ihdr[11] = 0; // filter method
  ihdr[12] = 0; // interlace method (none)

  return concat([
    SIGNATURE,
    chunk('IHDR', ihdr),
    chunk('IDAT', idat),
    chunk('IEND', new Uint8Array(0)),
  ]);
}

// A PNG chunk: 4-byte length, 4-byte type, data, 4-byte CRC (over type+data).
function chunk(type: string, data: Uint8Array): Uint8Array {
  const typeBytes = Uint8Array.from([
    type.charCodeAt(0),
    type.charCodeAt(1),
    type.charCodeAt(2),
    type.charCodeAt(3),
  ]);
  const out = new Uint8Array(12 + data.length);
  writeU32(out, 0, data.length);
  out.set(typeBytes, 4);
  out.set(data, 8);
  const crcInput = new Uint8Array(4 + data.length);
  crcInput.set(typeBytes, 0);
  crcInput.set(data, 4);
  writeU32(out, 8 + data.length, crc32(crcInput));
  return out;
}

function writeU32(buf: Uint8Array, offset: number, value: number): void {
  buf[offset] = (value >>> 24) & 0xff;
  buf[offset + 1] = (value >>> 16) & 0xff;
  buf[offset + 2] = (value >>> 8) & 0xff;
  buf[offset + 3] = value & 0xff;
}

function concat(parts: ReadonlyArray<Uint8Array>): Uint8Array {
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

// CRC-32 (ISO 3309 / PNG Annex D), lazily-tabulated, polynomial 0xEDB88320.
let crcTable: Uint32Array | undefined;

function crc32(bytes: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (const b of bytes) crc = crcTable[(crc ^ b) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
