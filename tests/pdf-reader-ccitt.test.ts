// E-PDF EP15 — CCITT Group 3 / Group 4 fax decode. The decoder is exercised
// three ways: the run/mode codes are checked to be a valid prefix set; two
// hand-computed bitstreams anchor the run + mode codes independently of the
// encoder; and an in-test encoder (built from the same exported code tables, but
// the standard mode-selection algorithm) round-trips varied bitmaps through the
// real decoder, including the end-to-end image pipeline.

import { readFileSync } from 'node:fs';

import { unzlibSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { prepareImage } from '@/core/images';
import { Ream } from '@/core/converter/ream';
import {
  BLACK_CODES,
  MODE_CODES,
  SHARED_MAKEUP,
  WHITE_CODES,
  decodeCcitt,
} from '@/pdf-reader/ccitt';
import { PdfFile } from '@/pdf-reader/document';
import { decodePdfImage } from '@/pdf-reader/image-decode';
import { dict, name, stream } from '@/pdf/objects';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// --- an independent encoder (mode selection is the standard G4 algorithm) -----

class BitWriter {
  private bits: Array<number> = [];
  str(s: string): void {
    for (const c of s) this.bits.push(c === '1' ? 1 : 0);
  }
  bytes(): Uint8Array {
    const out = new Uint8Array((this.bits.length + 7) >> 3);
    for (let i = 0; i < this.bits.length; i++) if (this.bits[i]) out[i >> 3]! |= 0x80 >> (i & 7);
    return out;
  }
}

const runBits = (codes: ReadonlyArray<readonly [number, string]>): Map<number, string> =>
  new Map([...codes, ...SHARED_MAKEUP].map(([run, s]) => [run, s]));
const WHITE = runBits(WHITE_CODES);
const BLACK = runBits(BLACK_CODES);
const modeBits = (kind: string, d: number): string =>
  MODE_CODES.find(([, m]) => m.kind === kind && m.d === d)![0];

function writeRun(bw: BitWriter, map: Map<number, string>, run: number): void {
  if (run >= 64) {
    const makeup = Math.floor(run / 64) * 64;
    bw.str(map.get(makeup)!);
    run -= makeup;
  }
  bw.str(map.get(run)!);
}

// Changing elements of one bitmap row (1 = black), starting white, with the two
// `width` sentinels the decoder also appends.
function changesOf(bitmap: Uint8Array, y: number, width: number): Array<number> {
  const ch: Array<number> = [];
  let color = 0;
  for (let x = 0; x < width; x++) {
    const p = bitmap[y * width + x]!;
    if (p !== color) {
      ch.push(x);
      color = p;
    }
  }
  ch.push(width, width);
  return ch;
}

const firstGreater = (arr: Array<number>, v: number): number =>
  arr.find((a) => a > v) ?? arr[arr.length - 1]!;

// b1/b2 — must match the decoder's reference-line scan.
function findB(ref: Array<number>, a0: number, color: number, width: number): [number, number] {
  let i = 0;
  while (i < ref.length && ref[i]! <= a0) i++;
  const wantEven = color === 0;
  if (i < ref.length && (i % 2 === 0) !== wantEven) i++;
  return [i < ref.length ? ref[i]! : width, i + 1 < ref.length ? ref[i + 1]! : width];
}

function encodeG4(bitmap: Uint8Array, width: number, height: number): Uint8Array {
  const bw = new BitWriter();
  let ref: Array<number> = [width, width];
  for (let y = 0; y < height; y++) {
    const cur = changesOf(bitmap, y, width);
    let a0 = -1;
    let color = 0;
    while (a0 < width) {
      const a1 = firstGreater(cur, a0);
      const [b1, b2] = findB(ref, a0, color, width);
      if (b2 < a1) {
        bw.str(modeBits('pass', 0));
        a0 = b2;
      } else if (a1 - b1 >= -3 && a1 - b1 <= 3) {
        bw.str(modeBits('vertical', a1 - b1));
        a0 = a1;
        color ^= 1;
      } else {
        const a2 = firstGreater(cur, a1);
        const start = a0 < 0 ? 0 : a0;
        bw.str(modeBits('horizontal', 0));
        writeRun(bw, color === 0 ? WHITE : BLACK, a1 - start);
        writeRun(bw, color === 0 ? BLACK : WHITE, a2 - a1);
        a0 = a2;
      }
    }
    ref = cur;
  }
  return bw.bytes();
}

function encode1D(bitmap: Uint8Array, width: number, height: number): Uint8Array {
  const bw = new BitWriter();
  for (let y = 0; y < height; y++) {
    let x = 0;
    let color = 0;
    while (x < width) {
      let run = 0;
      while (x < width && bitmap[y * width + x] === color) {
        run++;
        x++;
      }
      writeRun(bw, color === 0 ? WHITE : BLACK, run);
      color ^= 1;
    }
  }
  return bw.bytes();
}

// Unpack the decoder's packed bitmap (bit 1 = black) to a 0/1 array.
function unpackBitmap(packed: Uint8Array, width: number, height: number): Uint8Array {
  const rowBytes = (width + 7) >> 3;
  const out = new Uint8Array(width * height);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      out[y * width + x] = (packed[y * rowBytes + (x >> 3)]! >> (7 - (x & 7))) & 1;
    }
  }
  return out;
}

// A deterministic bilevel pattern (1 = black).
function pattern(width: number, height: number): Uint8Array {
  const b = new Uint8Array(width * height);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      b[y * width + x] = (x * 7 + y * 13 + ((x * y) >> 2)) % 3 === 0 ? 1 : 0;
    }
  }
  return b;
}

describe('CCITT fax decode (E-PDF EP15)', () => {
  it('the run and mode codes form a valid prefix set', () => {
    const prefixFree = (codes: ReadonlyArray<string>): boolean =>
      codes.every((a, i) => codes.every((b, j) => i === j || !b.startsWith(a)));
    const white = [...WHITE_CODES, ...SHARED_MAKEUP].map(([, s]) => s);
    const black = [...BLACK_CODES, ...SHARED_MAKEUP].map(([, s]) => s);
    expect(prefixFree(white)).toBe(true);
    expect(prefixFree(black)).toBe(true);
    expect(prefixFree(MODE_CODES.map(([s]) => s))).toBe(true);
    // anchors against the published T.4 / T.6 codes
    expect(new Map(WHITE_CODES).get(0)).toBe('00110101');
    expect(new Map(BLACK_CODES).get(0)).toBe('0000110111');
    expect(MODE_CODES.find(([, m]) => m.kind === 'vertical' && m.d === 0)![0]).toBe('1');
  });

  it('decodes a hand-built all-white Group 4 row (V0)', () => {
    // a single vertical-zero code paints the imaginary white reference line.
    const packed = decodeCcitt(Uint8Array.of(0b10000000), {
      k: -1,
      columns: 8,
      rows: 1,
      byteAlign: false,
    });
    expect(packed && [...packed]).toEqual([0x00]); // no black bits
  });

  it('decodes a hand-built all-black Group 4 row (Horizontal: white 0 + black 8)', () => {
    // 001 (horizontal) · 00110101 (white run 0) · 000101 (black run 8)
    const packed = decodeCcitt(Uint8Array.of(0x26, 0xa2, 0x80), {
      k: -1,
      columns: 8,
      rows: 1,
      byteAlign: false,
    });
    expect(packed && [...packed]).toEqual([0xff]); // all 8 pixels black
  });

  it('round-trips varied bitmaps through Group 4', () => {
    for (const [w, h] of [
      [8, 4],
      [16, 5],
      [40, 12],
      [33, 7],
    ] as const) {
      const bitmap = pattern(w, h);
      const packed = decodeCcitt(encodeG4(bitmap, w, h), {
        k: -1,
        columns: w,
        rows: h,
        byteAlign: false,
      });
      expect(packed).toBeDefined();
      expect([...unpackBitmap(packed!, w, h)]).toEqual([...bitmap]);
    }
  });

  it('round-trips a bitmap through Group 3 one-dimensional', () => {
    const [w, h] = [48, 9];
    const bitmap = pattern(w, h);
    const packed = decodeCcitt(encode1D(bitmap, w, h), {
      k: 0,
      columns: w,
      rows: h,
      byteAlign: false,
    });
    expect(packed).toBeDefined();
    expect([...unpackBitmap(packed!, w, h)]).toEqual([...bitmap]);
  });

  it('declines Group 3 two-dimensional (K > 0)', () => {
    expect(decodeCcitt(Uint8Array.of(0xff), { k: 1, columns: 8, rows: 1, byteAlign: false })).toBe(
      undefined,
    );
  });
});

describe('CCITT image XObject decode (E-PDF EP15)', () => {
  let file: PdfFile;
  beforeAll(async () => {
    const pdf = await Ream.parse(buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>')).convert(
      'pdf',
      { fonts: FONTS },
    );
    file = PdfFile.parse(pdf);
  });

  it('decodes a Group 4 fax image XObject to a DeviceGray PNG', () => {
    const w = 24;
    const h = 16;
    const bitmap = pattern(w, h);
    const xobj = stream(
      {
        Width: w,
        Height: h,
        ColorSpace: name('DeviceGray'),
        BitsPerComponent: 1,
        Filter: name('CCITTFaxDecode'),
        DecodeParms: dict({ K: -1, Columns: w, Rows: h }),
      },
      encodeG4(bitmap, w, h),
    );
    const decoded = decodePdfImage(file, xobj);
    expect(decoded.ok).toBe(true);
    if (!decoded.ok) return;
    expect(decoded.format).toBe('png');
    // The PNG's gray samples: black → 0, white → 255.
    const gray = unzlibSync(prepareImage(decoded.bytes).data);
    const expected = Uint8Array.from(bitmap, (b) => (b ? 0 : 255));
    expect([...gray]).toEqual([...expected]);
  });
});
