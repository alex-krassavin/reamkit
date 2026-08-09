// ITU-T T.81 — the baseline JPEG decoder. Ream embeds a JPEG verbatim, so the
// only reason to take one apart is to CHANGE it: an /SMask (ISO 32000-1
// §8.9.5.4) is a second image supplying the alpha, and a JPEG has nowhere to
// put alpha. The fixtures below are hand-assembled byte for byte, because a
// baseline JPEG small enough to read is small enough to write.

import { describe, expect, it } from 'vitest';

import { decodeJpeg } from '@/pdf-reader/jpeg';

const seg = (marker: number, body: ReadonlyArray<number>): Array<number> => [
  0xff,
  marker,
  (body.length + 2) >> 8,
  (body.length + 2) & 0xff,
  ...body,
];

/** One code of length 1 for `symbol` — the smallest legal Huffman table. */
const oneCodeTable = (tcth: number, symbol: number): Array<number> => [
  tcth,
  1,
  ...new Array<number>(15).fill(0), // BITS: a single code, of length 1
  symbol, // HUFFVAL
];

/**
 * An 8x8 grayscale baseline JPEG of one flat tone.
 *
 * The single block codes a DC difference of 15 and no AC at all. With every
 * quantizer at 8 the coefficient is 120, and §A.3.3's inverse DCT of a DC-only
 * block spreads it evenly: 120 · (1/(2√2))² = 15, level-shifted to 143.
 */
function flatGrayJpeg(): Uint8Array {
  const bits = 0b0111_1011; // DC code `0`, magnitude `1111`, AC EOB `0`, padding
  return new Uint8Array([
    0xff,
    0xd8, // SOI
    ...seg(0xdb, [0x00, ...new Array<number>(64).fill(8)]), // DQT: table 0, all 8
    ...seg(0xc0, [8, 0, 8, 0, 8, 1, 1, 0x11, 0]), // SOF0: 8x8, one component
    ...seg(0xc4, oneCodeTable(0x00, 0x04)), // DHT DC0: symbol 4 (a 4-bit magnitude)
    ...seg(0xc4, oneCodeTable(0x10, 0x00)), // DHT AC0: symbol 0 (EOB)
    ...seg(0xda, [1, 1, 0x00, 0x00, 0x3f, 0x00]), // SOS
    bits,
    0xff,
    0xd9, // EOI
  ]);
}

describe('baseline JPEG decoder (ITU-T T.81)', () => {
  it('decodes a flat 8x8 grayscale block', () => {
    const out = decodeJpeg(flatGrayJpeg());
    expect(out).toBeDefined();
    expect(out!.width).toBe(8);
    expect(out!.height).toBe(8);
    expect(out!.components).toBe(1);
    expect(out!.samples).toHaveLength(64);
    // Every sample is the same tone, and it is the one the arithmetic predicts.
    expect([...new Set(out!.samples)]).toEqual([143]);
  });

  it('declines a progressive scan rather than guessing at it', () => {
    // SOF2. The caller keeps carrying the JPEG through untouched.
    const progressive = flatGrayJpeg();
    const sof = progressive.indexOf(0xc0, 2);
    progressive[sof] = 0xc2;
    expect(decodeJpeg(progressive)).toBeUndefined();
  });

  it('declines bytes that are not a JPEG at all', () => {
    expect(decodeJpeg(new Uint8Array([0x89, 0x50, 0x4e, 0x47]))).toBeUndefined();
    expect(decodeJpeg(new Uint8Array())).toBeUndefined();
  });

  it('declines a truncated stream instead of throwing', () => {
    const cut = flatGrayJpeg().slice(0, 40);
    expect(() => decodeJpeg(cut)).not.toThrow();
  });
});
