// XLS-1 — the Compound File Binary (OLE2) reader, the container the legacy
// binary Office formats live in. A synthesized CFB (build-cfb) round-trips
// through the reader: a big stream via the regular FAT, a small stream via the
// mini stream + mini-FAT, an empty stream, and case-insensitive lookup.

import { describe, expect, it } from 'vitest';

import { buildCfb } from './fixtures/build-cfb';
import { isCfb, openCfb } from '@/core/ole/cfb';

// A deterministic byte pattern of a given length.
const pattern = (n: number, seed: number): Uint8Array =>
  Uint8Array.from({ length: n }, (_, i) => (i * 7 + seed) % 251);

describe('CFB / OLE2 reader (XLS-1)', () => {
  it('reads a big stream through the regular FAT', () => {
    const data = pattern(5000, 1); // ≥ 4096 cutoff → regular sectors
    const cfb = openCfb(buildCfb([{ name: 'Workbook', data }]));
    expect(cfb.readStream('Workbook')).toEqual(data);
  });

  it('reads a small stream through the mini stream / mini-FAT', () => {
    const data = pattern(100, 2); // < 4096 → mini sectors
    const cfb = openCfb(buildCfb([{ name: 'Book', data }]));
    expect(cfb.readStream('Book')).toEqual(data);
  });

  it('reads several streams of mixed sizes from one container', () => {
    const big = pattern(9000, 3);
    const small = pattern(64 * 3 + 5, 4); // spans 4 mini sectors
    const tiny = pattern(3, 5);
    const cfb = openCfb(
      buildCfb([
        { name: 'Workbook', data: big },
        { name: 'SmallStr', data: small },
        { name: 'Tiny', data: tiny },
        { name: 'Empty', data: new Uint8Array(0) },
      ]),
    );
    expect(cfb.readStream('Workbook')).toEqual(big);
    expect(cfb.readStream('SmallStr')).toEqual(small);
    expect(cfb.readStream('Tiny')).toEqual(tiny);
    expect(cfb.readStream('Empty')).toEqual(new Uint8Array(0));
  });

  it('looks streams up case-insensitively and reports missing ones', () => {
    const cfb = openCfb(buildCfb([{ name: 'Workbook', data: pattern(200, 6) }]));
    expect(cfb.hasStream('workbook')).toBe(true);
    expect(cfb.hasStream('WORKBOOK')).toBe(true);
    expect(cfb.hasStream('nope')).toBe(false);
    expect(cfb.readStream('nope')).toBeUndefined();
  });

  it('exposes the directory entries (root + streams)', () => {
    const cfb = openCfb(buildCfb([{ name: 'Workbook', data: pattern(50, 7) }]));
    expect(cfb.entries.some((e) => e.type === 'root')).toBe(true);
    expect(cfb.entries.find((e) => e.name === 'Workbook')?.type).toBe('stream');
  });

  it('sniffs the CFB magic', () => {
    expect(isCfb(buildCfb([{ name: 'Book', data: pattern(10, 8) }]))).toBe(true);
    expect(isCfb(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false); // ZIP
    expect(isCfb(new Uint8Array(10))).toBe(false);
  });
});
