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

  it('ignores garbage in the reserved high 4 bytes of a v3 stream size (MS-CFB §2.6.1)', () => {
    // Word and other Office writers leave non-zero junk in the high dword of a
    // v3 directory entry's 64-bit size field. A reader that uses the full value
    // sees an absurd size (~1.8e19 B) and rejects the file with "stream exceeds
    // size limit"; a correct reader masks it to the low 4 bytes. Cover both the
    // regular-FAT (big) and mini-stream (small) read paths.
    const big = pattern(5000, 11); // ≥ 4096 cutoff
    const small = pattern(1000, 12); // < 4096 cutoff
    const cfb = openCfb(
      buildCfb(
        [
          { name: 'WordDocument', data: big },
          { name: '1Table', data: small },
        ],
        { garbageSizeHighDword: true },
      ),
    );
    expect(cfb.readStream('WordDocument')).toEqual(big);
    expect(cfb.readStream('1Table')).toEqual(small);
  });

  it('prefers a top-level stream over an embedded-storage stream of the same name', () => {
    // Models a .doc that embeds an OLE object: both the main document and the
    // embedded object own a `WordDocument` stream, and the embedded one comes
    // first in the directory. A flat first-wins reader returns the embedded
    // stream — and the .doc parse then comes up empty — so the tree-aware reader
    // must return the main document's (the root storage's direct child).
    const mainWd = pattern(200, 21);
    const embeddedWd = pattern(120, 22);
    const cfb = openCfb(
      buildCfb(
        [
          { name: 'WordDocument', data: mainWd },
          { name: '1Table', data: pattern(80, 23) },
        ],
        {
          storages: [{ name: 'ObjectPool', streams: [{ name: 'WordDocument', data: embeddedWd }] }],
        },
      ),
    );
    expect(cfb.readStream('WordDocument')).toEqual(mainWd);
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
