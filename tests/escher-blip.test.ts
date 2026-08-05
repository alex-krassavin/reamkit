// MS-ODRAW §2.2.23 — the picture inside an Escher BLIP record, shared by the
// `.ppt` picture store and the `.xls` BLIP store.
//
// A raster is found by scanning for its signature. A METAFILE is not: it carries
// an OfficeArtMetafileHeader (§2.2.31) and its data is almost always deflated,
// so a scan walks straight past it — which is why every EMF and WMF in a legacy
// file used to be dropped, and a slide whose whole content was one drew its
// shape's fill colour instead.

import { zlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { readEscherBlip } from '@/core/ole/escher-blip';

const BLIP_EMF = 0xf01a; // EMF, WMF and PICT share the header this exercises
const BLIP_PNG = 0xf01e;

// A minimal standard WMF: type 1, a 9-word header.
const WMF = Uint8Array.from([1, 0, 9, 0, 0, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
const PNG = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

/** A metafile blip: one UID, the §2.2.31 header, then the data. */
function metafileBlip(data: Uint8Array, uncompressed: number, compression: number): Uint8Array {
  const out = new Uint8Array(16 + 34 + data.length);
  const v = new DataView(out.buffer);
  v.setUint32(16, uncompressed, true); // cbSize
  v.setUint32(16 + 24, data.length, true); // cbSave
  out[16 + 32] = compression;
  out[16 + 33] = 0xfe; // filter
  out.set(data, 50);
  return out;
}

describe('an Escher BLIP', () => {
  it('inflates a metafile stored deflated, which is how they all are', () => {
    const packed = zlibSync(WMF);
    const blip = metafileBlip(packed, WMF.length, 0x00);
    expect(readEscherBlip(BLIP_EMF, 0x3d4, blip)).toEqual(WMF);
  });

  it('takes a metafile stored plainly (compression 0xFE) as it stands', () => {
    expect(readEscherBlip(BLIP_EMF, 0x3d4, metafileBlip(WMF, WMF.length, 0xfe))).toEqual(WMF);
  });

  it('skips the second UID when the instance says there is one', () => {
    const packed = zlibSync(WMF);
    const blip = new Uint8Array(16 + metafileBlip(packed, WMF.length, 0x00).length);
    blip.set(metafileBlip(packed, WMF.length, 0x00), 16); // a second UID in front
    expect(readEscherBlip(BLIP_EMF, 0x3d5, blip)).toEqual(WMF);
  });

  it('finds a raster by its signature, past whatever UIDs precede it', () => {
    const blip = new Uint8Array(16 + PNG.length);
    blip.set(PNG, 16);
    expect(readEscherBlip(BLIP_PNG, 0x6e0, blip)).toEqual(PNG);
  });

  it('gives nothing for bytes it cannot unpack, rather than garbage', () => {
    expect(readEscherBlip(BLIP_EMF, 0x3d4, new Uint8Array(10))).toBeUndefined();
    // A header that promises deflated data over bytes that are not.
    expect(
      readEscherBlip(BLIP_EMF, 0x3d4, metafileBlip(Uint8Array.from([1, 2, 3, 4]), 99, 0x00)),
    ).toBeUndefined();
    expect(readEscherBlip(BLIP_PNG, 0x6e0, new Uint8Array(40))).toBeUndefined();
  });
});
