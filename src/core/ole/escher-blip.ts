// MS-ODRAW §2.2.23 OfficeArtBlip — the picture inside an Escher BLIP record,
// shared by the `.ppt` picture store and the `.xls` BLIP store because both
// reach the same record by different routes.
//
// A raster blip (PNG, JPEG) is its own bytes after a UID or two, and scanning
// for the signature finds it. A METAFILE is not: EMF, WMF and PICT carry an
// OfficeArtMetafileHeader (§2.2.31) stating the uncompressed size and whether
// what follows is deflated — which it almost always is, so nothing downstream
// ever saw a metafile at all. Both readers dropped every one of them: a `.ppt`
// slide whose whole content was an EMF drew the shape's fill colour instead, a
// blue rectangle where the reference has a diagram.
//
// The layout replays a metafile from its raw bytes (an EMF/WMF resource takes
// no raster name and is drawn as primitives), so unpacking here is all it takes.

import { unzlibSync } from 'fflate';

// §2.2.23 blip record types.
const BLIP_EMF = 0xf01a;
const BLIP_WMF = 0xf01b;
const BLIP_PICT = 0xf01c;
const BLIP_JPEG = 0xf01d;
const BLIP_PNG = 0xf01e;
const BLIP_DIB = 0xf01f;
const BLIP_TIFF = 0xf029;
const BLIP_JPEG_CMYK = 0xf02a;

// §2.2.31 — the metafile header that precedes the (usually deflated) data.
const METAFILE_HEADER_BYTES = 34;
const COMPRESSION_DEFLATE = 0x00;

// The recInstance values that mean the blip carries TWO UIDs rather than one
// (§2.2.24–§2.2.30: each type has a one-UID and a two-UID variant).
const TWO_UID_INSTANCES = new Set([0x3d5, 0x217, 0x543, 0x6e1, 0x6e3, 0x7a9, 0x6e5, 0x1e1]);

/** The signatures of the rasters a blip may hold, for the scan fallback. */
const RASTER_MAGICS: ReadonlyArray<ReadonlyArray<number>> = [
  [0x89, 0x50, 0x4e, 0x47], // PNG
  [0xff, 0xd8, 0xff], // JPEG
  [0x47, 0x49, 0x46, 0x38], // GIF
  [0x42, 0x4d], // BMP
];

/**
 * The picture bytes inside one Escher BLIP record — a raster as it stands, or a
 * metafile unpacked out of its header.
 *
 * @param type     The blip record type (`0xF01A`…`0xF02A`).
 * @param instance The record's `recInstance`, which says how many UIDs precede
 *                 the payload.
 * @param data     The record's body, without the 8-byte header.
 * @returns The bytes to hand the renderer, or `undefined` when the blip holds
 *          nothing this can read.
 */
export function readEscherBlip(
  type: number,
  instance: number,
  data: Uint8Array,
): Uint8Array | undefined {
  if (type === BLIP_EMF || type === BLIP_WMF || type === BLIP_PICT) {
    return readMetafileBlip(instance, data);
  }
  if (
    type === BLIP_JPEG ||
    type === BLIP_PNG ||
    type === BLIP_DIB ||
    type === BLIP_TIFF ||
    type === BLIP_JPEG_CMYK
  ) {
    return scanRaster(data);
  }
  // An unknown type still gets the scan: a producer that writes a raster under
  // a type this does not model should not lose its picture over it.
  return scanRaster(data);
}

// §2.2.24/§2.2.25 — UID(s), then the metafile header, then the data. The header
// states both sizes, so a truncated or mis-sized record is caught before it
// reaches the inflater.
function readMetafileBlip(instance: number, data: Uint8Array): Uint8Array | undefined {
  const uidBytes = TWO_UID_INSTANCES.has(instance) ? 32 : 16;
  const head = uidBytes + METAFILE_HEADER_BYTES;
  if (data.length < head) return undefined;
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const uncompressed = view.getUint32(uidBytes, true); // cbSize
  const compressed = view.getUint32(uidBytes + 24, true); // cbSave
  const compression = data[uidBytes + 32];
  const payload = data.subarray(head, Math.min(data.length, head + (compressed || Infinity)));
  if (payload.length === 0) return undefined;
  if (compression !== COMPRESSION_DEFLATE) return payload; // 0xFE — stored as is
  try {
    const out = unzlibSync(payload, uncompressed > 0 ? { out: new Uint8Array(uncompressed) } : {});
    return out.length > 0 ? out : undefined;
  } catch {
    return undefined; // a blip we cannot unpack is a picture we do not draw
  }
}

// The raster's signature, within the short window the UID(s) and any BSE prefix
// occupy. Robust against the variants rather than counting bytes for each.
function scanRaster(data: Uint8Array): Uint8Array | undefined {
  const limit = Math.min(data.length, 80);
  for (let off = 0; off < limit; off++) {
    for (const magic of RASTER_MAGICS) {
      if (matches(data, off, magic)) return data.subarray(off);
    }
  }
  return undefined;
}

function matches(d: Uint8Array, off: number, magic: ReadonlyArray<number>): boolean {
  if (off + magic.length > d.length) return false;
  for (let i = 0; i < magic.length; i++) if (d[off + i] !== magic[i]) return false;
  return true;
}
