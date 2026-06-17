// Office Drawing (Escher / MS-ODRAW) reader for `.xls` (XLS-5) — the binary
// predecessor of DrawingML that carries a workbook's pictures and shapes. The
// workbook globals hold one MSODrawingGroup record whose BLIP store is the image
// pool; each sheet holds an MSODrawing record whose shape containers reference a
// pool entry (the `pib` property) and carry a cell anchor. This module walks the
// nested record tree and pulls out the pictures; the BIFF reader maps them onto
// the SheetImageRef model the renderer already draws.
//
// Records are `[verInstance:u16][type:u16][len:u32]` then `len` bytes; a record
// is a container (children follow) when its version nibble is 0xF.

const FBT_BSTORE_CONTAINER = 0xf001;
const FBT_BSE = 0xf007;
const FBT_SP_CONTAINER = 0xf004;
const FBT_OPT = 0xf00b;
const FBT_CLIENT_ANCHOR = 0xf010;
const PROP_PIB = 0x0104; // OPT property: 1-based index into the BLIP store

const MAX_DEPTH = 24;

interface EscherRecord {
  readonly type: number;
  readonly instance: number;
  readonly isContainer: boolean;
  readonly data: Uint8Array;
}

// A picture shape: the BLIP store index it references plus its cell anchor.
export interface EscherPicture {
  readonly blipIndex: number; // 1-based (pib)
  readonly anchor?: EscherAnchor;
}

export interface EscherAnchor {
  readonly col1: number;
  readonly row1: number;
  readonly col2: number;
  readonly row2: number;
  readonly dx1: number;
  readonly dy1: number;
  readonly dx2: number;
  readonly dy2: number;
}

// Iterate the Escher records at one nesting level.
function* records(d: Uint8Array): Generator<EscherRecord> {
  let off = 0;
  while (off + 8 <= d.length) {
    const verInstance = u16(d, off);
    const type = u16(d, off + 2);
    const len = u32(d, off + 4);
    const start = off + 8;
    const end = Math.min(d.length, start + len);
    yield {
      type,
      instance: verInstance >> 4,
      isContainer: (verInstance & 0x0f) === 0x0f,
      data: d.subarray(start, end),
    };
    if (end <= off) break; // zero-length guard
    off = end;
  }
}

// The BLIP store: one entry per BSE (so a `pib` index lines up), holding the
// extracted image bytes — or undefined for an entry we cannot read (a metafile
// blip, an empty slot), which keeps the indices aligned.
export function parseBlipStore(msoDrawingGroup: Uint8Array): Array<Uint8Array | undefined> {
  const store = findContainer(msoDrawingGroup, FBT_BSTORE_CONTAINER, 0);
  if (!store) return [];
  const out: Array<Uint8Array | undefined> = [];
  for (const r of records(store)) {
    if (r.type === FBT_BSE) out.push(findImageBytes(r.data));
  }
  return out;
}

// Depth-first search for the first container of `type`.
function findContainer(d: Uint8Array, type: number, depth: number): Uint8Array | undefined {
  if (depth > MAX_DEPTH) return undefined;
  for (const r of records(d)) {
    if (r.type === type) return r.data;
    if (r.isContainer) {
      const found = findContainer(r.data, type, depth + 1);
      if (found) return found;
    }
  }
  return undefined;
}

// The picture shapes in a sheet's MSODrawing: every SpContainer that carries a
// `pib` (a picture-fill reference) with its anchor.
export function parseSheetPictures(msoDrawing: Uint8Array): Array<EscherPicture> {
  const out: Array<EscherPicture> = [];
  collectPictures(msoDrawing, out, 0);
  return out;
}

function collectPictures(d: Uint8Array, out: Array<EscherPicture>, depth: number): void {
  if (depth > MAX_DEPTH) return;
  for (const r of records(d)) {
    if (r.type === FBT_SP_CONTAINER) {
      let blipIndex: number | undefined;
      let anchor: EscherAnchor | undefined;
      for (const child of records(r.data)) {
        if (child.type === FBT_OPT) blipIndex = optProperty(child.data, child.instance, PROP_PIB);
        else if (child.type === FBT_CLIENT_ANCHOR) anchor = parseAnchor(child.data);
      }
      if (blipIndex !== undefined && blipIndex > 0) {
        out.push({ blipIndex, ...(anchor ? { anchor } : {}) });
      }
    } else if (r.isContainer) {
      collectPictures(r.data, out, depth + 1);
    }
  }
}

// §2.3.7.2 OfficeArtFOPT — `instance` is the property count; each entry is a
// 2-byte id (low 14 bits) + 4-byte value. Returns the simple value of `wantId`.
function optProperty(d: Uint8Array, count: number, wantId: number): number | undefined {
  for (let i = 0; i < count && i * 6 + 6 <= d.length; i++) {
    const id = u16(d, i * 6);
    if ((id & 0x3fff) === wantId && (id & 0x8000) === 0) return u32(d, i * 6 + 2);
  }
  return undefined;
}

// §2.5.193 OfficeArtClientAnchorSheet — the from/to cell anchor (each cell index
// is a u16, each offset a u16 in 1/1024 of the cell).
function parseAnchor(d: Uint8Array): EscherAnchor | undefined {
  if (d.length < 18) return undefined;
  return {
    col1: u16(d, 2),
    dx1: u16(d, 4),
    row1: u16(d, 6),
    dy1: u16(d, 8),
    col2: u16(d, 10),
    dx2: u16(d, 12),
    row2: u16(d, 14),
    dy2: u16(d, 16),
  };
}

// Image magics, scanned for inside a BSE's bytes — robust against the variable
// BLIP header (1 vs 2 UIDs, metafile sub-headers). The renderer sniffs the
// format from these same bytes. Metafiles (EMF/WMF) carry a sub-header before
// their payload, so they are not extracted by a raw scan and are skipped.
const MAGICS: ReadonlyArray<ReadonlyArray<number>> = [
  [0x89, 0x50, 0x4e, 0x47], // PNG
  [0xff, 0xd8, 0xff], // JPEG
  [0x47, 0x49, 0x46, 0x38], // GIF
  [0x42, 0x4d], // BMP
  [0x49, 0x49, 0x2a, 0x00], // TIFF (LE)
  [0x4d, 0x4d, 0x00, 0x2a], // TIFF (BE)
];

function findImageBytes(bse: Uint8Array): Uint8Array | undefined {
  // The header (BSE prefix + BLIP UID) is short; scan a bounded window.
  const limit = Math.min(bse.length, 80);
  for (let off = 0; off < limit; off++) {
    for (const magic of MAGICS) {
      if (matches(bse, off, magic)) return bse.subarray(off);
    }
  }
  return undefined;
}

function matches(d: Uint8Array, off: number, magic: ReadonlyArray<number>): boolean {
  if (off + magic.length > d.length) return false;
  for (let i = 0; i < magic.length; i++) if (d[off + i] !== magic[i]) return false;
  return true;
}

function u16(d: Uint8Array, off: number): number {
  return (d[off] ?? 0) | ((d[off + 1] ?? 0) << 8);
}
function u32(d: Uint8Array, off: number): number {
  return (
    ((d[off] ?? 0) |
      ((d[off + 1] ?? 0) << 8) |
      ((d[off + 2] ?? 0) << 16) |
      ((d[off + 3] ?? 0) << 24)) >>>
    0
  );
}
