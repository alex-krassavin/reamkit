// ISO/IEC 14496-22 — TrueType glyph subsetting.
//
// Strategy: keep all tables intact (cmap, hmtx, maxp.numGlyphs, etc.) so the
// CIDFontType2 /CIDToGIDMap /Identity contract holds. Only the glyf table is
// rebuilt with unused glyphs replaced by zero-length entries; loca is updated
// to match. The composite-glyph closure is taken (composites reference other
// glyphs by GID and those must be retained too).
//
// All other tables (GPOS, GSUB, GDEF, name, post, OS/2, …) are copied through
// untouched. The /head table is updated in two places:
//   - indexToLocFormat is forced to "long" (1) so we don't need to detect
//     whether the new offsets fit in the short format
//   - checkSumAdjustment is recomputed after the file is fully assembled

import type { ParsedTtf, TtfTableInfo } from '@/font/ttf-parser';

const HEAD_CHECKSUM_ADJUSTMENT_OFFSET = 8;
const HEAD_INDEX_TO_LOC_FORMAT_OFFSET = 50;
const MAGIC_HEAD_CHECKSUM = 0xb1b0afba;

const COMPOSITE_ARG_1_AND_2_ARE_WORDS = 0x0001;
const COMPOSITE_WE_HAVE_A_SCALE = 0x0008;
const COMPOSITE_MORE_COMPONENTS = 0x0020;
const COMPOSITE_WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
const COMPOSITE_WE_HAVE_A_TWO_BY_TWO = 0x0080;

export function subsetTtf(parsed: ParsedTtf, usedGids: Iterable<number>): Uint8Array {
  const glyf = parsed.tables.get('glyf');
  const loca = parsed.tables.get('loca');
  if (!glyf || !loca) {
    throw new Error('Subsetting requires a TrueType font with glyf+loca tables');
  }

  const closure = computeGlyphClosure(parsed, glyf, usedGids);
  const { newGlyf, newLoca } = rewriteGlyfLoca(parsed, glyf, closure);
  return assembleSubsetTtf(parsed, newGlyf, newLoca);
}

// The glyph ids actually present in a subset for `usedGids`: the seeds, their
// composite components (transitively), and GID 0. With Identity CID↔GID this is
// exactly the set of CIDs present in the embedded subset, used to build the
// PDF/A /CIDSet (ISO 19005-1 §6.3.5). Mirrors subsetTtf's own closure.
export function glyphClosure(parsed: ParsedTtf, usedGids: Iterable<number>): Set<number> {
  const glyf = parsed.tables.get('glyf');
  if (!glyf) return new Set<number>([0, ...usedGids]);
  return computeGlyphClosure(parsed, glyf, usedGids);
}

function computeGlyphClosure(
  parsed: ParsedTtf,
  glyf: TtfTableInfo,
  seed: Iterable<number>,
): Set<number> {
  const closure = new Set<number>([0]);
  for (const g of seed) closure.add(g);

  const stack: Array<number> = [...closure];
  while (stack.length > 0) {
    const gid = stack.pop()!;
    const data = readGlyphRaw(parsed, glyf, gid);
    if (!data || data.length < 10) continue;
    const numberOfContours = signedI16(data, 0);
    if (numberOfContours >= 0) continue;
    for (const ref of parseCompositeRefs(data)) {
      if (!closure.has(ref)) {
        closure.add(ref);
        stack.push(ref);
      }
    }
  }
  return closure;
}

function readGlyphRaw(parsed: ParsedTtf, glyf: TtfTableInfo, gid: number): Uint8Array | null {
  if (gid < 0 || gid >= parsed.numGlyphs) return null;
  const start = parsed.glyphOffsets[gid]!;
  const end = parsed.glyphOffsets[gid + 1]!;
  if (end <= start) return null;
  return parsed.raw.subarray(glyf.offset + start, glyf.offset + end);
}

function parseCompositeRefs(data: Uint8Array): Array<number> {
  const refs: Array<number> = [];
  let off = 10;
  let flags = 0;
  do {
    if (off + 4 > data.length) break;
    flags = (data[off]! << 8) | data[off + 1]!;
    off += 2;
    const glyphIndex = (data[off]! << 8) | data[off + 1]!;
    off += 2;
    refs.push(glyphIndex);
    off += flags & COMPOSITE_ARG_1_AND_2_ARE_WORDS ? 4 : 2;
    if (flags & COMPOSITE_WE_HAVE_A_SCALE) off += 2;
    else if (flags & COMPOSITE_WE_HAVE_AN_X_AND_Y_SCALE) off += 4;
    else if (flags & COMPOSITE_WE_HAVE_A_TWO_BY_TWO) off += 8;
  } while (flags & COMPOSITE_MORE_COMPONENTS);
  return refs;
}

interface RewrittenGlyfLoca {
  readonly newGlyf: Uint8Array;
  readonly newLoca: ReadonlyArray<number>;
}

function rewriteGlyfLoca(
  parsed: ParsedTtf,
  glyf: TtfTableInfo,
  closure: ReadonlySet<number>,
): RewrittenGlyfLoca {
  const parts: Array<Uint8Array> = [];
  const offsets: Array<number> = new Array(parsed.numGlyphs + 1);
  let cursor = 0;
  for (let gid = 0; gid < parsed.numGlyphs; gid++) {
    offsets[gid] = cursor;
    if (!closure.has(gid)) continue;
    const data = readGlyphRaw(parsed, glyf, gid);
    if (!data) continue;
    parts.push(data);
    cursor += data.length;
    if (cursor & 1) {
      parts.push(new Uint8Array(1));
      cursor++;
    }
  }
  offsets[parsed.numGlyphs] = cursor;
  return { newGlyf: concatBytes(parts), newLoca: offsets };
}

function assembleSubsetTtf(
  parsed: ParsedTtf,
  newGlyf: Uint8Array,
  newLoca: ReadonlyArray<number>,
): Uint8Array {
  const newLocaBytes = new Uint8Array((parsed.numGlyphs + 1) * 4);
  const locaView = new DataView(newLocaBytes.buffer);
  for (let i = 0; i <= parsed.numGlyphs; i++) {
    locaView.setUint32(i * 4, newLoca[i]!, false);
  }

  const headInfo = parsed.tables.get('head')!;
  const oldHead = parsed.raw.subarray(headInfo.offset, headInfo.offset + headInfo.length);
  const newHead = new Uint8Array(oldHead);
  newHead[HEAD_CHECKSUM_ADJUSTMENT_OFFSET] = 0;
  newHead[HEAD_CHECKSUM_ADJUSTMENT_OFFSET + 1] = 0;
  newHead[HEAD_CHECKSUM_ADJUSTMENT_OFFSET + 2] = 0;
  newHead[HEAD_CHECKSUM_ADJUSTMENT_OFFSET + 3] = 0;
  newHead[HEAD_INDEX_TO_LOC_FORMAT_OFFSET] = 0;
  newHead[HEAD_INDEX_TO_LOC_FORMAT_OFFSET + 1] = 1;

  const orderedTags = [...parsed.tables.keys()];
  const tableData = new Map<string, Uint8Array>();
  for (const tag of orderedTags) {
    if (tag === 'glyf') tableData.set(tag, newGlyf);
    else if (tag === 'loca') tableData.set(tag, newLocaBytes);
    else if (tag === 'head') tableData.set(tag, newHead);
    else {
      const info = parsed.tables.get(tag)!;
      tableData.set(tag, parsed.raw.subarray(info.offset, info.offset + info.length));
    }
  }

  const numTables = orderedTags.length;
  const directoryBytes = 12 + 16 * numTables;

  interface PlacedTable {
    readonly tag: string;
    readonly offset: number;
    readonly length: number;
    readonly checksum: number;
    readonly data: Uint8Array;
  }
  const placed: Array<PlacedTable> = [];
  let cursor = directoryBytes;
  for (const tag of orderedTags) {
    const data = tableData.get(tag)!;
    const offset = cursor;
    const length = data.length;
    const checksum = paddedChecksum(data);
    placed.push({ tag, offset, length, checksum, data });
    cursor += length;
    while (cursor & 3) cursor++;
  }

  const totalSize = cursor;
  const out = new Uint8Array(totalSize);
  const outView = new DataView(out.buffer);

  outView.setUint32(0, parsed.sfntVersion, false);
  outView.setUint16(4, numTables, false);
  const { searchRange, entrySelector, rangeShift } = directoryGeometry(numTables);
  outView.setUint16(6, searchRange, false);
  outView.setUint16(8, entrySelector, false);
  outView.setUint16(10, rangeShift, false);

  for (let i = 0; i < placed.length; i++) {
    const entry = placed[i]!;
    const dirOff = 12 + i * 16;
    out[dirOff] = entry.tag.charCodeAt(0);
    out[dirOff + 1] = entry.tag.charCodeAt(1);
    out[dirOff + 2] = entry.tag.charCodeAt(2);
    out[dirOff + 3] = entry.tag.charCodeAt(3);
    outView.setUint32(dirOff + 4, entry.checksum, false);
    outView.setUint32(dirOff + 8, entry.offset, false);
    outView.setUint32(dirOff + 12, entry.length, false);
  }

  for (const entry of placed) {
    out.set(entry.data, entry.offset);
  }

  const fileChecksum = paddedChecksum(out);
  const adjustment = (MAGIC_HEAD_CHECKSUM - fileChecksum) >>> 0;
  const head = placed.find((p) => p.tag === 'head')!;
  outView.setUint32(head.offset + HEAD_CHECKSUM_ADJUSTMENT_OFFSET, adjustment, false);

  return out;
}

function directoryGeometry(numTables: number): {
  searchRange: number;
  entrySelector: number;
  rangeShift: number;
} {
  let entrySelector = 0;
  let pow2 = 1;
  while (pow2 * 2 <= numTables) {
    pow2 *= 2;
    entrySelector++;
  }
  const searchRange = pow2 * 16;
  const rangeShift = numTables * 16 - searchRange;
  return { searchRange, entrySelector, rangeShift };
}

function paddedChecksum(data: Uint8Array): number {
  let sum = 0;
  const len = data.length;
  let i = 0;
  while (i + 4 <= len) {
    const b0 = data[i]!;
    const b1 = data[i + 1]!;
    const b2 = data[i + 2]!;
    const b3 = data[i + 3]!;
    const word = b0 * 0x1000000 + ((b1 << 16) | (b2 << 8) | b3);
    sum = (sum + word) >>> 0;
    i += 4;
  }
  if (i < len) {
    let tail = 0;
    for (let j = 0; j < 4; j++) {
      tail = tail * 0x100 + (i + j < len ? data[i + j]! : 0);
    }
    sum = (sum + tail) >>> 0;
  }
  return sum;
}

function signedI16(data: Uint8Array, offset: number): number {
  const v = (data[offset]! << 8) | data[offset + 1]!;
  return v >= 0x8000 ? v - 0x10000 : v;
}

function concatBytes(parts: ReadonlyArray<Uint8Array>): Uint8Array {
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
