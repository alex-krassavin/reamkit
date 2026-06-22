// ISO/IEC 14496-22 (Open Font Format) — TrueType / OpenType parser.
//
// Parses the subset of tables needed to embed the font in a PDF as a
// CIDFontType2 with a FontDescriptor:
//   sfnt header + table directory  (§5.1)
//   head   — unitsPerEm, FontBBox  (§5.2.4)
//   hhea   — ascender/descender/numberOfHMetrics  (§5.2.5)
//   maxp   — numGlyphs  (§5.2.6)
//   hmtx   — glyph advance widths  (§5.2.7)
//   cmap   — Unicode codepoint → glyph index  (§5.2.1.3)
//   name   — PostScript name (nameID 6)  (§5.2.2)
//   post   — italicAngle, isFixedPitch  (§5.2.18)
//   OS/2   — capHeight, xHeight, weight class, italic flag  (§5.2.3)
//
// glyf/loca/CFF table data is left untouched — the entire raw font is
// embedded in the PDF /FontFile2 stream, so we do not need to decode glyph
// outlines ourselves.

import type { ArabicJoiningForms, KerningMap, LigatureMap } from '@/core/font/opentype-layout';
import { BigEndianReader } from '@/core/font/binary-reader';
import {
  parseGposKerning,
  parseGsubArabicForms,
  parseGsubLigatures,
} from '@/core/font/opentype-layout';

/** A table's location within the sfnt: byte offset and length. */
export interface TtfTableInfo {
  /** Byte offset of the table from the start of the file. */
  readonly offset: number;
  /** Table length in bytes. */
  readonly length: number;
}

/**
 * A parsed TrueType / OpenType font (ISO/IEC 14496-22): metrics, the
 * cmap-derived codepoint→glyph mapping, OpenType layout maps (ligatures,
 * kerning, Arabic joining) and the raw bytes for embedding. Glyph outlines
 * (glyf/CFF) are not decoded — the whole font is embedded in the PDF /FontFile2.
 */
export interface ParsedTtf {
  readonly raw: Uint8Array;
  readonly sfntVersion: number;
  readonly tables: ReadonlyMap<string, TtfTableInfo>;
  readonly postScriptName: string;
  readonly unitsPerEm: number;
  readonly fontBBox: readonly [number, number, number, number];
  readonly ascender: number;
  readonly descender: number;
  /**
   * hhea line gap (font units); with ascender/descender it forms the hhea
   * vertical triple that drives the `'libreoffice'` line-height profile (E-PARITY).
   */
  readonly lineGap: number;
  readonly capHeight: number;
  readonly xHeight: number;
  /**
   * E-PARITY line-height metrics: OS/2 win/typo verticals (font units) selected
   * by layout profile. `'word'` uses winAscent/winDescent (the GDI cell box);
   * `'libreoffice'` uses the hhea triple above, or typo* when `useTypoMetrics` is
   * set. All fall back to hhea when the font has no OS/2 table.
   */
  readonly vmetrics: {
    readonly typoAscent: number;
    readonly typoDescent: number;
    readonly typoLineGap: number;
    readonly winAscent: number;
    readonly winDescent: number;
    readonly useTypoMetrics: boolean;
  };
  readonly italicAngle: number;
  readonly stemV: number;
  readonly flags: number;
  readonly numGlyphs: number;
  readonly indexToLocFormat: 0 | 1;
  readonly glyphOffsets: ReadonlyArray<number>;
  readonly advanceWidths: ReadonlyArray<number>;
  readonly glyphForCodepoint: (cp: number) => number;
  /** OpenType ligature substitutions: input GID sequence `"gid1,gid2[,gid3…]"` → output GID. */
  readonly ligatures: LigatureMap;
  /** OpenType pair kerning: `"gid1,gid2"` → advance adjustment in font units. */
  readonly kerning: KerningMap;
  /**
   * Arabic cursive-joining substitutions (init/medi/fina). Empty for non-Arabic
   * fonts, so the shaper leaves output unchanged.
   */
  readonly joiningForms: ArabicJoiningForms;
}

const SFNT_TRUETYPE = 0x00010000;
const SFNT_OPENTYPE_CFF = 0x4f54544f;

/**
 * Parse TrueType / OpenType font bytes into a {@link ParsedTtf}.
 *
 * @param raw The font bytes (sfnt: TrueType `0x00010000` or OpenType-CFF `OTTO`).
 * @returns The parsed metrics, mappings and layout tables.
 * @throws Error when the bytes are not a TrueType or OpenType font.
 */
export function parseTtf(raw: Uint8Array): ParsedTtf {
  const r = new BigEndianReader(raw);
  const sfntVersion = r.u32();
  if (sfntVersion !== SFNT_TRUETYPE && sfntVersion !== SFNT_OPENTYPE_CFF) {
    throw new Error(
      `Not a TrueType or OpenType font (sfntVersion=0x${sfntVersion.toString(16).padStart(8, '0')})`,
    );
  }
  const numTables = r.u16();
  r.skip(6);

  const tables = new Map<string, { offset: number; length: number }>();
  for (let i = 0; i < numTables; i++) {
    const tag = r.tag();
    r.u32();
    const offset = r.u32();
    const length = r.u32();
    tables.set(tag, { offset, length });
  }

  const require = (tag: string): { offset: number; length: number } => {
    const t = tables.get(tag);
    if (!t) throw new Error(`TTF missing required table: ${tag}`);
    return t;
  };

  const head = require('head');
  r.seek(head.offset + 18);
  const unitsPerEm = r.u16();
  r.skip(16);
  const xMin = r.i16();
  const yMin = r.i16();
  const xMax = r.i16();
  const yMax = r.i16();
  r.skip(6);
  const indexToLocFormatRaw = r.i16();
  if (indexToLocFormatRaw !== 0 && indexToLocFormatRaw !== 1) {
    throw new Error(`Invalid head.indexToLocFormat: ${indexToLocFormatRaw}`);
  }
  const indexToLocFormat = indexToLocFormatRaw;

  const hhea = require('hhea');
  r.seek(hhea.offset + 4);
  const ascender = r.i16();
  const descender = r.i16();
  const lineGap = r.i16();
  r.skip(24);
  const numberOfHMetrics = r.u16();

  const maxp = require('maxp');
  r.seek(maxp.offset + 4);
  const numGlyphs = r.u16();

  const hmtx = require('hmtx');
  r.seek(hmtx.offset);
  const advanceWidths = new Array<number>(numGlyphs);
  let lastAdvance = 0;
  const longCount = Math.min(numberOfHMetrics, numGlyphs);
  for (let i = 0; i < longCount; i++) {
    lastAdvance = r.u16();
    r.skip(2);
    advanceWidths[i] = lastAdvance;
  }
  for (let i = longCount; i < numGlyphs; i++) {
    advanceWidths[i] = lastAdvance;
  }

  const cmap = require('cmap');
  const glyphForCodepoint = buildCmap(raw, cmap.offset);

  const name = require('name');
  const postScriptName = readPostScriptName(raw, name.offset);

  let italicAngle = 0;
  let isFixedPitch = false;
  const post = tables.get('post');
  if (post) {
    r.seek(post.offset);
    r.skip(4);
    italicAngle = r.i32() / 0x10000;
    r.skip(4);
    isFixedPitch = r.u32() !== 0;
  }

  let capHeight = 0;
  let xHeight = 0;
  let isItalic = false;
  let usWeightClass = 400;
  // E-PARITY vertical metrics — default to the hhea triple; OS/2 overrides below.
  let typoAscent = ascender;
  let typoDescent = descender;
  let typoLineGap = lineGap;
  let winAscent = Math.abs(ascender);
  let winDescent = Math.abs(descender);
  let useTypoMetrics = false;
  const os2 = tables.get('OS/2');
  if (os2) {
    r.seek(os2.offset);
    const os2Version = r.u16();
    r.skip(2);
    usWeightClass = r.u16();

    r.seek(os2.offset + 62);
    const fsSelection = r.u16();
    isItalic = (fsSelection & 0x01) !== 0;
    useTypoMetrics = (fsSelection & 0x80) !== 0;

    // sTypoAscender/Descender/LineGap (68/70/72) + usWinAscent/Descent (74/76)
    // are part of OS/2 v0 (table length >= 78).
    if (os2.length >= 78) {
      r.seek(os2.offset + 68);
      typoAscent = r.i16();
      typoDescent = r.i16();
      typoLineGap = r.i16();
      winAscent = r.u16();
      winDescent = r.u16();
    }

    if (os2Version >= 2 && os2.length >= 90) {
      r.seek(os2.offset + 86);
      xHeight = r.i16();
      capHeight = r.i16();
    }
  }

  if (capHeight === 0) capHeight = Math.round(ascender * 0.7);
  if (xHeight === 0) xHeight = Math.round(ascender * 0.5);

  let flags = 0;
  if (isFixedPitch) flags |= 1;
  flags |= 32;
  if (isItalic || italicAngle !== 0) flags |= 64;

  const stemV = usWeightClass > 500 ? 120 : 80;

  const glyphOffsets = readGlyphOffsets(raw, tables, numGlyphs, indexToLocFormat);

  const ligatures = parseGsubLigatures(raw, tables.get('GSUB'));
  const kerning = parseGposKerning(raw, tables.get('GPOS'));
  const joiningForms = parseGsubArabicForms(raw, tables.get('GSUB'));

  return {
    raw,
    sfntVersion,
    tables,
    postScriptName,
    unitsPerEm,
    fontBBox: [xMin, yMin, xMax, yMax] as const,
    ascender,
    descender,
    lineGap,
    capHeight,
    xHeight,
    vmetrics: { typoAscent, typoDescent, typoLineGap, winAscent, winDescent, useTypoMetrics },
    italicAngle,
    stemV,
    flags,
    numGlyphs,
    indexToLocFormat,
    glyphOffsets,
    advanceWidths,
    glyphForCodepoint,
    ligatures,
    kerning,
    joiningForms,
  };
}

function readGlyphOffsets(
  raw: Uint8Array,
  tables: ReadonlyMap<string, TtfTableInfo>,
  numGlyphs: number,
  indexToLocFormat: 0 | 1,
): Array<number> {
  const loca = tables.get('loca');
  if (!loca) return [];
  const r = new BigEndianReader(raw, loca.offset);
  const out = new Array<number>(numGlyphs + 1);
  if (indexToLocFormat === 0) {
    for (let i = 0; i <= numGlyphs; i++) out[i] = r.u16() * 2;
  } else {
    for (let i = 0; i <= numGlyphs; i++) out[i] = r.u32();
  }
  return out;
}

function buildCmap(raw: Uint8Array, cmapOffset: number): (cp: number) => number {
  const r = new BigEndianReader(raw, cmapOffset);
  r.u16();
  const numTables = r.u16();

  interface SubtableRecord {
    readonly platform: number;
    readonly encoding: number;
    readonly absoluteOffset: number;
    readonly format: number;
  }
  const subtables: Array<SubtableRecord> = [];
  for (let i = 0; i < numTables; i++) {
    const platform = r.u16();
    const encoding = r.u16();
    const relOffset = r.u32();
    const absoluteOffset = cmapOffset + relOffset;
    const peek = new BigEndianReader(raw, absoluteOffset);
    const format = peek.u16();
    subtables.push({ platform, encoding, absoluteOffset, format });
  }

  const priorities: ReadonlyArray<{ readonly platform: number; readonly encoding: number }> = [
    { platform: 3, encoding: 10 },
    { platform: 0, encoding: 4 },
    { platform: 0, encoding: 6 },
    { platform: 3, encoding: 1 },
    { platform: 0, encoding: 3 },
    { platform: 0, encoding: 1 },
    { platform: 0, encoding: 0 },
  ];

  let chosen: SubtableRecord | undefined;
  for (const want of priorities) {
    chosen = subtables.find((s) => s.platform === want.platform && s.encoding === want.encoding);
    if (chosen) break;
  }
  if (!chosen) chosen = subtables[0];
  if (!chosen) return () => 0;

  if (chosen.format === 4) return buildCmapFormat4(raw, chosen.absoluteOffset);
  if (chosen.format === 12) return buildCmapFormat12(raw, chosen.absoluteOffset);
  throw new Error(`Unsupported cmap subtable format: ${chosen.format}`);
}

function buildCmapFormat4(raw: Uint8Array, offset: number): (cp: number) => number {
  const r = new BigEndianReader(raw, offset);
  r.u16();
  r.u16();
  r.u16();
  const segCountX2 = r.u16();
  const segCount = segCountX2 >> 1;
  r.skip(6);

  const endCode = new Uint16Array(segCount);
  for (let i = 0; i < segCount; i++) endCode[i] = r.u16();
  r.u16();
  const startCode = new Uint16Array(segCount);
  for (let i = 0; i < segCount; i++) startCode[i] = r.u16();
  const idDelta = new Int16Array(segCount);
  for (let i = 0; i < segCount; i++) idDelta[i] = r.i16();

  const idRangeOffsetPos = r.offset;
  const idRangeOffset = new Uint16Array(segCount);
  for (let i = 0; i < segCount; i++) idRangeOffset[i] = r.u16();

  return (cp: number): number => {
    if (cp > 0xffff) return 0;
    let seg = -1;
    for (let i = 0; i < segCount; i++) {
      if (endCode[i]! >= cp) {
        seg = i;
        break;
      }
    }
    if (seg < 0) return 0;
    if (startCode[seg]! > cp) return 0;
    if (idRangeOffset[seg] === 0) {
      return (cp + idDelta[seg]!) & 0xffff;
    }
    const glyphIdByteOffset =
      idRangeOffsetPos + seg * 2 + idRangeOffset[seg]! + (cp - startCode[seg]!) * 2;
    const peek = new BigEndianReader(raw, glyphIdByteOffset);
    const glyphId = peek.u16();
    if (glyphId === 0) return 0;
    return (glyphId + idDelta[seg]!) & 0xffff;
  };
}

function buildCmapFormat12(raw: Uint8Array, offset: number): (cp: number) => number {
  const r = new BigEndianReader(raw, offset);
  r.u16();
  r.u16();
  r.u32();
  r.u32();
  const numGroups = r.u32();

  const starts = new Uint32Array(numGroups);
  const ends = new Uint32Array(numGroups);
  const startGids = new Uint32Array(numGroups);
  for (let i = 0; i < numGroups; i++) {
    starts[i] = r.u32();
    ends[i] = r.u32();
    startGids[i] = r.u32();
  }

  return (cp: number): number => {
    let lo = 0;
    let hi = numGroups - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >>> 1;
      if (cp < starts[mid]!) hi = mid - 1;
      else if (cp > ends[mid]!) lo = mid + 1;
      else return startGids[mid]! + (cp - starts[mid]!);
    }
    return 0;
  };
}

function readPostScriptName(raw: Uint8Array, offset: number): string {
  const r = new BigEndianReader(raw, offset);
  r.u16();
  const count = r.u16();
  const storageOffset = r.u16();
  const stringsBase = offset + storageOffset;

  interface NameRecord {
    readonly platform: number;
    readonly encoding: number;
    readonly language: number;
    readonly nameId: number;
    readonly length: number;
    readonly stringOffset: number;
  }
  const records: Array<NameRecord> = [];
  for (let i = 0; i < count; i++) {
    records.push({
      platform: r.u16(),
      encoding: r.u16(),
      language: r.u16(),
      nameId: r.u16(),
      length: r.u16(),
      stringOffset: r.u16(),
    });
  }

  const filters: ReadonlyArray<(rec: NameRecord) => boolean> = [
    (rec) => rec.platform === 3 && rec.encoding === 1 && rec.language === 0x0409,
    (rec) => rec.platform === 3 && rec.encoding === 1,
    (rec) => rec.platform === 1 && rec.encoding === 0 && rec.language === 0,
    (rec) => rec.platform === 0,
    () => true,
  ];

  for (const filter of filters) {
    const rec = records.find((x) => x.nameId === 6 && filter(x));
    if (rec) return decodeNameRecord(raw, stringsBase + rec.stringOffset, rec.length, rec.platform);
  }
  throw new Error('TTF name table missing PostScript name (nameID 6)');
}

function decodeNameRecord(
  raw: Uint8Array,
  offset: number,
  length: number,
  platform: number,
): string {
  const bytes = raw.subarray(offset, offset + length);
  if (platform === 1) {
    return new TextDecoder('macintosh', { fatal: false }).decode(bytes);
  }
  return new TextDecoder('utf-16be', { fatal: false }).decode(bytes);
}
