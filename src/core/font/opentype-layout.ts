// OpenType GSUB / GPOS table parser. We extract only the bits the renderer
// needs:
//
//   GSUB Lookup Type 4 — Ligature substitution. Produces a map keyed on a
//   sequence of input GIDs (comma-separated) → output GID. Drives Latin /
//   Cyrillic ligatures like "fi" → ﬁ.
//
//   GPOS Lookup Type 2 — Pair positioning. Produces a map keyed on (GID1,
//   GID2) → advance adjustment in font units. Drives kerning.
//
// Both lookups are only consulted when their owning feature is one of the
// "default-on" Latin features: liga, clig, kern. Script/lang filtering is
// intentionally skipped — we always read the default lang sys.

import { assignArabicForms } from '@/core/font/arabic-joining';
import { BigEndianReader } from '@/core/font/binary-reader';

export type LigatureMap = ReadonlyMap<string, number>;
export type KerningMap = ReadonlyMap<string, number>;

// GSUB single-substitution maps (gid → positional gid) for the Arabic cursive
// features init / medi / fina. Empty for fonts without Arabic shaping.
export interface ArabicJoiningForms {
  readonly init: ReadonlyMap<number, number>;
  readonly medi: ReadonlyMap<number, number>;
  readonly fina: ReadonlyMap<number, number>;
}

export const EMPTY_JOINING_FORMS: ArabicJoiningForms = {
  init: new Map(),
  medi: new Map(),
  fina: new Map(),
};

interface TableInfo {
  readonly offset: number;
  readonly length: number;
}

// Parse the GSUB table and return only the liga / clig substitutions.
export function parseGsubLigatures(raw: Uint8Array, table: TableInfo | undefined): LigatureMap {
  if (!table) return new Map();
  try {
    const lookupOffsets = collectLookupOffsets(raw, table, ['liga', 'clig']);
    const ligatures = new Map<string, number>();
    for (const lookupOffset of lookupOffsets) {
      readLookup(raw, lookupOffset, (lookupType, subtableOffset) => {
        if (lookupType === 4) parseLigatureSubtable(raw, subtableOffset, ligatures);
      });
    }
    return ligatures;
  } catch {
    return new Map();
  }
}

// Parse the GPOS table and return only the kern lookup as a map keyed on
// "gid1,gid2".
export function parseGposKerning(raw: Uint8Array, table: TableInfo | undefined): KerningMap {
  if (!table) return new Map();
  try {
    const lookupOffsets = collectLookupOffsets(raw, table, ['kern']);
    const kerning = new Map<string, number>();
    for (const lookupOffset of lookupOffsets) {
      readLookup(raw, lookupOffset, (lookupType, subtableOffset) => {
        if (lookupType === 2) parsePairPosSubtable(raw, subtableOffset, kerning);
      });
    }
    return kerning;
  } catch {
    return new Map();
  }
}

// Parse GSUB single substitution (Lookup Type 1) for the Arabic positional
// features. Each produces a gid → gid map the shaper applies per cursive form.
export function parseGsubArabicForms(
  raw: Uint8Array,
  table: TableInfo | undefined,
): ArabicJoiningForms {
  if (!table) return EMPTY_JOINING_FORMS;
  try {
    const forFeature = (tag: string): Map<number, number> => {
      const map = new Map<number, number>();
      for (const lookupOffset of collectLookupOffsets(raw, table, [tag])) {
        readLookup(raw, lookupOffset, (lookupType, subtableOffset) => {
          if (lookupType === 1) parseSingleSubstSubtable(raw, subtableOffset, map);
        });
      }
      return map;
    };
    const forms = { init: forFeature('init'), medi: forFeature('medi'), fina: forFeature('fina') };
    if (forms.init.size === 0 && forms.medi.size === 0 && forms.fina.size === 0) {
      return EMPTY_JOINING_FORMS;
    }
    return forms;
  } catch {
    return EMPTY_JOINING_FORMS;
  }
}

// GSUB Lookup Type 1 — Single Substitution (Format 1 delta, Format 2 array).
function parseSingleSubstSubtable(
  raw: Uint8Array,
  subtableOffset: number,
  map: Map<number, number>,
): void {
  const r = new BigEndianReader(raw, subtableOffset);
  const format = r.u16();
  const coverageOffset = r.u16();
  const coverage = readCoverage(raw, subtableOffset + coverageOffset);
  if (format === 1) {
    const raw16 = r.u16();
    const delta = raw16 >= 0x8000 ? raw16 - 0x10000 : raw16; // signed
    for (const gid of coverage) map.set(gid, (gid + delta) & 0xffff);
  } else if (format === 2) {
    const count = r.u16();
    for (let i = 0; i < count; i++) {
      const sub = r.u16();
      if (i < coverage.length) map.set(coverage[i]!, sub);
    }
  }
}

// Walk Script → DefaultLangSys → FeatureIndexes → Feature → Lookup indices.
// Returns the absolute offsets of every lookup referenced by the named
// features, deduplicated.
function collectLookupOffsets(
  raw: Uint8Array,
  table: TableInfo,
  wantedFeatures: ReadonlyArray<string>,
): Array<number> {
  const base = table.offset;
  const r = new BigEndianReader(raw, base);
  r.u16(); // majorVersion
  r.u16(); // minorVersion
  const scriptListOff = r.u16();
  const featureListOff = r.u16();
  const lookupListOff = r.u16();

  // Walk feature list: gather all Feature tables with one of the wanted tags.
  const featureListBase = base + featureListOff;
  const featureWanted = new Set<number>(); // feature indices
  {
    const fr = new BigEndianReader(raw, featureListBase);
    const featureCount = fr.u16();
    for (let i = 0; i < featureCount; i++) {
      const tag = fr.tag();
      const featureOff = fr.u16();
      if (wantedFeatures.includes(tag)) {
        // Read lookups directly from feature table.
        const featBase = featureListBase + featureOff;
        const fr2 = new BigEndianReader(raw, featBase);
        fr2.u16(); // featureParamsOffset
        const lookupCount = fr2.u16();
        for (let k = 0; k < lookupCount; k++) featureWanted.add(fr2.u16());
      }
      void i;
    }
  }

  // Build absolute lookup offsets.
  const lookupListBase = base + lookupListOff;
  const lookupListReader = new BigEndianReader(raw, lookupListBase);
  const lookupCount = lookupListReader.u16();
  const lookupOffsets: Array<number> = [];
  for (let i = 0; i < lookupCount; i++) {
    const offset = lookupListReader.u16();
    lookupOffsets.push(lookupListBase + offset);
  }

  void scriptListOff;
  const result: Array<number> = [];
  const seen = new Set<number>();
  for (const idx of featureWanted) {
    const off = lookupOffsets[idx];
    if (off !== undefined && !seen.has(off)) {
      seen.add(off);
      result.push(off);
    }
  }
  return result;
}

function readLookup(
  raw: Uint8Array,
  lookupOffset: number,
  visit: (lookupType: number, subtableOffset: number) => void,
): void {
  const r = new BigEndianReader(raw, lookupOffset);
  const lookupType = r.u16();
  r.u16(); // lookupFlag
  const subTableCount = r.u16();
  for (let i = 0; i < subTableCount; i++) {
    const subOff = r.u16();
    visit(lookupType, lookupOffset + subOff);
  }
}

// LookupType 4 — Ligature Substitution Format 1.
function parseLigatureSubtable(
  raw: Uint8Array,
  subtableOffset: number,
  out: Map<string, number>,
): void {
  const r = new BigEndianReader(raw, subtableOffset);
  const substFormat = r.u16();
  if (substFormat !== 1) return;
  const coverageOff = r.u16();
  const ligatureSetCount = r.u16();
  const ligatureSetOffs: Array<number> = new Array(ligatureSetCount);
  for (let i = 0; i < ligatureSetCount; i++) ligatureSetOffs[i] = r.u16();

  const coverageGlyphs = readCoverage(raw, subtableOffset + coverageOff);
  for (let i = 0; i < ligatureSetCount; i++) {
    const firstGlyph = coverageGlyphs[i];
    if (firstGlyph === undefined) continue;
    parseLigatureSet(raw, subtableOffset + ligatureSetOffs[i]!, firstGlyph, out);
  }
}

function parseLigatureSet(
  raw: Uint8Array,
  setOffset: number,
  firstGlyph: number,
  out: Map<string, number>,
): void {
  const r = new BigEndianReader(raw, setOffset);
  const ligatureCount = r.u16();
  const ligOffsets: Array<number> = new Array(ligatureCount);
  for (let i = 0; i < ligatureCount; i++) ligOffsets[i] = r.u16();

  for (let i = 0; i < ligatureCount; i++) {
    const r2 = new BigEndianReader(raw, setOffset + ligOffsets[i]!);
    const ligGlyph = r2.u16();
    const componentCount = r2.u16();
    const components: Array<number> = [firstGlyph];
    for (let k = 1; k < componentCount; k++) components.push(r2.u16());
    out.set(components.join(','), ligGlyph);
  }
}

// LookupType 2 — Pair Positioning. Two formats:
//   Format 1: per-glyph PairSet (most common).
//   Format 2: class-based grid (large fonts).
function parsePairPosSubtable(
  raw: Uint8Array,
  subtableOffset: number,
  out: Map<string, number>,
): void {
  const r = new BigEndianReader(raw, subtableOffset);
  const posFormat = r.u16();
  if (posFormat === 1) parsePairPosFormat1(raw, subtableOffset, out);
  else if (posFormat === 2) parsePairPosFormat2(raw, subtableOffset, out);
}

function parsePairPosFormat1(
  raw: Uint8Array,
  subtableOffset: number,
  out: Map<string, number>,
): void {
  const r = new BigEndianReader(raw, subtableOffset);
  r.u16(); // posFormat
  const coverageOff = r.u16();
  const valueFormat1 = r.u16();
  const valueFormat2 = r.u16();
  const pairSetCount = r.u16();
  const pairSetOffs: Array<number> = new Array(pairSetCount);
  for (let i = 0; i < pairSetCount; i++) pairSetOffs[i] = r.u16();

  const value1Size = valueRecordSize(valueFormat1);
  const value2Size = valueRecordSize(valueFormat2);
  const coverage = readCoverage(raw, subtableOffset + coverageOff);

  for (let i = 0; i < pairSetCount; i++) {
    const first = coverage[i];
    if (first === undefined) continue;
    const psBase = subtableOffset + pairSetOffs[i]!;
    const psr = new BigEndianReader(raw, psBase);
    const pairCount = psr.u16();
    for (let j = 0; j < pairCount; j++) {
      const second = psr.u16();
      const v1 = readValueRecordAdvance(psr, valueFormat1);
      // Skip the rest of value1's record (we only care about xAdvance).
      psr.skip(value1Size - bytesRead(valueFormat1));
      const v2 = readValueRecordAdvance(psr, valueFormat2);
      psr.skip(value2Size - bytesRead(valueFormat2));
      // Convention: kerning adjusts the first glyph's advance.
      const adj = v1 !== 0 ? v1 : v2;
      if (adj !== 0) out.set(`${first},${second}`, adj);
    }
  }
}

function parsePairPosFormat2(
  raw: Uint8Array,
  subtableOffset: number,
  out: Map<string, number>,
): void {
  const r = new BigEndianReader(raw, subtableOffset);
  r.u16(); // posFormat
  const coverageOff = r.u16();
  const valueFormat1 = r.u16();
  const valueFormat2 = r.u16();
  const classDef1Off = r.u16();
  const classDef2Off = r.u16();
  const class1Count = r.u16();
  const class2Count = r.u16();

  const value1Size = valueRecordSize(valueFormat1);
  const value2Size = valueRecordSize(valueFormat2);
  const pairRecordSize = value1Size + value2Size;

  const coverage = readCoverage(raw, subtableOffset + coverageOff);
  const class1 = readClassDef(raw, subtableOffset + classDef1Off);
  const class2 = readClassDef(raw, subtableOffset + classDef2Off);

  const recordsBase = r.offset;
  for (let i = 0; i < class1Count; i++) {
    for (let j = 0; j < class2Count; j++) {
      const recordOff = recordsBase + (i * class2Count + j) * pairRecordSize;
      const pr = new BigEndianReader(raw, recordOff);
      const v1 = readValueRecordAdvance(pr, valueFormat1);
      pr.skip(value1Size - bytesRead(valueFormat1));
      const v2 = readValueRecordAdvance(pr, valueFormat2);
      pr.skip(value2Size - bytesRead(valueFormat2));
      const adj = v1 !== 0 ? v1 : v2;
      if (adj === 0) continue;
      // Emit a kerning entry for every glyph pair that maps to this class pair.
      const firsts = matchingGlyphs(coverage, class1, i);
      const seconds = matchingGlyphs(undefined, class2, j);
      for (const a of firsts) {
        for (const b of seconds) {
          out.set(`${a},${b}`, adj);
        }
      }
    }
  }
}

// Compute the byte size of a ValueRecord given its bit mask.
function valueRecordSize(valueFormat: number): number {
  let bits = valueFormat;
  let count = 0;
  while (bits) {
    count += bits & 1;
    bits >>>= 1;
  }
  return count * 2;
}

// Read the xAdvance field (if present) and leave the reader at the position
// immediately after it. The caller is expected to skip the rest of the record.
function readValueRecordAdvance(r: BigEndianReader, valueFormat: number): number {
  // ValueRecord field order: xPlacement, yPlacement, xAdvance, yAdvance, ...
  if (valueFormat & 0x0001) r.i16(); // xPlacement
  if (valueFormat & 0x0002) r.i16(); // yPlacement
  if (valueFormat & 0x0004) return r.i16(); // xAdvance
  return 0;
}

// Returns the number of bytes that readValueRecordAdvance just read so the
// caller can compute the remaining record size and skip it.
function bytesRead(valueFormat: number): number {
  let n = 0;
  if (valueFormat & 0x0001) n += 2;
  if (valueFormat & 0x0002) n += 2;
  if (valueFormat & 0x0004) n += 2;
  return n;
}

function readCoverage(raw: Uint8Array, offset: number): Array<number> {
  const r = new BigEndianReader(raw, offset);
  const format = r.u16();
  if (format === 1) {
    const glyphCount = r.u16();
    const out: Array<number> = new Array(glyphCount);
    for (let i = 0; i < glyphCount; i++) out[i] = r.u16();
    return out;
  }
  if (format === 2) {
    const rangeCount = r.u16();
    const out: Array<number> = [];
    for (let i = 0; i < rangeCount; i++) {
      const startGlyph = r.u16();
      const endGlyph = r.u16();
      const startCoverageIndex = r.u16();
      void startCoverageIndex;
      for (let g = startGlyph; g <= endGlyph; g++) out.push(g);
    }
    return out;
  }
  return [];
}

function readClassDef(raw: Uint8Array, offset: number): Map<number, number> {
  const r = new BigEndianReader(raw, offset);
  const format = r.u16();
  const out = new Map<number, number>();
  if (format === 1) {
    const startGlyph = r.u16();
    const glyphCount = r.u16();
    for (let i = 0; i < glyphCount; i++) {
      const cls = r.u16();
      if (cls !== 0) out.set(startGlyph + i, cls);
    }
  } else if (format === 2) {
    const classRangeCount = r.u16();
    for (let i = 0; i < classRangeCount; i++) {
      const startGlyph = r.u16();
      const endGlyph = r.u16();
      const cls = r.u16();
      if (cls !== 0) {
        for (let g = startGlyph; g <= endGlyph; g++) out.set(g, cls);
      }
    }
  }
  return out;
}

// Return the GIDs in `coverage` (if present) that map to a given class. When
// coverage is omitted, return all glyphs in the class def — used for the
// second glyph of a class-based pair.
function matchingGlyphs(
  coverage: Array<number> | undefined,
  classMap: Map<number, number>,
  cls: number,
): Array<number> {
  if (coverage) {
    const out: Array<number> = [];
    for (const g of coverage) {
      // Class 0 = "any glyph not in the class def". We only look at non-zero
      // classes; coverage glyphs without an entry implicitly belong to class 0.
      const actual = classMap.get(g) ?? 0;
      if (actual === cls) out.push(g);
    }
    return out;
  }
  const out: Array<number> = [];
  for (const [gid, c] of classMap) if (c === cls) out.push(gid);
  return out;
}

// ----- Shaping API -----

export interface ShapedRun {
  readonly gids: Array<number>;
  readonly advances: Array<number>; // font-unit advance per output gid
}

export function shapeText(
  text: string,
  glyphForCodepoint: (cp: number) => number,
  advanceWidths: ReadonlyArray<number>,
  ligatures: LigatureMap,
  kerning: KerningMap,
  joiningForms: ArabicJoiningForms = EMPTY_JOINING_FORMS,
): ShapedRun {
  // Step 1: map characters to GIDs.
  const cps: Array<number> = [];
  for (const ch of text) cps.push(ch.codePointAt(0)!);
  const initialGids: Array<number> = cps.map(glyphForCodepoint);

  // Step 1b: Arabic cursive joining — swap each letter for its contextual form
  // (initial/medial/final) via the font's init/medi/fina GSUB lookups. Skipped
  // when the font has none (the maps are empty → output unchanged).
  if (joiningForms.init.size + joiningForms.medi.size + joiningForms.fina.size > 0) {
    const forms = assignArabicForms(cps);
    for (let k = 0; k < initialGids.length; k++) {
      const map =
        forms[k] === 'init'
          ? joiningForms.init
          : forms[k] === 'medi'
            ? joiningForms.medi
            : forms[k] === 'fina'
              ? joiningForms.fina
              : undefined;
      const sub = map?.get(initialGids[k]!);
      if (sub !== undefined) initialGids[k] = sub;
    }
  }

  // Step 2: apply ligature substitution greedily, preferring longer matches.
  const ligated: Array<number> = [];
  let i = 0;
  while (i < initialGids.length) {
    let bestLen = 1;
    let bestOut = initialGids[i]!;
    const maxTry = Math.min(4, initialGids.length - i);
    for (let len = 2; len <= maxTry; len++) {
      const key = initialGids.slice(i, i + len).join(',');
      const out = ligatures.get(key);
      if (out !== undefined && len > bestLen) {
        bestLen = len;
        bestOut = out;
      }
    }
    ligated.push(bestOut);
    i += bestLen;
  }

  // Step 3: assign advances; apply pair kerning by adjusting the *previous*
  // glyph's advance (matches how PDF text rendering interprets advances).
  const advances: Array<number> = new Array(ligated.length);
  for (let j = 0; j < ligated.length; j++) {
    const gid = ligated[j]!;
    advances[j] = advanceWidths[gid] ?? 0;
  }
  for (let j = 1; j < ligated.length; j++) {
    const adj = kerning.get(`${ligated[j - 1]},${ligated[j]}`);
    if (adj !== undefined && adj !== 0) advances[j - 1]! += adj;
  }

  return { gids: ligated, advances };
}
