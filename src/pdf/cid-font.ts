// ISO 32000-1:2008 §9.7 — Composite (Type 0) fonts with CIDFontType2 descendants.
//
// Pipeline:
//   raw TTF  ->  /FontFile2 stream  (§9.9)
//            ->  /FontDescriptor    (§9.8)
//            ->  /CIDFontType2      (§9.7.4)  +  /W array  +  /CIDToGIDMap = Identity
//            ->  /Type0             (§9.7.2)  with /Encoding /Identity-H
//                                              and /ToUnicode CMap (Annex D)

import type { FontMeasure, ParsedTtf } from '@/core/font';
import type { PdfRef } from '@/pdf/objects';
import type { PdfDocument } from '@/pdf/writer';
import { createFontMeasure, glyphClosure, lettersForLigature, subsetTtf } from '@/core/font';
import { deflatedStream, dict, name, ref } from '@/pdf/objects';

const encoder = new TextEncoder();

/** Options for {@link embedTtfFont}. */
export interface EmbedTtfOptions {
  /** The glyph ids actually used; given, the font is subset to just these (plus their closure). */
  readonly usedGids?: Iterable<number>;
  /**
   * Emit a `/CIDSet` in the font descriptor (required by PDF/A-1; omitted
   * otherwise, since PDF/A-2/3 only constrain it when present).
   */
  readonly cidSet?: boolean;
}

/** The result of {@link embedTtfFont}: the Type 0 font reference plus measurement/encoding helpers. */
export interface EmbeddedFont {
  /** Indirect reference to the `/Type0` font dictionary, for page resource `/Font` entries. */
  readonly fontRef: PdfRef;
  /** The parsed source font. */
  readonly parsed: ParsedTtf;
  /** Glyph advance width in PDF text-space units (1000/em) for a given glyph id. */
  readonly pdfWidthForGid: (gid: number) => number;
  /** Measured width of `text` at `fontSize`, in points. */
  readonly textWidthPt: (text: string, fontSize: number) => number;
  /** Encode `text` as the hex CID string used inside a `Tj`/`TJ` show operator (Identity-H). */
  readonly encodeTextAsCidHex: (text: string) => string;
}

/**
 * Embed a TrueType font as a composite (Type 0) font with a CIDFontType2
 * descendant (ISO 32000-1 §9.7), optionally subsetting it to the used glyphs.
 * Adds the `/FontFile2`, `/FontDescriptor`, `/CIDFontType2` and `/Type0` objects
 * (plus a `/ToUnicode` CMap and an optional `/CIDSet`) to `doc`.
 *
 * @param parsed  The parsed TrueType font.
 * @param options Subset glyph set and/or PDF/A-1 `/CIDSet` emission.
 * @returns The font reference and measurement/encoding helpers ({@link EmbeddedFont}).
 */
export function embedTtfFont(
  doc: PdfDocument,
  parsed: ParsedTtf,
  options: EmbedTtfOptions = {},
): EmbeddedFont {
  const scale = 1000 / parsed.unitsPerEm;
  const toPdfUnits = (v: number): number => Math.round(v * scale);

  const usedGidsArr = options.usedGids ? [...options.usedGids] : undefined;
  // What the subset actually contains: the used glyphs plus whatever their
  // composites pull in — the bound for /ToUnicode.
  const subsetGids = usedGidsArr
    ? [...glyphClosure(parsed, usedGidsArr)].sort((a, b) => a - b)
    : undefined;
  const fontFileBytes = usedGidsArr ? subsetTtf(parsed, usedGidsArr) : parsed.raw;
  // §9.9 wants the UNCOMPRESSED length in /Length1; the stream itself deflates.
  const fontFileRef = doc.add(deflatedStream({ Length1: fontFileBytes.byteLength }, fontFileBytes));

  // ISO 19005-1 §6.3.5 — PDF/A-1 *requires* a /CIDSet in a CIDFont subset's
  // descriptor, marking the CIDs present (Identity ordering ⇒ CID = GID, so the
  // subset's glyph closure is the CID set). PDF/A-2/3 (§6.2.11.4.2) make it
  // optional and demand an exact font-program match, which is brittle — so we
  // emit /CIDSet only when asked (PDF/A-1) and omit it otherwise.
  const cidSetRef =
    usedGidsArr !== undefined && options.cidSet
      ? doc.add(deflatedStream({}, buildCidSet(glyphClosure(parsed, usedGidsArr))))
      : undefined;

  // ISO 32000-1 §9.6.4 / PDF/A §6.3.5 — a subsetted font's name must carry a
  // 6-uppercase-letter tag prefix ("ABCDEF+Name"), unique per subset.
  const baseFontName = usedGidsArr
    ? `${subsetTag(parsed.postScriptName, usedGidsArr)}+${parsed.postScriptName}`
    : parsed.postScriptName;

  const descriptorRef = doc.add(
    dict({
      Type: name('FontDescriptor'),
      FontName: name(baseFontName),
      Flags: parsed.flags,
      FontBBox: [
        toPdfUnits(parsed.fontBBox[0]),
        toPdfUnits(parsed.fontBBox[1]),
        toPdfUnits(parsed.fontBBox[2]),
        toPdfUnits(parsed.fontBBox[3]),
      ],
      ItalicAngle: parsed.italicAngle,
      Ascent: toPdfUnits(parsed.ascender),
      Descent: toPdfUnits(parsed.descender),
      CapHeight: toPdfUnits(parsed.capHeight),
      XHeight: toPdfUnits(parsed.xHeight),
      StemV: parsed.stemV,
      FontFile2: ref(fontFileRef.id),
      ...(cidSetRef ? { CIDSet: ref(cidSetRef.id) } : {}),
    }),
  );

  // §9.7.4.3 — /W lists the CIDs that differ from /DW, and a subset draws a
  // handful of them. One entry per glyph in the SOURCE font wrote 3388 numbers
  // for a document using fifteen: 13 KB of 55906-MultiSheetRefs.xlsx's 192.
  const widthRuns = consecutiveRuns(
    subsetGids ?? Array.from({ length: parsed.numGlyphs }, (_, i) => i),
  ).map((run) => [run[0]!, run.map((g) => toPdfUnits(parsed.advanceWidths[g] ?? 0))] as const);

  const cidFontRef = doc.add(
    dict({
      Type: name('Font'),
      Subtype: name('CIDFontType2'),
      BaseFont: name(baseFontName),
      CIDSystemInfo: dict({
        Registry: 'Adobe',
        Ordering: 'Identity',
        Supplement: 0,
      }),
      FontDescriptor: ref(descriptorRef.id),
      DW: 1000,
      W: widthRuns.flatMap(([first, ws]) => [first, ws]),
      CIDToGIDMap: name('Identity'),
    }),
  );

  const toUnicodeRef = doc.add(deflatedStream({}, buildToUnicodeCMap(parsed, subsetGids)));

  const type0Ref = doc.add(
    dict({
      Type: name('Font'),
      Subtype: name('Type0'),
      BaseFont: name(baseFontName),
      Encoding: name('Identity-H'),
      DescendantFonts: [ref(cidFontRef.id)],
      ToUnicode: ref(toUnicodeRef.id),
    }),
  );

  const measure = createFontMeasure(parsed);
  return {
    fontRef: type0Ref,
    parsed,
    pdfWidthForGid: measure.pdfWidthForGid,
    textWidthPt: measure.textWidthPt,
    encodeTextAsCidHex: measure.encodeTextAsCidHex,
  };
}

// Build a /CIDSet bit stream (ISO 19005-1 §6.3.5): bit c is set iff CID c is
// present in the subset, counting MSB-first within each byte. Length is
// ceil((maxCid + 1) / 8) bytes.
function buildCidSet(cids: ReadonlySet<number>): Uint8Array {
  let maxCid = 0;
  for (const c of cids) if (c > maxCid) maxCid = c;
  const bytes = new Uint8Array((maxCid >> 3) + 1);
  for (const c of cids) bytes[c >> 3]! |= 0x80 >> (c & 7);
  return bytes;
}

// Derive a deterministic 6-uppercase-letter subset tag from the font name and
// the set of glyphs retained. Different subsets of the same font get different
// tags (so two subsets don't collide in one file), and the same subset always
// yields the same tag (stable output).
function subsetTag(postScriptName: string, gids: ReadonlyArray<number>): string {
  let h = 0x811c9dc5;
  const mix = (n: number) => {
    h = (h ^ (n & 0xff)) >>> 0;
    h = Math.imul(h, 0x01000193) >>> 0;
  };
  for (let i = 0; i < postScriptName.length; i++) mix(postScriptName.charCodeAt(i));
  // Sort gids so ordering differences don't change the tag.
  for (const g of [...gids].sort((a, b) => a - b)) {
    mix(g);
    mix(g >> 8);
  }
  let tag = '';
  for (let i = 0; i < 6; i++) {
    tag += String.fromCharCode(65 + (h % 26));
    h = Math.floor(h / 26) + 1;
  }
  return tag;
}

/** Split a sorted id list into runs of consecutive ids. */
function consecutiveRuns(ids: ReadonlyArray<number>): Array<Array<number>> {
  const runs: Array<Array<number>> = [];
  for (const id of ids) {
    const last = runs[runs.length - 1];
    if (last && id === last[last.length - 1]! + 1) last.push(id);
    else runs.push([id]);
  }
  return runs;
}

/**
 * How much a code point says about the character a glyph stands for.
 *
 * A presentation form is a shape — the ligature two letters make, the form a
 * letter takes where it joins — but a shape that names its letters. A
 * private-use code is a font's own numbering and names nothing at all. A glyph
 * several codes reach keeps the one that says the most.
 *
 * @param cp The code point.
 * @returns 0 for a character, 1 for a presentation form, 2 for private use.
 */
function codeRank(cp: number): 0 | 1 | 2 {
  if (cp >= 0xe000 && cp <= 0xf8ff) return 2;
  // Alphabetic and Arabic presentation forms, A and B.
  if ((cp >= 0xfb00 && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff)) return 1;
  return 0;
}

// PDF spec Annex D — ToUnicode CMap.
// We scan the BMP (U+0020..U+FFFF) once and emit a bfchar entry for every
// codepoint mapped to a non-.notdef glyph. Each glyph keeps only its first
// codepoint mapping — adequate for copy-paste of the dominant script.
//
// Supplementary plane characters (U+10000+) are not yet enumerated; they
// will render correctly but won't appear in copy-paste output.
function buildToUnicodeCMap(parsed: ParsedTtf, subsetGids?: ReadonlyArray<number>): Uint8Array {
  // A subset draws a handful of glyphs, and mapping every glyph the SOURCE font
  // can reach wrote 2767 bfchar entries for a document that uses fifteen — 40 KB
  // of 55906-MultiSheetRefs.xlsx's 192. §9.10.3 asks for the CIDs that are used
  // and PDF/A §6.3.8 for every glyph that is drawn: the same set, PROVIDED every
  // drawn glyph is registered. Scoping here is what proved one was not — see the
  // comment apparatus in collectFontResources.
  const inSubset = subsetGids ? new Set(subsetGids) : undefined;
  // Map each glyph to the code point sequence it represents. Direct glyphs map
  // to a single code point; ligature glyphs (fi, ffi, …) map to their component
  // code points so text extraction recovers the original characters — required
  // for PDF/A §6.3.8 (every glyph used for rendering needs a ToUnicode value).
  const gidToCps = new Map<number, Array<number>>();
  const order: Array<number> = [];
  // Scan from U+0009 (TAB) so whitespace control glyphs that fonts map to real
  // outlines — e.g. the TAB glyph used in list markers ("1.\t") — get a
  // ToUnicode value too (PDF/A §6.3.8 needs every rendered glyph mapped).
  for (let cp = 0x0009; cp <= 0xffff; cp++) {
    if (cp >= 0xd800 && cp <= 0xdfff) continue;
    const gid = parsed.glyphForCodepoint(cp);
    if (gid === 0) continue;
    const had = gidToCps.get(gid);
    if (had) {
      // Two codes reaching one glyph: keep the one that says the most about the
      // character. Arimo's fl ligature answers both to U+FB02 and to U+F002,
      // the legacy private-use slot, and the scan runs upward — so every "fl"
      // we set went out mapped to a code that means nothing outside that one
      // font, and the text of our own page could not be searched.
      if (had.length === 1 && codeRank(cp) < codeRank(had[0]!)) gidToCps.set(gid, [cp]);
      continue;
    }
    gidToCps.set(gid, [cp]);
    order.push(gid);
  }
  // A ligature the font gives a code point of its own is still those letters:
  // "ﬂ" is a shape, and a search for "flavor" does not find "ﬂavor".
  for (const [gid, cps] of gidToCps) {
    const letters = cps.length === 1 ? lettersForLigature(cps[0]!) : undefined;
    if (letters !== undefined)
      gidToCps.set(
        gid,
        [...letters].map((c) => c.codePointAt(0)!),
      );
  }
  // Resolve ligatures to a fixpoint so chained ligatures (e.g. ffi = ff + i)
  // expand fully even if their component is itself a ligature glyph.
  let changed = true;
  while (changed) {
    changed = false;
    for (const [key, ligGid] of parsed.ligatures) {
      const have = gidToCps.get(ligGid);
      // A ligature the `cmap` also reaches directly is still a ligature: the
      // code it answers to is its SHAPE (U+FB02 is "the fl glyph"), and §9.10.3
      // asks for the characters it stands for. A search for "flavor" does not
      // find "ﬂavor".
      if (have && !(have.length === 1 && codeRank(have[0]!) > 0)) continue;
      const cps: Array<number> = [];
      let ok = true;
      for (const comp of key.split(',')) {
        const c = gidToCps.get(Number(comp));
        if (!c || c === have) {
          ok = false;
          break;
        }
        cps.push(...c);
      }
      if (!ok || cps.length === 0) continue;
      if (have?.join(',') === cps.join(',')) continue;
      gidToCps.set(ligGid, cps);
      if (!have) order.push(ligGid);
      changed = true;
    }
  }
  const pairs = order
    .filter((g) => inSubset === undefined || inSubset.has(g))
    .map((g): [number, Array<number>] => [g, gidToCps.get(g)!])
    .sort((a, b) => a[0] - b[0]);

  const utf16beHex = (cps: ReadonlyArray<number>): string => {
    let hex = '';
    for (const cp of cps) {
      if (cp <= 0xffff) {
        hex += cp.toString(16).padStart(4, '0').toUpperCase();
      } else {
        const adj = cp - 0x10000;
        hex += (0xd800 + (adj >> 10)).toString(16).padStart(4, '0').toUpperCase();
        hex += (0xdc00 + (adj & 0x3ff)).toString(16).padStart(4, '0').toUpperCase();
      }
    }
    return hex;
  };

  const lines: Array<string> = [];
  lines.push('/CIDInit /ProcSet findresource begin');
  lines.push('12 dict begin');
  lines.push('begincmap');
  lines.push('/CIDSystemInfo <</Registry (Adobe) /Ordering (UCS) /Supplement 0>> def');
  lines.push('/CMapName /Adobe-Identity-UCS def');
  lines.push('/CMapType 2 def');
  lines.push('1 begincodespacerange');
  lines.push('<0000> <FFFF>');
  lines.push('endcodespacerange');

  const CHUNK = 100;
  for (let i = 0; i < pairs.length; i += CHUNK) {
    const chunk = pairs.slice(i, i + CHUNK);
    lines.push(`${chunk.length} beginbfchar`);
    for (const [gid, cps] of chunk) {
      lines.push(`<${gid.toString(16).padStart(4, '0').toUpperCase()}> <${utf16beHex(cps)}>`);
    }
    lines.push('endbfchar');
  }

  lines.push('endcmap');
  lines.push('CMapName currentdict /CMap defineresource pop');
  lines.push('end');
  lines.push('end');

  return encoder.encode(lines.join('\n'));
}
