// ISO 32000-1:2008 §9.7 — Composite (Type 0) fonts with CIDFontType2 descendants.
//
// Pipeline:
//   raw TTF  ->  /FontFile2 stream  (§9.9)
//            ->  /FontDescriptor    (§9.8)
//            ->  /CIDFontType2      (§9.7.4)  +  /W array  +  /CIDToGIDMap = Identity
//            ->  /Type0             (§9.7.2)  with /Encoding /Identity-H
//                                              and /ToUnicode CMap (Annex D)

import type { ParsedTtf } from '@/font';
import type { PdfRef } from '@/pdf/objects';
import type { PdfDocument } from '@/pdf/writer';
import { glyphClosure, shapeText, subsetTtf } from '@/font';
import { dict, name, ref, stream } from '@/pdf/objects';

const encoder = new TextEncoder();

export interface EmbedTtfOptions {
  readonly usedGids?: Iterable<number>;
  // Emit a /CIDSet in the font descriptor (required by PDF/A-1; omitted
  // otherwise, since PDF/A-2/3 only constrain it when present).
  readonly cidSet?: boolean;
}

export interface EmbeddedFont {
  readonly fontRef: PdfRef;
  readonly parsed: ParsedTtf;
  readonly pdfWidthForGid: (gid: number) => number;
  readonly textWidthPt: (text: string, fontSize: number) => number;
  readonly encodeTextAsCidHex: (text: string) => string;
}

// Measurement/encoding functions derived purely from the parsed font — no
// PdfDocument involved. The layout phase measures with these; embedTtfFont
// reuses them so emit encodes identically (ir-design stage 3b).
export interface FontMeasure {
  readonly pdfWidthForGid: (gid: number) => number;
  readonly textWidthPt: (text: string, fontSize: number) => number;
  readonly encodeTextAsCidHex: (text: string) => string;
}

export function createFontMeasure(parsed: ParsedTtf): FontMeasure {
  const scale = 1000 / parsed.unitsPerEm;
  const widths: Array<number> = new Array(parsed.numGlyphs);
  for (let i = 0; i < parsed.numGlyphs; i++) {
    widths[i] = Math.round((parsed.advanceWidths[i] ?? 0) * scale);
  }
  const pdfWidthForGid = (gid: number): number => {
    if (gid < 0 || gid >= parsed.numGlyphs) return 1000;
    return widths[gid]!;
  };
  const textWidthPt = (text: string, fontSize: number): number => {
    const shaped = shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      parsed.kerning,
      parsed.joiningForms,
    );
    let totalEm = 0;
    for (const a of shaped.advances) totalEm += a;
    return (totalEm * fontSize) / parsed.unitsPerEm;
  };
  const encodeTextAsCidHex = (text: string): string => {
    const shaped = shapeText(
      text,
      parsed.glyphForCodepoint,
      parsed.advanceWidths,
      parsed.ligatures,
      parsed.kerning,
      parsed.joiningForms,
    );
    let out = '';
    for (const gid of shaped.gids) {
      out += gid.toString(16).padStart(4, '0').toUpperCase();
    }
    return out;
  };
  return { pdfWidthForGid, textWidthPt, encodeTextAsCidHex };
}

export function embedTtfFont(
  doc: PdfDocument,
  parsed: ParsedTtf,
  options: EmbedTtfOptions = {},
): EmbeddedFont {
  const scale = 1000 / parsed.unitsPerEm;
  const toPdfUnits = (v: number): number => Math.round(v * scale);

  const usedGidsArr = options.usedGids ? [...options.usedGids] : undefined;
  const fontFileBytes = usedGidsArr ? subsetTtf(parsed, usedGidsArr) : parsed.raw;
  const fontFileRef = doc.add(stream({ Length1: fontFileBytes.byteLength }, fontFileBytes));

  // ISO 19005-1 §6.3.5 — PDF/A-1 *requires* a /CIDSet in a CIDFont subset's
  // descriptor, marking the CIDs present (Identity ordering ⇒ CID = GID, so the
  // subset's glyph closure is the CID set). PDF/A-2/3 (§6.2.11.4.2) make it
  // optional and demand an exact font-program match, which is brittle — so we
  // emit /CIDSet only when asked (PDF/A-1) and omit it otherwise.
  const cidSetRef =
    usedGidsArr !== undefined && options.cidSet
      ? doc.add(stream({}, buildCidSet(glyphClosure(parsed, usedGidsArr))))
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

  const widths: Array<number> = new Array(parsed.numGlyphs);
  for (let i = 0; i < parsed.numGlyphs; i++) {
    widths[i] = toPdfUnits(parsed.advanceWidths[i] ?? 0);
  }

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
      W: [0, widths],
      CIDToGIDMap: name('Identity'),
    }),
  );

  const toUnicodeRef = doc.add(stream({}, buildToUnicodeCMap(parsed)));

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

// PDF spec Annex D — ToUnicode CMap.
// We scan the BMP (U+0020..U+FFFF) once and emit a bfchar entry for every
// codepoint mapped to a non-.notdef glyph. Each glyph keeps only its first
// codepoint mapping — adequate for copy-paste of the dominant script.
//
// Supplementary plane characters (U+10000+) are not yet enumerated; they
// will render correctly but won't appear in copy-paste output.
function buildToUnicodeCMap(parsed: ParsedTtf): Uint8Array {
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
    if (gid === 0 || gidToCps.has(gid)) continue;
    gidToCps.set(gid, [cp]);
    order.push(gid);
  }
  // Resolve ligatures to a fixpoint so chained ligatures (e.g. ffi = ff + i)
  // expand fully even if their component is itself a ligature glyph.
  let changed = true;
  while (changed) {
    changed = false;
    for (const [key, ligGid] of parsed.ligatures) {
      if (gidToCps.has(ligGid)) continue;
      const cps: Array<number> = [];
      let ok = true;
      for (const comp of key.split(',')) {
        const c = gidToCps.get(Number(comp));
        if (!c) {
          ok = false;
          break;
        }
        cps.push(...c);
      }
      if (ok && cps.length > 0) {
        gidToCps.set(ligGid, cps);
        order.push(ligGid);
        changed = true;
      }
    }
  }
  const pairs = order
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
