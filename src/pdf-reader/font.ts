// E-PDF EP2 — font resolution. Builds a ContentFont (the interpreter's decode +
// advance hooks) from a /Font dictionary: Unicode from the /ToUnicode CMap, and
// glyph advances from a simple font's /Widths or a composite font's /W array.

import { parseToUnicodeCMap } from './cmap';
import { decodePredefined, predefinedCMap, splitPredefined } from './predefined-cmap';
import { textForGlyphName } from './glyph-names';
import { standardFace, standardWidth } from './standard-widths';
import { embeddedFontName } from './embedded-fonts';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { ContentFont, Matrix, Type3Face } from './content';
import type { PdfFile } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';
import { parseTtf } from '@/core/font/ttf-parser';

/**
 * Build a {@link ContentFont} (the interpreter's decode + advance hooks) from a
 * `/Font` dictionary (E-PDF EP2). Unicode comes from the `/ToUnicode` CMap;
 * glyph advances from a simple font's `/Widths` (§9.6.2.1) or a composite
 * `/Type0` font's descendant `/W` array (§9.7.4.3). A code with no `/ToUnicode`
 * entry decodes to its Latin-1 character for a simple font, or to nothing for a
 * composite one.
 *
 * @param file     The owning {@link PdfFile}, used to resolve indirect references.
 * @param fontDict The `/Font` dictionary.
 * @returns The decode/advance hooks plus the code width (1 or 2 bytes per code).
 */
export function buildContentFont(file: PdfFile, fontDict: PdfDict): ContentFont {
  const isType0 = asName(file.resolve(fontDict.get('Subtype') ?? PDF_NULL)) === 'Type0';

  let toUnicode: ReadonlyMap<number, string> = new Map();
  let codeBytes: 1 | 2 = isType0 ? 2 : 1;
  // §9.7.5.2 — a composite font's `/Encoding` may NAME a CMap Adobe published
  // rather than embed one. Its codes are the bytes of a known encoding, and how
  // many bytes each takes is the leading byte's business: issue11555.pdf shows
  // `<6162632082a082a282a4>`, which is "abc " in one-byte codes and あいう in
  // two, and read as Identity-H it came apart into six codes of nonsense.
  const named = isType0
    ? predefinedCMap(asName(file.resolve(fontDict.get('Encoding') ?? PDF_NULL)))
    : undefined;
  const tu = file.resolve(fontDict.get('ToUnicode') ?? PDF_NULL);
  if (tu instanceof PdfStream) {
    const parsed = parseToUnicodeCMap(file.streamData(tu));
    toUnicode = parsed.map;
    // §9.6: a SIMPLE font's glyphs are always selected by one-byte codes —
    // only a composite Type0 font takes its code width from a CMap. The
    // `codespacerange` of a /ToUnicode belongs to that CMap's own convention,
    // and Distiller writes `<0000> <FFFF>` there whatever the font is. Read as
    // two, an Arial subset's every string came apart into pairs of bytes:
    // 160F-2019.pdf's "rémunérations brutes" arrived as "isr".
    if (isType0) codeBytes = parsed.codeBytes;
  }

  // §9.10.2 — a composite font that ships no `/ToUnicode` still says what its
  // glyphs are, in the font program it embeds. Read as nothing, every run in it
  // decoded to the empty string and was dropped where it stood:
  // Brotli-Prototype-FileA.pdf sets a floor plan's room names in one, and
  // "LIVING ROOM" and "DINING" never reached the page at all.
  const fromProgram = isType0 && toUnicode.size === 0 ? embeddedCmap(file, fontDict) : undefined;
  // …and a `/ToUnicode` that EXISTS may not cover the codes the page actually
  // shows. bug911034.pdf ships one describing 95 codes and then draws glyphs
  // 0x2000 upward out of a 222 KB Arial Unicode subset; every one of them
  // decoded to the empty string and the whole page came back blank. The
  // program is asked for the codes the map has no answer for — built on the
  // first miss, because walking a font's `cmap` is not free and most fonts
  // never need it.
  let programFallback: ReadonlyMap<number, string> | undefined | null = null;
  const fromProgramFor = (code: number): string | undefined => {
    if (!isType0 || toUnicode.size === 0) return undefined;
    programFallback ??= embeddedCmap(file, fontDict);
    return programFallback?.get(code);
  };
  // §9.6.6.1 — a SIMPLE font that states no `/ToUnicode` still says what its
  // codes are, in `/Encoding /Differences`: a list of glyph NAMES. A PDF from
  // TeX is nothing but this — a subset font whose codes start wherever the
  // subset does, and not a `/ToUnicode` in the file. Read as Latin-1, which is
  // all that is left without the names, issue10640.pdf's title came back as
  // "!48 SUPPORT" where it reads "LaTeX support", and its author as
  // "-OHAMED %LORABITY".
  const fromNames = !isType0 ? namedGlyphs(file, fontDict) : undefined;
  const unicode = fromProgram ?? (toUnicode.size > 0 ? toUnicode : (fromNames ?? toUnicode));

  const bytesPerCode = codeBytes;
  const style = faceStyle(file, fontDict, isType0);
  const name = embeddedFontName(file, fontDict);
  const type3 =
    asName(file.resolve(fontDict.get('Subtype') ?? PDF_NULL)) === 'Type3'
      ? type3Face(file, fontDict)
      : undefined;

  // §9.6.2.2 — a standard face need not carry a `/Widths` array at all, and
  // where it does not the reader's own metrics are the only ones there are.
  // §9.10.2 — a composite font that states NOTHING about its characters: no
  // `/ToUnicode`, and a program whose `cmap` is missing or unreadable. Its
  // codes are glyph indices and there is no way back to text from them.
  // Decoded to the empty string every run in it was dropped where it stood and
  // the page came back blank with nothing said about it —
  // issue11131_reduced.pdf is one line, "Operating Account Consolidated
  // Statement", in a subset whose program carries neither `cmap` nor `post`.
  // Marked unreadable it is stripped from the output just the same, and the
  // reconstruction reports it (see `./layout`).
  const decodeOne = (code: number): string =>
    unicode.get(code) ??
    fromProgramFor(code) ??
    (named ? decodePredefined(named, code) : undefined) ??
    // A composite code nothing could answer for is unrecoverable text, whether
    // the font stated NO map or a map that does not reach this code.
    // bug911034.pdf ships a `/ToUnicode` describing 95 codes and then draws
    // glyphs 0x2000 upward; decoded to the empty string every run was dropped
    // where it stood and the page came back blank with nothing said about it,
    // which is the one loss this reader must never take in silence.
    (bytesPerCode === 1 ? latin1(code) : '\uFFFD');
  const simple = simpleWidths(file, fontDict, decodeOne);
  // §9.6.5 — a Type 3 font states its widths in GLYPH space, which its
  // `/FontMatrix` maps to text space; every other font states them in
  // thousandths. Scaling here keeps the advance arithmetic (§9.4.4) one rule.
  const width = isType0
    ? cidWidths(file, fontDict)
    : type3
      ? (code: number) => simple(code) * type3.matrix[0] * 1000
      : simple;

  return {
    bytesPerCode,
    ...(named ? { splitCodes: (b: Uint8Array): Array<number> => splitPredefined(named, b) } : {}),
    ...(type3 ? { type3 } : {}),
    ...(name !== undefined ? { name } : {}),
    // Map each code to Unicode; an unmapped code in a simple font falls back to
    // its Latin-1 character, a composite font's to nothing (no sensible guess).
    decode: (codes) => codes.map((c) => readable(decodeOne(c))).join(''),
    width,
    ...style,
  };
}

/**
 * The character a code stands for when nothing in the font says.
 *
 * Latin-1 is the only guess there is, and it is a good one for a text font —
 * but a C0 control is not a glyph. §9.4.3 shows GLYPHS, and a code falling back
 * to one has landed there by accident: issue11549_reduced.pdf's one line came
 * back as U+0007 through U+0011 and was drawn as six empty boxes over a page
 * that shows nothing at all. A `/ToUnicode` that STATES a control is stating
 * something and is left alone.
 */
function latin1(code: number): string {
  return code < 0x20 || code === 0x7f ? '\uFFFD' : String.fromCharCode(code);
}

/**
 * What a `/ToUnicode` gives, less what Unicode says is not a character.
 *
 * `U+FFFE` and `U+FFFF` are noncharacters and the `U+FDD0`–`U+FDEF` block with
 * them; a lone surrogate is half of a pair that never came. A producer that
 * maps its glyphs to any of these has said "no text here" in the only way the
 * format lets it, and carrying them on writes bytes no reader can show —
 * arial_unicode_ab_cidfont.pdf maps its four Arabic letters to `U+FFFF` and the
 * page came back holding four of them. They become `U+FFFD`, which the
 * reconstruction counts and reports rather than passing along.
 */
function readable(text: string): string {
  let out = '';
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    const noncharacter =
      (cp & 0xfffe) === 0xfffe ||
      (cp >= 0xfdd0 && cp <= 0xfdef) ||
      // A C0 control is not a glyph. §9.4.3 shows GLYPHS, and a code that
      // decodes to one has fallen back to Latin-1 from a font that said
      // nothing: issue11549_reduced.pdf's one line came back as U+0007 through
      // U+0011 and was drawn as six empty boxes over a page that shows nothing.
      cp === 0 ||
      (cp >= SURROGATE_FIRST && cp <= SURROGATE_LAST);
    out += noncharacter ? '\uFFFD' : ch;
  }
  return out;
}

/** The last Unicode code point in the BMP, and the surrogate block inside it. */
const BMP_END = 0xffff;
const SURROGATE_FIRST = 0xd800;
const SURROGATE_LAST = 0xdfff;

/**
 * §9.10.2 — code → Unicode read out of an embedded TrueType program's `cmap`,
 * for a composite font that states no `/ToUnicode`.
 *
 * A `cmap` maps the other way, code point → glyph, so it is walked once and
 * turned round. With `Identity-H` and no `/CIDToGIDMap` a code IS a glyph
 * index, which is the case this exists for; a `/CIDToGIDMap` stream is read
 * where one is present.
 *
 * Only TrueType (`/FontFile2`) is read. A CFF program (`/FontFile3`) carries
 * its own charset and is a separate reading; a font with neither says nothing
 * about its glyphs and nothing is invented.
 */
function embeddedCmap(file: PdfFile, fontDict: PdfDict): Map<number, string> | undefined {
  const cidFont = descendantFont(file, fontDict);
  const descriptor = file.resolve(cidFont.get('FontDescriptor') ?? PDF_NULL);
  if (!(descriptor instanceof Map)) return undefined;
  const program = file.resolve(descriptor.get('FontFile2') ?? PDF_NULL);
  if (!(program instanceof PdfStream)) return undefined;
  let glyphOf: (cp: number) => number;
  try {
    glyphOf = parseTtf(file.streamData(program)).glyphForCodepoint;
  } catch {
    return undefined; // A font program we cannot read says nothing we can use.
  }
  const byGlyph = new Map<number, string>();
  for (let cp = 0x20; cp <= BMP_END; cp++) {
    if (cp >= SURROGATE_FIRST && cp <= SURROGATE_LAST) continue;
    let gid = 0;
    try {
      gid = glyphOf(cp);
    } catch {
      continue;
    }
    // The first code point to reach a glyph wins: a face maps several onto one
    // (a non-breaking space onto the space), and the first is the plainer.
    if (gid > 0 && !byGlyph.has(gid)) byGlyph.set(gid, String.fromCodePoint(cp));
  }
  if (byGlyph.size === 0) return undefined;
  const cidToGid = readCidToGid(file, cidFont);
  if (!cidToGid) return byGlyph;
  const out = new Map<number, string>();
  cidToGid.forEach((gid, cid) => {
    const text = byGlyph.get(gid);
    if (text !== undefined) out.set(cid, text);
  });
  return out.size > 0 ? out : undefined;
}

/** §9.7.4.2 `/CIDToGIDMap` — a stream of two-byte glyph indices, CID by CID. */
function readCidToGid(file: PdfFile, cidFont: PdfDict): Array<number> | undefined {
  const map = file.resolve(cidFont.get('CIDToGIDMap') ?? PDF_NULL);
  if (!(map instanceof PdfStream)) return undefined; // `/Identity`, or absent.
  const bytes = file.streamData(map);
  const out: Array<number> = [];
  for (let i = 0; i + 1 < bytes.length; i += 2) out.push((bytes[i]! << 8) | bytes[i + 1]!);
  return out;
}

/**
 * §9.6.5 — a Type 3 font's glyphs are content streams, not outlines: what the
 * face draws is whatever each procedure paints, in the resources the font
 * states. `/Encoding` `/Differences` names the procedure a code selects and
 * `/CharProcs` holds it.
 *
 * ContentStreamCycleType3insideType3.pdf is a page of them — a stroked square
 * and a stroked triangle, with a second Type 3 font shown from inside the
 * square — and with the procedures unread the page came back as two letters of
 * substituted type an eighth of an inch tall.
 */
function type3Face(file: PdfFile, fontDict: PdfDict): Type3Face | undefined {
  const procs = file.resolve(fontDict.get('CharProcs') ?? PDF_NULL);
  if (!(procs instanceof Map)) return undefined;
  const names = differences(file, fontDict);
  const resourcesVal = file.resolve(fontDict.get('Resources') ?? PDF_NULL);
  return {
    matrix: fontMatrix(file, fontDict),
    resources: resourcesVal instanceof Map ? resourcesVal : undefined,
    proc: (code) => {
      const glyph = names.get(code);
      if (glyph === undefined) return undefined;
      const stream = file.resolve(procs.get(glyph) ?? PDF_NULL);
      return stream instanceof PdfStream ? stream : undefined;
    },
  };
}

/** §9.6.5 `/FontMatrix` — glyph space to text space; a thousandth by default. */
function fontMatrix(file: PdfFile, fontDict: PdfDict): Matrix {
  const m = file.resolve(fontDict.get('FontMatrix') ?? PDF_NULL);
  if (!Array.isArray(m) || m.length !== 6) return [0.001, 0, 0, 0.001, 0, 0];
  const n = m.map((v) => asNumber(file.resolve(v), 0));
  return [n[0]!, n[1]!, n[2]!, n[3]!, n[4]!, n[5]!];
}

/**
 * §9.6.6.1 — code → text for a simple font, out of the glyph names its
 * `/Encoding /Differences` gives.
 *
 * Returns `undefined` where the font names nothing, so the caller keeps its
 * Latin-1 reading rather than replacing it with an empty map.
 */
function namedGlyphs(file: PdfFile, fontDict: PdfDict): Map<number, string> | undefined {
  const names = differences(file, fontDict);
  if (names.size === 0) return undefined;
  const out = new Map<number, string>();
  for (const [code, name] of names) {
    const text = textForGlyphName(name);
    if (text !== undefined) out.set(code, text);
  }
  if (out.size > 0) return out;
  // §9.6.5 — a TYPE 3 font's `/Encoding` is the only mapping it has: its codes
  // select CharProcs, which are drawings, and there is no standard encoding
  // underneath to fall back on. So a Type 3 face that named its glyphs and
  // whose names are not characters has unreadable text — bug1011159.pdf calls
  // its glyphs `LW010000`, and read as Latin-1 its line came back as "¦¦¦K".
  //
  // Any other font keeps the fallback: a subset TrueType commonly maps its
  // codes to `/g34`-style names that say nothing, while the codes themselves
  // are still the characters. Marking those unreadable cost TAMReview.pdf
  // eight thousand of its nine thousand words.
  if (asName(file.resolve(fontDict.get('Subtype') ?? PDF_NULL)) !== 'Type3') return undefined;
  const unreadable = new Map<number, string>();
  for (const code of names.keys()) unreadable.set(code, '\uFFFD');
  return unreadable;
}

/** §9.6.6.1 `/Encoding` `/Differences` — code → glyph name, as the array runs. */
function differences(file: PdfFile, fontDict: PdfDict): Map<number, string> {
  const out = new Map<number, string>();
  const encoding = file.resolve(fontDict.get('Encoding') ?? PDF_NULL);
  if (!(encoding instanceof Map)) return out;
  const list = file.resolve(encoding.get('Differences') ?? PDF_NULL);
  if (!Array.isArray(list)) return out;
  let code = 0;
  for (const entry of list) {
    const value = file.resolve(entry);
    if (typeof value === 'number') code = value;
    else if (value instanceof PdfName) out.set(code++, value.value);
  }
  return out;
}

/** §9.8.2 `/Flags` — bit 7 is Italic, bit 19 ForceBold (bits numbered from 1). */
const FLAG_ITALIC = 1 << 6;
const FLAG_FORCE_BOLD = 1 << 18;

/** §9.8.1 `/FontWeight` — the lightest a face may state; below it is no weight. */
const LIGHTEST_WEIGHT = 100;

/** §9.8.1 `/FontWeight` — 400 is normal, 700 bold; 600 is where "bold" begins. */
const BOLD_WEIGHT = 600;

/**
 * §9.8.1 — whether the face a run is shown in is bold or slanted.
 *
 * A descriptor is the witness where there is one, and the ONLY witness: it
 * states `/FontWeight`, `/ItalicAngle` and the `/Flags` bits, so one that gives
 * neither a weight nor the ForceBold bit is saying the face is not bold.
 * ArabicCIDTrueType.pdf shows why that matters — two of its four faces are
 * called `NewBasrahBold` and `DamascusBold`, which is the family's own name and
 * not a weight, and reading the name over the descriptor set two lines heavy
 * that no reader sets heavy.
 *
 * The name is read only where no descriptor exists at all, which is the
 * standard-14 case (§9.6.2.2): `Helvetica-BoldOblique` has nothing else to go
 * on. The subset prefix (`ISVAYD+`) is dropped first — six arbitrary capitals
 * may spell anything.
 */
function faceStyle(
  file: PdfFile,
  fontDict: PdfDict,
  isType0: boolean,
): { bold?: boolean; italic?: boolean } {
  const owner = isType0 ? descendantFont(file, fontDict) : fontDict;
  const descriptor = file.resolve(owner.get('FontDescriptor') ?? PDF_NULL);
  const named = styleFromName(asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL)));
  if (descriptor instanceof Map) {
    const flags = asNumber(file.resolve(descriptor.get('Flags') ?? PDF_NULL), 0);
    const weightVal = file.resolve(descriptor.get('FontWeight') ?? PDF_NULL);
    const slant = asNumber(file.resolve(descriptor.get('ItalicAngle') ?? PDF_NULL), 0);
    // §9.8.1 — a descriptor decides only what it STATES. One that gives no
    // /FontWeight and does not force bold has said nothing about weight, and
    // reading that silence as "regular" is how TAMReview.pdf's Times-Bold came
    // back light: every bold word on the page — its title, "Abstract",
    // "Keywords:" — set in the same face as the body.
    //
    // §9.8.2 — but ForceBold is a HINTING flag: it says whether the rasteriser
    // should thicken the stems at very small sizes, not that the face is a bold
    // cut. Where the descriptor states a weight, that weight is the answer:
    // issue10084_reduced.pdf sets "abcdefg" in a Helvetica of /FontWeight 400
    // with ForceBold set, and read off the flag the whole page came back bold.
    //
    // §9.8.1 gives the weight as one of 100…900, and a producer writing
    // anything else has written a placeholder rather than a weight:
    // issue10519_reduced.pdf states `/FontWeight 0` on a face called
    // "Calibri,Bold", and taken at its word every bold word went light.
    const stated = typeof weightVal === 'number' && weightVal >= LIGHTEST_WEIGHT;
    const bold = stated
      ? asNumber(weightVal, 0) >= BOLD_WEIGHT
      : (flags & FLAG_FORCE_BOLD) !== 0 || named.bold;
    // The slant and the flag each state italic outright; where neither does,
    // the name is the only witness left.
    const italic = slant !== 0 || (flags & FLAG_ITALIC) !== 0 || named.italic;
    return { ...(bold ? { bold: true } : {}), ...(italic ? { italic: true } : {}) };
  }
  return { ...(named.bold ? { bold: true } : {}), ...(named.italic ? { italic: true } : {}) };
}

/**
 * §9.6.2.2 — the style a font's NAME states, by the PostScript convention:
 * `Family-Style`, or `Family,Style` as Word writes it.
 *
 * The separator is what makes this safe. A family whose name merely CONTAINS
 * the word — "New Basrah Bold", "Damascus Bold", both real faces in
 * ArabicCIDTrueType.pdf — is not a bold cut of anything, and reading it as one
 * set two lines heavy that no reader sets heavy. `Times-Bold` is.
 */
function styleFromName(baseFont: string): { bold: boolean; italic: boolean } {
  // §9.6.4 — six arbitrary capitals and a plus sign mark a subset, and they may
  // spell anything at all.
  const name = baseFont.replace(/^[A-Z]{6}\+/u, '');
  const style = /[-,]([A-Za-z]+)$/u.exec(name)?.[1] ?? '';
  return {
    bold: /bold|black|heavy|semib|demi/iu.test(style),
    italic: /italic|oblique/iu.test(style),
  };
}

/** §9.7.4 — a `/Type0` font's one descendant CIDFont, which owns the descriptor. */
function descendantFont(file: PdfFile, fontDict: PdfDict): PdfDict {
  const descFonts = file.resolve(fontDict.get('DescendantFonts') ?? PDF_NULL);
  const first = Array.isArray(descFonts) ? file.resolve(descFonts[0] ?? PDF_NULL) : PDF_NULL;
  return first instanceof Map ? first : new Map<string, PdfValue>();
}

// §9.6.2.1 — a simple font's /Widths array is indexed by (code − /FirstChar).
function simpleWidths(
  file: PdfFile,
  fontDict: PdfDict,
  decodeOne: (code: number) => string,
): (code: number) => number {
  const first = asNumber(file.resolve(fontDict.get('FirstChar') ?? PDF_NULL), 0);
  const widthsVal = file.resolve(fontDict.get('Widths') ?? PDF_NULL);
  const widths = Array.isArray(widthsVal) ? widthsVal : [];
  const descriptor = file.resolve(fontDict.get('FontDescriptor') ?? PDF_NULL);
  const missing =
    descriptor instanceof Map
      ? asNumber(file.resolve(descriptor.get('MissingWidth') ?? PDF_NULL), 0)
      : 0;
  // §9.6.2.2 — the built-in metrics, for the face this one asks to be measured
  // as. Consulted only where the file itself states no width: a file that says
  // its Helvetica is 700 wide has said so, however unlike Helvetica that is.
  const face = standardFace(asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL)));
  return (code) => {
    const w = widths[code - first];
    if (typeof w === 'number') return w;
    const built = face === undefined ? undefined : standardWidth(face, code, decodeOne(code));
    if (built !== undefined) return built;
    return missing > 0 ? missing : 500;
  };
}

// §9.7.4.3 — a composite font's widths live on its descendant CIDFont as /DW
// (default) plus a /W array. With Identity encoding the CID equals the code.
function cidWidths(file: PdfFile, fontDict: PdfDict): (cid: number) => number {
  const descFonts = file.resolve(fontDict.get('DescendantFonts') ?? PDF_NULL);
  const desc0 = Array.isArray(descFonts) ? file.resolve(descFonts[0] ?? PDF_NULL) : PDF_NULL;
  const cidFont = desc0 instanceof Map ? desc0 : new Map<string, PdfValue>();
  const dw = asNumber(file.resolve(cidFont.get('DW') ?? PDF_NULL), 1000);
  const wMap = parseCidW(file, file.resolve(cidFont.get('W') ?? PDF_NULL));
  return (cid) => wMap.get(cid) ?? (dw || 1000);
}

// The /W array is a sequence of `c [w0 w1 …]` (per-CID widths from c) or
// `cFirst cLast w` (one width across a CID range).
function parseCidW(file: PdfFile, wVal: PdfValue): Map<number, number> {
  const out = new Map<number, number>();
  if (!Array.isArray(wVal)) return out;
  let i = 0;
  while (i < wVal.length) {
    const c = file.resolve(wVal[i++]!);
    if (typeof c !== 'number') break;
    const next = file.resolve(wVal[i] ?? PDF_NULL);
    if (Array.isArray(next)) {
      i++;
      next.forEach((w, k) => {
        if (typeof w === 'number') out.set(c + k, w);
      });
    } else if (typeof next === 'number') {
      i++;
      const w = file.resolve(wVal[i++] ?? PDF_NULL);
      if (typeof w === 'number') {
        for (let cc = c; cc <= next && cc - c < 65_536; cc++) out.set(cc, w);
      }
    } else {
      break;
    }
  }
  return out;
}

function asName(v: PdfValue): string {
  return v instanceof PdfName ? v.value : '';
}

function asNumber(v: PdfValue, dflt: number): number {
  return typeof v === 'number' ? v : dflt;
}
