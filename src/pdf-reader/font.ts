// E-PDF EP2 — font resolution. Builds a ContentFont (the interpreter's decode +
// advance hooks) from a /Font dictionary: Unicode from the /ToUnicode CMap, and
// glyph advances from a simple font's /Widths or a composite font's /W array.

import { parseToUnicodeCMap } from './cmap';
import { decodePredefined, predefinedCMap, splitPredefined } from './predefined-cmap';
import { textForGlyphName } from './glyph-names';
import {
  cffCidToGid,
  cffFsType,
  cffNameToGid,
  cffOutlineSource,
  cffSpaceAdvance,
  openTypeCff,
} from './cff-outline';
import { type1Font } from './type1-outline';
import {
  cmapSubtables,
  outlineSource,
  postGlyphNames,
  sfntFsType,
  sfntSpaceAdvance,
} from './glyf-outline';
import { standardFace, standardWidth } from './standard-widths';
import { eachPageFont, embeddedFontName, hasLiftableProgram, programStyle } from './embedded-fonts';
import { isZapfDingbats, zapfDingbatsChar } from './dingbats';
import {
  baseEncodingTable,
  isStandardLatinFace,
  macGlyphName,
  standardEncodingTable,
  winAnsiLatinName,
} from './encodings';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { ContentFont, FaceProgram, GlyphOutline, Matrix, PathSeg, Type3Face } from './content';
import type { OutlineSource } from './glyf-outline';
import type { PdfFile, PdfPage } from './document';
import type { FaceFamily } from '@/core/ir/flow';
import { knowsFamily, resolveFamilyStyle } from '@/core/fonts';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';
import { parseTtf } from '@/core/font/ttf-parser';

/** §9.10.2 — what a composite code nothing can answer for comes to. */
const UNANSWERABLE = '\uFFFD';

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
  const encodingName = asName(file.resolve(fontDict.get('Encoding') ?? PDF_NULL));
  // §9.7.5 — a name stands for a CMap Adobe published, and that CMap says how
  // codes are built. An `/Encoding` that is a STREAM is the font's own, and
  // nothing here reads one.
  const encodingNamed = isType0 && encodingName.length > 0;
  const named = isType0 ? predefinedCMap(encodingName) : undefined;
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
    //
    // §9.7.6.2 — and where the composite font NAMES its CMap, that name is what
    // says how wide a code is; a `/ToUnicode` is a second mapping and has no
    // vote. `Identity-H` is two bytes by definition, whatever else the file
    // holds: issue11549_reduced.pdf ships a `/ToUnicode` truncated mid-stream,
    // which decodes to nothing and read as one byte split every code in half —
    // each glyph came out preceded by a `.notdef` box.
    if (isType0 && !encodingNamed) codeBytes = parsed.codeBytes;
  }

  // §9.10.2 — a composite font that ships no `/ToUnicode` still says what its
  // glyphs are, in the font program it embeds. Read as nothing, every run in it
  // decoded to the empty string and was dropped where it stood:
  // Brotli-Prototype-FileA.pdf sets a floor plan's room names in one, and
  // "LIVING ROOM" and "DINING" never reached the page at all.
  const fromProgram = isType0 && toUnicode.size === 0 ? embeddedCmap(file, fontDict) : undefined;
  // …and a program with no `cmap` still says, through the `/CIDToGIDMap`,
  // whether its CIDs are CHARACTERS (see `unicodeCids`).
  const fromCids =
    isType0 && toUnicode.size === 0 && fromProgram === undefined
      ? unicodeCids(file, fontDict)
      : undefined;
  // …and a font that embeds nothing, named for a face every machine carries,
  // has its codes as glyph indices into THAT face (see `coreGlyphOrder`).
  const fromCore =
    isType0 && toUnicode.size === 0 && fromProgram === undefined && fromCids === undefined
      ? coreGlyphOrder(file, fontDict)
      : undefined;
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
  // §9.6.6 — the outlines the program holds, which decide whether a name that
  // says nothing can be DRAWN. A name nothing can draw keeps the old reading:
  // marked unreadable it would take the words away and put nothing in their
  // place, which is how bug1151216.pdf's three lines of prices vanished.
  const glyphs = isType0 ? undefined : simpleGlyphs(file, fontDict);
  // §9.6.6.4 — a glyph NAMED by its index (`g18`, `glyph152`) says nothing
  // itself, and the program it indexes still does: its `cmap` maps characters
  // onto those very glyphs, and read backwards it says what each one is.
  // TAMReview.pdf sets its body in a Cambria subset named that way, and seven
  // thousand of its nine thousand characters were traced as drawings instead
  // of read — forty thousand shapes where a page of text stands.
  let byIndex: ReadonlyMap<number, string> | undefined | null = null;
  const numbering = isType0 ? undefined : numberingOf(differences(file, fontDict).values());
  const fromIndex = (name: string): string | undefined => {
    const gid = numberedGlyph(name, numbering);
    if (gid === undefined) return undefined;
    byIndex ??= programCharacters(file, fontDict);
    return byIndex?.get(gid);
  };
  const fromNames = !isType0
    ? namedGlyphs(file, fontDict, toUnicode.size > 0, {
        draws: (name) => glyphs?.byName(name) !== undefined,
        blank: (name) => glyphs?.blank(name) === true,
        character: fromIndex,
      })
    : undefined;
  // The names themselves, not what they come to: a name that is no character
  // still selects a glyph, which is what the outline path draws.
  const glyphNames = isType0 ? new Map<number, string>() : differences(file, fontDict);
  const unicode =
    fromProgram ??
    fromCids ??
    fromCore ??
    (toUnicode.size > 0 ? toUnicode : (fromNames ?? toUnicode));

  const bytesPerCode = codeBytes;
  // §9.6.6.4 — a simple TrueType whose program has NO `cmap`, and which names
  // nothing: `/Differences` absent, `/ToUnicode` absent. Nothing in the file
  // maps a code to a character, and nothing in the program can be reached by
  // one — so the codes are glyph INDICES, and every reading of them as text is
  // invention. TrueType_without_cmap.pdf draws four Armenian letters and came
  // back "'>in", which is what its indices happen to spell in Latin-1.
  const indices =
    !isType0 &&
    glyphs?.indexed === true &&
    fromNames === undefined &&
    toUnicode.size === 0 &&
    glyphNames.size === 0;
  // §9.6.2.2 — a standard face whose own encoding is not the Latin one.
  const dingbats =
    !isType0 && isZapfDingbats(asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL)));
  // Annex D.2 — the encoding the codes are read through under /Differences.
  // Annex D.2 — the base encoding as glyph NAMES, which serve twice: the text a
  // code stands for, and the glyph it selects in a program addressed by name.
  //
  // §9.6.6.1 — and a font that states no encoding is read through the one its
  // program is built with. A NAME that is none of the encodings there are
  // states nothing: bug859204.pdf embeds a Type 1 News Gothic under
  // `/Encoding /NULL`, whose code 0x95 its program names `bullet` — read as
  // Latin-1 it was a control character, and the list lost its bullet.
  const stated = file.resolve(fontDict.get('Encoding') ?? PDF_NULL);
  const statesEncoding =
    stated instanceof Map ||
    (stated instanceof PdfName && baseEncodingTable(stated.value) !== undefined);
  const baseNames = isType0
    ? undefined
    : statesEncoding
      ? baseEncoding(file, fontDict)
      : (glyphs?.builtIn ?? baseEncoding(file, fontDict));
  const fromBase = new Map<number, string>();
  for (const [code, glyph] of baseNames ?? []) {
    const text = textForGlyphName(glyph);
    if (text !== undefined) fromBase.set(code, text);
  }
  // §9.6.6 — the name a code selects, `/Differences` first and the base
  // encoding under it. A legacy eight-bit face states nothing but the base one,
  // and its glyphs are reached through the program's `post` table.
  const namesOf = new Map<number, string>([...(baseNames ?? []), ...glyphNames]);
  const style = faceStyle(file, fontDict, isType0);
  const name = runFontName(file, fontDict, isType0);
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
    (indices ? UNANSWERABLE : undefined) ??
    unicode.get(code) ??
    fromProgramFor(code) ??
    (named ? decodePredefined(named, code) : undefined) ??
    // §9.10.2 — the `/Encoding`'s glyph NAMES, for the codes a `/ToUnicode`
    // does not reach. A map that covers a font's letters and not its space is
    // ordinary, and read as Latin-1 the space is a control and is dropped:
    // TAMReview.pdf's figure labels came back "SystemFeatures".
    fromNames?.get(code) ??
    // §9.6.6.1 — a code the font's own `/Differences` name `.notdef` selects
    // no glyph and is no character, whatever the base encoding would have
    // made of it. issue11403_reduced.pdf writes a UTF-8 no-break space into
    // a Helvetica that names both its bytes `.notdef`, and read through
    // StandardEncoding the first came back an acute accent before the line.
    (glyphNames.get(code) === NOTDEF ? '' : undefined) ??
    // Annex D.6 — the built-in encoding of a standard face that has one of its
    // own. ZapfDingbats is a font of PICTURES: its 0x4B is not the letter K but
    // `a38`, the six-pointed star, and read through the Latin encoding below —
    // which is all a reader can do for a font that states none —
    // ZapfDingbats.pdf's five hundred pictures came back as the alphabet.
    (dingbats ? zapfDingbatsChar(code) : undefined) ??
    // Annex D.2 — the BASE encoding, which is the font's own reading of every
    // code `/Differences` does not restate. Latin-1 below is a good guess for a
    // text font and it is only a guess: 0xD0 is Eth there and an em dash in
    // StandardEncoding, and ZapfDingbats.pdf's title — Times, no /Encoding at
    // all — came back as "Character Sets Ð Zapf Dingbats".
    fromBase.get(code) ??
    // A composite code nothing could answer for is unrecoverable text, whether
    // the font stated NO map or a map that does not reach this code.
    // bug911034.pdf ships a `/ToUnicode` describing 95 codes and then draws
    // glyphs 0x2000 upward; decoded to the empty string every run was dropped
    // where it stood and the page came back blank with nothing said about it,
    // which is the one loss this reader must never take in silence.
    (bytesPerCode === 1 ? latin1(code) : UNANSWERABLE);
  // §9.7.4 — a composite font's program, read by glyph once for both uses.
  const composite = isType0 ? compositeGlyphs(file, fontDict) : undefined;
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
    // §9.6.6 — where a code answers with nothing, its SHAPE is still in the
    // file. Drawn, the page shows what it showed; left out, it is a blank
    // sheet. Only for a code that has no character: a face this reads is set
    // as type, not traced.
    ...(isType0 ? outlineOf(composite, decodeOne) : simpleOutlineOf(decodeOne, namesOf, glyphs)),
    // §9.9 — and the program itself, for a writer that embeds the face. Only a
    // face a run can name, and never a Type 3 one, whose glyphs are drawings.
    ...(name !== undefined && !type3
      ? programOf(
          file,
          fontDict,
          isType0,
          isType0
            ? compositeGlyph(composite, fontDict, file)
            : simpleGlyph(glyphs, namesOf, baseNames === baseEncodingTable('WinAnsiEncoding')),
          isType0 ? composite : glyphs,
        )
      : {}),
    ...(named?.vertical ? { verticalAdvance: cidVerticalAdvance(file, fontDict) } : {}),
    ...(type3 ? { type3 } : {}),
    ...(name !== undefined ? { name } : {}),
    // Map each code to Unicode; an unmapped code in a simple font falls back to
    // its Latin-1 character, a composite font's to nothing (no sensible guess).
    decode: (codes) => codes.map((c) => lettersOf(readable(decodeOne(c)))).join(''),
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

/**
 * Unicode's presentation forms of the Latin ligatures, U+FB00–U+FB06, as the
 * letters they join.
 *
 * A producer maps a ligature glyph to its compatibility character, and carried
 * as that the word neither searches nor edits as the letters it is — and a
 * face with no such form sets it in another's, wider: copy_paste_ligatures.pdf
 * is one line in Times, and its "ﬀﬁﬂﬃﬄﬅﬆ" pushed the line off its page. The long
 * s of `ﬅ` is kept: it is a letter of its own.
 */
function lettersOf(text: string): string {
  return LIGATURE.test(text) ? text.replace(LIGATURES, (c) => LATIN_LIGATURES.get(c) ?? c) : text;
}

const LIGATURE = /[\uFB00-\uFB06]/u;
const LIGATURES = /[\uFB00-\uFB06]/gu;
const LATIN_LIGATURES: ReadonlyMap<string, string> = new Map([
  ['\uFB00', 'ff'],
  ['\uFB01', 'fi'],
  ['\uFB02', 'fl'],
  ['\uFB03', 'ffi'],
  ['\uFB04', 'ffl'],
  ['\uFB05', '\u017Ft'],
  ['\uFB06', 'st'],
]);

/** The glyph name that selects no glyph (§9.6.6.1). */
const NOTDEF = '.notdef';

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
  const byGlyph = programCharacters(file, cidFont);
  if (byGlyph === undefined) return undefined;
  const cidToGid = readCidToGid(file, cidFont);
  if (!cidToGid) return byGlyph;
  const out = new Map<number, string>();
  cidToGid.forEach((gid, cid) => {
    const text = byGlyph.get(gid);
    if (text !== undefined) out.set(cid, text);
  });
  return out.size > 0 ? out : undefined;
}

/**
 * §9.7.4 — the characters of a composite font that embeds no program, names a
 * core face and states no `/ToUnicode`: its codes are glyph indices into the
 * face the reader's machine carries under that name, and the core faces put
 * their first 258 glyphs in the standard Macintosh order (the `post` table's
 * format 1 names).
 *
 * issue11242_reduced.pdf shows "VAT Code" in a Verdana it does not embed.
 * Nothing in the file says what its codes are but the face's own order, and
 * read as characters they are "9$7&RGH"; read as nothing, the page came back
 * blank. Only faces known to keep that order are read this way — a face's
 * order is its own, and Calibri's is not this one.
 *
 * @param file     The owning file.
 * @param fontDict The `/Type0` font dictionary.
 * @returns Code → character, or `undefined` where the font is not such a one.
 */
function coreGlyphOrder(file: PdfFile, fontDict: PdfDict): Map<number, string> | undefined {
  const encoding = asName(file.resolve(fontDict.get('Encoding') ?? PDF_NULL));
  if (encoding !== 'Identity-H' && encoding !== 'Identity-V') return undefined;
  const cidFont = descendantFont(file, fontDict);
  if (asName(file.resolve(cidFont.get('Subtype') ?? PDF_NULL)) !== 'CIDFontType2') return undefined;
  const map = file.resolve(cidFont.get('CIDToGIDMap') ?? PDF_NULL);
  if (map !== PDF_NULL && !(map instanceof PdfName && map.value === 'Identity')) return undefined;
  const descriptor = file.resolve(cidFont.get('FontDescriptor') ?? PDF_NULL);
  if (
    descriptor instanceof Map &&
    ['FontFile', 'FontFile2', 'FontFile3'].some((k) => descriptor.has(k))
  )
    return undefined;
  const family = familyOfFace(asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL)));
  if (!CORE_FACES.has(family.toLowerCase().replace(/[^a-z]/gu, ''))) return undefined;
  const out = new Map<number, string>();
  for (let gid = 0; gid < MAC_GLYPHS; gid++) {
    const name = macGlyphName(gid);
    const text = name === undefined ? undefined : textForGlyphName(name);
    if (text !== undefined && text !== '') out.set(gid, text);
  }
  return out;
}

/** The core faces whose first glyphs stand in the standard Macintosh order. */
const CORE_FACES: ReadonlySet<string> = new Set([
  'arial',
  'timesnewroman',
  'couriernew',
  'verdana',
  'tahoma',
  'georgia',
  'trebuchetms',
]);

/** How many glyphs the standard Macintosh order names. */
const MAC_GLYPHS = 258;

/**
 * §9.7.4.2 — the CIDs of a composite font that ARE characters.
 *
 * A `/CIDToGIDMap` stream says the CID is not the glyph index, and TCPDF —
 * with tFPDF and mPDF after it — makes it the character's own Unicode value,
 * mapping every character of the face to its glyph whether the subset kept
 * the glyph or not. Without a `/ToUnicode`, and with a subset that carries no
 * `cmap`, that was the only statement of what the glyphs are and it went
 * unread: bug1650302_reduced.pdf's "Výbava na přání" came back as drawings,
 * the `ř` missing from them. Read as characters, the words come back.
 *
 * Only where the map itself says so. A producer that keeps the ORIGINAL
 * glyph index as the CID routes it through a stream too, and read as Unicode
 * complex_ttf_font.pdf's Arabic would come back as `$&')`. What tells the two
 * apart is the space: a map of characters sends U+0020 to a glyph that draws
 * nothing and no control character anywhere, where a map of indices sends 3 —
 * the space in the fonts they come from — and 32 to a glyph with ink on it.
 */
function unicodeCids(file: PdfFile, fontDict: PdfDict): Map<number, string> | undefined {
  const encoding = asName(file.resolve(fontDict.get('Encoding') ?? PDF_NULL));
  if (encoding !== 'Identity-H' && encoding !== 'Identity-V') return undefined;
  const cidFont = descendantFont(file, fontDict);
  const cidToGid = readCidToGid(file, cidFont);
  if (!cidToGid || cidToGid.slice(0, SPACE).some((gid) => gid !== 0)) return undefined;
  const space = cidToGid[SPACE];
  if (space === undefined || space === 0) return undefined;
  const descriptor = file.resolve(cidFont.get('FontDescriptor') ?? PDF_NULL);
  if (!(descriptor instanceof Map)) return undefined;
  const program = file.resolve(descriptor.get('FontFile2') ?? PDF_NULL);
  if (!(program instanceof PdfStream)) return undefined;
  let blank: boolean;
  try {
    const glyf = outlineSource(file.streamData(program));
    blank = glyf !== undefined && space < glyf.count && glyf.path(space) === undefined;
  } catch {
    return undefined;
  }
  if (!blank) return undefined;
  const out = new Map<number, string>();
  cidToGid.forEach((gid, cid) => {
    if (gid !== 0 && !(cid >= SURROGATE_FIRST && cid <= SURROGATE_LAST)) {
      out.set(cid, String.fromCharCode(cid));
    }
  });
  return out;
}

/** U+0020, the one character whose glyph every text face leaves blank. */
const SPACE = 0x20;

/**
 * The character each glyph of a font's embedded TrueType program stands for:
 * its `cmap` read backwards.
 *
 * @param file  The document.
 * @param owner The dictionary that holds the descriptor — the font itself, or
 *              a composite font's descendant.
 * @returns Glyph index → character, or `undefined` where there is no program
 *          or it maps nothing.
 */
function programCharacters(file: PdfFile, owner: PdfDict): Map<number, string> | undefined {
  const descriptor = file.resolve(owner.get('FontDescriptor') ?? PDF_NULL);
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
  return byGlyph.size > 0 ? byGlyph : undefined;
}

/**
 * §9.7.4 — a composite font's embedded program, read by glyph: the code is a
 * CID, `/CIDToGIDMap` (or, for a CID-keyed CFF, its charset) turns that into a
 * glyph index, and the program holds the contours.
 *
 * @param file     The owning file.
 * @param fontDict The Type 0 font dictionary.
 * @returns The program's outlines and how a CID reaches them, or `undefined`
 *          where the font embeds no program this reads.
 */
function compositeGlyphs(file: PdfFile, fontDict: PdfDict): CompositeGlyphs | undefined {
  const cidFont = descendantFont(file, fontDict);
  const descriptor = file.resolve(cidFont.get('FontDescriptor') ?? PDF_NULL);
  if (!(descriptor instanceof Map)) return undefined;
  const truetype = file.resolve(descriptor.get('FontFile2') ?? PDF_NULL);
  const compact = file.resolve(descriptor.get('FontFile3') ?? PDF_NULL);
  const program = truetype instanceof PdfStream ? truetype : compact;
  if (!(program instanceof PdfStream)) return undefined;
  let source: OutlineSource | undefined;
  let charsetCids: Map<number, number> | undefined;
  let fsType: number | undefined;
  let spaceAdvance: number | undefined;
  try {
    const bytes = file.streamData(program);
    // §9.9 — `/FontFile3` is a CFF program, either bare or wrapped in an
    // OpenType shell; the shell may carry TrueType outlines instead.
    source = outlineSource(bytes);
    fsType = sfntFsType(bytes);
    spaceAdvance = sfntSpaceAdvance(bytes);
    if (!source) {
      const cff = openTypeCff(bytes) ?? bytes;
      source = cffOutlineSource(cff);
      // TN 5176 §10 — a CID-keyed CFF is ordered by CID, so the code is not
      // the index into its charstrings; its charset says which glyph is which.
      charsetCids = source ? cffCidToGid(cff) : undefined;
      fsType ??= cffFsType(cff);
    }
  } catch {
    return undefined;
  }
  if (!source) return undefined;
  const cidToGid = readCidToGid(file, cidFont);
  return {
    source,
    glyphOf: (cid) => {
      const mapped = cidToGid ? cidToGid[cid] : cid;
      if (mapped === undefined) return undefined;
      return charsetCids ? charsetCids.get(mapped) : mapped;
    },
    looseGlyphOf: (cid) => {
      const mapped = cidToGid ? (cidToGid[cid] ?? 0) : cid;
      return charsetCids ? (charsetCids.get(mapped) ?? mapped) : mapped;
    },
    ...(fsType !== undefined ? { fsType } : {}),
    ...(spaceAdvance !== undefined ? { spaceAdvance } : {}),
  };
}

/** A composite font's program, and the ways a CID reaches its glyphs. */
interface CompositeGlyphs extends ProgramFacts {
  readonly source: OutlineSource;
  /** The glyph a CID selects, or `undefined` where the maps name none for it. */
  readonly glyphOf: (cid: number) => number | undefined;
  /**
   * The same, falling back to the CID itself where the maps are silent — the
   * last resort of a code that otherwise draws nothing at all.
   */
  readonly looseGlyphOf: (cid: number) => number;
}

/**
 * §9.6.6 — the outline a code draws, for a code that stands for no character.
 *
 * The glyph is there even when the character is not: `/Encoding /Identity-H`
 * makes the code a CID, `/CIDToGIDMap` turns that into a glyph index, and the
 * embedded program holds the contours. complex_ttf_font.pdf is eight lines of
 * Arabic in a subset with no `cmap` and no `/ToUnicode`, and every one of them
 * was dropped.
 *
 * Deliberately NOT a fallback for text: a code the font CAN answer for is set
 * as type, and only the unanswerable ones are traced.
 *
 * @param glyphs    The font's program, read by glyph (see `compositeGlyphs`).
 * @param decodeOne What one code comes to, to tell the two cases apart.
 * @returns The `outline` field of a {@link ContentFont}, or nothing where the
 *          program carries no outlines this reads.
 */
function outlineOf(
  glyphs: CompositeGlyphs | undefined,
  decodeOne: (code: number) => string,
): { outline?: GlyphOutline } {
  if (!glyphs) return {};
  return {
    outline: {
      // The reader gives a one-unit em, which is a `/FontMatrix` of 1/upem
      // already applied — so glyph space to text space is the identity.
      matrix: [1, 0, 0, 1, 0, 0],
      path: (code: number): Array<PathSeg> | undefined => {
        // The same test the text takes: a `/ToUnicode` that STATES a
        // noncharacter has said "no text here" as plainly as one that says
        // nothing at all. arial_unicode_ab_cidfont.pdf maps its four Arabic
        // letters to U+FFFF.
        if (readable(decodeOne(code)) !== UNANSWERABLE) return undefined;
        return glyphs.source.path(glyphs.looseGlyphOf(code));
      },
    },
  };
}

/**
 * §9.7.5.2 — the glyph ANY code of a composite font draws, for a writer that
 * embeds the face.
 *
 * Only under `Identity-H` or `Identity-V`, where the code is the CID: any other
 * CMap turns codes into CIDs by a table of its own, and a glyph looked up by
 * the code there would be some other character's.
 */
function compositeGlyph(
  glyphs: CompositeGlyphs | undefined,
  fontDict: PdfDict,
  file: PdfFile,
): ((code: number) => ReadonlyArray<PathSeg> | undefined) | undefined {
  const encoding = asName(file.resolve(fontDict.get('Encoding') ?? PDF_NULL));
  if (!glyphs || (encoding !== 'Identity-H' && encoding !== 'Identity-V')) return undefined;
  return (code) => {
    const gid = glyphs.glyphOf(code);
    // Glyph 0 is `.notdef`: the program has nothing for the code.
    if (gid === undefined || gid <= 0 || gid >= glyphs.source.count) return undefined;
    return glyphs.source.path(gid) ?? [];
  };
}

/**
 * §9.6.6 — the outline a SIMPLE font's code draws, for a code that stands for
 * no character.
 *
 * A simple font addresses its glyphs by NAME: `/Differences` (or the program's
 * own `/Encoding`) says code 2 is `g18`, and the program says what `g18` looks
 * like. Where that name is not a character there is nothing to write and there
 * is still something to draw — TAMReview.pdf sets most of its body in a Cambria
 * subset whose glyphs are named `g18`, `g152`, `g135`, and seven thousand of
 * its nine thousand characters were dropped.
 *
 * @param file      The owning file.
 * @param fontDict  The simple font's dictionary.
 * @param decodeOne What one code comes to, to tell the two cases apart.
 * @param nameOf    The glyph name a code selects, where the font states one.
 * @returns The `outline` field of a {@link ContentFont}, or nothing.
 */
function simpleOutlineOf(
  decodeOne: (code: number) => string,
  nameOf: ReadonlyMap<number, string>,
  source: SimpleGlyphs | undefined,
): { outline?: GlyphOutline } {
  if (!source) return {};
  return {
    outline: {
      // Every reader gives a one-unit em, so glyph space to text space is the
      // identity.
      matrix: [1, 0, 0, 1, 0, 0],
      path: (code: number): Array<PathSeg> | undefined => {
        if (readable(decodeOne(code)) !== UNANSWERABLE) return undefined;
        const name = nameOf.get(code) ?? source.builtIn?.get(code);
        // A program that names nothing and maps nothing is addressed the only
        // way that is left: by index, which is what the code is.
        if (name === undefined) return source.byIndex?.(code);
        return source.byName(name);
      },
    },
  };
}

/** The outlines a simple font's program holds, addressed the way it addresses them. */
interface SimpleGlyphs extends ProgramFacts {
  readonly byName: (name: string) => Array<PathSeg> | undefined;
  /**
   * Whether the program HOLDS that glyph and it draws nothing. A blank glyph
   * in a text font is a space, not a character the reader lost: TAMReview.pdf
   * names its space `g3` and, dropped, its figure labels came back as
   * "SystemFeatures".
   */
  readonly blank: (name: string) => boolean;
  /** §5 — a Type 1 program's own `/Encoding`, where the file states none. */
  readonly builtIn?: ReadonlyMap<number, string>;
  /** The glyph at an INDEX, for a program whose codes are indices. */
  readonly byIndex?: (gid: number) => Array<PathSeg> | undefined;
  /**
   * Whether the program can be reached by character at all. A TrueType with no
   * `cmap` cannot (§9.6.6.4): its codes are glyph indices, and every reading of
   * them as text is invention.
   */
  readonly indexed?: boolean;
  /**
   * §9.6.6 — the glyph a code draws, reached the way the program is: by the
   * name the encoding gives the code, or — for a TrueType program — through
   * whichever `cmap` it carries. Empty for a blank glyph, `undefined` where the
   * program holds none for the code.
   */
  readonly glyph: (code: number, name: string | undefined) => ReadonlyArray<PathSeg> | undefined;
}

/**
 * §9.6.6 — the program a SIMPLE font embeds, ready to draw a glyph BY NAME.
 *
 * All three formats appear here: a Type 1 program keys its charstrings by name
 * outright, a CFF says which name each glyph has in its charset, and a TrueType
 * says nothing at all — but a subsetter that renames glyphs `g24` has written
 * the glyph's index into the name, which is the only handle such a program
 * gives.
 *
 * @param file     The owning file.
 * @param fontDict The simple font's dictionary.
 * @returns How to draw one of its glyphs, or `undefined` where nothing here
 *          can read the program.
 */
function simpleGlyphs(file: PdfFile, fontDict: PdfDict): SimpleGlyphs | undefined {
  const descriptor = file.resolve(fontDict.get('FontDescriptor') ?? PDF_NULL);
  if (!(descriptor instanceof Map)) return undefined;
  const typeOne = file.resolve(descriptor.get('FontFile') ?? PDF_NULL);
  const truetype = file.resolve(descriptor.get('FontFile2') ?? PDF_NULL);
  const compact = file.resolve(descriptor.get('FontFile3') ?? PDF_NULL);
  const numbering = numberingOf(differences(file, fontDict).values());
  try {
    if (typeOne instanceof PdfStream) {
      const face = type1Font(file.streamData(typeOne));
      if (!face) return undefined;
      return {
        byName: face.path,
        blank: (name: string): boolean => face.has(name) && face.path(name) === undefined,
        ...(face.encoding ? { builtIn: face.encoding } : {}),
        glyph: (_code, name) =>
          name === undefined || name === NOTDEF || !face.has(name)
            ? undefined
            : (face.path(name) ?? []),
        ...(face.fsType !== undefined ? { fsType: face.fsType } : {}),
        ...spaceOf(face.advance('space')),
      };
    }
    const stream = compact instanceof PdfStream ? compact : truetype;
    if (!(stream instanceof PdfStream)) return undefined;
    const bytes = file.streamData(stream);
    // A TrueType program, or the TrueType half of an OpenType shell.
    const glyf = outlineSource(bytes);
    const fsType = sfntFsType(bytes);
    if (glyf) {
      // §post — the names the program itself gives its glyphs, which is how a
      // legacy eight-bit face is reached: it has no `cmap`, and its shapes sit
      // under Latin names.
      const named = postGlyphNames(bytes);
      const gidOf = (name: string): number | undefined => {
        const gid = named?.get(name) ?? numberedGlyph(name, numbering);
        return gid !== undefined && gid < glyf.count ? gid : undefined;
      };
      const flags = asNumber(file.resolve(descriptor.get('Flags') ?? PDF_NULL), 0);
      const encoding = file.resolve(fontDict.get('Encoding') ?? PDF_NULL);
      const byCode = trueTypeGlyphs(
        bytes,
        (flags & FLAG_SYMBOLIC) !== 0 && (flags & FLAG_NONSYMBOLIC) === 0,
        encoding instanceof Map || encoding instanceof PdfName,
        gidOf,
      );
      return {
        glyph: (code, name) => {
          const gid = byCode(code, name);
          if (gid === undefined || gid <= 0 || gid >= glyf.count) return undefined;
          return glyf.path(gid) ?? [];
        },
        ...(fsType !== undefined ? { fsType } : {}),
        ...spaceOf(sfntSpaceAdvance(bytes)),
        byName: (name: string): Array<PathSeg> | undefined => {
          const gid = gidOf(name);
          return gid === undefined ? undefined : glyf.path(gid);
        },
        blank: (name: string): boolean => {
          const gid = gidOf(name);
          return gid !== undefined && glyf.path(gid) === undefined;
        },
        byIndex: (gid: number): Array<PathSeg> | undefined =>
          gid < glyf.count ? glyf.path(gid) : undefined,
        // §9.6.6.4 — with no `cmap` the program cannot be reached by character
        // at all, which is the file saying what its codes are.
        indexed: !glyf.cmap,
      };
    }
    const cff = openTypeCff(bytes) ?? bytes;
    const outlines = cffOutlineSource(cff);
    if (!outlines) return undefined;
    const names = cffNameToGid(cff);
    const gidOf = (name: string): number | undefined => {
      const gid = names?.get(name) ?? numberedGlyph(name, numbering);
      return gid !== undefined && gid < outlines.count ? gid : undefined;
    };
    const stated = fsType ?? cffFsType(cff);
    return {
      glyph: (_code, name) => {
        const gid = name === undefined || name === NOTDEF ? undefined : gidOf(name);
        return gid === undefined || gid === 0 ? undefined : (outlines.path(gid) ?? []);
      },
      ...(stated !== undefined ? { fsType: stated } : {}),
      ...spaceOf(sfntSpaceAdvance(bytes) ?? cffSpaceAdvance(cff)),
      byName: (name: string): Array<PathSeg> | undefined => {
        const gid = gidOf(name);
        return gid === undefined ? undefined : outlines.path(gid);
      },
      blank: (name: string): boolean => {
        const gid = gidOf(name);
        return gid !== undefined && outlines.path(gid) === undefined;
      },
    };
  } catch {
    return undefined;
  }
}

/** The `spaceAdvance` field, where a program states a sensible one. */
function spaceOf(advance: number | undefined): { spaceAdvance?: number } {
  return advance !== undefined && advance > 0 && advance < 1000 ? { spaceAdvance: advance } : {};
}

/**
 * §9.6.6.4 — how a SIMPLE TrueType font's code reaches a glyph: through the
 * program's `cmap`, and which subtable decides what the code is looked up as.
 *
 * A face that states an encoding turns the code into its glyph NAME, and the
 * name into a character for the Unicode subtable (3,1), a Mac Roman code for
 * (1,0), or an entry of the program's own `post` names — and past ASCII, only
 * that: the same byte is another letter in Mac Roman, and a glyph found by the
 * code there is the wrong one. A symbolic face with a Windows symbol subtable
 * (3,0), and a face that states no encoding, is looked up by the CODE: in
 * (3,0), where it stands in one of four ranges (itself, or behind 0xF000,
 * 0xF100 or 0xF200), or byte for byte in (1,0) — its names the fallback, as a
 * viewer tries them. A program with no `cmap` at all is read by index: the
 * code IS the glyph.
 *
 * @param program     The raw `/FontFile2` bytes.
 * @param symbolic    Whether `/Flags` call the face symbolic (bit 3, not bit 6).
 * @param hasEncoding Whether the font dictionary states an `/Encoding`.
 * @param byName      The glyph a name selects through `post` or the
 *                    numbered-name convention (see {@link numberedGlyph}).
 * @returns Code (and the name the encoding gives it) → glyph index, or
 *          `undefined` where nothing reaches one.
 */
function trueTypeGlyphs(
  program: Uint8Array,
  symbolic: boolean,
  hasEncoding: boolean,
  byName: (name: string) => number | undefined,
): (code: number, name: string | undefined) => number | undefined {
  let tables: Map<string, (code: number) => number> | undefined;
  return (code, name) => {
    tables ??= cmapSubtables(program);
    const unicode =
      tables.get('3,1') ?? tables.get('0,3') ?? tables.get('0,1') ?? tables.get('0,0');
    const symbol = tables.get('3,0');
    const roman = tables.get('1,0');
    const bySymbol = (): number => {
      if (symbol) {
        for (const high of SYMBOL_RANGES) {
          const gid = symbol(high | code);
          if (gid > 0) return gid;
        }
        return 0;
      }
      return roman ? roman(code) : 0;
    };
    const byGlyphName = (): number => {
      if (name === undefined || name === NOTDEF) return 0;
      const text = textForGlyphName(name);
      const cp = text !== undefined && [...text].length === 1 ? text.codePointAt(0) : undefined;
      const viaUnicode = unicode && cp !== undefined ? unicode(cp) : 0;
      if (viaUnicode > 0) return viaUnicode;
      const mac = macRomanCode(name);
      const viaRoman = roman && mac !== undefined ? roman(mac) : 0;
      if (viaRoman > 0) return viaRoman;
      return byName(name) ?? 0;
    };
    const gid =
      symbolic && symbol !== undefined
        ? bySymbol() || byGlyphName()
        : hasEncoding
          ? byGlyphName() || (code < ASCII_END ? bySymbol() : 0)
          : // A Unicode subtable asked with the code itself: what a face with
            // no encoding of its own leaves to be tried last.
            bySymbol() || byGlyphName() || (unicode ? unicode(code) : 0);
    if (gid > 0) return gid;
    // §9.6.6.4 — no `cmap` at all: the codes are glyph indices.
    return tables.size === 0 ? code : undefined;
  };
}

/** Below this, a byte is the same character in every encoding a simple font uses. */
const ASCII_END = 0x80;

/** §9.6.6.4 — where a symbolic face's (3,0) subtable may put a one-byte code. */
const SYMBOL_RANGES: ReadonlyArray<number> = [0x0000, 0xf000, 0xf100, 0xf200];

/** §9.8.2 `/Flags` — bit 3 is Symbolic, bit 6 Nonsymbolic (bits numbered from 1). */
const FLAG_SYMBOLIC = 1 << 2;
const FLAG_NONSYMBOLIC = 1 << 5;

/** The Mac Roman code of a glyph name, which is what a (1,0) `cmap` is keyed by. */
function macRomanCode(name: string): number | undefined {
  if (!macRomanCodes) {
    macRomanCodes = new Map();
    for (const [code, glyph] of baseEncodingTable('MacRomanEncoding') ?? []) {
      if (!macRomanCodes.has(glyph)) macRomanCodes.set(glyph, code);
    }
  }
  return macRomanCodes.get(name);
}

let macRomanCodes: Map<string, number> | undefined;

/**
 * §9.6.6 — the glyph ANY code of a simple font draws, for a writer that embeds
 * the face: the name the encoding gives the code, as the outline path reads it.
 *
 * WinAnsiEncoding's Latin-1 half is named too (Annex D.2), though the text
 * reads it as Latin-1 without a name: unnamed, an é in a program addressed by
 * name selected nothing — and in a TrueType one, the Mac Roman subtable's È.
 *
 * @param glyphs  The program, read by name and code.
 * @param nameOf  The name each code has: `/Differences` over the base encoding.
 * @param winAnsi Whether that base encoding is WinAnsiEncoding.
 */
function simpleGlyph(
  glyphs: SimpleGlyphs | undefined,
  nameOf: ReadonlyMap<number, string>,
  winAnsi: boolean,
): ((code: number) => ReadonlyArray<PathSeg> | undefined) | undefined {
  if (!glyphs) return undefined;
  return (code) =>
    glyphs.glyph(
      code,
      nameOf.get(code) ??
        (winAnsi ? winAnsiLatinName(code) : undefined) ??
        glyphs.builtIn?.get(code),
    );
}

/** What a font program states of itself, beside its glyphs. */
interface ProgramFacts {
  /** OS/2 `fsType`, or the `/FSType` a CFF or Type 1 program states. */
  readonly fsType?: number;
  /** How far the program's own space advances, in thousandths of an em. */
  readonly spaceAdvance?: number;
}

/**
 * §9.9 — the {@link FaceProgram} of a font whose program this reads: the glyph
 * lookup, the licence, and what the descriptor states (§9.8.1).
 *
 * @param file     The owning file.
 * @param fontDict The font dictionary.
 * @param isType0  Whether it is a composite font, whose descendant owns the
 *                 descriptor.
 * @param glyph    The glyph each code draws, where the program can be read.
 * @param facts    What the program states of itself: its licence and its space.
 * @returns The `program` field of a {@link ContentFont}, or nothing.
 */
function programOf(
  file: PdfFile,
  fontDict: PdfDict,
  isType0: boolean,
  glyph: ((code: number) => ReadonlyArray<PathSeg> | undefined) | undefined,
  facts: ProgramFacts | undefined,
): { program?: FaceProgram } {
  if (!glyph) return {};
  const owner = isType0 ? descendantFont(file, fontDict) : fontDict;
  const descriptor = file.resolve(owner.get('FontDescriptor') ?? PDF_NULL);
  if (!(descriptor instanceof Map)) return {};
  const stated = (key: string): number | undefined => {
    const v = file.resolve(descriptor.get(key) ?? PDF_NULL);
    return typeof v === 'number' && Number.isFinite(v) ? v : undefined;
  };
  const flags = stated('Flags') ?? 0;
  const capHeight = stated('CapHeight');
  const xHeight = stated('XHeight');
  const base = asName(file.resolve(owner.get('BaseFont') ?? fontDict.get('BaseFont') ?? PDF_NULL));
  return {
    program: {
      glyph,
      ...(facts?.fsType !== undefined ? { fsType: facts.fsType } : {}),
      ...(facts?.spaceAdvance !== undefined ? { spaceAdvance: facts.spaceAdvance } : {}),
      postScriptName: base.replace(/^[A-Z]{6}\+/u, ''),
      ...(capHeight !== undefined && capHeight > 0 ? { capHeight } : {}),
      ...(xHeight !== undefined && xHeight > 0 ? { xHeight } : {}),
      italicAngle: stated('ItalicAngle') ?? 0,
      fixedPitch: (flags & FLAG_FIXED_PITCH) !== 0,
    },
  };
}

/**
 * The glyph index a name carries, for the names that carry one.
 *
 * A subsetter that drops a font's `cmap` renames its glyphs after their
 * INDEX — `g24`, `glyph24`, `index24`, `cid24` — and that name is then the only
 * way back to the outline. bug1151216.pdf names them `g24`, `g381`, `g3`, and
 * its three lines of prices are drawn from nothing else.
 *
 * @param name      The glyph's name.
 * @param numbering How the font writes the number, where its names say (see
 *                  {@link numberingOf}); four digits are hexadecimal otherwise.
 */
function numberedGlyph(name: string, numbering?: 'hex' | 'decimal'): number | undefined {
  const hex = numbering !== 'decimal' ? /^g([0-9a-f]{4})$/u.exec(name) : null;
  if (hex) return Number.parseInt(hex[1]!, 16);
  const m = /^(?:g|glyph|index|cid|G)(\d+)$/u.exec(name);
  const n = m ? Number(m[1]) : Number.NaN;
  return Number.isFinite(n) && n >= 0 && n < MAX_GLYPH_INDEX ? n : undefined;
}

/** No font holds more glyphs than this; a bigger number is not an index. */
const MAX_GLYPH_INDEX = 65536;

/**
 * How a font's index names write their numbers, where its names say.
 *
 * A subsetter that pads the number to four digits writes it in hexadecimal —
 * bug1027533.pdf's `g0024` is glyph 36 — and one that does not pad writes it
 * in decimal. One name cannot tell the two apart: bug1151216.pdf names its
 * glyphs `g24`, `g381` and `g1004`, and `g1004` is four digits either way.
 * Read as hexadecimal it is glyph 4100, which no subset of a thousand glyphs
 * holds: five of the file's codes drew nothing, fell back to Latin-1, and its
 * prices came back "$@'' for 1" set over the glyphs that were drawn. The rest
 * of the font's names say which it is — a leading zero or a letter is
 * hexadecimal, a number of any length but four is decimal.
 *
 * @param names The glyph names the font's `/Differences` gives.
 * @returns The numbering, or `undefined` where the names do not say.
 */
function numberingOf(names: Iterable<string>): 'hex' | 'decimal' | undefined {
  let hex = false;
  let decimal = false;
  for (const name of names) {
    const digits = /^g([0-9a-f]+)$/u.exec(name)?.[1];
    if (digits === undefined) continue;
    if (/[a-f]/u.test(digits) || (digits.length === 4 && digits.startsWith('0'))) hex = true;
    else if (digits.length !== 4) decimal = true;
  }
  return hex === decimal ? undefined : hex ? 'hex' : 'decimal';
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
function namedGlyphs(
  file: PdfFile,
  fontDict: PdfDict,
  statesUnicode: boolean,
  glyph: {
    draws: (name: string) => boolean;
    blank: (name: string) => boolean;
    character: (name: string) => string | undefined;
  },
): Map<number, string> | undefined {
  const names = differences(file, fontDict);
  if (names.size === 0) return undefined;
  const out = new Map<number, string>();
  for (const [code, name] of names) {
    const text = textForGlyphName(name) ?? glyph.character(name);
    if (text !== undefined) {
      out.set(code, text);
      continue;
    }
    // A name whose glyph draws NOTHING is a space, whatever it is called. A
    // subsetter names the space glyph after its index like every other, and
    // dropped as unreadable it took the gaps between the words with it —
    // TAMReview.pdf's figure labels came back "SystemFeatures".
    if (glyph.blank(name)) {
      out.set(code, ' ');
      continue;
    }
    // A name that is no character and a glyph that CAN be drawn: the text is
    // unrecoverable and the shape is not. Only where the program can draw it —
    // marked unreadable with nothing to put in its place, the words would
    // simply be gone.
    if (!statesUnicode && glyph.draws(name)) out.set(code, UNANSWERABLE);
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
  // eight thousand of its nine thousand words — and where the glyph CAN be
  // drawn the loop above has already taken them, which is the case that file
  // actually is.
  if (asName(file.resolve(fontDict.get('Subtype') ?? PDF_NULL)) !== 'Type3') return undefined;
  const unreadable = new Map<number, string>();
  for (const code of names.keys()) unreadable.set(code, '\uFFFD');
  return unreadable;
}

/**
 * §9.7.4.3 `/DW2` and `/W2` — how far the pen drops for one glyph, in 1000-unit
 * text space.
 *
 * `/DW2`'s default is `[880 -1000]`: the vertical origin sits 880 above the
 * horizontal one and the displacement is a full em DOWN. Only the displacement
 * is wanted here — the origin shifts where the glyph is drawn, which a reader
 * re-setting the words in another face does not reproduce anyway.
 */
function cidVerticalAdvance(file: PdfFile, fontDict: PdfDict): (code: number) => number {
  const cid = descendantFont(file, fontDict);
  const dw2 = file.resolve(cid.get('DW2') ?? PDF_NULL);
  const stated =
    Array.isArray(dw2) && typeof file.resolve(dw2[1] ?? PDF_NULL) === 'number'
      ? (file.resolve(dw2[1]!) as number)
      : DEFAULT_DW2_DISPLACEMENT;
  const perCode = new Map<number, number>();
  // §9.7.4.3 `/W2` — `c [w1y v1x v1y …]` or `cFirst cLast w1y v1x v1y`.
  const w2 = file.resolve(cid.get('W2') ?? PDF_NULL);
  if (Array.isArray(w2)) {
    for (let i = 0; i < w2.length; ) {
      const first = file.resolve(w2[i] ?? PDF_NULL);
      const next = file.resolve(w2[i + 1] ?? PDF_NULL);
      if (typeof first !== 'number') break;
      if (Array.isArray(next)) {
        for (let k = 0; k + 2 < next.length; k += 3) {
          const v = file.resolve(next[k] ?? PDF_NULL);
          if (typeof v === 'number') perCode.set(first + k / 3, v);
        }
        i += 2;
      } else if (
        typeof next === 'number' &&
        typeof file.resolve(w2[i + 2] ?? PDF_NULL) === 'number'
      ) {
        const v = file.resolve(w2[i + 2]!) as number;
        for (let c = first; c <= next && c - first < MAX_W2_RANGE; c++) perCode.set(c, v);
        i += 5;
      } else {
        break;
      }
    }
  }
  return (code: number): number => perCode.get(code) ?? stated;
}

/** §9.7.4.3 — `/DW2`'s default displacement: one em down the page. */
const DEFAULT_DW2_DISPLACEMENT = -1000;

/** A `/W2` range wider than this is not walked out code by code. */
const MAX_W2_RANGE = 65_536;

/**
 * Annex D.2 — code → text for the encoding a simple font is read through, under
 * whatever `/Differences` restates.
 *
 * `/Encoding` either names one of the base encodings, or is a dictionary that
 * may name one in `/BaseEncoding`, or is absent — and absent means the encoding
 * built into the FACE. Only the standard Latin faces can be answered for there:
 * a file that names them embeds no program, and what they are is known
 * (§9.6.2.2). An embedded font's built-in encoding lives in the program and a
 * substituted face has none worth guessing at, so both keep the Latin-1
 * reading, which is what the codes of such a file nearly always are.
 *
 * @returns Code → glyph NAME, or `undefined` where nothing better than Latin-1
 *          is known.
 */
function baseEncoding(file: PdfFile, fontDict: PdfDict): ReadonlyMap<number, string> | undefined {
  const encoding = file.resolve(fontDict.get('Encoding') ?? PDF_NULL);
  const named =
    encoding instanceof PdfName
      ? encoding.value
      : encoding instanceof Map
        ? asName(file.resolve(encoding.get('BaseEncoding') ?? PDF_NULL))
        : '';
  const table =
    named.length > 0
      ? baseEncodingTable(named)
      : isStandardLatinFace(asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL)))
        ? standardEncodingTable()
        : undefined;
  return table;
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

/**
 * §9.8.2 — the family a run states: the file's own name for the face, plus what
 * the file says the face IS.
 *
 * The name is a hint for the substitution and nothing more, unless the reader
 * can lift the program itself (§9.9 `/FontFile2`) — then it is the key the face
 * is filed under and must be left exactly as it is. Everything else is
 * substituted by name, and TeX and PostScript producers embed Type 1 and CFF
 * programs under names no table knows: `NimbusRomNo9L-Regu`, `LMRoman10`,
 * `CMR10`, `stonesans`. Every one of those pages came back set in a grotesque
 * — the loudest single difference between our page and a serif document's.
 *
 * The descriptor states the class outright, so where the name says nothing the
 * flags do: bit 1 is FixedPitch and bit 2 Serif.
 *
 * @param file     The document.
 * @param fontDict The font dictionary.
 * @param isType0  Whether it is a composite font, whose descendant owns the
 *                 descriptor.
 * @returns The family name for the run, or `undefined` where the font has none.
 */
function runFontName(file: PdfFile, fontDict: PdfDict, isType0: boolean): string | undefined {
  const name = embeddedFontName(file, fontDict);
  if (name === undefined || hasLiftableProgram(file, fontDict)) return name;
  // A name the substitution knows already says what it is — a sans as much as
  // a serif, and over flags that say otherwise: bug898853.pdf's descriptor
  // calls its FrutigerLTStd-Light Serif, and read off the flag "Canadian" came
  // back in a roman.
  const face = familyOfFace(asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL)));
  if (knowsFamily(name) || knowsFamily(face)) return name;
  const owner = isType0 ? descendantFont(file, fontDict) : fontDict;
  const descriptor = file.resolve(owner.get('FontDescriptor') ?? PDF_NULL);
  if (!(descriptor instanceof Map)) return name;
  const flags = asNumber(file.resolve(descriptor.get('Flags') ?? PDF_NULL), 0);
  if ((flags & FLAG_FIXED_PITCH) !== 0) return `${name} monospace`;
  if ((flags & FLAG_SERIF) !== 0) return `${name} serif`;
  return name;
}

/**
 * For every face a run may name, the family a word processor knows it by —
 * keyed by the name the run carries (see {@link FaceFamily}).
 *
 * A PDF names the FACE and a .docx names a family: written as the face,
 * `inter-semibold` was a font no reader has, and LibreOffice set the whole of
 * an invoice drawn in Inter in its default serif. The family is the one the
 * descriptor states (§9.8.1 `/FontFamily`) or, where it states none, the one
 * the PostScript name is made of.
 *
 * @param file  The document.
 * @param pages The pages whose faces are wanted.
 * @returns Run font name → its family.
 */
export function collectFaceFamilies(
  file: PdfFile,
  pages: ReadonlyArray<PdfPage>,
): Map<string, FaceFamily> {
  const out = new Map<string, FaceFamily>();
  eachPageFont(file, pages, (fontDict) => {
    const isType0 = asName(file.resolve(fontDict.get('Subtype') ?? PDF_NULL)) === 'Type0';
    const key = runFontName(file, fontDict, isType0);
    if (key === undefined || out.has(key)) return;
    const owner = isType0 ? descendantFont(file, fontDict) : fontDict;
    const descriptor = file.resolve(owner.get('FontDescriptor') ?? PDF_NULL);
    const stated =
      descriptor instanceof Map ? file.resolve(descriptor.get('FontFamily') ?? PDF_NULL) : PDF_NULL;
    const family =
      typeof stated === 'string' && /^[\x20-\x7e]+$/u.test(stated.trim())
        ? stated.trim()
        : familyOfFace(cidFontName(file, fontDict, isType0));
    if (family.length === 0) return;
    const flags =
      descriptor instanceof Map
        ? asNumber(file.resolve(descriptor.get('Flags') ?? PDF_NULL), 0)
        : 0;
    out.set(key, { family, generic: genericOf(family, flags) });
  });
  return out;
}

/**
 * §9.7.6.1 — a composite font's `/BaseFont` is its CIDFont's name, a hyphen and
 * the name of the CMap it is encoded by (`HeiseiMin-W3-UniJIS-UCS2-H`); the
 * face is the part before the CMap's.
 */
function cidFontName(file: PdfFile, fontDict: PdfDict, isType0: boolean): string {
  const base = asName(file.resolve(fontDict.get('BaseFont') ?? PDF_NULL));
  const cmap = isType0 ? asName(file.resolve(fontDict.get('Encoding') ?? PDF_NULL)) : '';
  return cmap !== '' && base.endsWith(`-${cmap}`) ? base.slice(0, -cmap.length - 1) : base;
}

/**
 * The family a PostScript face name is made of: `Inter-SemiBold` → `Inter`,
 * `ArialMT` → `Arial`, `TimesNewRomanPS-BoldMT` → `Times New Roman`.
 *
 * §9.6.2.1 names a face `Family-Style` (Word writes `Family,Style`), with the
 * family's words run together. What follows the separator is a style only if
 * it is made of style words — `MS-Mincho` is a family of its own — and the
 * words come apart where the capitals say they do.
 */
export function familyOfFace(baseFont: string): string {
  const name = plainFace(baseFont);
  const cut = /^(.+?)[-,]([^-,]+)$/u.exec(name);
  const whole =
    cut && STYLE_WORDS.test(cut[2]!)
      ? cut[1]!
      : cut
        ? name.replace(/-/gu, ' ')
        : (GLUED_STYLE.exec(name)?.[1] ?? name);
  const family = whole.replace(/(?:PSMT|PS|MT)$/u, '') || whole;
  return family
    .replace(/([a-z])([A-Z])/gu, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/gu, '$1 $2')
    .replace(/\bDeja Vu\b/u, 'DejaVu')
    .replace(/\s+/gu, ' ')
    .trim();
}

/**
 * A face name with what producers write around it taken off: the subset tag
 * (§9.6.4) and, after the face, the charset Acrobat's Office export appends
 * (`Arial,Bold-WinCharSetFFFF`, `-H2`), the encoding a CIDFont was cut for
 * (`-Identity-H`, `-OneByteIdentityH`), a serial number, and the slant or
 * stretch an interpreter synthesised (`-Slant_167`, `-Extend_850`). Left on,
 * bug900822.pdf's .docx asked for "Courier New,Bold Win Char Set FFFF", which
 * no machine has, over the Courier New every machine does.
 */
function plainFace(baseFont: string): string {
  let name = baseFont
    .replace(/^[A-Z]{6}\+/u, '')
    // ASCII punctuation only: `*Arial-68771-Identity-H`. A name in some
    // other encoding read as Latin-1 is letters all the same.
    .replace(/^[!-/:-@[-`{-~]+/u, '')
    .trim();
  for (let next = name.replace(PRODUCER_TAIL, ''); next !== name; ) {
    name = next;
    next = name.replace(PRODUCER_TAIL, '');
  }
  return name;
}

/** One piece of producer debris at the end of a face name (see `plainFace`). */
const PRODUCER_TAIL =
  /[-_](?:WinCharSet[0-9A-F]+|Identity-[HV]|OneByteIdentity[HV]|H\d+|\d+|(?:Slant|Extend)_\d+)$/iu;

/**
 * A style run on to a family with no separator — `CalibriBold` — which is the
 * family's name only with the style taken off. Whether the face IS bold is the
 * program's to say (see `faceStyle`): `NewBasrahBold` is a family of its own.
 */
const GLUED_STYLE =
  /^(.*?[a-z])((?:Semi|Demi|Extra|Ultra)?(?:Bold|Black|Heavy)(?:Italic|Oblique)?|Italic|Oblique)$/u;

/** A PostScript style part, made of nothing but style words (`SemiBoldItalic`, `BoldMT`, `Regu`). */
const STYLE_WORDS =
  /^(?:regular|regu|roman|book|normal|plain|medium|medi|light|thin|hairline|extra|ultra|semi|demi|bold|bd|black|heavy|italic|ital|it|oblique|obl|condensed|cond|narrow|compressed|extended|mt|ps)+$/iu;

/** §17.8.3.10 — the kind of face a family is, from the descriptor or, failing that, its name. */
function genericOf(family: string, flags: number): FaceFamily['generic'] {
  // The flags speak for a family the tables do not know, as in `runFontName`.
  if (!knowsFamily(family)) {
    if ((flags & FLAG_FIXED_PITCH) !== 0) return 'modern';
    if ((flags & FLAG_SERIF) !== 0) return 'roman';
  }
  const key = resolveFamilyStyle(family).key;
  if (key === 'cousine') return 'modern';
  return key === 'tinos' || key === 'caladea' ? 'roman' : 'swiss';
}

/** §9.8.2 `/Flags` — bit 1 is FixedPitch, bit 2 Serif (bits numbered from 1). */
const FLAG_FIXED_PITCH = 1 << 0;
const FLAG_SERIF = 1 << 1;

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
    // §9.9 — where the descriptor is silent the PROGRAM is a witness too: its
    // header states its own style (the `head` table's macStyle).
    // bigboundingbox.pdf names its faces `CalibriBold` and `Calibri`, states
    // no weight for either, and only the program says which is the bold cut.
    // It says so for the Bold of a family and not for its SemiBold, which is
    // a family of its own there — so the name `Inter-SemiBold` still counts,
    // and `NewBasrahBold`, a family whose program says Regular, still does not.
    //
    // Nor does the flag outweigh a name that states a weight of its own under
    // the bold one: bug898853.pdf sets "Canadian" in FrutigerLTStd-Light with
    // ForceBold set, and read off the flag the light word came back heavy.
    const own = programStyle(file, fontDict);
    const forced = (flags & FLAG_FORCE_BOLD) !== 0 && !named.light;
    const bold = stated
      ? asNumber(weightVal, 0) >= BOLD_WEIGHT
      : forced || own?.bold === true || named.bold;
    // The slant and the flag each state italic outright; where neither does,
    // the program and the name are the witnesses left.
    const italic =
      slant !== 0 || (flags & FLAG_ITALIC) !== 0 || own?.italic === true || named.italic;
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
function styleFromName(baseFont: string): { bold: boolean; italic: boolean; light: boolean } {
  // §9.6.4 — six arbitrary capitals and a plus sign mark a subset, and they may
  // spell anything at all.
  const name = plainFace(baseFont);
  const style = /[-,]([A-Za-z]+)$/u.exec(name)?.[1] ?? '';
  const bold = /bold|black|heavy|semib|demi/iu.test(style);
  return {
    bold,
    italic: /italic|oblique/iu.test(style),
    // A weight the name states at or under the medium one: `-Light`, `-Book`.
    light: !bold && /light|thin|hairline|book|regular|roman|normal|medium/iu.test(style),
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
