// §9.9 — the outlines each face drew a document's characters with, gathered so
// a writer can embed the face (see `src/pdf-reader/face-outlines.ts`). Each
// program format reaches its glyphs its own way, and each way is one a viewer
// takes: a glyph found another way is some other letter's.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { GlyphSeg } from '@/core/font';
import type { FaceOutlines } from '@/core/ir/flow';
import { buildTrueType, parseTtf } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { Ream } from '@/core/converter/ream';
import { outlineSource, sfntFsType, sfntSpaceAdvance } from '@/pdf-reader/glyf-outline';
import { writeDocx } from '@/word/docx-writer';

const ROBOTO = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));

/** The faces a PDF's reconstruction gathered, by the name its runs carry. */
function facesOf(pdf: Uint8Array): ReadonlyMap<string, FaceOutlines> {
  return Ream.parse(pdf).flow.faceOutlines ?? new Map();
}

/** Roboto's own outline for a character, as the reader gives any glyph. */
function robotoGlyph(char: string): ReadonlyArray<GlyphSeg> | undefined {
  const gid = parseTtf(ROBOTO).glyphForCodepoint(char.codePointAt(0)!);
  return outlineSource(ROBOTO)?.path(gid);
}

describe('the outlines a PDF face drew its characters with', () => {
  it('reaches a simple TrueType face through its Unicode cmap, by the names of its encoding', () => {
    // WinAnsiEncoding's 0xE9 is `eacute`, which is U+00E9 to the program's cmap.
    const faces = facesOf(
      simpleFontPdf({
        show: '(Hi \\351) Tj',
        font: '/Subtype /TrueType /Encoding /WinAnsiEncoding',
        flags: 32,
        program: ROBOTO,
      }),
    );
    const face = faces.get('roboto');
    expect(face).toBeDefined();
    for (const char of ['H', 'i', 'é'])
      expect(face!.glyphs.get(char)?.outline).toEqual(robotoGlyph(char));
    // The advance is the page's (§9.2.4), not the program's.
    expect(face!.glyphs.get('H')?.advance).toBe(610);
    expect(face!.glyphs.get(' ')?.outline).toEqual([]);
    expect(face!.fsType).toBe(sfntFsType(ROBOTO));
    expect(face!.postScriptName).toBe('Roboto');
  });

  it("reads a symbolic face's code through its Windows symbol cmap", () => {
    // §9.6.6.4 — a (3,0) subtable keys a one-byte code behind 0xF000.
    const faces = facesOf(
      simpleFontPdf({
        show: '(A) Tj',
        font: '/Subtype /TrueType',
        flags: 4,
        program: symbolFont(4),
      }),
    );
    const face = faces.get('roboto');
    // The program's box, closed by a line back to where it began.
    expect(units(face?.glyphs.get('A')?.outline ?? [])).toEqual([...units(BOX), [100, 700]]);
    // …and carries the licence the program states.
    expect(face?.fsType).toBe(4);
  });

  it('reaches a composite face by CID', () => {
    const parsed = parseTtf(ROBOTO);
    const [h, i] = ['H', 'i'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const faces = facesOf(identityPdf(`<${hex4(h!)}${hex4(i!)}> Tj`, [h!, i!], 'Hi'));
    const face = faces.get('roboto');
    expect(face?.glyphs.get('H')?.outline).toEqual(robotoGlyph('H'));
    expect(face?.glyphs.get('i')?.outline).toEqual(robotoGlyph('i'));
    expect(face?.glyphs.get('i')?.advance).toBe(600); // `/DW`
  });

  it('gives a space the page never showed the width the program gives its own', () => {
    const parsed = parseTtf(ROBOTO);
    const [h, i] = ['H', 'i'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const face = facesOf(identityPdf(`<${hex4(h!)}${hex4(i!)}> Tj`, [h!, i!], 'Hi')).get('roboto');
    expect(face?.glyphs.get(' ')?.advance).toBeCloseTo(sfntSpaceAdvance(ROBOTO)!, 6);
    expect(face?.glyphs.get('\u00a0')).toBe(face?.glyphs.get(' '));
  });

  it('gives it the white the page leaves between words, where the program keeps no space', () => {
    // A subset with neither `cmap` nor `post` has no way to its space; the page
    // moves the pen three tenths of an em between words, and that is the space.
    const parsed = parseTtf(ROBOTO);
    const [h, i] = ['H', 'i'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const words = `[<${hex4(h!)}> -300 <${hex4(i!)}> -300 <${hex4(h!)}> -300 <${hex4(i!)}>] TJ`;
    const face = facesOf(
      identityPdf(words, [h!, i!], 'Hi', withoutTables(ROBOTO, ['cmap', 'post'])),
    ).get('roboto');
    expect(face?.glyphs.get(' ')?.advance).toBeCloseTo(300, 6);
  });

  it('gives the space no width the file never states for it', () => {
    // pdfTeX writes /Widths from the first code a subset shows, and it shows no
    // space: code 32 had no width, which reads as the 500 a code with none
    // falls back to — a space twice as wide as the face's own, in every face
    // such a paper embeds.
    const face = facesOf(
      simpleFontPdf({
        show: '(Hi) Tj',
        font: '/Subtype /TrueType',
        flags: 32,
        program: ROBOTO,
        firstChar: 33,
      }),
    ).get('roboto');
    expect(face?.glyphs.get(' ')?.advance).toBeCloseTo(sfntSpaceAdvance(ROBOTO)!, 6);
  });

  it('takes no width for a space from a code its encoding names another glyph', () => {
    // Computer Modern's code 32 is the stroke of the Polish l.
    const face = facesOf(
      simpleFontPdf({
        show: '(Hi) Tj',
        font: '/Subtype /TrueType',
        flags: 32,
        program: ROBOTO,
        encoding: '/Encoding << /Differences [32 /suppress] >> ',
      }),
    ).get('roboto');
    expect(face?.glyphs.get(' ')?.advance).toBeCloseTo(sfntSpaceAdvance(ROBOTO)!, 6);
  });

  it('reads the pairs a page kerns a face by off the nudges of its TJ arrays', () => {
    // §9.4.3 — a number in a TJ array moves the pen back that many thousandths
    // of an em: 80 between H and i is a kern of -80.
    const parsed = parseTtf(ROBOTO);
    const [h, i] = ['H', 'i'].map((c) => hex4(parsed.glyphForCodepoint(c.codePointAt(0)!)));
    const cids = [h!, i!].map((x) => parseInt(x, 16));
    const flow = Ream.parse(identityPdf(`[<${h!}> 80 <${i!}> 60 <${h!}>] TJ`, cids, 'Hi')).flow;
    const face = flow.faceOutlines?.get('roboto');
    expect(face?.kerning?.get('Hi')).toBeCloseTo(-80, 6);
    expect(face?.kerning?.get('iH')).toBeCloseTo(-60, 6);
    // §17.3.2.19 — and a run set in a face the page kerned asks to be kerned.
    const runs = flow.body.flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs : []));
    expect(runs.length).toBeGreaterThan(0);
    for (const run of runs) expect(run.properties.kerningMinPt).toBe(1);
  });

  it('takes a nudge under a twip and a half for rounding, not a kern', () => {
    // Five thousandths of an em at 10pt is 0.05pt: Word writes the difference
    // between its twip grid and the face's widths that way, between any letters.
    const parsed = parseTtf(ROBOTO);
    const [h, i] = ['H', 'i'].map((c) => hex4(parsed.glyphForCodepoint(c.codePointAt(0)!)));
    const cids = [h!, i!].map((x) => parseInt(x, 16));
    const flow = Ream.parse(
      identityPdf(`[<${h!}> 5 <${i!}> 5 <${h!}>] TJ`, cids, 'Hi', ROBOTO, 10),
    ).flow;
    expect(flow.faceOutlines?.get('roboto')?.kerning).toBeUndefined();
    const runs = flow.body.flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs : []));
    for (const run of runs) expect(run.properties.kerningMinPt).toBeUndefined();
  });

  it('takes one nudge between every two letters and beside every space for tracking, not kerning', () => {
    // A page that sets a line tight moves the pen back the same amount after
    // every glyph, its spaces too — and no face is kerned against its space.
    const parsed = parseTtf(ROBOTO);
    const [h, i, sp] = ['H', 'i', ' '].map((c) =>
      hex4(parsed.glyphForCodepoint(c.codePointAt(0)!)),
    );
    const cids = [h!, i!, sp!].map((x) => parseInt(x, 16));
    const line = [h, i, sp, h, i, sp, h, i].map((g) => `<${g!}>`).join(' 50 ');
    const flow = Ream.parse(identityPdf(`[${line}] TJ`, cids, 'Hi ')).flow;
    expect(flow.faceOutlines?.get('roboto')?.kerning).toBeUndefined();
  });

  it("keeps the glyph a page drew for a run of letters as the face's ligature of them", () => {
    // §9.10.2 — Roboto's "fi" is one glyph, which /ToUnicode maps to two letters.
    const parsed = parseTtf(ROBOTO);
    const [fi, f, i] = ['\ufb01', 'f', 'i'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const show = `<${hex4(fi!)}${hex4(f!)}${hex4(i!)}> Tj`;
    const flow = Ream.parse(identityPdf(show, [fi!, f!, i!], ['fi', 'f', 'i'])).flow;
    const face = flow.faceOutlines?.get('roboto');
    expect(face?.ligatures?.get('fi')?.outline).toEqual(outlineSource(ROBOTO)?.path(fi!));
    expect(face?.ligatures?.get('fi')?.advance).toBe(600); // `/DW`
    expect(face?.glyphs.get('f')?.outline).toEqual(robotoGlyph('f'));
    // [MS-DOCX] `w14:ligatures` — and a run set in a face the page ligated asks
    // for the face's ligatures.
    const runs = flow.body.flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs : []));
    expect(runs.length).toBeGreaterThan(0);
    for (const run of runs) expect(run.properties.ligatures).toBe('standard');
  });

  it('asks for the layout of a current Word, the only one that forms them', () => {
    // [MS-DOCX] compatibilityMode — Word opens a document that states none in
    // Compatibility Mode, and forms no OpenType ligature there.
    const parsed = parseTtf(ROBOTO);
    const [fi, f, i] = ['\ufb01', 'f', 'i'].map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!));
    const show = `<${hex4(fi!)}${hex4(f!)}${hex4(i!)}> Tj`;
    const flow = Ream.parse(identityPdf(show, [fi!, f!, i!], ['fi', 'f', 'i'])).flow;
    expect(flow.compatibilityMode).toBe(15);
    const settings = OpcPackage.open(writeDocx(flow).bytes).getPart('word/settings.xml');
    expect(new TextDecoder().decode(settings)).toContain(
      '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>',
    );
  });

  it('takes no ligature of more than letters, nor of letters the face never draws alone', () => {
    // A glyph mapped to a name is a logo, not a ligature; and with no "l" of its
    // own, the face's "fl" is one a reader never lays out in it.
    const parsed = parseTtf(ROBOTO);
    const [fi, fl, f, i] = ['\ufb01', '\ufb02', 'f', 'i'].map((c) =>
      parsed.glyphForCodepoint(c.codePointAt(0)!),
    );
    const show = `<${hex4(fi!)}${hex4(fl!)}${hex4(f!)}${hex4(i!)}> Tj`;
    const flow = Ream.parse(identityPdf(show, [fi!, fl!, f!, i!], ['ACME', 'fl', 'f', 'i'])).flow;
    expect(flow.faceOutlines?.get('roboto')?.glyphs.get('f')).toBeDefined();
    expect(flow.faceOutlines?.get('roboto')?.ligatures).toBeUndefined();
    const runs = flow.body.flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs : []));
    for (const run of runs) expect(run.properties.ligatures).toBeUndefined();
  });

  it('draws the accented letter a CFF program composes of two glyphs (seac)', () => {
    // TN 5177 Appendix C — `eacute` is `adx ady bchar achar endchar`: the e,
    // and the acute moved by (adx, ady). Drawn as nothing, the é embedded as a
    // blank, and as its e alone it would be another letter.
    const faces = facesOf(
      simpleFontPdf({
        show: '(\\351) Tj',
        font: '/Subtype /Type1 /Encoding << /Differences [101 /e 233 /eacute] >>',
        flags: 32,
        program: accentedCff(),
        compact: true,
      }),
    );
    const corners = units(faces.get('accented')?.glyphs.get('é')?.outline ?? []);
    expect(corners).toContainEqual([100, 100]); // the e
    expect(corners).toContainEqual([600, 600]);
    expect(corners).toContainEqual([50, 600]); // the acute, 50 units along
    expect(corners).toContainEqual([150, 700]);
  });

  it('puts the glyphs in the slot its runs look the face up in', () => {
    // `Roboto,Bold` names a bold its regular program is not, and the page draws
    // the regular glyphs as they are: in the bold slot a reader shows them so,
    // where in the regular one it would thicken them into a bold never drawn.
    const face = facesOf(
      simpleFontPdf({
        show: '(Hi) Tj',
        font: '/Subtype /TrueType /Encoding /WinAnsiEncoding',
        flags: 32,
        program: ROBOTO,
        base: 'ABCDEF+Roboto,Bold',
      }),
    ).get('roboto,bold');
    expect(face?.bold).toBe(true);
    expect(face?.glyphs.get('H')?.outline).toEqual(robotoGlyph('H'));
  });

  it('gathers nothing from words the page does not paint', () => {
    // §9.3.6 — a scanned page's recognised words are shown in mode 3, in a
    // face with no ink in it; embedded, the words would be invisible.
    const faces = facesOf(
      simpleFontPdf({
        show: '3 Tr (Hi) Tj',
        font: '/Subtype /TrueType /Encoding /WinAnsiEncoding',
        flags: 32,
        program: ROBOTO,
      }),
    );
    expect(faces.size).toBe(0);
  });
});

/** The points an outline passes through, in thousandths of an em. */
function units(outline: ReadonlyArray<GlyphSeg>): Array<[number, number]> {
  return outline.flatMap(
    (seg): Array<[number, number]> =>
      seg.op === 'close' ? [] : [[Math.round(seg.x * 1000), Math.round(seg.y * 1000)]],
  );
}

/** A 400 × 700 box, clockwise, as a TrueType program draws it. */
const BOX: ReadonlyArray<GlyphSeg> = [
  { op: 'move', x: 0.1, y: 0.7 },
  { op: 'line', x: 0.5, y: 0.7 },
  { op: 'line', x: 0.5, y: 0 },
  { op: 'line', x: 0.1, y: 0 },
  { op: 'close' },
];

/**
 * A symbol font: one glyph at U+F041 behind a (3,0) subtable — the built font
 * with its Windows record's encoding turned from Unicode to symbol.
 */
function symbolFont(fsType: number): Uint8Array {
  const font = buildTrueType({
    family: 'Symbolic',
    bold: false,
    italic: false,
    postScriptName: 'Symbolic',
    glyphs: [{ codePoints: [0xf041], outline: BOX, advance: 600 }],
    ascent: 800,
    descent: -200,
    fsType,
  });
  const view = new DataView(font.buffer);
  for (let i = 0; i < view.getUint16(4); i++) {
    const at = 12 + i * 16;
    if (String.fromCharCode(...font.subarray(at, at + 4)) !== 'cmap') continue;
    const cmap = view.getUint32(at + 8);
    for (let r = 0; r < view.getUint16(cmap + 2); r++) {
      const record = cmap + 4 + r * 8;
      if (view.getUint16(record) === 3) view.setUint16(record + 2, 0);
    }
  }
  return font;
}

/**
 * A CFF program of four glyphs: `.notdef`, `e` (a 500-unit square), `acute` (a
 * 100-unit box over it), and `eacute`, which is the two by `seac`.
 */
function accentedCff(): Uint8Array {
  const n = (v: number): Array<number> =>
    v >= -107 && v <= 107 ? [v + 139] : [28, (v >> 8) & 0xff, v & 0xff];
  const box = (x: number, y: number, w: number, h: number): Array<number> => [
    ...n(x),
    ...n(y),
    21, // rmoveto
    ...n(w),
    ...n(0),
    5, // rlineto
    ...n(0),
    ...n(h),
    5,
    ...n(-w),
    ...n(0),
    5,
    14, // endchar
  ];
  const charStrings = index([
    Uint8Array.from([14]),
    Uint8Array.from(box(100, 100, 500, 500)),
    Uint8Array.from(box(0, 600, 100, 100)),
    Uint8Array.from([...n(50), ...n(0), ...n(101), ...n(194), 14]), // seac: e + acute
  ]);
  const name = index([new TextEncoder().encode('Accented')]);
  const empty = index([]);
  // The Top DICT: `charset` (15) and `CharStrings` (17), each a five-byte integer.
  const topSize = index([new Uint8Array(12)]).length;
  const charsetAt = 4 + name.length + topSize + empty.length * 2;
  // Format 0: one SID per glyph after `.notdef` — e, acute, eacute.
  const charset = Uint8Array.from([0, 0, 70, 0, 125, 0, 207]);
  const charStringsAt = charsetAt + charset.length;
  const int32 = (v: number): Array<number> => [
    29,
    (v >> 24) & 0xff,
    (v >> 16) & 0xff,
    (v >> 8) & 0xff,
    v & 0xff,
  ];
  const top = index([Uint8Array.from([...int32(charsetAt), 15, ...int32(charStringsAt), 17])]);
  return Uint8Array.from([
    1,
    0,
    4,
    1,
    ...name,
    ...top,
    ...empty,
    ...empty,
    ...charset,
    ...charStrings,
  ]);
}

/** TN 5176 §5 — an INDEX over the given items, with one-byte offsets. */
function index(items: ReadonlyArray<Uint8Array>): Uint8Array {
  if (items.length === 0) return Uint8Array.from([0, 0]);
  const offsets: Array<number> = [1];
  for (const item of items) offsets.push(offsets[offsets.length - 1]! + item.length);
  return Uint8Array.from([
    items.length >> 8,
    items.length & 0xff,
    1,
    ...offsets,
    ...items.flatMap((item) => [...item]),
  ]);
}

/** Four hex digits, as a two-byte code is written in a shown string. */
function hex4(code: number): string {
  return code.toString(16).padStart(4, '0');
}

/** A one-page PDF showing `show` in a simple font whose program is `program`. */
function simpleFontPdf(options: {
  show: string;
  font: string;
  flags: number;
  program: Uint8Array;
  compact?: boolean;
  base?: string;
  /** The first code the `/Widths` array states (32 unless said). */
  firstChar?: number;
  /** An `/Encoding` entry for the font dictionary, whole. */
  encoding?: string;
}): Uint8Array {
  const content = `BT /F0 40 Tf 20 40 Td ${options.show} ET`;
  const base = options.base ?? (options.compact === true ? 'ABCDEF+Accented' : 'ABCDEF+Roboto');
  const first = options.firstChar ?? 32;
  const widths = Array.from({ length: 256 - first }, (_, i) => (i + first === 72 ? 610 : 600)).join(
    ' ',
  );
  return assemble([
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 100] /Contents 4 0 R ' +
      '/Resources << /Font << /F0 5 0 R >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    `<< /Type /Font ${options.font} /BaseFont /${base} /FirstChar ${String(first)} /LastChar 255 ` +
      `/Widths [${widths}] ${options.encoding ?? ''}/FontDescriptor 6 0 R >>`,
    `<< /Type /FontDescriptor /FontName /${base} /Flags ${String(options.flags)} /ItalicAngle 0 ` +
      '/StemV 80 /Ascent 900 /Descent -200 /CapHeight 700 /FontBBox [-500 -300 1500 1000] ' +
      `/${options.compact === true ? 'FontFile3' : 'FontFile2'} 7 0 R >>`,
    fontStream(options.program, options.compact === true ? '/Subtype /Type1C ' : ''),
  ]);
}

/**
 * A one-page PDF showing `show` in Roboto through an `Identity-H` composite
 * font, whose `/ToUnicode` maps the given CIDs to `text`'s characters — or,
 * given a list, each CID to its string: a ligature's letters.
 */
function identityPdf(
  show: string,
  cids: ReadonlyArray<number>,
  text: string | ReadonlyArray<string>,
  program: Uint8Array = ROBOTO,
  size = 40,
): Uint8Array {
  const content = `BT /F0 ${String(size)} Tf 20 40 Td ${show} ET`;
  const strings = typeof text === 'string' ? [...text] : text;
  const utf16 = (s: string): string =>
    Array.from({ length: s.length }, (_, k) => hex4(s.charCodeAt(k))).join('');
  const pairs = cids.map((cid, i) => `<${hex4(cid)}> <${utf16(strings[i]!)}>`).join('\n');
  const toUnicode =
    '/CIDInit /ProcSet findresource begin 12 dict begin begincmap\n' +
    '1 begincodespacerange <0000> <FFFF> endcodespacerange\n' +
    `${String(cids.length)} beginbfchar\n${pairs}\nendbfchar\nendcmap end end`;
  return assemble([
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 100] /Contents 4 0 R ' +
      '/Resources << /Font << /F0 5 0 R >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type0 /BaseFont /ABCDEF+Roboto /Encoding /Identity-H ' +
      '/DescendantFonts [6 0 R] /ToUnicode 8 0 R >>',
    '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /ABCDEF+Roboto /DW 600 ' +
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> ' +
      '/FontDescriptor 7 0 R /CIDToGIDMap /Identity >>',
    '<< /Type /FontDescriptor /FontName /ABCDEF+Roboto /Flags 32 /ItalicAngle 0 /StemV 80 ' +
      '/Ascent 900 /Descent -200 /CapHeight 700 /FontBBox [-500 -300 1500 1000] /FontFile2 9 0 R >>',
    `<< /Length ${String(toUnicode.length)} >>\nstream\n${toUnicode}\nendstream`,
    fontStream(program, ''),
  ]);
}

/** A program with the named tables renamed out of reach, as a subsetter drops them. */
function withoutTables(program: Uint8Array, tags: ReadonlyArray<string>): Uint8Array {
  const out = Uint8Array.from(program);
  const count = (out[4]! << 8) | out[5]!;
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    if (tags.includes(String.fromCharCode(...out.subarray(at, at + 4)))) out[at] = 0x7a;
  }
  return out;
}

/** A font-program stream object, whose body is the program's own bytes. */
function fontStream(program: Uint8Array, subtype: string): Uint8Array {
  const head = new TextEncoder().encode(
    `<< ${subtype}/Length ${String(program.length)} >>\nstream\n`,
  );
  const tail = new TextEncoder().encode('\nendstream');
  return Uint8Array.from([...head, ...program, ...tail]);
}

/** Numbered objects, an xref built over their offsets, and a trailer. */
function assemble(objects: ReadonlyArray<string | Uint8Array>): Uint8Array {
  const encoder = new TextEncoder();
  const parts: Array<Uint8Array> = [encoder.encode('%PDF-1.7\n')];
  const offsets: Array<number> = [];
  let at = parts[0]!.length;
  objects.forEach((body, i) => {
    offsets.push(at);
    const open = encoder.encode(`${String(i + 1)} 0 obj\n`);
    const bytes = typeof body === 'string' ? encoder.encode(body) : body;
    const close = encoder.encode('\nendobj\n');
    parts.push(open, bytes, close);
    at += open.length + bytes.length + close.length;
  });
  let tail = `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) tail += `${String(off).padStart(10, '0')} 00000 n \n`;
  tail += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(at)}\n%%EOF\n`;
  parts.push(encoder.encode(tail));
  return Uint8Array.from(parts.flatMap((p) => [...p]));
}
