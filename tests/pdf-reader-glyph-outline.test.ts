// §9.6.6 — a composite font read by glyph index says nothing about its
// characters: `/Encoding /Identity-H` makes every code a CID, a subset program
// has had its `cmap` stripped, and `/ToUnicode` is absent or unusable. There is
// no text to recover and there is a shape, so the glyph is DRAWN.
//
// Eleven files of the pdf.js corpus are exactly this and every one of them came
// back a blank sheet — complex_ttf_font.pdf is eight lines of Arabic.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import type { PdfDict } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { buildContentFont } from '@/pdf-reader/font';
import { cffOutlineSource } from '@/pdf-reader/cff-outline';
import { type1Font } from '@/pdf-reader/type1-outline';
import { outlineSource, postGlyphNames } from '@/pdf-reader/glyf-outline';
import { macGlyphName } from '@/pdf-reader/encodings';
import { parseTtf } from '@/core/font';

const ROBOTO = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));

/**
 * The same program with its `cmap` renamed out of reach — which is what a
 * subsetter leaves behind, and the whole reason a code can reach no character.
 */
function withoutCmap(): Uint8Array {
  const out = Uint8Array.from(ROBOTO);
  const count = (out[4]! << 8) | out[5]!;
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    if (String.fromCharCode(out[at]!, out[at + 1]!, out[at + 2]!, out[at + 3]!) !== 'cmap')
      continue;
    out[at] = 'z'.charCodeAt(0);
    break;
  }
  return out;
}

/** The glyph index Roboto uses for one character. */
function glyphFor(ch: string): number {
  return parseTtf(ROBOTO).glyphForCodepoint(ch.codePointAt(0)!);
}

describe('glyph outlines (§9.6.6)', () => {
  it('reads a glyph as contours in a one-unit em', () => {
    const source = outlineSource(ROBOTO);
    expect(source).toBeDefined();
    const path = source?.path(glyphFor('A'));
    expect(path).toBeDefined();
    if (!path) return;
    // A capital A is one closed outline plus the counter inside it.
    expect(path.filter((s) => s.op === 'move').length).toBeGreaterThanOrEqual(2);
    expect(path[0]?.op).toBe('move');
    expect(path[path.length - 1]?.op).toBe('close');
    const xs = path.flatMap((s) => (s.op === 'close' ? [] : [s.x]));
    const ys = path.flatMap((s) => (s.op === 'close' ? [] : [s.y]));
    // In a one-unit em an upper-case letter stands about 0.7 high and rests on
    // the baseline; anything outside 0..1.2 means the scaling went wrong.
    expect(Math.max(...ys)).toBeGreaterThan(0.55);
    expect(Math.max(...ys)).toBeLessThan(1.2);
    expect(Math.min(...ys)).toBeGreaterThan(-0.1);
    expect(Math.max(...xs)).toBeLessThan(1.2);
  });

  it('gives nothing for a glyph the program does not hold', () => {
    expect(outlineSource(ROBOTO)?.path(999999)).toBeUndefined();
  });

  it('is not offered for a program with no outlines this reads', () => {
    // A CFF program under `/FontFile2` parses as an sfnt and carries no
    // `glyf`/`loca` pair; its charstrings are a separate reading.
    const header = Uint8Array.from([0x4f, 0x54, 0x54, 0x4f, 0, 0, 0, 0, 0, 0, 0, 0]);
    expect(outlineSource(header)).toBeUndefined();
  });

  it('draws the glyphs of a font that cannot say what its characters are', () => {
    const pdf = identityPdf(`<${hex4(glyphFor('A'))}${hex4(glyphFor('B'))}> Tj`);
    const doc = Ream.parse(pdf);
    const shapes = doc.flow.body.filter((b) => b.kind === 'shape');
    expect(shapes).toHaveLength(2);
    // Set at 40pt, a capital stands around 28pt tall.
    for (const shape of shapes) {
      expect(shape.shape.height).toBeGreaterThan(20);
      expect(shape.shape.height).toBeLessThan(40);
    }
    // And it is drawn BECAUSE the text could not be read, which the
    // reconstruction still says out loud.
    expect(doc.losses.some((l) => /map to no character/u.test(l.detail))).toBe(true);
  });

  it('runs a CFF charstring for a face whose outlines are a program', () => {
    // Adobe TN 5177 — a CFF glyph is not a table of points but a program on a
    // stack machine, so reading one means running it.
    // cff_bluescale_small_zones.pdf sets "Public Mobile" in one and came back a
    // blank sheet.
    const path = cffOutlineSource(squareCff())?.path(1);
    expect(path).toBeDefined();
    if (!path) return;
    expect(path[0]?.op).toBe('move');
    const xs = path.flatMap((seg) => (seg.op === 'close' ? [] : [seg.x]));
    const ys = path.flatMap((seg) => (seg.op === 'close' ? [] : [seg.y]));
    // The charstring draws a 500-unit square whose corner is at (100, 100),
    // and a one-unit em divides all four by a thousand.
    expect(Math.min(...xs)).toBeCloseTo(0.1, 6);
    expect(Math.max(...xs)).toBeCloseTo(0.6, 6);
    expect(Math.min(...ys)).toBeCloseTo(0.1, 6);
    expect(Math.max(...ys)).toBeCloseTo(0.6, 6);
  });

  it('draws the glyphs of a CFF face embedded in a PDF', () => {
    const pdf = identityPdf('<0001> Tj', undefined, squareCff(), 'FontFile3');
    const doc = Ream.parse(pdf);
    const shapes = doc.flow.body.filter((b) => b.kind === 'shape');
    expect(shapes).toHaveLength(1);
    // Half an em square, set at 40pt.
    expect(shapes[0]?.shape.width).toBeCloseTo(20, 1);
    expect(shapes[0]?.shape.height).toBeCloseTo(20, 1);
  });

  it('runs a Type 1 charstring out from under its two layers of encryption', () => {
    // Adobe "Type 1 Font Format" §6–7 — `eexec` hides the private dictionary
    // and every charstring inside it is encrypted again on its own, so reading
    // one means undoing both and then running it.
    const face = type1Font(squareType1());
    expect(face).toBeDefined();
    const path = face?.path('square');
    expect(path).toBeDefined();
    if (!path) return;
    const xs = path.flatMap((seg) => (seg.op === 'close' ? [] : [seg.x]));
    const ys = path.flatMap((seg) => (seg.op === 'close' ? [] : [seg.y]));
    // `hsbw` puts the origin at the sidebearing, so the square drawn 100 to the
    // right of a 50-unit sidebearing stands at 150, and its far side at 650.
    expect(Math.min(...xs)).toBeCloseTo(0.15, 6);
    expect(Math.max(...xs)).toBeCloseTo(0.65, 6);
    expect(Math.min(...ys)).toBeCloseTo(0.1, 6);
    expect(Math.max(...ys)).toBeCloseTo(0.6, 6);
    // And the program's own `/Encoding` is read, which is all a face with no
    // `/Differences` says about its codes.
    expect(face?.encoding?.get(65)).toBe('square');
  });

  it('takes the glyph index out of a name that carries one', () => {
    // A subsetter that drops a font's `cmap` renames its glyphs after their
    // INDEX. bug1151216.pdf writes them `g24` and bug1027533.pdf `g0024`, which
    // is the same number in hexadecimal — read as decimal its eight letters
    // came back as eight different ones.
    const gid = glyphFor('A');
    for (const name of [`g${String(gid)}`, `g${gid.toString(16).padStart(4, '0')}`]) {
      const doc = Ream.parse(simpleTruetypePdf(name));
      const shapes = doc.flow.body.filter((b) => b.kind === 'shape');
      expect(shapes).toHaveLength(1);
      // A capital set at 40pt stands around 28pt tall.
      expect(shapes[0]?.shape.height).toBeGreaterThan(20);
      expect(shapes[0]?.shape.height).toBeLessThan(40);
    }
  });

  it('reaches a legacy eight-bit face by the NAME its program gives the glyph', () => {
    // §9.6.6.4 — a TrueType with no `cmap` cannot be reached by character at
    // all. What is left is the encoding's glyph NAME and the program's `post`
    // table: TrueType_without_cmap.pdf is Masis, an Armenian face whose `i`
    // draws ի. Read as text its line came back "'>in"; drawn by name it is the
    // four letters the page shows.
    const file = PdfFile.parse(legacyTruetypePdf());
    const fonts = file.get(file.pages()[0]!.resources!, 'Font');
    if (!(fonts instanceof Map)) throw new Error('the page has a font');
    const font = buildContentFont(file, file.resolve(fonts.get('F0')!) as PdfDict);
    // The name says `A` and the shape is whatever the foundry put there, so
    // there is no text to recover…
    expect(font.decode([65])).toBe('�');
    // …and the shape is the one the program holds under that name.
    expect(font.outline?.path(65)).toEqual(outlineSource(withoutCmap())?.path(glyphFor('A')));
  });

  it('knows the standard order a post index below 258 stands for', () => {
    expect(macGlyphName(0)).toBe('.notdef');
    expect(macGlyphName(36)).toBe('A');
    expect(macGlyphName(76)).toBe('i');
    expect(macGlyphName(257)).toBe('dcroat');
    expect(macGlyphName(258)).toBeUndefined();
    // …and the table itself, read out of a program that states names.
    expect(postGlyphNames(ROBOTO)?.get('A')).toBe(glyphFor('A'));
  });

  it('reads a blank glyph as the space it is', () => {
    // A subsetter names the space glyph after its index like every other, so
    // its name says nothing and the glyph draws nothing — and dropped as
    // unreadable it took the gaps between the words with it. TAMReview.pdf's
    // figure labels came back "SystemFeatures".
    const file = PdfFile.parse(simpleTruetypePdf(`g${String(glyphFor(' '))}`));
    const fonts = file.get(file.pages()[0]!.resources!, 'Font');
    expect(fonts).toBeInstanceOf(Map);
    if (!(fonts instanceof Map)) return;
    const font = buildContentFont(file, file.resolve(fonts.get('F0')!) as PdfDict);
    expect(font.decode([65])).toBe(' ');
    // A space is not artwork: nothing is drawn for it.
    expect(font.outline?.path(65)).toBeUndefined();
  });

  it('keeps two bytes to a code under Identity-H whatever a /ToUnicode says', () => {
    // §9.7.6.2 — the CMap the font NAMES says how wide a code is; a
    // `/ToUnicode` is a second mapping and has no vote.
    // issue11549_reduced.pdf ships one truncated mid-stream, and read as one
    // byte every code split in half: each glyph came out preceded by the
    // `.notdef` box that the leading zero drew.
    const toUnicode =
      '/CIDInit /ProcSet findresource begin 12 dict begin begincmap\n' +
      '1 begincodespacerange\n<00><ff>\nendcodespacerange\nendcmap end end';
    const pdf = identityPdf(`<${hex4(glyphFor('A'))}${hex4(glyphFor('B'))}> Tj`, toUnicode);
    const doc = Ream.parse(pdf);
    // Two codes, so two glyphs — not four, and no `.notdef` among them.
    expect(doc.flow.body.filter((b) => b.kind === 'shape')).toHaveLength(2);
  });
});

/**
 * A one-page PDF setting code 65 in a SIMPLE TrueType font whose `/Differences`
 * name that code `name` — the shape of a subset whose `cmap` was dropped.
 */
function simpleTruetypePdf(name: string): Uint8Array {
  const content = 'BT /F0 40 Tf 20 40 Td (A) Tj ET';
  return assemble([
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R ' +
      '/Resources << /Font << /F0 5 0 R >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /TrueType /BaseFont /Roboto /FirstChar 65 /LastChar 65 ' +
      `/Widths [600] /FontDescriptor 6 0 R /Encoding << /Type /Encoding /Differences [65 /${name}] >> >>`,
    '<< /Type /FontDescriptor /FontName /Roboto /Flags 4 /ItalicAngle 0 /StemV 80 ' +
      '/Ascent 900 /Descent -200 /CapHeight 700 /FontBBox [-500 -300 1500 1000] /FontFile2 7 0 R >>',
    fontStreamObject(withoutCmap()),
  ]);
}

/**
 * A one-page PDF setting code 65 in a SIMPLE TrueType whose program carries no
 * `cmap` and which names nothing of its own — the shape of a legacy eight-bit
 * face, reachable only through its encoding's glyph names.
 */
function legacyTruetypePdf(): Uint8Array {
  const content = 'BT /F0 40 Tf 20 40 Td (A) Tj ET';
  return assemble([
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R ' +
      '/Resources << /Font << /F0 5 0 R >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /TrueType /BaseFont /Masis /FirstChar 65 /LastChar 65 ' +
      '/Widths [600] /FontDescriptor 6 0 R /Encoding /WinAnsiEncoding >>',
    '<< /Type /FontDescriptor /FontName /Masis /Flags 32 /ItalicAngle 0 /StemV 80 ' +
      '/Ascent 900 /Descent -200 /CapHeight 700 /FontBBox [-500 -300 1500 1000] /FontFile2 7 0 R >>',
    fontStreamObject(withoutCmap()),
  ]);
}

/**
 * A Type 1 program of one glyph, `square`, drawing a 500-unit box — the
 * smallest thing the charstring interpreter can be asked to run.
 */
function squareType1(): Uint8Array {
  const encrypt = (data: Uint8Array, key: number, lead: number): Uint8Array => {
    let r = key;
    const out = new Uint8Array(data.length + lead);
    const plain = new Uint8Array(out.length);
    plain.set(data, lead);
    for (let i = 0; i < plain.length; i++) {
      const p = plain[i]!;
      const c = p ^ (r >> 8);
      out[i] = c & 0xff;
      r = ((out[i]! + r) * 52845 + 22719) & 0xffff;
    }
    return out;
  };
  const num = (v: number): Array<number> =>
    v >= -107 && v <= 107
      ? [v + 139]
      : [255, (v >> 24) & 0xff, (v >> 16) & 0xff, (v >> 8) & 0xff, v & 0xff];
  const charstring = Uint8Array.from([
    ...num(50),
    ...num(600),
    13, // 50 600 hsbw — sidebearing, then the width
    ...num(100),
    ...num(100),
    21, // 100 100 rmoveto
    ...num(500),
    6, // 500 hlineto
    ...num(500),
    7, // 500 vlineto
    ...num(-500),
    6, // -500 hlineto
    9, // closepath
    14, // endchar
  ]);
  const glyph = encrypt(charstring, 4330, 4);
  const encoder = new TextEncoder();
  const priv = [
    ...encoder.encode('XXXXdup /Private 8 dict dup begin\n/lenIV 4 def\n/Subrs 0 array ND\n'),
    ...encoder.encode(`/CharStrings 1 dict dup begin\n/square ${String(glyph.length)} RD `),
    ...glyph,
    ...encoder.encode(' ND\nend end\n'),
  ];
  // The first four bytes of the eexec plaintext are discarded on the way back.
  const body = encrypt(Uint8Array.from(priv.slice(4)), 55665, 4);
  const head = encoder.encode(
    '%!PS-AdobeFont-1.0: Square\n/FontMatrix [0.001 0 0 0.001 0 0] readonly def\n' +
      '/Encoding 256 array\ndup 65 /square put\nreadonly def\ncurrentdict end\ncurrentfile eexec\n',
  );
  const out = new Uint8Array(head.length + body.length);
  out.set(head, 0);
  out.set(body, head.length);
  return out;
}

/** Four hex digits, as a two-byte code is written in a shown string. */
function hex4(code: number): string {
  return code.toString(16).padStart(4, '0');
}

/**
 * A one-page PDF setting `content` in Roboto through an `Identity-H` composite
 * font — the shape of every file this feature exists for.
 */
function identityPdf(
  show: string,
  toUnicode?: string,
  program?: Uint8Array,
  key: 'FontFile2' | 'FontFile3' = 'FontFile2',
): Uint8Array {
  const content = `BT /F0 40 Tf 20 40 Td ${show} ET`;
  const objects: Array<string | Uint8Array> = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R ' +
      '/Resources << /Font << /F0 5 0 R >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type0 /BaseFont /Roboto /Encoding /Identity-H ' +
      `/DescendantFonts [6 0 R]${toUnicode !== undefined ? ' /ToUnicode 8 0 R' : ''} >>`,
    '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Roboto /DW 600 ' +
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> ' +
      '/FontDescriptor 7 0 R /CIDToGIDMap /Identity >>',
    '<< /Type /FontDescriptor /FontName /Roboto /Flags 4 /ItalicAngle 0 /StemV 80 ' +
      `/Ascent 900 /Descent -200 /CapHeight 700 /FontBBox [-500 -300 1500 1000] /${key} 9 0 R >>`,
    toUnicode !== undefined
      ? `<< /Length ${String(toUnicode.length)} >>\nstream\n${toUnicode}\nendstream`
      : '<< >>',
    fontStreamObject(program ?? withoutCmap()),
  ];
  return assemble(objects);
}

/** The font-program object, whose body is the face's own bytes. */
function fontStreamObject(program: Uint8Array): Uint8Array {
  const head = new TextEncoder().encode(
    `<< /Length ${String(program.length)} /Length1 ${String(program.length)} >>\nstream\n`,
  );
  const tail = new TextEncoder().encode('\nendstream');
  const out = new Uint8Array(head.length + program.length + tail.length);
  out.set(head, 0);
  out.set(program, head.length);
  out.set(tail, head.length + program.length);
  return out;
}

/**
 * A CFF program of two glyphs: `.notdef`, and one charstring drawing a
 * 500-unit square at (100, 100) — the smallest thing the interpreter can run.
 */
function squareCff(): Uint8Array {
  const int16 = (v: number): Array<number> => [28, (v >> 8) & 0xff, v & 0xff];
  const square = Uint8Array.from([
    ...[100 + 139, 100 + 139, 21], // 100 100 rmoveto
    ...int16(500),
    139,
    5, // 500 0 rlineto
    139,
    ...int16(500),
    5, // 0 500 rlineto
    ...int16(-500),
    139,
    5, // -500 0 rlineto
    14, // endchar
  ]);
  const charStrings = index([Uint8Array.from([14]), square]);
  const name = index([new TextEncoder().encode('Square')]);
  const strings = index([]);
  const gsubrs = index([]);
  // The Top DICT holds one entry — `CharStrings` (operator 17) — written as a
  // five-byte integer so its own size does not depend on the value.
  const header = 4;
  const dictSize = 6;
  const top = index([new Uint8Array(dictSize)]);
  const charStringsAt = header + name.length + top.length + strings.length + gsubrs.length;
  const dict = Uint8Array.from([
    29,
    (charStringsAt >> 24) & 0xff,
    (charStringsAt >> 16) & 0xff,
    (charStringsAt >> 8) & 0xff,
    charStringsAt & 0xff,
    17,
  ]);
  const out = new Uint8Array(charStringsAt + charStrings.length);
  out.set([1, 0, header, 1], 0);
  out.set(name, header);
  out.set(index([dict]), header + name.length);
  out.set(strings, header + name.length + top.length);
  out.set(gsubrs, header + name.length + top.length + strings.length);
  out.set(charStrings, charStringsAt);
  return out;
}

/** TN 5176 §5 — an INDEX over the given items, with one-byte offsets. */
function index(items: ReadonlyArray<Uint8Array>): Uint8Array {
  if (items.length === 0) return Uint8Array.from([0, 0]);
  const offsets: Array<number> = [1];
  for (const item of items) offsets.push((offsets[offsets.length - 1] ?? 1) + item.length);
  const head = [items.length >> 8, items.length & 0xff, 1, ...offsets];
  const body = items.reduce<Array<number>>((acc, item) => [...acc, ...item], []);
  return Uint8Array.from([...head, ...body]);
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
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}
