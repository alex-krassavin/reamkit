// E-PDF EP2 — /ToUnicode CMap parsing and end-to-end page text extraction. The
// honest test reads the text back out of a PDF Ream itself wrote.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { TextRun } from '@/pdf-reader/content';
import type { PdfVector } from '@/pdf-reader/vector';
import { Ream } from '@/core/converter/ream';
import { parseToUnicodeCMap } from '@/pdf-reader/cmap';
import { PdfFile } from '@/pdf-reader/document';
import { collectEmbeddedFonts } from '@/pdf-reader/embedded-fonts';
import { textForGlyphName } from '@/pdf-reader/glyph-names';
import { patternTint } from '@/pdf-reader/pattern-tint';
import { extractPageText } from '@/pdf-reader/text';
import { collectPageAppearances } from '@/pdf-reader/annots';
import { withMeasuredMargins } from '@/pdf-reader/flow-build';
import { displayOf, placeRuns, placeVectors } from '@/pdf-reader/display';
import { markDrawnRules } from '@/pdf-reader/text-rules';
import { collectPageVectors } from '@/pdf-reader/vector';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { parseTtf } from '@/core/font/ttf-parser';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

describe('/ToUnicode CMap parser (E-PDF EP2)', () => {
  it('parses bfchar and bfrange blocks and the code width', () => {
    const cmap = [
      'begincmap',
      '1 begincodespacerange <0000> <FFFF> endcodespacerange',
      '2 beginbfchar <0003> <0020> <0011> <0048> endbfchar',
      '1 beginbfrange <0041> <0043> <0061> endbfrange',
      'endcmap',
    ].join('\n');
    const { map, codeBytes } = parseToUnicodeCMap(new TextEncoder().encode(cmap));
    expect(codeBytes).toBe(2);
    expect(map.get(0x0003)).toBe(' '); // bfchar → U+0020
    expect(map.get(0x0011)).toBe('H'); // bfchar → U+0048
    expect(map.get(0x0041)).toBe('a'); // bfrange base U+0061
    expect(map.get(0x0042)).toBe('b'); // +1
    expect(map.get(0x0043)).toBe('c'); // +2
  });

  it('parses an array-form bfrange', () => {
    const cmap = '1 beginbfrange <0001> <0002> [<0058> <0059>] endbfrange';
    const { map } = parseToUnicodeCMap(new TextEncoder().encode(cmap));
    expect(map.get(0x0001)).toBe('X');
    expect(map.get(0x0002)).toBe('Y');
  });
});

/**
 * A one-page PDF holding one text-showing operator in a SIMPLE TrueType font
 * whose `/ToUnicode` declares the two-byte codespace Distiller writes for every
 * font, simple or not. Hand-built because Ream's own writer never emits that
 * combination — and it is exactly the combination that broke.
 */
function simpleFontPdf(text: string): Uint8Array {
  const cmap = [
    '/CIDInit /ProcSet findresource begin 12 dict begin begincmap',
    '1 begincodespacerange <0000> <FFFF> endcodespacerange',
    // One entry per distinct byte of `text`, mapped to itself.
    `${String(new Set(text).size)} beginbfchar`,
    ...[...new Set(text)].map((c) => {
      const hex = c.charCodeAt(0).toString(16).padStart(2, '0').toUpperCase();
      return `<${hex}> <00${hex}>`;
    }),
    'endbfchar',
    'endcmap end end',
  ].join('\n');
  const content = `BT /F1 12 Tf 72 720 Td (${text}) Tj ET`;
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] ' +
      '/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /TrueType /BaseFont /Arial /FirstChar 32 /LastChar 255 ' +
      '/Encoding /WinAnsiEncoding /ToUnicode 6 0 R >>',
    `<< /Length ${String(cmap.length)} >>\nstream\n${cmap}\nendstream`,
  ];

  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\n`;
  pdf += `startxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

/**
 * One line of text in a font named `baseFont`. `descriptor` is the
 * `/FontDescriptor` body, or `undefined` for a font that carries none at all —
 * the standard-14 case (§9.6.2.2), where the name is the only witness.
 */
function styledFontPdf(baseFont: string, descriptor?: string): Uint8Array {
  const content = 'BT /F1 12 Tf 72 720 Td (Styled) Tj ET';
  const font =
    `<< /Type /Font /Subtype /TrueType /BaseFont /${baseFont} ` +
    `/FirstChar 32 /LastChar 255 /Encoding /WinAnsiEncoding` +
    `${descriptor !== undefined ? ' /FontDescriptor 5 0 R' : ''} >>`;
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] ' +
      '/Resources << /Font << /F1 4 0 R >> >> /Contents 6 0 R >>',
    font,
    descriptor !== undefined
      ? `<< /Type /FontDescriptor /FontName /${baseFont} ${descriptor} >>`
      : '<< >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\n`;
  pdf += `startxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

const styleOf = (baseFont: string, descriptor?: string) => {
  const file = PdfFile.parse(styledFontPdf(baseFont, descriptor));
  const run = extractPageText(file, file.pages()[0]!)[0];
  return { bold: run?.bold, italic: run?.italic };
};

/**
 * A simple Type1 font that states NO `/ToUnicode` and says what its codes are
 * only through `/Encoding /Differences` — the shape every PDF from TeX takes.
 * The codes start at 5 because a subset font's codes start wherever the subset
 * does, not at any character's value.
 */
function namedGlyphPdf(): Uint8Array {
  const content = 'BT /F1 12 Tf 72 720 Td <050607080905> Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] ' +
      '/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /KUEKMM+OpenSans ' +
      '/FirstChar 5 /LastChar 10 /Encoding 6 0 R >>',
    '<< /Type /Encoding /Differences [5 /L /a /T /e /X /five.os] >>',
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\n`;
  pdf += `startxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

describe('a font that names its glyphs rather than mapping them (§9.6.6.1)', () => {
  it('reads the names, where reading the codes gives rubble', () => {
    // issue10640.pdf is a LaTeX document: a subset font whose codes start at 5,
    // and not a /ToUnicode in the file. Read as Latin-1 — all that is left
    // without the names — its title came back as "!48 SUPPORT" where it reads
    // "LaTeX support", and its author as "-OHAMED %LORABITY".
    const file = PdfFile.parse(namedGlyphPdf());
    const run = extractPageText(file, file.pages()[0]!)[0];
    // Codes 5,6,7,8,9 are L,a,T,e,X; code 5 again closes it.
    expect(run?.text).toBe('LaTeXL');
  });

  it('reads a variant suffix as the shape it is, not a character', () => {
    // `five.os` is the oldstyle five — a shape, and still a "5".
    expect(textForGlyphName('five.os')).toBe('5');
    expect(textForGlyphName('a.sc')).toBe('a');
  });

  it('composes an accented name, and joins a ligature name', () => {
    expect(textForGlyphName('eacute')).toBe('é');
    expect(textForGlyphName('ccedilla')).toBe('ç');
    expect(textForGlyphName('f_i')).toBe('fi');
    expect(textForGlyphName('ff')).toBe('ﬀ');
  });

  it('reads the algorithmic names, and says nothing for a slot number', () => {
    expect(textForGlyphName('uni0041')).toBe('A');
    expect(textForGlyphName('uni00410042')).toBe('AB');
    expect(textForGlyphName('u1F600')).toBe('😀');
    // `g42` names a slot in the program, not a character.
    expect(textForGlyphName('g42')).toBeUndefined();
  });
});

describe('the face a run was shown in (§9.8.1)', () => {
  it('reads the weight and the slant off the descriptor', () => {
    // 160F-2019.pdf sets its title in Arial-BoldMT at /FontWeight 700, and
    // rebuilt without it the whole heading came back light.
    expect(styleOf('Arial-BoldMT', '/Flags 32 /FontWeight 700 /ItalicAngle 0')).toEqual({
      bold: true,
      italic: undefined,
    });
    expect(styleOf('ArialMT', '/Flags 32 /FontWeight 400 /ItalicAngle 0')).toEqual({
      bold: undefined,
      italic: undefined,
    });
    expect(styleOf('ArialMT', '/Flags 96 /ItalicAngle -12')).toEqual({
      bold: undefined,
      italic: true,
    });
  });

  it('takes the ForceBold flag as a weight of its own', () => {
    expect(styleOf('Whatever', '/Flags 262176').bold).toBe(true);
  });

  it('believes a descriptor that states no weight, whatever the name says', () => {
    // "New Basrah Bold" and "Damascus Bold" are family names, not weights, and
    // their descriptors give neither /FontWeight nor ForceBold.
    // ArabicCIDTrueType.pdf sets two of its four lines in them, and reading the
    // name over the descriptor set two lines heavy that no reader sets heavy.
    expect(styleOf('NewBasrahBold', '/Flags 4 /ItalicAngle 0').bold).toBeUndefined();
  });

  it('falls back to the name only where there is no descriptor at all', () => {
    // §9.6.2.2 — a standard-14 font carries none, so the name is all there is.
    expect(styleOf('Helvetica-BoldOblique')).toEqual({ bold: true, italic: true });
    expect(styleOf('Helvetica')).toEqual({ bold: undefined, italic: undefined });
    // The subset prefix is six arbitrary capitals and may spell anything.
    expect(styleOf('BOLDLY+Helvetica').bold).toBeUndefined();
  });
});

/**
 * A composite `Identity-H` font with NO `/ToUnicode`, embedding a real
 * TrueType: the codes are glyph indices, and the only place their Unicode is
 * written down is the font program's own `cmap`.
 */
function identityHNoToUnicodePdf(word: string): Uint8Array {
  const face = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));
  const gidOf = parseTtf(face).glyphForCodepoint;
  const codes = [...word].map((c) => gidOf(c.codePointAt(0)!));
  const hex = codes.map((g) => g.toString(16).padStart(4, '0')).join('');
  const content = `BT /F1 12 Tf 20 100 Td <${hex}> Tj ET`;
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] ' +
      '/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type0 /BaseFont /Roboto /Encoding /Identity-H ' +
      '/DescendantFonts [6 0 R] >>',
    '<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Roboto /DW 500 ' +
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> ' +
      '/FontDescriptor 7 0 R >>',
    `<< /Type /FontDescriptor /FontName /Roboto /Flags 4 /FontFile2 8 0 R >>`,
    `<< /Length ${String(face.length)} /Length1 ${String(face.length)} >>`,
  ];
  const parts: Array<Uint8Array> = [];
  const push = (s: string): void => void parts.push(new TextEncoder().encode(s));
  let at = 0;
  const offsets: Array<number> = [];
  const add = (s: string): void => {
    offsets.push(at);
    push(s);
    at += s.length;
  };
  push('%PDF-1.7\n');
  at = 9;
  objects.forEach((body, i) => {
    if (i === 7) {
      offsets.push(at);
      const head = `8 0 obj\n${body}\nstream\n`;
      push(head);
      parts.push(face);
      const tail = '\nendstream\nendobj\n';
      push(tail);
      at += head.length + face.length + tail.length;
      return;
    }
    add(`${String(i + 1)} 0 obj\n${body}\nendobj\n`);
  });
  const xref = at;
  let trailer = `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) trailer += `${String(off).padStart(10, '0')} 00000 n \n`;
  trailer += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  push(trailer);
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let cursor = 0;
  for (const p of parts) {
    out.set(p, cursor);
    cursor += p.length;
  }
  return out;
}

describe('a composite font with no /ToUnicode (§9.10.2)', () => {
  it('reads its glyphs’ Unicode out of the font program it embeds', () => {
    // Brotli-Prototype-FileA.pdf sets a floor plan's room names in an
    // Identity-H font that ships no /ToUnicode. Decoded to nothing, every run
    // in it was dropped where it stood: "LIVING ROOM" never reached the page.
    const file = PdfFile.parse(identityHNoToUnicodePdf('LIVING'));
    const runs = extractPageText(file, file.pages()[0]!);
    expect(runs.map((r) => r.text).join('')).toBe('LIVING');
  });
});

/** A page whose text is filled with a tiling pattern that paints magenta. */
function patternTextPdf(): Uint8Array {
  const tile = '1 0 1 rg 0 0 5 5 re f';
  const content = '/Pattern cs /P1 scn BT /F1 24 Tf 20 100 Td (tinted) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
      '/Resources << /Font << /F1 5 0 R >> /Pattern << /P1 6 0 R >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] ' +
      `/XStep 10 /YStep 10 /Resources << >> /Length ${String(tile.length)} >>\n` +
      `stream\n${tile}\nendstream`,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

describe('type filled with a pattern (§8.6.6.2)', () => {
  it('takes the pattern’s colour, at the pattern’s own strength', () => {
    // ContentStreamCycleType3insideType3.pdf fills its glyphs with a magenta
    // tiling pattern. A pattern is a content stream, not a colour, and the fill
    // still standing from before painted the words solid black.
    //
    // The tile here paints a 5×5 square on a 10×10 cell — a quarter of it — so
    // the words read as a quarter-strength magenta over white paper, not as
    // solid magenta.
    const file = PdfFile.parse(patternTextPdf());
    const run = extractPageText(file, file.pages()[0]!)[0];
    expect(run?.text).toBe('tinted');
    expect(run?.fillPatternName).toBe('P1');
    expect(run?.colorHex).toBe('FFBFFF');
  });

  it('measures the cell it repeats on, not the marks it paints', () => {
    // A tile is clipped to its /BBox and repeats on /XStep × /YStep, and the
    // marks need agree with neither: the fixture that started this tiles on
    // 55 × 32 and paints two 300 × 300 squares, one of them starting 400 units
    // out. Summed by area, its ink is a hundred times the cell.
    const file = PdfFile.parse(patternTextPdf());
    const page = file.pages()[0]!;
    const tint = patternTint(file, page.resources, 'P1');
    expect(tint?.colorHex).toBe('FF00FF');
    expect(tint?.coverage).toBeCloseTo(0.25, 1);
  });

  it('reads a pattern that covers its whole cell as the solid colour', () => {
    const solid = new TextDecoder()
      .decode(patternTextPdf())
      .replace('1 0 1 rg 0 0 5 5 re f', '1 0 1 rg 0 0 9 9 re f');
    const file = PdfFile.parse(new TextEncoder().encode(solid));
    expect(patternTint(file, file.pages()[0]!.resources, 'P1')?.coverage).toBeGreaterThan(0.75);
  });

  it('says so, because the pattern’s shape is lost even where its colour is not', () => {
    const { losses } = reconstructByLayout(PdfFile.parse(patternTextPdf()), 'positional');
    expect(losses.map((l) => l.severity)).toContain('degraded');
    expect(losses.some((l) => l.detail.includes('tiling pattern'))).toBe(true);
  });
});

describe('the faces a file carries (§9.9)', () => {
  it('lifts the embedded program under the name a run will ask for', () => {
    // A substituted face is never the one the author used: its glyphs are a
    // fraction wider, so every word drifts from where the page put it and the
    // drift runs on down the line.
    const file = PdfFile.parse(identityHNoToUnicodePdf('LIVING'));
    const fonts = collectEmbeddedFonts(file, file.pages());
    expect([...fonts.keys()]).toEqual(['roboto']);
    // A one-face registry answers every style request with the face it has.
    expect(fonts.get('roboto')!.resolveByStyle(false, false).parsed.numGlyphs).toBeGreaterThan(0);
  });

  it('names that face on the run, which is how the layout finds it', () => {
    const file = PdfFile.parse(identityHNoToUnicodePdf('LIVING'));
    const runs = extractPageText(file, file.pages()[0]!);
    expect(runs[0]!.fontName).toBe('roboto');
    const { doc } = reconstructByLayout(file, 'positional');
    expect([...(doc.embeddedFonts ?? new Map()).keys()]).toEqual(['roboto']);
    const shape = doc.body.find((b) => b.kind === 'shape');
    if (shape?.kind !== 'shape') throw new Error('expected a placed line');
    const para = shape.shape.text?.content[0];
    if (para?.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(para.paragraph.runs[0]?.properties.fontFamily?.ascii).toBe('roboto');
  });
});

describe('a simple font reads its codes one byte at a time (§9.6)', () => {
  it('ignores the /ToUnicode codespace width, which is the CMap’s and not the font’s', () => {
    // Read two bytes at a time, every string came apart into pairs: 160F-2019.pdf
    // showed "rémunérations brutes" as "isr".
    const file = PdfFile.parse(simpleFontPdf('Certificat'));
    const runs = extractPageText(file, file.pages()[0]!);
    expect(runs.map((r) => r.text).join('')).toBe('Certificat');
  });
});

describe('page text extraction — real Ream output (E-PDF EP2)', () => {
  it('reads the text back out of a docx → pdf conversion', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>Extract this text</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const file = PdfFile.parse(pdf);
    const runs = extractPageText(file, file.pages()[0]!);
    const all = runs
      .map((r) => r.text)
      .join('')
      .replace(/\s+/g, '');
    expect(all).toContain('Extractthistext');
  });

  it('extracts text positioned down the page in reading order', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>FirstLine</w:t></w:r></w:p><w:p><w:r><w:t>SecondLine</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const file = PdfFile.parse(pdf);
    const runs = extractPageText(file, file.pages()[0]!);
    const first = runs.find((r) => r.text.replace(/\s/g, '').includes('FirstLine'));
    const second = runs.find((r) => r.text.replace(/\s/g, '').includes('SecondLine'));
    expect(first).toBeDefined();
    expect(second).toBeDefined();
    // PDF y grows upward, so the first line sits higher on the page.
    expect(first!.y).toBeGreaterThan(second!.y);
  });
});

/**
 * One line of text with a text-markup annotation over it. `annot` states the
 * subtype and its colour; the quad is the box round the line.
 */
function markedTextPdf(annot: string): Uint8Array {
  const content = 'BT /F1 12 Tf 72 720 Td (Marked) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] ' +
      '/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R ' +
      `/Annots [<< /Type /Annot ${annot} /Rect [70 716 140 734] ` +
      '/QuadPoints [70 734 140 734 70 716 140 716] >>] >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /TrueType /BaseFont /Arial /FirstChar 32 /LastChar 255 ' +
      '/Encoding /WinAnsiEncoding >>',
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\n`;
  pdf += `startxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

describe('text-markup annotations (§12.5.6.10)', () => {
  it('marks the WORDS a highlight covers, not the paper under them', () => {
    // Lifted as artwork a highlight is a band anchored to the page, which is
    // right until the text re-sets and is then a band between two paragraphs
    // it does not mark. Carried on the run it survives the reflow, and Word
    // gets what the annotation meant.
    const file = PdfFile.parse(markedTextPdf('/Subtype /Highlight /C [1 1 0]'));
    const [run] = extractPageText(file, file.pages()[0]!);
    expect(run?.markup?.highlightHex).toBe('FFFF00');
    // And the band itself is NOT painted a second time.
    expect(collectPageAppearances(file, file.pages()[0]!)).toHaveLength(0);
  });

  it('reads an underline, a squiggle and a strikeout off the same quads', () => {
    const under = PdfFile.parse(markedTextPdf('/Subtype /Underline /C [0 0.6 0]'));
    expect(extractPageText(under, under.pages()[0]!)[0]?.markup).toEqual({
      underline: 'single',
      underlineHex: '009900',
    });
    const wavy = PdfFile.parse(markedTextPdf('/Subtype /Squiggly /C [0 0 1]'));
    expect(extractPageText(wavy, wavy.pages()[0]!)[0]?.markup?.underline).toBe('wave');
    const struck = PdfFile.parse(markedTextPdf('/Subtype /StrikeOut /C [1 0 0]'));
    expect(extractPageText(struck, struck.pages()[0]!)[0]?.markup?.strike).toBe(true);
  });

  it('leaves a line the quads do not cover unmarked', () => {
    // A quad drawn round one line of a paragraph must not claim the lines
    // above and below it, so the test is the baseline falling inside the box.
    const file = PdfFile.parse(
      new TextEncoder().encode(
        new TextDecoder().decode(markedTextPdf('/Subtype /Highlight /C [1 1 0]')).replace(
          '/QuadPoints [70 734 140 734 70 716 140 716]',
          // The same quad, two lines further up the page.
          '/QuadPoints [70 764 140 764 70 746 140 746]',
        ),
      ),
    );
    const [run] = extractPageText(file, file.pages()[0]!);
    // The line is still read — it is only no longer claimed by the quad.
    expect(run?.text).toBe('Marked');
    expect(run?.markup).toBeUndefined();
  });
});

/**
 * One line of text with a filled bar drawn near it — the shape a PDF gives an
 * underline, a strikeout, and also a table's cell border.
 */
function ruledTextPdf(bar: string): Uint8Array {
  const content = `BT /F1 10 Tf 72 720 Td (Underlined) Tj ET ${bar}`;
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] ' +
      '/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /TrueType /BaseFont /Arial /FirstChar 32 /LastChar 255 ' +
      '/Encoding /WinAnsiEncoding >>',
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\n`;
  pdf += `startxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

/** The runs and the paths of a one-page PDF, both placed on the shown page. */
function pageOf(bytes: Uint8Array): { runs: Array<TextRun>; vectors: Array<PdfVector> } {
  const file = PdfFile.parse(bytes);
  const page = file.pages()[0]!;
  const shown = displayOf(page);
  return {
    runs: placeRuns(extractPageText(file, page), shown),
    vectors: placeVectors(collectPageVectors(file, page, []).vectors, shown),
  };
}

describe('the underline a PDF draws rather than states', () => {
  it('reads a bar under the words onto the run, and takes the bar', () => {
    // Nothing in ISO 32000-1 says a run is underlined: the page paints a thin
    // bar beneath it. Anchored where it was drawn it is a dash between two
    // lines the moment the text re-sets.
    const { runs, vectors } = pageOf(ruledTextPdf('0 0 0.5 rg 72 718 48 0.7 re f'));
    const ruled = markDrawnRules(runs, vectors);
    expect(ruled.runs[0]?.markup).toEqual({ underline: 'single', underlineHex: '000080' });
    expect(ruled.consumed.size).toBe(1);
  });

  it('leaves a table’s rule alone, however like an underline it looks', () => {
    // A cell border begins at the CELL and ends at the cell's other edge,
    // however short the text stops: TAMReview.pdf rules its tables 86 to 509
    // across a measure whose text stops well before it, and a proportional
    // allowance read all fourteen of them as underlines.
    const { runs, vectors } = pageOf(ruledTextPdf('0 0 0.5 rg 72 718 300 0.7 re f'));
    const ruled = markDrawnRules(runs, vectors);
    expect(ruled.runs[0]?.markup).toBeUndefined();
    expect(ruled.consumed.size).toBe(0);
  });

  it('does not underline in white, and does not read a seam as a mark', () => {
    const white = pageOf(ruledTextPdf('1 1 1 rg 72 718 48 0.7 re f'));
    expect(markDrawnRules(white.runs, white.vectors).consumed.size).toBe(0);
    const seam = pageOf(ruledTextPdf('0 0 0.5 rg 72 718 48 0.05 re f'));
    expect(markDrawnRules(seam.runs, seam.vectors).consumed.size).toBe(0);
  });

  it('reads a bar ACROSS the words as a strikeout', () => {
    const { runs, vectors } = pageOf(ruledTextPdf('0 0 0 rg 72 723 48 0.7 re f'));
    expect(markDrawnRules(runs, vectors).runs[0]?.markup?.strike).toBe(true);
  });
});

describe('the margins a page\u2019s words say it had', () => {
  const runAt = (y: number, sizePt: number): TextRun =>
    ({
      text: 'x',
      x: 72,
      y,
      endX: 200,
      endY: y,
      fontSizePt: sizePt,
      fontKey: 'F1',
    }) as TextRun;

  it('measures the top margin to the top of the LINE, not to its baseline', () => {
    // A run carries a baseline and a margin is to the top of the line, so a
    // margin measured to the baseline puts every converted PDF a whole
    // ascender too low: on annotation-stamp.pdf the word "Stamp" slid under
    // the stamp anchored above it.
    const section = { pageSize: { width: 612, height: 792 } } as never;
    const page = [{ width: 612, height: 792 }];
    const small = withMeasuredMargins(section, page, [[runAt(700, 10)]]);
    const large = withMeasuredMargins(section, page, [[runAt(700, 30)]]);
    // The same baseline in a bigger face starts higher up the page, so less
    // paper is left above it.
    expect((small?.margins?.top ?? 0) as number).toBeCloseTo(792 - 700 - 8, 3);
    expect((large?.margins?.top ?? 0) as number).toBeCloseTo(792 - 700 - 24, 3);
  });
});
