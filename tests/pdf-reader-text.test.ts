// E-PDF EP2 — /ToUnicode CMap parsing and end-to-end page text extraction. The
// honest test reads the text back out of a PDF Ream itself wrote.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { parseToUnicodeCMap } from '@/pdf-reader/cmap';
import { PdfFile } from '@/pdf-reader/document';
import { collectEmbeddedFonts } from '@/pdf-reader/embedded-fonts';
import { extractPageText } from '@/pdf-reader/text';
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
