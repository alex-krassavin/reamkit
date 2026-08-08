// E-PDF EP2 — /ToUnicode CMap parsing and end-to-end page text extraction. The
// honest test reads the text back out of a PDF Ream itself wrote.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { parseToUnicodeCMap } from '@/pdf-reader/cmap';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

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
