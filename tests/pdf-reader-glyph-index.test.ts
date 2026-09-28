// §9.6.6.4 — a glyph named by its index (`g18`) says nothing itself, and the
// TrueType program it indexes still does: its `cmap`, read backwards.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { parseTtf } from '@/core/font/ttf-parser';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

/** A one-page PDF showing codes 1 and 2 of a simple TrueType font named by index. */
function indexNamedPdf(names: ReadonlyArray<string>): Uint8Array {
  const face = new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf'));
  const content = 'BT /F1 12 Tf 20 100 Td <0102> Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] ' +
      '/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /TrueType /BaseFont /ABCDEF+Roboto /FirstChar 1 /LastChar 2 ' +
      `/Widths [600 600] /Encoding << /Differences [1 ${names.map((n) => `/${n}`).join(' ')}] >> ` +
      '/FontDescriptor 6 0 R >>',
    '<< /Type /FontDescriptor /FontName /ABCDEF+Roboto /Flags 4 /FontFile2 7 0 R >>',
  ];
  const chunks: Array<Uint8Array> = [];
  const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
  let at = 0;
  const offsets: Array<number> = [];
  const push = (bytes: Uint8Array): void => {
    chunks.push(bytes);
    at += bytes.length;
  };
  push(enc('%PDF-1.7\n'));
  objects.forEach((body, i) => {
    offsets.push(at);
    push(enc(`${String(i + 1)} 0 obj\n${body}\nendobj\n`));
  });
  offsets.push(at);
  push(enc(`7 0 obj\n<< /Length ${String(face.length)} >>\nstream\n`));
  push(face);
  push(enc('\nendstream\nendobj\n'));
  const xref = at;
  let trailer = `xref\n0 8\n0000000000 65535 f \n`;
  for (const off of offsets) trailer += `${String(off).padStart(10, '0')} 00000 n \n`;
  trailer += `trailer\n<< /Size 8 /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  push(enc(trailer));
  const out = new Uint8Array(at);
  let pos = 0;
  for (const c of chunks) {
    out.set(c, pos);
    pos += c.length;
  }
  return out;
}

describe('a glyph named by its index (§9.6.6.4)', () => {
  it('reads the character the program maps onto that glyph', () => {
    // TAMReview.pdf sets its text in a subset whose glyphs are named `g18`,
    // `g152`; unread, every one of them was traced as a drawing instead.
    const gidOf = parseTtf(
      new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
    ).glyphForCodepoint;
    const file = PdfFile.parse(
      indexNamedPdf([`g${String(gidOf(0x5a))}`, `glyph${String(gidOf(0x71))}`]),
    );
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('');
    expect(text).toBe('Zq');
  });
});
