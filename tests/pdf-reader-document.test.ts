// E-PDF EP1 — the document layer: classic xref + trailer + page tree + stream
// decoding. The honest test is a round-trip — read back PDFs that Ream's own
// writer produced (a writer-built micro-PDF, and a real docx → pdf conversion).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';
import { zlibSync } from 'fflate';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { PdfDict } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';
import { PdfFile } from '@/pdf-reader/document';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const dec = (b: Uint8Array): string => new TextDecoder().decode(b);

// A minimal one-page PDF built with the writer — exercises the classic xref the
// reader must walk, with no fonts or compression in the way.
function tinyPdf(content: Uint8Array, contentDict: PdfDict): Uint8Array {
  const doc = new PdfDocument();
  const contentRef = doc.add(stream(Object.fromEntries(contentDict), content));
  const page = dict({ Type: name('Page'), MediaBox: [0, 0, 200, 100], Contents: contentRef });
  const pageRef = doc.add(page);
  const pagesRef = doc.add(dict({ Type: name('Pages'), Kids: [pageRef], Count: 1 }));
  page.set('Parent', pagesRef);
  const catalog = doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }));
  return doc.build(catalog);
}

describe('PDF document layer — classic xref + page tree (E-PDF EP1)', () => {
  it('reads back a writer-produced PDF: pages, MediaBox, content', () => {
    const body = new TextEncoder().encode('BT /F1 24 Tf 72 60 Td (Hello) Tj ET');
    const file = PdfFile.parse(tinyPdf(body, new Map([['Length', body.length]])));
    const pages = file.pages();
    expect(pages.length).toBe(1);
    expect(pages[0]!.mediaBox).toEqual([0, 0, 200, 100]);
    expect(dec(file.pageContent(pages[0]!))).toContain('(Hello) Tj');
  });

  it('decodes a FlateDecode content stream', () => {
    const raw = new TextEncoder().encode('BT (compressed) Tj ET');
    const comp = zlibSync(raw);
    const file = PdfFile.parse(
      tinyPdf(
        comp,
        new Map([
          ['Length', comp.length],
          ['Filter', name('FlateDecode')],
        ]),
      ),
    );
    expect(dec(file.pageContent(file.pages()[0]!))).toBe('BT (compressed) Tj ET');
  });

  it('recovers via a brute-force scan when the xref offset is corrupt', () => {
    const body = new TextEncoder().encode('BT (recovered) Tj ET');
    const bytes = tinyPdf(body, new Map([['Length', body.length]]));
    // Corrupt the startxref offset so the classic path fails and recovery kicks in.
    const sx = dec(bytes).lastIndexOf('startxref');
    const corrupt = bytes.slice();
    corrupt[sx + 10] = '9'.charCodeAt(0); // point startxref into nowhere
    corrupt[sx + 11] = '9'.charCodeAt(0);
    const file = PdfFile.parse(corrupt);
    expect(file.pages().length).toBe(1);
    expect(dec(file.pageContent(file.pages()[0]!))).toContain('(recovered) Tj');
  });

  it('round-trips a real Ream-written PDF (docx → pdf → read back)', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>Round trip</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const file = PdfFile.parse(pdf);
    const pages = file.pages();
    expect(pages.length).toBeGreaterThanOrEqual(1);
    const mb = pages[0]!.mediaBox;
    expect(mb[2] - mb[0]).toBeGreaterThan(500); // ~A4 width in pt
    expect(mb[3] - mb[1]).toBeGreaterThan(700); // ~A4 height in pt
    expect(dec(file.pageContent(pages[0]!))).toContain('BT'); // a real text object
  });
});
