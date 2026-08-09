// E-PDF EP1 — the document layer: classic xref + trailer + page tree + stream
// decoding. The honest test is a round-trip — read back PDFs that Ream's own
// writer produced (a writer-built micro-PDF, and a real docx → pdf conversion).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';
import { zlibSync } from 'fflate';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { PdfDict, PdfValue } from '@/pdf/objects';
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

describe('a filter nothing here can undo says so (§7.4)', () => {
  it('reports the filter by name instead of returning an empty document', () => {
    // Brotli-Prototype-FileA.pdf compresses its CROSS-REFERENCE stream with
    // `/BrotliDecode` (PDF 2.0). Undecoded, the page tree never resolves and
    // all twenty-five pages go missing — with nothing in the report to say why.
    const content = 'BT ET';
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>',
      `<< /Filter /BrotliDecode /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    ];
    let pdf = '%PDF-2.0\n';
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

    const file = PdfFile.parse(new TextEncoder().encode(pdf));
    file.streamData(file.resolve(file.pages()[0]!.dict.get('Contents')!) as never);
    expect([...file.unknownFilters]).toContain('BrotliDecode');
  });

  it('says nothing about the filters it does undo, or leaves to others', () => {
    const raw = zlibSync(new TextEncoder().encode('BT ET'));
    const file = PdfFile.parse(
      new TextEncoder().encode('%PDF-1.7\ntrailer\n<< /Size 1 >>\n%%EOF\n'),
    );
    file.streamData(stream({ Filter: name('FlateDecode') }, raw));
    file.streamData(stream({ Filter: name('DCTDecode') }, raw));
    expect([...file.unknownFilters]).toHaveLength(0);
  });
});

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
        new Map<string, PdfValue>([
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
