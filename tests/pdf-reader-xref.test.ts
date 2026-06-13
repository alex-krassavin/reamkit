// E-PDF EP7 — cross-reference streams + object streams. A hand-built compressed
// PDF (objects packed in an /ObjStm, cross-reference as an /XRef stream) is the
// modern form Ream's writer never emits; the reader must resolve objects out of
// the object stream and walk the page tree through them.

import { describe, expect, it } from 'vitest';

import { buildCompressedPdf } from './fixtures/build-compressed-pdf';
import { PdfFile } from '@/pdf-reader/document';
import { PdfName } from '@/pdf/objects';

describe('xref streams + object streams (E-PDF EP7)', () => {
  it('reads a /Catalog packed inside an object stream via an xref stream', () => {
    const file = PdfFile.parse(buildCompressedPdf());
    const type = file.catalog.get('Type');
    expect(type instanceof PdfName ? type.value : '').toBe('Catalog');
  });

  it('walks the page tree through compressed objects', () => {
    const pages = PdfFile.parse(buildCompressedPdf()).pages();
    expect(pages).toHaveLength(1);
    expect(pages[0]!.mediaBox).toEqual([0, 0, 200, 200]);
  });

  it('recovers compressed objects via brute force when startxref is broken', () => {
    // No usable xref stream → the brute-force scan must find the object stream
    // and lift the /Catalog out of it.
    const file = PdfFile.parse(buildCompressedPdf({ brokenStartxref: true }));
    const type = file.catalog.get('Type');
    expect(type instanceof PdfName ? type.value : '').toBe('Catalog');
    expect(file.pages()).toHaveLength(1);
  });
});
