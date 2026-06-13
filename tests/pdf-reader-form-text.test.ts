// E-PDF EP13 — text drawn inside a Form XObject. The page only paints the form
// (`/Fm0 Do`); the reader must recurse into it (composing its /Matrix and using
// its own fonts) to recover the text, which page-level interpretation misses.

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import { dict, name, stream } from '@/pdf/objects';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';
import { PdfDocument } from '@/pdf/writer';

function buildFormTextPdf(): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const form = doc.add(
    stream(
      {
        Type: name('XObject'),
        Subtype: name('Form'),
        BBox: [0, 0, 300, 300],
        Resources: dict({ Font: dict({ F1: font }) }),
      },
      new TextEncoder().encode('BT /F1 12 Tf 40 250 Td (FormXObjectText) Tj ET'),
    ),
  );
  const content = doc.add(stream({}, new TextEncoder().encode('q 1 0 0 1 0 0 cm /Fm0 Do Q')));
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 1 });
  const pagesRef = doc.add(pagesMap);
  const page = doc.add(
    dict({
      Type: name('Page'),
      Parent: pagesRef,
      MediaBox: [0, 0, 300, 300],
      Resources: dict({ XObject: dict({ Fm0: form }) }),
      Contents: content,
    }),
  );
  (pagesMap.get('Kids') as Array<PdfValue>).push(page);
  const catalog = doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }));
  return doc.build(catalog);
}

describe('form-XObject text extraction (E-PDF EP13)', () => {
  it('recovers text drawn inside a Form XObject', () => {
    const file = PdfFile.parse(buildFormTextPdf());
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('')
      .replace(/\s/g, '');
    expect(text).toContain('FormXObjectText');
  });
});
