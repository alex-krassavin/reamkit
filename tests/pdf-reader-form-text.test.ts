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

/**
 * A page that paints nothing itself and carries one widget annotation whose
 * appearance draws a caption — the shape a form's push button takes.
 */
function buildWidgetTextPdf(): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const appearance = doc.add(
    stream(
      {
        Type: name('XObject'),
        Subtype: name('Form'),
        BBox: [0, 0, 120, 20],
        Resources: dict({ Font: dict({ F1: font }) }),
      },
      new TextEncoder().encode('BT /F1 12 Tf 4 5 Td (Reinitialiser) Tj ET'),
    ),
  );
  const content = doc.add(stream({}, new TextEncoder().encode('')));
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 1 });
  const pagesRef = doc.add(pagesMap);
  const page = doc.add(
    dict({
      Type: name('Page'),
      Parent: pagesRef,
      MediaBox: [0, 0, 300, 300],
      Contents: content,
      Annots: [
        dict({
          Type: name('Annot'),
          Subtype: name('Widget'),
          FT: name('Btn'),
          Rect: [40, 200, 160, 220],
          AP: dict({ N: appearance }),
        }),
      ],
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

  it('recovers the words an annotation draws, at the /Rect it draws them in', () => {
    // §12.5.5 — a field's value and a button's caption live in the widget's own
    // appearance, not in the page content. Only the ARTWORK was being lifted, so
    // 160F-2019.pdf's reset button arrived as a tinted rectangle with nothing
    // written on it.
    const file = PdfFile.parse(buildWidgetTextPdf());
    const runs = extractPageText(file, file.pages()[0]!);
    const caption = runs.find((r) => r.text.includes('Reinitialiser'));
    expect(caption).toBeDefined();
    // The appearance is authored at its own origin; the /Rect is what places it.
    expect(caption!.x).toBeCloseTo(44, 0);
    expect(caption!.y).toBeCloseTo(205, 0);
  });
});
