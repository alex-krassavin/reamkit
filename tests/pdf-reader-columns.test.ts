// E-PDF EP17 — multi-column reconstruction. An untagged two-column page has the
// left and right columns sharing baselines; grouping by baseline alone would
// interleave them (L1 R1 L2 R2 …). The reader detects the central gutter and
// reads each column in full (L1 L2 … then R1 R2 …).

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

const ROWS = 18;

// A page with two columns of short runs (left x=72, right x=380) sharing each
// baseline. 36 runs clears the heuristic's confidence threshold.
function twoColumnPdf(): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const ops = ['BT /F1 10 Tf'];
  for (let i = 0; i < ROWS; i++) {
    const y = 720 - i * 24;
    const n = String(i + 1).padStart(2, '0');
    ops.push(`1 0 0 1 72 ${y} Tm (L${n}) Tj`);
    ops.push(`1 0 0 1 380 ${y} Tm (R${n}) Tj`);
  }
  ops.push('ET');
  const content = doc.add(stream({}, new TextEncoder().encode(ops.join('\n'))));
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: 1 });
  const pagesRef = doc.add(pagesMap);
  const page = doc.add(
    dict({
      Type: name('Page'),
      Parent: pagesRef,
      MediaBox: [0, 0, 612, 792],
      Resources: dict({ Font: dict({ F1: font }) }),
      Contents: content,
    }),
  );
  (pagesMap.get('Kids') as Array<PdfValue>).push(page);
  const catalog = doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }));
  return doc.build(catalog);
}

const bodyTokens = (pdf: Uint8Array): Array<string> => {
  const text = Ream.parse(pdf)
    .flow.body.map((el) =>
      el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '',
    )
    .join(' ');
  return text.match(/[LR]\d\d/g) ?? [];
};

describe('two-column reconstruction (E-PDF EP17)', () => {
  it('reads the left column fully before the right', () => {
    const tokens = bodyTokens(twoColumnPdf());
    const left = Array.from({ length: ROWS }, (_, i) => `L${String(i + 1).padStart(2, '0')}`);
    const right = Array.from({ length: ROWS }, (_, i) => `R${String(i + 1).padStart(2, '0')}`);
    expect(tokens.slice(0, ROWS)).toEqual(left); // left column, top-to-bottom
    expect(tokens.slice(ROWS)).toEqual(right); // then the right column
  });
});
