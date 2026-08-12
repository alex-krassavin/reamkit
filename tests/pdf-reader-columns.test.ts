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

/** A one-page PDF of `ops`, drawn in Helvetica on a letter sheet. */
function onePage(ops: ReadonlyArray<string>): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
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

// A page with two columns of short runs (left x=72, right x=380) sharing each
// baseline. 36 runs clears the heuristic's confidence threshold.
// `heading` prepends a line reaching right across the page, the way a paper's
// title does.
function twoColumnPdf(heading = false): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const ops = ['BT /F1 10 Tf'];
  if (heading) {
    ops.push('1 0 0 1 72 760 Tm (TITLE ACROSS THE WHOLE PAGE WIDTH HERE) Tj');
    ops.push('1 0 0 1 72 744 Tm (AUTHORS ACROSS THE WHOLE PAGE WIDTH TOO) Tj');
  }
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
  return text.match(/[LR]\d\d|TITLE|AUTHORS|FOOTER/g) ?? [];
};

describe('two-column reconstruction (E-PDF EP17)', () => {
  it('reads the left column fully before the right', () => {
    const tokens = bodyTokens(twoColumnPdf());
    const left = Array.from({ length: ROWS }, (_, i) => `L${String(i + 1).padStart(2, '0')}`);
    const right = Array.from({ length: ROWS }, (_, i) => `R${String(i + 1).padStart(2, '0')}`);
    expect(tokens.slice(0, ROWS)).toEqual(left); // left column, top-to-bottom
    expect(tokens.slice(ROWS)).toEqual(right); // then the right column
  });

  it('reads a full-width heading before the columns it stands over', () => {
    // A page is rarely two columns and nothing else. comments.pdf is a
    // conference paper — a full-width title, a full-width author block, then
    // two columns — and asking for a band NO line crosses found no gutter at
    // all: its columns were joined line by line, "Abstract and is used for the
    // application logic of browser-based productivity Dynamic languages…".
    const tokens = bodyTokens(twoColumnPdf(true));
    expect(tokens.slice(0, 2)).toEqual(['TITLE', 'AUTHORS']);
    const left = Array.from({ length: ROWS }, (_, i) => `L${String(i + 1).padStart(2, '0')}`);
    const right = Array.from({ length: ROWS }, (_, i) => `R${String(i + 1).padStart(2, '0')}`);
    expect(tokens.slice(2, 2 + ROWS)).toEqual(left);
    expect(tokens.slice(2 + ROWS)).toEqual(right);
  });

  it('reads a page of THREE columns one at a time', () => {
    // A page is not always two columns and a middle.
    // chrome-text-selection-markedContent.pdf is an analyst's report — two
    // columns of comment and a sidebar of figures down the right — and asked
    // for the ONE best gutter it took the body's and read the sidebar as part
    // of the text, opening the page with the guidance box from the margin.
    // The lines FILL their columns, which is what tells a page set in columns
    // from a page ruled into them (see the ruled test below).
    const ops = ['BT /F1 10 Tf'];
    for (let i = 0; i < ROWS; i++) {
      const y = 720 - i * 24;
      const n = String(i + 1).padStart(2, '0');
      ops.push(`1 0 0 1 72 ${String(y)} Tm (L${n} and a line of prose that fills) Tj`);
      ops.push(`1 0 0 1 280 ${String(y)} Tm (M${n} and a line of prose to fill) Tj`);
      ops.push(`1 0 0 1 480 ${String(y)} Tm (R${n} and prose filling it) Tj`);
    }
    ops.push('ET');
    const text = Ream.parse(onePage(ops))
      .flow.body.map((el) =>
        el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '',
      )
      .join(' ');
    const tokens = text.match(/[LMR]\d\d/gu) ?? [];
    const column = (letter: string): Array<string> =>
      Array.from({ length: ROWS }, (_, i) => `${letter}${String(i + 1).padStart(2, '0')}`);
    expect(tokens.slice(0, ROWS)).toEqual(column('L'));
    expect(tokens.slice(ROWS, ROWS * 2)).toEqual(column('M'));
    expect(tokens.slice(ROWS * 2)).toEqual(column('R'));
  });

  it('reads a page RULED into columns by its rows, not by its columns', () => {
    // A page of columns and a page of a table look alike from here: both have
    // gutters, and both put their lines on one baseline grid. What separates
    // them is the CELL — a line of prose fills its measure and a cell does not.
    // ZapfDingbats.pdf is five hundred entries in a table of three groups, and
    // read by column every row of it was torn into three.
    const ops = ['BT /F1 9 Tf'];
    for (let i = 0; i < ROWS; i++) {
      const y = 720 - i * 14;
      const n = String(i + 1).padStart(2, '0');
      ops.push(`1 0 0 1 72 ${String(y)} Tm (A${n}) Tj`);
      ops.push(`1 0 0 1 180 ${String(y)} Tm (B${n}) Tj`);
      ops.push(`1 0 0 1 300 ${String(y)} Tm (C${n}) Tj`);
      ops.push(`1 0 0 1 420 ${String(y)} Tm (D${n}) Tj`);
    }
    ops.push('ET');
    const text = Ream.parse(onePage(ops))
      .flow.body.map((el) =>
        el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '',
      )
      .join(' ');
    const tokens = text.match(/[A-D]\d\d/gu) ?? [];
    // Row by row: A01 B01 C01 D01, A02 B02 …
    expect(tokens.slice(0, 8)).toEqual(['A01', 'B01', 'C01', 'D01', 'A02', 'B02', 'C02', 'D02']);
  });

  it('reads a full-width FOOTER after the columns it stands under', () => {
    // A spanning line is what breaks a band, so it always stands at the foot of
    // its own — the columns of that band are the ones above it. Read ahead of
    // them, bug1997343.pdf's page number came out between the date and the
    // abstract.
    const ops = ['BT /F1 10 Tf'];
    for (let i = 0; i < ROWS; i++) {
      const y = 720 - i * 24;
      const n = String(i + 1).padStart(2, '0');
      ops.push(`1 0 0 1 72 ${String(y)} Tm (L${n}) Tj`);
      ops.push(`1 0 0 1 380 ${String(y)} Tm (R${n}) Tj`);
    }
    ops.push('1 0 0 1 72 200 Tm (FOOTER ACROSS THE WHOLE PAGE WIDTH HERE) Tj');
    ops.push('ET');
    const tokens = bodyTokens(onePage(ops));
    expect(tokens[tokens.length - 1]).toBe('FOOTER');
  });
});
