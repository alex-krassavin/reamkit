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
    const table = Ream.parse(onePage(ops)).flow.body.find((el) => el.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') return;
    expect(table.table.rows).toHaveLength(ROWS);
    expect(table.table.grid).toHaveLength(4);
    const first = table.table.rows[0]!.cells.map((c) =>
      c.content
        .flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text) : []))
        .join(''),
    );
    expect(first).toEqual(['A01', 'B01', 'C01', 'D01']);
  });

  it('keeps a line that overhangs its column in ONE cell, beside its neighbour', () => {
    // A gutter is a BAND, and the middle of it is only a guess at where the
    // columns divide: a gutter is where the FEWEST lines cross, not where none
    // do. ZapfDingbats.pdf's lead paragraph runs a few points past the middle
    // of the first gutter, and cut there it either wrapped inside a cell too
    // narrow for it — every wrapped line costing the groups beside it an
    // entry, two pages coming out as four — or swallowed the glyph standing in
    // the next column.
    const ops = ['BT /F1 9 Tf'];
    for (let i = 0; i < ROWS; i++) {
      const y = 720 - i * 14;
      const n = String(i + 1).padStart(2, '0');
      // The first four rows carry prose that overruns the first column — past
      // the middle of the gutter, and short of the column beyond it.
      if (i < 4) ops.push(`1 0 0 1 72 ${String(y)} Tm (Prose overhanging ${n}) Tj`);
      else ops.push(`1 0 0 1 72 ${String(y)} Tm (A${n}) Tj`);
      ops.push(`1 0 0 1 200 ${String(y)} Tm (B${n}) Tj`);
      ops.push(`1 0 0 1 300 ${String(y)} Tm (C${n}) Tj`);
      ops.push(`1 0 0 1 420 ${String(y)} Tm (D${n}) Tj`);
    }
    ops.push('ET');
    const table = Ream.parse(onePage(ops)).flow.body.find((el) => el.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') return;
    const cells = (row: number): Array<string> =>
      table.table.rows[row]!.cells.map((c) =>
        c.content
          .flatMap((el) => (el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text) : []))
          .join(''),
      );
    // The prose is one cell, and B01 is still its own.
    expect(cells(0)[0]).toBe('Prose overhanging 01');
    expect(cells(0)[1]).toBe('B01');
    // …and the rows below are untouched.
    expect(cells(5)).toEqual(['A06', 'B06', 'C06', 'D06']);
  });

  it('leaves the line the page hangs ABOVE its ruling where it stands', () => {
    // A line that crosses every column is not a row of the table.
    // ZapfDingbats.pdf heads each sheet with two red lines of provenance that
    // run wider than the frame drawn under them, and squeezed into a cell they
    // wrapped — and every wrapped line cost the groups beside them an entry.
    // The table under them starts at its OWN left, not at theirs: read from
    // there, three groups and five hundred entries stood sixteen points left of
    // where the file has them.
    const ops = [
      'BT /F1 9 Tf',
      '1 0 0 1 40 740 Tm (A line of provenance right across the sheet, wider than the table) Tj',
    ];
    for (let i = 0; i < ROWS; i++) {
      const y = 720 - i * 14;
      const n = String(i + 1).padStart(2, '0');
      ops.push(`1 0 0 1 72 ${String(y)} Tm (A${n}) Tj`);
      ops.push(`1 0 0 1 200 ${String(y)} Tm (B${n}) Tj`);
      ops.push(`1 0 0 1 300 ${String(y)} Tm (C${n}) Tj`);
      ops.push(`1 0 0 1 420 ${String(y)} Tm (D${n}) Tj`);
    }
    ops.push('ET');
    const body = Ream.parse(onePage(ops)).flow.body;
    const first = body[0];
    expect(first?.kind).toBe('paragraph');
    if (first?.kind === 'paragraph') {
      expect(first.paragraph.runs.map((r) => r.text).join('')).toContain('A line of provenance');
    }
    const table = body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('the page is a table');
    expect(table.table.rows).toHaveLength(ROWS);
    // The table stands in from the page's text by as much as its own left does.
    expect(table.table.properties.indentPt as number).toBeGreaterThan(20);
  });

  it('centres a cell the page centred, and leaves the rest flush', () => {
    // ZapfDingbats.pdf centres its title over the first group, inside the grey
    // panel drawn behind it; set flush left it came out of that panel at the
    // wrong end. An ordinary line that merely reaches the far edge of a wide
    // column is NOT centred, however even its margins look.
    const ops = ['BT /F1 9 Tf'];
    for (let i = 0; i < ROWS; i++) {
      const y = 720 - i * 14;
      const n = String(i + 1).padStart(2, '0');
      // Row 0 stands in from both sides of a column whose rows start at 72.
      if (i === 0) ops.push(`1 0 0 1 113 ${String(y)} Tm (THE GROUP) Tj`);
      else ops.push(`1 0 0 1 72 ${String(y)} Tm (A${n} entry) Tj`);
      ops.push(`1 0 0 1 250 ${String(y)} Tm (B${n}) Tj`);
      ops.push(`1 0 0 1 330 ${String(y)} Tm (C${n}) Tj`);
      ops.push(`1 0 0 1 430 ${String(y)} Tm (D${n}) Tj`);
    }
    ops.push('ET');
    const table = Ream.parse(onePage(ops)).flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('the page is a table');
    const alignment = (row: number): string | undefined => {
      const cell = table.table.rows[row]!.cells[0]!.content[0];
      return cell?.kind === 'paragraph' ? cell.paragraph.properties.alignment : undefined;
    };
    expect(alignment(0)).toBe('center');
    expect(alignment(1)).not.toBe('center');
  });

  it('gives each row the height the PAGE gave it', () => {
    // A row laid out by the height of its own text closes up wherever the page
    // left air, and everything anchored to the sheet — ZapfDingbats.pdf's grey
    // title panel, the frame around its table — then stands somewhere else.
    const ops = ['BT /F1 9 Tf'];
    for (let i = 0; i < ROWS; i++) {
      // A double step after the third row: the page left air there.
      const y = 720 - i * 14 - (i > 2 ? 20 : 0);
      const n = String(i + 1).padStart(2, '0');
      ops.push(`1 0 0 1 72 ${String(y)} Tm (A${n}) Tj`);
      ops.push(`1 0 0 1 180 ${String(y)} Tm (B${n}) Tj`);
      ops.push(`1 0 0 1 300 ${String(y)} Tm (C${n}) Tj`);
      ops.push(`1 0 0 1 420 ${String(y)} Tm (D${n}) Tj`);
    }
    ops.push('ET');
    const table = Ream.parse(onePage(ops)).flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('the page is a table');
    const heights = table.table.rows.map((r) => r.properties.height as number | undefined);
    expect(heights[0]).toBeCloseTo(14, 1);
    expect(heights[2]).toBeCloseTo(34, 1); // the row the page left air under
    expect(table.table.rows[0]!.properties.heightRule).toBe('atLeast');
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
