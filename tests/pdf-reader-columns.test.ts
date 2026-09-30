// E-PDF EP17 — multi-column reconstruction. An untagged two-column page has the
// left and right columns sharing baselines; grouping by baseline alone would
// interleave them (L1 R1 L2 R2 …). The reader detects the central gutter and
// reads each column in full (L1 L2 … then R1 R2 …).

import { describe, expect, it } from 'vitest';

import type { BodyElement } from '@/core/document-model';
import type { PdfValue } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { standardWidth } from '@/pdf-reader/standard-widths';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

const ROWS = 18;

/** A one-page PDF of `ops`, drawn in Helvetica (or `face`) on a letter sheet. */
function onePage(ops: ReadonlyArray<string>, face = 'Helvetica'): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name(face) }));
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

  it('reads a page set in two columns of TABLES by its two columns', () => {
    // A table's columns stand apart line after line as a page's do. canvas.pdf
    // sets two columns of Name, Type and Default tables under headings that
    // run across them, and read at every gap its sheets came back in five
    // columns a word wide each, on seven pages where it has two.
    const ops = ['BT /F1 9 Tf'];
    for (const [side, x] of [
      ['L', 40],
      ['R', 320],
    ] as const) {
      let y = 740;
      for (let s = 0; s < 4; s++) {
        const head = `${side}h${String(s)} a heading over the table under it`;
        ops.push(`1 0 0 1 ${String(x)} ${String(y)} Tm (${head}) Tj`);
        for (let r = 0; r < 5; r++) {
          y -= 12;
          const n = `${String(s)}${String(r)}`;
          ops.push(`1 0 0 1 ${String(x)} ${String(y)} Tm (${side}n${n}) Tj`);
          ops.push(`1 0 0 1 ${String(x + 70)} ${String(y)} Tm (${side}t${n}) Tj`);
          ops.push(`1 0 0 1 ${String(x + 140)} ${String(y)} Tm (${side}v${n}) Tj`);
        }
        y -= 24;
      }
    }
    ops.push('ET');
    const textIn = (els: ReadonlyArray<BodyElement>): string =>
      els
        .map((el) =>
          el.kind === 'paragraph'
            ? el.paragraph.runs.map((r) => r.text).join('')
            : el.kind === 'table'
              ? el.table.rows.flatMap((row) => row.cells.map((c) => textIn(c.content))).join(' ')
              : '',
        )
        .join(' ');
    const tokens = textIn(Ream.parse(onePage(ops)).flow.body).match(/[LR][hntv]\d+/gu) ?? [];
    expect(tokens).toHaveLength(2 * 4 * 16);
    // Every token of the left column before any of the right…
    const right = tokens.findIndex((t) => t.startsWith('R'));
    expect(tokens.slice(0, right).every((t) => t.startsWith('L'))).toBe(true);
    expect(tokens.slice(right).every((t) => t.startsWith('R'))).toBe(true);
    // …and a row's cells read across it, as the page sets them.
    expect(tokens.slice(0, 7)).toEqual(['Lh0', 'Ln00', 'Lt00', 'Lv00', 'Ln01', 'Lt01', 'Lv01']);
  });

  it('reads a JUSTIFIED page in its columns', () => {
    // Every full line of a justified column ends where the column does, and
    // only a paragraph's first line starts anywhere but its edge: the right
    // edges agree more than the left ones, which is what a column of amounts
    // looks like. comments.pdf's pages came back read straight across both
    // columns, a line of one and a line of the other.
    const line = (tag: string, first: boolean): string =>
      `${tag} ${'set in the column and filling it '.repeat(3)}`.slice(0, first ? 42 : 44);
    const ops = ['BT /F1 9 Tf'];
    for (let k = 0; k < 30; k++) {
      const first = k % 5 === 0;
      const y = String(720 - k * 11);
      const n = String(k + 1).padStart(2, '0');
      // 44 characters of 9pt Courier fill the 237.6 points of each column.
      ops.push(`1 0 0 1 ${String(54 + (first ? 10.8 : 0))} ${y} Tm (${line(`L${n}`, first)}) Tj`);
      ops.push(`1 0 0 1 ${String(317 + (first ? 10.8 : 0))} ${y} Tm (${line(`R${n}`, first)}) Tj`);
    }
    ops.push('ET');
    const tokens =
      Ream.parse(onePage(ops, 'Courier'))
        .flow.body.map((el) =>
          el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '',
        )
        .join(' ')
        .match(/[LR]\d\d/gu) ?? [];
    const column = (letter: string): Array<string> =>
      Array.from({ length: 30 }, (_, i) => `${letter}${String(i + 1).padStart(2, '0')}`);
    expect(tokens).toEqual([...column('L'), ...column('R')]);
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

  it('reads an index down its columns, entries run on into the gutter and all', () => {
    // freeculture.pdf sets its index two columns to the page, six points apart
    // at their widest, under a heading centred over both. Entries that ran
    // past the middle of the white between them were read as lines across the
    // page — "167–Apple Corporation" one line — and the heading, standing
    // wholly right of that middle, as the head of the right column.
    const size = 8;
    const width = (text: string): number =>
      ([...text].reduce((w, c) => w + (standardWidth('Helvetica', c.charCodeAt(0), c) ?? 0), 0) *
        size) /
      1000;
    // An entry that runs on to within seven points of the right column.
    const long = (head: string): string => {
      let text = head;
      while (72 + width(`${text}1`) < 243) text += '1';
      return text;
    };
    const ops = ['BT /F1 15 Tf 1 0 0 1 215 740 Tm (INDEX) Tj', `/F1 ${String(size)} Tf`];
    for (let k = 0; k < 30; k++) {
      const y = 700 - k * 10;
      const n = String(k + 1).padStart(2, '0');
      const left =
        k % 6 === 3 || k >= 25 ? long(`L${n} an entry that runs on, `) : `L${n} entry, 12`;
      ops.push(`1 0 0 1 72 ${String(y)} Tm (${left}) Tj`);
      // The right column is the shorter: the last lines of the left one stand alone.
      const right = `R${n} an entry in the right column, ${String(34 + k)}${k % 3 === 0 ? ', 35' : ''}`;
      if (k < 24) ops.push(`1 0 0 1 250 ${String(y)} Tm (${right}) Tj`);
    }
    ops.push('ET');
    const text = Ream.parse(onePage(ops))
      .flow.body.map((el) =>
        el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '',
      )
      .join(' ');
    const column = (letter: string, rows: number): Array<string> =>
      Array.from({ length: rows }, (_, i) => `${letter}${String(i + 1).padStart(2, '0')}`);
    expect(text.match(/INDEX|[LR]\d\d/gu)).toEqual([
      'INDEX',
      ...column('L', 30),
      ...column('R', 24),
    ]);
  });

  it("sets a ragged column's gutter as the white its longest lines leave, not where most end", () => {
    // freeculture.pdf's index sets its entries ragged, and its gutter was
    // measured from where most of them end: set that far apart, its columns
    // came back a third narrower than the page's, and the thirteen pages of
    // the index ran to twenty-seven.
    const size = 8;
    const width = (text: string): number =>
      ([...text].reduce((w, c) => w + (standardWidth('Helvetica', c.charCodeAt(0), c) ?? 0), 0) *
        size) /
      1000;
    const long = (head: string): string => {
      let text = head;
      while (72 + width(`${text}1`) < 305) text += '1';
      return text;
    };
    const ops = [`BT /F1 ${String(size)} Tf`];
    for (let k = 0; k < 30; k++) {
      const y = 700 - k * 10;
      const n = String(k + 1).padStart(2, '0');
      const left =
        k % 5 === 2
          ? long(`L${n} an entry that runs on, `)
          : `L${n} ${'an entry '.repeat(1 + (k % 4))}12`;
      ops.push(`1 0 0 1 72 ${String(y)} Tm (${left}) Tj`);
      ops.push(
        `1 0 0 1 320 ${String(y)} Tm (R${n} an entry in the right-hand column, ${String(34 + k)}) Tj`,
      );
    }
    ops.push('ET');
    const longest = 72 + width(long('L03 an entry that runs on, '));
    const columns = Ream.parse(onePage(ops)).flow.sections.find(
      (section) => section.properties.columns,
    )?.properties.columns;
    expect(columns?.count).toBe(2);
    expect(columns?.spacePt).toBeCloseTo(320 - longest, 1);
  });

  it('measures a gutter past a line the page set over its measure', () => {
    // TeX leaves a line over its measure where it can break a paragraph no
    // better, and it stands out into the gutter alone: comments.pdf's, ten
    // points past the edge of its column, was taken for where the column
    // ends, and every justified line of it stopped short.
    const size = 8;
    const width = (text: string): number =>
      ([...text].reduce((w, c) => w + (standardWidth('Helvetica', c.charCodeAt(0), c) ?? 0), 0) *
        size) /
      1000;
    const line = (n: string): string => `L${n} ${'x'.repeat(48)}`;
    const ops = [`BT /F1 ${String(size)} Tf`];
    for (let k = 0; k < 30; k++) {
      const y = 700 - k * 10;
      const n = String(k + 1).padStart(2, '0');
      ops.push(`1 0 0 1 72 ${String(y)} Tm (${line(n)}${k === 12 ? 'xx' : ''}) Tj`);
      ops.push(
        `1 0 0 1 300 ${String(y)} Tm (R${n} an entry in the right-hand column, ${String(34 + k)}) Tj`,
      );
    }
    ops.push('ET');
    const columns = Ream.parse(onePage(ops)).flow.sections.find(
      (section) => section.properties.columns,
    )?.properties.columns;
    expect(columns?.count).toBe(2);
    expect(columns?.spacePt).toBeCloseTo(300 - (72 + width(line('01'))), 1);
  });

  it('reads a line with a word space over a point as running across it', () => {
    // A column of long lines spaces its words anywhere, and now and then a
    // few of them over the same point: freeculture.pdf's index spaced four
    // entries over one point seventy points short of its right column, and a
    // gutter was voted there — every entry longer than that was cut in two.
    const size = 8;
    const width = (text: string): number =>
      ([...text].reduce((w, c) => w + (standardWidth('Helvetica', c.charCodeAt(0), c) ?? 0), 0) *
        size) /
      1000;
    const ops = [`BT /F1 ${String(size)} Tf`];
    for (let k = 0; k < 30; k++) {
      const y = 700 - k * 10;
      const n = String(k + 1).padStart(2, '0');
      if (k < 14) ops.push(`1 0 0 1 72 ${String(y)} Tm (L${n} short, 12) Tj`);
      else if (k % 4 !== 3) {
        ops.push(
          `1 0 0 1 72 ${String(y)} Tm (L${n} a long entry, 12, 34, 56, 78, 90, 123, 456) Tj`,
        );
      } else {
        // The same entry set in two runs, the space between them stepped.
        const head = `L${n} a long entry, 12, 34,`;
        ops.push(`1 0 0 1 72 ${String(y)} Tm (${head}) Tj`);
        ops.push(
          `1 0 0 1 ${String(72 + width(head) + 2.2)} ${String(y)} Tm (56, 78, 90, 123, 456) Tj`,
        );
      }
      ops.push(
        `1 0 0 1 320 ${String(y)} Tm (R${n} an entry in the right-hand column, ${String(34 + k)}) Tj`,
      );
    }
    ops.push('ET');
    const text = Ream.parse(onePage(ops))
      .flow.body.map((el) =>
        el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '',
      )
      .join(' ');
    const column = (letter: string, rows: number): Array<string> =>
      Array.from({ length: rows }, (_, i) => `${letter}${String(i + 1).padStart(2, '0')}`);
    expect(text.match(/[LR]\d\d/gu)).toEqual([...column('L', 30), ...column('R', 30)]);
    expect(text).toContain('L16 a long entry, 12, 34, 56, 78, 90, 123, 456');
  });

  it('reads an index whose entries hang, a line apiece where the page ends them', () => {
    // freeculture.pdf sets "democracy:" at the edge, its sub-entries a level
    // in, and the lines that carry an entry on a level further in again.
    // Read by the measure alone, a line back at the edge after one carried
    // on was taken for the second line of a paragraph set in.
    const ops = ['BT /F1 8 Tf'];
    const lines: Array<[number, string]> = [
      [72, 'democracy:'],
      [82, 'digital sharing within, 184'],
      [82, 'media concentration and, 166'],
      [82, 'in technologies of expression, 33, 35, 36, 37, 38, 39, 40, 41,'],
      [92, '42, 43, 44-45'],
      [72, 'Democratic Party, 249'],
      [72, 'derivative works, 329n'],
      [82, 'historical shift in copyright coverage of, 136, 137, 138, 139,'],
      [92, '170-72'],
      [72, 'developing countries, foreign patent costs in, 63, 64, 65, 66,'],
      [92, '257-61, 313n'],
      [72, 'Diamond Multimedia Systems, 323n'],
    ];
    lines.forEach(([x, text], k) =>
      ops.push(`1 0 0 1 ${String(x)} ${String(700 - k * 10)} Tm (${text}) Tj`),
    );
    ops.push('ET');
    const paragraphs = Ream.parse(onePage(ops)).flow.body.flatMap((el) =>
      el.kind === 'paragraph' ? [el.paragraph.runs.map((r) => r.text).join('')] : [],
    );
    expect(paragraphs.map((p) => p.replace(/\s+/gu, ' '))).toEqual([
      'democracy:',
      'digital sharing within, 184',
      'media concentration and, 166',
      'in technologies of expression, 33, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44-45',
      'Democratic Party, 249',
      'derivative works, 329n',
      'historical shift in copyright coverage of, 136, 137, 138, 139, 170-72',
      'developing countries, foreign patent costs in, 63, 64, 65, 66, 257-61, 313n',
      'Diamond Multimedia Systems, 323n',
    ]);
  });

  it('reads a page of prose with a quotation set in as prose, not as an index', () => {
    // freeculture.pdf sets its quotations in from the edge, reached from a
    // full line of the paragraph they break, as an index's entries go on; read
    // as hanging, a page of its prose came back a paragraph to every line.
    const full = 'the argument runs on from one line to the next as prose does';
    const ops = ['BT /F1 10 Tf'];
    const lines: Array<[number, string]> = [
      ...Array.from({ length: 5 }, (): [number, string] => [72, full]),
      [72, 'and ends here.'],
      ...Array.from({ length: 4 }, (): [number, string] => [72, full]),
      [92, 'a quotation set in from the edge of the column, as long'],
      [92, 'as the lines of the prose are, or nearly so, runs on'],
      [92, 'and ends.'],
      ...Array.from({ length: 4 }, (): [number, string] => [72, full]),
      [72, 'the end.'],
    ];
    lines.forEach(([x, text], k) =>
      ops.push(`1 0 0 1 ${String(x)} ${String(700 - k * 12)} Tm (${text}) Tj`),
    );
    ops.push('ET');
    const paragraphs = Ream.parse(onePage(ops)).flow.body.flatMap((el) =>
      el.kind === 'paragraph' ? [el.paragraph.runs.map((r) => r.text).join('')] : [],
    );
    expect(paragraphs[0]?.replace(/\s+/gu, ' ')).toBe(
      `${Array(5).fill(full).join(' ')} and ends here.`,
    );
  });

  it('carries an entry on at the level its lines go on at, however short the line before', () => {
    // freeculture.pdf breaks "RPI, see Rensselaer Polytechnic" a third of
    // the column short, for want of room for "Institute", and carries it on a
    // level in: taken for an entry ended short, it came back two.
    const ops = ['BT /F1 8 Tf'];
    const lines: Array<[number, string]> = [
      [72, 'Rensselaer Polytechnic Institute (RPI), 48, 49, 50, 51, 185,'],
      [92, '200, 206'],
      [72, 'Rhapsody, 191'],
      [72, 'Rise of the Creative Class, The (Florida), 21'],
      [72, 'Roberts, Richard, 309n, 310n, 311n, 312n, 313n, 314n, 315n,'],
      [92, '316n'],
      [72, 'Rogers, Fred, 111'],
      [72, 'RPI, see Rensselaer Polytechnic'],
      [92, 'Institute'],
      [72, 'Rubin, Jed, 44'],
      [72, 'Russia, commercial piracy in, 63, 64, 65, 66, 67, 68, 69, 70,'],
      [92, '71, 302'],
      [72, 'Safire, William, 128, 129, 130, 131, 132, 133, 134, 135, 136,'],
      [92, '137, 138'],
    ];
    lines.forEach(([x, text], k) =>
      ops.push(`1 0 0 1 ${String(x)} ${String(700 - k * 10)} Tm (${text}) Tj`),
    );
    ops.push('ET');
    const paragraphs = Ream.parse(onePage(ops)).flow.body.flatMap((el) =>
      el.kind === 'paragraph' ? [el.paragraph.runs.map((r) => r.text).join('')] : [],
    );
    expect(paragraphs.map((p) => p.replace(/\s+/gu, ' '))).toEqual([
      'Rensselaer Polytechnic Institute (RPI), 48, 49, 50, 51, 185, 200, 206',
      'Rhapsody, 191',
      'Rise of the Creative Class, The (Florida), 21',
      'Roberts, Richard, 309n, 310n, 311n, 312n, 313n, 314n, 315n, 316n',
      'Rogers, Fred, 111',
      'RPI, see Rensselaer Polytechnic Institute',
      'Rubin, Jed, 44',
      'Russia, commercial piracy in, 63, 64, 65, 66, 67, 68, 69, 70, 71, 302',
      'Safire, William, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138',
    ]);
  });

  it('reads an invoice ACROSS, not down: figures against the right margin are not a column', () => {
    // A page whose amounts stand against the right margin breaks a dozen lines
    // at the same x, which is what a gutter looks like from the outside and
    // nothing like two columns from the inside. Read down, every amount was
    // torn off the label it belongs to and carried to the end of the document
    // — a receipt came back with "Total excluding tax" on one sheet and
    // "$100.00" on the next — and the labels were then indented past the strip
    // they had been given, one letter per line.
    const ops = ['BT /F1 10 Tf'];
    const amounts = ['$100.00', '$16.00', '$116.00', '$4.00', '$8.00', '$12.00'];
    for (let i = 0; i < ROWS; i++) {
      const y = 720 - i * 24;
      const n = String(i + 1).padStart(2, '0');
      ops.push(`1 0 0 1 72 ${String(y)} Tm (LABEL ${n} of the account) Tj`);
      // Flush RIGHT: each figure ends at the same place and starts wherever its
      // own width puts it.
      const amount = amounts[i % amounts.length]!;
      const x = 540 - amount.length * 5;
      ops.push(`1 0 0 1 ${String(x)} ${String(y)} Tm (${amount}) Tj`);
    }
    ops.push('ET');
    const tokens = Ream.parse(onePage(ops))
      .flow.body.map((el) =>
        el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '',
      )
      .filter((line) => line.includes('LABEL'));
    // Every label keeps its figure on its own line…
    expect(tokens).toHaveLength(ROWS);
    expect(tokens[0]).toContain('$100.00');
    // …and the section is not set in columns.
    const columns = Ream.parse(onePage(ops)).flow.sections[0]?.properties.columns;
    expect(columns?.count ?? 1).toBe(1);
  });

  it('measures the left margin to INK, not to a space set in the corner', () => {
    // A Stripe invoice opens its page with one non-breaking space at x=0.
    // Taken for the leftmost thing on the sheet it put the left margin at zero
    // and moved every line of the document flush against the edge, thirty
    // points left of where the page sets them.
    const ops = ['BT /F1 10 Tf', '1 0 0 1 0 779 Tm ( ) Tj'];
    for (let i = 0; i < ROWS; i++)
      ops.push(`1 0 0 1 72 ${String(720 - i * 24)} Tm (a line of the document here) Tj`);
    ops.push('ET');
    const section = reconstructByLayout(PdfFile.parse(onePage(ops))).doc.section;
    expect(section?.margins?.left as number).toBeGreaterThan(60);
  });
});
