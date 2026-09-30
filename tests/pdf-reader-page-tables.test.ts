// A table set across the head of a page in columns is read as the table it is,
// and the prose under it in its columns: the white between a table's columns
// is no gutter of the page's. comments.pdf's Figure 13 is a benchmark and nine
// figures a row, twenty-six rows deep, over two columns of text, and the whole
// page came back as one table — each line of the text a row of it.

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import type { BodyElement } from '@/core/document-model';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { standardWidth } from '@/pdf-reader/standard-widths';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';

/** Letter pages of `ops` in Helvetica, one page each. */
function pages(sheets: ReadonlyArray<ReadonlyArray<string>>): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: sheets.length });
  const pagesRef = doc.add(pagesMap);
  for (const ops of sheets) {
    const content = doc.add(stream({}, new TextEncoder().encode(ops.join('\n'))));
    const leaf = doc.add(
      dict({
        Type: name('Page'),
        Parent: pagesRef,
        MediaBox: [0, 0, 612, 792],
        Resources: dict({ Font: dict({ F1: font }) }),
        Contents: content,
      }),
    );
    (pagesMap.get('Kids') as Array<PdfValue>).push(leaf);
  }
  const catalog = doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }));
  return doc.build(catalog);
}

const at = (x: number, y: number, text: string): string =>
  `1 0 0 1 ${String(x)} ${String(y)} Tm (${text}) Tj`;

/** How wide `text` is set in 9pt Helvetica. */
const widthOf = (text: string): number =>
  [...text].reduce((sum, c) => sum + (standardWidth('Helvetica', c.charCodeAt(0), c) ?? 0), 0) *
  0.009;

/** `text` set to end at `right`. */
const endingAt = (right: number, y: number, text: string): string =>
  at(right - widthOf(text), y, text);

/** The table's columns: a name, then figures whose right edges line up. */
const COLUMNS = [150, 190, 230, 270, 310, 350, 390, 430];

/**
 * A table of `rows` rows from `top`, a heading row first whose words stand
 * only an em and a half apart — narrower than a tab's white.
 */
function table(top: number, rows: number): Array<string> {
  const ops = COLUMNS.map((right, k) => endingAt(right, top, `Head${String(k)}`));
  for (let k = 1; k <= rows; k++) {
    const y = top - k * 10;
    ops.push(at(70, y, `bench-${String(k)}`));
    COLUMNS.forEach((right, c) => ops.push(endingAt(right, y, `${String((k * c) % 90)}`)));
  }
  return ops;
}

/** Two columns of prose from `top`: `L..` at the left, `R..` at the right. */
function columns(top: number, rows: number, tag = ''): Array<string> {
  const ops: Array<string> = [];
  for (let k = 0; k < rows; k++) {
    const n = String(k + 1).padStart(2, '0');
    const y = top - k * 12;
    ops.push(at(54, y, `L${tag}${n} a line of prose that runs across its column`));
    ops.push(at(317, y, `R${tag}${n} a line of prose that runs across its column`));
  }
  return ops;
}

const textOf = (el: BodyElement): string =>
  el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : '';

describe('a table across the head of a page in columns', () => {
  const sheet = ['BT /F1 9 Tf', ...table(760, 28), ...columns(440, 30), 'ET'];
  const { doc } = reconstructByLayout(PdfFile.parse(pages([sheet])));

  it('reads the prose under it in its columns', () => {
    const prose =
      doc.body
        .map(textOf)
        .join(' ')
        .match(/[LR]\d\d/gu) ?? [];
    const column = (letter: string): Array<string> =>
      Array.from({ length: 30 }, (_, k) => `${letter}${String(k + 1).padStart(2, '0')}`);
    expect(prose).toEqual([...column('L'), ...column('R')]);
    expect(doc.sections.some((s) => s.properties.columns?.count === 2)).toBe(true);
  });

  it('reads the table as one table, its heading its first row', () => {
    const tables = doc.body.filter((el) => el.kind === 'table');
    expect(tables).toHaveLength(1);
    const [t] = tables;
    if (t?.kind !== 'table') throw new Error('a table');
    expect(t.table.rows).toHaveLength(29);
    const cell = (r: number, c: number): string =>
      (t.table.rows[r]?.cells[c]?.content ?? []).map(textOf).join('');
    expect(cell(0, 1)).toBe('Head0');
    expect(cell(1, 0)).toBe('bench-1');
    expect(t.table.rows.every((row) => row.cells.length === COLUMNS.length + 1)).toBe(true);
  });
});

describe('a table of figures as TeX sets one', () => {
  // comments.pdf's Figure 13: a name and figures set against the right of
  // their columns, headings over them wider than they are, a dash for a
  // figure there is none of, and a rule under the headings — the table
  // centred on the page, its names set in from the margin.
  const RIGHTS = [200, 260, 330, 400, 470, 540];
  const HEADS = ['Loops', 'Flushes', 'Trees/Loop', 'Traces/Tree', 'Traces/Loop', 'Speedup'];
  const figures = (k: number): Array<string> =>
    k === 3
      ? ['12', '0', '-', '-', '-', '0.93x']
      : [String(k * 7), '0', `${String(k)}.5`, '1.0', `${String(k % 3)}.2`, `${String(k)}.20x`];
  const ops = ['BT /F1 9 Tf'];
  RIGHTS.forEach((right, c) => ops.push(endingAt(right, 700, HEADS[c]!)));
  for (let k = 1; k <= 12; k++) {
    const y = 700 - 3 - k * 10;
    ops.push(at(72, y, `bench-${String(k)}`));
    figures(k).forEach((figure, c) => ops.push(endingAt(RIGHTS[c]!, y, figure)));
  }
  ops.push(...columns(520, 30), 'ET', '0.4 w 66 696.8 m 545 696.8 l S');
  const { doc } = reconstructByLayout(PdfFile.parse(pages([ops])));
  const tables = doc.body.filter((el) => el.kind === 'table');
  const [t] = tables;
  if (t?.kind !== 'table') throw new Error('a table');
  const cell = (r: number, c: number) => t.table.rows[r]!.cells[c]!;
  const textAt = (r: number, c: number): string => cell(r, c).content.map(textOf).join('');

  it('is one table, a dash a cell', () => {
    expect(tables).toHaveLength(1);
    expect(t.table.rows).toHaveLength(13);
    expect([3, 4, 5].map((c) => textAt(3, c))).toEqual(['-', '-', '-']);
  });

  it('sets each heading in one piece, in a column that holds it', () => {
    // Cut to the figures under it, "Traces/Loop" came back "Traces/" over "Loop".
    expect(HEADS.map((_, c) => textAt(0, c + 1))).toEqual(HEADS);
    HEADS.forEach((head, c) => {
      const para = cell(0, c + 1).content[0];
      const inset = para?.kind === 'paragraph' ? (para.paragraph.properties.indentRight ?? 0) : 0;
      expect(t.table.grid[c + 1]! - inset).toBeGreaterThan(widthOf(head));
    });
  });

  it('sets figures against the right of their column, noughts and all', () => {
    // Every figure under "Flushes" is one nought, as wide as the next: only
    // the heading over them says which side they are set against.
    for (let c = 1; c <= HEADS.length; c++) {
      const alignment = (r: number): string | undefined => {
        const para = cell(r, c).content[0];
        return para?.kind === 'paragraph' ? para.paragraph.properties.alignment : undefined;
      };
      expect([0, 1, 5].map(alignment)).toEqual(['right', 'right', 'right']);
    }
  });

  it('stands in from the margin as far as the page set its names', () => {
    expect(t.table.properties.indentPt).toBeCloseTo(72 - 54, 0);
  });

  it('rules its headings off with a border, and draws the rule no more', () => {
    expect(t.table.rows[1]!.cells.every((c) => c.properties.borders?.top !== undefined)).toBe(true);
    expect(doc.body.some((el) => el.kind === 'shape')).toBe(false);
  });
});

describe('a rule in the white between two rows of a table', () => {
  it('stays where the page drew it', () => {
    // An invoice's rule under its headings stands in the white between them
    // and the item under them: moved onto the item's edge it rose nine points.
    const RIGHTS = [380, 450, 500, 558];
    const row = (y: number, cells: ReadonlyArray<string>): Array<string> => [
      at(54, y, cells[0]!),
      ...RIGHTS.map((right, c) => endingAt(right, y, cells[c + 1]!)),
    ];
    const ops = [
      'BT /F1 9 Tf',
      ...row(500, ['Description', 'Qty', 'Unit price', 'Tax', 'Amount']),
      ...row(478, ['Max plan', '1', '$100.00', '16%', '$100.00']),
      'ET',
      '0.75 w 54 491.5 m 558 491.5 l S',
    ];
    const { doc } = reconstructByLayout(PdfFile.parse(pages([ops])));
    const t = doc.body.find((el) => el.kind === 'table');
    if (t?.kind !== 'table') throw new Error('a table');
    const ruled = t.table.rows.flatMap((r) => r.cells).filter((c) => c.properties.borders);
    expect(ruled).toHaveLength(0);
    expect(doc.body.some((el) => el.kind === 'shape')).toBe(true);
  });
});

describe('a line over a table its columns cannot hold', () => {
  it('stays a line of its own', () => {
    // A form's foot: "Bankverbindung:" heads a label and its value together,
    // wider than the labels under it, and set over the labels alone it broke
    // in two.
    const ops = ['BT /F1 9 Tf', at(54, 300, 'Internet:'), at(200, 300, 'Adresse:')];
    ops.push(at(340, 300, 'Bankverbindung:'));
    const rows = [
      ['www:', 'Street 5', 'Bank:', 'Sparkasse'],
      ['e-mail:', '5020 City', 'IBAN:', 'AT92 2040'],
      ['irc:', 'Country', 'BIC:', 'SBGSAT2S'],
    ];
    rows.forEach((row, k) => {
      const y = 286 - k * 12;
      [54, 200, 340, 395].forEach((x, c) => ops.push(at(x, y, row[c]!)));
    });
    ops.push('ET');
    const { doc } = reconstructByLayout(PdfFile.parse(pages([ops])));
    const t = doc.body.find((el) => el.kind === 'table');
    if (t?.kind !== 'table') throw new Error('a table');
    expect(t.table.rows).toHaveLength(3);
    expect(doc.body.map(textOf)).toContain('Internet:\tAdresse:\tBankverbindung:');
  });
});

describe('a table down either column, their rows abreast', () => {
  it('reads each in its own column', () => {
    // canvas.pdf sets a table down either side of the sheet, three rows
    // abreast, and read as one table across the page the two came back in a
    // row: "void | setTransform( | lineWidth | float | 1.0".
    const ops = ['BT /F1 9 Tf', ...columns(740, 10)];
    const left = [
      ['Return', 'Name'],
      ['void', 'save()'],
      ['void', 'restore()'],
    ];
    const right = [
      ['Name', 'Type', 'Default'],
      ['lineWidth', 'float', '1.0'],
      ['lineCap', 'string', 'butt'],
    ];
    for (let k = 0; k < 3; k++) {
      const y = 610 - k * 12;
      ops.push(at(54, y, left[k]![0]!), at(130, y, left[k]![1]!));
      [317, 400, 480].forEach((x, c) => ops.push(at(x, y, right[k]![c]!)));
    }
    ops.push(...columns(574, 10, 'x'), 'ET');
    const { doc } = reconstructByLayout(PdfFile.parse(pages([ops])));
    // The words of a table too, cell by cell.
    const words = (el: BodyElement): string =>
      el.kind === 'table'
        ? el.table.rows.flatMap((r) => r.cells.flatMap((c) => c.content.map(words))).join(' ')
        : textOf(el);
    const text = doc.body.map(words).join(' ');
    const order = ['restore', 'Lx01', 'R01', 'lineWidth', 'Rx01'].map((w) => text.indexOf(w));
    expect(order.every((at) => at >= 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
  });
});

describe('the white under a line across the page', () => {
  it('is the space after it, over both columns under it', () => {
    // comments.pdf leaves twenty points under Figure 13's caption before the
    // columns begin, and each column read from nothing put the white nowhere:
    // the caption came down onto the text. As space before each column's
    // first paragraph, Word and LibreOffice drop it at the head of the second.
    const caption = 'Figure 13. A caption set across the page, wider than a column is';
    const ops = ['BT /F1 9 Tf', at(54, 600, caption), ...columns(575, 20), 'ET'];
    const { doc } = reconstructByLayout(PdfFile.parse(pages([ops])));
    const paragraphs = doc.body.flatMap((el) => (el.kind === 'paragraph' ? [el.paragraph] : []));
    const over = paragraphs.find((p) => p.runs.some((r) => r.text.startsWith('Figure 13.')));
    // Baselines 25 points apart, less the boxes the two lines stand in.
    expect(over?.properties.spacingAfter).toBeGreaterThan(10);
    const under = paragraphs.filter((p) => p.runs.some((r) => /^[LR]01/u.test(r.text)));
    expect(under).toHaveLength(2);
    expect(under.map((p) => p.properties.spacingBefore ?? 0)).toEqual([0, 0]);
  });
});

describe('a table inside a column', () => {
  it('stays in its column, and the column beside it reads on', () => {
    const ops = ['BT /F1 9 Tf'];
    for (let k = 0; k < 40; k++) {
      ops.push(
        at(
          317,
          740 - k * 12,
          `R${String(k + 1).padStart(2, '0')} a line of prose in the right column`,
        ),
      );
    }
    // Four narrow columns of figures in the left column, and prose over them.
    for (let k = 0; k < 10; k++) {
      ops.push(
        at(
          54,
          740 - k * 12,
          `P${String(k + 1).padStart(2, '0')} a line of prose in the left column`,
        ),
      );
    }
    for (let k = 0; k < 12; k++) {
      const y = 600 - k * 10;
      ops.push(at(54, y, `row${String(k)}`), at(120, y, '12'), at(170, y, '34'), at(220, y, '56'));
    }
    ops.push('ET');
    const { doc } = reconstructByLayout(PdfFile.parse(pages([ops])));
    const right = doc.body.map(textOf).join(' ').match(/R\d\d/gu) ?? [];
    expect(right).toEqual(
      Array.from({ length: 40 }, (_, k) => `R${String(k + 1).padStart(2, '0')}`),
    );
  });
});
