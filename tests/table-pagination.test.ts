import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

function layoutOf(docx: Uint8Array) {
  const flow = Ream.parse(docx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
}

/** The text of every line drawn, page by page. */
function textByPage(laid: ReturnType<typeof layoutOf>): Array<Array<string>> {
  return laid.pages.map((p) =>
    p.commands
      .filter((c) => c.type === 'line')
      .map((c) =>
        (c as unknown as { line: { tokens: ReadonlyArray<{ text?: string }> } }).line.tokens
          .map((t) => t.text ?? '')
          .join(''),
      )
      .filter((t) => t !== ''),
  );
}

// Half a page of body text, then a one-cell row that needs most of a page.
const FILLER = Array.from(
  { length: 24 },
  (_, i) => `<w:p><w:r><w:t>Filler${String(i)}</w:t></w:r></w:p>`,
).join('');
const CELL = Array.from(
  { length: 30 },
  (_, i) => `<w:p><w:r><w:t>Row${String(i)}</w:t></w:r></w:p>`,
).join('');

const doc = (trPr = '') =>
  buildDocxFromBody(
    `${FILLER}<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>` +
      `<w:tr>${trPr}<w:tc>${CELL}</w:tc></w:tr></w:tbl>`,
  );

describe('a table row at the page edge (§17.4.6)', () => {
  it('breaks the row where the page ends', () => {
    // table-row-data-displayed-twice.docx is one 713pt row: drawn whole it ran
    // off the bottom of the page and over the section that follows it.
    const pages = textByPage(layoutOf(doc()));
    expect(pages.length).toBe(2);
    expect(pages[0]!).toContain('Row0'); // the row starts beside the filler…
    expect(pages[1]!).toContain('Row29'); // …and ends on the next page
    // Nothing is lost or repeated: every line of the cell is drawn once.
    const all = pages.flat().filter((t) => t.startsWith('Row'));
    expect(all.length).toBe(30);
    expect(new Set(all).size).toBe(30);
  });

  it('moves a `w:cantSplit` row whole instead', () => {
    const pages = textByPage(layoutOf(doc('<w:trPr><w:cantSplit/></w:trPr>')));
    expect(pages[0]!.some((t) => t.startsWith('Row'))).toBe(false);
    expect(pages[1]!).toContain('Row0');
  });

  it('counts the space between a cell’s paragraphs when it splits one', () => {
    // The chunk sizer measured lines only, so a cell of spaced paragraphs
    // reported less height than it drew — and the piece it cut off overflowed
    // the page it was cut to fit.
    const spaced = Array.from(
      { length: 30 },
      (_, i) =>
        `<w:p><w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr>` +
        `<w:r><w:t>Row${String(i)}</w:t></w:r></w:p>`,
    ).join('');
    const laid = layoutOf(
      buildDocxFromBody(
        `${FILLER}<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>` +
          `<w:tr><w:tc>${spaced}</w:tc></w:tr></w:tbl>`,
      ),
    );
    // Letter minus 1" margins: nothing may be drawn below the bottom margin.
    for (const page of laid.pages) {
      for (const c of page.commands) {
        if (c.type !== 'line') continue;
        expect((c as unknown as { baselineY: number }).baselineY).toBeLessThanOrEqual(721);
      }
    }
  });
});

// Issue #139 and its neighbours: rows of a long table, each with a height rule,
// laid out on A4 with 2 cm margins (728.5pt of body) — the probes measured in
// Word for Mac. Every cell paragraph sets its own exact line pitch, so the rows'
// heights do not hang on the test font's metrics.
const A4 =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"' +
  ' w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>';

/** A paragraph `pitch` twips tall per line, with `after` twips of space after. */
const para = (text: string, pitch = 300, after = 0): string =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="${String(after)}" w:line="${String(pitch)}"` +
  ` w:lineRule="exact"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/**
 * A two-column table: `R<i>` on the left, `right(i)` on the right, every row
 * with `trPr`, and `tblPr` as the table's properties.
 */
function tableDoc(
  rows: number,
  trPr: string,
  right: (i: number) => string,
  tblPr = '',
): Uint8Array {
  let body = `<w:tbl><w:tblPr>${tblPr}</w:tblPr>`;
  body += '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="7638"/></w:tblGrid>';
  for (let i = 1; i <= rows; i++) {
    body += `<w:tr><w:trPr>${trPr}</w:trPr><w:tc>${para(`R${String(i)}`)}</w:tc><w:tc>${right(i)}</w:tc></w:tr>`;
  }
  return buildDocxFromBody(`${body}</w:tbl>${para('')}${A4}`);
}

/** Every drawn line's text and its baseline (down from the page's top), page by page. */
function linesByPage(laid: ReturnType<typeof layoutOf>): Array<Array<{ text: string; y: number }>> {
  return laid.pages.map((p) =>
    p.commands
      .filter((c) => c.type === 'line')
      .map((c) => {
        const l = c as unknown as {
          line: { tokens: ReadonlyArray<{ text?: string }> };
          baselineY: number;
        };
        return { text: l.line.tokens.map((t) => t.text ?? '').join(''), y: l.baselineY };
      })
      .filter((l) => l.text !== ''),
  );
}

/** The baselines of the left-hand `R<i>` labels, page by page. */
function labelsByPage(laid: ReturnType<typeof layoutOf>): Array<Array<{ row: number; y: number }>> {
  return linesByPage(laid).map((page) =>
    page
      .filter((l) => /^R\d+$/.test(l.text))
      .map((l) => ({ row: Number(l.text.slice(1)), y: l.y })),
  );
}

/** Row pitches on each page, from one label's baseline to the next. */
const pitches = (page: ReadonlyArray<{ y: number }>): Array<number> =>
  page.slice(1).map((l, i) => l.y - page[i]!.y);

const EXACT = '<w:trHeight w:val="680" w:hRule="exact"/>';
const AT_LEAST = '<w:trHeight w:val="680" w:hRule="atLeast"/>';

describe('a row with a height rule at a page break (§17.4.81, issue #139)', () => {
  it('moves an exact row to the next page whole, at its height', () => {
    // Word: 21 rows of 34pt a page (21 × 34 = 714; the 22nd has 14.5pt and
    // moves), every one of them 34pt. The row carried over came out 15.24pt —
    // its line alone — and two more were squeezed in at the foot of page 2.
    const pages = labelsByPage(layoutOf(tableDoc(50, EXACT, (i) => para(`Row ${String(i)}`))));
    expect(pages.map((p) => p.length)).toEqual([21, 21, 8]);
    for (const page of pages) {
      for (const p of pitches(page)) expect(p).toBeCloseTo(34, 2);
      // …and each page's first row stands where the table's first does.
      expect(page[0]!.y).toBeCloseTo(pages[0]![0]!.y, 2);
    }
  });

  it('never breaks an exact row, even one whose text is taller than it', () => {
    // Three 15pt lines in a 34pt row: Word moves it whole like any other.
    const three = (i: number) =>
      para(`Row ${String(i)} a`) + para(`Row ${String(i)} b`) + para(`Row ${String(i)} c`);
    const laid = layoutOf(tableDoc(50, EXACT, three));
    const pages = labelsByPage(laid);
    expect(pages.map((p) => p.length)).toEqual([21, 21, 8]);
    for (const page of pages) for (const p of pitches(page)) expect(p).toBeCloseTo(34, 2);
    // What a row draws stays on the page its label is on.
    linesByPage(laid).forEach((page, k) => {
      for (const l of page) {
        const row = /^Row (\d+) /.exec(l.text);
        if (row) expect(pages[k]!.some((p) => p.row === Number(row[1]))).toBe(true);
      }
    });
  });

  it('moves a one-line atLeast row whole, at its minimum', () => {
    const pages = labelsByPage(layoutOf(tableDoc(50, AT_LEAST, (i) => para(`Row ${String(i)}`))));
    expect(pages.map((p) => p.length)).toEqual([21, 21, 8]);
    for (const page of pages) for (const p of pitches(page)) expect(p).toBeCloseTo(34, 2);
  });

  it('moves an atLeast row whole where its minimum does not fit', () => {
    // Word: a 100pt-minimum row of six 22.5pt lines with 52pt left — room
    // for two of them — moves to the next page whole. Here: six 19pt lines
    // (114pt) a row, six rows a page, 44.5pt left.
    const six = (i: number) =>
      Array.from({ length: 6 }, (_, k) => para(`Row ${String(i)} line ${String(k + 1)}`, 380)).join(
        '',
      );
    const laid = layoutOf(tableDoc(14, '<w:trHeight w:val="2000" w:hRule="atLeast"/>', six));
    const pages = linesByPage(laid);
    expect(labelsByPage(laid).map((p) => p.length)).toEqual([6, 6, 2]);
    for (let i = 1; i <= 14; i++) {
      const on = pages.findIndex((p) => p.some((l) => l.text === `Row ${String(i)} line 1`));
      expect(pages[on]!.some((l) => l.text === `Row ${String(i)} line 6`)).toBe(true);
    }
  });

  it('gives the part of an atLeast row a break carries over its minimum', () => {
    // Word: a 34pt-minimum row of three lines broken after two ends 34.8pt
    // below the top margin, where its third line alone would end at 23. Here:
    // three 19pt lines (57pt) a row, twelve rows a page, 44.5pt left — two
    // lines stay, the third goes over and stands in a 34pt part.
    const three = (i: number) =>
      ['a', 'b', 'c'].map((k) => para(`Row ${String(i)} ${k}`, 380)).join('');
    const laid = layoutOf(tableDoc(30, AT_LEAST, three));
    const pages = linesByPage(laid);
    const find = (text: string) => {
      const k = pages.findIndex((p) => p.some((l) => l.text === text));
      return { page: k, y: pages[k]!.find((l) => l.text === text)!.y };
    };
    expect(find('Row 13 a').page).toBe(0);
    expect(find('Row 13 b').page).toBe(0);
    const carried = find('Row 13 c');
    expect(carried.page).toBe(1);
    // At the top, where a row's first line stands…
    expect(carried.y).toBeCloseTo(find('Row 1 a').y, 2);
    // …and the next row 34pt below it, not 19.
    expect(find('Row 14 a').y - carried.y).toBeCloseTo(34, 2);
  });

  it('moves a one-line row whole with its space after (no height rule)', () => {
    // Word: 23pt rows (a line and 8pt after), 31 on the first page and the
    // 32nd moved: its line fits in the 15.5pt left, its space after does not.
    // It was squeezed in as 15pt instead, and the next one carried over as 15.
    const laid = layoutOf(tableDoc(50, '', (i) => para(`Row ${String(i)}`, 300, 160)));
    const pages = labelsByPage(laid);
    expect(pages.map((p) => p.length)).toEqual([31, 19]);
    for (const page of pages) for (const p of pitches(page)) expect(p).toBeCloseTo(23, 2);
  });

  it('starts a carried-over part at the top and keeps its last space after', () => {
    // Word, a row of three spaced paragraphs broken after two: the third line
    // at the top margin (the gap above it goes, as above a paragraph a break
    // carries over) and the next row a line plus its 8pt after below it.
    // Here: four 19pt lines with 8pt after (108pt a row), six rows a page and
    // 80.5pt left: three lines stay, the fourth goes over.
    const four = (i: number) =>
      ['a', 'b', 'c', 'd'].map((k) => para(`Row ${String(i)} ${k}`, 380, 160)).join('');
    const laid = layoutOf(tableDoc(12, '', four));
    const pages = linesByPage(laid);
    const find = (text: string) => {
      const k = pages.findIndex((p) => p.some((l) => l.text === text));
      return { page: k, y: pages[k]!.find((l) => l.text === text)!.y };
    };
    expect(find('Row 7 c').page).toBe(0);
    const carried = find('Row 7 d');
    expect(carried.page).toBe(1);
    expect(carried.y).toBeCloseTo(find('Row 1 a').y, 2);
    expect(find('Row 8 a').y - carried.y).toBeCloseTo(27, 2);
  });

  it('gives an exact row taller than the page a page of its own, cut at the margin', () => {
    // Word: an 800pt exact row after a 34pt one stands alone on page 2 — the
    // 728.5pt it has — and the next row starts page 3. It used to be broken
    // over two pages, 897pt of it, every line drawn.
    let body = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="7638"/></w:tblGrid>';
    const row = (trPr: string, label: string, right: string) =>
      `<w:tr><w:trPr>${trPr}</w:trPr><w:tc>${para(label)}</w:tc><w:tc>${right}</w:tc></w:tr>`;
    body += row(EXACT, 'R1', para('First'));
    const forty = Array.from({ length: 40 }, (_, k) => para(`Huge ${String(k + 1)}`, 300, 160));
    body += row('<w:trHeight w:val="16000" w:hRule="exact"/>', 'R2', forty.join(''));
    for (let i = 3; i <= 5; i++) body += row(EXACT, `R${String(i)}`, para(`Row ${String(i)}`));
    const pages = linesByPage(layoutOf(buildDocxFromBody(`${body}</w:tbl>${para('')}${A4}`)));
    expect(pages.length).toBe(3);
    expect(pages[0]!.map((l) => l.text)).toEqual(['R1', 'First']);
    const huge = pages[1]!.filter((l) => l.text.startsWith('Huge'));
    expect(huge[0]!.text).toBe('Huge 1');
    // 23pt a line from 56.7pt down: line 32 is the last to start above the
    // bottom margin (785.2pt), and nothing past it is drawn.
    expect(huge.length).toBe(32);
    expect(pages[2]!.map((l) => l.text)).toEqual(['R3', 'Row 3', 'R4', 'Row 4', 'R5', 'Row 5']);
  });
});

describe('a row height under borders (§17.4.81)', () => {
  // The probes measured in Word for Mac: rows of one line, rules on every edge
  // or on the cells' bottoms only.
  const rule = (e: string, sz: number) =>
    `<w:${e} w:val="single" w:sz="${String(sz)}" w:space="0" w:color="000000"/>`;
  const ruled = (sz: number) =>
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((e) => rule(e, sz)).join('') +
    '</w:tblBorders>';

  /** `rows` rows, each cell with `tcPr`, then a paragraph "After". */
  function ruledDoc(rows: number, trPr: string, tblPr: string, tcPr = ''): Uint8Array {
    let body = `<w:tbl><w:tblPr>${tblPr}</w:tblPr>`;
    body += '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="7638"/></w:tblGrid>';
    for (let i = 1; i <= rows; i++) {
      body +=
        `<w:tr><w:trPr>${trPr}</w:trPr>` +
        `<w:tc><w:tcPr>${tcPr}</w:tcPr>${para(`R${String(i)}`)}</w:tc>` +
        `<w:tc><w:tcPr>${tcPr}</w:tcPr>${para(`Row ${String(i)}`)}</w:tc></w:tr>`;
    }
    return buildDocxFromBody(`${body}</w:tbl>${para('After')}${A4}`);
  }

  /** Inner row pitches, and how far the paragraph after stands below the last row's text. */
  function measure(docx: Uint8Array): { inner: Array<number>; after: number } {
    const page = linesByPage(layoutOf(docx))[0]!;
    const labels = page.filter((l) => /^R\d+$/.test(l.text));
    const after = page.find((l) => l.text === 'After')!;
    return { inner: pitches(labels).slice(1), after: after.y - labels[labels.length - 1]!.y };
  }

  it('puts the rule above a row on top of its atLeast minimum', () => {
    // Word: 34pt minimum, one line a row — 37.00pt apart under 3pt rules
    // (compatibility modes 15, 14 and none), 34.5 under 0.5pt ones.
    for (const [sz, pitch] of [
      [24, 37],
      [4, 34.5],
    ] as const) {
      const { inner } = measure(ruledDoc(10, AT_LEAST, ruled(sz)));
      for (const p of inner) expect(p).toBeCloseTo(pitch, 2);
    }
  });

  it('counts a rule that is only the bottom of the cell above', () => {
    // tdf127814.docx's rows: no top border of their own, the rule between
    // them the upper cell's bottom. Word: 37.00pt apart still (LibreOffice
    // gives 34).
    const tcPr =
      '<w:tcBorders><w:top w:val="nil"/>' +
      rule('left', 24) +
      rule('bottom', 24) +
      rule('right', 24) +
      '</w:tcBorders>';
    const { inner } = measure(ruledDoc(10, AT_LEAST, '', tcPr));
    for (const p of inner) expect(p).toBeCloseTo(37, 2);
  });

  it('keeps the rule above a row inside an exact height', () => {
    // Word: 33.98pt apart under 3pt rules, and 34pt under 0.5pt ones.
    for (const sz of [24, 4]) {
      const { inner } = measure(ruledDoc(10, EXACT, ruled(sz)));
      for (const p of inner) expect(p).toBeCloseTo(34, 2);
    }
  });

  it('puts the table’s bottom edge under its last row, whatever the rule', () => {
    // Word, 3pt rules: the paragraph after stands 34.05pt below the text of a
    // last exact row and 36.4 below an atLeast one's, the text 3.2pt down
    // from the row's top — rows of 37 and 40.
    expect(measure(ruledDoc(10, EXACT, ruled(24))).after).toBeCloseTo(34, 2);
    expect(measure(ruledDoc(10, AT_LEAST, ruled(24))).after).toBeCloseTo(37, 2);
  });
});
