// A Word table's column widths when the cells' `w:tcW` and the `w:tblGrid`
// disagree: Word 16.114's layout of our own probes (Word's PDF, read back),
// in compatibility modes 15, 14 and none alike. The probes stand on a 523.3pt
// text column (A4, half-inch margins) with Word's default 5.4pt cell margins.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { Table, TableCell } from '@/core/document-model';
import type { TextLineItem } from '@/layout/page-doc';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { FontRegistry } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { wordColumnWidths } from '@/layout/word-table-columns';
import { parseDocument } from '@/word';

const LINE = 523.3;
const PAD = 5.4;

interface Probe {
  /** `w:tblPr` children. */
  readonly tblPr?: string;
  /** `w:gridCol` widths in twips. */
  readonly grid: ReadonlyArray<number>;
  /** Each row: each cell's `w:tcPr` children and its text. */
  readonly rows: ReadonlyArray<ReadonlyArray<readonly [tcPr: string, text: string]>>;
  /** Each row's `w:gridBefore`, where it has one. */
  readonly gridBefore?: ReadonlyArray<number>;
}

const FIXED = '<w:tblLayout w:type="fixed"/>';
const tblW = (type: string, w: number | string): string => `<w:tblW w:w="${w}" w:type="${type}"/>`;
const tcW = (type: string, w: number | string): string => `<w:tcW w:w="${w}" w:type="${type}"/>`;
const AUTO = tblW('auto', 0);
const FULL = tblW('pct', 5000);
const IN1 = tcW('dxa', 1440);
const IN3 = tcW('dxa', 4320);

function xml(p: Probe): string {
  const grid = p.grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('');
  const rows = p.rows
    .map(
      (r, k) =>
        '<w:tr>' +
        (p.gridBefore?.[k]
          ? `<w:trPr><w:gridBefore w:val="${String(p.gridBefore[k])}"/></w:trPr>`
          : '') +
        r
          .map(
            ([tcPr, text]) =>
              `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`,
          )
          .join('') +
        '</w:tr>',
    )
    .join('');
  return `<w:tbl><w:tblPr>${p.tblPr ?? ''}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

function parse(p: Probe): Table {
  const el = parseDocument(OpcPackage.open(buildDocxFromBody(xml(p))).getMainDocument().data).find(
    (b) => b.kind === 'table',
  );
  if (el?.kind !== 'table') throw new Error('expected a table');
  return el.table;
}

/**
 * Column widths with text measured at 5pt a character: a cell's narrowest
 * setting is its longest word, its widest the whole text on one line.
 */
function widths(p: Probe): Array<number> | undefined {
  const table = parse(p);
  const textOf = (cell: TableCell): string =>
    cell.content
      .map((el) => (el.kind === 'paragraph' ? el.paragraph.runs.map((r) => r.text).join('') : ''))
      .join(' ');
  return wordColumnWidths(table, {
    colCount: Math.max(table.grid.length, ...p.rows.map((r) => r.length)),
    availablePt: LINE,
    padding: (cell) => ({
      left: cell.properties.margins?.left ?? table.properties.defaultCellMargins?.left ?? PAD,
      right: cell.properties.margins?.right ?? table.properties.defaultCellMargins?.right ?? PAD,
    }),
    content: (cell) => {
      const text = textOf(cell);
      const longest = Math.max(0, ...text.split(' ').map((w) => w.length));
      return { minPt: longest * 5, maxPt: text.length * 5 };
    },
  });
}

const near = (got: ReadonlyArray<number> | undefined, want: ReadonlyArray<number>): void => {
  expect(got).toBeDefined();
  expect(got!.length).toBe(want.length);
  got!.forEach((w, i) => expect(w, `column ${String(i + 1)}`).toBeCloseTo(want[i]!, 1));
};

const ab = (a: string, b: string): Probe['rows'] => [
  [
    [a, 'A'],
    [b, 'B'],
  ],
];

describe('a fixed table: the cells decide, the grid fills in', () => {
  const fixed = (tw: string, grid: ReadonlyArray<number>, rows: Probe['rows']): Probe => ({
    tblPr: tw + FIXED,
    grid,
    rows,
  });

  it('stretches cell widths to the table width, in proportion', () => {
    // Word: 1in+3in on a 100+100 twip grid drew 25% and 75% of the line
    // (the second column at 172.3pt); in a 6in table 108+324pt.
    near(widths(fixed(FULL, [100, 100], ab(IN1, IN3))), [LINE / 4, (LINE * 3) / 4]);
    near(widths(fixed(tblW('dxa', 8640), [100, 100], ab(IN1, IN3))), [108, 324]);
  });

  it('keeps cell widths as they are in a table of no width', () => {
    near(widths(fixed(AUTO, [100, 100], ab(IN1, IN3))), [72, 216]);
    near(widths(fixed(AUTO, [4000, 2000], ab(IN1, IN3))), [72, 216]);
    // Nor caps them at the line: 5in+5in ran 720pt.
    near(widths(fixed(AUTO, [100, 100], ab(tcW('dxa', 7200), tcW('dxa', 7200)))), [360, 360]);
  });

  it('shrinks by what each column has above its minimum', () => {
    // Word: 1in+3in in a 2in table drew 39.1+105.2pt, not 36+108.
    const min = 2 * PAD + 1;
    const take = (w: number): number => 144 * ((w - min) / (288 - 2 * min));
    near(widths(fixed(tblW('dxa', 2880), [100, 100], ab(IN1, IN3))), [
      72 - take(72),
      216 - take(216),
    ]);
  });

  it('gives a column no cell sizes its grid width', () => {
    near(widths(fixed(AUTO, [2000, 4000], ab(IN1, ''))), [72, 200]);
    // …which then stretches with the rest: 138.2pt for the inch.
    near(widths(fixed(FULL, [2000, 4000], ab(IN1, ''))), [(72 * LINE) / 272, (200 * LINE) / 272]);
    // A zero grid column is a quarter inch.
    near(widths(fixed(AUTO, [0, 2000], ab('', ''))), [18, 100]);
  });

  it('settles percentages first and gives the rest to absolute widths', () => {
    // 1in beside 50%: half the line each; with no table width, a 2in table.
    near(widths(fixed(FULL, [2000, 4000], ab(IN1, tcW('pct', 2500)))), [LINE / 2, LINE / 2]);
    near(widths(fixed(AUTO, [2000, 4000], ab(IN1, tcW('pct', 2500)))), [72, 72]);
    // 25% beside a 200pt grid column: a 266.7pt table.
    near(widths(fixed(AUTO, [2000, 4000], ab(tcW('pct', 1250), ''))), [200 / 3, 200]);
    // 75% beside 75%: the second gets what is left.
    near(widths(fixed(FULL, [100, 100], ab(tcW('pct', 3750), tcW('pct', 3750)))), [
      LINE * 0.75,
      LINE * 0.25,
    ]);
  });

  it('makes a table of percentages alone as narrow as its smallest share allows', () => {
    // 25%+75% of an auto-width table: 11.8+35.4pt.
    near(widths(fixed(AUTO, [2000, 4000], ab(tcW('pct', 1250), tcW('pct', 3750)))), [11.8, 35.4]);
  });

  it('never makes a cell narrower than its margins and 1pt', () => {
    // A 50-twip cell, a zero one and a nil one all drew 11.8pt.
    near(widths(fixed(AUTO, [100, 100], ab(tcW('dxa', 50), IN3))), [11.8, 216]);
    near(widths(fixed(AUTO, [100, 100], ab(tcW('dxa', 0), tcW('nil', 0)))), [11.8, 11.8]);
    // So did a table asking for nothing, or for less than its cells need.
    near(widths(fixed(tblW('pct', 0), [100, 100], ab('', ''))), [11.8, 11.8]);
    near(widths(fixed(tblW('dxa', 100), [100, 100], ab('', ''))), [11.8, 11.8]);
    // The minimum is the cell's own margins: 20pt a side made 41pt.
    const wide =
      '<w:tcMar><w:left w:w="400" w:type="dxa"/><w:right w:w="400" w:type="dxa"/></w:tcMar>';
    near(widths(fixed(AUTO, [100, 100], ab(wide, ''))), [41, 11.8]);
    // A negative width is no width.
    near(widths(fixed(AUTO, [100, 100], ab(tcW('dxa', -1440), IN3))), [11.8, 216]);
  });

  it('lets each column take the widest claim of its rows', () => {
    const rows: Probe['rows'] = [
      [
        [IN1, 'A'],
        [IN3, 'B'],
      ],
      [
        [IN3, 'C'],
        [IN1, 'D'],
      ],
    ];
    near(widths(fixed(AUTO, [100, 100], rows)), [216, 216]);
  });

  it('gives no width to a cell past the grid’s end that asks for none', () => {
    // A .doc row's trailing cells beyond its last edge (53379.doc).
    near(
      widths(
        fixed(
          AUTO,
          [2000, 4000],
          [
            [
              [tcW('dxa', 2000), 'A'],
              [tcW('dxa', 4000), 'B'],
              ['', 'C'],
            ],
          ],
        ),
      ),
      [100, 200, 0],
    );
  });

  it('sizes a spanned column from the cells around it', () => {
    const span2 = tcW('dxa', 2880) + '<w:gridSpan w:val="2"/>';
    const rows: Probe['rows'] = [
      [
        [span2, 'A'],
        [IN1, 'B'],
      ],
      [
        [IN1, 'C'],
        [span2, 'D'],
      ],
    ];
    near(widths(fixed(AUTO, [100, 100, 100], rows)), [72, 72, 72]);
  });
});

describe('an autofit table: laid out afresh, the grid ignored', () => {
  const autofit = (tw: string, grid: ReadonlyArray<number>, rows: Probe['rows']): Probe => ({
    tblPr: tw,
    grid,
    rows,
  });
  const cell = (text: string): number => text.length * 5 + 2 * PAD;

  it('takes the cells’ widths over a grid that says otherwise', () => {
    near(widths(autofit(AUTO, [100, 100], ab(IN1, IN3))), [72, 216]);
    near(widths(autofit(FULL, [2000, 4000], ab(IN1, IN3))), [LINE / 4, (LINE * 3) / 4]);
    // The docx package's own shape: gridCol 100 each, the widths in tcW.
    near(
      widths(autofit(tblW('auto', 100), [100, 100], ab(tcW('dxa', 3505), tcW('dxa', 5505)))),
      [175.25, 275.25],
    );
  });

  it('sizes columns no cell gives a width by their text, not by the grid', () => {
    // A grid of 2000+4000 twips around "A" and "B" drew 17.5+17.5pt.
    near(widths(autofit(AUTO, [2000, 4000], ab('', ''))), [cell('A'), cell('B')]);
    // In a full-width table they share the line by their text: 50/50 here.
    near(widths(autofit(FULL, [2000, 4000], ab('', ''))), [LINE / 2, LINE / 2]);
  });

  it('gives the slack to the columns of no width first', () => {
    // 1in beside an empty-width cell in a full-width table kept its inch.
    near(widths(autofit(FULL, [2000, 4000], ab(IN1, ''))), [72, LINE - 72]);
    // A percentage beside one: the share is of the table the text makes.
    const w = cell('A') / 0.25;
    near(widths(autofit(AUTO, [2000, 4000], ab(tcW('pct', 1250), ''))), [w / 4, (w * 3) / 4]);
  });

  it('caps a table of no width at the line, taking from its widest first', () => {
    near(widths(autofit(AUTO, [100, 100], ab(tcW('dxa', 7200), tcW('dxa', 7200)))), [
      LINE / 2,
      LINE / 2,
    ]);
    const long = Array.from({ length: 40 }, () => 'aaa').join(' ');
    const got = widths({
      grid: [100, 100],
      tblPr: AUTO,
      rows: [
        [
          ['', long],
          ['', 'B'],
        ],
      ],
    });
    near(got, [LINE - cell('B'), cell('B')]);
  });

  it('widens a column to its longest word, but not to its whole text', () => {
    const word = 'MMMMMMMMMMMMMMM';
    near(
      widths({
        tblPr: AUTO,
        grid: [1440, 1440],
        rows: [
          [
            [IN1, word],
            [IN1, 'B'],
          ],
        ],
      }),
      [cell(word), 72],
    );
    near(
      widths({
        tblPr: AUTO,
        grid: [100, 100],
        rows: [
          [
            [IN1, 'aaa bbb ccc ddd eee fff'],
            [IN1, 'B'],
          ],
        ],
      }),
      [72, 72],
    );
  });

  it('keeps a column a row starts in past the others', () => {
    // gridbefore.docx: one row starts its only cell two columns in.
    const got = widths({
      grid: [132, 1872, 685],
      rows: [
        [['', 'A3']],
        [
          ['', 'B1'],
          ['', 'B2'],
        ],
      ],
      gridBefore: [2, 0],
    });
    near(got, [cell('B1'), cell('B2'), cell('A3')]);
  });

  it('breaks the words rather than run past the line', () => {
    // A 600pt word beside "B": the text, not the margins, shrinks to fit.
    const word = 'M'.repeat(120);
    const text = LINE - 4 * PAD;
    near(
      widths({
        tblPr: AUTO,
        grid: [100, 100],
        rows: [
          [
            ['', word],
            ['', 'B'],
          ],
        ],
      }),
      [(600 * text) / 605 + 2 * PAD, (5 * text) / 605 + 2 * PAD],
    );
    // A stated width the words outgrow stops at the line too: 261.4pt each.
    const half = 'M'.repeat(80);
    near(
      widths({
        tblPr: tblW('dxa', 8640),
        grid: [100, 100],
        rows: [
          [
            ['', half],
            ['', half],
          ],
        ],
      }),
      [LINE / 2, LINE / 2],
    );
  });

  it('widens a span too narrow for its word by each column’s width and text', () => {
    // Word set a 400pt word over 1in and 3in at 113.8+296.9pt.
    const span = '<w:gridSpan w:val="2"/>';
    const need = 400 + 2 * PAD;
    const a = 72 + cell('A');
    const b = 216 + cell('B');
    near(
      widths({
        tblPr: AUTO,
        grid: [100, 100],
        rows: [
          [[span, 'M'.repeat(80)]],
          [
            [IN1, 'A'],
            [IN3, 'B'],
          ],
        ],
      }),
      [(need * a) / (a + b), (need * b) / (a + b)],
    );
  });

  it('stretches every column under a full-width span, sized ones too', () => {
    // Word widened "A. Each", "B. Every", a 117.5pt cell and "D. A lot" alike
    // under a first row of the table's full width.
    const fullRow = tcW('dxa', 10466) + '<w:gridSpan w:val="4"/>';
    const auto = tcW('auto', 0);
    const row: Probe['rows'][number] = [
      [auto, 'A. Each'],
      [auto, 'B. Every'],
      [tcW('dxa', 2350), 'C. Most'],
      [auto, 'D. A lot'],
    ];
    const grid = [2213, 2281, 3446, 2919];
    const want = [cell('A. Each'), cell('B. Every'), 117.5, cell('D. A lot')];
    const sum = want.reduce((x, y) => x + y, 0);
    near(
      widths({ tblPr: tblW('dxa', 10466), grid, rows: [[[fullRow, 'Title']], row] }),
      want.map((w) => (w * 523.3) / sum),
    );
    // Without that row the sized column keeps its width and the rest share
    // the slack by their text.
    const free = 523.3 - 117.5;
    const autos = want[0]! + want[1]! + want[3]!;
    near(widths({ tblPr: tblW('dxa', 10466), grid, rows: [row] }), [
      (want[0]! * free) / autos,
      (want[1]! * free) / autos,
      117.5,
      (want[3]! * free) / autos,
    ]);
  });

  it('keeps the grid where a span’s columns add up to what it declares', () => {
    // tdf102466.docx: no cell of its own sizes the middle column; the spans
    // around it say it is 69.9pt, and so does the grid.
    const rows: Probe['rows'] = [
      [
        [tcW('dxa', 3662), 'A'],
        [tcW('dxa', 4842) + '<w:gridSpan w:val="2"/>', 'B'],
      ],
      [
        [tcW('dxa', 5058) + '<w:gridSpan w:val="2"/>', 'C'],
        [tcW('dxa', 3446), 'D'],
      ],
    ];
    expect(widths(autofit(FULL, [3662, 1396, 3446], rows))).toBeUndefined();
    // On a grid that does not say so, the spans size it: 2in less 1in.
    near(
      widths(
        autofit(
          AUTO,
          [100, 100, 100],
          [
            [
              [tcW('dxa', 1440), 'A'],
              [tcW('dxa', 2880) + '<w:gridSpan w:val="2"/>', 'B'],
            ],
            [
              [tcW('dxa', 2880) + '<w:gridSpan w:val="2"/>', 'C'],
              [tcW('dxa', 1440), 'D'],
            ],
          ],
        ),
      ),
      [72, 72, 72],
    );
  });

  it('keeps the grid where it is what the cells declare', () => {
    // 2_table_doc.docx: two 4428-twip cells on a 4428+4428 grid.
    const cells = ab(tcW('dxa', 4428), tcW('dxa', 4428));
    expect(widths(autofit(AUTO, [4428, 4428], cells))).toBeUndefined();
    // Not where a word would not fit it.
    const tight = [
      [
        [IN1, 'MMMMMMMMMMMMMMM'],
        [IN1, 'B'],
      ],
    ] as const;
    expect(widths({ tblPr: AUTO, grid: [1440, 1440], rows: tight })).toBeDefined();
    // Nor where a column has no width of its own to check it against.
    expect(widths(autofit(AUTO, [4428, 4428], ab(tcW('dxa', 4428), '')))).toBeDefined();
  });
});

describe('Word table columns through the layout', () => {
  const FONTS = {
    regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
    bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  };
  const SECT =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>';

  // Word 2013's layout: the table's edge at the margin (see table-compat-outdent).
  const MODE15 =
    '<w:compat><w:compatSetting w:name="compatibilityMode" ' +
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';

  /** Where the text "B" starts on the page, in a document of `mode` ('none': states no mode). */
  function startOfB(p: Probe, mode: number | 'none' = 15): number {
    const settingsXml =
      mode === 'none' ? '' : MODE15.replace('w:val="15"', `w:val="${String(mode)}"`);
    const flow = Ream.parse(buildDocxFromBody(xml(p) + SECT, { settingsXml })).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const line = laid.pages[0]!.commands.find(
      (c): c is TextLineItem =>
        c.type === 'line' && c.line.tokens.some((t) => t.kind === 'text' && t.text === 'B'),
    );
    return line!.originX;
  }

  it('draws the docx package’s table at its cells’ widths', () => {
    // Word: the second column's text at 217.0pt (36 + 175.25 + 5.4).
    expect(
      startOfB({
        tblPr: tblW('auto', 100),
        grid: [100, 100],
        rows: ab(tcW('dxa', 3505), tcW('dxa', 5505)),
      }),
    ).toBeCloseTo(36 + 175.25 + PAD, 1);
  });

  it('stretches a fixed table’s inches over its full width', () => {
    // Word: 172.3pt.
    expect(startOfB({ tblPr: FULL + FIXED, grid: [100, 100], rows: ab(IN1, IN3) })).toBeCloseTo(
      36 + LINE / 4 + PAD,
      1,
    );
  });

  it('breaks a word after its hyphens when sizing a column', () => {
    // fdo68607.docx: "Theatre-in-the-round" in a 1in column Word did not widen.
    const p = {
      tblPr: AUTO,
      grid: [1440, 1440],
      rows: [
        [
          [IN1, 'Theatre-in-the-round'],
          [IN1, 'B'],
        ],
      ],
    } as const;
    expect(startOfB(p)).toBeCloseTo(36 + 72 + PAD, 1);
  });

  it('keeps cells of no width at least their margins and 1pt apart', () => {
    // Word: "B" at 53.5pt; the 5pt grid alone set it over "A".
    expect(
      startOfB({ tblPr: tblW('pct', 0) + FIXED, grid: [100, 100], rows: ab('', '') }),
    ).toBeCloseTo(36 + 11.8 + PAD, 1);
  });

  it('measures a share against the column and the outer cell margins before Word 2013', () => {
    // Word, modes 14 and none: a 50% table of two empty-width cells was
    // 267.05pt, half of 523.3 + 5.4 + 5.4, standing out 5.4pt; mode 15 halves
    // the column alone.
    const half: Probe = { tblPr: tblW('pct', 2500) + FIXED, grid: [100, 100], rows: ab('', '') };
    for (const mode of [14, 'none'] as const) {
      expect(startOfB(half, mode)).toBeCloseTo(36 - PAD + (LINE + 2 * PAD) / 4 + PAD, 1);
    }
    expect(startOfB(half, 15)).toBeCloseTo(36 + LINE / 4 + PAD, 1);
    // An autofit table of no width stops at that wider line too.
    const wide: Probe = {
      tblPr: AUTO,
      grid: [100, 100],
      rows: ab(tcW('dxa', 7200), tcW('dxa', 7200)),
    };
    expect(startOfB(wide, 14)).toBeCloseTo(36 - PAD + (LINE + 2 * PAD) / 2 + PAD, 1);
  });
});
