// Where the grid actually lands on the page.
//
// The rest of the xlsx suite asks what is on the page; this asks where. Both
// questions matter, but only the first had coverage, which is why a sheet could
// render every cell correctly and still put the grid somewhere Excel never
// would.
//
// These assert against the spreadsheet model rather than against LibreOffice:
// a declared column width is a number in the file, and honouring it is not a
// matter of taste. The golden suite (xlsx-golden.test.ts) is where agreement
// with another renderer is measured.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { FontRegistry } from '@/core/font';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { readXlsx, readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { parsePrinterSettings } from '@/excel/printer-settings';
import { projectSheetDoc } from '@/excel/sheet-to-flow';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

interface PlacedText {
  readonly text: string;
  readonly x: number;
  readonly y: number;
}

/** Every text line the layout placed, with its position. */
function placed(xlsx: Uint8Array): Array<PlacedText> {
  const flow = Ream.parse(xlsx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const out: Array<PlacedText> = [];
  for (const page of laid.pages) {
    for (const command of page.commands) {
      if (command.type !== 'line') continue;
      const c = command as unknown as {
        originX: number;
        baselineY: number;
        line: { tokens: ReadonlyArray<{ text?: string }> };
      };
      const text = c.line.tokens
        .map((t) => t.text ?? '')
        .join('')
        .trim();
      if (text.length > 0) out.push({ text, x: c.originX, y: c.baselineY });
    }
  }
  return out;
}

/** How many pages the sheet lays out to. */
function pageCount(xlsx: Uint8Array): number {
  const flow = Ream.parse(xlsx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  }).pages.length;
}

const at = (items: Array<PlacedText>, text: string): PlacedText => {
  const hit = items.find((i) => i.text === text);
  if (!hit) throw new Error(`no placed text "${text}" among ${items.map((i) => i.text).join('|')}`);
  return hit;
};

describe('grid geometry', () => {
  it('honours declared column widths instead of fitting them to content', () => {
    // A is declared WIDE and holds a short value; B is declared NARROW and holds
    // a long one. Auto-fit — the table default, since tblGrid is only a hint in
    // WordprocessingML — sizes them by content and produces the reverse. In a
    // spreadsheet the declared width is not a hint: it is what the author set,
    // it is what Excel and LibreOffice print, and getting it wrong moves every
    // column after it.
    const xlsx = buildXlsx({
      rows: [['x', 'a very long value indeed']],
      columns: [
        { min: 1, max: 1, widthChars: 40 },
        { min: 2, max: 2, widthChars: 6 },
      ],
    });
    const items = placed(xlsx);
    const a = at(items, 'x');
    const b = at(items, 'a very long value indeed');

    // Column A is 40 chars ≈ 210pt wide, so B starts far to the right of A.
    // Under auto-fit the two nearly touch.
    expect(b.x - a.x).toBeGreaterThan(150);
  });

  it('gives a declared width the 5-pixel padding the spec puts on it', () => {
    // §18.3.1.13: px = chars × MaximumDigitWidth + 5. At 96 DPI that padding is
    // 3.75pt, and dropping it made every column that much narrow — an error
    // that compounds across the sheet. Excel's own documented 8.43-character
    // default only reaches its documented 64px with the padding included.
    const xlsx = buildXlsx({
      rows: [['A', 'B']],
      columns: [{ min: 1, max: 1, widthChars: 12 }],
    });
    const items = placed(xlsx);
    // 12 × 5.25pt + 3.75pt = 66.75pt.
    expect(at(items, 'B').x - at(items, 'A').x).toBeCloseTo(66.75, 1);
  });

  it("uses Excel's page margins, not a word processor's, when the sheet declares none", () => {
    // §18.3.1.62: 0.7" left/right, 0.75" top/bottom. The renderer's fallback is
    // 1 inch, which pushed the grid 0.3" (21.6pt) right on every sheet that
    // omits <pageMargins> — which most do.
    const noMargins = buildXlsx([['A']]);
    const declared = buildXlsx({
      rows: [['A']],
      pageMargins: { left: 2, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    });
    const left = at(placed(noMargins), 'A').x;
    // 0.7 inch = 50.4pt, plus the cell's own left padding.
    expect(left).toBeGreaterThan(50.4);
    expect(left).toBeLessThan(50.4 + 12);
    // An explicit margin still wins: 2 inches = 144pt.
    expect(at(placed(declared), 'A').x).toBeGreaterThan(144);
  });

  it('applies §18.8.1 general alignment — numbers right, text left, booleans centred', () => {
    // "General" is not "left": it is decided by the value's type. A column of
    // figures rendered flush left is the most visible way a spreadsheet can
    // look wrong, and it was the largest remaining gap against LibreOffice
    // once column widths were fixed.
    const xlsx = buildXlsx({
      rows: [['text', 1234, true]],
      columns: [{ min: 1, max: 3, widthChars: 20 }],
    });
    const items = placed(xlsx);
    const colWidth = 20 * 5.25 + 3.75;
    const text = at(items, 'text');
    const number = at(items, '1234');
    const bool = at(items, 'TRUE');

    // Text sits at its column's left edge; the number is pushed to the right of
    // its own column, and the boolean sits between the two.
    const numberOffset = number.x - (text.x + colWidth);
    const boolOffset = bool.x - (text.x + 2 * colWidth);
    expect(numberOffset).toBeGreaterThan(colWidth / 2);
    expect(boolOffset).toBeGreaterThan(colWidth / 4);
    expect(boolOffset).toBeLessThan(numberOffset);
  });

  it('shrinks column widths with the print scale, not just fonts and rows', () => {
    // <pageSetup scale="50"> shrinks the whole sheet. The scale reached fonts
    // and row heights but not the emitted column grid, which was invisible
    // while the layout auto-fitted that grid away — and became a clipped page
    // the moment the grid started to count. 49156.xlsx (scale="47") went from
    // 349 extracted characters to 43.
    const rows = [['A', 'B', 'C']];
    const columns = [{ min: 1, max: 3, widthChars: 20 }];
    const full = placed(buildXlsx({ rows, columns }));
    const half = placed(buildXlsx({ rows, columns, pageSetup: { paperSize: 9, scale: 50 } }));
    const pitch = (items: Array<PlacedText>): number => at(items, 'B').x - at(items, 'A').x;
    expect(pitch(half)).toBeCloseTo(pitch(full) / 2, 0);
  });

  // Five 300pt rows asked to fit two A4 pages. The closed form divides an area:
  // 2 × 734pt of body over 1500pt of rows is 97 %, which packs 2 rows per page
  // and takes three. Rows do not split, so the answer is 81 % — three per page.
  const tallFit = (fitToHeight: number) =>
    buildXlsx({
      rows: [['r0'], ['r1'], ['r2'], ['r3'], ['r4']],
      rowHeights: [0, 1, 2, 3, 4].map((row) => ({ row, heightPt: 300 })),
      fitToPage: true,
      pageSetup: { paperSize: 9, fitToWidth: 1, fitToHeight },
    });

  it('shrinks fit-to-page until the rows really fit, not until the areas do', () => {
    expect(pageCount(tallFit(2))).toBe(2);
    expect(pageCount(tallFit(3))).toBe(3);
  });

  it('leaves the scale alone when no amount of shrinking reaches the target', () => {
    // A manual break costs a page at every scale, so `fitToHeight="1"` here is
    // unreachable. Searching for it anyway drove the scale into its 10 % floor
    // and rendered AverageTaxRates.xlsx at 1pt — when the target cannot be met,
    // the closed form stands and the sheet simply takes the pages it takes.
    const rows = [['A', 'B'], ['x'], ['y']];
    const columns = [{ min: 1, max: 2, widthChars: 20 }];
    const plain = placed(buildXlsx({ rows, columns }));
    const unreachable = placed(
      buildXlsx({
        rows,
        columns,
        rowBreaks: [1],
        fitToPage: true,
        pageSetup: { paperSize: 9, fitToWidth: 1, fitToHeight: 1 },
      }),
    );
    const pitch = (items: Array<PlacedText>): number => at(items, 'B').x - at(items, 'A').x;
    expect(pitch(unreachable)).toBeCloseTo(pitch(plain), 1);
  });

  it("gives every row a definite height instead of the font's natural leading", () => {
    // §18.3.1.81. A row without an explicit `ht` had no height at all, so its
    // pitch came out as whatever leading the rendering font wanted — 13.2pt for
    // 11pt Roboto — making the row pitch a property of the typesetter rather
    // than of the document. Excel's default is 15pt regardless of the font it
    // is drawn with.
    const items = placed(buildXlsx([['r0'], ['r1'], ['r2'], ['r3']]));
    const ys = ['r0', 'r1', 'r2', 'r3'].map((t) => at(items, t).y);
    for (let i = 1; i < ys.length; i++) {
      expect(ys[i]! - ys[i - 1]!).toBeCloseTo(15, 1);
    }
  });

  it('honours <sheetFormatPr defaultRowHeight> over the 15pt fallback', () => {
    const items = placed(buildXlsx({ rows: [['r0'], ['r1'], ['r2']], defaultRowHeightPt: 24 }));
    const ys = ['r0', 'r1', 'r2'].map((t) => at(items, t).y);
    expect(ys[1]! - ys[0]!).toBeCloseTo(24, 1);
    expect(ys[2]! - ys[1]!).toBeCloseTo(24, 1);
  });

  it('lets an explicit row ht override the sheet default', () => {
    const items = placed(
      buildXlsx({
        rows: [['r0'], ['r1'], ['r2']],
        rowHeights: [{ row: 1, heightPt: 40, customHeight: true }],
      }),
    );
    const ys = ['r0', 'r1', 'r2'].map((t) => at(items, t).y);
    // A cell sits at the BOTTOM of its box (§18.8.1), so the step between two
    // rows' text is the height of the LOWER one: r1 is pinned to 40pt and its
    // text drops to the foot of that box, then r2's default 15pt follows.
    expect(ys[1]! - ys[0]!).toBeCloseTo(40, 1);
    expect(ys[2]! - ys[1]!).toBeCloseTo(15, 1);
  });

  it('runs unwrapped text across empty neighbours on ONE line', () => {
    // Excel does not wrap a cell unless wrapText is set: the text runs over the
    // empty cells to its right. We kept the full text but left the cell one
    // column wide, so the layout broke it into a stack of lines — three where
    // Excel and LibreOffice draw one (RepeatingRowsCols.xlsx).
    const phrase = 'no repeating rows or columns';
    const items = placed(
      buildXlsx({
        // Row 2 anchors the used range at four columns; row 1 leaves B..D empty
        // so the phrase has somewhere to run.
        rows: [[phrase], [null, null, null, 'anchor']],
        columns: [{ min: 1, max: 4, widthChars: 10 }],
      }),
    );
    const lines = items.filter((i) => phrase.startsWith(i.text));
    expect(lines).toHaveLength(1);
    expect(lines[0]!.text).toBe(phrase);
  });

  it('cuts a non-wrapping cell at the glyph, not back to the last space', () => {
    // A cell that does not wrap is cut where its box ends. Taking the line
    // breaker's first line looks equivalent and is not: the breaker only breaks
    // at spaces, so a string a hair too wide loses its whole last WORD.
    // tdf82984's "Carta geologica - litologica" printed as "Carta" — a quarter
    // of the column, with 22pt of white beside neighbours that ran full.
    const phrase = 'Carta geologica - litologica';
    const items = placed(
      buildXlsx({
        // The neighbour holds content, so the text cannot overflow into it and
        // has to be cut inside its own column.
        rows: [[phrase, 'x']],
        columns: [{ min: 1, max: 2, widthChars: 9 }],
      }),
    );
    const shown = at(items, items.find((i) => i.text.startsWith('Carta'))?.text ?? '').text;
    expect(phrase.startsWith(shown)).toBe(true);
    // More than the first word, and not the whole string either.
    expect(shown.length).toBeGreaterThan('Carta'.length);
    expect(shown.length).toBeLessThan(phrase.length);
  });

  it('stops the overflow at a neighbour that paints its own fill', () => {
    // An empty cell carrying a fill is not free space: spanning over it to give
    // the text room would take its paint with it. Clip instead — visibly short
    // beats visibly wrong.
    const stylesXml = `
      <fonts count="1"><font/></fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
      </fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
      </cellXfs>`;
    const phrase = 'no repeating rows or columns';
    const rows = [
      // B1 is EMPTY but filled green, so the text may not run over it.
      [phrase, { value: null, styleIndex: 1 }],
      [null, null, null, 'anchor'],
    ];
    const columns = [{ min: 1, max: 4, widthChars: 10 }];
    const blocked = placed(buildXlsx({ rows, columns, stylesXml }));
    // Same sheet without the fill: there the phrase does run across.
    const free = placed(
      buildXlsx({ rows: [[phrase], [null, null, null, 'anchor']], columns, stylesXml }),
    );
    expect(free.map((i) => i.text)).toContain(phrase);
    expect(blocked.map((i) => i.text)).not.toContain(phrase);
  });

  it('runs the text over the rule that CLOSES its band, taking it along', () => {
    // A vertical rule inside the run would be erased by the span, so the text
    // stops at it. The rule that closes the run is different: it is the band's
    // own right edge, and the span's right border lands in the same place. We
    // refused it, and tdf171828.xlsx printed "ohne Sondertil" for "ohne
    // Sondertilgung" on every page — five labels cut inside a filled box whose
    // far edge is exactly where the reference ends each of them.
    const stylesXml = `
      <fonts count="1"><font/></fonts>
      <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
      </fills>
      <borders count="2">
        <border/>
        <border><right style="thin"><color rgb="FF000000"/></right></border>
      </borders>
      <cellXfs count="3">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="1" applyFill="1" applyBorder="1"/>
      </cellXfs>`;
    const phrase = 'no repeating rows or columns';
    const columns = [{ min: 1, max: 4, widthChars: 10 }];
    const closed = placed(
      buildXlsx({
        rows: [
          // A green run: the text, one plain green cell, then the green cell
          // that carries the band's closing rule.
          [
            { value: phrase, styleIndex: 1 },
            { value: null, styleIndex: 1 },
            { value: null, styleIndex: 2 },
          ],
          [null, null, null, 'anchor'],
        ],
        columns,
        stylesXml,
      }),
    );
    expect(closed.map((i) => i.text)).toContain(phrase);
  });

  it('derives the default column width from <sheetFormatPr baseColWidth>', () => {
    // baseColWidth is what Excel computes the default column FROM when
    // defaultColWidth is absent. Ignoring a sheet that asks for 10 left every
    // unlisted column at 8.43 and the whole grid narrow.
    const wide = placed(buildXlsx({ rows: [['A', 'B']], baseColWidthChars: 20 }));
    const narrow = placed(buildXlsx({ rows: [['A', 'B']] }));
    const pitch = (items: Array<PlacedText>): number => at(items, 'B').x - at(items, 'A').x;
    expect(pitch(wide)).toBeCloseTo(20 * 5.25 + 3.75, 1);
    expect(pitch(narrow)).toBeCloseTo(48, 1); // Excel's 8.43-char default
  });

  it('gives each sheet its own page geometry', () => {
    // A workbook's sheets set their paper independently, and applying the first
    // sheet's setup to the whole document silently reprints the others on the
    // wrong paper. tdf171828_fail_to_import_file.xlsx is A4 landscape then
    // Letter portrait then A4 landscape; LibreOffice emits 17 landscape pages
    // and 39 portrait ones, we emitted 56 landscape.
    const flow = Ream.parse(
      buildXlsx({
        sheets: [
          { name: 'Wide', rows: [['a']], pageSetup: { paperSize: 9, orientation: 'landscape' } },
          { name: 'Tall', rows: [['b']], pageSetup: { paperSize: 1, orientation: 'portrait' } },
        ],
      }),
    ).flow;

    expect(flow.sections).toHaveLength(2);
    const [first, second] = flow.sections;
    expect(first!.properties.pageSize?.orientation).toBe('landscape');
    expect(second!.properties.pageSize?.orientation).toBe('portrait');
    // A4 landscape is wider than it is tall; Letter portrait is the reverse.
    expect(first!.properties.pageSize!.width).toBeGreaterThan(first!.properties.pageSize!.height);
    expect(second!.properties.pageSize!.width).toBeLessThan(second!.properties.pageSize!.height);
    // Every section ends inside the body, in order.
    expect(first!.endIndex).toBeLessThan(second!.endIndex);
    expect(second!.endIndex).toBe(flow.body.length);
  });

  it('paginates empty rows the same as filled ones', () => {
    // Pagination was gated on whether a page had received a drawable ITEM, and
    // a page holding none was discarded outright. Empty rows consume space and
    // draw nothing, so no break ever fired and they marched off the bottom of
    // page one — a sheet of 292 rows came out on a single page. A spreadsheet
    // is mostly empty rows; they are not nothing, they are blank space that has
    // to be paged like any other.
    const tall = 60;
    const count = 40;
    const withText = buildXlsx({
      rows: Array.from({ length: count }, (_, i) => [`r${i}`]),
      rowHeights: Array.from({ length: count }, (_, i) => ({
        row: i,
        heightPt: tall,
        customHeight: true,
      })),
    });
    // Same geometry, but only the first and last rows carry a value — the rest
    // are blank rows of the same declared height.
    const mostlyEmpty = buildXlsx({
      rows: Array.from({ length: count }, (_, i) => (i === 0 || i === count - 1 ? [`r${i}`] : [])),
      rowHeights: Array.from({ length: count }, (_, i) => ({
        row: i,
        heightPt: tall,
        customHeight: true,
      })),
    });
    expect(pageCount(mostlyEmpty)).toBe(pageCount(withText));
    expect(pageCount(withText)).toBeGreaterThan(1);
  });

  it('renders a sheet as wide as SpreadsheetML allows, without an extra limit', () => {
    // 53105.xlsx is 2 rows across all 16 384 columns. A 1024-column cap cut it
    // to 103 pages where Excel and LibreOffice print 1639 — a limit of ours, on
    // the reasoning that a wider table is unreadable, applied to a document
    // that is merely wide. The memory bound is the total-cell budget, which a
    // 2-row sheet comes nowhere near.
    const wide = buildXlsx([Array.from({ length: 5000 }, (_, i) => `c${i}`)]);
    const { doc, losses } = readXlsx(wide);
    let columns = 0;
    for (const element of doc.body) {
      if (element.kind === 'table') columns += element.table.grid.length;
    }
    expect(columns).toBe(5000);
    expect(losses.filter((l) => /columns/.test(l.detail))).toEqual([]);
  });

  it('still bounds a sheet by its total cells', () => {
    // The budget that actually protects memory stays, and still reports. Two
    // cells at opposite corners declare a 2000×600 used range — 1.2M cells,
    // past the budget — without costing anything to build, which is exactly the
    // amplification shape the budget exists for.
    const lastRow = [...Array.from<null>({ length: 599 }).fill(null), 'y'];
    const sparse = Array.from({ length: 2000 }, (_, r) =>
      r === 0 ? ['x'] : r === 1999 ? lastRow : [],
    );
    const { losses } = readXlsx(buildXlsx(sparse));
    expect(losses.some((l) => l.severity === 'dropped' && /rows/.test(l.detail))).toBe(true);
  });

  it('does not let an empty merge stretch the used range', () => {
    // A merge follows content; it does not create any. 57893-many-merges.xlsx
    // is 50 000 rows of merges over cells that hold nothing at all, and letting
    // those merges set the used range turned it into 1042 blank pages once
    // blank space started paginating honestly. LibreOffice prints one.
    const anchored = pageCount(buildXlsx({ rows: [['x']], mergeRefs: ['A1:C1'] }));
    const stranded = pageCount(buildXlsx({ rows: [['x']], mergeRefs: ['A1:C1', 'E200:F400'] }));
    expect(stranded).toBe(anchored);
    expect(stranded).toBe(1);
  });

  it('still follows a merge that starts on a cell with content', () => {
    // The merge loop exists for a reason: a merged cell's value lives in the
    // origin and the span reaches past it, so the grid has to cover the span or
    // the merge is clipped.
    const { doc } = readXlsx(buildXlsx({ rows: [['spanning value']], mergeRefs: ['A1:D1'] }));
    const table = doc.body.find((e) => e.kind === 'table');
    expect(table?.kind === 'table' && table.table.grid.length).toBe(4);
  });

  it('prints neither a hidden column nor a hidden row', () => {
    // §18.3.1.13 / §18.3.1.73 `hidden`. Excel and LibreOffice print neither, and
    // rendering them is not a cosmetic difference: AverageTaxRates.xlsx hides a
    // currency column between two visible ones and seven rows in the middle of
    // its table, so showing them inserted a column no other reader shows and
    // broke the table's structure around it.
    const xlsx = buildXlsx({
      rows: [
        ['keep-a', 'HIDE-COL', 'keep-b'],
        ['HIDE-ROW', 'HIDE-ROW', 'HIDE-ROW'],
        ['keep-c', 'HIDE-COL', 'keep-d'],
      ],
      columns: [
        { min: 1, max: 1, widthChars: 12 },
        { min: 2, max: 2, widthChars: 12, hidden: true },
        { min: 3, max: 3, widthChars: 12 },
      ],
      rowHeights: [{ row: 1, heightPt: 15, hidden: true }],
    });
    const drawn = placed(xlsx).map((i) => i.text);
    expect(drawn).toEqual(expect.arrayContaining(['keep-a', 'keep-b', 'keep-c', 'keep-d']));
    expect(drawn).not.toContain('HIDE-COL');
    expect(drawn).not.toContain('HIDE-ROW');

    // The grid loses the column too, so nothing after it is displaced.
    const { doc } = readXlsx(xlsx);
    const table = doc.body.find((e) => e.kind === 'table');
    expect(table?.kind === 'table' && table.table.grid.length).toBe(2);
  });

  it('counts an overflow span in VISIBLE columns, so it cannot swallow a value', () => {
    // A span is consumed against the emitted grid, which skips hidden columns.
    // Counting absolute ones instead overruns by however many are hidden inside
    // the run, and an overrunning span covers the cell PAST its end — taking
    // its content off the page. tdf100034.xlsx hides two columns mid-row and
    // lost the value in the one after them on two of its four pages.
    const xlsx = buildXlsx({
      // A runs long over the empty B and the hidden C/D; E holds a value that
      // stops the run and must survive it.
      rows: [['a label far too long for its own column', null, null, null, 'KEEP']],
      columns: [
        { min: 1, max: 2, widthChars: 8 },
        { min: 3, max: 4, widthChars: 8, hidden: true },
        { min: 5, max: 5, widthChars: 8 },
      ],
    });
    expect(placed(xlsx).map((i) => i.text)).toContain('KEEP');
    const { doc } = readXlsx(xlsx);
    const table = doc.body.find((e) => e.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    // Three visible columns; the run covers the first two and leaves the third.
    expect(table.table.grid).toHaveLength(3);
    const cells = table.table.rows[0]!.cells;
    expect(cells[0]?.properties.colSpan).toBe(2);
    expect(cells).toHaveLength(2);
  });

  it('spills a CENTRED cell too, not only a left-aligned one', () => {
    // Excel and Calc run a centred cell out both ways. Gating overflow on left
    // alignment kept tdf171828.xlsx's centred "unter Berücksichtung der
    // Sondertilgungen" inside its own 73pt column, cut to "unter Berücksic".
    const phrase = 'a centred label much wider than its column';
    const stylesXml =
      `<fonts count="1"><font/></fonts><fills count="1"><fill/></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellXfs count="2"><xf/>` +
      `<xf applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs>`;
    const items = placed(
      buildXlsx({
        rows: [[{ value: phrase, styleIndex: 1 }], [null, null, null, 'anchor']],
        columns: [{ min: 1, max: 4, widthChars: 10 }],
        stylesXml,
      }),
    );
    expect(items.map((i) => i.text)).toContain(phrase);
  });

  it('does not overflow into a cell that belongs to a merge', () => {
    // An empty cell inside a merge is not free space — it is that merge's. We
    // spanned over a merge ORIGIN, which made the origin's own span disappear
    // and left the rest of the merge painting nothing: tdf171828.xlsx lost the
    // fill and the rule over a whole column of one row.
    const phrase = 'a label much wider than its own column';
    const xlsx = buildXlsx({
      rows: [
        [phrase, null, null, null],
        [null, null, null, 'anchor'],
      ],
      columns: [{ min: 1, max: 4, widthChars: 10 }],
      mergeRefs: ['B1:C1'],
    });
    const { doc } = readXlsx(xlsx);
    const table = doc.body.find((e) => e.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    const first = table.table.rows[0]!.cells;
    // The label stays in its own column, and the merge keeps its two.
    expect(first[0]?.properties.colSpan ?? 1).toBe(1);
    expect(first[1]?.properties.colSpan).toBe(2);
  });

  it('keeps equal declared widths equally spaced', () => {
    // Four columns declared identical must come out identical. Auto-fit makes
    // each one as wide as its own content, so the pitch wanders — which is what
    // pushed our columns out of step with LibreOffice's by up to 80pt.
    const xlsx = buildXlsx({
      rows: [['A', 'BB', 'CCC', 'DDDD']],
      columns: [{ min: 1, max: 4, widthChars: 12 }],
    });
    const items = placed(xlsx);
    const xs = ['A', 'BB', 'CCC', 'DDDD'].map((t) => at(items, t).x);
    const pitches = [xs[1]! - xs[0]!, xs[2]! - xs[1]!, xs[3]! - xs[2]!];
    for (const pitch of pitches) {
      expect(pitch).toBeCloseTo(pitches[0]!, 1);
    }
  });
});

describe('column width unit (§18.3.1.13)', () => {
  it('measures a column in the digit width of the font it is drawn in', () => {
    // The unit is the workbook default font's Maximum Digit Width. When the
    // file names a font — Calibri 11, Arial 10 — that width is Excel's
    // documented 7px, and a metric-compatible substitute reproduces it. When it
    // names none, the reader's own face decides: tdf122336.xlsx declares
    // `<font/>` and LibreOffice laid its 40-character columns out one per page
    // where we fitted two.
    const named = buildXlsx({
      rows: [['x']],
      columns: [{ min: 1, max: 1, widthChars: 40 }],
      stylesXml:
        `<fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>` +
        `<fills count="1"><fill/></fills><borders count="1"><border/></borders>` +
        `<cellXfs count="1"><xf/></cellXfs>`,
    });
    const anonymous = buildXlsx({
      rows: [['x']],
      columns: [{ min: 1, max: 1, widthChars: 40 }],
    });
    // The projection takes the measured digit width; the reader alone has no
    // font, so both would fall back to Excel's unit.
    const width = (bytes: Uint8Array): number => {
      const flow = projectSheetDoc(readXlsxToSheetDoc(bytes), { digitWidthPt: 6.18 });
      const table = flow.body.find((e) => e.kind === 'table');
      if (table?.kind !== 'table') throw new Error('expected a table');
      return table.table.grid[0]!;
    };
    // Excel's own unit: 40 × 7px + 5px of padding.
    expect(width(named)).toBeCloseTo(40 * 5.25 + 3.75, 1);
    // The reader's face is wider than Calibri's, so the column is too.
    expect(width(anonymous)).toBeGreaterThan(width(named));
  });
});

describe('a table that spills over a page break', () => {
  // Every cell a thin box; 80 rows so the grid takes more than one page.
  const bordered = (): Uint8Array =>
    buildXlsx({
      rows: Array.from({ length: 80 }, (_, i) => [{ value: `r${i}`, styleIndex: 1 }]),
      stylesXml:
        `<fonts count="1"><font/></fonts><fills count="1"><fill/></fills>` +
        `<borders count="2"><border/><border>` +
        ['left', 'right', 'top', 'bottom']
          .map((s) => `<${s} style="thin"><color rgb="FF000000"/></${s}>`)
          .join('') +
        `</border></borders>` +
        `<cellXfs count="2"><xf/><xf borderId="1" applyBorder="1"/></cellXfs>`,
    });

  it('closes the table at the bottom of the page it leaves', () => {
    // Only the last row paints its bottom edge — every other horizontal rule is
    // the row below it painting its top. Across a page break that row is on the
    // NEXT page, so the table was left hanging open: on tdf58243.xlsx the
    // vertical rules ran past the last row into white space.
    const flow = Ream.parse(bordered()).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    expect(laid.pages.length).toBeGreaterThan(1);
    const lowestBottom = (page: (typeof laid.pages)[number]): number =>
      Math.max(
        ...page.commands
          .filter((c): c is typeof c & { side: string; y: number; height: number } => {
            const item = c as { type: string; side?: string };
            return item.type === 'border' && item.side === 'bottom';
          })
          .map((c) => c.y + c.height),
      );
    const lowestTop = (page: (typeof laid.pages)[number]): number =>
      Math.max(
        ...page.commands
          .filter((c): c is typeof c & { side: string; y: number; height: number } => {
            const item = c as { type: string; side?: string };
            return item.type === 'border' && item.side === 'top';
          })
          .map((c) => c.y + c.height),
      );
    // The last row on page 1 is closed: its bottom edge is level with the
    // bottom of the last box drawn there, not a rule short of it.
    expect(lowestBottom(laid.pages[0]!)).toBeCloseTo(lowestTop(laid.pages[0]!), 1);
  });
});

describe('vertical alignment (§18.8.1)', () => {
  it('sits a cell at the bottom of a box taller than its text', () => {
    // Excel's default is bottom, and LibreOffice does the same; we drew every
    // cell against the top, which on tdf144642's 28.35pt rows floated the text
    // a line-height above where both of them put it.
    const items = placed(
      buildXlsx({ rows: [['tall']], rowHeights: [{ row: 0, heightPt: 60, customHeight: true }] }),
    );
    const y = at(items, 'tall').y;
    const flat = placed(buildXlsx({ rows: [['tall']] }));
    // The 60pt row pushes its text down by the slack over the 15pt default.
    expect(y - at(flat, 'tall').y).toBeCloseTo(45, 0);
  });

  it('honours an explicit vertical="top"', () => {
    const items = placed(
      buildXlsx({
        rows: [[{ value: 'tall', styleIndex: 1 }]],
        rowHeights: [{ row: 0, heightPt: 60, customHeight: true }],
        stylesXml:
          `<fonts count="1"><font/></fonts><fills count="1"><fill/></fills>` +
          `<borders count="1"><border/></borders>` +
          `<cellXfs count="2"><xf/><xf applyAlignment="1"><alignment vertical="top"/></xf></cellXfs>`,
      }),
    );
    // The same 60pt row without the override drops its text to the foot of the
    // box; with it the text stays at the head, 45pt of slack higher.
    const bottom = placed(
      buildXlsx({ rows: [['tall']], rowHeights: [{ row: 0, heightPt: 60, customHeight: true }] }),
    );
    // 60pt of box less the line the text occupies — the slack it drops through.
    expect(at(bottom, 'tall').y - at(items, 'tall').y).toBeCloseTo(46.8, 0);
  });
});

describe('paper size from the printer settings part', () => {
  it('takes the paper a <pageSetup> leaves to its printerSettings DEVMODE', () => {
    // `<pageSetup>` naming no paperSize does not mean "the default": Excel
    // records the print dialog's choice in the related part, and LibreOffice
    // reads it. simple-monthly-budget.xlsx prints on Letter that way.
    const bytes = new Uint8Array(readFileSync('tests/fixtures/real/simple-monthly-budget.xlsx'));
    const setup = readXlsxToSheetDoc(bytes).sheets[0]!.grid.pageSetup;
    expect(setup?.paperSize).toBe(1); // §18.3.1.63 — 1 is Letter
    expect(setup?.orientation).toBe('landscape');
    // The relationship id is a spelling, not a fact: it must not survive into
    // the model, or two dialects of one workbook parse differently.
    expect(setup).not.toHaveProperty('printerSettingsRelId');
  });

  it('ignores a printerSettings part that is not a DEVMODE', () => {
    expect(parsePrinterSettings(new Uint8Array(4))).toEqual({});
    expect(parsePrinterSettings(new Uint8Array(200))).toEqual({});
  });
});
