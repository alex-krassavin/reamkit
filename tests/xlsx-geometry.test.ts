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
import { readXlsx } from '@/excel/xlsx-reader';

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
    expect(ys[1]! - ys[0]!).toBeCloseTo(15, 1); // r0's own height: the default
    expect(ys[2]! - ys[1]!).toBeCloseTo(40, 1); // r1 was pinned to 40pt
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
