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
import { columnTwips } from '@/excel/print-model';
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

  it('does not pad a declared width a second time (§18.3.1.13)', () => {
    // The stored `width` ALREADY carries the 5px: the forward formula is
    // `Truncate([chars × MDW + 5] / MDW × 256) / 256`, so the padding is inside
    // the number. Adding it again on the way out made every explicitly sized
    // column 5px too wide, which is why our columns measured a fixed 3.75pt
    // over Excel's however narrow they were.
    const xlsx = buildXlsx({
      rows: [['A', 'B']],
      columns: [{ min: 1, max: 1, widthChars: 12 }],
    });
    const items = placed(xlsx);
    // 12 × 7px = 84px = 63pt.
    expect(at(items, 'B').x - at(items, 'A').x).toBeCloseTo(63, 1);
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

  it('does not grow a row whose height the author pinned', () => {
    // §18.3.1.73 customHeight="1" — Excel does not auto-fit such a row, it
    // clips. Growing it is not a harmless safety margin: the extra height moves
    // every row after it and changes where the page breaks.
    // simple-monthly-budget.xlsx pins all 23 of its rows, their sum fits its
    // page with 5pt to spare, and one row carrying a 20pt "62%" grew past its
    // pin and pushed the last row — and the chart anchored beside it — onto a
    // second page.
    const big = `<fonts count="2"><font><sz val="11"/></font><font><sz val="36"/></font></fonts>
      <fills count="1"><fill/></fills><borders count="1"><border/></borders>
      <cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>`;
    const rows = [[{ value: 'huge', styleIndex: 1 }], ['after']];
    const pinned = placed(
      buildXlsx({
        rows,
        stylesXml: big,
        rowHeights: [{ row: 0, heightPt: 15, customHeight: true }],
      }),
    );
    const loose = placed(
      buildXlsx({
        rows,
        stylesXml: big,
        rowHeights: [{ row: 0, heightPt: 15, customHeight: false }],
      }),
    );
    // Pinned, the row keeps its 15pt and the row below starts where 15pt says.
    // Unpinned, the 36pt line grows it and pushes everything after it down.
    expect(at(loose, 'after').y - at(pinned, 'after').y).toBeGreaterThan(20);
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

  it('wraps a cell whose word is wider than the column, instead of one long line', () => {
    // Knuth-Plass scores a line by how badly it fits, and a line that cannot
    // stretch — one word alone in a narrow cell — is infinitely loose, while an
    // overfull line clamps to a small fixed badness. So the total-fit answer for
    // a wrapping cell holding "SELF EMPLOYED" in a column that fits neither word
    // was ONE overfull line spilling across the cells beside it, where breaking
    // after "SELF" costs nothing but white space.
    const stylesXml =
      `<fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill/></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellXfs count="2"><xf/><xf applyAlignment="1"><alignment wrapText="1"/></xf></cellXfs>`;
    const cell = (text: string): Array<string> =>
      placed(
        buildXlsx({
          rows: [[{ value: text, styleIndex: 1 }, 'x']],
          columns: [{ min: 1, max: 2, widthChars: 8 }],
          stylesXml,
        }),
      ).map((i) => i.text);
    expect(cell('SELF EMPLOYED')).toEqual(expect.arrayContaining(['SELF', 'EMPLOYED']));
    // A paragraph that fits keeps the total-fit break it had.
    expect(cell('alpha beta gamma')).toEqual(expect.arrayContaining(['alpha', 'beta', 'gamma']));
    // And a single word wider than the column has nowhere to break: it stays.
    expect(cell('supercalifragilistic')).toContain('supercalifragilistic');
  });

  it('fills a number that cannot fit its column with #, never truncating it', () => {
    // §18.8.1 — a truncated number is a DIFFERENT number: "4/30/201" reads as a
    // date and is not the one in the cell. Excel and LibreOffice fill such a
    // cell with `#`, which says "widen me" and cannot be misread.
    // column-style-autofilter.xlsx has 3196 of them.
    const stylesXml =
      `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000000"/></numFmts>` +
      `<fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill/></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellXfs count="2"><xf/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs>`;
    const drawn = (widthChars: number): Array<string> =>
      placed(
        buildXlsx({
          rows: [[{ value: 1234.567891, styleIndex: 1 }, 'x']],
          columns: [{ min: 1, max: 2, widthChars }],
          stylesXml,
        }),
      ).map((i) => i.text);
    // Two characters of column for an eleven-character number.
    expect(drawn(2).some((t) => /^#+$/.test(t))).toBe(true);
    // Room enough, and the number itself is drawn.
    expect(drawn(20)).toContain('1234.567891');
    // A value that overruns by a hair is drawn and left to clip: filling the
    // cell with `#` erases it, so the test has to be sure. A ten-character date
    // in a column measured for about nine — which is what a default-width
    // column holds — keeps its date (forum-mso-de-104083.xlsx hashed 439 cells
    // that way, where the reference hashes none of them).
    const dates = placed(
      buildXlsx({
        rows: [[{ value: 42324, styleIndex: 2 }, 'x']],
        columns: [{ min: 1, max: 2, widthChars: 8.43 }],
        stylesXml:
          `<fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill/></fills>` +
          `<borders count="1"><border/></borders>` +
          `<cellXfs count="3"><xf/><xf/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs>`,
      }),
    ).map((i) => i.text);
    expect(dates.some((t) => /^#+$/.test(t))).toBe(false);

    // Text is exempt — a clipped word is still recognisably that word.
    const text = placed(
      buildXlsx({
        rows: [['abcdefghijklmnop', 'x']],
        columns: [{ min: 1, max: 2, widthChars: 2 }],
        stylesXml,
      }),
    ).map((i) => i.text);
    expect(text.some((t) => /^#+$/.test(t))).toBe(false);
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
    //
    // §18.3.1.81 pads it TWICE: `defaultColWidth` "includes margin padding and
    // extra padding for gridlines" and `baseColWidth` explicitly does not, so
    // deriving one from the other adds the 5px once, and rendering it adds the
    // 5px again. 47668.xlsx declares baseColWidth="10" and caches its picture
    // at 768pt over 12 columns plus 48pt — 60pt a column, not 56.25.
    const wide = placed(buildXlsx({ rows: [['A', 'B']], baseColWidthChars: 20 }));
    const narrow = placed(buildXlsx({ rows: [['A', 'B']] }));
    const pitch = (items: Array<PlacedText>): number => at(items, 'B').x - at(items, 'A').x;
    expect(pitch(wide)).toBeCloseTo(20 * 5.25 + 3.75 * 2, 1);
    // baseColWidth 10 → 60pt, the number this file's own cached extent implies.
    expect(pitch(placed(buildXlsx({ rows: [['A', 'B']], baseColWidthChars: 10 })))).toBeCloseTo(
      60,
      1,
    );
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
        columns: [{ min: 1, max: 4, widthChars: 15 }],
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
    // Excel's own unit: 40 × 7px, the padding already inside the stored width.
    expect(width(named)).toBeCloseTo(40 * 5.25, 1);
    // The reader's face is wider than Calibri's, so the column is too.
    expect(width(anonymous)).toBeGreaterThan(width(named));
  });

  it('rounds a sub-character width by its own formula, as Excel documents', () => {
    // Excel gives a column narrower than one character its own conversion,
    // `px = Trunc(width × MDW + 0.5)`, instead of the 256ths formula it uses
    // above one — and it matters because a constant error is ruinous at that
    // scale: tdf118668.xlsx rules a form over 168 columns of 0.855 characters,
    // and 3.75pt of stray padding nearly doubled a 4.5pt column. The sheet came
    // out 1384pt wide against the reference's 754 and took two pages.
    expect(columnTwips(0.855, 105)).toBe(90); // Trunc(0.855×7 + 0.5) = 6px
    expect(columnTwips(0.1, 105)).toBe(15); // Trunc(0.7 + 0.5) = 1px
    // At one character and above, §18.3.1.13's own inverse — and the number it
    // is anchored on: Excel's default column is stored 9.140625 and is 64px.
    expect(columnTwips(1, 105)).toBe(105);
    expect(columnTwips(9.140625, 105)).toBe(960);
  });
});

describe('cell indent (§18.8.1)', () => {
  it('offsets the text inside the cell, not just the model', () => {
    // The projection has always put `indent` on the paragraph; the table path
    // drew every cell at its padding and ignored it, so 45544.xlsx printed five
    // indented rows flush with the heading above them.
    const stylesXml =
      `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellXfs count="2">` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">` +
      `<alignment horizontal="left" indent="1"/></xf>` +
      `</cellXfs>`;
    const items = placed(
      buildXlsx({
        rows: [['flush'], [{ value: 'indented', styleIndex: 1 }]],
        columns: [{ min: 1, max: 1, widthChars: 30 }],
        stylesXml,
      }),
    );
    expect(at(items, 'indented').x).toBeGreaterThan(at(items, 'flush').x + 8);
  });
});

/** Every fill the layout painted, with its box. */
const fills = (
  xlsx: Uint8Array,
): Array<{ y: number; height: number; width: number; hex: string }> => {
  const flow = Ream.parse(xlsx).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  const out: Array<{ y: number; height: number; width: number; hex: string }> = [];
  for (const page of laid.pages) {
    for (const command of page.commands) {
      if (command.type !== 'fill') continue;
      const f = command as unknown as {
        y: number;
        height: number;
        width: number;
        fillColorHex: string;
      };
      out.push({ y: f.y, height: f.height, width: f.width, hex: f.fillColorHex });
    }
  }
  return out;
};

describe('a vertical merge is ONE box (§18.3.1.55)', () => {
  const green = { value: 'M', styleIndex: 1 };
  const filled = (
    rows: ReadonlyArray<ReadonlyArray<typeof green | string>>,
    mergeRefs: Array<string>,
  ): Uint8Array =>
    buildXlsx({
      rows,
      mergeRefs,
      stylesXml:
        `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="3"><fill><patternFill patternType="none"/></fill>` +
        `<fill><patternFill patternType="gray125"/></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>` +
        `</fills><borders count="1"><border/></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
        `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs>`,
    });

  it('paints its whole height, not only its first row', () => {
    // The continuation rows arrived with no paint of their own and were skipped
    // by the fill pass, so a two-row merge painted one row's worth: 50299.xlsx
    // drew 13.6pt of grey where both references draw 27.
    const greenHeight = (xlsx: Uint8Array): number =>
      fills(xlsx)
        .filter((f) => f.hex === '00FF00')
        .reduce((sum, f) => sum + f.height, 0);
    const oneRow = greenHeight(filled([[green], ['after']], []));
    const merged = greenHeight(filled([[green], [{ ...green, value: '' }], ['after']], ['A1:A2']));
    expect(oneRow).toBeGreaterThan(0);
    expect(merged).toBeGreaterThan(oneRow * 1.8);
  });

  it('sits its text at the bottom of the merged box, not of its first row', () => {
    // A spreadsheet cell sits at the bottom of a box taller than its content,
    // and a merge's box is every row it spans — Excel and LibreOffice both put
    // the label on the merge's last row.
    const items = placed(
      buildXlsx({ rows: [['a', 'b'], [null, 'c']], mergeRefs: ['A1:A2'] }),
    );
    expect(at(items, 'a').y).toBeCloseTo(at(items, 'c').y, 1);
  });

  it('cuts its text at its own edge instead of running it across the row', () => {
    // Excluded from the clip pass — which is really the neighbour-claiming rule
    // a merge must not follow, with the clip every cell needs bolted to it —
    // 50299.xlsx ran `$R[]{A,hideDuplicate=true}` out of a 48pt merge and over
    // the three cells beside it.
    const long = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    const items = placed(
      buildXlsx({
        rows: [[long, null, 'right']],
        mergeRefs: ['A1:B1'],
        columns: [{ min: 1, max: 3, widthChars: 6 }],
      }),
    );
    const drawn = items.find((i) => long.startsWith(i.text))?.text ?? '';
    expect(drawn.length).toBeGreaterThan(0);
    expect(drawn.length).toBeLessThan(long.length);
    expect(at(items, 'right').x).toBeGreaterThan(at(items, drawn).x);
  });
});

describe('the used range', () => {
  it('reaches a value-less cell that PAINTS something (50299)', () => {
    // Column H of this sheet is ten cells with a solid fill each and no value —
    // a colour key beside the data. Bounding the grid to cells that carry
    // content dropped the column before it was ever materialised, and the whole
    // band was missing from a page both references print it on.
    const flow = Ream.parse(new Uint8Array(readFileSync('tests/fixtures/real/50299.xlsx'))).flow;
    const shadings = new Set<string>();
    for (const el of flow.body) {
      if (el.kind !== 'table') continue;
      for (const row of el.table.rows) {
        for (const cell of row.cells) {
          const hex = cell.properties.shading?.colorHex;
          if (hex) shadings.add(hex);
        }
      }
    }
    // The ten swatches, in fills[2..11] order.
    for (const hex of ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '7030A0']) {
      expect(shadings).toContain(hex);
    }
  });

  it('does not let a fill CREATE one (bnc762542)', () => {
    // The same file that made this rule get written twice and reverted twice:
    // every value on it lives in a styled-empty cell, so a fill that could seed
    // the used range gives a grid to a sheet that prints nothing but a callout.
    // Painting extends a range; it never opens one.
    const bytes = new Uint8Array(readFileSync('tests/fixtures/real/bnc762542.xlsx'));
    expect(pageCount(bytes)).toBe(1);
  });
});

describe('a sheet with nothing on it', () => {
  it('is left out of the print (tdf115159)', () => {
    // Two untouched tabs beside one sheet of data. Each sheet is its own
    // section — the data sheet keeps its printer's margins and the blank ones
    // the default — and a section whose geometry differs starts a page, so the
    // empty tabs printed one. Excel and LibreOffice both leave them out.
    const bytes = new Uint8Array(readFileSync('tests/fixtures/real/tdf115159.xlsx'));
    expect(readXlsxToSheetDoc(bytes).sheets).toHaveLength(3);
    expect(pageCount(bytes)).toBe(1);
  });

  it('is left out even when it carries a header (47737)', () => {
    // A header is furniture, not content. 47737's "Sheet 2" is eighteen rows of
    // cells that carry a style and no value, and one `<oddHeader>`: Excel
    // answers "We didn't find anything to print" and LibreOffice prints no page
    // for it either. Counting the header as content printed it.
    const bytes = new Uint8Array(readFileSync('tests/fixtures/real/47737.xlsx'));
    expect(readXlsxToSheetDoc(bytes).sheets).toHaveLength(2);
    expect(placed(bytes).map((p) => p.text)).not.toContain('Agency Footnotes');
  });

  it('still prints when it is the only sheet there is', () => {
    // A workbook of nothing but empty sheets is still a document.
    expect(
      pageCount(
        buildXlsx({
          sheets: [
            { name: 'A', rows: [] },
            { name: 'B', rows: [] },
          ],
        }),
      ),
    ).toBe(1);
  });
});

describe('a row pinned to a height (§18.3.1.73 customHeight)', () => {
  it('cuts what does not fit instead of running it over the row below', () => {
    // Excel and Word both clip a cell to its row. Drawing the overflow put
    // tdf118668.xlsx's wrapped labels on top of the next row's text — a form of
    // 9.75pt rows, every one of them pinned.
    const stylesXml =
      `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
      `<borders count="1"><border/></borders>` +
      `<cellXfs count="2">` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1">` +
      `<alignment wrapText="1"/></xf>` +
      `</cellXfs>`;
    const sheet = (pinned: boolean): Uint8Array =>
      buildXlsx({
        rows: [[{ value: 'one two three four five six seven eight', styleIndex: 1 }], ['next']],
        columns: [{ min: 1, max: 1, widthChars: 6 }],
        rowHeights: [{ row: 0, heightPt: 12, customHeight: pinned }],
        stylesXml,
      });
    // Pinned: only the lines that fit are drawn. Unpinned: the row grows and
    // keeps every one of them.
    const pinnedLines = placed(sheet(true)).length;
    const grownLines = placed(sheet(false)).length;
    expect(pinnedLines).toBeLessThan(grownLines);
    expect(pinnedLines).toBeGreaterThan(0);
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

describe('a cell that is present but empty (§18.3.1.4)', () => {
  it('draws nothing where a styled blank sits', () => {
    // `t` defaults to "n", so a cell written only to carry a style arrives
    // typed numeric with an empty value — and `Number('')` is 0, not NaN.
    // Rounding General to the column width printed that zero: every styled
    // blank in invalid_ext_data_validation.xlsx grew a `0`, and one of them
    // took a page of its own.
    // The blank has to earn its place in the grid: style 1 paints it green, so
    // it survives the trim that removes cells drawing nothing at all.
    const stylesXml = `
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
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
    const items = placed(
      buildXlsx({
        rows: [
          ['left', { value: null, styleIndex: 1 }, 'right'],
          ['a', 'b', 'c'],
        ],
        stylesXml,
      }),
    );
    expect(items.map((i) => i.text).sort()).toEqual(['a', 'b', 'c', 'left', 'right']);
  });
});

describe('a column of one colour is painted once', () => {
  it('joins the rows into a single rectangle', () => {
    // Runs ACROSS a row were already merged; down the page each row painted its
    // own slice and leaned on a 0.07pt bleed to hide the join — a fifth of a
    // device pixel at 300 DPI, so the boundary pixel took partial coverage from
    // each side and came out pale. 51710.xlsx paints its column A grey down 46
    // pages and showed a rung at every row boundary.
    const green = { value: 'x', styleIndex: 1 };
    const xlsx = buildXlsx({
      rows: [[green], [green], [green], [green]],
      stylesXml:
        `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="3"><fill><patternFill patternType="none"/></fill>` +
        `<fill><patternFill patternType="gray125"/></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>` +
        `</fills><borders count="1"><border/></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
        `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs>`,
    });
    const painted = fills(xlsx).filter((f) => f.hex === '00FF00');
    expect(painted).toHaveLength(1);
    expect(painted[0]!.height).toBeGreaterThan(4 * 14);
  });
});

describe('overflow claims a neighbour only when the text needs it', () => {
  it('leaves a short centred cell in its own column, fill and all', () => {
    // There was no width test on the rightward path at all, so any left- or
    // centre-aligned string claimed every empty neighbour to the end of its
    // band — and since overflow is modelled as a colSpan, the cell's fill and
    // its centring went with them. 54524.xlsx paints "X", 6.9pt of text in a
    // 49.5pt column, across two columns; 54206.xlsx runs a header's fill into
    // an empty column K. The mirror path for right-aligned cells has always
    // had this predicate.
    const xlsx = buildXlsx({
      rows: [[{ value: 'X', styleIndex: 1 }], [null, 'anchor']],
      columns: [{ min: 1, max: 3, widthChars: 12 }],
      stylesXml:
        `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="3"><fill><patternFill patternType="none"/></fill>` +
        `<fill><patternFill patternType="gray125"/></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>` +
        `</fills><borders count="1"><border/></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
        `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1" applyAlignment="1">` +
        `<alignment horizontal="center"/></xf></cellXfs>`,
    });
    const green = fills(xlsx).filter((f) => f.hex === '00FF00');
    expect(green).toHaveLength(1);
    // One column of 12 characters: 84px = 63pt, not two.
    expect(green[0]!.width).toBeLessThan(70);
  });

  it('still lets a cell that genuinely overruns take the room', () => {
    const long = 'a label far wider than the column it sits in';
    const items = placed(
      buildXlsx({
        rows: [[long], [null, 'anchor']],
        columns: [{ min: 1, max: 3, widthChars: 8 }],
      }),
    );
    // Drawn on ONE line — the span it claimed is what keeps it off three.
    expect(items.filter((i) => long.startsWith(i.text))).toHaveLength(1);
  });
});

describe('the column-width unit is the normal style font’s digit (§18.3.1.13)', () => {
  const pitchWith = (font: string, sizePt: number): number => {
    const items = placed(
      buildXlsx({
        rows: [['A', 'B']],
        columns: [{ min: 1, max: 1, widthChars: 10 }],
        stylesXml:
          `<fonts count="1"><font><sz val="${sizePt}"/><name val="${font}"/></font></fonts>` +
          `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
          `<borders count="1"><border/></borders>` +
          `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>`,
      }),
    );
    return at(items, 'B').x - at(items, 'A').x;
  };

  it('is 7px for Calibri 11 and Arial 10, and 8px for Arial 11', () => {
    // "…the maximum digit width of the numbers 0, 1, 2, …, 9 AS RENDERED IN THE
    // NORMAL STYLE'S FONT". The test used to be whether the file named a font
    // at all, and every named one got Excel's 7px — right for the two faces the
    // note cited, wrong for anything else. 55850.xlsx and 55927.xlsx are Arial
    // ELEVEN, 8px, so their columns came out 14% narrow and each lost the last
    // character of every date it holds.
    expect(pitchWith('Calibri', 11)).toBeCloseTo(10 * 5.25, 1); // 70px
    expect(pitchWith('Arial', 10)).toBeCloseTo(10 * 5.25, 1);
    expect(pitchWith('Arial', 11)).toBeCloseTo(10 * 6.0, 1); // 80px
  });

  it('keeps 7px for a face it cannot measure', () => {
    expect(pitchWith('Some Unknown Face', 11)).toBeCloseTo(10 * 5.25, 1);
  });
});

describe('an overflow span is not a paint box', () => {
  it('runs the text across a neighbour without painting it', () => {
    // Excel runs the TEXT of an over-long cell over the empty cell beside it
    // and paints nothing there. Modelled as one colSpan, the fill went with the
    // text: 54436.xlsx ran its pivot header's blue a whole column past the
    // pivot, so its header band and its total band — the same two columns —
    // came out different widths.
    const green = { value: 'a label wider than its own column', styleIndex: 1 };
    const xlsx = buildXlsx({
      rows: [[green], [null, 'anchor']],
      columns: [{ min: 1, max: 3, widthChars: 8 }],
      stylesXml:
        `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
        `<fills count="3"><fill><patternFill patternType="none"/></fill>` +
        `<fill><patternFill patternType="gray125"/></fill>` +
        `<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>` +
        `</fills><borders count="1"><border/></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
        `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/></cellXfs>`,
    });
    // One column of 8 characters is 56px = 42pt; the text spans more than that.
    const painted = fills(xlsx).filter((f) => f.hex === '00FF00');
    expect(painted).toHaveLength(1);
    expect(painted[0]!.width).toBeLessThan(50);
    // …and the text is still one line, not wrapped into its own column.
    expect(placed(xlsx).filter((i) => green.value.startsWith(i.text))).toHaveLength(1);
  });
});

describe('a cell cuts its text mid-glyph (ISO 32000-1 §8.5.4)', () => {
  it('keeps the glyph that straddles the edge and clips the line to the cell', () => {
    // Excel and LibreOffice paint the straddling glyph and cut it through its
    // bowl. Requiring its right edge to fit dropped it whole — three characters
    // on 56511.xlsx, and on 56574.xlsx a formatted date came out `5/29/201`,
    // which is not a shorter date but a different one.
    const long = 'dlgkdflgdfjkl';
    const items = placed(
      buildXlsx({
        rows: [[long, long]],
        columns: [{ min: 1, max: 2, widthChars: 10 }],
      }),
    );
    const drawn = items.find((i) => long.startsWith(i.text))?.text ?? '';
    expect(drawn.length).toBeGreaterThan(0);
    // The cut keeps one more character than a whole-glyph fit would: the
    // emitter's clip rectangle is what makes it safe to draw.
    expect(long.startsWith(drawn)).toBe(true);
  });
});

describe('the cell cut reaches every writer', () => {
  const overrun = buildXlsx({
    rows: [['dlgkdflgdfjkl', 'dlgkdflgdfjkl']],
    columns: [{ min: 1, max: 2, widthChars: 10 }],
  });

  it('clips the line in SVG the way the PDF emitter does', async () => {
    const svg = new TextDecoder().decode(
      await Ream.parse(overrun).convert('svg', { fonts: FONTS }),
    );
    expect(svg).toContain('<clipPath');
    expect(svg).toContain('clip-path="url(#');
  });

  it('tells the browser not to wrap or spill in HTML (§18.8.1)', async () => {
    // HTML renders the document model rather than a laid-out page, so the cut
    // is the browser's to make — it only has to be told.
    const html = new TextDecoder().decode(await Ream.parse(overrun).convert('html'));
    expect(html).toContain('white-space:nowrap');
    expect(html).toContain('overflow:hidden');
  });
});
