// §20.5 — a sheet's drawings on paper (TableOverlay). A drawing is anchored to
// cells, so it prints with them: each page shows the part of every drawing
// that lies over the rows and columns it prints, cut off where they end.
// Measured against Excel's own PDF of probe workbooks (2026-10-01): a chart
// across a page break prints in two pieces, a drawing in the print titles on
// every page, nothing of what lies outside the print area, a hidden column or
// row takes no room, and the half of a frame's line that falls past the grid's
// edge is cut off with it.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { LaidOutPage, ShapeItem } from '@/layout/page-doc';
import { FontRegistry } from '@/core/font';
import { Ream } from '@/core/converter/ream';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { writeSvg } from '@/svg/svg-writer';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { projectSheetDoc } from '@/excel/sheet-to-flow';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/** One value per row, as many rows as asked for. */
const rows = (n: number): Array<Array<string>> =>
  Array.from({ length: n }, (_, i) => [`row ${i + 1}`]);

function pages(xlsx: Uint8Array): ReadonlyArray<LaidOutPage> {
  const flow = Ream.parse(xlsx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  }).pages;
}

/** The shape items of a sheet's drawings on a page: they stand over its text. */
const drawn = (page: LaidOutPage): Array<ShapeItem> =>
  page.commands.filter((c): c is ShapeItem => c.type === 'shape' && c.over === true);

/** A shape item's top edge, y-down, for a drawing `heightPt` tall. */
const topOf = (item: ShapeItem, heightPt: number): number => item.shape.transform[5] - heightPt;

describe('a sheet drawing on paper (TableOverlay)', () => {
  it('prints in pieces across a page break, cut off where each page ends', () => {
    // 80 rows of 15pt run onto a second page; the shape spans rows 41 to 60.
    const laid = pages(
      buildXlsx({ rows: rows(80), sheetShape: { anchor: { from: [1, 40], to: [3, 60] } } }),
    );
    expect(laid.length).toBe(2);
    const [first] = drawn(laid[0]!);
    const [second] = drawn(laid[1]!);
    if (!first || !second) throw new Error('the shape is not on both pages');
    const height = 20 * 15;
    const w1 = first.window!;
    const w2 = second.window!;
    // Page one shows it down to the page's last row, page two the rest from
    // its first: what the first window shows is what the second starts past.
    const shownFirst = w1.y + w1.height - topOf(first, height);
    expect(shownFirst).toBeGreaterThan(0);
    expect(shownFirst).toBeLessThan(height);
    expect(topOf(second, height)).toBeCloseTo(w2.y - shownFirst, 1);
  });

  it('prints in the print titles on every page that repeats them', () => {
    const laid = pages(
      buildXlsx({
        rows: rows(120),
        definedNames: [{ name: '_xlnm.Print_Titles', localSheetId: 0, value: 'Sheet1!$1:$2' }],
        sheetShape: { anchor: { from: [1, 0], to: [3, 2] } },
      }),
    );
    expect(laid.length).toBeGreaterThan(2);
    for (const page of laid) expect(drawn(page).length).toBeGreaterThan(0);
  });

  it('is cut off on an SVG page as on a PDF one', () => {
    const flow = Ream.parse(
      buildXlsx({ rows: rows(80), sheetShape: { anchor: { from: [1, 40], to: [3, 60] } } }),
    ).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const svg = new TextDecoder().decode(writeSvg(laid).bytes);
    // One window a page, each the clip of the piece drawn there.
    expect(svg.match(/<clipPath id="win\d+"><rect /g)?.length).toBe(2);
    expect(svg).toMatch(/<g clip-path="url\(#win\d+\)">/);
  });

  it('prints nothing of what lies outside the print area, and cuts what straddles it', () => {
    const book = (anchor: { from: [number, number]; to: [number, number] }): Uint8Array =>
      buildXlsx({
        rows: Array.from({ length: 40 }, (_row, r) =>
          Array.from({ length: 8 }, (_cell, c) => r * 8 + c),
        ),
        definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 0, value: 'Sheet1!$C$6:$H$30' }],
        sheetShape: { anchor },
      });
    // A3:E10 runs out of the area's top left corner: drawn, through a window
    // that starts where the area does.
    const [straddling] = drawn(pages(book({ from: [0, 2], to: [4, 9] }))[0]!);
    if (!straddling) throw new Error('the straddling shape is not drawn');
    expect(straddling.shape.transform[4]).toBeLessThan(straddling.window!.x);
    expect(topOf(straddling, 7 * 15)).toBeLessThan(straddling.window!.y);
    // J1:K3 is nowhere near it.
    expect(drawn(pages(book({ from: [9, 0], to: [11, 3] }))[0]!)).toHaveLength(0);
  });

  it('takes no room for a hidden column or row', () => {
    const doc = readXlsxToSheetDoc(
      buildXlsx({
        rows: rows(10),
        columns: [{ min: 3, max: 3, widthChars: 9.140625, hidden: true }],
        rowHeights: [
          { row: 1, heightPt: 15, hidden: true },
          { row: 2, heightPt: 15, hidden: true },
        ],
        sheetShape: { anchor: { from: [1, 4], to: [4, 6] } },
      }),
    );
    const shape = doc.sheets[0]!.shapes![0]!;
    // B to E over a hidden C is B and D; row 5 under two hidden rows stands
    // two rows down.
    expect(shape.width).toBeCloseTo(2 * 48, 1);
    expect(shape.float!.posH!.offsetPt).toBeCloseTo(48, 1);
    expect(shape.float!.posV!.offsetPt).toBeCloseTo(2 * 15, 1);
  });

  it('is cut at the grid edge it stands on, its line with it', () => {
    // A shape at A1 with a 4.5pt line: Excel prints the outer half of its top
    // and left line nowhere, since it falls outside the sheet's first cell.
    const [item] = drawn(
      pages(buildXlsx({ rows: rows(5), sheetShape: { anchor: { from: [0, 0], to: [2, 3] } } }))[0]!,
    );
    if (!item) throw new Error('no shape drawn');
    expect(item.shape.transform[4]).toBeCloseTo(item.window!.x, 1);
    expect(topOf(item, 3 * 15)).toBeCloseTo(item.window!.y, 1);
  });

  it('prints the page a drawing reaches onto, though nothing else is on it', () => {
    // Twenty rows, a manual break after them, and a shape that runs on past
    // the break: Excel prints a second page for its tail — and for a shape
    // that ends exactly on the break as well, which is the half of its line.
    for (const bottom of [21, 20]) {
      const laid = pages(
        buildXlsx({
          rows: rows(20),
          rowBreaks: [20],
          sheetShape: { anchor: { from: [1, 3], to: [3, bottom] } },
        }),
      );
      expect(laid.length).toBe(2);
      expect(drawn(laid[1]!).length).toBeGreaterThan(0);
    }
  });

  it('starts each column band on a page of its own when the headings print', () => {
    // The letters lead each band, and the break that starts a band is on them.
    const doc = readXlsxToSheetDoc(
      buildXlsx({
        rows: [Array.from({ length: 6 }, (_, c) => `c${c}`)],
        columns: [{ min: 1, max: 6, widthChars: 40 }],
        printOptions: { headings: true },
      }),
    );
    expect(projectSheetDoc(doc).body.filter((el) => el.kind === 'table')).toHaveLength(3);
    expect(
      pages(
        buildXlsx({
          rows: [Array.from({ length: 6 }, (_, c) => `c${c}`)],
          columns: [{ min: 1, max: 6, widthChars: 40 }],
          printOptions: { headings: true },
        }),
      ),
    ).toHaveLength(3);
  });
});
