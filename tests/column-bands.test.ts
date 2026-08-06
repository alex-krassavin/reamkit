// E-SHEET SE1 — wide-sheet column-band pagination. A sheet wider than the page
// (at 100% scale) paginates ACROSS columns into bands ("down, then over") instead
// of being squeezed onto one page width; fit-to-page / scaled sheets keep the
// uniform-shrink path. The pure band maths is unit-tested; the projection +
// layout are exercised end-to-end via the band-table and page counts.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildTinyPng } from './fixtures/build-png';
import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { computeColumnBands } from '@/excel/column-bands';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// How many band tables the projection emits for a single-sheet workbook.
const bandCount = (xlsx: Uint8Array): number =>
  Ream.parse(xlsx).flow.body.filter((el) => el.kind === 'table').length;

const pageCount = (xlsx: Uint8Array): number => {
  const flow = Ream.parse(xlsx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  }).pages.length;
};

const grid = (rows: number, cols: number): Array<Array<number>> =>
  Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => r * cols + c + 1));

// 6 columns of 40 chars ≈ 4200 twips each = 25 200 twips, well past the default
// A4 content width (11906 − 2×1440 = 9026 twips) → 3 bands of 2 columns.
const wideCols = [{ min: 1, max: 6, widthChars: 40 }];

describe('computeColumnBands (E-SHEET SE1)', () => {
  it('keeps a sheet that fits in a single band', () => {
    expect(computeColumnBands([100, 100, 100], 1000, new Set())).toEqual([{ start: 0, end: 2 }]);
  });

  it('greedily splits columns that overflow the content width', () => {
    expect(computeColumnBands([400, 400, 400, 400], 1000, new Set())).toEqual([
      { start: 0, end: 1 },
      { start: 2, end: 3 },
    ]);
  });

  it('starts a new band at a manual column break even when the rest fits', () => {
    expect(computeColumnBands([100, 100, 100, 100], 1000, new Set([2]))).toEqual([
      { start: 0, end: 1 },
      { start: 2, end: 3 },
    ]);
  });

  it('gives a single over-wide column its own band', () => {
    expect(computeColumnBands([2000, 100], 1000, new Set())).toEqual([
      { start: 0, end: 0 },
      { start: 1, end: 1 },
    ]);
  });

  it('returns nothing for an empty grid', () => {
    expect(computeColumnBands([], 1000, new Set())).toEqual([]);
  });
});

describe('wide-sheet column-band pagination (E-SHEET SE1)', () => {
  it('paginates a wide unscaled sheet across columns (down, then over)', () => {
    const xlsx = buildXlsx({ rows: grid(3, 6), columns: wideCols });
    expect(bandCount(xlsx)).toBe(3); // 3 bands of 2 columns
    expect(pageCount(xlsx)).toBe(3); // each band starts on its own page
  });

  it('does not band a sheet that already fits the page', () => {
    const xlsx = buildXlsx({ rows: grid(3, 2), columns: [{ min: 1, max: 2, widthChars: 5 }] });
    expect(bandCount(xlsx)).toBe(1);
    expect(pageCount(xlsx)).toBe(1);
  });

  it('does not band a fit-to-page sheet (uniform shrink instead)', () => {
    const xlsx = buildXlsx({
      rows: grid(3, 6),
      columns: wideCols,
      fitToPage: true,
      pageSetup: { fitToWidth: 1, fitToHeight: 1 },
    });
    expect(bandCount(xlsx)).toBe(1);
  });

  it('bands a fitToWidth=2 sheet across pages on its SCALED widths (SE-T)', () => {
    // 6×4200-twip columns scaled to fit two 9026-twip pages → ~3 columns per band.
    const xlsx = buildXlsx({
      rows: grid(3, 6),
      columns: wideCols,
      fitToPage: true,
      pageSetup: { fitToWidth: 2, fitToHeight: 1 },
    });
    expect(bandCount(xlsx)).toBe(2);
    expect(pageCount(xlsx)).toBe(2);
  });

  it('bands a sheet that is only ZOOMED, on its scaled widths (§18.3.1.63)', () => {
    // `<pageSetup scale>` and fit-to-page are not the same instruction. Fitting
    // is a promise about the page count, so the grid may be squeezed onto one
    // width; a scale is a plain zoom, and Excel paginates after applying it.
    // Treating a scaled sheet as fitted left 47737.xlsx — `scale="63"`, still
    // 583pt of columns on a 487pt page — as one over-wide table, and its last
    // column ran off the paper.
    const xlsx = buildXlsx({ rows: grid(3, 6), columns: wideCols, pageSetup: { scale: 50 } });
    expect(bandCount(xlsx)).toBe(2); // 6×2100 scaled twips over a 9026-twip page
    expect(pageCount(xlsx)).toBe(2);
  });

  it('honours a manual column break as a band boundary on a sheet that fits', () => {
    const xlsx = buildXlsx({
      rows: grid(3, 3),
      columns: [{ min: 1, max: 3, widthChars: 10 }], // 3×1050 twips — fits the page
      colBreaks: [1], // manual break before column B (0-based index 1)
    });
    expect(bandCount(xlsx)).toBe(2);
    expect(pageCount(xlsx)).toBe(2);
  });
});

describe('printed headings + bands (§18.3.1.70)', () => {
  it('starts a headed band on its own page, letters and all', () => {
    // The band break is set on the band's own leading row; the letters row is
    // prepended in front of it afterwards. Left where it was, the letters
    // printed alone on one page and the band they label began the next —
    // ElapsedFormatTests.xlsx grew a blank page between its two bands.
    const xlsx = buildXlsx({
      rows: grid(4, 6),
      columns: wideCols,
      printOptions: { headings: true },
    });
    const tables = Ream.parse(xlsx).flow.body.filter((el) => el.kind === 'table');
    expect(tables.length).toBeGreaterThan(1);
    for (const [i, t] of tables.entries()) {
      // Row 0 is the column letters; it leads the band and carries the break.
      const first = t.table.rows[0]!;
      const second = t.table.rows[1];
      expect(first.properties.pageBreakBefore ?? false).toBe(i > 0);
      expect(second?.properties.pageBreakBefore ?? false).toBe(false);
    }
  });
});

// A drawing WIDER than a band does not stop at the band's edge: the page after
// it shows the rest, and the one after that the rest of that. picture.xlsx
// anchors a coin across 19 columns of a sheet that bands three ways, and
// emitted once it ended at the first page's edge with the two pages behind it
// blank — where every reader carries it across all three.
describe('a drawing wider than its column band (E-SHEET SE1)', () => {
  // Six 40-character columns band three ways; a picture anchored across all of
  // them reaches into every band.
  const across = buildXlsx({
    rows: grid(3, 6),
    columns: wideCols,
    sheetImage: {
      pngBytes: buildTinyPng(4, 4, [0, 0, 255, 255]),
      anchor: { from: [0, 0], to: [6, 3] },
    },
  });
  const images = (xlsx: Uint8Array): Array<number> =>
    Ream.parse(xlsx)
      .flow.body.filter((el) => el.kind === 'image')
      .map((el) => el.image.float?.posH?.offsetPt ?? 0);

  it('is emitted once per band it reaches, rebased to that band', () => {
    const lefts = images(across);
    expect(lefts.length).toBe(3);
    // The first copy keeps the sheet's own anchor; each next one is measured
    // from its band's left edge, so the offsets step DOWN.
    expect(lefts[1]).toBeLessThan(lefts[0]!);
    expect(lefts[2]).toBeLessThan(lefts[1]!);
  });

  it('leaves a drawing inside one band alone', () => {
    const inside = buildXlsx({
      rows: grid(3, 6),
      columns: wideCols,
      sheetImage: {
        pngBytes: buildTinyPng(4, 4, [0, 0, 255, 255]),
        anchor: { from: [0, 0], to: [1, 3] },
      },
    });
    expect(images(inside)).toHaveLength(1);
  });

  it('still starts a page for the band after one that only a drawing crosses', () => {
    // Every band's page is a page: the second and third hold no cell of their
    // own here, and counted by their cells alone the break never fired and all
    // three bands' drawings landed together.
    expect(pageCount(across)).toBe(3);
  });
});
