// E-SHEET SE1 — wide-sheet column-band pagination. A sheet wider than the page
// (at 100% scale) paginates ACROSS columns into bands ("down, then over") instead
// of being squeezed onto one page width; fit-to-page / scaled sheets keep the
// uniform-shrink path. The pure band maths is unit-tested; the projection +
// layout are exercised end-to-end via the band-table and page counts.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

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
