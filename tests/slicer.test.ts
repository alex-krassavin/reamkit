// E-SHEET SV2 — slicers. A slicer is a visual filter panel (xl/slicers +
// xl/slicerCaches) anchored over a sheet. A native-table slicer resolves its
// buttons from the referenced table column's distinct values, with the column's
// autofilter giving selection; an OLAP/pivot slicer degrades to a caption box.
// The panel renders as a styled mini-table after the grid (like chart frames).

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

// A sheet with a Region table (A1:A5) and a slicer over its first column.
const regionXlsx = (opts: {
  filters?: ReadonlyArray<{ colId: number; values: ReadonlyArray<string> }>;
  columnCount?: number;
  styleName?: string;
}) =>
  buildXlsx({
    rows: [['Region'], ['North'], ['South'], ['East'], ['West']],
    tables: [{ ref: 'A1:A5', name: 'Data', ...(opts.filters ? { filters: opts.filters } : {}) }],
    slicers: [
      {
        name: 'Region',
        caption: 'Region',
        cacheName: 'Slicer_Region',
        ...(opts.columnCount ? { columnCount: opts.columnCount } : {}),
        styleName: opts.styleName ?? 'SlicerStyleLight2',
        cache: { sourceName: 'Region', tableId: 1, column: 1 },
      },
    ],
  });

const firstSlicer = (xlsx: Uint8Array) => readXlsxToSheetDoc(xlsx).sheets[0]!.slicers?.[0];

describe('slicers — resolve native-table items (E-SHEET SV2)', () => {
  it('resolves the table column values as buttons, all selected by default', () => {
    const slicer = firstSlicer(regionXlsx({}));
    expect(slicer?.caption).toBe('Region');
    expect(slicer?.columnCount).toBe(1);
    expect(slicer?.items.map((i) => i.label)).toEqual(['North', 'South', 'East', 'West']);
    expect(slicer?.items.every((i) => i.selected)).toBe(true);
  });

  it('takes selection from the table column autofilter', () => {
    const slicer = firstSlicer(regionXlsx({ filters: [{ colId: 0, values: ['North', 'South'] }] }));
    const selected = slicer!.items.filter((i) => i.selected).map((i) => i.label);
    const unselected = slicer!.items.filter((i) => !i.selected).map((i) => i.label);
    expect(selected).toEqual(['North', 'South']);
    expect(unselected).toEqual(['East', 'West']);
  });

  it('carries the column count and a resolved style accent', () => {
    const slicer = firstSlicer(regionXlsx({ columnCount: 2 }));
    expect(slicer?.columnCount).toBe(2);
    expect(slicer?.headerHex).toBeDefined();
    expect(slicer?.selectedHex).toBe(slicer?.headerHex);
    expect(slicer?.headerTextHex).toBe('FFFFFF');
  });

  it('degrades an OLAP/pivot slicer (no tableSlicerCache) to a caption-only box', () => {
    const xlsx = buildXlsx({
      rows: [['x']],
      slicers: [
        {
          name: 'Cat',
          caption: 'Category',
          cacheName: 'Slicer_Cat',
          styleName: 'SlicerStyleLight1',
          cache: { sourceName: 'Category' }, // no tableId/column → not a table slicer
        },
      ],
    });
    const slicer = firstSlicer(xlsx);
    expect(slicer?.caption).toBe('Category');
    expect(slicer?.items).toEqual([]);
  });

  it('leaves a sheet with no slicers untouched (no slicers field)', () => {
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1]] })).sheets[0]!;
    expect(sheet.slicers).toBeUndefined();
  });
});

describe('slicers — projection to a styled box (E-SHEET SV2)', () => {
  it('emits a slicer table after the grid with a caption header and button cells', () => {
    const xlsx = regionXlsx({ filters: [{ colId: 0, values: ['North'] }] });
    const model = firstSlicer(xlsx)!;
    const tables = Ream.parse(xlsx).flow.body.filter((el) => el.kind === 'table');
    // body[0] is the grid (the table's cached cells); the slicer box follows it.
    expect(tables.length).toBe(2);
    const box = tables[1];
    if (box?.kind !== 'table') throw new Error('expected the slicer table');
    // Caption header.
    const header = box.table.rows[0]?.cells[0];
    expect(header?.content[0]).toMatchObject({
      kind: 'paragraph',
      paragraph: { runs: [{ text: 'Region', properties: { bold: true } }] },
    });
    expect(header?.properties.shading?.colorHex).toBe(model.headerHex);
    // The selected item (North) takes the accent fill; an unselected one a band.
    const labelCell = (label: string) =>
      box.table.rows
        .flatMap((r) => r.cells)
        .find((c) => {
          const p = c.content[0];
          return p?.kind === 'paragraph' && p.paragraph.runs[0]?.text === label;
        });
    expect(labelCell('North')?.properties.shading?.colorHex).toBe(model.selectedHex);
    expect(labelCell('West')?.properties.shading?.colorHex).toBe('F2F2F2');
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('slicers — render (E-SHEET SV2)', () => {
  it('renders the slicer caption and buttons into HTML', async () => {
    const xlsx = regionXlsx({});
    const html = new TextDecoder().decode(await Ream.parse(xlsx).convert('html'));
    expect(html).toContain('Region'); // caption
    expect(html).toContain('North'); // a button
    expect(html).toContain('West');
    const accent = firstSlicer(xlsx)!.selectedHex!;
    expect(html).toContain(`#${accent}`); // the accent fill on selected buttons
  });

  it('renders a sheet with a slicer to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(regionXlsx({ columnCount: 2 }), { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
