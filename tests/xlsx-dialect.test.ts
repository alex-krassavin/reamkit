// Dialect invariance: how a document is SPELLED must not change what it means.
//
// The rest of the xlsx suite builds its fixtures with build-xlsx.ts, which
// emits exactly one spelling of SpreadsheetML — the same one our writer
// produces. Those tests therefore establish that the parser can read our
// writer, which is not the question anyone is asking.
//
// Rather than restate every assertion once per dialect, this asserts the
// property underneath them: re-spell a package (namespace prefix, GUID
// relationship ids, mc:Ignorable extension attributes) and the parse must come
// back identical. It holds for every fixture rather than the handful of real
// documents we managed to adopt, and it fails loudly for a parser that has
// become accidentally literal about a spelling — which is how tdf122336.xlsx
// managed to render blank.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { toDialect } from './fixtures/xlsx-dialect';
import type { DialectOptions } from './fixtures/xlsx-dialect';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { parseWorksheet } from '@/excel/worksheet-parser';

const here = dirname(fileURLToPath(import.meta.url));
const SML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const DIALECTS: ReadonlyArray<{ name: string; options: DialectOptions }> = [
  { name: 'x: namespace prefix', options: { nsPrefix: 'x' } },
  { name: 'GUID relationship ids', options: { guidRelIds: true } },
  { name: 'mc:Ignorable extension attributes', options: { mcIgnorable: true } },
  {
    name: 'all three at once',
    options: { nsPrefix: 'x', guidRelIds: true, mcIgnorable: true },
  },
];

/** The parse, reduced to what a renderer actually consumes. */
function digest(bytes: Uint8Array): unknown {
  const doc = readXlsxToSheetDoc(bytes);
  return {
    date1904: doc.date1904,
    sharedStrings: doc.sharedStrings,
    definedNames: doc.definedNames,
    sheets: doc.sheets.map((s) => ({
      name: s.name,
      cells: s.grid.cells,
      merges: s.grid.merges,
      columns: s.grid.columns,
      rowHeights: s.grid.rowHeights,
      maxRow: s.grid.maxRow,
      maxColumn: s.grid.maxColumn,
      pageSetup: s.grid.pageSetup,
      pageMargins: s.grid.pageMargins,
      rowBreaks: s.grid.rowBreaks,
      colBreaks: s.grid.colBreaks,
    })),
  };
}

const SYNTHETIC = buildXlsx({
  rows: [
    ['Region', 'Revenue', 'Growth'],
    ['Moscow', 1200000, true],
    ['SPb', 720000, false],
  ],
  columns: [{ min: 1, max: 1, widthChars: 18 }],
  mergeRefs: ['A1:C1'],
  pageSetup: { paperSize: 9, orientation: 'landscape' },
  pageMargins: { left: 0.5, right: 0.5, top: 1, bottom: 1, header: 0.3, footer: 0.3 },
  rowBreaks: [1],
});

// Adopted documents make the property meaningful beyond our own generator:
// re-spelling a file we did not write must be a no-op too.
const REAL = ['tdf58243.xlsx', 'RepeatingRowsCols.xlsx', 'tdf100034.xlsx'];

describe('dialect invariance', () => {
  for (const { name, options } of DIALECTS) {
    it(`synthetic workbook survives ${name}`, () => {
      expect(digest(toDialect(SYNTHETIC, options))).toEqual(digest(SYNTHETIC));
    });
  }

  for (const file of REAL) {
    const bytes = new Uint8Array(readFileSync(resolve(here, 'fixtures/real', file)));
    for (const { name, options } of DIALECTS) {
      it(`${file} survives ${name}`, () => {
        expect(digest(toDialect(bytes, options))).toEqual(digest(bytes));
      });
    }
  }

  it('the transform actually re-spells the package (guards a no-op test)', () => {
    // A transform that quietly did nothing would make every assertion above
    // pass while proving nothing at all.
    const prefixed = toDialect(SYNTHETIC, {
      nsPrefix: 'x',
      guidRelIds: true,
      mcIgnorable: true,
    });
    const parts = unzipSync(prefixed);
    const read = (path: string): string => new TextDecoder().decode(parts[path]);

    const sheet = read('xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<x:worksheet');
    expect(sheet).toContain('<x:sheetData');
    expect(sheet).toContain(`xmlns:x="${SML_NS}"`);
    expect(sheet).not.toContain(`xmlns="${SML_NS}"`);
    expect(sheet).toContain('mc:Ignorable="x14ac"');

    const rels = read('xl/_rels/workbook.xml.rels');
    expect(rels).not.toMatch(/"rId\d+"/);
    expect(read('xl/workbook.xml')).not.toMatch(/rId\d+/);
  });
});
describe('a chartsheet root (§18.3.1.99)', () => {
  it("reads its drawing and page setup, not just a worksheet's", () => {
    // A `<chartsheet>` is a tab that is nothing but a chart: no sheetData, but
    // the same pageMargins / pageSetup / `<drawing r:id>` a worksheet carries.
    // Reading only `<worksheet>` made it an empty sheet, which then dropped out
    // of the print entirely — 47813.xlsx lost its whole "Chart" tab, a page the
    // reference fills with a plot of 1700 points.
    const xml = new TextEncoder().encode(
      '<?xml version="1.0"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheetPr/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>' +
        '<drawing r:id="rId1"/></chartsheet>',
    );
    const parsed = parseWorksheet(xml);
    expect(parsed.drawingRelId).toBe('rId1');
    expect(parsed.cells).toHaveLength(0);
    expect(parsed.pageMargins?.leftInches).toBeCloseTo(0.75, 3);
  });
});
