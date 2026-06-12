// E-SHEET SD2 — xlsx roundtrip gate (analog of E-DOCX D6). Asserts that the
// SheetDoc is a stable fixpoint of the read↔write pair over a fixture matrix:
//
//   xlsx → SheetDoc(S1) → xlsx(b1) → SheetDoc(S2) → xlsx(b2)
//
//   • IR-identity: S2 equals S1 on the written surface (cells, shared strings,
//     styles, merges, columns, row heights, defined names, multi-sheet, dates).
//   • Byte-stable: b2 === b1 — the writer is deterministic on its own output, so
//     a faithful read→write loop is idempotent.
//
// Page setup / print options / conditional formats / sparklines / table parts
// are NOT yet written back (SD1 writes the grid core); they are excluded here
// and arrive in SD3.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { ParsedWorksheet } from '@/core/spreadsheet-model';
import type { SheetDoc } from '@/core/ir/sheet';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';

const STYLES = `
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><i/><sz val="12"/><color rgb="FF0070C0"/><name val="Arial"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>
  <borders count="2"><border/><border><left style="thin"><color rgb="FF000000"/></left><bottom style="medium"/></border></borders>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" wrapText="1"/></xf></cellXfs>`;

// The grid fields SD1 writes back (page setup / print options / CF / sparklines
// / tables are SD3 and deliberately excluded from the identity check).
function normGrid(g: ParsedWorksheet): unknown {
  return {
    cells: g.cells,
    maxRow: g.maxRow,
    maxColumn: g.maxColumn,
    columns: g.columns,
    merges: g.merges,
    rowHeights: g.rowHeights,
  };
}

function norm(d: SheetDoc): unknown {
  return {
    names: d.sheets.map((s) => s.name),
    grids: d.sheets.map((s) => normGrid(s.grid)),
    styles: d.styles,
    sharedStrings: d.sharedStrings,
    definedNames: d.definedNames,
    date1904: d.date1904,
  };
}

function roundtrip(xlsx: Uint8Array): {
  s1: SheetDoc;
  s2: SheetDoc;
  b1: Uint8Array;
  b2: Uint8Array;
} {
  const s1 = readXlsxToSheetDoc(xlsx);
  const b1 = writeXlsx(s1).bytes;
  const s2 = readXlsxToSheetDoc(b1);
  const b2 = writeXlsx(s2).bytes;
  return { s1, s2, b1, b2 };
}

const fixtures: Array<{ name: string; xlsx: Uint8Array }> = [
  {
    name: 'numbers',
    xlsx: buildXlsx({
      rows: [
        [1, 2, 3.5],
        [-4, 0, 1000000],
      ],
    }),
  },
  {
    name: 'shared strings (interned)',
    xlsx: buildXlsx({
      rows: [
        ['a', 'b'],
        ['a', 'c'],
        ['b', 'a'],
      ],
    }),
  },
  { name: 'booleans', xlsx: buildXlsx({ rows: [[true, false]] }) },
  {
    name: 'mixed values',
    xlsx: buildXlsx({
      rows: [
        [1, 'text'],
        ['more', 2.5],
      ],
    }),
  },
  { name: 'special chars in strings', xlsx: buildXlsx({ rows: [['a & b', '<x> "q"']] }) },
  {
    name: 'styled cells',
    xlsx: buildXlsx({
      rows: [
        [
          { value: 1, styleIndex: 1 },
          { value: 'hi', styleIndex: 1 },
        ],
      ],
      stylesXml: STYLES,
    }),
  },
  {
    name: 'merges',
    xlsx: buildXlsx({
      rows: [
        [1, 2, 3],
        [4, 5, 6],
      ],
      mergeRefs: ['A1:C1', 'A2:B2'],
    }),
  },
  {
    name: 'columns',
    xlsx: buildXlsx({ rows: [[1, 2]], columns: [{ min: 1, max: 1, widthChars: 14.5 }] }),
  },
  {
    name: 'row heights',
    xlsx: buildXlsx({
      rows: [[1], [2], [3]],
      rowHeights: [
        { row: 0, heightPt: 30, customHeight: true },
        { row: 2, heightPt: 18 },
      ],
    }),
  },
  { name: 'full style table', xlsx: buildXlsx({ rows: [[1]], stylesXml: STYLES }) },
  {
    name: 'multi-sheet',
    xlsx: buildXlsx({
      sheets: [
        { name: 'First', rows: [[1, 'a']] },
        {
          name: 'Second',
          rows: [
            [2, 'b'],
            [3, 'c'],
          ],
        },
      ],
    }),
  },
  { name: 'date1904', xlsx: buildXlsx({ rows: [[40000]], date1904: true }) },
  {
    name: 'defined names',
    xlsx: buildXlsx({
      rows: [[1, 2]],
      definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 0, value: 'Sheet1!$A$1:$B$1' }],
    }),
  },
];

describe('xlsx roundtrip gate (E-SHEET SD2)', () => {
  for (const { name, xlsx } of fixtures) {
    it(`${name}: SheetDoc identity + byte-stable`, () => {
      const { s1, s2, b1, b2 } = roundtrip(xlsx);
      expect(norm(s2)).toEqual(norm(s1));
      expect(b2).toEqual(b1);
    });
  }

  it('an empty sheet round-trips', () => {
    const { s1, s2, b1, b2 } = roundtrip(buildXlsx({ rows: [] }));
    expect(norm(s2)).toEqual(norm(s1));
    expect(b2).toEqual(b1);
  });
});
