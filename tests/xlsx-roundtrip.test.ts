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
// The whole grid surface now round-trips — page model (SD3a), conditional
// formats / sparklines / table parts (SD3b). Only embedded charts (a sheet's
// drawing) are not yet written back.

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

const DXF_STYLES = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>`;

// Every grid field the writer round-trips (everything but the sheet's embedded
// charts, which the writer does not yet re-emit).
function normGrid(g: ParsedWorksheet): unknown {
  return {
    cells: g.cells,
    maxRow: g.maxRow,
    maxColumn: g.maxColumn,
    columns: g.columns,
    merges: g.merges,
    rowHeights: g.rowHeights,
    pageMargins: g.pageMargins,
    pageSetup: g.pageSetup,
    fitToPage: g.fitToPage,
    printOptions: g.printOptions,
    rowBreaks: g.rowBreaks,
    colBreaks: g.colBreaks,
    conditionalFormats: g.conditionalFormats,
    dataValidations: g.dataValidations,
    sparklines: g.sparklines,
    tables: g.tables,
    pane: g.pane,
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
  {
    name: 'page margins (SD3a)',
    xlsx: buildXlsx({
      rows: [[1]],
      pageMargins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    }),
  },
  {
    name: 'page setup landscape (SD3a)',
    xlsx: buildXlsx({
      rows: [[1]],
      pageSetup: { paperSize: 9, orientation: 'landscape', scale: 80 },
    }),
  },
  {
    name: 'fit to page (SD3a)',
    xlsx: buildXlsx({ rows: [[1]], fitToPage: true, pageSetup: { fitToWidth: 1, fitToHeight: 0 } }),
  },
  {
    name: 'print options (SD3a)',
    xlsx: buildXlsx({ rows: [[1]], printOptions: { gridLines: true, horizontalCentered: true } }),
  },
  { name: 'row breaks (SD3a)', xlsx: buildXlsx({ rows: [[1], [2], [3]], rowBreaks: [1] }) },
  { name: 'col breaks (SE1)', xlsx: buildXlsx({ rows: [[1, 2, 3]], colBreaks: [1] }) },
  {
    name: 'frozen pane — rows + cols (SE2)',
    xlsx: buildXlsx({
      rows: [
        [1, 2],
        [3, 4],
      ],
      freeze: { rows: 1, cols: 1 },
    }),
  },
  {
    name: 'frozen pane — top row (SE2)',
    xlsx: buildXlsx({ rows: [[1], [2]], freeze: { rows: 1 } }),
  },
  {
    name: 'conditional formatting cellIs (SD3b)',
    xlsx: buildXlsx({
      rows: [[2], [6]],
      stylesXml: DXF_STYLES,
      conditionalFormattingXml:
        '<conditionalFormatting sqref="A1:A2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>',
    }),
  },
  {
    name: 'conditional formatting colorScale (SD3b)',
    xlsx: buildXlsx({
      rows: [[1], [5], [10]],
      conditionalFormattingXml:
        '<conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>',
    }),
  },
  {
    name: 'conditional formatting dataBar (SD3b)',
    xlsx: buildXlsx({
      rows: [[1], [10]],
      conditionalFormattingXml:
        '<conditionalFormatting sqref="A1:A2"><cfRule type="dataBar" priority="1"><dataBar minLength="0" maxLength="100"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>',
    }),
  },
  {
    name: 'conditional formatting iconSet (SD3b)',
    xlsx: buildXlsx({
      rows: [[10], [50], [90]],
      conditionalFormattingXml:
        '<conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>',
    }),
  },
  {
    name: 'data validations (SV1)',
    xlsx: buildXlsx({
      rows: [['Yes'], [50]],
      dataValidationsXml:
        '<dataValidations count="2">' +
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">' +
        '<formula1>"Yes,No,Maybe"</formula1></dataValidation>' +
        '<dataValidation type="whole" operator="between" showErrorMessage="1" errorStyle="stop" sqref="A2">' +
        '<formula1>1</formula1><formula2>100</formula2></dataValidation>' +
        '</dataValidations>',
    }),
  },
  {
    name: 'sparkline (SD3b)',
    xlsx: buildXlsx({
      rows: [[1, 5, 3, 8]],
      extLstXml:
        '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup type="column"><x14:colorSeries rgb="FF376092"/><x14:sparklines><x14:sparkline><xm:f>Sheet1!A1:D1</xm:f><xm:sqref>E1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>',
    }),
  },
  {
    name: 'table parts (SD3b)',
    xlsx: buildXlsx({
      rows: [
        [1, 2],
        [3, 4],
        [5, 6],
      ],
      tables: [
        { ref: 'A1:B3', name: 'Data', styleName: 'TableStyleMedium2', showRowStripes: true },
      ],
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
