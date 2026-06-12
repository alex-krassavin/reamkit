// E-SHEET SA1/SA2 — the xlsx reader builds a SheetDoc (the SpreadsheetML IR
// node) before the print model projects it to a FlowDoc. These cover the IR
// node directly; the end-to-end render is covered (byte-identically) by the
// byte gate and xlsx.test.ts.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { buildDocxFromBody } from './fixtures/build-docx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';

const BAR_CHART =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>' +
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/>' +
  '<c:order val="0"/><c:val><c:numRef><c:numCache><c:ptCount val="2"/>' +
  '<c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numCache></c:numRef>' +
  '</c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';

describe('readXlsxToSheetDoc (E-SHEET SA1/SA2)', () => {
  it('builds a workbook SheetDoc: sheets, shared strings, defined names, date system', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        sheets: [
          {
            name: 'Alpha',
            rows: [
              ['hello', 42],
              ['world', 7],
            ],
          },
          { name: 'Beta', rows: [['second']] },
        ],
        date1904: true,
        definedNames: [{ name: '_xlnm.Print_Area', localSheetId: 0, value: 'Alpha!$A$1:$B$1' }],
      }),
    );

    expect(sheet.kind).toBe('sheet');
    expect(sheet.sheets.map((s) => s.name)).toEqual(['Alpha', 'Beta']);
    expect(sheet.date1904).toBe(true);
    expect(sheet.sharedStrings).toContain('hello');
    expect(sheet.definedNames).toHaveLength(1);
    expect(sheet.definedNames[0]!.name).toBe('_xlnm.Print_Area');
    // The grid keeps RAW cells + style indices — resolution is the projection's
    // job (so SA2 stays byte-zero). The first sheet's grid carries its cells.
    expect(sheet.sheets[0]!.grid.cells.length).toBeGreaterThan(0);
  });

  it('resolves sheet chart frames into the SheetDoc (per-sheet refs + chart data)', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        rows: [
          ['A', 4],
          ['B', 9],
        ],
        sheetChart: { chartXml: BAR_CHART },
      }),
    );

    expect(sheet.chartData?.size).toBe(1);
    expect(sheet.chartData?.has('xl/charts/chart1.xml')).toBe(true);
    const refs = sheet.sheets[0]!.charts;
    expect(refs).toHaveLength(1);
    expect(refs![0]!.chartPartPath).toBe('xl/charts/chart1.xml');
    expect(refs![0]!.widthPt).toBeGreaterThan(0);
  });
});

describe('Ream exposes the native SheetDoc (E-SHEET SB1)', () => {
  it('parse(xlsx).sheet is the SheetDoc; .flow is its projection', () => {
    const doc = Ream.parse(buildXlsx({ sheets: [{ name: 'Numbers', rows: [['x', 1]] }] }));
    expect(doc.sheet?.kind).toBe('sheet');
    expect(doc.sheet?.sheets[0]!.name).toBe('Numbers');
    // The render interlayer is the projected FlowDoc (a table block per sheet).
    expect(doc.flow.kind).toBe('flow');
    expect(doc.flow.body.some((el) => el.kind === 'table')).toBe(true);
  });

  it('parse(docx).sheet is undefined (a flow source has no native sheet)', () => {
    const doc = Ream.parse(buildDocxFromBody('<w:p><w:r><w:t>plain</w:t></w:r></w:p>'));
    expect(doc.sheet).toBeUndefined();
    expect(doc.flow.kind).toBe('flow');
  });
});
