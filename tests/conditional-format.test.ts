// E-SHEET SC1 — conditional formatting (cellIs). The first Excel feature the
// flat table projection could not express: a value-driven per-cell fill/font.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';

const STYLES_WITH_DXF = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <dxfs count="2">
    <dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>
    <dxf><font><b/><color rgb="FF0000FF"/></font></dxf>
  </dxfs>`;

const cfRule = (operator: string, dxfId: number, ...formulas: Array<number>): string =>
  `<cfRule type="cellIs" dxfId="${dxfId}" priority="${dxfId + 1}" operator="${operator}">` +
  formulas.map((f) => `<formula>${f}</formula>`).join('') +
  '</cfRule>';

describe('conditional formatting — cellIs (E-SHEET SC1)', () => {
  it('applies a dxf highlight fill to cells matching a greaterThan rule', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${cfRule('greaterThan', 0, 5)}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[2], [6], [9]],
        stylesXml: STYLES_WITH_DXF,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    const shadingOf = (row: number): string | undefined =>
      table.table.rows[row]?.cells[0]?.properties.shading?.colorHex;

    expect(shadingOf(0)).toBeUndefined(); // A1 = 2, not > 5
    expect(shadingOf(1)).toBe('FF0000'); // A2 = 6 → red highlight (dxf bgColor)
    expect(shadingOf(2)).toBe('FF0000'); // A3 = 9 → red highlight
  });

  it('the highest-priority (lowest number) matching rule wins', () => {
    // Two overlapping rules on A1=8: rule dxf 0 (priority 1, red fill) beats
    // rule dxf 1 (priority 2, blue bold font) because its priority is lower.
    const cf =
      '<conditionalFormatting sqref="A1">' +
      cfRule('greaterThan', 0, 5) +
      cfRule('greaterThan', 1, 5) +
      '</conditionalFormatting>';
    const flow = Ream.parse(
      buildXlsx({ rows: [[8]], stylesXml: STYLES_WITH_DXF, conditionalFormattingXml: cf }),
    ).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.table.rows[0]?.cells[0]?.properties.shading?.colorHex).toBe('FF0000');
  });

  it('a between rule highlights only the in-range cell', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${cfRule('between', 0, 3, 7)}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[1], [5], [9]],
        stylesXml: STYLES_WITH_DXF,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    const shadingOf = (row: number): string | undefined =>
      table.table.rows[row]?.cells[0]?.properties.shading?.colorHex;
    expect(shadingOf(0)).toBeUndefined(); // 1 — below
    expect(shadingOf(1)).toBe('FF0000'); // 5 — between 3 and 7
    expect(shadingOf(2)).toBeUndefined(); // 9 — above
  });

  it('parses cellIs rules onto the sheet and dxfs onto the styles', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${cfRule('greaterThan', 0, 5)}</conditionalFormatting>`;
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [[10]], stylesXml: STYLES_WITH_DXF, conditionalFormattingXml: cf }),
    );
    const cfs = sheet.sheets[0]!.grid.conditionalFormats;
    expect(cfs).toHaveLength(1);
    const rule = cfs![0]!.rules[0]!;
    expect(rule.type).toBe('cellIs');
    expect(rule.operator).toBe('greaterThan');
    expect(rule.formulas).toEqual(['5']);
    expect(sheet.styles.dxfs).toHaveLength(2);
  });
});
