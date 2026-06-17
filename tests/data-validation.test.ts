// E-SHEET SV1 — data validation. A worksheet <dataValidations> carries per-range
// input constraints; the only one with a visual signature is type="list", which
// paints an in-cell dropdown affordance. Every validation round-trips through the
// writer so the SheetDoc stays a byte-stable fixpoint.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { writeXlsx } from '@/excel/xlsx-writer';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

// One <dataValidations> with a single <dataValidation>. `formula1` carries the
// raw inner text (a list source is the quoted "A,B,C" or a $A$1:$A$3 range).
const dv = (
  type: string,
  sqref: string,
  opts: {
    formula1?: string;
    formula2?: string;
    operator?: string;
    showDropDown?: boolean;
  } = {},
): string =>
  `<dataValidations count="1"><dataValidation type="${type}"` +
  (opts.operator ? ` operator="${opts.operator}"` : '') +
  (opts.showDropDown ? ' showDropDown="1"' : '') +
  ` allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${sqref}">` +
  (opts.formula1 !== undefined ? `<formula1>${opts.formula1}</formula1>` : '') +
  (opts.formula2 !== undefined ? `<formula2>${opts.formula2}</formula2>` : '') +
  '</dataValidation></dataValidations>';

const dropdownAt = (
  flow: ReturnType<typeof Ream.parse>['flow'],
  row: number,
  col: number,
): boolean | undefined => {
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table.table.rows[row]?.cells[col]?.properties.dropdown;
};

describe('data validation — parse (E-SHEET SV1)', () => {
  it('parses a list validation onto the sheet, formulas and flags included', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        rows: [['Yes'], ['No']],
        dataValidationsXml: dv('list', 'A1:A2', { formula1: '"Yes,No,Maybe"' }),
      }),
    );
    const dvs = sheet.sheets[0]!.grid.dataValidations;
    expect(dvs).toHaveLength(1);
    const v = dvs![0]!;
    expect(v.type).toBe('list');
    expect(v.formula1).toBe('"Yes,No,Maybe"');
    expect(v.allowBlank).toBe(true);
    expect(v.showInputMessage).toBe(true);
    expect(v.ranges[0]).toMatchObject({ startColumn: 0, startRow: 0, endColumn: 0, endRow: 1 });
  });

  it('parses a numeric between validation with two formulas', () => {
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        rows: [[50]],
        dataValidationsXml: dv('whole', 'A1', {
          operator: 'between',
          formula1: '1',
          formula2: '100',
        }),
      }),
    );
    const v = sheet.sheets[0]!.grid.dataValidations![0]!;
    expect(v.type).toBe('whole');
    expect(v.operator).toBe('between');
    expect(v.formula1).toBe('1');
    expect(v.formula2).toBe('100');
  });
});

describe('data validation — dropdown affordance (E-SHEET SV1)', () => {
  it('marks every cell of a list range with a dropdown', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['Yes'], ['No'], ['Maybe']],
        dataValidationsXml: dv('list', 'A1:A3', { formula1: '"Yes,No,Maybe"' }),
      }),
    ).flow;
    expect(dropdownAt(flow, 0, 0)).toBe(true);
    expect(dropdownAt(flow, 1, 0)).toBe(true);
    expect(dropdownAt(flow, 2, 0)).toBe(true);
  });

  it('leaves cells outside the validation range unmarked', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [
          ['Yes', 'free'],
          ['No', 'text'],
        ],
        dataValidationsXml: dv('list', 'A1:A2', { formula1: '"Yes,No"' }),
      }),
    ).flow;
    expect(dropdownAt(flow, 0, 0)).toBe(true);
    expect(dropdownAt(flow, 0, 1)).toBeUndefined();
  });

  it('covers each area of a multi-area sqref', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['a', 'b', 'c']],
        dataValidationsXml: dv('list', 'A1 C1', { formula1: '"a,b,c"' }),
      }),
    ).flow;
    expect(dropdownAt(flow, 0, 0)).toBe(true);
    expect(dropdownAt(flow, 0, 1)).toBeUndefined();
    expect(dropdownAt(flow, 0, 2)).toBe(true);
  });

  it('a non-list validation paints no dropdown', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [[5]],
        dataValidationsXml: dv('whole', 'A1', { operator: 'greaterThan', formula1: '0' }),
      }),
    ).flow;
    expect(dropdownAt(flow, 0, 0)).toBeUndefined();
  });

  it('showDropDown="1" suppresses the dropdown (ECMA inverted sense)', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['Yes']],
        dataValidationsXml: dv('list', 'A1', { formula1: '"Yes,No"', showDropDown: true }),
      }),
    ).flow;
    expect(dropdownAt(flow, 0, 0)).toBeUndefined();
  });
});

// The dropdown must survive the full render path — the layout shape pass (PDF)
// and the floated inline-SVG marker (HTML), not just reach the FlowDoc.
const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('data validation — render (E-SHEET SV1)', () => {
  it('renders a list dropdown as a floated inline SVG in HTML', async () => {
    const xlsx = buildXlsx({
      rows: [['Yes'], ['No']],
      dataValidationsXml: dv('list', 'A1:A2', { formula1: '"Yes,No,Maybe"' }),
    });
    const html = new TextDecoder().decode(await Ream.parse(xlsx).convert('html'));
    expect(html).toContain('float:right'); // the dropdown button floats to the right edge
    expect(html).toContain('#595959'); // the ▾ arrow colour
    expect(html).toContain('3.4,5 9.6,5 6.5,8.6'); // the ▾ polygon
  });

  it('renders a sheet with a list validation to a valid PDF', () => {
    const xlsx = buildXlsx({
      rows: [['Yes'], ['No'], ['Maybe']],
      dataValidationsXml: dv('list', 'A1:A3', { formula1: '"Yes,No,Maybe"' }),
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});

describe('data validation — write-back (E-SHEET SV1)', () => {
  it('round-trips the validations through the xlsx writer', () => {
    const s1 = readXlsxToSheetDoc(
      buildXlsx({
        rows: [['Yes'], [50]],
        dataValidationsXml:
          '<dataValidations count="2">' +
          '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">' +
          '<formula1>"Yes,No,Maybe"</formula1></dataValidation>' +
          '<dataValidation type="whole" operator="between" showErrorMessage="1" errorStyle="stop" sqref="A2">' +
          '<formula1>1</formula1><formula2>100</formula2></dataValidation>' +
          '</dataValidations>',
      }),
    );
    const s2 = readXlsxToSheetDoc(writeXlsx(s1).bytes);
    expect(s2.sheets[0]!.grid.dataValidations).toEqual(s1.sheets[0]!.grid.dataValidations);
    const list = s2.sheets[0]!.grid.dataValidations!.find((v) => v.type === 'list')!;
    expect(list.formula1).toBe('"Yes,No,Maybe"');
    const whole = s2.sheets[0]!.grid.dataValidations!.find((v) => v.type === 'whole')!;
    expect(whole.operator).toBe('between');
    expect(whole.errorStyle).toBe('stop');
    expect(whole.formula2).toBe('100');
  });
});
