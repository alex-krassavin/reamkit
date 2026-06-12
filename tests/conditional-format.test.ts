// E-SHEET SC1 — conditional formatting (cellIs). The first Excel feature the
// flat table projection could not express: a value-driven per-cell fill/font.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

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
    if (rule.type !== 'cellIs') throw new Error('unreachable');
    expect(rule.operator).toBe('greaterThan');
    expect(rule.formulas).toEqual(['5']);
    expect(sheet.styles.dxfs).toHaveLength(2);
  });
});

// E-SHEET SC1b — colorScale: a gradient fill interpolated across the range's
// value extent. The cross-cell feature cellIs (compare-to-constant) cannot do.
const PLAIN_STYLES = `
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>`;

// stops/colors length must match; each color is given as 6 hex (ARGB alpha FF
// is prepended). `cfvos` is a list of `type[:val]` stop descriptors.
const colorScale = (priority: number, cfvos: Array<string>, ...colors6: Array<string>): string => {
  const cfvoXml = cfvos
    .map((s) => {
      const [type, val] = s.split(':');
      return val !== undefined ? `<cfvo type="${type}" val="${val}"/>` : `<cfvo type="${type}"/>`;
    })
    .join('');
  const colorXml = colors6.map((c) => `<color rgb="FF${c}"/>`).join('');
  return `<cfRule type="colorScale" priority="${priority}"><colorScale>${cfvoXml}${colorXml}</colorScale></cfRule>`;
};

const shadingAt = (
  flow: ReturnType<typeof Ream.parse>['flow'],
  row: number,
): string | undefined => {
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table.table.rows[row]?.cells[0]?.properties.shading?.colorHex;
};

describe('conditional formatting — colorScale (E-SHEET SC1b)', () => {
  it('interpolates a 2-stop min→max gradient over the range', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${colorScale(1, ['min', 'max'], 'FF0000', '00FF00')}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({ rows: [[0], [5], [10]], stylesXml: PLAIN_STYLES, conditionalFormattingXml: cf }),
    ).flow;
    expect(shadingAt(flow, 0)).toBe('FF0000'); // 0 = min → red endpoint
    expect(shadingAt(flow, 1)).toBe('808000'); // 5 = midpoint → halfway red↔green
    expect(shadingAt(flow, 2)).toBe('00FF00'); // 10 = max → green endpoint
  });

  it('interpolates a 3-stop min/percentile/max gradient', () => {
    const cf = `<conditionalFormatting sqref="A1:A5">${colorScale(1, ['min', 'percentile:50', 'max'], 'FF0000', '00FF00', '0000FF')}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[1], [2], [3], [4], [5]],
        stylesXml: PLAIN_STYLES,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    expect(shadingAt(flow, 0)).toBe('FF0000'); // 1 = min
    expect(shadingAt(flow, 1)).toBe('808000'); // 2 = halfway min↔median
    expect(shadingAt(flow, 2)).toBe('00FF00'); // 3 = median (percentile 50)
    expect(shadingAt(flow, 3)).toBe('008080'); // 4 = halfway median↔max
    expect(shadingAt(flow, 4)).toBe('0000FF'); // 5 = max
  });

  it('a higher-priority cellIs rule wins over a colorScale on the same cell', () => {
    const cellIsYellow = `
      <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
      <dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs>`;
    const cf =
      '<conditionalFormatting sqref="A1:A2">' +
      '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>10</formula></cfRule>' +
      colorScale(2, ['min', 'max'], 'FF0000', '0000FF') +
      '</conditionalFormatting>';
    const flow = Ream.parse(
      buildXlsx({ rows: [[1], [10]], stylesXml: cellIsYellow, conditionalFormattingXml: cf }),
    ).flow;
    expect(shadingAt(flow, 0)).toBe('FF0000'); // A1=1: cellIs misses → colorScale min (red)
    expect(shadingAt(flow, 1)).toBe('FFFF00'); // A2=10: cellIs (prio 1) wins over scale max
  });

  it('parses a colorScale rule onto the sheet (cfvos + colors)', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${colorScale(1, ['min', 'max'], 'FF0000', '00FF00')}</conditionalFormatting>`;
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [[1], [2], [3]], stylesXml: PLAIN_STYLES, conditionalFormattingXml: cf }),
    );
    const rule = sheet.sheets[0]!.grid.conditionalFormats![0]!.rules[0]!;
    expect(rule.type).toBe('colorScale');
    if (rule.type !== 'colorScale') throw new Error('unreachable');
    expect(rule.cfvos.map((c) => c.type)).toEqual(['min', 'max']);
    expect(rule.colorsHex).toEqual(['FF0000', '00FF00']);
  });
});

// E-SHEET SC1c — dataBar: an in-cell bar whose length encodes the value within
// the range extent (the cell carries CellProperties.dataBar { fraction, color }).
const dataBar = (
  priority: number,
  color6: string,
  bounds?: { min?: number; max?: number },
): string => {
  const lens =
    `${bounds?.min !== undefined ? ` minLength="${bounds.min}"` : ''}` +
    `${bounds?.max !== undefined ? ` maxLength="${bounds.max}"` : ''}`;
  return (
    `<cfRule type="dataBar" priority="${priority}"><dataBar${lens}>` +
    `<cfvo type="min"/><cfvo type="max"/><color rgb="FF${color6}"/></dataBar></cfRule>`
  );
};

const barAt = (
  flow: ReturnType<typeof Ream.parse>['flow'],
  row: number,
): { fraction: number; colorHex: string; startFraction?: number } | undefined => {
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table.table.rows[row]?.cells[0]?.properties.dataBar;
};

describe('conditional formatting — dataBar (E-SHEET SC1c)', () => {
  it('sets a bar fraction from the value position in the range extent', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${dataBar(1, '638EC6', { min: 0, max: 100 })}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({ rows: [[0], [5], [10]], stylesXml: PLAIN_STYLES, conditionalFormattingXml: cf }),
    ).flow;
    expect(barAt(flow, 0)?.colorHex).toBe('638EC6');
    expect(barAt(flow, 0)?.fraction).toBeCloseTo(0); // 0 = min → empty bar
    expect(barAt(flow, 1)?.fraction).toBeCloseTo(0.5); // 5 = midpoint
    expect(barAt(flow, 2)?.fraction).toBeCloseTo(1); // 10 = max → full bar
  });

  it('clamps the bar between minLength and maxLength', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${dataBar(1, '638EC6', { min: 20, max: 80 })}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({ rows: [[0], [5], [10]], stylesXml: PLAIN_STYLES, conditionalFormattingXml: cf }),
    ).flow;
    expect(barAt(flow, 0)?.fraction).toBeCloseTo(0.2); // min value → minLength
    expect(barAt(flow, 1)?.fraction).toBeCloseTo(0.5); // 0.2 + 0.5·0.6
    expect(barAt(flow, 2)?.fraction).toBeCloseTo(0.8); // max value → maxLength
  });

  it('a data bar coexists with a cellIs highlight on the same cell', () => {
    const cf =
      '<conditionalFormatting sqref="A1:A2">' +
      cfRule('greaterThan', 0, 5) + // cellIs >5 → red fill (dxf 0), priority 1
      dataBar(2, '638EC6', { min: 0, max: 100 }) +
      '</conditionalFormatting>';
    const flow = Ream.parse(
      buildXlsx({ rows: [[1], [10]], stylesXml: STYLES_WITH_DXF, conditionalFormattingXml: cf }),
    ).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    const cell = (row: number) => table.table.rows[row]?.cells[0]?.properties;
    // A1=1: cellIs misses (no fill); bar empty.
    expect(cell(0)?.shading).toBeUndefined();
    expect(barAt(flow, 0)?.fraction).toBeCloseTo(0);
    // A2=10: cellIs red fill AND a full bar — both apply.
    expect(cell(1)?.shading?.colorHex).toBe('FF0000');
    expect(barAt(flow, 1)?.fraction).toBeCloseTo(1);
    expect(barAt(flow, 1)?.colorHex).toBe('638EC6');
  });

  it('parses a dataBar rule onto the sheet', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${dataBar(1, '638EC6', { min: 0, max: 100 })}</conditionalFormatting>`;
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [[1], [2], [3]], stylesXml: PLAIN_STYLES, conditionalFormattingXml: cf }),
    );
    const rule = sheet.sheets[0]!.grid.conditionalFormats![0]!.rules[0]!;
    expect(rule.type).toBe('dataBar');
    if (rule.type !== 'dataBar') throw new Error('unreachable');
    expect(rule.colorHex).toBe('638EC6');
    expect(rule.cfvos.map((c) => c.type)).toEqual(['min', 'max']);
    expect(rule.minLength).toBe(0);
    expect(rule.maxLength).toBe(100);
  });

  it('draws a mixed-sign data bar around a zero axis (Tail TC4)', () => {
    // Extent [-10, 10] → the axis sits at 0.5; negatives run left in red.
    const cf = `<conditionalFormatting sqref="A1:A3">${dataBar(1, '638EC6')}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[-10], [0], [10]],
        stylesXml: PLAIN_STYLES,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    const neg = barAt(flow, 0);
    expect(neg?.colorHex).toBe('FF0000'); // negative → Excel's default red
    expect(neg?.fraction).toBeCloseTo(0.5);
    expect(neg?.startFraction).toBeCloseTo(0); // runs from the left up to the axis
    const pos = barAt(flow, 2);
    expect(pos?.colorHex).toBe('638EC6'); // positive → series colour
    expect(pos?.fraction).toBeCloseTo(0.5);
    expect(pos?.startFraction).toBeCloseTo(0.5); // runs from the axis to the right
  });

  it('keeps an all-positive data bar left-anchored (no axis)', () => {
    const cf = `<conditionalFormatting sqref="A1:A2">${dataBar(1, '638EC6', { min: 0, max: 100 })}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({ rows: [[5], [10]], stylesXml: PLAIN_STYLES, conditionalFormattingXml: cf }),
    ).flow;
    expect(barAt(flow, 1)?.startFraction).toBeUndefined();
  });
});

// E-SHEET SC1c — iconSet: a glyph per cell from a named family, chosen by the
// value's bucket among the cfvo thresholds (the cell carries CellProperties.icon
// { shape, colorHex }; the xlsx layer maps families → format-neutral shapes).
const iconSet = (name: string, cfvos: Array<string>, reverse = false): string => {
  const cfvoXml = cfvos
    .map((s) => {
      const [type, val] = s.split(':');
      return `<cfvo type="${type}" val="${val}"/>`;
    })
    .join('');
  return (
    `<cfRule type="iconSet" priority="1"><iconSet iconSet="${name}"${reverse ? ' reverse="1"' : ''}>` +
    `${cfvoXml}</iconSet></cfRule>`
  );
};

const iconAt = (
  flow: ReturnType<typeof Ream.parse>['flow'],
  row: number,
): { shape: string; colorHex: string } | undefined => {
  const table = flow.body.find((el) => el.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected a table');
  return table.table.rows[row]?.cells[0]?.properties.icon;
};

describe('conditional formatting — iconSet (E-SHEET SC1c)', () => {
  it('picks a traffic-light icon by the value bucket', () => {
    // percent 0/33/67 over [10,90] → thresholds 10, 36.4, 63.6.
    const cf = `<conditionalFormatting sqref="A1:A3">${iconSet('3TrafficLights1', ['percent:0', 'percent:33', 'percent:67'])}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[10], [50], [90]],
        stylesXml: PLAIN_STYLES,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    expect(iconAt(flow, 0)).toEqual({ shape: 'circle', colorHex: 'FF0000' }); // bucket 0
    expect(iconAt(flow, 1)).toEqual({ shape: 'circle', colorHex: 'FFC000' }); // bucket 1
    expect(iconAt(flow, 2)).toEqual({ shape: 'circle', colorHex: '00B050' }); // bucket 2
  });

  it('maps 3Arrows to down/right/up triangles', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${iconSet('3Arrows', ['percent:0', 'percent:33', 'percent:67'])}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[10], [50], [90]],
        stylesXml: PLAIN_STYLES,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    expect(iconAt(flow, 0)?.shape).toBe('triangleDown'); // lowest bucket
    expect(iconAt(flow, 1)?.shape).toBe('triangleRight');
    expect(iconAt(flow, 2)?.shape).toBe('triangleUp'); // highest bucket
  });

  it('reverse flips the icon order (low value → top icon)', () => {
    const cf = `<conditionalFormatting sqref="A1:A2">${iconSet('3TrafficLights1', ['percent:0', 'percent:33', 'percent:67'], true)}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({ rows: [[10], [90]], stylesXml: PLAIN_STYLES, conditionalFormattingXml: cf }),
    ).flow;
    expect(iconAt(flow, 0)?.colorHex).toBe('00B050'); // value 10, reversed → green
    expect(iconAt(flow, 1)?.colorHex).toBe('FF0000'); // value 90, reversed → red
  });

  it('maps 3Signs to diamond / triangle / circle by bucket (Tail TC2)', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${iconSet('3Signs', ['percent:0', 'percent:33', 'percent:67'])}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[10], [50], [90]],
        stylesXml: PLAIN_STYLES,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    expect(iconAt(flow, 0)?.shape).toBe('diamond'); // low
    expect(iconAt(flow, 1)?.shape).toBe('triangleUp');
    expect(iconAt(flow, 2)?.shape).toBe('circle'); // high
  });

  it('renders a *Gray family monochrome (Tail TC2)', () => {
    const cf = `<conditionalFormatting sqref="A1:A3">${iconSet('3ArrowsGray', ['percent:0', 'percent:33', 'percent:67'])}</conditionalFormatting>`;
    const flow = Ream.parse(
      buildXlsx({
        rows: [[10], [50], [90]],
        stylesXml: PLAIN_STYLES,
        conditionalFormattingXml: cf,
      }),
    ).flow;
    // All three arrows are grey; only the direction differs.
    expect(iconAt(flow, 0)).toEqual({ shape: 'triangleDown', colorHex: '808080' });
    expect(iconAt(flow, 1)).toEqual({ shape: 'triangleRight', colorHex: '808080' });
    expect(iconAt(flow, 2)).toEqual({ shape: 'triangleUp', colorHex: '808080' });
  });

  it('parses an iconSet rule onto the sheet', () => {
    const cf = `<conditionalFormatting sqref="A1:A4">${iconSet('4Rating', ['percent:0', 'percent:25', 'percent:50', 'percent:75'])}</conditionalFormatting>`;
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        rows: [[1], [2], [3], [4]],
        stylesXml: PLAIN_STYLES,
        conditionalFormattingXml: cf,
      }),
    );
    const rule = sheet.sheets[0]!.grid.conditionalFormats![0]!.rules[0]!;
    expect(rule.type).toBe('iconSet');
    if (rule.type !== 'iconSet') throw new Error('unreachable');
    expect(rule.iconSet).toBe('4Rating');
    expect(rule.cfvos).toHaveLength(4);
  });
});

// A sheet carrying both decorations must survive the full layout + PDF emit
// path (the bar FillItem and the icon ShapeItem), not just reach the FlowDoc.
const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('conditional formatting — render smoke (E-SHEET SC1c)', () => {
  it('renders a sheet with a data bar and an icon set to a valid PDF', () => {
    const cf =
      `<conditionalFormatting sqref="A1:A3">${dataBar(2, '638EC6', { min: 0, max: 100 })}</conditionalFormatting>` +
      `<conditionalFormatting sqref="B1:B3">${iconSet('3TrafficLights1', ['percent:0', 'percent:33', 'percent:67'])}</conditionalFormatting>`;
    const xlsx = buildXlsx({
      rows: [
        [1, 10],
        [5, 50],
        [10, 90],
      ],
      stylesXml: PLAIN_STYLES,
      conditionalFormattingXml: cf,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
