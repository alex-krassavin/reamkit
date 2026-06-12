// E-SHEET SC2 — sparklines: per-cell mini charts parsed from the worksheet
// extLst, resolved to a value series on the host cell, and drawn as vector
// geometry sized to the cell.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';
import { buildSparkline } from '@/core/drawingml/sparkline-geometry';

// x14 sparkline extLst. removeNSPrefix strips the x14:/xm: prefixes on read, so
// the prefixes here are cosmetic (they mirror a real Excel file).
const sparklineExt = (
  groups: Array<{
    type?: string;
    color?: string;
    cells: Array<{ f: string; sqref: string }>;
  }>,
): string => {
  const groupXml = groups
    .map((g) => {
      const typeAttr = g.type ? ` type="${g.type}"` : '';
      const colorXml = g.color ? `<x14:colorSeries rgb="FF${g.color}"/>` : '';
      const sparks = g.cells
        .map(
          (c) =>
            `<x14:sparkline><xm:f>${c.f}</xm:f><xm:sqref>${c.sqref}</xm:sqref></x14:sparkline>`,
        )
        .join('');
      return `<x14:sparklineGroup${typeAttr}>${colorXml}<x14:sparklines>${sparks}</x14:sparklines></x14:sparklineGroup>`;
    })
    .join('');
  return `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups>${groupXml}</x14:sparklineGroups></ext></extLst>`;
};

describe('sparklines — parse (E-SHEET SC2)', () => {
  it('reads sparkline groups from the worksheet extLst', () => {
    const ext = sparklineExt([
      { type: 'column', color: '376092', cells: [{ f: 'Sheet1!A1:D1', sqref: 'E1' }] },
    ]);
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1, 5, 3, 8]], extLstXml: ext }));
    const sparks = sheet.sheets[0]!.grid.sparklines;
    expect(sparks).toHaveLength(1);
    expect(sparks![0]).toEqual({
      kind: 'column',
      dataRange: 'Sheet1!A1:D1',
      sqref: 'E1',
      colorHex: '376092',
    });
  });

  it('defaults an untyped group to a line sparkline', () => {
    const ext = sparklineExt([{ cells: [{ f: 'A1:D1', sqref: 'E1' }] }]);
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1, 5, 3, 8]], extLstXml: ext }));
    expect(sheet.sheets[0]!.grid.sparklines![0]!.kind).toBe('line');
  });
});

describe('sparklines — projection (E-SHEET SC2)', () => {
  it('attaches the resolved value series to the host cell', () => {
    const ext = sparklineExt([{ type: 'line', cells: [{ f: 'A1:D1', sqref: 'E1' }] }]);
    const flow = Ream.parse(buildXlsx({ rows: [[1, 5, 3, 8]], extLstXml: ext })).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    // E1 = row 0, col 4 — empty, kept in the used range as the sparkline host.
    const sparkline = table.table.rows[0]?.cells[4]?.properties.sparkline;
    expect(sparkline?.kind).toBe('line');
    expect(sparkline?.values).toEqual([1, 5, 3, 8]);
  });

  it('resolves a vertical data range in row order', () => {
    const ext = sparklineExt([{ type: 'column', cells: [{ f: 'A1:A3', sqref: 'B1' }] }]);
    const flow = Ream.parse(buildXlsx({ rows: [[2], [9], [4]], extLstXml: ext })).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.table.rows[0]?.cells[1]?.properties.sparkline?.values).toEqual([2, 9, 4]);
  });
});

describe('sparklines — cross-sheet + gaps (Tail TC3)', () => {
  it('resolves a sheet-qualified data range against the named sheet', () => {
    // Sparkline on "Main" cell B1, data on a different sheet "Data".
    const ext = sparklineExt([{ type: 'line', cells: [{ f: 'Data!A1:C1', sqref: 'B1' }] }]);
    const flow = Ream.parse(
      buildXlsx({
        sheets: [
          { name: 'Main', rows: [['hdr']], extLstXml: ext },
          { name: 'Data', rows: [[7, 8, 9]] },
        ],
      }),
    ).flow;
    // Main is the first sheet → the first body table.
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.table.rows[0]?.cells[1]?.properties.sparkline?.values).toEqual([7, 8, 9]);
  });

  it('keeps a blank cell in the range as a gap (null)', () => {
    const ext = sparklineExt([{ type: 'line', cells: [{ f: 'A1:D1', sqref: 'E1' }] }]);
    const flow = Ream.parse(buildXlsx({ rows: [[1, null, 3, 4]], extLstXml: ext })).flow;
    const table = flow.body.find((el) => el.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected a table');
    expect(table.table.rows[0]?.cells[4]?.properties.sparkline?.values).toEqual([1, null, 3, 4]);
  });

  it('a line breaks across a gap (two move segments)', () => {
    const prims = buildSparkline('line', [1, null, 3], 56, 16);
    const segs = prims[0]!.paths[0]!.segments;
    expect(segs.filter((s) => s.op === 'move')).toHaveLength(2); // restart after the gap
    expect(segs.filter((s) => s.op === 'line')).toHaveLength(0);
  });

  it('a column skips a gap slot but keeps the others', () => {
    const prims = buildSparkline('column', [5, null, 7, 9], 56, 16);
    expect(prims[0]!.paths).toHaveLength(3); // 3 bars, the gap omitted
  });
});

describe('sparklines — geometry (E-SHEET SC2)', () => {
  it('builds a line as one stroked polyline through every point', () => {
    const prims = buildSparkline('line', [1, 5, 3, 8], 56, 16);
    expect(prims).toHaveLength(1);
    expect(prims[0]!.stroke).toBeDefined();
    expect(prims[0]!.fillColorHex).toBeUndefined();
    // 4 points → 1 move + 3 line segments.
    const segs = prims[0]!.paths[0]!.segments;
    expect(segs.filter((s) => s.op === 'move')).toHaveLength(1);
    expect(segs.filter((s) => s.op === 'line')).toHaveLength(3);
  });

  it('builds a column as one filled rect per value', () => {
    const prims = buildSparkline('column', [1, 5, 3, 8], 56, 16, 'AA0000');
    expect(prims).toHaveLength(1);
    expect(prims[0]!.fillColorHex).toBe('AA0000');
    expect(prims[0]!.paths).toHaveLength(4);
  });

  it('splits win/loss into a positive and a negative group', () => {
    const prims = buildSparkline('winLoss', [1, -2, 3, -4], 56, 16, '008000');
    expect(prims).toHaveLength(2);
    const [wins, losses] = prims;
    expect(wins!.fillColorHex).toBe('008000');
    expect(wins!.paths).toHaveLength(2); // 1, 3
    expect(losses!.paths).toHaveLength(2); // -2, -4
    expect(losses!.fillColorHex).not.toBe('008000');
  });

  it('returns nothing for an empty series or a zero-size box', () => {
    expect(buildSparkline('line', [], 56, 16)).toEqual([]);
    expect(buildSparkline('line', [1, 2], 0, 16)).toEqual([]);
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('sparklines — render smoke (E-SHEET SC2)', () => {
  it('renders line, column and win-loss sparklines to a valid PDF', () => {
    const ext = sparklineExt([
      { type: 'line', cells: [{ f: 'A1:D1', sqref: 'E1' }] },
      { type: 'column', cells: [{ f: 'A2:D2', sqref: 'E2' }] },
      { type: 'stacked', cells: [{ f: 'A3:D3', sqref: 'E3' }] },
    ]);
    const xlsx = buildXlsx({
      rows: [
        [1, 5, 3, 8],
        [4, 2, 7, 1],
        [1, -2, 3, -4],
      ],
      extLstXml: ext,
    });
    const pdf = convertXlsxToPdfSync(xlsx, { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});
