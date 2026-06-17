// XLS-6 — embedded charts in a legacy `.xls`. A chart is a nested BOF…EOF
// substream of chart records that reference worksheet ranges through AI records;
// the values are read straight from the sheet's cells (the cached-grid approach),
// the chart type from the group record, and the result maps onto the Chart model
// the renderer already draws.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXls, chartRecords, labelSstRec, numberRec } from './fixtures/build-xls';
import { parseBiffChart } from '@/excel/xls/biff-chart';
import { readXlsToSheetDoc } from '@/excel/xls/biff-reader';
import { Ream } from '@/core/converter/ream';

// A workbook with A1:A3 = values and B1:B3 = category labels, plus a column chart
// plotting them.
const chartXls = (kind?: 'bar' | 'line' | 'pie'): Uint8Array =>
  buildXls({
    sst: ['Q1', 'Q2', 'Q3'],
    sheets: [
      {
        name: 'S',
        records: [
          numberRec(0, 0, 10),
          numberRec(1, 0, 20),
          numberRec(2, 0, 30),
          labelSstRec(0, 1, 0),
          labelSstRec(1, 1, 1),
          labelSstRec(2, 1, 2),
          ...chartRecords({
            kind,
            values: { r0: 0, r1: 2, c0: 0, c1: 0 },
            categories: { r0: 0, r1: 2, c0: 1, c1: 1 },
            name: 'Sales',
          }),
        ],
      },
    ],
  });

describe('xls embedded charts (XLS-6)', () => {
  it('plots a chart from the worksheet cells its AI records reference', () => {
    const doc = readXlsToSheetDoc(chartXls());
    const ref = doc.sheets[0]?.charts?.[0];
    expect(ref).toBeDefined();
    const chart = doc.chartData?.get(ref!.chartPartPath);
    expect(chart?.type).toBe('bar');
    expect(chart?.barDir).toBe('col');
    expect(chart?.series).toHaveLength(1);
    expect(chart?.series[0]?.values).toEqual([10, 20, 30]);
    expect(chart?.series[0]?.name).toBe('Sales');
    expect(chart?.categories).toEqual(['Q1', 'Q2', 'Q3']);
  });

  it('reads the chart type from its group record', () => {
    expect(readChartType('line')).toBe('line');
    expect(readChartType('pie')).toBe('pie');
  });

  it('does not truncate the sheet at the nested chart EOF', () => {
    // A cell after the chart substream must still be read (depth-tracked substream).
    const xls = buildXls({
      sheets: [
        {
          name: 'S',
          records: [
            numberRec(0, 0, 5),
            ...chartRecords({ kind: 'bar', values: { r0: 0, r1: 0, c0: 0, c1: 0 } }),
            numberRec(9, 9, 99),
          ],
        },
      ],
    });
    const cells = readXlsToSheetDoc(xls).sheets[0]!.grid.cells;
    expect(cells.find((c) => c.row === 9 && c.column === 9)?.rawValue).toBe('99');
  });

  it('renders an .xls with an embedded chart to a valid PDF', async () => {
    const pdf = await Ream.parse(chartXls()).convert('pdf', {
      fonts: {
        regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
        bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
      },
    });
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});

function readChartType(kind: 'line' | 'pie'): string | undefined {
  const doc = readXlsToSheetDoc(chartXls(kind));
  const ref = doc.sheets[0]?.charts?.[0];
  return ref ? doc.chartData?.get(ref.chartPartPath)?.type : undefined;
}

// Direct parser unit: a hand-built chart substream over a tiny cell set.
describe('parseBiffChart unit (XLS-6)', () => {
  it('resolves series values + names without the full reader', () => {
    const recs = chartRecords({
      values: { r0: 0, r1: 1, c0: 0, c1: 0 },
      name: 'X',
    }).map((r) => ({ type: (r[0]! | (r[1]! << 8)) >>> 0, data: r.subarray(4) }));
    const cells = [
      { row: 0, column: 0, type: 'n' as const, rawValue: '7' },
      { row: 1, column: 0, type: 'n' as const, rawValue: '8' },
    ];
    const chart = parseBiffChart(recs, cells, []);
    expect(chart?.series[0]?.values).toEqual([7, 8]);
    expect(chart?.series[0]?.name).toBe('X');
  });
});
