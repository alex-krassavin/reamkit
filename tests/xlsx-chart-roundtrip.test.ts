// E-SHEET WT1 — embedded chart write-back. A sheet chart survives the full
// round-trip xlsx → SheetDoc → xlsx (the chart-serializer re-emits the drawing +
// chart parts), so the re-read FlowDoc carries the same chart data.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';

const C_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const BAR_CHART = `<c:chartSpace ${C_NS}>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sheet Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea><c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/><c:order val="0"/>
        <c:spPr><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill></c:spPr>
        <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
    </c:barChart></c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>`;

const ROWS = [
  ['Item', 'Qty'],
  ['A', 4],
  ['B', 9],
];

describe('xlsx embedded chart write-back (WT1)', () => {
  it('round-trips an embedded chart through xlsx → SheetDoc → xlsx', async () => {
    const xlsx = buildXlsx({ rows: ROWS, sheetChart: { chartXml: BAR_CHART } });
    const { bytes, losses } = await Ream.parse(xlsx).convertWithReport('xlsx');
    // No "charts not written back" loss any more.
    expect(losses.some((l) => /chart/i.test(l.detail))).toBe(false);

    const flow = Ream.parse(bytes).flow;
    expect(flow.charts?.size).toBe(1);
    const chart = flow.charts?.get('xl/charts/chart1.xml');
    expect(chart?.type).toBe('bar');
    expect(chart?.title).toBe('Sheet Sales');
    expect(chart?.barDir).toBe('col');
    expect(chart?.categories).toEqual(['A', 'B']);
    expect(chart?.series[0]?.values).toEqual([4, 9]);
    expect(chart?.series[0]?.colorHex).toBe('C0504D');
    expect(chart?.hasLegend).toBe(true);
    expect(flow.body.some((el) => el.kind === 'chart')).toBe(true);
  });
});
