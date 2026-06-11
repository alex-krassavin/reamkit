import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const C_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const BAR_CHART = `<c:chartSpace ${C_NS}>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sheet Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea><c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/><c:order val="0"/>
        <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
    </c:barChart></c:plotArea>
  </c:chart>
</c:chartSpace>`;

// A custom two-colour cycle: first series must come out 112233.
const COLORS = `<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
  <a:srgbClr val="112233"/>
  <a:srgbClr val="445566"/>
</cs:colorStyle>`;

const ROWS = [
  ['Item', 'Qty'],
  ['A', 4],
  ['B', 9],
];

describe('charts on xlsx sheets (§20.5 SpreadsheetDrawingML)', () => {
  it('loads the drawing part and renders the chart after the grid', () => {
    const xlsx = buildXlsx({ rows: ROWS, sheetChart: { chartXml: BAR_CHART } });
    const flow = Ream.parse(xlsx).flow;
    expect(flow.charts?.size).toBe(1);
    const chartBlock = flow.body.find((el) => el.kind === 'chart');
    expect(chartBlock).toBeDefined();
    if (chartBlock?.kind !== 'chart') throw new Error('unreachable');
    expect(chartBlock.chart.chartRelId).toBe('xl/charts/chart1.xml');
    // twoCellAnchor B2..H17: 6 default columns × 48pt, 15 default rows × 15pt.
    expect(chartBlock.chart.width).toBeCloseTo(6 * 48, 0);
    expect(chartBlock.chart.height).toBeCloseTo(15 * 15, 0);
    // The chart block comes after the sheet's table in the body.
    const tableIdx = flow.body.findIndex((el) => el.kind === 'table');
    const chartIdx = flow.body.findIndex((el) => el.kind === 'chart');
    expect(chartIdx).toBeGreaterThan(tableIdx);
  });

  it('renders chart geometry into the PDF page commands', () => {
    const xlsx = buildXlsx({ rows: ROWS, sheetChart: { chartXml: BAR_CHART } });
    const flow = Ream.parse(xlsx).flow;
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const shapes = laid.pages[0]!.commands.filter((c) => c.type === 'shape');
    expect(shapes.length).toBeGreaterThan(2); // bars + axes
    // The chart title text rides along as a line command.
    const text = laid.pages[0]!.commands.filter((c) => c.type === 'line')
      .map((c) =>
        (c as unknown as { line: { tokens: ReadonlyArray<{ text?: string }> } }).line.tokens
          .map((t) => t.text ?? '')
          .join(''),
      )
      .join('\n');
    expect(text).toContain('Sheet Sales');
  });

  it('applies the colorsN.xml series cycle over the default palette', () => {
    const xlsx = buildXlsx({
      rows: ROWS,
      sheetChart: { chartXml: BAR_CHART, colorsXml: COLORS },
    });
    const flow = Ream.parse(xlsx).flow;
    const chart = flow.charts?.get('xl/charts/chart1.xml');
    expect(chart?.seriesColorCycle).toEqual(['112233', '445566']);
    const laid = layoutStyledDocument(flow.body, {
      registry: FontRegistry.fromBytes(FONTS),
      ...flowRenderOptions(flow),
    });
    const fills = laid.pages[0]!.commands.filter((c) => c.type === 'shape').map(
      (c) => (c as unknown as { shape: { fillColorHex?: string } }).shape.fillColorHex,
    );
    expect(fills).toContain('112233'); // bars use the custom cycle
    expect(fills).not.toContain('4472C4'); // not the built-in accent palette
  });

  it('sheets without drawings are untouched', () => {
    const xlsx = buildXlsx({ rows: ROWS });
    const flow = Ream.parse(xlsx).flow;
    expect(flow.charts).toBeUndefined();
    expect(flow.body.some((el) => el.kind === 'chart')).toBe(false);
  });
});
