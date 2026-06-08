import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { convertDocxToPdfSync } from '@/converter';
import { defaultColorResolver } from '@/ooxml/drawingml/colors';
import { parseChart } from '@/ooxml/drawingml/chart-parser';
import { OpcPackage } from '@/opc';
import { parseDocument } from '@/ooxml/wordproc';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};
const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);
const enc = new TextEncoder();

const C_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const BAR_CHART = `<c:chartSpace ${C_NS}>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quarterly Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>2023</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>2024</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></c:spPr>
          <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>18</c:v></c:pt><c:pt idx="2"><c:v>25</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:axId val="111"/><c:axId val="222"/>
      </c:barChart>
      <c:catAx><c:axId val="111"/></c:catAx>
      <c:valAx><c:axId val="222"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;

function chartDrawing(rId: string, cx = 5486400, cy = 3200400): string {
  return `<w:p><w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="${cx}" cy="${cy}"/>
      <wp:docPr id="1" name="Chart 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                   r:id="${rId}"/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;
}

describe('parseChart', () => {
  it('reads a clustered column chart with two series', () => {
    const chart = parseChart(enc.encode(BAR_CHART), defaultColorResolver);
    expect(chart).not.toBeNull();
    expect(chart!.type).toBe('bar');
    expect(chart!.barDir).toBe('col');
    expect(chart!.grouping).toBe('clustered');
    expect(chart!.title).toBe('Quarterly Sales');
    expect(chart!.categories).toEqual(['Q1', 'Q2', 'Q3']);
    expect(chart!.hasLegend).toBe(true);
    expect(chart!.legendPos).toBe('b');
    expect(chart!.series).toHaveLength(2);
    expect(chart!.series[0]).toMatchObject({
      name: '2023',
      values: [10, 20, 15],
      colorHex: '4472C4',
    });
    expect(chart!.series[1]).toMatchObject({
      name: '2024',
      values: [12, 18, 25],
      colorHex: 'ED7D31',
    });
  });

  it('returns unknown type for an unsupported chart group', () => {
    const radar = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:radarChart/></c:plotArea></c:chart></c:chartSpace>`;
    expect(parseChart(enc.encode(radar), defaultColorResolver)!.type).toBe('unknown');
  });

  it('reads a scatter chart from c:xVal / c:yVal', () => {
    const scatter = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:scatterChart>
      <c:ser><c:idx val="0"/>
        <c:xVal><c:numRef><c:numCache><c:ptCount val="3"/>
          <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt>
        </c:numCache></c:numRef></c:xVal>
        <c:yVal><c:numRef><c:numCache><c:ptCount val="3"/>
          <c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt>
        </c:numCache></c:numRef></c:yVal>
      </c:ser>
    </c:scatterChart></c:plotArea></c:chart></c:chartSpace>`;
    const chart = parseChart(enc.encode(scatter), defaultColorResolver);
    expect(chart!.type).toBe('scatter');
    expect(chart!.series[0]!.values).toEqual([5, 9, 4]);
    expect(chart!.series[0]!.xValues).toEqual([1, 2, 3]);
  });

  it('reads data labels (showVal) and axis titles', () => {
    const withExtras = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:dLbls><c:showVal val="1"/></c:dLbls>
      <c:ser><c:idx val="0"/>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/>
          <c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
      </c:barChart>
      <c:catAx><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>
      <c:valAx><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const chart = parseChart(enc.encode(withExtras), defaultColorResolver)!;
    expect(chart.showValues).toBe(true);
    expect(chart.catAxisTitle).toBe('Quarter');
    expect(chart.valAxisTitle).toBe('Sales');
  });

  it('flags a doughnut chart (renders as a pie with a hole)', () => {
    const doughnut = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:doughnutChart>
      <c:ser><c:idx val="0"/>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/>
          <c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
    </c:doughnutChart></c:plotArea></c:chart></c:chartSpace>`;
    const chart = parseChart(enc.encode(doughnut), defaultColorResolver);
    expect(chart!.type).toBe('pie');
    expect(chart!.doughnut).toBe(true);
  });
});

describe('chart drawing parsing', () => {
  it('produces a chart BodyElement carrying the relationship id', () => {
    const docx = buildDocxFromBody(chartDrawing('rId5'));
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    expect(parsed).toHaveLength(1);
    expect(parsed[0]!.kind).toBe('chart');
    if (parsed[0]!.kind !== 'chart') throw new Error('unreachable');
    expect(parsed[0]!.chart.chartRelId).toBe('rId5');
    expect(parsed[0]!.chart.widthEmu).toBe(5486400);
    expect(parsed[0]!.chart.heightEmu).toBe(3200400);
  });
});

const LINE_CHART = `<c:chartSpace ${C_NS}><c:chart>
  <c:plotArea>
    <c:lineChart><c:grouping val="standard"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Alpha</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:spPr><a:ln><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln></c:spPr>
        <c:cat><c:strRef><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt><c:pt idx="3"><c:v>Apr</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="4"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>9</c:v></c:pt><c:pt idx="2"><c:v>7</c:v></c:pt><c:pt idx="3"><c:v>12</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:lineChart>
    <c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx>
  </c:plotArea>
</c:chart></c:chartSpace>`;

describe('line chart', () => {
  it('parses a line series colour from its outline (a:ln)', () => {
    const chart = parseChart(enc.encode(LINE_CHART), defaultColorResolver);
    expect(chart!.type).toBe('line');
    expect(chart!.series[0]).toMatchObject({
      name: 'Alpha',
      values: [5, 9, 7, 12],
      colorHex: '4472C4',
    });
    expect(chart!.categories).toEqual(['Jan', 'Feb', 'Mar', 'Apr']);
  });

  it('renders a stroked polyline in the series colour', () => {
    const docx = buildDocxFromBody(chartDrawing('rId7'), { charts: { rId7: LINE_CHART } });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // Series line stroked in 4472C4 (stroking RG, not fill rg).
    expect(text).toContain('0.266667 0.447059 0.768627 RG');
    expect(text).toContain(' l\n'); // polyline segments
    expect(text).toMatch(/<[0-9A-F]+> Tj/); // axis/category labels
  });
});

const PIE_CHART = `<c:chartSpace ${C_NS}><c:chart>
  <c:plotArea>
    <c:pieChart>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Red</c:v></c:pt><c:pt idx="1"><c:v>Green</c:v></c:pt><c:pt idx="2"><c:v>Blue</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>50</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt><c:pt idx="2"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea>
  <c:legend><c:legendPos val="r"/></c:legend>
</c:chart></c:chartSpace>`;

describe('pie chart', () => {
  it('parses a single-series pie with categories', () => {
    const chart = parseChart(enc.encode(PIE_CHART), defaultColorResolver);
    expect(chart!.type).toBe('pie');
    expect(chart!.categories).toEqual(['Red', 'Green', 'Blue']);
    expect(chart!.series[0]!.values).toEqual([50, 30, 20]);
  });

  it('renders filled wedges (arc Béziers) with a category legend', () => {
    const docx = buildDocxFromBody(chartDrawing('rId8'), { charts: { rId8: PIE_CHART } });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // Wedges are filled arc paths → contain Bézier curves.
    expect(text).toMatch(/ c\n/);
    // First slice in accent1 (4472C4) fill.
    expect(text).toContain('0.266667 0.447059 0.768627 rg');
    expect(text).toMatch(/<[0-9A-F]+> Tj/); // % labels / legend categories
  });
});

describe('column chart rendering (end-to-end)', () => {
  it('renders clustered bars in series colours with axes and labels', () => {
    const docx = buildDocxFromBody(chartDrawing('rId5'), { charts: { rId5: BAR_CHART } });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // Filled bars in each series colour (4472C4 / ED7D31).
    expect(text).toContain('0.266667 0.447059 0.768627 rg'); // 4472C4
    expect(text).toContain('0.929412 0.490196 0.192157 rg'); // ED7D31
    expect(text).toMatch(/\nh\nf\n/); // a filled bar rect
    // Axis lines stroked in 595959 (→ 0.34902).
    expect(text).toContain('0.34902 0.34902 0.34902 RG');
    // Labels (categories / ticks / title) rendered as text.
    expect(text).toMatch(/<[0-9A-F]+> Tj/);
  });
});
