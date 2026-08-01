import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';

import { convertDocxToPdfSync } from '@/core/converter';
import { defaultColorResolver } from '@/core/drawingml/colors';
import { buildChartScene } from '@/core/drawingml/chart-geometry';
import { parseChart } from '@/core/drawingml/chart-parser';
import { OpcPackage } from '@/core/opc';
import { parseDocument } from '@/word';
import { readDocx } from '@/word/docx-reader';

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

  it('keeps the value-axis ends the author fixed (§21.2.2.157)', () => {
    // A chart whose cells all read zero still has the axis its author pinned.
    // Scaling to the data drew shape-macro-ext-ref.xlsx's axis 0…1 where every
    // reader draws 0…300.
    const withScaling = BAR_CHART.replace(
      '<c:valAx><c:axId val="222"/></c:valAx>',
      '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/><c:max val="300"/></c:scaling></c:valAx>',
    );
    const chart = parseChart(enc.encode(withScaling), defaultColorResolver);
    expect(chart!.valAxisMax).toBe(300);
    expect(chart!.valAxisMin).toBeUndefined();
    // …and an axis the author left alone stays automatic.
    expect(parseChart(enc.encode(BAR_CHART), defaultColorResolver)!.valAxisMax).toBeUndefined();
  });

  it('reads the chart-space frame beside <c:chart> (§21.2.2.198)', () => {
    const framed = BAR_CHART.replace(
      '<c:chart>',
      '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr><c:chart>',
    );
    const chart = parseChart(enc.encode(framed), defaultColorResolver);
    expect(chart!.frameFillHex).toBe('FFFFFF');
    expect(chart!.frameLineHex).toBe('D9D9D9');
    // A chart that states no frame at all gets the one both references draw
    // around a plain chart: white inside a light grey rule (chart-prop.docx).
    const plain = parseChart(enc.encode(BAR_CHART), defaultColorResolver);
    expect(plain!.frameFillHex).toBe('FFFFFF');
    expect(plain!.frameLineHex).toBe('D9D9D9');
    // …but a chart that asks for neither gets neither (chart-dupe.docx).
    const bare = parseChart(
      enc.encode(BAR_CHART.replace('<c:chart>', '<c:spPr><a:noFill/></c:spPr><c:chart>')),
      defaultColorResolver,
    );
    expect(bare!.frameFillHex).toBeUndefined();
    expect(bare!.frameLineHex).toBeUndefined();
  });

  it('takes a gradient-filled series from its first stop, not from its outline', () => {
    // §20.1.8.33 — a series filled with a gradient still has a colour, and the
    // scene model carries one per series. Falling through to the outline
    // painted 123233_charts.xlsx's five gradient bars in the black of their
    // own hairline.
    const grad = BAR_CHART.replace(
      '<c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>',
      '<c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="8599D3"/></a:gs>' +
        '<a:gs pos="100000"><a:srgbClr val="5876AE"/></a:gs></a:gsLst></a:gradFill>' +
        '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>',
    );
    const chart = parseChart(enc.encode(grad), defaultColorResolver);
    expect(chart!.series[0]?.colorHex).toBe('8599D3');
  });

  it('keeps the references a cache-less chart reads its data from', () => {
    // A chart written without caches is not a chart without data — the reader
    // resolves these against the workbook (123233_charts.xlsx, four charts that
    // drew as empty axes).
    const noCache = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:barChart>
      <c:ser><c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:f>data!B1</c:f></c:strRef></c:tx>
        <c:cat><c:strRef><c:f>data!A2:A4</c:f></c:strRef></c:cat>
        <c:val><c:numRef><c:f>data!B2:B4</c:f></c:numRef></c:val>
      </c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    const chart = parseChart(enc.encode(noCache), defaultColorResolver);
    expect(chart!.series[0]?.valuesRef).toBe('data!B2:B4');
    expect(chart!.series[0]?.nameRef).toBe('data!B1');
    expect(chart!.categoriesRef).toBe('data!A2:A4');
    // A cached chart records them too, and keeps its cache.
    expect(parseChart(enc.encode(BAR_CHART), defaultColorResolver)!.series[0]?.values).toEqual([
      10, 20, 15,
    ]);
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

  it("draws the axis and its data labels in the axis's own number format", () => {
    // §21.2.2.121 c:valAx/c:numFmt. simple-monthly-budget.xlsx declares
    // `"$"#,##0` and we printed 0/1000/2000 down the axis and 3750 on the bar,
    // where every other reader shows $0/$1,000/$2,000 and $3,750.
    const money = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:dLbls><c:showVal val="1"/></c:dLbls>
      <c:ser><c:idx val="0"/>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/>
          <c:pt idx="0"><c:v>3750</c:v></c:pt><c:pt idx="1"><c:v>2336</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
      </c:barChart>
      <c:valAx><c:numFmt formatCode="&quot;$&quot;#,##0" sourceLinked="0"/></c:valAx>
    </c:plotArea></c:chart></c:chartSpace>`;
    const chart = parseChart(enc.encode(money), defaultColorResolver)!;
    expect(chart.numberFormat).toBe('"$"#,##0');
    const scene = buildChartScene(chart, 320, 200, (t, sz) => t.length * sz * 0.5);
    const texts = scene!.labels.map((t) => t.text);
    expect(texts).toContain('$3,750');
    expect(texts).toContain('$2,336');
    // The ticks carry it too — no bare 4000 anywhere on the axis.
    expect(texts.some((t) => /^\$[\d,]+$/.test(t) && t !== '$3,750' && t !== '$2,336')).toBe(true);
    expect(texts).not.toContain('4000');

    // General is not a format: it means "plain", and a chart that says so keeps
    // the numeric render.
    const general = parseChart(
      enc.encode(money.replace('&quot;$&quot;#,##0', 'General')),
      defaultColorResolver,
    )!;
    expect(general.numberFormat).toBeUndefined();
  });

  it('prints a data label the author typed, not the one it would compute', () => {
    // §21.2.2.49 — a <c:dLbl> carrying its own <c:tx><c:rich> replaces whatever
    // the chart would generate for that point, and it is the only place that
    // text exists. orderOfCNumFmtElements.xlsx labels every slice of its pie
    // with a sentence ("Промышленные потребители; 22,7млрд.кВтч; 67,3%") and we
    // drew a bare "67%" over each one.
    const pie = `<c:chartSpace ${C_NS}><c:chart><c:plotArea><c:pieChart>
      <c:ser><c:idx val="0"/>
        <c:dLbls>
          <c:dLbl><c:idx val="0"/><c:tx><c:rich><a:p><a:r><a:t>Big slice; 60%</a:t></a:r></a:p></c:rich></c:tx></c:dLbl>
          <c:showVal val="1"/>
        </c:dLbls>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/>
          <c:pt idx="0"><c:v>60</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt>
        </c:numCache></c:numRef></c:val>
      </c:ser>
      </c:pieChart></c:plotArea></c:chart></c:chartSpace>`;
    const chart = parseChart(enc.encode(pie), defaultColorResolver)!;
    expect(chart.series[0]?.pointLabels).toEqual([{ idx: 0, text: 'Big slice; 60%' }]);
    const texts = buildChartScene(chart, 320, 240, (t, sz) => t.length * sz * 0.5)!.labels.map(
      (l) => l.text,
    );
    expect(texts).toContain('Big slice; 60%');
    // The point with no label of its own keeps the computed one.
    expect(texts).toContain('40%');
    expect(texts).not.toContain('60%');
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
    expect(parsed[0]!.chart.width).toBe(emuToPt(5486400));
    expect(parsed[0]!.chart.height).toBe(emuToPt(3200400));
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

describe('a text box anchored beside text (chart-size.docx)', () => {
  // The box shares its paragraph with a run of body text, so the drawing is
  // read from the run — which used to parse it with no body parser and no
  // image resolver at all. The box came out empty: no "Before.", no chart, no
  // "After.".
  const boxed =
    `<w:p><w:r><w:drawing>` +
    `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>1000000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>100000</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="3943350" cy="1404620"/><wp:wrapNone/><wp:docPr id="2" name="Text Box 2"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:cNvSpPr txBox="1"/>` +
    `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3943350" cy="1404620"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
    `<wps:txbx><w:txbxContent>` +
    `<w:p><w:r><w:t>Before.</w:t></w:r></w:p>` +
    chartDrawing('rId5', 2162810, 1297940) +
    `<w:p><w:r><w:t>After.</w:t></w:r></w:p>` +
    `</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
    `</a:graphicData></a:graphic></wp:anchor>` +
    `</w:drawing></w:r><w:r><w:t>Body text.</w:t></w:r></w:p>`;

  it('parses the box with everything in it', () => {
    const { doc } = readDocx(buildDocxFromBody(boxed, { charts: { rId5: BAR_CHART } }));
    const shape = doc.body.find((b) => b.kind === 'shape');
    expect(
      shape?.kind === 'shape' ? shape.shape.text?.content.map((c) => c.kind) : undefined,
    ).toEqual(['paragraph', 'chart', 'paragraph']);
  });

  it('draws the chart inside it', () => {
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(boxed, { charts: { rId5: BAR_CHART } }), {
        fonts: FONTS,
      }),
    );
    expect(text).toContain('0.266667 0.447059 0.768627 rg'); // 4472C4 bars
    expect(text).toContain('0.929412 0.490196 0.192157 rg'); // ED7D31 bars
  });
});

describe('a chart in a paragraph that holds text too (chart-dupe.docx)', () => {
  // A lone drawing collapses to a block; anything else in the paragraph kept
  // the drawing in the run, where only pictures render — the chart was
  // dropped. chart-dupe.docx sets one beside a trailing space.
  const body = chartDrawing('rId5').replace(
    '</w:r></w:p>',
    `</w:r><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>`,
  );

  it('keeps the chart, as the block it is, and the paragraph beside it', () => {
    const { doc } = readDocx(buildDocxFromBody(body, { charts: { rId5: BAR_CHART } }));
    expect(doc.body.map((b) => b.kind)).toEqual(['chart', 'paragraph']);
    const first = doc.body[0];
    expect(first?.kind === 'chart' ? first.chart.chartRelId : undefined).toBe(
      'word/charts/chart1.xml',
    );
  });

  it('draws it', () => {
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(body, { charts: { rId5: BAR_CHART } }), {
        fonts: FONTS,
      }),
    );
    expect(text).toContain('0.266667 0.447059 0.768627 rg'); // 4472C4 bars
    expect(text).toMatch(/\nh\nf\n/);
  });
});

describe('a chart anchored in a footer (chart-in-footer.docx)', () => {
  // Charts were collected from the main document's rels only, and the band
  // renderer drew paragraphs, tables and images but not charts — a document
  // whose only content is a footer chart came out blank. Relationship ids are
  // scoped to their owning part (OPC §9.3), so the footer's rId5 here is a
  // different chart from the body's.
  const docx = buildDocxFromBody(
    `<w:p><w:r><w:t>body</w:t></w:r></w:p>` +
      `<w:sectPr><w:footerReference w:type="default" r:id="rId11"/></w:sectPr>`,
    {
      footerXml: chartDrawing('rId5'),
      footerCharts: { rId5: BAR_CHART },
      charts: { rId5: PIE_CHART },
    },
  );

  it('draws the footer chart on every page', () => {
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toContain('0.266667 0.447059 0.768627 rg'); // 4472C4 bars
    expect(text).toContain('0.929412 0.490196 0.192157 rg'); // ED7D31 bars
    expect(text).toMatch(/\nh\nf\n/); // a filled bar rect
  });

  it('keeps the body and the footer rId5 apart', () => {
    const { doc } = readDocx(docx);
    // Two chart parts, filed under their part paths rather than a shared id.
    expect([...(doc.charts?.keys() ?? [])].sort()).toEqual([
      'word/charts/chart1.xml',
      'word/charts/chart2.xml',
    ]);
    const footer = doc.headersFooters?.get('rId11');
    const block = footer?.find((b) => b.kind === 'chart');
    expect(block?.kind === 'chart' ? block.chart.chartRelId : undefined).toBe(
      'word/charts/chart2.xml',
    );
    expect(doc.charts?.get('word/charts/chart2.xml')?.type).toBe('bar');
  });
});

describe('an auto-generated chart title (§21.2.2.213)', () => {
  const chartXml = (title: string, sers: string): string =>
    `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<c:chart>${title}<c:plotArea><c:barChart><c:barDir val="col"/>${sers}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  const oneSeries =
    `<c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Demo</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
    `<c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>`;

  it('takes the series name when there is exactly one, and the placeholder otherwise', () => {
    // `c:tx` is optional: a `<c:title>` without one asks the application to make
    // the title up, and §21.2.2.10's `autoTitleDeleted val="0"` says the made-up
    // title stands. Reading only the `a:t` runs left it undefined, and eleven
    // chart parts across eight corpus files are written this way.
    const single = parseChart(
      enc.encode(
        chartXml('<c:title><c:layout/></c:title><c:autoTitleDeleted val="0"/>', oneSeries),
      ),
      defaultColorResolver,
    );
    expect(single?.title).toBe('Demo');
    const two = parseChart(
      enc.encode(
        chartXml('<c:title><c:layout/></c:title>', oneSeries + oneSeries.replace('Demo', 'Other')),
      ),
      defaultColorResolver,
    );
    expect(two?.title).toBe('Chart Title');
  });

  it('stays silent when the author deleted it, or asked for no title at all', () => {
    expect(
      parseChart(
        enc.encode(
          chartXml('<c:title><c:layout/></c:title><c:autoTitleDeleted val="1"/>', oneSeries),
        ),
        defaultColorResolver,
      )?.title,
    ).toBeUndefined();
    expect(
      parseChart(enc.encode(chartXml('', oneSeries)), defaultColorResolver)?.title,
    ).toBeUndefined();
  });
});
