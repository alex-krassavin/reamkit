// E-DOCX WT3 (charts) — a docx chart survives docx → FlowDoc → docx: the body
// keeps an inline chart drawing and a re-serialized word/charts/chartN.xml
// carries the data (the shared chart-serializer, also used by the xlsx writer).

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';

const C_NS =
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const BAR_CHART = `<c:chartSpace ${C_NS}>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quarterly Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea><c:barChart>
      <c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/><c:order val="0"/>
        <c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>2024</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
        <c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt></c:numCache></c:numRef></c:val>
      </c:ser>
      <c:axId val="111"/><c:axId val="222"/>
    </c:barChart><c:catAx><c:axId val="111"/></c:catAx><c:valAx><c:axId val="222"/></c:valAx></c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>`;

const chartDrawing = (rId: string): string =>
  `<w:p><w:r><w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}"/>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing></w:r></w:p>`;

describe('docx chart write-back (WT3)', () => {
  it('round-trips a chart through docx → FlowDoc → docx', async () => {
    const docx = buildDocxFromBody(chartDrawing('rId50'), { charts: { rId50: BAR_CHART } });
    expect(Ream.parse(docx).flow.charts?.size).toBe(1);

    const { bytes, losses } = await Ream.parse(docx).convertWithReport('docx');
    expect(losses.some((l) => /chart/i.test(l.detail))).toBe(false);

    const flow = Ream.parse(bytes).flow;
    expect(flow.charts?.size).toBe(1);
    const chart = [...(flow.charts?.values() ?? [])][0];
    expect(chart?.type).toBe('bar');
    expect(chart?.title).toBe('Quarterly Sales');
    expect(chart?.categories).toEqual(['Q1', 'Q2', 'Q3']);
    expect(chart?.series[0]?.values).toEqual([10, 20, 15]);
    expect(chart?.series[0]?.name).toBe('2024');
    expect(flow.body.some((el) => el.kind === 'chart')).toBe(true);
  });
});
