// ECMA-376 Part 1 §21.2 — serialize a Chart back to a DrawingML chart part
// (chart1.xml). The inverse of chart-parser.ts: it emits exactly the cached
// data the parser reads (c:numCache / c:strCache), so a parse → serialize →
// parse round-trip preserves type, series, categories, colours, title, legend
// and axis titles. Shared by the xlsx writer (embedded charts, WT1) and the
// docx writer (drawing charts, WT3).

import type { Chart, ChartSeries } from '@/core/document-model';

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const GROUP_TAG: Readonly<Record<string, string>> = {
  bar: 'c:barChart',
  line: 'c:lineChart',
  pie: 'c:pieChart',
  area: 'c:areaChart',
  scatter: 'c:scatterChart',
  unknown: 'c:barChart',
};

const CAT_AX_ID = 111111111;
const VAL_AX_ID = 222222222;

/**
 * Serialize a {@link Chart} to a DrawingML chart part (chart1.xml) — the inverse
 * of `parseChart`. Emits exactly the cached data the parser reads
 * (`c:numCache` / `c:strCache`), so a parse → serialize → parse round-trip
 * preserves type, series, categories, colours, title, legend and axis titles.
 * Pie/doughnut charts omit axes; scatter emits two value axes (`c:valAx`);
 * everything else gets a category + value axis pair. Shared by the xlsx writer
 * (embedded charts, WT1) and the docx writer (drawing charts, WT3).
 *
 * @param chart The chart to serialize.
 * @returns The chart1.xml document as a string (with the XML declaration).
 */
export function chartSpaceXml(chart: Chart): string {
  const isScatter = chart.type === 'scatter';
  const isPie = chart.type === 'pie' || chart.doughnut === true;
  const groupTag = chart.doughnut ? 'c:doughnutChart' : (GROUP_TAG[chart.type] ?? 'c:barChart');

  const inner =
    (chart.type === 'bar' ? `<c:barDir val="${chart.barDir ?? 'col'}"/>` : '') +
    (chart.grouping ? `<c:grouping val="${chart.grouping}"/>` : '') +
    '<c:varyColors val="0"/>' +
    chart.series.map((s, i) => seriesXml(s, i, isScatter, chart.categories)).join('') +
    (chart.showValues ? '<c:dLbls><c:showVal val="1"/></c:dLbls>' : '') +
    (chart.doughnut ? '<c:holeSize val="50"/>' : '') +
    (isPie ? '' : `<c:axId val="${CAT_AX_ID}"/><c:axId val="${VAL_AX_ID}"/>`);

  const axes = isPie
    ? ''
    : isScatter
      ? valAx(CAT_AX_ID, 'b', VAL_AX_ID, chart.catAxisTitle) +
        valAx(VAL_AX_ID, 'l', CAT_AX_ID, chart.valAxisTitle)
      : catAx(CAT_AX_ID, VAL_AX_ID, chart.catAxisTitle) +
        valAx(VAL_AX_ID, 'l', CAT_AX_ID, chart.valAxisTitle);

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
    '<c:chart>' +
    titleXml(chart.title) +
    `<c:autoTitleDeleted val="${chart.title ? 0 : 1}"/>` +
    `<c:plotArea><c:layout/><${groupTag}>${inner}</${groupTag}>${axes}</c:plotArea>` +
    (chart.hasLegend
      ? `<c:legend><c:legendPos val="${chart.legendPos ?? 'r'}"/><c:overlay val="0"/></c:legend>`
      : '') +
    '<c:plotVisOnly val="1"/>' +
    '</c:chart>' +
    '</c:chartSpace>'
  );
}

function seriesXml(
  s: ChartSeries,
  idx: number,
  isScatter: boolean,
  categories: ReadonlyArray<string>,
): string {
  const tx = s.name
    ? `<c:tx><c:strRef><c:f>Sheet1!$A$${idx + 1}</c:f>` +
      `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(s.name)}</c:v></c:pt></c:strCache>` +
      '</c:strRef></c:tx>'
    : '';
  const spPr = s.colorHex
    ? `<c:spPr><a:solidFill><a:srgbClr val="${s.colorHex}"/></a:solidFill></c:spPr>`
    : '';
  const dPts = (s.pointColors ?? [])
    .map(
      (p) =>
        `<c:dPt><c:idx val="${p.idx}"/><c:bubble3D val="0"/>` +
        `<c:spPr><a:solidFill><a:srgbClr val="${p.colorHex}"/></a:solidFill></c:spPr></c:dPt>`,
    )
    .join('');
  const cat = isScatter
    ? `<c:xVal>${numRef('A', s.xValues ?? [])}</c:xVal>`
    : categories.length > 0
      ? `<c:cat>${strRef(categories)}</c:cat>`
      : '';
  const val = isScatter
    ? `<c:yVal>${numRef('B', s.values)}</c:yVal>`
    : `<c:val>${numRef('B', s.values)}</c:val>`;
  return (
    `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>` +
    tx +
    spPr +
    dPts +
    cat +
    val +
    '</c:ser>'
  );
}

function strRef(values: ReadonlyArray<string>): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('');
  return (
    '<c:strRef><c:f>Sheet1!$A$1</c:f>' +
    `<c:strCache><c:ptCount val="${values.length}"/>${pts}</c:strCache></c:strRef>`
  );
}

function numRef(col: string, values: ReadonlyArray<number>): string {
  const pts = values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('');
  return (
    `<c:numRef><c:f>Sheet1!$${col}$1</c:f>` +
    `<c:numCache><c:formatCode>General</c:formatCode>` +
    `<c:ptCount val="${values.length}"/>${pts}</c:numCache></c:numRef>`
  );
}

function titleXml(title: string | undefined): string {
  if (!title) return '';
  return (
    '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>' +
    `<a:p><a:r><a:t>${esc(title)}</a:t></a:r></a:p>` +
    '</c:rich></c:tx><c:overlay val="0"/></c:title>'
  );
}

function catAx(id: number, crossId: number, title: string | undefined): string {
  return (
    `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="b"/>${titleXml(title)}<c:crossAx val="${crossId}"/></c:catAx>`
  );
}

function valAx(id: number, pos: string, crossId: number, title: string | undefined): string {
  return (
    `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
    `<c:delete val="0"/><c:axPos val="${pos}"/>${titleXml(title)}<c:crossAx val="${crossId}"/></c:valAx>`
  );
}

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
