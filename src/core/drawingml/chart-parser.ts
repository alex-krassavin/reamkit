// ECMA-376 Part 1 §21.2 — DrawingML charts (chart1.xml).
//
// Reads the chart's CACHED data (c:numCache / c:strCache), not the embedded
// spreadsheet — the cache holds the last-computed categories and values, which
// is exactly what Word renders. Supports bar/column, line and pie; other chart
// types parse as 'unknown' (the renderer reserves their box but draws nothing).

import { XMLParser } from 'fast-xml-parser';

import type { Chart, ChartDataPoint, ChartSeries, ChartType } from '@/core/document-model';
import type { OpcPackage } from '@/core/opc';
import type { ColorMod, ColorResolver } from '@/core/drawingml/colors';
import type { PoNode } from '@/core/po-helpers';
import { resolveColorNode } from '@/core/drawingml/colors';
import {
  poAttr,
  poChildren,
  poFindByPath,
  poFindDescendant,
  poIntAttr,
  poIs,
  poTag,
  poText,
  poVal,
} from '@/core/po-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

// Plot-area chart-group elements → our coarse ChartType.
const TYPE_OF_TAG: Readonly<Record<string, ChartType>> = {
  'c:barChart': 'bar',
  'c:bar3DChart': 'bar',
  'c:lineChart': 'line',
  'c:line3DChart': 'line',
  'c:pieChart': 'pie',
  'c:pie3DChart': 'pie',
  'c:doughnutChart': 'pie',
  'c:areaChart': 'area',
  'c:area3DChart': 'area',
  'c:scatterChart': 'scatter',
};

/**
 * Parse a DrawingML chart part (chart1.xml) into a {@link Chart}, reading the
 * CACHED data (`c:numCache` / `c:strCache`) rather than the embedded spreadsheet
 * — the cache holds the last-computed categories and values, exactly what Word
 * renders. Supports bar/column, line, pie/doughnut, area and scatter; other
 * chart types parse with `type: 'unknown'` (the renderer reserves the box but
 * draws nothing). Categories are shared and taken from the first series carrying
 * them.
 *
 * @param chartXml     The raw chart1.xml part bytes.
 * @param resolveColor Maps a DrawingML colour reference to a 6-hex string.
 * @returns The parsed chart, or `null` when there is no `c:chart` / `c:plotArea`.
 */
export function parseChart(chartXml: Uint8Array, resolveColor: ColorResolver): Chart | null {
  const tree = parser.parse(decoder.decode(chartXml)) as Array<PoNode>;
  const chart = poFindByPath(tree, ['c:chartSpace', 'c:chart']);
  if (!chart) return null;
  const plotArea = poChildren(chart).find((c) => poIs(c, 'c:plotArea'));
  if (!plotArea) return null;

  // §21.2.2.145 — a plot area holds a SEQUENCE of chart groups, not one. A combo
  // writes `c:barChart` and `c:lineChart` beside each other, and taking the first
  // dropped every series of the rest: 57362.xlsx printed its bars and lost its
  // line. The first group still gives the chart its type (and its bar direction,
  // grouping and gap); the others' series carry their own.
  const groups = poChildren(plotArea).filter((c) => (poTag(c) ?? '') in TYPE_OF_TAG);
  const group = groups[0];
  const type: ChartType = group ? (TYPE_OF_TAG[poTag(group)!] ?? 'unknown') : 'unknown';

  const serNodes: Array<PoNode> = [];
  const series: Array<ChartSeries> = [];
  for (const g of groups) {
    const groupType: ChartType = TYPE_OF_TAG[poTag(g)!] ?? 'unknown';
    for (const s of poChildren(g).filter((c) => poIs(c, 'c:ser'))) {
      serNodes.push(s);
      const parsed = parseSeries(s, resolveColor);
      series.push(groupType === type ? parsed : { ...parsed, type: groupType });
    }
  }

  // Categories are shared; take them from the first series that carries them.
  let categories: Array<string> = [];
  let categoriesRef: string | undefined;
  for (const s of serNodes) {
    const cat = poChildren(s).find((c) => poIs(c, 'c:cat'));
    if (cat) {
      categories = denseStrings(cat);
      categoriesRef ??= refFormula(cat);
      break;
    }
  }

  const barDir = group ? poVal(poChildren(group).find((c) => poIs(c, 'c:barDir'))) : undefined;
  const grouping = group ? poVal(poChildren(group).find((c) => poIs(c, 'c:grouping'))) : undefined;
  const doughnut = group ? poIs(group, 'c:doughnutChart') : false;
  const showValues = group ? chartShowsValues(group) : false;
  const catAxisTitle = axisTitle(plotArea, 'c:catAx');
  const valAxisTitle = axisTitle(plotArea, 'c:valAx');
  const valAxisMin = axisScaling(plotArea, 'c:min');
  const valAxisMax = axisScaling(plotArea, 'c:max');
  // §21.2.2.198 — the chart-space frame sits beside <c:chart>, not inside it.
  const chartSpace = tree.find((c) => poIs(c, 'c:chartSpace'));
  const spaceSpPr = poChildren(chartSpace).find((c) => poIs(c, 'c:spPr'));
  const frameFillHex = fillColorOf(spaceSpPr, resolveColor);
  const frameLine = spaceSpPr ? poChildren(spaceSpPr).find((c) => poIs(c, 'a:ln')) : undefined;
  const frameLineHex = frameLine ? fillColorOf(frameLine, resolveColor) : undefined;
  const numberFormat = valueFormatCode(plotArea);

  const legend = poChildren(chart).find((c) => poIs(c, 'c:legend'));
  const legendPos = legend
    ? poVal(poChildren(legend).find((c) => poIs(c, 'c:legendPos')))
    : undefined;

  const title = chartTitle(chart, series);
  // §21.2.2.75 — the gap between category slots, as a percentage of the bar
  // width. Unread, every bar took 0.63 of its slot; 57362.xlsx asks for 219 and
  // its bars came out 2.6× too wide. The schema default is 150.
  const gapNode = poChildren(plotArea)
    .flatMap((g) => poChildren(g))
    .find((c) => poIs(c, 'c:gapWidth'));
  const gapPercent = Number(poVal(gapNode) ?? '');

  return {
    type,
    ...(title ? { title } : {}),
    ...(Number.isFinite(gapPercent) && gapPercent >= 0 ? { gapPercent } : {}),
    categories,
    ...(categoriesRef ? { categoriesRef } : {}),
    series,
    hasLegend: legend !== undefined,
    ...(isLegendPos(legendPos) ? { legendPos } : {}),
    ...(barDir === 'col' || barDir === 'bar' ? { barDir } : {}),
    ...(isGrouping(grouping) ? { grouping } : {}),
    ...(doughnut ? { doughnut: true } : {}),
    ...(showValues ? { showValues: true } : {}),
    ...(catAxisTitle ? { catAxisTitle } : {}),
    ...(valAxisTitle ? { valAxisTitle } : {}),
    ...(valAxisMin !== undefined ? { valAxisMin } : {}),
    ...(valAxisMax !== undefined ? { valAxisMax } : {}),
    ...(frameFillHex ? { frameFillHex } : {}),
    ...(frameLineHex ? { frameLineHex } : {}),
    ...(numberFormat ? { numberFormat } : {}),
  };
}

function parseSeries(ser: PoNode, resolveColor: ColorResolver): ChartSeries {
  // Category charts carry values in c:val; scatter carries them in c:yVal with
  // the independent variable in c:xVal.
  const valNode =
    poChildren(ser).find((c) => poIs(c, 'c:val')) ?? poChildren(ser).find((c) => poIs(c, 'c:yVal'));
  const values = valNode ? denseNumbers(valNode) : [];
  const xValNode = poChildren(ser).find((c) => poIs(c, 'c:xVal'));
  const xValues = xValNode ? denseNumbers(xValNode) : undefined;
  const name = seriesName(ser);
  const colorHex = fillColorOf(
    poChildren(ser).find((c) => poIs(c, 'c:spPr')),
    resolveColor,
  );
  const pointColors = dataPointColors(ser, resolveColor);
  const pointLabels = customDataLabels(ser);
  // Keep the references so the reader can resolve them when nothing is cached.
  const valuesRef = valNode ? refFormula(valNode) : undefined;
  const nameRef = refFormula(poChildren(ser).find((c) => poIs(c, 'c:tx')));
  return {
    values,
    ...(valuesRef ? { valuesRef } : {}),
    ...(nameRef ? { nameRef } : {}),
    ...(xValues && xValues.length > 0 ? { xValues } : {}),
    ...(name ? { name } : {}),
    ...(colorHex ? { colorHex } : {}),
    ...(pointColors.length > 0 ? { pointColors } : {}),
    ...(pointLabels.length > 0 ? { pointLabels } : {}),
  };
}

function seriesName(ser: PoNode): string | undefined {
  const tx = poChildren(ser).find((c) => poIs(c, 'c:tx'));
  if (!tx) return undefined;
  const direct = poChildren(tx).find((c) => poIs(c, 'c:v'));
  if (direct) return poText(direct) || undefined;
  return readPts(tx)[0]?.v || undefined;
}

/**
 * §21.2.2.49 — the labels the author typed, by point index. A `<c:dLbl>` with
 * its own `<c:tx><c:rich>` replaces whatever the chart would have computed for
 * that point, and it is the only place that text exists.
 */
function customDataLabels(ser: PoNode): Array<{ idx: number; text: string }> {
  const dLbls = poChildren(ser).find((c) => poIs(c, 'c:dLbls'));
  if (!dLbls) return [];
  const out: Array<{ idx: number; text: string }> = [];
  for (const dLbl of poChildren(dLbls)) {
    if (!poIs(dLbl, 'c:dLbl')) continue;
    const idxNode = poChildren(dLbl).find((c) => poIs(c, 'c:idx'));
    const tx = poChildren(dLbl).find((c) => poIs(c, 'c:tx'));
    if (!tx) continue;
    const text = collectAT(tx).trim();
    if (text.length === 0) continue;
    out.push({ idx: idxNode ? (poIntAttr(idxNode, 'val') ?? 0) : 0, text });
  }
  return out;
}

function dataPointColors(ser: PoNode, resolveColor: ColorResolver): Array<ChartDataPoint> {
  const out: Array<ChartDataPoint> = [];
  for (const dPt of poChildren(ser)) {
    if (!poIs(dPt, 'c:dPt')) continue;
    const idxNode = poChildren(dPt).find((c) => poIs(c, 'c:idx'));
    const idx = idxNode ? (poIntAttr(idxNode, 'val') ?? 0) : 0;
    const colorHex = fillColorOf(
      poChildren(dPt).find((c) => poIs(c, 'c:spPr')),
      resolveColor,
    );
    if (colorHex) out.push({ idx, colorHex });
  }
  return out;
}

// Series colour from a c:spPr: the fill (direct a:solidFill, used by bars/pie)
// or, failing that, the outline (a:ln/a:solidFill, used by line charts).
function fillColorOf(spPr: PoNode | undefined, resolveColor: ColorResolver): string | undefined {
  if (!spPr) return undefined;
  const directFill = poChildren(spPr).find((c) => poIs(c, 'a:solidFill'));
  const fromFill = directFill ? colorFromSolidFill(directFill, resolveColor) : undefined;
  if (fromFill) return fromFill;
  // §20.1.8.33 a:gradFill — a series filled with a gradient still has a colour;
  // the scene model carries one per series, so take the first stop. Falling
  // through to the outline instead painted 123233_charts.xlsx's five gradient
  // bars in the black of their own hairline.
  const grad = poChildren(spPr).find((c) => poIs(c, 'a:gradFill'));
  const firstStop = grad
    ? poChildren(grad)
        .filter((c) => poIs(c, 'a:gsLst'))
        .flatMap((lst) => poChildren(lst))
        .find((gs) => poIs(gs, 'a:gs'))
    : undefined;
  const fromGrad = firstStop ? colorFromSolidFill(firstStop, resolveColor) : undefined;
  if (fromGrad) return fromGrad;
  const ln = poChildren(spPr).find((c) => poIs(c, 'a:ln'));
  const lnFill = ln ? poChildren(ln).find((c) => poIs(c, 'a:solidFill')) : undefined;
  return lnFill ? colorFromSolidFill(lnFill, resolveColor) : undefined;
}

function colorFromSolidFill(solid: PoNode, resolveColor: ColorResolver): string | undefined {
  for (const c of poChildren(solid)) {
    const isSrgb = poIs(c, 'a:srgbClr');
    if (!isSrgb && !poIs(c, 'a:schemeClr')) continue;
    if (!poAttr(c, 'val')) continue;
    // Chart semantics: stop at the first colour node, even when the resolver
    // does not know the colour (the word drawing-parser continues instead).
    return resolveColorNode(c, resolveColor);
  }
  return undefined;
}

// Concatenate every a:t run beneath a node (a c:title or rich-text body).
function collectAT(node: PoNode): string {
  let text = '';
  const walk = (n: PoNode): void => {
    for (const c of poChildren(n)) {
      if (poIs(c, 'a:t')) text += poText(c);
      else walk(c);
    }
  };
  walk(node);
  return text;
}

/**
 * §21.2.2.213 `c:title` — the chart's title, authored or generated.
 *
 * `c:tx` is OPTIONAL: a `<c:title>` with none asks the application to make the
 * title up, and §21.2.2.10's `c:autoTitleDeleted val="0"` says that generated
 * title has not been removed. Excel and LibreOffice agree on the rule — one
 * series means the series' name, anything else the placeholder "Chart Title" —
 * and reading only the `a:t` runs gave 56557.xlsx a blank page where the
 * reference draws its title. Eleven chart parts across eight corpus files are
 * written this way.
 *
 * The placeholder is an application string and therefore a locale: this file
 * was authored in Swedish, where Excel prints "Diagramrubrik". English is what
 * both references print here, and what we have to pick.
 *
 * @param chart  The `c:chart` element.
 * @param series The chart's series, for the single-series case.
 * @returns The title text, or undefined when there is no title to draw.
 */
function chartTitle(chart: PoNode, series: ReadonlyArray<{ name?: string }>): string | undefined {
  const title = poChildren(chart).find((c) => poIs(c, 'c:title'));
  if (!title) return undefined;
  const authored = collectAT(title) || undefined;
  if (authored) return authored;
  const deleted =
    poVal(poChildren(chart).find((c) => poIs(c, 'c:autoTitleDeleted'))) === '1';
  if (deleted) return undefined;
  return (series.length === 1 ? series[0]?.name : undefined) ?? 'Chart Title';
}

// c:catAx / c:valAx → c:title text.
//
// Like the chart's own title (§21.2.2.213), an axis title element with no
// `c:tx` asks the reader to generate one; Excel and LibreOffice both write the
// placeholder "Axis Title". 57362.xlsx leaves both of its value axes that way
// and we drew neither.
function axisTitle(plotArea: PoNode, axTag: string): string | undefined {
  const ax = poChildren(plotArea).find((c) => poIs(c, axTag));
  const title = ax ? poChildren(ax).find((c) => poIs(c, 'c:title')) : undefined;
  if (!title) return undefined;
  return collectAT(title) || 'Axis Title';
}

/**
 * §21.2.2.121 `c:numFmt` on the value axis — the number format its tick labels
 * and the chart's data labels are drawn in. It is the same code grammar cells
 * use (§18.8.31), so a currency axis reads as currency: without it a monthly
 * budget's axis ran 0/1000/2000 where every other reader shows $0/$1,000/$2,000.
 *
 * `sourceLinked="1"` means "whatever the source cells use", which the chart part
 * does not carry — those keep the plain numeric render.
 */
// §21.2.2.157 c:valAx/c:scaling/c:min|c:max — an axis end the author fixed.
function axisScaling(plotArea: PoNode, tag: 'c:min' | 'c:max'): number | undefined {
  const ax = poChildren(plotArea).find((c) => poIs(c, 'c:valAx'));
  const scaling = ax ? poChildren(ax).find((c) => poIs(c, 'c:scaling')) : undefined;
  const node = scaling ? poChildren(scaling).find((c) => poIs(c, tag)) : undefined;
  const v = node ? Number(poAttr(node, 'val')) : Number.NaN;
  return Number.isFinite(v) ? v : undefined;
}

function valueFormatCode(plotArea: PoNode): string | undefined {
  const ax = poChildren(plotArea).find((c) => poIs(c, 'c:valAx'));
  const numFmt = ax ? poChildren(ax).find((c) => poIs(c, 'c:numFmt')) : undefined;
  const code = numFmt ? poAttr(numFmt, 'formatCode') : undefined;
  if (code === undefined || code.trim().length === 0) return undefined;
  return code.trim().toLowerCase() === 'general' ? undefined : code;
}

// A c:dLbls with <c:showVal val="1"/> — group-level or on any series.
function dLblsShowVal(dLbls: PoNode | undefined): boolean {
  if (!dLbls) return false;
  const v = poVal(poChildren(dLbls).find((c) => poIs(c, 'c:showVal')));
  return v === '1' || v === 'true';
}

function chartShowsValues(group: PoNode): boolean {
  if (dLblsShowVal(poChildren(group).find((c) => poIs(c, 'c:dLbls')))) return true;
  for (const ser of poChildren(group)) {
    if (poIs(ser, 'c:ser') && dLblsShowVal(poChildren(ser).find((c) => poIs(c, 'c:dLbls'))))
      return true;
  }
  return false;
}

/** The `<c:f>` inside a c:val / c:cat / c:tx, or undefined when there is none. */
function refFormula(container: PoNode | undefined): string | undefined {
  if (!container) return undefined;
  const f = poFindDescendant(container, 'c:f');
  const text = f ? poText(f).trim() : '';
  return text.length > 0 ? text : undefined;
}

// Read c:pt entries from the numCache/strCache inside a c:cat / c:val / c:tx.
function readPts(container: PoNode): Array<{ idx: number; v: string }> {
  const cache =
    poFindDescendant(container, 'c:numCache') ?? poFindDescendant(container, 'c:strCache');
  if (!cache) return [];
  const out: Array<{ idx: number; v: string }> = [];
  for (const pt of poChildren(cache)) {
    if (!poIs(pt, 'c:pt')) continue;
    const idx = poIntAttr(pt, 'idx') ?? 0;
    const vNode = poChildren(pt).find((c) => poIs(c, 'c:v'));
    out.push({ idx, v: vNode ? poText(vNode) : '' });
  }
  return out;
}

function ptCountOf(container: PoNode): number {
  const cache =
    poFindDescendant(container, 'c:numCache') ?? poFindDescendant(container, 'c:strCache');
  const pc = cache ? poChildren(cache).find((c) => poIs(c, 'c:ptCount')) : undefined;
  return pc ? (poIntAttr(pc, 'val') ?? 0) : 0;
}

function denseLength(container: PoNode, pts: ReadonlyArray<{ idx: number }>): number {
  let max = ptCountOf(container);
  for (const p of pts) max = Math.max(max, p.idx + 1);
  return max;
}

function denseNumbers(container: PoNode): Array<number> {
  const pts = readPts(container);
  const arr = new Array<number>(denseLength(container, pts)).fill(0);
  for (const p of pts) {
    const n = Number(p.v);
    if (Number.isFinite(n)) arr[p.idx] = n;
  }
  return arr;
}

function denseStrings(container: PoNode): Array<string> {
  const pts = readPts(container);
  const arr = new Array<string>(denseLength(container, pts)).fill('');
  for (const p of pts) arr[p.idx] = p.v;
  return arr;
}

function isLegendPos(v: string | undefined): v is 'r' | 'l' | 't' | 'b' {
  return v === 'r' || v === 'l' || v === 't' || v === 'b';
}

function isGrouping(
  v: string | undefined,
): v is 'clustered' | 'stacked' | 'percentStacked' | 'standard' {
  return v === 'clustered' || v === 'stacked' || v === 'percentStacked' || v === 'standard';
}

/**
 * MS-ODRAWXML chartColorStyle (charts/colorsN.xml): the top-level colour list is
 * the series cycle (`meth="cycle"` — the common case; variations are luminance
 * tweaks for `>N` series and are ignored in v1).
 *
 * @param colorsXml    The raw colorsN.xml part bytes.
 * @param resolveColor Maps a DrawingML colour reference to a 6-hex string.
 * @returns The resolved series-colour cycle, in order (empty if none resolve).
 */
export function parseChartColorStyle(
  colorsXml: Uint8Array,
  resolveColor: ColorResolver,
): Array<string> {
  const tree = parser.parse(new TextDecoder().decode(colorsXml)) as Array<PoNode>;
  const root = tree.find((n) => {
    const tag = Object.keys(n).find((k) => k !== ':@' && k !== '#text');
    return tag !== undefined && tag.endsWith('colorStyle');
  });
  if (!root) return [];
  const out: Array<string> = [];
  for (const child of poChildren(root)) {
    const hex = resolveColorNode(child, resolveColor);
    if (hex !== undefined) out.push(hex);
  }
  return out;
}

/**
 * Augment a parsed {@link Chart} with its custom series-colour cycle when the
 * chart part's own relationships carry a chartColorStyle (colorsN.xml). Returns
 * the chart unchanged when no such relationship resolves to a non-empty cycle.
 * Shared by the docx and xlsx readers.
 *
 * @param chart         The parsed chart to augment.
 * @param pkg           The OPC package, for relationship lookup.
 * @param chartPartPath The chart part path, used as the relationship source.
 * @param resolveColor  Maps a DrawingML colour reference to a 6-hex string.
 * @returns The chart, with `seriesColorCycle` set when a cycle is found.
 */
export function withChartColorStyle(
  chart: Chart,
  pkg: OpcPackage,
  chartPartPath: string,
  resolveColor: ColorResolver,
): Chart {
  for (const rel of pkg.getPartRelationships(chartPartPath)) {
    if (rel.type !== REL_CHART_COLOR_STYLE) continue;
    const resolved = pkg.resolveRelatedPart(chartPartPath, rel);
    if (!resolved) continue;
    const cycle = parseChartColorStyle(resolved.data, resolveColor);
    if (cycle.length > 0) return { ...chart, seriesColorCycle: cycle };
  }
  return chart;
}

const REL_CHART_COLOR_STYLE =
  'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
