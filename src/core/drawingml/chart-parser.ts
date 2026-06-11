// ECMA-376 Part 1 §21.2 — DrawingML charts (chart1.xml).
//
// Reads the chart's CACHED data (c:numCache / c:strCache), not the embedded
// spreadsheet — the cache holds the last-computed categories and values, which
// is exactly what Word renders. Supports bar/column, line and pie; other chart
// types parse as 'unknown' (the renderer reserves their box but draws nothing).

import { XMLParser } from 'fast-xml-parser';

import type { Chart, ChartDataPoint, ChartSeries, ChartType } from '@/core/document-model';
import type { ColorMod, ColorResolver } from '@/core/drawingml/colors';
import type { PoNode } from '@/core/po-helpers';
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

export function parseChart(chartXml: Uint8Array, resolveColor: ColorResolver): Chart | null {
  const tree = parser.parse(decoder.decode(chartXml)) as Array<PoNode>;
  const chart = poFindByPath(tree, ['c:chartSpace', 'c:chart']);
  if (!chart) return null;
  const plotArea = poChildren(chart).find((c) => poIs(c, 'c:plotArea'));
  if (!plotArea) return null;

  const group = poChildren(plotArea).find((c) => (poTag(c) ?? '') in TYPE_OF_TAG);
  const type: ChartType = group ? (TYPE_OF_TAG[poTag(group)!] ?? 'unknown') : 'unknown';

  const serNodes = group ? poChildren(group).filter((c) => poIs(c, 'c:ser')) : [];
  const series = serNodes.map((s) => parseSeries(s, resolveColor));

  // Categories are shared; take them from the first series that carries them.
  let categories: Array<string> = [];
  for (const s of serNodes) {
    const cat = poChildren(s).find((c) => poIs(c, 'c:cat'));
    if (cat) {
      categories = denseStrings(cat);
      break;
    }
  }

  const barDir = group ? poVal(poChildren(group).find((c) => poIs(c, 'c:barDir'))) : undefined;
  const grouping = group ? poVal(poChildren(group).find((c) => poIs(c, 'c:grouping'))) : undefined;
  const doughnut = group ? poIs(group, 'c:doughnutChart') : false;
  const showValues = group ? chartShowsValues(group) : false;
  const catAxisTitle = axisTitle(plotArea, 'c:catAx');
  const valAxisTitle = axisTitle(plotArea, 'c:valAx');

  const legend = poChildren(chart).find((c) => poIs(c, 'c:legend'));
  const legendPos = legend
    ? poVal(poChildren(legend).find((c) => poIs(c, 'c:legendPos')))
    : undefined;

  const title = chartTitle(chart);

  return {
    type,
    ...(title ? { title } : {}),
    categories,
    series,
    hasLegend: legend !== undefined,
    ...(isLegendPos(legendPos) ? { legendPos } : {}),
    ...(barDir === 'col' || barDir === 'bar' ? { barDir } : {}),
    ...(isGrouping(grouping) ? { grouping } : {}),
    ...(doughnut ? { doughnut: true } : {}),
    ...(showValues ? { showValues: true } : {}),
    ...(catAxisTitle ? { catAxisTitle } : {}),
    ...(valAxisTitle ? { valAxisTitle } : {}),
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
  return {
    values,
    ...(xValues && xValues.length > 0 ? { xValues } : {}),
    ...(name ? { name } : {}),
    ...(colorHex ? { colorHex } : {}),
    ...(pointColors.length > 0 ? { pointColors } : {}),
  };
}

function seriesName(ser: PoNode): string | undefined {
  const tx = poChildren(ser).find((c) => poIs(c, 'c:tx'));
  if (!tx) return undefined;
  const direct = poChildren(tx).find((c) => poIs(c, 'c:v'));
  if (direct) return poText(direct) || undefined;
  return readPts(tx)[0]?.v || undefined;
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
  const ln = poChildren(spPr).find((c) => poIs(c, 'a:ln'));
  const lnFill = ln ? poChildren(ln).find((c) => poIs(c, 'a:solidFill')) : undefined;
  return lnFill ? colorFromSolidFill(lnFill, resolveColor) : undefined;
}

function colorMods(colorNode: PoNode): Array<ColorMod> {
  const mods: Array<ColorMod> = [];
  for (const c of poChildren(colorNode)) {
    for (const kind of ['lumMod', 'lumOff', 'shade', 'tint', 'alpha'] as const) {
      if (poIs(c, `a:${kind}`)) {
        const v = poIntAttr(c, 'val');
        if (v !== undefined) mods.push({ kind, val: v / 100000 });
      }
    }
  }
  return mods;
}

function colorFromSolidFill(solid: PoNode, resolveColor: ColorResolver): string | undefined {
  for (const c of poChildren(solid)) {
    const isSrgb = poIs(c, 'a:srgbClr');
    if (!isSrgb && !poIs(c, 'a:schemeClr')) continue;
    const v = poAttr(c, 'val');
    if (!v) continue;
    const mods = colorMods(c);
    const raw = isSrgb ? { srgb: v } : { scheme: v };
    return resolveColor(mods.length > 0 ? { ...raw, mods } : raw);
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

function chartTitle(chart: PoNode): string | undefined {
  const title = poChildren(chart).find((c) => poIs(c, 'c:title'));
  return title ? collectAT(title) || undefined : undefined;
}

// c:catAx / c:valAx → c:title text.
function axisTitle(plotArea: PoNode, axTag: string): string | undefined {
  const ax = poChildren(plotArea).find((c) => poIs(c, axTag));
  const title = ax ? poChildren(ax).find((c) => poIs(c, 'c:title')) : undefined;
  return title ? collectAT(title) || undefined : undefined;
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
