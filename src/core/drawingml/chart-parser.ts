// ECMA-376 Part 1 §21.2 — DrawingML charts (chart1.xml).
//
// Reads the chart's CACHED data (c:numCache / c:strCache), not the embedded
// spreadsheet — the cache holds the last-computed categories and values, which
// is exactly what Word renders. Supports bar/column, line and pie; other chart
// types parse as 'unknown' (the renderer reserves their box but draws nothing).

import { XMLParser } from 'fast-xml-parser';

import type {
  Chart,
  ChartDataLabels,
  ChartDataPoint,
  ChartLabelPosition,
  ChartLineStyle,
  ChartMarker,
  ChartMarkerSymbol,
  ChartSeries,
  ChartType,
  ShapeDash,
} from '@/core/document-model';
import type { ResourceId } from '@/core/ir';
import type { OpcPackage } from '@/core/opc';
import type { ColorMod, ColorResolver } from '@/core/drawingml/colors';
import type { ChartTextDefaults } from '@/core/drawingml/chart-text';
import type { PoNode } from '@/core/po-helpers';
import { chartTextStyles } from '@/core/drawingml/chart-text';
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
 * The most points a chart keeps, all its series together: as many as a sheet
 * has rows, the longest range a series can name. A chart part states its own
 * point counts — a cache's `c:ptCount`, each point's `idx`, the range of a
 * `c:f` — and taken at their word, a count of two billion in a part a few
 * hundred bytes long asked for a dense array that ran the process out of
 * memory, and one past 2³² − 1 threw.
 */
export const MOST_CHART_POINTS = 1_048_576;

/**
 * How many points each of a chart's `seriesCount` series keeps: an even share
 * of {@link MOST_CHART_POINTS}, so that the bound holds for the chart and not
 * series by series — two thousand series, each declaring a sheet's worth of
 * points, are two billion. A series' points past its share are not read.
 *
 * @param seriesCount How many series the chart has.
 * @returns The points each keeps, at least one.
 */
export function pointsPerSeries(seriesCount: number): number {
  return Math.max(1, Math.floor(MOST_CHART_POINTS / Math.max(1, seriesCount)));
}

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
 * @param resolveImage Maps a blip's relationship id to its image resource.
 * @param textDefaults The host's defaults for chart text; given, the chart's
 *                     text is resolved role by role ({@link Chart.text}).
 * @returns The parsed chart, or `null` when there is no `c:chart` / `c:plotArea`.
 */
export function parseChart(
  chartXml: Uint8Array,
  resolveColor: ColorResolver,
  resolveImage?: (relId: string) => ResourceId | undefined,
  textDefaults?: ChartTextDefaults,
): Chart | null {
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

  // §21.2.2.9 — each group names the two axes it plots against. The first
  // group's value axis is the primary one; a group naming another is on the
  // secondary axis, drawn opposite (57362.xlsx's line, at `axPos="r"`).
  const groupAxIds = (g: PoNode): Array<string> =>
    poChildren(g)
      .filter((c) => poIs(c, 'c:axId'))
      .map((c) => poAttr(c, 'val') ?? '');
  const primaryAxIds = new Set(group ? groupAxIds(group) : []);
  // Each series' share of the chart's points, known before any is read.
  const most = pointsPerSeries(
    groups.reduce((n, g) => n + poChildren(g).filter((c) => poIs(c, 'c:ser')).length, 0),
  );
  const serNodes: Array<PoNode> = [];
  const series: Array<ChartSeries> = [];
  let secondaryValAxId: string | undefined;
  for (const g of groups) {
    const groupType: ChartType = TYPE_OF_TAG[poTag(g)!] ?? 'unknown';
    const own = groupAxIds(g).filter((id) => !primaryAxIds.has(id));
    const secondary = own.length > 0;
    for (const s of poChildren(g).filter((c) => poIs(c, 'c:ser'))) {
      serNodes.push(s);
      series.push({
        ...parseSeries(s, resolveColor, most),
        ...(groupType === type ? {} : { type: groupType }),
        ...(secondary ? { secondaryAxis: true as const } : {}),
      });
    }
    if (secondary && series.length > 0) {
      // Of the pair, the one a `c:valAx` claims is the value axis.
      secondaryValAxId ??= own.find((id) =>
        poChildren(plotArea).some(
          (c) =>
            poIs(c, 'c:valAx') &&
            poChildren(c).some((k) => poIs(k, 'c:axId') && poAttr(k, 'val') === id),
        ),
      );
    }
  }
  const secondaryValAx = secondaryValAxId
    ? poChildren(plotArea).find(
        (c) =>
          poIs(c, 'c:valAx') &&
          poChildren(c).some((k) => poIs(k, 'c:axId') && poAttr(k, 'val') === secondaryValAxId),
      )
    : undefined;
  // §21.2.2.40 `c:delete` — an axis the author hid is not drawn.
  const secondaryDeleted =
    secondaryValAx !== undefined &&
    poChildren(secondaryValAx).some((c) => poIs(c, 'c:delete') && poAttr(c, 'val') === '1');
  const secondaryTitleNode =
    secondaryValAx && !secondaryDeleted
      ? poChildren(secondaryValAx).find((c) => poIs(c, 'c:title'))
      : undefined;
  const secondaryValAxisTitle = secondaryTitleNode
    ? collectAT(secondaryTitleNode) || 'Axis Title'
    : undefined;

  // Categories are shared; take them from the first series that carries them.
  let categories: Array<string> = [];
  let categoriesRef: string | undefined;
  let categoryGroups: Chart['categoryGroups'];
  for (const s of serNodes) {
    const cat = poChildren(s).find((c) => poIs(c, 'c:cat'));
    if (cat) {
      const levels = multiLevelCategories(cat, most);
      categories = levels ? levels.categories : denseStrings(cat, most);
      categoryGroups = levels?.groups;
      categoriesRef ??= refFormula(cat);
      break;
    }
  }

  // §21.2.2.161 — a scatter joins its points, marks them, or both.
  const scatterStyleRaw = group
    ? poVal(poChildren(group).find((c) => poIs(c, 'c:scatterStyle')))
    : undefined;
  const scatterStyle = SCATTER_STYLES.has(scatterStyleRaw ?? '')
    ? (scatterStyleRaw as Chart['scatterStyle'])
    : undefined;
  // §21.2.2.106 — on a LINE group `c:marker` is a switch, not a symbol: it says
  // whether the group's series stamp their points at all.
  const lineMarkers =
    group && poIs(group, 'c:lineChart')
      ? poVal(poChildren(group).find((c) => poIs(c, 'c:marker'))) === '1'
      : false;
  const barDir = group ? poVal(poChildren(group).find((c) => poIs(c, 'c:barDir'))) : undefined;
  const grouping = group ? poVal(poChildren(group).find((c) => poIs(c, 'c:grouping'))) : undefined;
  const doughnut = group ? poIs(group, 'c:doughnutChart') : false;
  // §21.2.2.143/§21.2.2.228 — a 3-D pie is a tilted disc, by its view's
  // elevation and turn; Excel writes `rotX="30"` for the one it inserts, and
  // that is its view where the part states none.
  const view3D = poChildren(chart).find((c) => poIs(c, 'c:view3D'));
  const viewAngle = (tag: string, fallback: number): number => {
    const v = Number(poVal(view3D ? poChildren(view3D).find((c) => poIs(c, tag)) : undefined));
    return Number.isFinite(v) ? v : fallback;
  };
  // An absent `c:rotX` is the schema's 0 — the disc seen edge on, as Excel's
  // own PDF draws a 3-D pie whose view says only its perspective.
  const pie3D =
    group && poIs(group, 'c:pie3DChart')
      ? { rotX: viewAngle('c:rotX', 0), rotY: viewAngle('c:rotY', 0) }
      : undefined;
  const plotBox = plotBoxOf(plotArea);
  const firstSliceRaw =
    group && (poIs(group, 'c:pieChart') || poIs(group, 'c:doughnutChart'))
      ? Number(poVal(poChildren(group).find((c) => poIs(c, 'c:firstSliceAng'))))
      : Number.NaN;
  const firstSliceAngle = Number.isFinite(firstSliceRaw) ? firstSliceRaw : undefined;
  const showValues = group ? chartShowsValues(group) : false;
  const firstSerNode = group ? poChildren(group).find((c) => poIs(c, 'c:ser')) : undefined;
  const dataLabels = dataLabelsOf(group, firstSerNode);
  const catAxNode = poChildren(plotArea).find((c) => poIs(c, 'c:catAx'));
  const valAxNode = poChildren(plotArea).find((c) => poIs(c, 'c:valAx'));
  // §21.2.2.28 `c:axPos` — an axis line and its gridlines are geometry, so bind
  // them by WHERE the axis sits and not by which element declared it. A scatter
  // has two `c:valAx` and no `c:catAx` at all: chartTitle_noTitle.xlsx asks for
  // a 0.75pt #BFBFBF rule along the bottom and got the 1pt #595959 we fall back
  // to, and its left axis took the BOTTOM axis's styling.
  const axisAt = (...where: ReadonlyArray<string>): PoNode | undefined =>
    poChildren(plotArea).find(
      (c) =>
        (poIs(c, 'c:catAx') || poIs(c, 'c:valAx') || poIs(c, 'c:dateAx')) &&
        c !== secondaryValAx &&
        where.includes(poVal(poChildren(c).find((k) => poIs(k, 'c:axPos'))) ?? ''),
    );
  const bottomAxNode = axisAt('b', 't') ?? catAxNode;
  const leftAxNode = axisAt('l', 'r') ?? valAxNode;
  const catAxisLine = lineStyleOf(bottomAxNode, resolveColor);
  const valAxisLine = lineStyleOf(leftAxNode, resolveColor);
  const secondaryValAxisLine = lineStyleOf(secondaryValAx, resolveColor);
  const gridLine = lineStyleOf(
    leftAxNode ? poChildren(leftAxNode).find((c) => poIs(c, 'c:majorGridlines')) : undefined,
    resolveColor,
  );
  // §21.2.2.226 — a scatter has two value axes and no category axis, and which
  // is which is where each SITS: its upright one is the values', its lying one
  // the x's. Taken as "the first c:valAx", Excel's own order put the x axis's
  // ends, step and format on the y axis (DataTableCities.xlsx runs its longitude
  // −180…180 by 60 in `0"°"`), and its title on the wrong side.
  const isScatter = type === 'scatter';
  const catAxisTitle = isScatter ? axisTitleOf(bottomAxNode) : axisTitle(plotArea, 'c:catAx');
  const valAxisTitle = isScatter ? axisTitleOf(leftAxNode) : axisTitle(plotArea, 'c:valAx');
  const valAxisMin = isScatter
    ? axisScalingOf(leftAxNode, 'c:min')
    : axisScaling(plotArea, 'c:min');
  const valAxisMax = isScatter
    ? axisScalingOf(leftAxNode, 'c:max')
    : axisScaling(plotArea, 'c:max');
  const valAxisMajorUnit = isScatter ? majorUnitOfAxis(leftAxNode) : majorUnitOf(plotArea);
  const xAxisMin = isScatter ? axisScalingOf(bottomAxNode, 'c:min') : undefined;
  const xAxisMax = isScatter ? axisScalingOf(bottomAxNode, 'c:max') : undefined;
  const xAxisMajorUnit = isScatter ? majorUnitOfAxis(bottomAxNode) : undefined;
  const xNumberFormat = isScatter ? formatCodeOf(bottomAxNode) : undefined;
  // §21.2.2.134 — the category axis may run the other way, which is how a
  // ranked bar chart puts its first row at the top.
  const catAxisReversed = axisOrientation(catAxNode) === 'maxMin';
  // §21.2.2.33/§21.2.2.34/§21.2.2.207 — where the category axis crosses the
  // value axis, and where its labels stand.
  const catAxisCrosses = axisCrossing(catAxNode);
  const valAxisCrosses = axisCrossing(valAxNode);
  const tickLblPos = catAxNode
    ? poVal(poChildren(catAxNode).find((c) => poIs(c, 'c:tickLblPos')))
    : undefined;
  const catTickLabelPos =
    tickLblPos === 'low' || tickLblPos === 'high' || tickLblPos === 'none' ? tickLblPos : undefined;
  // §21.2.2.198 — the chart-space frame sits beside <c:chart>, not inside it.
  const chartSpace = tree.find((c) => poIs(c, 'c:chartSpace'));
  const spaceSpPr = poChildren(chartSpace).find((c) => poIs(c, 'c:spPr'));
  const frameLine = spaceSpPr ? poChildren(spaceSpPr).find((c) => poIs(c, 'a:ln')) : undefined;
  // A chart that states no frame at all is not a chart without one: both
  // references draw plain charts on white inside a light grey rule (chart-prop
  // .docx, chart-size.docx). A chart that DOES state one is taken at its word,
  // `<a:noFill/>` included — chart-dupe.docx asks for neither and gets neither.
  const frameStyle = lineStyleOf(chartSpace, resolveColor);
  const frameFillHex = spaceSpPr ? frameFillOf(spaceSpPr, resolveColor) : 'FFFFFF';
  // §20.1.8.14 — …or a PICTURE, which is what chart-texture-bg.pptx papers its
  // whole chart with. The blip is named through the CHART part's own
  // relationships, so the resolver is the caller's.
  const frameFillImage = spaceSpPr ? blipFillOf(spaceSpPr, resolveImage) : undefined;
  // §21.2.2.145 — the plot rectangle carries its own fill and rule.
  const plotSpPr = poChildren(plotArea).find((c) => poIs(c, 'c:spPr'));
  const plotFillHex = plotSpPr ? frameFillOf(plotSpPr, resolveColor) : undefined;
  const plotLine = lineStyleOf(plotArea, resolveColor);
  const frameLineHex = spaceSpPr
    ? frameLine
      ? frameFillOf(frameLine, resolveColor)
      : undefined
    : 'D9D9D9';
  const numberFormat = isScatter ? formatCodeOf(leftAxNode) : valueFormatCode(plotArea);

  const legend = poChildren(chart).find((c) => poIs(c, 'c:legend'));
  const legendPos = legend
    ? poVal(poChildren(legend).find((c) => poIs(c, 'c:legendPos')))
    : undefined;

  // §21.2.2.216 — the text each part of the chart is set in, when the host
  // asked for it. The labels belong to the axis that DRAWS them: a scatter's
  // horizontal value axis labels as the category axis does elsewhere.
  const firstSer = serNodes[0];
  const text = textDefaults
    ? chartTextStyles(
        {
          ...(chartSpace ? { chartSpace } : {}),
          ...withNode(
            'title',
            poChildren(chart).find((c) => poIs(c, 'c:title')),
          ),
          ...withNode('legend', legend),
          ...withNode(
            'catAxis',
            isScatter
              ? bottomAxNode
              : (catAxNode ?? poChildren(plotArea).find((c) => poIs(c, 'c:dateAx'))),
          ),
          ...withNode(
            'valAxis',
            isScatter
              ? leftAxNode
              : poChildren(plotArea).find((c) => poIs(c, 'c:valAx') && c !== secondaryValAx),
          ),
          ...withNode('secondaryValAxis', secondaryValAx),
          ...withNode(
            'dataLabels',
            (firstSer ? poChildren(firstSer).find((c) => poIs(c, 'c:dLbls')) : undefined) ??
              (group ? poChildren(group).find((c) => poIs(c, 'c:dLbls')) : undefined),
          ),
        },
        resolveColor,
        (fill) => colorFromSolidFill(fill, resolveColor),
        textDefaults,
      )
    : undefined;

  const title = chartTitle(chart, series);
  // §21.2.2.75 — the gap between category slots, as a percentage of the bar
  // width. Unread, every bar took 0.63 of its slot; 57362.xlsx asks for 219 and
  // its bars came out 2.6× too wide. The schema default is 150.
  // …and a chart that does not write one gets that default, which is NOT the
  // same as a gap of zero: `Number(undefined ?? '')` is 0, a finite number that
  // passed the guard below, so every chart silent about its gap drew its bars
  // touching (60255_extra_drawingparts.xlsx fills its whole slot where both
  // references leave half of it).
  const gapNode = poChildren(plotArea)
    .flatMap((g) => poChildren(g))
    .find((c) => poIs(c, 'c:gapWidth'));
  const gapRaw = gapNode ? poVal(gapNode) : undefined;
  const gapPercent = gapRaw === undefined ? NaN : Number(gapRaw);

  return {
    type,
    ...(title ? { title } : {}),
    ...(text ? { text } : {}),
    ...(Number.isFinite(gapPercent) && gapPercent >= 0 ? { gapPercent } : {}),
    categories,
    ...(categoriesRef ? { categoriesRef } : {}),
    ...(categoryGroups && categoryGroups.length > 0 ? { categoryGroups } : {}),
    series,
    hasLegend: legend !== undefined,
    ...(isLegendPos(legendPos) ? { legendPos } : {}),
    ...(barDir === 'col' || barDir === 'bar' ? { barDir } : {}),
    ...(isGrouping(grouping) ? { grouping } : {}),
    ...(doughnut ? { doughnut: true } : {}),
    ...(pie3D ? { pie3D } : {}),
    ...(plotBox ? { plotBox } : {}),
    ...(firstSliceAngle ? { firstSliceAngle } : {}),
    ...(showValues ? { showValues: true } : {}),
    ...(dataLabels ? { dataLabels } : {}),
    ...(catAxisTitle ? { catAxisTitle } : {}),
    ...(catAxisReversed ? { catAxisReversed } : {}),
    ...(catAxisCrosses !== undefined ? { catAxisCrosses } : {}),
    ...(valAxisCrosses !== undefined ? { valAxisCrosses } : {}),
    ...(catTickLabelPos ? { catTickLabelPos } : {}),
    ...(valAxisTitle ? { valAxisTitle } : {}),
    ...(secondaryValAxisTitle ? { secondaryValAxisTitle } : {}),
    ...(scatterStyle ? { scatterStyle } : {}),
    ...(lineMarkers ? { lineMarkers: true } : {}),
    ...(catAxisLine ? { catAxisLine } : {}),
    ...(valAxisLine ? { valAxisLine } : {}),
    ...(secondaryValAxisLine ? { secondaryValAxisLine } : {}),
    ...(gridLine ? { gridLine } : {}),
    ...(valAxisMin !== undefined ? { valAxisMin } : {}),
    ...(valAxisMax !== undefined ? { valAxisMax } : {}),
    ...(valAxisMajorUnit !== undefined ? { valAxisMajorUnit } : {}),
    ...(xAxisMin !== undefined ? { xAxisMin } : {}),
    ...(xAxisMax !== undefined ? { xAxisMax } : {}),
    ...(xAxisMajorUnit !== undefined ? { xAxisMajorUnit } : {}),
    ...(xNumberFormat ? { xNumberFormat } : {}),
    ...(frameFillHex ? { frameFillHex } : {}),
    ...(frameFillImage ? { frameFillImage } : {}),
    ...(frameLineHex ? { frameLineHex } : {}),
    ...(frameStyle?.widthPt !== undefined ? { frameLineWidthPt: frameStyle.widthPt } : {}),
    ...(frameStyle?.dash ? { frameLineDash: frameStyle.dash } : {}),
    ...(plotFillHex ? { plotFillHex } : {}),
    ...(plotLine && plotLine.none !== true ? { plotLine } : {}),
    ...(numberFormat ? { numberFormat } : {}),
  };
}

/** `{ [role]: node }` when there is a node — the shape `ChartTextNodes` takes. */
function withNode<TRole extends string>(
  role: TRole,
  node: PoNode | undefined,
): { [P in TRole]?: PoNode } {
  return (node ? { [role]: node } : {}) as { [P in TRole]?: PoNode };
}

function parseSeries(ser: PoNode, resolveColor: ColorResolver, most: number): ChartSeries {
  // Category charts carry values in c:val; scatter carries them in c:yVal with
  // the independent variable in c:xVal.
  const valNode =
    poChildren(ser).find((c) => poIs(c, 'c:val')) ?? poChildren(ser).find((c) => poIs(c, 'c:yVal'));
  const values = valNode ? denseNumbers(valNode, most) : [];
  const xValNode = poChildren(ser).find((c) => poIs(c, 'c:xVal'));
  const xValues = xValNode ? denseNumbers(xValNode, most) : undefined;
  const name = seriesName(ser);
  const colorHex = fillColorOf(
    poChildren(ser).find((c) => poIs(c, 'c:spPr')),
    resolveColor,
  );
  const pointColors = dataPointColors(ser, resolveColor);
  const pointLabels = customDataLabels(ser);
  const explosion = explosionOf(ser);
  const pointExplosions = poChildren(ser)
    .filter((c) => poIs(c, 'c:dPt'))
    .flatMap((dPt) => {
      const percent = explosionOf(dPt);
      const idx = poIntAttr(
        poChildren(dPt).find((c) => poIs(c, 'c:idx')),
        'val',
      );
      return percent !== undefined && idx !== undefined ? [{ idx, percent }] : [];
    });
  const pointLabelPlacements = labelPlacements(ser);
  const marker = seriesMarker(ser);
  const line = lineStyleOf(ser, resolveColor);
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
    ...(explosion ? { explosion } : {}),
    ...(pointExplosions.length > 0 ? { pointExplosions } : {}),
    ...(pointLabelPlacements.length > 0 ? { pointLabelPlacements } : {}),
    ...(marker ? { marker } : {}),
    ...(line ? { line } : {}),
  };
}

const MARKER_SYMBOLS = new Set<string>([
  'circle',
  'dash',
  'diamond',
  'dot',
  'none',
  'plus',
  'square',
  'star',
  'triangle',
  'x',
]);

/**
 * §21.2.2.106 `c:marker` — the series' own point symbol. `auto` (and the
 * picture marker we cannot draw) read as absent, leaving the reader's default.
 */
function seriesMarker(ser: PoNode): ChartMarker | undefined {
  const node = poChildren(ser).find((c) => poIs(c, 'c:marker'));
  if (!node) return undefined;
  const symbol = poVal(poChildren(node).find((c) => poIs(c, 'c:symbol')));
  if (!symbol || !MARKER_SYMBOLS.has(symbol)) return undefined;
  // §21.2.2.153 — `c:size` is already in points (2–72).
  const sizePt = poIntAttr(
    poChildren(node).find((c) => poIs(c, 'c:size')),
    'val',
  );
  return {
    symbol: symbol as ChartMarkerSymbol,
    ...(sizePt !== undefined && sizePt > 0 ? { sizePt } : {}),
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
/**
 * §21.2.2.95 — the plot area's box where the author sized it: `edge` mode,
 * every one of x, y, w and h stated, each a fraction of the chart.
 */
function plotBoxOf(plotArea: PoNode): Chart['plotBox'] {
  const layout = poChildren(plotArea).find((c) => poIs(c, 'c:layout'));
  const manual = layout ? poChildren(layout).find((c) => poIs(c, 'c:manualLayout')) : undefined;
  if (!manual) return undefined;
  const kids = poChildren(manual);
  const val = (tag: string): string | undefined => poVal(kids.find((c) => poIs(c, tag)));
  if (val('c:xMode') !== 'edge' || val('c:yMode') !== 'edge') return undefined;
  const [x, y, w, h] = ['c:x', 'c:y', 'c:w', 'c:h'].map((tag) => Number(val(tag)));
  if (![x, y, w, h].every((v) => v !== undefined && Number.isFinite(v))) return undefined;
  if (w! <= 0 || h! <= 0) return undefined;
  return {
    x: x!,
    y: y!,
    w: w!,
    h: h!,
    ...(val('c:layoutTarget') === 'inner' ? { inner: true } : {}),
  };
}

/** §21.2.2.62 `c:explosion` — a percentage of the pie's radius, at most 400. */
function explosionOf(owner: PoNode): number | undefined {
  const v = Number(poVal(poChildren(owner).find((c) => poIs(c, 'c:explosion'))));
  return Number.isFinite(v) && v >= 0 ? Math.min(v, 400) : undefined;
}

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
/**
 * §20.1.8.14 `a:blipFill` — the picture a chart element is filled with, and
 * whether it TILES at its own size rather than stretching over the box.
 *
 * @param spPr         The element's `c:spPr`.
 * @param resolveImage The chart part's own blip resolver, when the caller has one.
 * @returns The stored picture, or `undefined` when the fill is not one.
 */
function blipFillOf(
  spPr: PoNode,
  resolveImage: ((relId: string) => ResourceId | undefined) | undefined,
): Chart['frameFillImage'] {
  const blipFill = poChildren(spPr).find((c) => poIs(c, 'a:blipFill'));
  const blip = blipFill ? poChildren(blipFill).find((c) => poIs(c, 'a:blip')) : undefined;
  const relId = blip ? poAttr(blip, 'embed') : undefined;
  const resource = relId !== undefined ? resolveImage?.(relId) : undefined;
  if (resource === undefined || !blipFill) return undefined;
  const tiled = poChildren(blipFill).some((c) => poIs(c, 'a:tile'));
  return { resource, ...(tiled ? { tiled } : {}) };
}

/**
 * The BACKGROUND fill of a frame (the chart space, the plot rectangle): its own
 * `a:solidFill` and nothing else. Unlike a series, a frame that states
 * `<a:noFill/>` is not filled in the colour of its own rule —
 * Chart_Plot_BorderLine_Style.docx rules its plot in orange and we painted the
 * whole plot orange.
 *
 * @param spPr         The frame's `c:spPr`, or `undefined`.
 * @param resolveColor Maps a DrawingML colour reference to 6-hex.
 * @returns The fill colour, or `undefined` when the frame states none.
 */
function frameFillOf(spPr: PoNode | undefined, resolveColor: ColorResolver): string | undefined {
  if (!spPr) return undefined;
  const solid = poChildren(spPr).find((c) => poIs(c, 'a:solidFill'));
  return solid ? colorFromSolidFill(solid, resolveColor) : undefined;
}

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

/** The colour elements a chart's fill names a colour by (§20.1.2.3). */
const CHART_COLOR_NODES = ['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr'] as const;

function colorFromSolidFill(solid: PoNode, resolveColor: ColorResolver): string | undefined {
  for (const c of poChildren(solid)) {
    // …a system or a preset colour as much as a scheme or an RGB one: Excel
    // 2007 and 2010 outline a chart in `<a:sysClr val="windowText">`, and
    // skipped, dataValidationTableRange.xlsx lost the black frame round its
    // chart and round its plot.
    if (!CHART_COLOR_NODES.some((name) => poIs(c, name))) continue;
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
  const authored = collectAT(title) || cachedTitleText(title);
  if (authored) return authored;
  const deleted = poVal(poChildren(chart).find((c) => poIs(c, 'c:autoTitleDeleted'))) === '1';
  if (deleted) return undefined;
  return (series.length === 1 ? series[0]?.name : undefined) ?? 'Chart Title';
}

/**
 * §21.2.2.215 — a title the author gave as a FORMULA: `c:tx/c:strRef` with the
 * text in its `c:strCache`, and not one `a:t` run anywhere. Read for rich text
 * alone, such a title came out as the generated placeholder —
 * chartTitle_withTitleFormula.xlsx printed "Chart Title" where both references
 * print "Formula Title from Excel 2016".
 *
 * @param title The `c:title` node.
 * @returns The cached text, or undefined when the title carries no cache.
 */
function cachedTitleText(title: PoNode): string | undefined {
  const tx = poChildren(title).find((c) => poIs(c, 'c:tx'));
  const ref = tx ? poChildren(tx).find((c) => poIs(c, 'c:strRef')) : undefined;
  if (!ref) return undefined;
  const cached = denseStrings(ref, MOST_CHART_POINTS).filter((t) => t.length > 0);
  return cached.length > 0 ? cached.join(' ') : undefined;
}

// c:catAx / c:valAx → c:title text.
//
// Like the chart's own title (§21.2.2.213), an axis title element with no
// `c:tx` asks the reader to generate one; Excel and LibreOffice both write the
// placeholder "Axis Title". 57362.xlsx leaves both of its value axes that way
// and we drew neither.
function axisTitle(plotArea: PoNode, axTag: string): string | undefined {
  return axisTitleOf(poChildren(plotArea).find((c) => poIs(c, axTag)));
}

/** An axis node's own `c:title` text. */
function axisTitleOf(ax: PoNode | undefined): string | undefined {
  const title = ax ? poChildren(ax).find((c) => poIs(c, 'c:title')) : undefined;
  if (!title) return undefined;
  return collectAT(title) || cachedTitleText(title) || 'Axis Title';
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
/** §21.2.2.134 `c:scaling/c:orientation` — `minMax` (the default) or `maxMin`. */
/**
 * §21.2.2.33/§21.2.2.34 — where an axis crosses the one it is drawn against:
 * `crossesAt` a value, or `crosses` at that axis's minimum or maximum.
 * undefined for `autoZero`, the default.
 */
function axisCrossing(ax: PoNode | undefined): 'min' | 'max' | number | undefined {
  if (!ax) return undefined;
  const at = poChildren(ax).find((c) => poIs(c, 'c:crossesAt'));
  const value = at ? Number(poAttr(at, 'val')) : Number.NaN;
  if (Number.isFinite(value)) return value;
  const crosses = poVal(poChildren(ax).find((c) => poIs(c, 'c:crosses')));
  return crosses === 'min' || crosses === 'max' ? crosses : undefined;
}

function axisOrientation(ax: PoNode | undefined): string | undefined {
  const scaling = ax ? poChildren(ax).find((c) => poIs(c, 'c:scaling')) : undefined;
  return scaling ? poVal(poChildren(scaling).find((c) => poIs(c, 'c:orientation'))) : undefined;
}

/** §21.2.2.98 `c:valAx/c:majorUnit` — a positive step, or undefined for "auto". */
function majorUnitOf(plotArea: PoNode): number | undefined {
  return majorUnitOfAxis(poChildren(plotArea).find((c) => poIs(c, 'c:valAx')));
}

function majorUnitOfAxis(ax: PoNode | undefined): number | undefined {
  const node = ax ? poChildren(ax).find((c) => poIs(c, 'c:majorUnit')) : undefined;
  const v = node ? Number(poAttr(node, 'val')) : Number.NaN;
  return Number.isFinite(v) && v > 0 ? v : undefined;
}

function axisScaling(plotArea: PoNode, tag: 'c:min' | 'c:max'): number | undefined {
  return axisScalingOf(
    poChildren(plotArea).find((c) => poIs(c, 'c:valAx')),
    tag,
  );
}

function axisScalingOf(ax: PoNode | undefined, tag: 'c:min' | 'c:max'): number | undefined {
  const scaling = ax ? poChildren(ax).find((c) => poIs(c, 'c:scaling')) : undefined;
  const node = scaling ? poChildren(scaling).find((c) => poIs(c, tag)) : undefined;
  const v = node ? Number(poAttr(node, 'val')) : Number.NaN;
  return Number.isFinite(v) ? v : undefined;
}

/**
 * §21.2.2.196 `c:spPr/a:ln` — an axis's (or a series') own rule.
 * `<a:ln><a:noFill/>` means it draws nothing, which is not the same as having
 * no `c:spPr` at all: the first hides the line, the second leaves it to the
 * renderer's default.
 *
 * @param owner        The `c:catAx` / `c:valAx` / `c:majorGridlines` / `c:ser`
 *                     node.
 * @param resolveColor Maps a DrawingML colour reference to 6 hex digits.
 * @returns The rule, or undefined when the node says nothing about it.
 */
const DASHES = new Set<string>([
  'solid',
  'dot',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
  'sysDash',
  'sysDot',
]);

function lineStyleOf(
  owner: PoNode | undefined,
  resolveColor: ColorResolver,
): ChartLineStyle | undefined {
  const spPr = owner ? poChildren(owner).find((c) => poIs(c, 'c:spPr')) : undefined;
  const ln = spPr ? poChildren(spPr).find((c) => poIs(c, 'a:ln')) : undefined;
  if (!ln) return undefined;
  if (poChildren(ln).some((c) => poIs(c, 'a:noFill'))) return { none: true };
  const solid = poChildren(ln).find((c) => poIs(c, 'a:solidFill'));
  const colorHex = solid ? colorFromSolidFill(solid, resolveColor) : undefined;
  // §20.1.2.1 `w` is in EMU; 12 700 to the point.
  const emu = Number(poAttr(ln, 'w'));
  const widthPt = Number.isFinite(emu) && emu > 0 ? emu / 12700 : undefined;
  // §20.1.10.49 — a preset dash on the rule; Chart_BorderLine_Style.docx
  // outlines each of its bars in a different one.
  const prst = poChildren(ln).find((c) => poIs(c, 'a:prstDash'));
  const dashVal = prst ? poAttr(prst, 'val') : undefined;
  const dash = dashVal !== undefined && DASHES.has(dashVal) ? (dashVal as ShapeDash) : undefined;
  if (!colorHex && widthPt === undefined && !dash) return undefined;
  return {
    ...(colorHex ? { colorHex } : {}),
    ...(widthPt !== undefined ? { widthPt } : {}),
    ...(dash ? { dash } : {}),
  };
}

function valueFormatCode(plotArea: PoNode): string | undefined {
  return formatCodeOf(poChildren(plotArea).find((c) => poIs(c, 'c:valAx')));
}

function formatCodeOf(ax: PoNode | undefined): string | undefined {
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

/**
 * §21.2.2.49 — what a chart's data labels show: its group's `c:dLbls` with the
 * first series' own over it. `c:delete` shows nothing; a chart with no
 * `c:dLbls` anywhere has no labels, which is not a chart that shows its
 * percentages — Excel prints nothing on such a pie.
 *
 * @param group     The chart group (`c:pieChart`, …).
 * @param firstSer  Its first series.
 * @returns The switches, or undefined for a chart without data labels.
 */
function dataLabelsOf(
  group: PoNode | undefined,
  firstSer: PoNode | undefined,
): ChartDataLabels | undefined {
  const own = (owner: PoNode | undefined): PoNode | undefined =>
    owner ? poChildren(owner).find((c) => poIs(c, 'c:dLbls')) : undefined;
  const groupLabels = own(group);
  const seriesLabels = own(firstSer);
  if (!groupLabels && !seriesLabels) return undefined;
  const read = (dLbls: PoNode | undefined): ChartDataLabels => {
    if (!dLbls) return {};
    const kids = poChildren(dLbls);
    const deleted = poVal(kids.find((c) => poIs(c, 'c:delete')));
    if (deleted === '1' || deleted === 'true') {
      return { showVal: false, showCatName: false, showSerName: false, showPercent: false };
    }
    const flag = (tag: string): Partial<Record<string, boolean>> => {
      const v = poVal(kids.find((c) => poIs(c, tag)));
      return v === undefined ? {} : { [tag.slice(2)]: v === '1' || v === 'true' };
    };
    const separatorNode = kids.find((c) => poIs(c, 'c:separator'));
    const numFmt = kids.find((c) => poIs(c, 'c:numFmt'));
    const code = numFmt ? poAttr(numFmt, 'formatCode') : undefined;
    const linked = numFmt ? poAttr(numFmt, 'sourceLinked') : undefined;
    const position = labelPosition(kids);
    return {
      ...flag('c:showVal'),
      ...flag('c:showCatName'),
      ...flag('c:showSerName'),
      ...flag('c:showPercent'),
      ...flag('c:showLeaderLines'),
      ...(position ? { position } : {}),
      ...(separatorNode ? { separator: poText(separatorNode) } : {}),
      ...(code && code.trim() !== '' && code.trim().toLowerCase() !== 'general' && linked !== '1'
        ? { numberFormat: code }
        : {}),
    };
  };
  return { ...read(groupLabels), ...read(seriesLabels) };
}

const LABEL_POSITIONS: ReadonlySet<string> = new Set([
  'bestFit',
  'b',
  'ctr',
  'inBase',
  'inEnd',
  'l',
  'outEnd',
  'r',
  't',
]);

/** §21.2.2.48 — a label's `c:dLblPos` among its siblings, when it states one. */
function labelPosition(kids: ReadonlyArray<PoNode>): ChartLabelPosition | undefined {
  const v = poVal(kids.find((c) => poIs(c, 'c:dLblPos')));
  return v !== undefined && LABEL_POSITIONS.has(v) ? (v as ChartLabelPosition) : undefined;
}

/**
 * §21.2.2.47 — where each point's own label stands, for the points that say:
 * their `c:dLblPos`, and the `c:manualLayout` an author dragged them to
 * (45544.xlsx sets every one of its slices' names out by hand).
 */
function labelPlacements(ser: PoNode): NonNullable<ChartSeries['pointLabelPlacements']> {
  const dLbls = poChildren(ser).find((c) => poIs(c, 'c:dLbls'));
  if (!dLbls) return [];
  const out: Array<NonNullable<ChartSeries['pointLabelPlacements']>[number]> = [];
  for (const dLbl of poChildren(dLbls)) {
    if (!poIs(dLbl, 'c:dLbl')) continue;
    const kids = poChildren(dLbl);
    const idx =
      poIntAttr(
        kids.find((c) => poIs(c, 'c:idx')),
        'val',
      ) ?? 0;
    const position = labelPosition(kids);
    const layout = kids.find((c) => poIs(c, 'c:layout'));
    const manual = layout ? poChildren(layout).find((c) => poIs(c, 'c:manualLayout')) : undefined;
    const num = (tag: string): number | undefined => {
      const v = Number(poVal(manual ? poChildren(manual).find((c) => poIs(c, tag)) : undefined));
      return Number.isFinite(v) ? v : undefined;
    };
    const x = num('c:x');
    const y = num('c:y');
    const mode = (tag: string): string | undefined =>
      poVal(manual ? poChildren(manual).find((c) => poIs(c, tag)) : undefined);
    const edge = mode('c:xMode') === 'edge' && mode('c:yMode') === 'edge';
    if (position === undefined && x === undefined && y === undefined) continue;
    out.push({
      idx,
      ...(position ? { position } : {}),
      ...(x !== undefined ? { x } : {}),
      ...(y !== undefined ? { y } : {}),
      ...(edge ? { edge: true } : {}),
    });
  }
  return out;
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

/** Whether `n` can index a point: a whole number from zero. */
const isPointIndex = (n: number): boolean => Number.isInteger(n) && n >= 0;

/**
 * How many slots `count` points and those at `indexes` take — the count, or
 * past it the last of them, and never more than `most`. A count or an index
 * that is no whole number (2.5) asked for an array of no possible length and
 * threw; it counts for nothing.
 */
function denseLength(count: number, indexes: ReadonlyArray<number>, most: number): number {
  let length = isPointIndex(count) ? count : 0;
  for (const idx of indexes) if (isPointIndex(idx)) length = Math.max(length, idx + 1);
  return Math.min(length, most);
}

/** A cache's values by point index, gaps 0, at most `most` of them. */
function denseNumbers(container: PoNode, most: number): Array<number> {
  const pts = readPts(container);
  const length = denseLength(
    ptCountOf(container),
    pts.map((p) => p.idx),
    most,
  );
  const arr = new Array<number>(length).fill(0);
  for (const p of pts) {
    const n = Number(p.v);
    if (Number.isFinite(n) && isPointIndex(p.idx) && p.idx < length) arr[p.idx] = n;
  }
  return arr;
}

/** A cache's texts by point index, gaps empty, at most `most` of them. */
function denseStrings(container: PoNode, most: number): Array<string> {
  const pts = readPts(container);
  const length = denseLength(
    ptCountOf(container),
    pts.map((p) => p.idx),
    most,
  );
  const arr = new Array<string>(length).fill('');
  for (const p of pts) if (isPointIndex(p.idx) && p.idx < length) arr[p.idx] = p.v;
  return arr;
}

/**
 * §21.2.2.115 `c:multiLvlStrCache` — categories labelled on several levels:
 * the first `c:lvl` labels each category, every later one groups them, a
 * group's label standing at the category it starts at. Read as one flat cache
 * it read as none, and an xlsx chart fell back to the cells its reference
 * names, both columns of them in turn: WithChartSheet.xlsx's six bars stood in
 * the first six of sixteen slots under a jumble of years and measure names.
 *
 * @param cat  The `c:cat` element.
 * @param most The most categories it keeps (see {@link pointsPerSeries}).
 * @returns The innermost labels and the outer levels' groups, or undefined
 *   for a category axis of one level.
 */
function multiLevelCategories(
  cat: PoNode,
  most: number,
): { categories: Array<string>; groups: NonNullable<Chart['categoryGroups']> } | undefined {
  const cache = poFindDescendant(cat, 'c:multiLvlStrCache');
  if (!cache) return undefined;
  const count = poIntAttr(poChildren(cache).find((c) => poIs(c, 'c:ptCount')) ?? cache, 'val') ?? 0;
  const levels = poChildren(cache)
    .filter((c) => poIs(c, 'c:lvl'))
    .map((lvl) =>
      poChildren(lvl)
        .filter((pt) => poIs(pt, 'c:pt'))
        .map((pt) => {
          const v = poChildren(pt).find((c) => poIs(c, 'c:v'));
          return { start: poIntAttr(pt, 'idx') ?? 0, label: v ? poText(v) : '' };
        }),
    );
  const [inner, ...outer] = levels;
  if (!inner) return undefined;
  const length = denseLength(
    count,
    inner.map((p) => p.start),
    most,
  );
  const categories = new Array<string>(length).fill('');
  // A label past the axis's last slot has none to stand in, and a group
  // starting past it groups nothing.
  const kept = (p: { start: number }): boolean => isPointIndex(p.start) && p.start < length;
  for (const p of inner) if (kept(p)) categories[p.start] = p.label;
  return {
    categories,
    groups: outer.map((level) => level.filter(kept).sort((a, b) => a.start - b.start)),
  };
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
  // No `colorsN.xml` — an Office 2011 extension a chart from 2007 does not
  // carry. The cycle is then the THEME's own accents (§20.1.4.1.5), which is
  // what both references paint: WithThreeCharts.xlsx keeps the Office 2007
  // scheme, so its second series is C0504D red and its pie runs blue, red,
  // green, purple, cyan, orange — where our built-in 2013 accents drew orange
  // and blue, orange, grey, yellow, light blue, green.
  const themed = ACCENT_SLOTS.map((scheme) => resolveColor({ scheme }));
  return themed.every((c): c is string => c !== undefined)
    ? { ...chart, seriesColorCycle: themed }
    : chart;
}

/** §20.1.4.1.5 — the accent slots a chart cycles its series through. */
const ACCENT_SLOTS: ReadonlyArray<string> = [
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
];

const SCATTER_STYLES = new Set(['none', 'line', 'lineMarker', 'marker', 'smooth', 'smoothMarker']);

const REL_CHART_COLOR_STYLE =
  'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
