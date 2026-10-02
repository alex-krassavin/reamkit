// How a chart sets its text — face, size, weight and colour — role by role.
//
// §21.2.2.216 `c:txPr` and §21.2.2.156 `c:rich` carry DrawingML run properties
// (§21.1.2.3.2 `a:defRPr`, §21.1.2.3.9 `a:rPr`) for one element of a chart; the
// chart's own `c:chartSpace/c:txPr` stands under all of them, and under that the
// application decides. Excel decides as its own PDF of probe workbooks shows
// (2026-10-02): every text in the theme's minor font at 10pt, in the theme's
// text colour, the chart title and the axis titles bold — not in the workbook's
// Normal font, which a chart never looks at. A workbook without a theme part
// has Calibri when its Normal font is Calibri and Aptos Narrow, the Office 2023
// theme's minor font, when it is anything else.

import type { PoNode } from '@/core/po-helpers';
import type { ChartTextStyle, ChartTextStyles } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { ThemeFonts } from '@/core/drawingml/theme-parser';
import { resolveThemeFont } from '@/core/drawingml/theme-parser';
import { poAttr, poChildren, poIs } from '@/core/po-helpers';

/**
 * What a host application sets the chart text a chart leaves unsaid in.
 * Passing it asks for the text to be resolved at all; without it a chart
 * carries no {@link ChartTextStyles} and the renderer keeps its own defaults.
 */
export interface ChartTextDefaults {
  /** The theme's font scheme — what `+mn-lt` and `+mj-lt` stand for. */
  readonly themeFonts?: ThemeFonts;
  /** The family of text no run names. */
  readonly family?: string;
}

/** Excel's size for chart text no run sizes. */
export const EXCEL_CHART_TEXT_PT = 10;

/** The nodes each role's own text properties hang from. */
export interface ChartTextNodes {
  readonly chartSpace?: PoNode;
  readonly title?: PoNode;
  readonly legend?: PoNode;
  readonly catAxis?: PoNode;
  readonly valAxis?: PoNode;
  readonly secondaryValAxis?: PoNode;
  /** `c:dLbls` — the first series' own, else its group's. */
  readonly dataLabels?: PoNode;
}

/** Reads an `a:solidFill` as a chart reads it, to RRGGBB. */
export type ChartFillReader = (solidFill: PoNode) => string | undefined;

/**
 * Resolve a chart's text, role by role: the element's own run properties over
 * the chart's over the host's defaults.
 *
 * @param nodes        Where each role's properties hang.
 * @param resolveColor Maps a colour reference to RRGGBB — the default text colour.
 * @param readFill     Reads a run's `a:solidFill`.
 * @param defaults     The host's defaults and theme fonts.
 * @returns Every role, resolved.
 */
export function chartTextStyles(
  nodes: ChartTextNodes,
  resolveColor: ColorResolver,
  readFill: ChartFillReader,
  defaults: ChartTextDefaults,
): ChartTextStyles {
  const ctx: ReadContext = { readFill, fonts: defaults.themeFonts };
  const base: ChartTextStyle = {
    ...(defaults.family ? { family: defaults.family } : {}),
    sizePt: EXCEL_CHART_TEXT_PT,
    colorHex: resolveColor({ scheme: 'tx1' }) ?? '000000',
  };
  const space = {
    ...base,
    ...bodyStyle(child(nodes.chartSpace, 'c:txPr'), ctx),
  };
  const titled = { ...space, bold: space.bold ?? true };
  const ofTitle = (title: PoNode | undefined): ChartTextStyle => ({
    ...titled,
    ...titleStyle(title, ctx),
  });
  const ofText = (node: PoNode | undefined): ChartTextStyle => ({
    ...space,
    ...bodyStyle(child(node, 'c:txPr'), ctx),
  });
  return {
    title: ofTitle(nodes.title),
    legend: ofText(nodes.legend),
    catAxis: ofText(nodes.catAxis),
    valAxis: ofText(nodes.valAxis),
    secondaryValAxis: ofText(nodes.secondaryValAxis ?? nodes.valAxis),
    catAxisTitle: ofTitle(child(nodes.catAxis, 'c:title')),
    valAxisTitle: ofTitle(child(nodes.valAxis, 'c:title')),
    secondaryValAxisTitle: ofTitle(child(nodes.secondaryValAxis, 'c:title')),
    dataLabels: ofText(nodes.dataLabels),
  };
}

/**
 * A title's text: its `c:txPr`, then the rich text's paragraph defaults, then
 * the first run's own properties — Excel keeps one style for a whole title.
 */
function titleStyle(title: PoNode | undefined, ctx: ReadContext): ChartTextStyle {
  if (!title) return {};
  const rich = child(child(title, 'c:tx'), 'c:rich');
  const firstRun = rich
    ? poChildren(rich)
        .filter((c) => poIs(c, 'a:p'))
        .flatMap((p) => poChildren(p))
        .find((c) => poIs(c, 'a:r'))
    : undefined;
  return {
    ...bodyStyle(child(title, 'c:txPr'), ctx),
    ...bodyStyle(rich, ctx),
    ...runStyle(child(firstRun, 'a:rPr'), ctx),
  };
}

/** A text body's paragraph defaults — the first paragraph's `a:pPr/a:defRPr`. */
function bodyStyle(body: PoNode | undefined, ctx: ReadContext): ChartTextStyle {
  const p = body ? poChildren(body).find((c) => poIs(c, 'a:p')) : undefined;
  return runStyle(child(child(p, 'a:pPr'), 'a:defRPr'), ctx);
}

/** §21.1.2.3.2 — one run-property node, read for what it states. */
function runStyle(node: PoNode | undefined, ctx: ReadContext): ChartTextStyle {
  if (!node) return {};
  const size = Number(poAttr(node, 'sz') ?? NaN);
  const bold = flag(poAttr(node, 'b'));
  const italic = flag(poAttr(node, 'i'));
  const typeface = poAttr(child(node, 'a:latin'), 'typeface')?.trim();
  const family = typeface ? resolveThemeFont(typeface, ctx.fonts) : undefined;
  const fill = child(node, 'a:solidFill');
  const colorHex = fill ? ctx.readFill(fill) : undefined;
  return {
    ...(family ? { family } : {}),
    // ST_TextFontSize — hundredths of a point, 100…400000.
    ...(Number.isFinite(size) && size >= 100 ? { sizePt: size / 100 } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(colorHex ? { colorHex } : {}),
  };
}

/** What reading one element's text properties needs. */
interface ReadContext {
  readonly readFill: ChartFillReader;
  readonly fonts: ThemeFonts | undefined;
}

/** xsd:boolean, as an attribute states it. */
function flag(value: string | undefined): boolean | undefined {
  if (value === '1' || value === 'true') return true;
  if (value === '0' || value === 'false') return false;
  return undefined;
}

function child(node: PoNode | undefined, tag: string): PoNode | undefined {
  return node ? poChildren(node).find((c) => poIs(c, tag)) : undefined;
}
