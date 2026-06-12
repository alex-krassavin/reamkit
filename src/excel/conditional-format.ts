// Conditional formatting evaluation (E-SHEET SC1/SC1b/SC1c) — the first Excel
// feature the flat table projection could not express. A sheet's
// <conditionalFormatting> rules plus the workbook's <dxfs> become a per-cell
// lookup; the print model overrides each cell's base format. Handles `cellIs`
// (compare to a constant → dxf), `colorScale` (interpolate a fill across the
// range's value extent), and `dataBar` (an in-cell bar whose length encodes the
// value); iconSet follows. A cell's text format (fill/font) is claimed by the
// first applicable cellIs/colorScale; a dataBar applies independently on top.

import type { CellIcon, CellIconShape } from '@/core/document-model';
import type {
  CfOperator,
  CfRuleColorScale,
  CfRuleDataBar,
  CfRuleIconSet,
  Cfvo,
  ConditionalFormat,
  Dxf,
  MergedRange,
  WorksheetCell,
  XlsxStyles,
} from '@/core/spreadsheet-model';

type Rgb = readonly [number, number, number];

// A colorScale resolved against its range's value extent: N monotonic
// thresholds paired with N pre-parsed RGB stop colours (N = 2 or 3).
interface ResolvedColorScale {
  readonly thresholds: ReadonlyArray<number>;
  readonly colors: ReadonlyArray<Rgb>;
}

// A dataBar resolved against its range's extent: the lower/upper value the bar
// scale spans, its fill colour, and the min/max bar length (0..1 of cell width).
interface ResolvedDataBar {
  readonly lower: number;
  readonly upper: number;
  readonly colorHex: string;
  readonly minLen: number;
  readonly maxLen: number;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

// A resolved per-cell override: a solid highlight fill, font tweaks, an in-cell
// data bar (fraction of the cell width 0..1 + colour), and/or a leading icon.
export interface CfOverride {
  readonly fillHex?: string;
  readonly fontColorHex?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly dataBar?: { readonly fraction: number; readonly colorHex: string };
  readonly icon?: CellIcon;
}

export type CellConditionalFormatter = (
  row: number,
  col: number,
  numericValue: number | undefined,
) => CfOverride | undefined;

interface FlatRule {
  readonly ranges: ReadonlyArray<MergedRange>;
  readonly rule: ConditionalFormat['rules'][number];
  // Present (and possibly undefined, when resolution failed) for colorScale /
  // dataBar / iconSet rules: the extent-derived data computed once from the
  // range values (gradient stops, bar bounds, or icon bucket thresholds).
  readonly scale?: ResolvedColorScale | undefined;
  readonly bar?: ResolvedDataBar | undefined;
  readonly iconThresholds?: ReadonlyArray<number> | undefined;
}

// Returns undefined when the sheet has no conditional formats (the common case)
// so callers skip the work entirely and stay byte-identical. `cells` supplies
// the range values colorScale rules need (min/max/percentile); cellIs ignores it.
export function buildConditionalFormatter(
  conditionalFormats: ReadonlyArray<ConditionalFormat> | undefined,
  styles: XlsxStyles,
  cells: ReadonlyArray<WorksheetCell>,
): CellConditionalFormatter | undefined {
  if (!conditionalFormats || conditionalFormats.length === 0) return undefined;
  const dxfs = styles.dxfs ?? [];
  const flat: Array<FlatRule> = [];
  for (const cf of conditionalFormats) {
    for (const rule of cf.rules) {
      if (rule.type === 'colorScale') {
        flat.push({ ranges: cf.ranges, rule, scale: resolveColorScale(rule, cf.ranges, cells) });
      } else if (rule.type === 'dataBar') {
        flat.push({ ranges: cf.ranges, rule, bar: resolveDataBar(rule, cf.ranges, cells) });
      } else if (rule.type === 'iconSet') {
        flat.push({
          ranges: cf.ranges,
          rule,
          iconThresholds: resolveIconThresholds(rule, cf.ranges, cells),
        });
      } else {
        flat.push({ ranges: cf.ranges, rule });
      }
    }
  }
  // §18.3.1.10 — a lower @priority wins. Evaluate in ascending priority. A cell's
  // text format (fill/font) is claimed by the first applicable cellIs/colorScale;
  // a dataBar fills its own slot independently, so a bar can sit over a fill.
  flat.sort((a, b) => a.rule.priority - b.rule.priority);

  return (row, col, value) => {
    if (value === undefined) return undefined;
    let textFmt: CfOverride | undefined;
    let textClaimed = false;
    let bar: CfOverride['dataBar'];
    let icon: CfOverride['icon'];
    for (const { ranges, rule, scale, bar: resolved, iconThresholds } of flat) {
      if (!coversCell(ranges, row, col)) continue;
      switch (rule.type) {
        case 'cellIs':
          if (!textClaimed && cellIsMatches(rule.operator, value, rule.formulas)) {
            textFmt = dxfToOverride(dxfs[rule.dxfId]);
            textClaimed = true;
          }
          break;
        case 'colorScale':
          if (!textClaimed && scale) {
            textFmt = { fillHex: rgbHex(colorScaleColor(scale, value)) };
            textClaimed = true;
          }
          break;
        case 'dataBar':
          if (!bar && resolved) {
            bar = { fraction: dataBarFraction(resolved, value), colorHex: resolved.colorHex };
          }
          break;
        case 'iconSet':
          if (!icon && iconThresholds) {
            const bucket = iconBucket(iconThresholds, value);
            icon = iconToCell(rule.iconSet, iconThresholds.length, bucket, rule.reverse ?? false);
          }
          break;
      }
    }
    if (!bar && !icon) return textFmt;
    return {
      ...(textFmt ?? {}),
      ...(bar ? { dataBar: bar } : {}),
      ...(icon ? { icon } : {}),
    };
  };
}

function coversCell(ranges: ReadonlyArray<MergedRange>, row: number, col: number): boolean {
  for (const r of ranges) {
    if (row >= r.startRow && row <= r.endRow && col >= r.startColumn && col <= r.endColumn) {
      return true;
    }
  }
  return false;
}

function cellIsMatches(
  operator: CfOperator,
  value: number,
  formulas: ReadonlyArray<string>,
): boolean {
  const a = Number(formulas[0]);
  if (!Number.isFinite(a)) return false;
  switch (operator) {
    case 'lessThan':
      return value < a;
    case 'lessThanOrEqual':
      return value <= a;
    case 'equal':
      return value === a;
    case 'notEqual':
      return value !== a;
    case 'greaterThanOrEqual':
      return value >= a;
    case 'greaterThan':
      return value > a;
    case 'between':
    case 'notBetween': {
      const b = Number(formulas[1]);
      if (!Number.isFinite(b)) return false;
      const inside = value >= Math.min(a, b) && value <= Math.max(a, b);
      return operator === 'between' ? inside : !inside;
    }
  }
}

function dxfToOverride(dxf: Dxf | undefined): CfOverride | undefined {
  if (!dxf) return undefined;
  // §18.8.20 — a dxf's solid fill conventionally carries its colour in bgColor.
  const fillHex = dxf.fill?.bgColorHex ?? dxf.fill?.fgColorHex;
  const out: Mutable<CfOverride> = {};
  if (fillHex) out.fillHex = fillHex;
  if (dxf.font?.colorHex) out.fontColorHex = dxf.font.colorHex;
  if (dxf.font?.bold !== undefined) out.bold = dxf.font.bold;
  if (dxf.font?.italic !== undefined) out.italic = dxf.font.italic;
  return Object.keys(out).length > 0 ? out : undefined;
}

// --- colorScale (E-SHEET SC1b) ---------------------------------------------

// Gather the numeric values within a colorScale's range(s). Iterates the sheet's
// actual (sparse) cells — bounded by real data — rather than the range box, so a
// whole-column scale (A:A) costs O(cells), not O(rows). Mirrors the cellIs value
// rule: only `n` cells holding a finite number contribute to the extent.
function collectRangeValues(
  cells: ReadonlyArray<WorksheetCell>,
  ranges: ReadonlyArray<MergedRange>,
): Array<number> {
  const out: Array<number> = [];
  for (const cell of cells) {
    if (cell.type !== 'n') continue;
    const n = Number(cell.rawValue);
    if (!Number.isFinite(n)) continue;
    if (coversCell(ranges, cell.row, cell.column)) out.push(n);
  }
  return out;
}

// Resolve a colorScale's cfvo stops to numeric thresholds (from the range's value
// distribution) and pre-parse its stop colours. Returns undefined — making the
// rule a no-op — when the range holds no numbers or any stop/colour can't resolve.
function resolveColorScale(
  rule: CfRuleColorScale,
  ranges: ReadonlyArray<MergedRange>,
  cells: ReadonlyArray<WorksheetCell>,
): ResolvedColorScale | undefined {
  const vals = collectRangeValues(cells, ranges);
  if (vals.length === 0) return undefined;
  const sorted = [...vals].sort((a, b) => a - b);
  const vmin = sorted[0]!;
  const vmax = sorted[sorted.length - 1]!;

  const thresholds: Array<number> = [];
  for (const cfvo of rule.cfvos) {
    const t = resolveCfvo(cfvo, vmin, vmax, sorted);
    if (t === undefined) return undefined;
    thresholds.push(t);
  }
  const colors: Array<Rgb> = [];
  for (const hex of rule.colorsHex) {
    const rgb = parseRgb(hex);
    if (!rgb) return undefined;
    colors.push(rgb);
  }
  if (thresholds.length < 2 || thresholds.length !== colors.length) return undefined;
  // Excel sorts the stops; clamp any out-of-order num/formula threshold up to its
  // predecessor so the gradient domain stays monotonic non-decreasing.
  for (let i = 1; i < thresholds.length; i++) {
    if (thresholds[i]! < thresholds[i - 1]!) thresholds[i] = thresholds[i - 1]!;
  }
  return { thresholds, colors };
}

// §18.3.1.11 — a cfvo stop's numeric threshold given the range's extent + sorted
// values. min/max take the extent; num/formula a literal; percent positions
// linearly in [min,max]; percentile interpolates the value distribution.
function resolveCfvo(
  cfvo: Cfvo,
  vmin: number,
  vmax: number,
  sorted: ReadonlyArray<number>,
): number | undefined {
  switch (cfvo.type) {
    case 'min':
    case 'autoMin':
      return vmin;
    case 'max':
    case 'autoMax':
      return vmax;
    case 'num':
    case 'formula': {
      const n = Number(cfvo.val);
      return Number.isFinite(n) ? n : undefined;
    }
    case 'percent': {
      const p = Number(cfvo.val);
      return Number.isFinite(p) ? vmin + (vmax - vmin) * (p / 100) : undefined;
    }
    case 'percentile': {
      const p = Number(cfvo.val);
      return Number.isFinite(p) ? percentile(sorted, p) : undefined;
    }
  }
}

// PERCENTILE.INC over an ascending array: linear interpolation between ranks.
function percentile(sorted: ReadonlyArray<number>, p: number): number {
  const n = sorted.length;
  if (n === 1) return sorted[0]!;
  const rank = (Math.min(100, Math.max(0, p)) / 100) * (n - 1);
  const lo = Math.floor(rank);
  const frac = rank - lo;
  const a = sorted[lo]!;
  const b = sorted[Math.min(lo + 1, n - 1)]!;
  return a + frac * (b - a);
}

// Position `value` in the gradient and interpolate (in RGB) between the two
// bracketing stops. Works for 2- and 3-stop scales alike; clamps to the domain.
function colorScaleColor(scale: ResolvedColorScale, value: number): Rgb {
  const { thresholds: t, colors } = scale;
  const last = t.length - 1;
  if (value <= t[0]!) return colors[0]!;
  if (value >= t[last]!) return colors[last]!;
  let i = 0;
  while (i < last && value > t[i + 1]!) i++;
  const a = t[i]!;
  const b = t[i + 1]!;
  const f = b > a ? (value - a) / (b - a) : 0;
  return lerpRgb(colors[i]!, colors[i + 1]!, f);
}

function lerpRgb(c0: Rgb, c1: Rgb, f: number): Rgb {
  return [c0[0] + (c1[0] - c0[0]) * f, c0[1] + (c1[1] - c0[1]) * f, c0[2] + (c1[2] - c0[2]) * f];
}

function parseRgb(hex: string): Rgb | undefined {
  const h = hex.length === 8 ? hex.slice(2) : hex;
  if (!/^[0-9A-Fa-f]{6}$/.test(h)) return undefined;
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

function rgbHex([r, g, b]: Rgb): string {
  const h = (n: number): string =>
    Math.max(0, Math.min(255, Math.round(n)))
      .toString(16)
      .padStart(2, '0')
      .toUpperCase();
  return `${h(r)}${h(g)}${h(b)}`;
}

// --- dataBar (E-SHEET SC1c) -------------------------------------------------

// Resolve a dataBar's lower/upper cfvo stops against the range extent and read
// its length bounds. Returns undefined (rule becomes a no-op) when the range has
// no numbers or a stop won't resolve. minLength/maxLength are percents; absent →
// 0/100 (modern solid bars span the full cell; ECMA's 10/90 default is dropped).
function resolveDataBar(
  rule: CfRuleDataBar,
  ranges: ReadonlyArray<MergedRange>,
  cells: ReadonlyArray<WorksheetCell>,
): ResolvedDataBar | undefined {
  const vals = collectRangeValues(cells, ranges);
  if (vals.length === 0) return undefined;
  const sorted = [...vals].sort((a, b) => a - b);
  const vmin = sorted[0]!;
  const vmax = sorted[sorted.length - 1]!;
  const lower = resolveCfvo(rule.cfvos[0]!, vmin, vmax, sorted);
  const upper = resolveCfvo(rule.cfvos[1]!, vmin, vmax, sorted);
  if (lower === undefined || upper === undefined) return undefined;
  return {
    lower,
    upper,
    colorHex: rule.colorHex,
    minLen: (rule.minLength ?? 0) / 100,
    maxLen: (rule.maxLength ?? 100) / 100,
  };
}

// Bar length (0..1 of cell width) for a value: its position in [lower,upper]
// scaled into [minLen,maxLen]. A degenerate extent (upper≤lower) → full at/above.
function dataBarFraction(bar: ResolvedDataBar, value: number): number {
  const t =
    bar.upper > bar.lower
      ? Math.min(1, Math.max(0, (value - bar.lower) / (bar.upper - bar.lower)))
      : value >= bar.upper
        ? 1
        : 0;
  return bar.minLen + t * (bar.maxLen - bar.minLen);
}

// --- iconSet (E-SHEET SC1c) -------------------------------------------------

// Resolve an iconSet's cfvo thresholds against the range extent (ascending).
function resolveIconThresholds(
  rule: CfRuleIconSet,
  ranges: ReadonlyArray<MergedRange>,
  cells: ReadonlyArray<WorksheetCell>,
): Array<number> | undefined {
  const vals = collectRangeValues(cells, ranges);
  if (vals.length === 0) return undefined;
  const sorted = [...vals].sort((a, b) => a - b);
  const vmin = sorted[0]!;
  const vmax = sorted[sorted.length - 1]!;
  const out: Array<number> = [];
  for (const cfvo of rule.cfvos) {
    const t = resolveCfvo(cfvo, vmin, vmax, sorted);
    if (t === undefined) return undefined;
    out.push(t);
  }
  return out;
}

// The icon bucket for a value: how many thresholds (past the floor) it reaches.
// thresholds[0] is the floor (bucket 0); each later threshold met bumps the bucket.
function iconBucket(thresholds: ReadonlyArray<number>, value: number): number {
  let idx = 0;
  for (let i = 1; i < thresholds.length; i++) {
    if (value >= thresholds[i]!) idx++;
  }
  return idx;
}

// Map Excel's named icon family + bucket onto a format-neutral shape + colour.
// `reverse` flips the bucket so the highest value takes the first icon. The
// `*Gray` families are monochrome (only the direction/shape carries meaning).
function iconToCell(setName: string, count: number, index: number, reverse: boolean): CellIcon {
  const e = reverse ? count - 1 - index : index;
  const colorHex = /Gray/i.test(setName) ? GRAY_HEX : iconColor(count, e);
  return { shape: iconShape(setName, e, count), colorHex };
}

const GRAY_HEX = '808080';

function iconShape(setName: string, e: number, count: number): CellIconShape {
  if (setName.includes('Arrows')) {
    if (e <= 0) return 'triangleDown';
    if (e >= count - 1) return 'triangleUp';
    return 'triangleRight';
  }
  if (setName.includes('Signs')) {
    // 3 Signs: red diamond (low), yellow triangle (mid), green circle (high).
    if (e <= 0) return 'diamond';
    if (e >= count - 1) return 'circle';
    return 'triangleUp';
  }
  if (setName.includes('Flags')) return 'square';
  return 'circle';
}

// A red→green ramp per icon count; the approximation Ream draws for every family
// (faithful glyph sets — signs, symbols, ratings — follow). e is the bucket.
const ICON_RAMPS: Readonly<Record<number, ReadonlyArray<string>>> = {
  3: ['FF0000', 'FFC000', '00B050'],
  4: ['FF0000', 'FF8C00', 'FFC000', '00B050'],
  5: ['FF0000', 'FF8C00', 'FFC000', '92D050', '00B050'],
};

function iconColor(count: number, e: number): string {
  const ramp = ICON_RAMPS[count] ?? ICON_RAMPS[3]!;
  return ramp[Math.max(0, Math.min(ramp.length - 1, e))]!;
}
