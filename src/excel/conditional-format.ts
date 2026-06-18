// Conditional formatting evaluation (E-SHEET SC1/SC1b/SC1c + W5) — the first Excel
// feature the flat table projection could not express. A sheet's
// <conditionalFormatting> rules plus the workbook's <dxfs> become a per-cell
// lookup; the print model overrides each cell's base format. Handles the
// value-driven families: `cellIs` (compare to a constant → dxf), `colorScale`
// (interpolate a fill across the range's value extent), `dataBar` (an in-cell bar
// whose length encodes the value) and `iconSet`; plus the W5 families that also
// resolve against the range — `top10` (top/bottom N or N%), `aboveAverage`
// (mean ± N·σ), `duplicate/uniqueValues` (value frequency) — and the per-cell
// text tests (`containsText`/`beginsWith`/`endsWith`/`notContainsText`). A cell's
// text format (fill/font) is claimed by the first applicable highlight rule; a
// dataBar/iconSet applies independently on top. W9 adds `expression` (an
// arbitrary formula evaluated per cell by the formula engine over the cached grid
// values) and `timePeriod` (a clock-relative date window against an injected
// reference day) — both deterministic, no recalculation, no wall clock.

import type { CellIcon, CellIconShape } from '@/core/document-model';
import type {
  CfOperator,
  CfRuleAboveAverage,
  CfRuleColorScale,
  CfRuleDataBar,
  CfRuleDupUnique,
  CfRuleIconSet,
  CfRuleText,
  CfRuleTop10,
  Cfvo,
  ConditionalFormat,
  DefinedName,
  Dxf,
  MergedRange,
  WorksheetCell,
  XlsxStyles,
} from '@/core/spreadsheet-model';
import type { CompiledFormula, EvalContext, FErr, FValue, Rect, Scalar } from '@/excel/formula';

import {
  BLANK,
  NO_SHIFT,
  bool,
  compileFormula,
  err,
  evaluate,
  evaluateToBool,
  num,
  serialFromDate,
  str,
  timePeriodMatches,
} from '@/excel/formula';

type Rgb = readonly [number, number, number];

// A colorScale resolved against its range's value extent: N monotonic
// thresholds paired with N pre-parsed RGB stop colours (N = 2 or 3).
interface ResolvedColorScale {
  readonly thresholds: ReadonlyArray<number>;
  readonly colors: ReadonlyArray<Rgb>;
}

// A dataBar resolved against its range's extent: the lower/upper value the bar
// scale spans, its positive/negative fill colours, and the min/max bar length
// (0..1 of cell width).
interface ResolvedDataBar {
  readonly lower: number;
  readonly upper: number;
  readonly colorHex: string;
  readonly negativeColorHex: string;
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
  readonly dataBar?: {
    readonly fraction: number;
    readonly colorHex: string;
    readonly startFraction?: number;
  };
  readonly icon?: CellIcon;
}

export type CellConditionalFormatter = (
  row: number,
  col: number,
  numericValue: number | undefined,
  // The cell's resolved text — needed by the text tests and to key duplicate /
  // unique comparisons for non-numeric cells. Empty/undefined for a blank cell.
  text: string | undefined,
) => CfOverride | undefined;

// A resolved top10/aboveAverage threshold: a value matches by comparing against
// `threshold` per the rule's direction (and inclusivity).
interface ResolvedThreshold {
  readonly threshold: number;
  readonly below: boolean; // bottom-N / below-average
  readonly orEqual: boolean; // include values exactly at the threshold
}

interface FlatRule {
  readonly ranges: ReadonlyArray<MergedRange>;
  readonly rule: ConditionalFormat['rules'][number];
  // Present (and possibly undefined, when resolution failed) for colorScale /
  // dataBar / iconSet rules: the extent-derived data computed once from the
  // range values (gradient stops, bar bounds, or icon bucket thresholds).
  readonly scale?: ResolvedColorScale | undefined;
  readonly bar?: ResolvedDataBar | undefined;
  readonly iconThresholds?: ReadonlyArray<number> | undefined;
  // top10 / aboveAverage: the value threshold resolved from the range (W5).
  readonly threshold?: ResolvedThreshold | undefined;
  // duplicate/uniqueValues: the set of value keys that qualify (W5).
  readonly dupKeys?: ReadonlySet<string> | undefined;
  // expression (W9): the formula compiled once. undefined when it failed to
  // parse — the rule then never applies (graceful loss).
  readonly compiled?: CompiledFormula | undefined;
  // expression (W9): the rule's relative-reference origin (the sqref's first
  // cell). A covered cell shifts the formula's unanchored refs by its offset
  // from here, exactly as Excel anchors a conditional-format expression.
  readonly origin?: { readonly row: number; readonly col: number };
}

// Returns undefined when the sheet has no conditional formats (the common case)
// so callers skip the work entirely and stay byte-identical. `cells` supplies
// the range values the extent rules need (min/max/percentile/mean/frequency);
// cellIs and the text tests ignore it. `resolveText` resolves a cell's string
// value (shared strings / number format) for the duplicate/unique frequency map —
// numeric cells key by value without it, so it is only needed for text cells.
// `date1904`/`now` (E-SHEET W9) feed the formula engine: `now` (an injected
// reference date — never the wall clock) drives TODAY()/NOW() and the timePeriod
// windows; absent ⇒ those constructs no-op, so the output stays deterministic.
export function buildConditionalFormatter(
  conditionalFormats: ReadonlyArray<ConditionalFormat> | undefined,
  styles: XlsxStyles,
  cells: ReadonlyArray<WorksheetCell>,
  resolveText?: (cell: WorksheetCell) => string,
  date1904 = false,
  now?: Date,
  // The whole workbook, for an `expression` rule that reaches another sheet
  // (Sheet2!A1) or a defined name. Absent ⇒ same-sheet references only (#REF!/
  // #NAME? for those, exactly as before).
  sheetGrids?: ReadonlyMap<string, { readonly cells: ReadonlyArray<WorksheetCell> }>,
  currentSheet?: string,
  definedNames?: ReadonlyArray<DefinedName>,
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
      } else if (rule.type === 'top10') {
        flat.push({ ranges: cf.ranges, rule, threshold: resolveTop10(rule, cf.ranges, cells) });
      } else if (rule.type === 'aboveAverage') {
        flat.push({
          ranges: cf.ranges,
          rule,
          threshold: resolveAboveAverage(rule, cf.ranges, cells),
        });
      } else if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') {
        flat.push({
          ranges: cf.ranges,
          rule,
          dupKeys: resolveDupUnique(rule, cf.ranges, cells, resolveText),
        });
      } else if (rule.type === 'expression') {
        // W9: compile the formula once; remember the sqref's first cell as the
        // relative-reference origin for the per-cell shift.
        const first = cf.ranges[0];
        flat.push({
          ranges: cf.ranges,
          rule,
          compiled: compileFormula(rule.formula),
          ...(first ? { origin: { row: first.startRow, col: first.startColumn } } : {}),
        });
      } else {
        // cellIs / the text tests / timePeriod: matched per-cell, no precompute.
        flat.push({ ranges: cf.ranges, rule });
      }
    }
  }
  // §18.3.1.10 — a lower @priority wins. Evaluate in ascending priority. A cell's
  // text format (fill/font) is claimed by the first applicable cellIs/colorScale;
  // a dataBar fills its own slot independently, so a bar can sit over a fill.
  flat.sort((a, b) => a.rule.priority - b.rule.priority);

  // W9: an `expression` rule can apply to a cell with no value of its own (it may
  // reference neighbours), so when one is present we cannot take the empty-cell
  // shortcut. The evaluator context (grid lookup) is built only then, keeping
  // every existing sheet's path byte-identical. `nowSerial` is the injected day.
  const hasExpr = flat.some((f) => f.rule.type === 'expression');
  const nowSerial = now !== undefined ? serialFromDate(now, date1904) : undefined;
  // The cross-sheet / defined-name layer is built only for an expression rule
  // that has the workbook wired in — keeping the common path byte-identical.
  const cross =
    hasExpr && sheetGrids
      ? {
          sheets: [...sheetGrids].map(([name, ws]) => ({ name, cells: ws.cells })),
          ...(currentSheet !== undefined ? { currentSheet } : {}),
          definedNames: definedNames ?? [],
        }
      : undefined;
  const ctx = hasExpr
    ? buildEvalContext(cells, resolveText, nowSerial, date1904, cross)
    : undefined;

  return (row, col, value, text) => {
    // A cell with neither a comparable number nor any text matches nothing — skip
    // the loop entirely so number-only sheets stay byte-identical to before W5.
    // (An expression rule may still target an empty cell, so keep going then.)
    if (!hasExpr && value === undefined && (text === undefined || text.length === 0)) {
      return undefined;
    }
    const dupKey = dupKeyOf(value, text);
    let textFmt: CfOverride | undefined;
    let textClaimed = false;
    let bar: CfOverride['dataBar'];
    let icon: CfOverride['icon'];
    const claim = (dxfId: number): void => {
      textFmt = dxfToOverride(dxfs[dxfId]);
      textClaimed = true;
    };
    for (const {
      ranges,
      rule,
      scale,
      bar: resolved,
      iconThresholds,
      threshold,
      dupKeys,
      compiled,
      origin,
    } of flat) {
      if (!coversCell(ranges, row, col)) continue;
      switch (rule.type) {
        case 'cellIs':
          if (
            !textClaimed &&
            value !== undefined &&
            cellIsMatches(rule.operator, value, rule.formulas)
          ) {
            claim(rule.dxfId);
          }
          break;
        case 'colorScale':
          if (!textClaimed && value !== undefined && scale) {
            textFmt = { fillHex: rgbHex(colorScaleColor(scale, value)) };
            textClaimed = true;
          }
          break;
        case 'dataBar':
          if (!bar && value !== undefined && resolved) {
            bar = dataBarBar(resolved, value);
          }
          break;
        case 'iconSet':
          if (!icon && value !== undefined && iconThresholds) {
            const bucket = iconBucket(iconThresholds, value);
            icon = iconToCell(rule.iconSet, iconThresholds.length, bucket, rule.reverse ?? false);
          }
          break;
        case 'top10':
        case 'aboveAverage':
          if (
            !textClaimed &&
            value !== undefined &&
            threshold &&
            thresholdMatches(threshold, value)
          ) {
            claim(rule.dxfId);
          }
          break;
        case 'duplicateValues':
        case 'uniqueValues':
          if (!textClaimed && dupKey !== undefined && dupKeys?.has(dupKey)) {
            claim(rule.dxfId);
          }
          break;
        case 'containsText':
        case 'notContainsText':
        case 'beginsWith':
        case 'endsWith':
          if (!textClaimed && text !== undefined && textRuleMatches(rule, text)) {
            claim(rule.dxfId);
          }
          break;
        case 'expression':
          // W9: evaluate the compiled formula at this cell, shifting unanchored
          // references by the cell's offset from the rule origin.
          if (!textClaimed && compiled && ctx && origin) {
            const shift = { dRow: row - origin.row, dCol: col - origin.col };
            if (evaluateToBool(compiled, ctx, shift)) claim(rule.dxfId);
          }
          break;
        case 'timePeriod':
          // W9: the cell's serial date against the window relative to options.now.
          if (
            !textClaimed &&
            value !== undefined &&
            nowSerial !== undefined &&
            timePeriodMatches(rule.timePeriod, value, nowSerial, date1904)
          ) {
            claim(rule.dxfId);
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

// --- expression engine context (E-SHEET W9) ---------------------------------

// Build a formula EvalContext over the worksheet's cached cell values. The grid
// is read once into a (row,col) → scalar map (point lookups) plus a flat list of
// the populated cells (range scans for SUM/COUNTIF). It is read-only and pure —
// precisely the cached-value lookup the engine needs, with no recalculation.
interface SheetMap {
  readonly byKey: Map<number, Scalar>;
  readonly populated: Array<{ row: number; col: number; value: Scalar }>;
}
const COLS = 16384; // key stride: column is 0..16383 (XFD)

// Read a sheet's cached cells into a (row,col)→scalar map plus a populated list.
function indexCells(
  cells: ReadonlyArray<WorksheetCell>,
  resolveText: ((cell: WorksheetCell) => string) | undefined,
): SheetMap {
  const byKey = new Map<number, Scalar>();
  const populated: Array<{ row: number; col: number; value: Scalar }> = [];
  for (const cell of cells) {
    const s = cellToScalar(cell, resolveText);
    if (s.t === 'blank') continue;
    byKey.set(cell.row * COLS + cell.column, s);
    populated.push({ row: cell.row, col: cell.column, value: s });
  }
  return { byKey, populated };
}

function eachInMap(
  map: SheetMap,
  rect: Rect,
  visit: (r: number, c: number, v: Scalar) => void,
): void {
  for (const c of map.populated) {
    if (c.row >= rect.r0 && c.row <= rect.r1 && c.col >= rect.c0 && c.col <= rect.c1) {
      visit(c.row, c.col, c.value);
    }
  }
}

function buildEvalContext(
  cells: ReadonlyArray<WorksheetCell>,
  resolveText: ((cell: WorksheetCell) => string) | undefined,
  nowSerial: number | undefined,
  date1904: boolean,
  cross?: {
    readonly sheets: ReadonlyArray<{ name: string; cells: ReadonlyArray<WorksheetCell> }>;
    readonly currentSheet?: string;
    readonly definedNames: ReadonlyArray<DefinedName>;
  },
): EvalContext {
  const self = indexCells(cells, resolveText);
  const base: EvalContext = {
    nowSerial,
    date1904,
    getCell: (row, col) => self.byKey.get(row * COLS + col) ?? BLANK,
    eachCell: (rect, visit) => eachInMap(self, rect, visit),
  };
  if (!cross) return base;

  // A sheet name (lower-cased) → its index; its scalar map is built on first use.
  const sheetByName = new Map<string, number>();
  cross.sheets.forEach((s, i) => sheetByName.set(s.name.toLowerCase(), i));
  const maps: Array<SheetMap | undefined> = new Array<SheetMap | undefined>(cross.sheets.length);
  const sheetData = (idx: number): SheetMap =>
    (maps[idx] ??= indexCells(cross.sheets[idx]!.cells, resolveText));
  const currentIdx =
    cross.currentSheet !== undefined
      ? sheetByName.get(cross.currentSheet.toLowerCase())
      : undefined;

  // A defined name (lower-cased) → its descriptor, preferring a name scoped to the
  // current sheet over a workbook-scoped one; other sheets' local names are skipped.
  const nameMap = new Map<string, DefinedName>();
  for (const dn of cross.definedNames) {
    if (dn.localSheetId !== undefined && dn.localSheetId !== currentIdx) continue;
    const key = dn.name.toLowerCase();
    if (dn.localSheetId !== undefined || !nameMap.has(key)) nameMap.set(key, dn);
  }

  const resolving = new Set<string>(); // guards a name that references another name
  const ctx: EvalContext = {
    ...base,
    sheetIndex: (name) => sheetByName.get(name.toLowerCase()),
    getCellOn: (sheet, row, col) => sheetData(sheet).byKey.get(row * COLS + col) ?? BLANK,
    eachCellOn: (sheet, rect, visit) => eachInMap(sheetData(sheet), rect, visit),
    resolveName: (rawName): FValue | undefined => {
      const key = rawName.toLowerCase();
      const dn = nameMap.get(key);
      if (!dn || resolving.has(key)) return undefined;
      const value = dn.value.startsWith('=') ? dn.value.slice(1) : dn.value;
      const compiled = compileFormula(value);
      if (!compiled) return undefined;
      const ast = compiled.ast;
      if (ast.k === 'num') return num(ast.v);
      if (ast.k === 'str') return str(ast.v);
      if (ast.k === 'bool') return bool(ast.v);
      // Only a plain reference name resolves; a name that is itself a formula no-ops.
      if (ast.k !== 'cell' && ast.k !== 'range') return undefined;
      resolving.add(key);
      try {
        return evaluate(ast, ctx, NO_SHIFT);
      } finally {
        resolving.delete(key);
      }
    },
  };
  return ctx;
}

// A stored worksheet cell → a formula scalar (its cached value). Strings resolve
// through `resolveText` (shared strings / inline / number format); without a
// resolver a shared-string cell falls back to its raw text. A non-finite number
// or an empty non-string cell is treated as blank.
function cellToScalar(
  cell: WorksheetCell,
  resolveText: ((cell: WorksheetCell) => string) | undefined,
): Scalar {
  switch (cell.type) {
    case 'n': {
      const n = Number(cell.rawValue);
      return Number.isFinite(n) ? num(n) : BLANK;
    }
    case 'b':
      return bool(cell.rawValue === '1' || cell.rawValue.toUpperCase() === 'TRUE');
    case 'e':
      return err(toFErr(cell.rawValue));
    case 's':
    case 'str':
    case 'inlineStr':
      return str(resolveText ? resolveText(cell) : (cell.inlineText ?? cell.rawValue));
    default:
      // 'd' (ISO date text) and any other type: expose the raw text verbatim.
      return cell.rawValue.length > 0 ? str(cell.rawValue) : BLANK;
  }
}

const FERRS: ReadonlySet<string> = new Set<FErr>([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
]);

// Map a stored error cell's raw text to a known error value (default #VALUE!).
function toFErr(raw: string): FErr {
  return (FERRS.has(raw) ? raw : '#VALUE!') as FErr;
}

// The duplicate/unique comparison key for a cell at evaluation time. Numbers key
// by value (`n:`), text case-insensitively (`s:`) — the two are namespaced so a
// numeric 5 and the string "5" never collide (Excel treats them as distinct).
function dupKeyOf(value: number | undefined, text: string | undefined): string | undefined {
  if (value !== undefined) return `n:${value}`;
  if (text === undefined) return undefined;
  const k = text.trim().toLowerCase();
  return k.length > 0 ? `s:${k}` : undefined;
}

// A case-insensitive substring test (W5). containsText / notContainsText match
// anywhere; beginsWith / endsWith anchor to the ends. Excel's SEARCH is
// case-insensitive, so the comparison folds case.
function textRuleMatches(rule: CfRuleText, text: string): boolean {
  const hay = text.toLowerCase();
  const needle = rule.text.toLowerCase();
  switch (rule.type) {
    case 'containsText':
      return hay.includes(needle);
    case 'notContainsText':
      return !hay.includes(needle);
    case 'beginsWith':
      return hay.startsWith(needle);
    case 'endsWith':
      return hay.endsWith(needle);
  }
}

function thresholdMatches(t: ResolvedThreshold, value: number): boolean {
  if (t.below) return t.orEqual ? value <= t.threshold : value < t.threshold;
  return t.orEqual ? value >= t.threshold : value > t.threshold;
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
    negativeColorHex: 'FF0000', // Excel's default negative bar colour
    minLen: (rule.minLength ?? 0) / 100,
    maxLen: (rule.maxLength ?? 100) / 100,
  };
}

// The bar for a value: width (0..1 of cell) + optional start offset + colour.
// A mixed-sign extent (lower<0<upper) puts an axis at zero inside the cell —
// positives run right of it, negatives run left in the negative colour (tail
// TC4). A single-sign extent fills from the left edge (the common case).
function dataBarBar(
  bar: ResolvedDataBar,
  value: number,
): { fraction: number; colorHex: string; startFraction?: number } {
  if (bar.lower < 0 && bar.upper > 0) {
    const axis = -bar.lower / (bar.upper - bar.lower); // zero's position in [0,1]
    if (value >= 0) {
      const w = (value / bar.upper) * (1 - axis);
      return { fraction: w, colorHex: bar.colorHex, startFraction: axis };
    }
    const w = (Math.abs(value) / Math.abs(bar.lower)) * axis;
    return { fraction: w, colorHex: bar.negativeColorHex, startFraction: axis - w };
  }
  return { fraction: dataBarFraction(bar, value), colorHex: bar.colorHex };
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
  // Meter families are monochrome — only the amount filled carries meaning, so
  // they bypass the red→green ramp. Ratings light bucket+1 of `count` bars; a
  // quarter pie fills `e` of its `count-1` slices (bucket 0 = empty circle).
  if (/Rating/i.test(setName)) {
    return { shape: 'bars', colorHex: METER_HEX, fill: { filled: e + 1, levels: count } };
  }
  if (/Quarters/i.test(setName)) {
    return { shape: 'pie', colorHex: METER_HEX, fill: { filled: e, levels: count - 1 } };
  }
  const colorHex = /Gray/i.test(setName) ? GRAY_HEX : iconColor(count, e);
  return { shape: iconShape(setName, e, count), colorHex };
}

const GRAY_HEX = '808080';
// Excel's rating/quarter glyphs are drawn near-black, not on the colour ramp.
const METER_HEX = '595959';

function iconShape(setName: string, e: number, count: number): CellIconShape {
  if (setName.includes('Arrows')) {
    if (e <= 0) return 'triangleDown';
    if (e >= count - 1) return 'triangleUp';
    return 'triangleRight';
  }
  if (setName.includes('Symbols')) {
    // 3 Symbols: red cross (low), yellow exclamation (mid), green check (high).
    if (e <= 0) return 'cross';
    if (e >= count - 1) return 'check';
    return 'exclamation';
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

// A red→green ramp per icon count, used by the shape families (lights, arrows,
// signs, symbols, flags). The meter families above are monochrome. e is the bucket.
const ICON_RAMPS: Readonly<Record<number, ReadonlyArray<string>>> = {
  3: ['FF0000', 'FFC000', '00B050'],
  4: ['FF0000', 'FF8C00', 'FFC000', '00B050'],
  5: ['FF0000', 'FF8C00', 'FFC000', '92D050', '00B050'],
};

function iconColor(count: number, e: number): string {
  const ramp = ICON_RAMPS[count] ?? ICON_RAMPS[3]!;
  return ramp[Math.max(0, Math.min(ramp.length - 1, e))]!;
}

// --- top10 / aboveAverage / duplicate-unique (E-SHEET W5) -------------------

// §18.3.1.10 top10 — resolve the cutoff value for the top (or bottom) N / N% of
// the range. A cell qualifies when it is at or beyond that cutoff, so ties at the
// boundary all match (as in Excel). Returns undefined (rule no-op) on an empty
// range. `rank` 0 clamps to 1; a percent rank floors the count.
function resolveTop10(
  rule: CfRuleTop10,
  ranges: ReadonlyArray<MergedRange>,
  cells: ReadonlyArray<WorksheetCell>,
): ResolvedThreshold | undefined {
  const vals = collectRangeValues(cells, ranges);
  if (vals.length === 0) return undefined;
  const n = vals.length;
  const raw = rule.percent ? Math.floor((n * rule.rank) / 100) : rule.rank;
  const count = Math.max(1, Math.min(n, raw));
  // bottom → ascending so sorted[count-1] is the N-th smallest; top → descending.
  const sorted = [...vals].sort((a, b) => (rule.bottom ? a - b : b - a));
  return { threshold: sorted[count - 1]!, below: rule.bottom, orEqual: true };
}

// §18.3.1.10 aboveAverage — resolve the threshold to the range mean, shifted by
// N population standard deviations when `stdDev` is set. `equalAverage` makes the
// comparison inclusive of the threshold itself.
function resolveAboveAverage(
  rule: CfRuleAboveAverage,
  ranges: ReadonlyArray<MergedRange>,
  cells: ReadonlyArray<WorksheetCell>,
): ResolvedThreshold | undefined {
  const vals = collectRangeValues(cells, ranges);
  if (vals.length === 0) return undefined;
  const mean = vals.reduce((s, v) => s + v, 0) / vals.length;
  let threshold = mean;
  if (rule.stdDev !== undefined && rule.stdDev > 0) {
    const variance = vals.reduce((s, v) => s + (v - mean) * (v - mean), 0) / vals.length;
    const sd = Math.sqrt(variance);
    threshold = rule.aboveAverage ? mean + rule.stdDev * sd : mean - rule.stdDev * sd;
  }
  return { threshold, below: !rule.aboveAverage, orEqual: rule.equalAverage };
}

// §18.3.1.10 duplicate/uniqueValues — the set of value keys that repeat within the
// range (duplicate) or occur exactly once (unique). Numbers key by value, text
// case-insensitively; blanks (absent from the sparse cell list) never qualify.
function resolveDupUnique(
  rule: CfRuleDupUnique,
  ranges: ReadonlyArray<MergedRange>,
  cells: ReadonlyArray<WorksheetCell>,
  resolveText: ((cell: WorksheetCell) => string) | undefined,
): ReadonlySet<string> | undefined {
  const counts = new Map<string, number>();
  for (const cell of cells) {
    if (!coversCell(ranges, cell.row, cell.column)) continue;
    const key = cellDupKey(cell, resolveText);
    if (key === undefined) continue;
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  if (counts.size === 0) return undefined;
  const qualifies =
    rule.type === 'duplicateValues' ? (c: number) => c >= 2 : (c: number) => c === 1;
  const out = new Set<string>();
  for (const [k, c] of counts) if (qualifies(c)) out.add(k);
  return out;
}

// The duplicate/unique key for a stored cell — mirrors dupKeyOf but reads the raw
// cell: numbers from rawValue, text via the supplied resolver (shared strings /
// inline / number format). Without a resolver, only numeric cells get a key.
function cellDupKey(
  cell: WorksheetCell,
  resolveText: ((cell: WorksheetCell) => string) | undefined,
): string | undefined {
  if (cell.type === 'n') {
    const n = Number(cell.rawValue);
    return Number.isFinite(n) ? `n:${n}` : undefined;
  }
  if (!resolveText) return undefined;
  const k = resolveText(cell).trim().toLowerCase();
  return k.length > 0 ? `s:${k}` : undefined;
}
