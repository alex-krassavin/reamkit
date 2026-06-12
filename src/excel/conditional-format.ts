// Conditional formatting evaluation (E-SHEET SC1) — the first Excel feature the
// flat table projection could not express. A sheet's <conditionalFormatting>
// rules plus the workbook's <dxfs> become a per-cell lookup; the print model
// overrides each cell's base fill/font with the highest-priority matching rule's
// differential format. v1 handles `cellIs` rules (colorScale/dataBar/iconSet,
// which need cross-cell statistics, follow).

import type {
  CfOperator,
  ConditionalFormat,
  Dxf,
  MergedRange,
  XlsxStyles,
} from '@/core/spreadsheet-model';

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

// A resolved per-cell override: a solid highlight fill and/or font tweaks.
export interface CfOverride {
  readonly fillHex?: string;
  readonly fontColorHex?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
}

export type CellConditionalFormatter = (
  row: number,
  col: number,
  numericValue: number | undefined,
) => CfOverride | undefined;

interface FlatRule {
  readonly ranges: ReadonlyArray<MergedRange>;
  readonly rule: ConditionalFormat['rules'][number];
}

// Returns undefined when the sheet has no conditional formats (the common case)
// so callers skip the work entirely and stay byte-identical.
export function buildConditionalFormatter(
  conditionalFormats: ReadonlyArray<ConditionalFormat> | undefined,
  styles: XlsxStyles,
): CellConditionalFormatter | undefined {
  if (!conditionalFormats || conditionalFormats.length === 0) return undefined;
  const dxfs = styles.dxfs ?? [];
  const flat: Array<FlatRule> = [];
  for (const cf of conditionalFormats) {
    for (const rule of cf.rules) flat.push({ ranges: cf.ranges, rule });
  }
  // §18.3.1.10 — a lower @priority wins. Evaluate in ascending priority and take
  // the first rule that both covers the cell and matches.
  flat.sort((a, b) => a.rule.priority - b.rule.priority);

  return (row, col, value) => {
    if (value === undefined) return undefined;
    for (const { ranges, rule } of flat) {
      // Only `cellIs` rules exist today, so `rule` narrows to CfRuleCellIs with
      // no discriminating guard. colorScale/dataBar/iconSet (SC1b) will widen
      // CfRule and reintroduce a `switch (rule.type)` here.
      if (!coversCell(ranges, row, col)) continue;
      if (cellIsMatches(rule.operator, value, rule.formulas)) {
        return dxfToOverride(dxfs[rule.dxfId]);
      }
    }
    return undefined;
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
