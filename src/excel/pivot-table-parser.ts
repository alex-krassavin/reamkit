// ECMA-376 §18.10.1.73 — xl/pivotTables/pivotTableN.xml. A pivot table over an
// output range with a named, banded style. Its OUTPUT cells are already cached
// in the worksheet (so they render as a normal grid); we read the <location>
// (range + header/data offsets) and the <pivotTableStyleInfo> flags so the
// reader can band the region in the pivot's own palette (E-PIVOT). The reader
// resolves the named style to fill colours against the workbook theme (PV2).

import { XMLParser } from 'fast-xml-parser';

import type { PivotTable } from '@/core/spreadsheet-model';

import { parseAreaRef } from '@/excel/defined-name-ref';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
});

/**
 * Parse `xl/pivotTables/pivotTableN.xml` (§18.10.1.73) into a {@link PivotTable}:
 * the output `<location>` (range + header/data offsets), the banded-style flags,
 * and the row/column total markers derived from `<rowItems>`/`<colItems>` (so the
 * reader can band the region in the pivot palette — E-PIVOT). The output cells are
 * already cached in the worksheet, so they render as a normal grid. Returns
 * undefined when the root or its `ref` is missing.
 */
export function parsePivotTablePart(data: Uint8Array): PivotTable | undefined {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const def = asObj(tree['pivotTableDefinition']);
  if (!def) return undefined;
  const loc = asObj(def['location']);
  const refStr = loc ? strAttr(loc, 'ref') : undefined;
  const ref = refStr ? parseAreaRef(refStr) : undefined;
  if (!ref) return undefined;
  const si = asObj(def['pivotTableStyleInfo']);
  const name = strAttr(def, 'name');
  const styleName = si ? strAttr(si, 'name') : undefined;
  // <rowItems>/<colItems> — one <i> per data row / column (in order); @t marks
  // total lines ('grand' = grand total, a subtotal-function name = subtotal,
  // absent = data). The i-th type maps to data row/col `firstData{Row,Col} + i`
  // (E-PIVOT PV3/PV4).
  const rowItemTypes = itemTypes(asObj(def['rowItems']));
  const colItemTypes = itemTypes(asObj(def['colItems']));
  const hasRowTotals = rowItemTypes?.some((t) => t !== undefined) ?? false;
  const hasColTotals = colItemTypes?.some((t) => t !== undefined) ?? false;
  return {
    ref,
    ...(name ? { name } : {}),
    ...(styleName ? { styleName } : {}),
    firstHeaderRow: (loc ? numAttr(loc, 'firstHeaderRow') : undefined) ?? 1,
    firstDataRow: (loc ? numAttr(loc, 'firstDataRow') : undefined) ?? 1,
    firstDataCol: (loc ? numAttr(loc, 'firstDataCol') : undefined) ?? 0,
    showRowStripes: si ? boolAttr(si, 'showRowStripes') : false,
    showColStripes: si ? boolAttr(si, 'showColStripes') : false,
    ...(hasRowTotals && rowItemTypes ? { rowItemTypes } : {}),
    ...(hasColTotals && colItemTypes ? { colItemTypes } : {}),
  };
}

// The @t of each <i> in a <rowItems>/<colItems> list, in order.
function itemTypes(
  node: Record<string, unknown> | undefined,
): Array<string | undefined> | undefined {
  if (!node) return undefined;
  return toArray(node['i']).map((i) => {
    const o = asObj(i);
    return o ? strAttr(o, 't') : undefined;
  });
}

function toArray(v: unknown): Array<unknown> {
  if (v === undefined || v === null) return [];
  return Array.isArray(v) ? v : [v];
}

function asObj(v: unknown): Record<string, unknown> | undefined {
  return v && typeof v === 'object' && !Array.isArray(v)
    ? (v as Record<string, unknown>)
    : undefined;
}

function strAttr(obj: Record<string, unknown>, name: string): string | undefined {
  const v = obj[`@_${name}`];
  return typeof v === 'string' ? v : undefined;
}

function numAttr(obj: Record<string, unknown>, name: string): number | undefined {
  const v = strAttr(obj, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

function boolAttr(obj: Record<string, unknown>, name: string): boolean {
  const v = strAttr(obj, name);
  return v === '1' || v === 'true';
}
