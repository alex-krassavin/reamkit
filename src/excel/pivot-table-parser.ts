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
  // <rowItems>/<i> — one item per data row (in order); @t marks total rows
  // ('grand' = grand total, a subtotal-function name = subtotal, absent = data).
  // The i-th type maps to data row `firstDataRow + i` (E-PIVOT PV3).
  const rowItemsNode = asObj(def['rowItems']);
  const rowItemTypes = rowItemsNode
    ? toArray(rowItemsNode['i']).map((i) => {
        const o = asObj(i);
        return o ? strAttr(o, 't') : undefined;
      })
    : undefined;
  const hasTotals = rowItemTypes?.some((t) => t !== undefined) ?? false;
  return {
    ref,
    ...(name ? { name } : {}),
    ...(styleName ? { styleName } : {}),
    firstHeaderRow: (loc ? numAttr(loc, 'firstHeaderRow') : undefined) ?? 1,
    firstDataRow: (loc ? numAttr(loc, 'firstDataRow') : undefined) ?? 1,
    firstDataCol: (loc ? numAttr(loc, 'firstDataCol') : undefined) ?? 0,
    showRowStripes: si ? boolAttr(si, 'showRowStripes') : false,
    showColStripes: si ? boolAttr(si, 'showColStripes') : false,
    ...(hasTotals && rowItemTypes ? { rowItemTypes } : {}),
  };
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
