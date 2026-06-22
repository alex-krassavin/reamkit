// ECMA-376 §18.5.1.2 — xl/tables/tableN.xml. A structured table over a cell
// range with a named, banded style. We read the range, header-row count,
// autofilter presence and the <tableStyleInfo> flags; the reader resolves the
// named style to header / band fill colours against the workbook theme (SC3).

import { XMLParser } from 'fast-xml-parser';

import type { ExcelTable } from '@/core/spreadsheet-model';

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
 * §18.5.1.1 `<autoFilter><filterColumn colId><filters><filter val>` — the values
 * a column's filter keeps visible. Used only for slicer selection (E-SHEET SV2);
 * not part of the persisted {@link ExcelTable} (so the writer/roundtrip stay
 * unchanged).
 */
export interface TableFilterColumn {
  /** 0-based column offset within the table. */
  readonly colId: number;
  /** The values the filter keeps visible. */
  readonly values: ReadonlyArray<string>;
}

/**
 * The richer parse the reader needs for slicer resolution: the persisted
 * {@link ExcelTable} plus the table's numeric id and any autofilter value filters.
 */
export interface ParsedTablePart {
  readonly table: ExcelTable;
  /** The table's numeric `@id` (matched against a slicer cache's `tableId`). */
  readonly id?: number;
  readonly filters: ReadonlyArray<TableFilterColumn>;
}

/**
 * Parse `xl/tables/tableN.xml` (§18.5.1.2) into the persisted {@link ExcelTable}
 * (range, header-row count, autofilter presence, banded-style flags), or undefined
 * when the root or its `ref` is missing. The slimmer of the two entry points;
 * {@link parseTablePartFull} also returns the id + filters.
 */
export function parseTablePart(data: Uint8Array): ExcelTable | undefined {
  return parseTablePartFull(data)?.table;
}

/**
 * Parse a table part into the full {@link ParsedTablePart} — the {@link ExcelTable}
 * plus the table's numeric id and any autofilter value filters (for slicer
 * resolution). Returns undefined when the root or its `ref` is missing.
 */
export function parseTablePartFull(data: Uint8Array): ParsedTablePart | undefined {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const t = asObj(tree['table']);
  if (!t) return undefined;
  const refStr = strAttr(t, 'ref');
  const ref = refStr ? parseAreaRef(refStr) : undefined;
  if (!ref) return undefined;
  const si = asObj(t['tableStyleInfo']);
  const name = strAttr(t, 'name');
  const styleName = si ? strAttr(si, 'name') : undefined;
  const af = asObj(t['autoFilter']);
  const table: ExcelTable = {
    ref,
    ...(name ? { name } : {}),
    ...(styleName ? { styleName } : {}),
    headerRowCount: numAttr(t, 'headerRowCount') ?? 1,
    showRowStripes: si ? boolAttr(si, 'showRowStripes') : false,
    showColumnStripes: si ? boolAttr(si, 'showColumnStripes') : false,
    showFirstColumn: si ? boolAttr(si, 'showFirstColumn') : false,
    showLastColumn: si ? boolAttr(si, 'showLastColumn') : false,
    autoFilter: af !== undefined,
  };
  const id = numAttr(t, 'id');
  return {
    table,
    ...(id !== undefined ? { id } : {}),
    filters: parseFilterColumns(af),
  };
}

function parseFilterColumns(af: Record<string, unknown> | undefined): Array<TableFilterColumn> {
  if (!af) return [];
  const out: Array<TableFilterColumn> = [];
  for (const fc of asArray(af['filterColumn'])) {
    const col = asObj(fc);
    if (!col) continue;
    const colId = numAttr(col, 'colId');
    if (colId === undefined) continue;
    const filters = asObj(col['filters']);
    if (!filters) continue;
    const values: Array<string> = [];
    for (const f of asArray(filters['filter'])) {
      const v = asObj(f) ? strAttr(asObj(f)!, 'val') : undefined;
      if (v !== undefined) values.push(v);
    }
    out.push({ colId, values });
  }
  return out;
}

function asArray(v: unknown): Array<unknown> {
  return Array.isArray(v) ? v : v === undefined || v === null ? [] : [v];
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
