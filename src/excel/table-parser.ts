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

export function parseTablePart(data: Uint8Array): ExcelTable | undefined {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const t = asObj(tree['table']);
  if (!t) return undefined;
  const refStr = strAttr(t, 'ref');
  const ref = refStr ? parseAreaRef(refStr) : undefined;
  if (!ref) return undefined;
  const si = asObj(t['tableStyleInfo']);
  const name = strAttr(t, 'name');
  const styleName = si ? strAttr(si, 'name') : undefined;
  return {
    ref,
    ...(name ? { name } : {}),
    ...(styleName ? { styleName } : {}),
    headerRowCount: numAttr(t, 'headerRowCount') ?? 1,
    showRowStripes: si ? boolAttr(si, 'showRowStripes') : false,
    showColumnStripes: si ? boolAttr(si, 'showColumnStripes') : false,
    showFirstColumn: si ? boolAttr(si, 'showFirstColumn') : false,
    showLastColumn: si ? boolAttr(si, 'showLastColumn') : false,
    autoFilter: asObj(t['autoFilter']) !== undefined,
  };
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
