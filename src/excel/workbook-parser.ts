// ECMA-376 Part 1 §18.2.20 — workbook.xml.
// We extract the ordered list of sheet references; the renderer currently
// processes only the first sheet.

import { XMLParser } from 'fast-xml-parser';

import type { DefinedName } from '@/core/spreadsheet-model';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  // Some producers (e.g. Haansoft HCell) write SpreadsheetML with an explicit
  // `x:` namespace prefix (<x:workbook>, <x:sheets>, <x:sheet r:id=…>) instead
  // of the default namespace. Strip prefixes so the local-name lookups below
  // (and `r:id` → `id`) match regardless of prefix. ECMA-376 permits any prefix.
  removeNSPrefix: true,
});

/** §18.2.19 `<sheet name sheetId r:id>` — one worksheet reference in tab order. */
export interface SheetReference {
  /** The sheet (tab) display name. */
  readonly name: string;
  /** The workbook-internal sheet id (`@sheetId`); `''` when absent. */
  readonly sheetId: string;
  /** The relationship id (`r:id`) resolving to the sheet's worksheet part. */
  readonly relationshipId: string;
}

// DefinedName (the workbook's named ranges — print areas/titles ride these)
// now lives in @/core/spreadsheet-model; imported above.

/** The parsed `workbook.xml`: the ordered sheet list, the date epoch and defined names. */
export interface ParsedWorkbook {
  /** The worksheet references in tab order. */
  readonly sheets: ReadonlyArray<SheetReference>;
  /**
   * ECMA-376 Part 1 §18.2.28 — `<workbookPr date1904="1"/>`. When true the
   * serial-to-date conversion uses 1904-01-01 as day 0 instead of the
   * standard 1899-12-30 epoch (legacy Mac Excel files).
   */
  readonly date1904: boolean;
  /** §18.2.5 workbook defined names (print areas/titles, named ranges). */
  readonly definedNames: ReadonlyArray<DefinedName>;
}

/**
 * Parse `workbook.xml` into its ordered {@link SheetReference} list plus the date
 * epoch and {@link ParsedWorkbook.definedNames}. A sheet entry missing a name or
 * relationship id is skipped; a malformed root yields an empty workbook.
 */
export function parseWorkbook(data: Uint8Array): ParsedWorkbook {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const workbook = tree['workbook'];
  if (!workbook || typeof workbook !== 'object') {
    return { sheets: [], date1904: false, definedNames: [] };
  }
  const wbObj = workbook as Record<string, unknown>;
  const date1904 = parseDate1904(wbObj);
  const definedNames = parseDefinedNames(wbObj);
  const sheets = wbObj['sheets'];
  if (!sheets || typeof sheets !== 'object') return { sheets: [], date1904, definedNames };
  const sheetRaw = (sheets as Record<string, unknown>)['sheet'];
  const items = Array.isArray(sheetRaw) ? sheetRaw : sheetRaw !== undefined ? [sheetRaw] : [];
  const out: Array<SheetReference> = [];
  for (const item of items) {
    if (!item || typeof item !== 'object') continue;
    const obj = item as Record<string, unknown>;
    const name = strAttr(obj, 'name');
    const sheetId = strAttr(obj, 'sheetId');
    const rId = strAttr(obj, 'r:id') ?? strAttr(obj, 'id');
    if (!name || !rId) continue;
    out.push({ name, sheetId: sheetId ?? '', relationshipId: rId });
  }
  return { sheets: out, date1904, definedNames };
}

function parseDefinedNames(wbObj: Record<string, unknown>): Array<DefinedName> {
  const node = wbObj['definedNames'];
  if (!node || typeof node !== 'object') return [];
  const raw = (node as Record<string, unknown>)['definedName'];
  const items = Array.isArray(raw) ? raw : raw !== undefined ? [raw] : [];
  const out: Array<DefinedName> = [];
  for (const item of items) {
    if (!item || typeof item !== 'object') continue;
    const obj = item as Record<string, unknown>;
    const name = strAttr(obj, 'name');
    if (!name) continue;
    const value = textOf(obj['#text']);
    const localRaw = strAttr(obj, 'localSheetId');
    const localSheetId = localRaw !== undefined ? Number(localRaw) : undefined;
    out.push({
      name,
      value,
      ...(localSheetId !== undefined && Number.isInteger(localSheetId) ? { localSheetId } : {}),
    });
  }
  return out;
}

function textOf(node: unknown): string {
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  return '';
}

function parseDate1904(wbObj: Record<string, unknown>): boolean {
  const pr = wbObj['workbookPr'];
  if (!pr || typeof pr !== 'object') return false;
  const raw = strAttr(pr as Record<string, unknown>, 'date1904');
  return raw === '1' || raw === 'true';
}

function strAttr(obj: Record<string, unknown>, key: string): string | undefined {
  const v = obj[`@_${key}`];
  return typeof v === 'string' ? v : undefined;
}
