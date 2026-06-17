// Slicers (E-SHEET SV2). A slicer is a visual filter panel anchored over a
// sheet. Its definition lives in xl/slicers/slicerN.xml (caption, column count,
// style, and the name of its cache); the cache in xl/slicerCaches/slicerCacheN
// .xml binds the panel to its data source. We resolve native-table slicers — the
// cache's <x14:tableSlicerCache> names a table + column whose distinct values are
// the buttons — and degrade an OLAP/pivot slicer (whose items live in a pivot
// cache) to a caption-only box. removeNSPrefix lets the sml/x14 tags read plainly.

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
});

// §18.3.1 (2009/9/main) <slicer> — one panel. `cacheName` (@cache) keys into the
// workbook's slicer caches; `caption` (@caption) titles the box.
export interface SlicerDef {
  readonly name: string;
  readonly cacheName: string;
  readonly caption: string;
  readonly columnCount: number;
  readonly styleName?: string;
}

// <slicerCacheDefinition> — binds a slicer to its source. A native-table slicer
// carries an <x14:tableSlicerCache tableId column> (the table + column whose
// values are the buttons). `sourceName` is the field/column name.
export interface SlicerCacheDef {
  readonly name: string;
  readonly sourceName?: string;
  readonly tableId?: number;
  readonly columnId?: number;
}

export function parseSlicerPart(data: Uint8Array): Array<SlicerDef> {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const root = asObj(tree['slicers']);
  if (!root) return [];
  const out: Array<SlicerDef> = [];
  for (const s of asArray(root['slicer'])) {
    const o = asObj(s);
    if (!o) continue;
    const name = strAttr(o, 'name');
    const cacheName = strAttr(o, 'cache');
    if (!name || !cacheName) continue;
    const caption = strAttr(o, 'caption') ?? name;
    const styleName = strAttr(o, 'style');
    out.push({
      name,
      cacheName,
      caption,
      columnCount: Math.max(1, numAttr(o, 'columnCount') ?? 1),
      ...(styleName ? { styleName } : {}),
    });
  }
  return out;
}

export function parseSlicerCachePart(data: Uint8Array): SlicerCacheDef | undefined {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const root = asObj(tree['slicerCacheDefinition']);
  if (!root) return undefined;
  const name = strAttr(root, 'name');
  if (!name) return undefined;
  const sourceName = strAttr(root, 'sourceName');
  // Native-table source: extLst → ext → (x14:)slicerCacheDefinition → tableSlicerCache.
  let tableId: number | undefined;
  let columnId: number | undefined;
  const extLst = asObj(root['extLst']);
  if (extLst) {
    for (const ext of asArray(extLst['ext'])) {
      const e = asObj(ext);
      const inner = e ? asObj(e['slicerCacheDefinition']) : undefined;
      const tsc = inner ? asObj(inner['tableSlicerCache']) : undefined;
      if (tsc) {
        tableId = numAttr(tsc, 'tableId');
        columnId = numAttr(tsc, 'column');
        break;
      }
    }
  }
  return {
    name,
    ...(sourceName ? { sourceName } : {}),
    ...(tableId !== undefined ? { tableId } : {}),
    ...(columnId !== undefined ? { columnId } : {}),
  };
}

function asObj(v: unknown): Record<string, unknown> | undefined {
  return v && typeof v === 'object' && !Array.isArray(v)
    ? (v as Record<string, unknown>)
    : undefined;
}

function asArray(v: unknown): Array<unknown> {
  return Array.isArray(v) ? v : v === undefined || v === null ? [] : [v];
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
