// Helpers for navigating fast-xml-parser preserveOrder trees.
//
// Format reminder:
//   { "w:p": [child1, child2, ...], ":@": { "@_w:val": "..." } }
// Each element has exactly one tag-name key (its tag) plus an optional ":@"
// attributes object. Children include leaf text nodes of shape
//   { "#text": "..." }.

export type PoNode = Record<string, unknown>;
export type PoTree = ReadonlyArray<PoNode>;

const ATTRS_KEY = ':@';
const TEXT_KEY = '#text';

export function poTag(node: PoNode | undefined): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  for (const key of Object.keys(node)) {
    if (key !== ATTRS_KEY && key !== TEXT_KEY) return key;
  }
  return undefined;
}

export function poIs(node: PoNode | undefined, tag: string): boolean {
  return poTag(node) === tag;
}

export function poChildren(node: PoNode | undefined): ReadonlyArray<PoNode> {
  const tag = poTag(node);
  if (!tag || !node) return [];
  const arr = node[tag];
  return Array.isArray(arr) ? (arr as Array<PoNode>) : [];
}

export function poChildrenWith(node: PoNode | undefined, tag: string): Array<PoNode> {
  return poChildren(node).filter((c) => poIs(c, tag));
}

export function poFirstChild(node: PoNode | undefined, tag: string): PoNode | undefined {
  return poChildren(node).find((c) => poIs(c, tag));
}

export function poAttr(node: PoNode | undefined, name: string): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const attrs = node[ATTRS_KEY];
  if (!attrs || typeof attrs !== 'object') return undefined;
  const a = attrs as Record<string, unknown>;
  // Try common namespaces in priority: w:, r:, m: (OfficeMath), xml:, then bare.
  // Header/footer references carry the relationship id under r:id; OfficeMath
  // properties carry their value under m:val (e.g. m:chr, m:type, m:sty).
  const v =
    a[`@_w:${name}`] ??
    a[`@_r:${name}`] ??
    a[`@_m:${name}`] ??
    a[`@_xml:${name}`] ??
    a[`@_${name}`];
  return typeof v === 'string' ? v : undefined;
}

export function poIntAttr(node: PoNode | undefined, name: string): number | undefined {
  const v = poAttr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

export function poVal(node: PoNode | undefined): string | undefined {
  return poAttr(node, 'val');
}

export function poToggle(node: PoNode | undefined): boolean | undefined {
  if (!node) return undefined;
  const v = poAttr(node, 'val');
  if (v === undefined) return true;
  return !(v === 'false' || v === '0' || v === 'off');
}

export function poText(node: PoNode | undefined): string {
  if (!node) return '';
  const children = poChildren(node);
  let out = '';
  for (const c of children as ReadonlyArray<unknown>) {
    if (c && typeof c === 'object' && TEXT_KEY in c) {
      const t = (c as Record<string, unknown>)[TEXT_KEY];
      if (typeof t === 'string') out += t;
      else if (typeof t === 'number') out += String(t);
    }
  }
  return out;
}

export function poFindByPath(tree: PoTree, path: ReadonlyArray<string>): PoNode | undefined {
  let cursor: ReadonlyArray<PoNode> = tree;
  let current: PoNode | undefined;
  for (const seg of path) {
    current = cursor.find((n) => poIs(n, seg));
    if (!current) return undefined;
    cursor = poChildren(current);
  }
  return current;
}

// Depth-first search for the first descendant with the given tag.
export function poFindDescendant(node: PoNode | undefined, tag: string): PoNode | undefined {
  if (!node) return undefined;
  for (const child of poChildren(node)) {
    if (poIs(child, tag)) return child;
    const found = poFindDescendant(child, tag);
    if (found) return found;
  }
  return undefined;
}
