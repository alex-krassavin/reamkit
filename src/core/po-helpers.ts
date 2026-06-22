// Helpers for navigating fast-xml-parser preserveOrder trees.
//
// Format reminder:
//   { "w:p": [child1, child2, ...], ":@": { "@_w:val": "..." } }
// Each element has exactly one tag-name key (its tag) plus an optional ":@"
// attributes object. Children include leaf text nodes of shape
//   { "#text": "..." }.

/**
 * One element node in a fast-xml-parser `preserveOrder` tree: a record with a
 * single tag-name key (its children) plus an optional `:@` attributes object.
 */
export type PoNode = Record<string, unknown>;

/** An ordered list of sibling {@link PoNode}s (the parser's array form). */
export type PoTree = ReadonlyArray<PoNode>;

const ATTRS_KEY = ':@';
const TEXT_KEY = '#text';

/** The element's tag name (the one key that is neither `:@` nor `#text`), or undefined. */
export function poTag(node: PoNode | undefined): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  for (const key of Object.keys(node)) {
    if (key !== ATTRS_KEY && key !== TEXT_KEY) return key;
  }
  return undefined;
}

/** Whether the node's tag equals `tag` exactly (prefix-sensitive). */
export function poIs(node: PoNode | undefined, tag: string): boolean {
  return poTag(node) === tag;
}

/** The node's child elements (the value under its tag key), or `[]`. */
export function poChildren(node: PoNode | undefined): ReadonlyArray<PoNode> {
  const tag = poTag(node);
  if (!tag || !node) return [];
  const arr = node[tag];
  return Array.isArray(arr) ? (arr as Array<PoNode>) : [];
}

/** The node's direct children with tag `tag`. */
export function poChildrenWith(node: PoNode | undefined, tag: string): Array<PoNode> {
  return poChildren(node).filter((c) => poIs(c, tag));
}

/** The node's first direct child with tag `tag`, or undefined. */
export function poFirstChild(node: PoNode | undefined, tag: string): PoNode | undefined {
  return poChildren(node).find((c) => poIs(c, tag));
}

/**
 * An attribute value by local `name`, trying the common OOXML namespace prefixes
 * (`w:`/`r:`/`m:`/`xml:`, then bare) in priority order. Returns undefined when
 * absent or non-string.
 */
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

/**
 * Like `poAttr` but matches by LOCAL name regardless of the prefix — for vendor
 * extensions whose namespace `poAttr` does not special-case (e.g. `w14:paraId`,
 * `w15:done`, `w15:paraIdParent`).
 */
export function poAttrLocal(node: PoNode | undefined, name: string): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const attrs = node[ATTRS_KEY];
  if (!attrs || typeof attrs !== 'object') return undefined;
  for (const [k, v] of Object.entries(attrs as Record<string, unknown>)) {
    if (!k.startsWith('@_')) continue;
    const local = k.slice(2).split(':').pop();
    if (local === name && typeof v === 'string') return v;
  }
  return undefined;
}

/**
 * Like `poIs` but matches a tag by LOCAL name, so an element authored under any
 * prefix (`w15:commentsEx`, `mc:commentsEx`, …) still resolves.
 */
export function poIsLocal(node: PoNode | undefined, name: string): boolean {
  const tag = poTag(node);
  return tag !== undefined && (tag === name || tag.split(':').pop() === name);
}

/** {@link poAttr} parsed as a finite number, or undefined when absent/non-numeric. */
export function poIntAttr(node: PoNode | undefined, name: string): number | undefined {
  const v = poAttr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** Shorthand for the `val` attribute (`@w:val`/…). */
export function poVal(node: PoNode | undefined): string | undefined {
  return poAttr(node, 'val');
}

/**
 * Read an OOXML boolean toggle property (§17.17.4): a present element with no
 * `val` (or `val` true/1/on) is `true`; `false`/`0`/`off` is `false`; an absent
 * element (`node` undefined) is undefined ("inherit").
 */
export function poToggle(node: PoNode | undefined): boolean | undefined {
  if (!node) return undefined;
  const v = poAttr(node, 'val');
  if (v === undefined) return true;
  return !(v === 'false' || v === '0' || v === 'off');
}

/** Concatenate the node's direct `#text` leaf children into a single string. */
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

/**
 * Walk a chain of tags from the tree root, descending into the first child that
 * matches each successive `path` segment. Returns the node at the path end, or
 * undefined if any segment is missing.
 */
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

/** Depth-first search for the first descendant with the given `tag`. */
export function poFindDescendant(node: PoNode | undefined, tag: string): PoNode | undefined {
  if (!node) return undefined;
  for (const child of poChildren(node)) {
    if (poIs(child, tag)) return child;
    const found = poFindDescendant(child, tag);
    if (found) return found;
  }
  return undefined;
}
