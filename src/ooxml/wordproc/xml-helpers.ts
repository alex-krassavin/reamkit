// Helpers for reading the tree produced by fast-xml-parser with
// attributeNamePrefix '@_' and textNodeName '#text'.

export type XmlAttrValue = string;
export type XmlElement = Record<string, unknown>;

export function asElement(v: unknown): XmlElement | undefined {
  if (v === null || v === undefined) return undefined;
  if (typeof v !== 'object' || Array.isArray(v)) return undefined;
  return v as XmlElement;
}

export function asArray(v: unknown): ReadonlyArray<unknown> {
  if (v === null || v === undefined) return [];
  return Array.isArray(v) ? v : [v];
}

export function getAttr(node: unknown, name: string): string | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  const v = el[`@_w:${name}`] ?? el[`@_${name}`];
  return typeof v === 'string' ? v : undefined;
}

export function getVal(node: unknown): string | undefined {
  return getAttr(node, 'val');
}

export function parseIntAttr(node: unknown, name: string): number | undefined {
  const v = getAttr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  if (!Number.isFinite(n)) return undefined;
  return n;
}

// ECMA-376 Part 1 §17.17.4 — Boolean Properties.
// Absent  → undefined (inherit)
// Present without val, or val ∈ {true, 1, on}  → true
// Present with val ∈ {false, 0, off}           → false
export function parseToggle(node: unknown): boolean | undefined {
  if (node === undefined) return undefined;
  if (node === null || node === '' || typeof node === 'boolean' || typeof node === 'number') {
    return true;
  }
  if (typeof node === 'string') return true;
  const v = getAttr(node, 'val');
  if (v === undefined) return true;
  return !(v === 'false' || v === '0' || v === 'off');
}

export function getText(node: unknown): string {
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  const el = asElement(node);
  if (!el) return '';
  const inner = el['#text'];
  if (typeof inner === 'string') return inner;
  if (typeof inner === 'number') return String(inner);
  return '';
}
