// Helpers for reading the tree produced by fast-xml-parser with
// attributeNamePrefix '@_' and textNodeName '#text'.

/** The value of a parsed XML attribute (always a string in this flat tree). */
export type XmlAttrValue = string;
/** A parsed XML element: attributes (`@_`-prefixed) and child tags keyed by name. */
export type XmlElement = Record<string, unknown>;

/**
 * Narrow an arbitrary parsed value to an {@link XmlElement}.
 *
 * @returns The value as an element record, or `undefined` when it is null, a
 *          primitive, or an array (i.e. not a single element node).
 */
export function asElement(v: unknown): XmlElement | undefined {
  if (v === null || v === undefined) return undefined;
  if (typeof v !== 'object' || Array.isArray(v)) return undefined;
  return v as XmlElement;
}

/**
 * Coerce a value to an array, treating a lone element as a one-item array.
 * `fast-xml-parser` collapses a single repeated child to a bare object; this
 * normalizes it so callers can always iterate.
 *
 * @returns The value as an array (empty for null/undefined).
 */
export function asArray(v: unknown): ReadonlyArray<unknown> {
  if (v === null || v === undefined) return [];
  return Array.isArray(v) ? v : [v];
}

/**
 * Read a string attribute off an element, trying the `w:`-prefixed name first
 * then the unprefixed fallback (e.g. `@_w:val` then `@_val`).
 *
 * @param node The element to read from.
 * @param name The local attribute name (without the `@_`/`w:` prefix).
 * @returns The attribute value, or `undefined` when the node is not an element
 *          or the attribute is absent / non-string.
 */
export function getAttr(node: unknown, name: string): string | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  const v = el[`@_w:${name}`] ?? el[`@_${name}`];
  return typeof v === 'string' ? v : undefined;
}

/** Shorthand for {@link getAttr}`(node, 'val')` — the ubiquitous `@w:val` attribute. */
export function getVal(node: unknown): string | undefined {
  return getAttr(node, 'val');
}

/**
 * Read an attribute and parse it as a finite number.
 *
 * @param node The element to read from.
 * @param name The local attribute name.
 * @returns The parsed number, or `undefined` when absent or not finite.
 */
export function parseIntAttr(node: unknown, name: string): number | undefined {
  const v = getAttr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  if (!Number.isFinite(n)) return undefined;
  return n;
}

/**
 * Parse a toggle (boolean) property per ECMA-376 Part 1 §17.17.4:
 * - Absent → `undefined` (inherit)
 * - Present without `val`, or `val ∈ {true, 1, on}` → `true`
 * - Present with `val ∈ {false, 0, off}` → `false`
 *
 * @returns The toggle state, or `undefined` when the property is absent.
 */
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

/**
 * Read the `#text` content of a node. A bare string/number is returned directly;
 * an element returns its `#text` child (coerced to string), else the empty string.
 */
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
