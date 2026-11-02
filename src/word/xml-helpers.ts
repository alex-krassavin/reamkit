// Helpers for reading the tree produced by fast-xml-parser with
// attributeNamePrefix '@_' and textNodeName '#text'.

import { universalMeasureTwips } from '@/core/po-helpers';
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
/**
 * §17.3.2.38 ST_HpsMeasure — a size in HALF-POINTS, or the same universal
 * measure with a unit on it ("20pt"). The two units differ by a factor of ten
 * from the twips a length is written in, so a size is read here rather than
 * through {@link parseIntAttr}: tdf108408.docx writes `w:sz w:val="20pt"` and
 * read as twips it set its sample text 200 points tall.
 *
 * @param node The element carrying the attribute.
 * @param name The attribute name.
 * @returns The size in half-points, or undefined.
 */
export function parseHalfPointAttr(node: unknown, name: string): number | undefined {
  const v = getAttr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  if (Number.isFinite(n)) return n;
  const twips = universalMeasureTwips(v);
  return twips === undefined ? undefined : twips / 10;
}

export function parseIntAttr(node: unknown, name: string): number | undefined {
  const v = getAttr(node, name);
  if (v === undefined) return undefined;
  const n = Number(v);
  if (Number.isFinite(n)) return n;
  // §22.9.2.15 — the value may name its UNIT instead of being a count of
  // twips, and producers do: tdf116410.docx writes its tab stops, indents and
  // spacing as "85.05pt" and every one of them was read as nothing, so the
  // paragraph lost the stop its dot leader hangs on.
  return universalMeasureTwips(v);
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
 * Narrow a `w:rPr` or `w:pPr` to an element, with every property written more
 * than once in it folded into the one Word applies.
 *
 * The schema allows each property once, but producers repeat them: LibreOffice
 * writes `<w:b w:val="false"/>` twice, and the flat tree hands the pair over as
 * an array, which every reader took for nothing and {@link parseToggle} for
 * "on", so the run came out bold. Word applies the occurrences in order, each
 * over the ones before it (Word for Mac, measured): what a later one states
 * wins, what it leaves unsaid stands. Two `w:ind` keep the left indent of one
 * and the right of the other, two `w:pBdr` draw the rules of both, and two
 * `w:rFonts` name the faces of both.
 *
 * @param node The parsed `w:rPr` / `w:pPr`.
 * @returns The element with its repeats folded (the same object when it has
 *          none), or `undefined` when `node` is not a single element.
 */
export function foldRepeatedProperties(node: unknown): XmlElement | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  let out: XmlElement | undefined;
  for (const [tag, v] of Object.entries(el)) {
    if (!Array.isArray(v)) continue;
    out ??= { ...el };
    out[tag] = fold(v);
  }
  return out ?? el;
}

function fold(v: unknown): unknown {
  return Array.isArray(v) ? (v as ReadonlyArray<unknown>).reduce(applyOver) : v;
}

/**
 * One occurrence of a property element applied over an earlier one: its
 * attributes win, its child elements apply over theirs in turn, and its tab
 * stops join the ones already set (Word keeps the stops of both `w:tabs`).
 */
function applyOver(earlier: unknown, later: unknown): unknown {
  const a = asElement(earlier) ?? {};
  const b = asElement(later) ?? {};
  // Between two that carry nothing but `w:val`, the later one is the whole
  // statement: an on/off property written bare says "on" (§17.17.4), and
  // `<w:b w:val="0"/><w:b/>` is bold in Word, where laying the bare one over
  // the other attribute by attribute would have kept the "0".
  if (valOnly(a) && valOnly(b)) return later;
  const out: XmlElement = { ...a };
  for (const [k, v] of Object.entries(b)) {
    out[k] =
      !Object.hasOwn(out, k) || k.startsWith('@_')
        ? v
        : k === 'w:tab'
          ? [...asArray(out[k]), ...asArray(v)]
          : applyOver(fold(out[k]), fold(v));
  }
  return out;
}

function valOnly(el: XmlElement): boolean {
  return Object.keys(el).every((k) => k === '@_w:val' || k === '@_val');
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
