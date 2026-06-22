// Adapter: preserveOrder element → flat object the property parsers expect.
//
// The flat shape is what fast-xml-parser produces *without* preserveOrder:
//   { "@_w:val": "...", "w:b": {…}, "w:r": [{…}, {…}], "#text": "…" }
// Property parsers (run-properties.ts, paragraph-properties.ts) walk that
// shape, so we convert once and reuse them unchanged.

import type { PoNode } from '@/core/po-helpers';
import { poChildren, poTag } from '@/core/po-helpers';

const ATTRS_KEY = ':@';
const TEXT_KEY = '#text';

/**
 * Convert one `preserveOrder` element ({@link PoNode}) into the flat object shape
 * `fast-xml-parser` produces *without* `preserveOrder`, so the flat-tree property
 * parsers (`run-properties.ts`, `paragraph-properties.ts`) can consume it unchanged.
 *
 * Attributes are flattened to `@_`-prefixed keys, text nodes concatenate into
 * `#text`, and repeated child tags collapse to an array under the tag name.
 *
 * @param node The PO element to flatten; `undefined` yields an empty object.
 * @returns The flat `{ "@_w:val": …, "w:b": {…}, "w:r": [{…}], "#text": "…" }` shape.
 */
export function poElementToFlat(node: PoNode | undefined): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (!node) return out;

  const attrs = node[ATTRS_KEY];
  if (attrs && typeof attrs === 'object') {
    for (const [k, v] of Object.entries(attrs as Record<string, unknown>)) {
      out[k] = v;
    }
  }

  for (const child of poChildren(node)) {
    if (TEXT_KEY in child) {
      const t = (child as Record<string, unknown>)[TEXT_KEY];
      const str = typeof t === 'string' ? t : typeof t === 'number' ? String(t) : '';
      out[TEXT_KEY] =
        (out[TEXT_KEY] as string | undefined) === undefined ? str : `${out[TEXT_KEY]}${str}`;
      continue;
    }
    const childTag = poTag(child);
    if (!childTag) continue;
    const childFlat = poElementToFlat(child);
    const existing = out[childTag];
    if (existing === undefined) {
      out[childTag] = childFlat;
    } else if (Array.isArray(existing)) {
      existing.push(childFlat);
    } else {
      out[childTag] = [existing, childFlat];
    }
  }
  return out;
}
