// Adapter: preserveOrder element → flat object the property parsers expect.
//
// The flat shape is what fast-xml-parser produces *without* preserveOrder:
//   { "@_w:val": "...", "w:b": {…}, "w:r": [{…}, {…}], "#text": "…" }
// Property parsers (run-properties.ts, paragraph-properties.ts) walk that
// shape, so we convert once and reuse them unchanged.

import type { PoNode } from '@/word/po-helpers';
import { poChildren, poTag } from '@/word/po-helpers';

const ATTRS_KEY = ':@';
const TEXT_KEY = '#text';

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
