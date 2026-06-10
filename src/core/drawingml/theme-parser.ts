// ECMA-376 §20.1.6.2 — theme colour scheme (word/theme/theme1.xml a:clrScheme).
//
// Reads the twelve scheme slots (dk1/lt1/dk2/lt2, accent1-6, hlink/folHlink)
// into a name→hex map. Each slot holds either an a:srgbClr (@val) or an
// a:sysClr (@lastClr — the resolved system colour). The converter merges this
// over the built-in default palette and builds a ColorResolver from it.

import { XMLParser } from 'fast-xml-parser';

import type { PoNode } from '@/word/po-helpers';
import { poAttr, poChildren, poFindByPath, poIs, poTag } from '@/word/po-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

const SCHEME_SLOTS = new Set([
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]);

export function parseTheme(themeXml: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  const tree = parser.parse(decoder.decode(themeXml)) as Array<PoNode>;
  const clrScheme = poFindByPath(tree, ['a:theme', 'a:themeElements', 'a:clrScheme']);
  if (!clrScheme) return out;
  for (const slot of poChildren(clrScheme)) {
    const tag = poTag(slot); // 'a:accent1' etc.
    if (!tag || !tag.startsWith('a:')) continue;
    const name = tag.slice(2);
    if (!SCHEME_SLOTS.has(name)) continue;
    const hex = colorOf(slot);
    if (hex) out.set(name, hex);
  }
  return out;
}

function colorOf(slot: PoNode): string | undefined {
  for (const c of poChildren(slot)) {
    if (poIs(c, 'a:srgbClr')) {
      const v = poAttr(c, 'val');
      if (v) return v.toUpperCase();
    } else if (poIs(c, 'a:sysClr')) {
      const v = poAttr(c, 'lastClr');
      if (v) return v.toUpperCase();
    }
  }
  return undefined;
}
