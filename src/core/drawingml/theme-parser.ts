// ECMA-376 §20.1.6.2 — theme colour scheme (word/theme/theme1.xml a:clrScheme).
//
// Reads the twelve scheme slots (dk1/lt1/dk2/lt2, accent1-6, hlink/folHlink)
// into a name→hex map. Each slot holds either an a:srgbClr (@val) or an
// a:sysClr (@lastClr — the resolved system colour). The converter merges this
// over the built-in default palette and builds a ColorResolver from it.

import { XMLParser } from 'fast-xml-parser';

import type { PoNode } from '@/core/po-helpers';
import { poAttr, poChildren, poFindByPath, poIs, poTag } from '@/core/po-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
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

/**
 * Parse a DrawingML theme part (ECMA-376 §20.1.6.2, `a:clrScheme`) into a
 * `name → hex` colour map. Reads the twelve scheme slots — `dk1`/`lt1`/`dk2`/
 * `lt2`, `accent1`–`accent6`, `hlink`/`folHlink` — taking each slot's
 * `a:srgbClr@val` or, for a system colour, its resolved `a:sysClr@lastClr`.
 * Unknown slots and slots with no resolvable colour are skipped.
 *
 * @param themeXml The raw `word/theme/theme1.xml` (or sibling) bytes, UTF-8.
 * @returns A map keyed by slot name (`'accent1'`, …) to uppercase RRGGBB hex;
 *          empty when no `a:clrScheme` is present.
 */
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

/**
 * Parse a theme's line-style widths (ECMA-376 §20.1.4.1.21 `a:lnStyleLst`).
 *
 * A shape drawn from the gallery keeps its outline as `<a:lnRef idx="N">`,
 * which is a 1-based index into this list — the reference names the colour and
 * the list holds the width. The standard Office theme is 0.75pt / 2pt / 3pt.
 *
 * @param themeXml The raw theme part bytes, UTF-8.
 * @returns The widths in points, in list order; empty when the theme declares
 *          no `a:fmtScheme`.
 */
export function parseThemeLineWidths(themeXml: Uint8Array): Array<number> {
  const tree = parser.parse(decoder.decode(themeXml)) as Array<PoNode>;
  const list = poFindByPath(tree, ['a:theme', 'a:themeElements', 'a:fmtScheme', 'a:lnStyleLst']);
  if (!list) return [];
  const out: Array<number> = [];
  for (const ln of poChildren(list)) {
    if (!poIs(ln, 'a:ln')) continue;
    const w = Number(poAttr(ln, 'w') ?? '');
    // §20.1.2.1 ST_LineWidth is EMU; a width the theme omits is a hairline.
    out.push(Number.isFinite(w) && w > 0 ? w / EMU_PER_POINT : 0.75);
  }
  return out;
}

const EMU_PER_POINT = 12700;

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
