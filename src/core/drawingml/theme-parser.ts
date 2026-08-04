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
 * §20.1.4.1.16 `a:fontScheme` — the two typefaces a theme names, each in three
 * scripts. A document does not repeat them: it refers to them by TOKEN
 * (`+mj-lt`, `+mn-ea`, …), which every reader is expected to resolve.
 */
export interface ThemeFonts {
  /** `a:majorFont` — headings. */
  readonly major: ThemeFontSlots;
  /** `a:minorFont` — body text. */
  readonly minor: ThemeFontSlots;
}

/** One font of a scheme, in the three scripts a run may pick from. */
export interface ThemeFontSlots {
  readonly latin?: string;
  readonly ea?: string;
  readonly cs?: string;
}

/**
 * Parse a theme's font scheme (§20.1.4.1.16).
 *
 * @param themeXml The raw theme part bytes, UTF-8.
 * @returns The two fonts; slots the theme leaves empty are absent.
 */
export function parseThemeFonts(themeXml: Uint8Array): ThemeFonts {
  const tree = parser.parse(decoder.decode(themeXml)) as Array<PoNode>;
  const scheme = poFindByPath(tree, ['a:theme', 'a:themeElements', 'a:fontScheme']);
  const read = (tag: string): ThemeFontSlots => {
    const font = scheme ? poChildren(scheme).find((c) => poIs(c, tag)) : undefined;
    const slot = (name: string): string | undefined => {
      const node = font ? poChildren(font).find((c) => poIs(c, name)) : undefined;
      const face = node ? poAttr(node, 'typeface')?.trim() : undefined;
      return face === undefined || face === '' ? undefined : face;
    };
    const [latin, ea, cs] = [slot('a:latin'), slot('a:ea'), slot('a:cs')];
    return {
      ...(latin !== undefined ? { latin } : {}),
      ...(ea !== undefined ? { ea } : {}),
      ...(cs !== undefined ? { cs } : {}),
    };
  };
  return { major: read('a:majorFont'), minor: read('a:minorFont') };
}

// §20.1.4.1.14 — the six tokens a typeface may be written as, and the slot each
// one stands for.
const THEME_FONT_TOKENS: ReadonlyMap<string, readonly ['major' | 'minor', keyof ThemeFontSlots]> =
  new Map([
    ['+mj-lt', ['major', 'latin']],
    ['+mn-lt', ['minor', 'latin']],
    ['+mj-ea', ['major', 'ea']],
    ['+mn-ea', ['minor', 'ea']],
    ['+mj-cs', ['major', 'cs']],
    ['+mn-cs', ['minor', 'cs']],
  ]);

/**
 * Resolve a typeface written as a theme TOKEN to the name it stands for.
 *
 * A slide states its fonts as `+mn-lt` far more often than by name, and left
 * unresolved that string travels into the model as if it WERE a typeface: no
 * substitution table knows it, so a deck whose theme is Times came out in a
 * grotesque (45541_Header).
 *
 * @param typeface The `@typeface` as the file writes it.
 * @param fonts    The theme's font scheme, when the part was read.
 * @returns The resolved name; the input unchanged when it is not a token, and
 *          `undefined` when it is a token the theme leaves empty.
 */
export function resolveThemeFont(
  typeface: string | undefined,
  fonts: ThemeFonts | undefined,
): string | undefined {
  if (typeface === undefined || !typeface.startsWith('+')) return typeface;
  const slot = THEME_FONT_TOKENS.get(typeface.toLowerCase());
  if (!slot || !fonts) return undefined;
  const [which, script] = slot;
  return fonts[which][script];
}

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

/**
 * Parse a theme's fill styles (§20.1.4.1.13 `a:fillStyleLst`), as the raw nodes.
 *
 * A shape drawn from the gallery carries no fill of its own — only
 * `<a:fillRef idx="N">`, a 1-based index into this list, and a colour to put
 * where the styles say `phClr`. The standard Office theme's slots are a solid,
 * a subtle gradient and a stronger one, so reading the reference's colour alone
 * paints slot 3 flat: 47504.xlsx's rectangle is a gradient in both references
 * and a single blue in ours.
 *
 * The nodes are handed back unparsed because what they hold is a whole fill —
 * solid, gradient, pattern — which the shape readers already know how to read.
 *
 * @param themeXml The raw theme part bytes, UTF-8.
 * @returns The `a:fillStyleLst` children in list order; empty when the theme
 *          declares no `a:fmtScheme`.
 */
export function parseThemeFillStyles(themeXml: Uint8Array): Array<PoNode> {
  const tree = parser.parse(decoder.decode(themeXml)) as Array<PoNode>;
  const list = poFindByPath(tree, ['a:theme', 'a:themeElements', 'a:fmtScheme', 'a:fillStyleLst']);
  return list ? [...poChildren(list)] : [];
}

/**
 * Parse a theme's BACKGROUND fill styles (§20.1.4.1.7 `a:bgFillStyleLst`), as
 * the raw nodes. An `<a:fillRef>` reaches them with an index past 1000 — slot
 * 1001 is the first — and that is where the page-sized backdrops Word's cover
 * pages are built from live (fdo78957.docx).
 *
 * @param themeXml The raw theme part bytes, UTF-8.
 * @returns The fill children in list order; empty when the theme declares none.
 */
export function parseThemeBgFillStyles(themeXml: Uint8Array): Array<PoNode> {
  const tree = parser.parse(decoder.decode(themeXml)) as Array<PoNode>;
  const list = poFindByPath(tree, [
    'a:theme',
    'a:themeElements',
    'a:fmtScheme',
    'a:bgFillStyleLst',
  ]);
  return list ? [...poChildren(list)] : [];
}

/**
 * Parse a theme's effect styles (§20.1.4.1.15 `a:effectStyleLst`), as the raw
 * nodes. `<a:effectRef idx="N">` is a 1-based index into this list, exactly as
 * the fill and line references index theirs.
 *
 * @param themeXml The raw theme part bytes, UTF-8.
 * @returns The `a:effectStyle` children in list order; empty when the theme
 *          declares no `a:fmtScheme`.
 */
export function parseThemeEffectStyles(themeXml: Uint8Array): Array<PoNode> {
  const tree = parser.parse(decoder.decode(themeXml)) as Array<PoNode>;
  const list = poFindByPath(tree, [
    'a:theme',
    'a:themeElements',
    'a:fmtScheme',
    'a:effectStyleLst',
  ]);
  return list ? [...poChildren(list)] : [];
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
