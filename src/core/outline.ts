// Heading detection — the one rule every target reads a document's outline by.
//
// ECMA-376 §17.3.1.20 `w:outlineLvl` states levels 0–8, which is nine; the
// media that consume them offer six (HTML's h1–h6, markdown's `#`×6, tagged
// PDF's H1–H6 in ISO 32000-1 §14.8.4), so 0–5 map straight across and 6–8
// clamp onto the sixth.
//
// A heading style need not carry an outline level at all: Word's own
// "Heading 2" leaves it to the style, so the style id is the fallback, and
// "Title"/"Subtitle" — which no outline level ever covers — read as level 1.

import type { ResolvedParagraphProperties } from '@/core/style-cascade';

/** The deepest heading level the consuming media can express. */
const MAX_HEADING_LEVEL = 6;

/**
 * The heading level (1–6) a paragraph's resolved properties describe, or
 * `undefined` when it is body text.
 *
 * @param resolved The paragraph's fully-resolved properties.
 * @returns The heading level 1–6, or `undefined` for body text.
 */
export function headingLevelOf(resolved: ResolvedParagraphProperties): number | undefined {
  const lvl = resolved.outlineLevel;
  if (lvl !== undefined && lvl >= 0 && lvl <= 8) {
    return Math.min(lvl + 1, MAX_HEADING_LEVEL);
  }
  return headingLevelFromStyleId(resolved.styleId);
}

/**
 * The heading level a `"Heading N"` / `"Title"` / `"Subtitle"` style id names,
 * for a heading style that declares no outline level of its own.
 *
 * @param styleId The paragraph's style id, if it has one.
 * @returns The heading level 1–6, or `undefined` when the id names no heading.
 */
export function headingLevelFromStyleId(styleId: string | undefined): number | undefined {
  if (styleId === undefined) return undefined;
  const m = /^Heading\s*([1-9])$/i.exec(styleId);
  if (m) return Math.min(Number(m[1]), MAX_HEADING_LEVEL);
  if (/^(Title|Subtitle)$/i.test(styleId)) return 1;
  return undefined;
}
