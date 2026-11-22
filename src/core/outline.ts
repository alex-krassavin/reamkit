// Heading detection — the one rule every target reads a document's outline by.
//
// ECMA-376 §17.3.1.20 `w:outlineLvl` states levels 0–8, which is nine; the
// media that consume them offer six (HTML's h1–h6, markdown's `#`×6, tagged
// PDF's H1–H6 in ISO 32000-1 §14.8.4), so 0–5 map straight across and 6–8
// clamp onto the sixth.
//
// A heading style need not carry an outline level at all. Word's own
// "heading N" styles are known by their NAME and the styles parser gives them
// their level (`Style.builtInOutlineLevel`); a document built without a style
// sheet — a FlowDoc made in code — has only the id, so the id is the fallback,
// and "Title"/"Subtitle" — which no outline level ever covers — read as level 1.

import type { Run } from '@/core/document-model';
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
  // §17.3.1.20 — level 9 is "body text", stated: a paragraph that says it is
  // no heading is not one, whatever its style id looks like.
  if (lvl === 9) return undefined;
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

/**
 * The level (1–9) at which a paragraph enters a document's NAVIGATION outline
 * — a PDF's bookmarks (ISO 32000-1 §12.3.3) — or `undefined` for one that
 * does not.
 *
 * All nine outline levels, unclamped: a bookmarks panel nests as deep as the
 * document does, and LibreOffice's export puts Heading 7–9 at levels 7–9.
 * "Title" and "Subtitle" stay out — no outline level covers them, Word reads
 * them as body text, and neither Word nor LibreOffice bookmarks them.
 *
 * @param resolved The paragraph's fully-resolved properties.
 * @returns The outline level 1–9, or `undefined` for body text.
 */
export function outlineLevelOf(resolved: ResolvedParagraphProperties): number | undefined {
  const lvl = resolved.outlineLevel;
  if (lvl !== undefined && lvl >= 0 && lvl <= 8) return lvl + 1;
  if (lvl === 9) return undefined;
  const m = resolved.styleId !== undefined ? /^Heading\s*([1-9])$/i.exec(resolved.styleId) : null;
  return m ? Number(m[1]) : undefined;
}

// What a title keeps of a run's text. A soft hyphen is a break opportunity, not
// a character (LibreOffice copies U+00AD into its titles, and viewers print it
// as a hyphen mid-word); a zero-width space and a BOM are no more than that.
// A private-use code is a symbol font's glyph number — a bullet in Wingdings —
// which means nothing in the face a viewer sets the panel in.
const DROPPED_IN_TITLE = /[\u00AD\u200B\uFEFF\uE000-\uF8FF]/gu;
// Any run of white space — the tab after a list number, a line break, an
// ideographic space — reads as one space on the single line a panel shows.
const SPACES_IN_TITLE = /\s+/gu;

/**
 * The text an outline item shows for a paragraph: its words as printed, on one
 * line — the list number included ("1.1 Background", as LibreOffice titles
 * it), the note references and pictures left out.
 *
 * @param runs The paragraph's runs.
 * @returns The title, trimmed; empty when the paragraph says nothing.
 */
export function outlineTitleOf(runs: ReadonlyArray<Run>): string {
  let text = '';
  for (const run of runs) {
    if (run.inlineImage !== undefined || run.math !== undefined) continue;
    if (run.footnoteRef !== undefined || run.endnoteRef !== undefined) continue;
    if (run.commentRef !== undefined || run.noteNumber === true) continue;
    text += run.listMarkerSpace === true ? ' ' : run.text;
  }
  return text.replace(DROPPED_IN_TITLE, '').replace(SPACES_IN_TITLE, ' ').trim();
}
