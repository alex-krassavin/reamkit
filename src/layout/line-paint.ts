// What drawing a laid-out page asks of every writer that draws one — the PDF
// emitter and the SVG writer alike: the order a line's tokens stand in, how
// far each of its spaces stretches, whether one matrix sets the whole of it,
// the lean of a faked italic and the pattern of a dashed rule. One owner, so
// the two pages cannot draw the same layout two ways.

import type { BorderStyle } from '@/core/document-model';
import type { Line } from '@/layout/page-doc';
import { reorderVisual } from '@/core/bidi';
import { GLUE_SHRINK_RATIO } from '@/layout/styled-layout';

/** The slant of a faked italic — tan(12°), the angle a text italic leans at. */
export const FAUX_ITALIC_SHEAR = 0.2126;

/**
 * §18.18.3 — the dash pattern of a rule that is not solid, in points: dashes
 * and gaps alternating, from a dash.
 */
export const BORDER_DASHES: ReadonlyMap<BorderStyle, ReadonlyArray<number>> = new Map([
  ['dashed', [8, 2.5]],
  // The SMALL-gap dash is a pattern of its own — Word spells it
  // `dashSmallGap`, Excel calls the same thing a thin `dashed` rule.
  ['dashSmallGap', [3, 1]],
  ['dotted', [0.5, 1]],
  // …and a dash-dot alternates the two, which is the whole difference
  // between it and a dash on the page (cell-borders.xlsx names five of them).
  ['dashDot', [8, 2.5, 2.5, 2.5]],
  ['dashDotDot', [8, 2.5, 2.5, 2.5, 2.5, 2.5]],
]);

/**
 * UAX #9 rule L2 over a line's tokens: the token indices in visual order. Each
 * token carries a single embedding level (the tokenizer split runs at level
 * boundaries), so reordering tokens is equivalent to reordering their
 * characters.
 *
 * @param line The laid-out line.
 * @returns Its token indices, left to right as they are drawn.
 */
export function lineVisualOrder(line: Line): Array<number> {
  return reorderVisual(line.tokens.map((t) => t.bidiLevel));
}

/**
 * The extra width each space of a justified line takes — 0 for a line that is
 * not justified, or the last line of a paragraph (which stays flush by
 * convention).
 *
 * @param line The laid-out line.
 * @returns Points added to (or, squeezed, taken from) each space token.
 */
export function computeJustifyExtra(line: Line): number {
  const alignment = line.resolved.alignment;
  if (alignment !== 'both' && alignment !== 'distribute') return 0;
  // §17.3.1.13 — `distribute` justifies EVERY line, the last one included;
  // `both` leaves the last alone. para-adjust-distribute.docx sets one of each
  // and we wrote both flush left.
  if (line.noJustify === true) return 0;
  if (line.isLastInParagraph && alignment !== 'distribute') return 0;
  // A space at the END of a line justifies nothing: there is no glyph after it
  // to push, and counted in it left the line short of the measure by its own
  // width. Word and LibreOffice both hang it past the margin instead.
  let last = line.tokens.length - 1;
  while (last >= 0 && line.tokens[last]!.kind === 'text' && line.tokens[last]!.isSpace) last--;
  let spaces = 0;
  let narrowest = Infinity;
  let contentWidthPt = 0;
  for (let i = 0; i <= last; i++) {
    const tok = line.tokens[i]!;
    contentWidthPt += tok.widthPt;
    if (!tok.isSpace) continue;
    spaces++;
    if (tok.widthPt < narrowest) narrowest = tok.widthPt;
  }
  if (spaces === 0) return 0;
  // Both ways. A line's spaces stretch to fill it out, and they SHRINK to pull
  // it in — the breaker weighs a line knowing it may squeeze each space by up
  // to {@link GLUE_SHRINK_RATIO}, so a line it packed that tight is one whose
  // natural width is over the measure. Drawn at that natural width it ran past
  // the right margin: IllustrativeCases.docx put three of its four opening
  // lines 1 to 10pt into the margin while the fourth sat short of it, which
  // reads as no justification at all.
  const extra = (line.availableWidthPt - contentWidthPt) / spaces;
  return extra < 0 ? Math.max(extra, -narrowest * GLUE_SHRINK_RATIO) : extra;
}

/**
 * Whether a line is ONE text matrix and a run of glyphs: nothing in it placed
 * on its own — a tab, a stretched space, a picture, a formula, a right-to-left
 * run, a rise, a faked italic — and nothing drawn beside its glyphs, a run's
 * shading, a highlight or a rule. Only such a line can be turned by its matrix
 * alone; any other is placed token by token in page space.
 *
 * @param line The laid-out line.
 * @returns True when a single text matrix sets the whole of it.
 */
export function settableInOneMatrix(line: Line): boolean {
  return (
    computeJustifyExtra(line) === 0 &&
    line.tokens.every(
      (t) =>
        t.kind === 'text' &&
        t.tab !== true &&
        t.bidiLevel % 2 === 0 &&
        t.risePt === undefined &&
        t.synthetic?.italic !== true &&
        t.highlight !== true &&
        t.resolvedRun.shadingColorHex === undefined &&
        t.resolvedRun.underline === 'none' &&
        !t.resolvedRun.strike,
    )
  );
}

/**
 * §14.1.2.10 — the wash a picture is drawn through, as the flat veil it is.
 * `gain`/`blacklevel` are contrast and brightness about mid grey, so a source
 * value `in` prints as `(in - 0.5) * gain + 0.5 + black` — and painting a
 * colour `c` at opacity `a` gives `in * (1 - a) + c * a`, the same line with
 * `a = 1 - gain`. Exact for every wash a document can state, and it needs no
 * decoder: the picture's own pixels are not rewritten.
 *
 * @param wash The picture's wash, if it has one.
 * @returns The grey (0…1) and opacity of the veil laid over the picture, or
 *          nothing where the wash changes nothing.
 */
export function washVeil(
  wash: { readonly gain: number; readonly black: number } | undefined,
): { alpha: number; grey: number } | undefined {
  if (!wash) return undefined;
  const alpha = 1 - wash.gain;
  if (!(alpha > 0.004)) return undefined;
  const offset = 0.5 * (1 - wash.gain) + wash.black;
  return { alpha: Math.min(1, alpha), grey: Math.max(0, Math.min(1, offset / alpha)) };
}
