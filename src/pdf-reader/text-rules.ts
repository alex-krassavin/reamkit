// A PDF has no underline.
//
// Nothing in ISO 32000-1 says a run of text is underlined or struck through:
// the file draws a thin filled rectangle under the words, or across them, and
// the reader is expected to see a line. annotation-link-text-popup.pdf paints
// `70.86 736.14 49.34 0.72 re f*` beneath a baseline at 738.12 and that is its
// hyperlink's underline, in the page content, indistinguishable from a rule.
//
// Lifted as artwork the rule is anchored where it was drawn, which is right
// until the text re-sets — and then it is a blue dash floating between two
// lines. Recognised for what it is, it goes on the RUN (§17.3.2.40 `w:u`,
// §17.3.2.37 `w:strike`) and travels with the words wherever they end up.
//
// The signature is narrow on purpose: a table rule and an underline are the
// same shape, and only their placement tells them apart.

import type { TextRun } from './content';
import type { PdfVector } from './vector';

/** A rule no thicker than this is a line under words, not a box. */
const THICKEST_PT = 2;

/** …and no thinner than this: below it the page drew a seam, not a mark. */
const THINNEST_PT = 0.2;

/** …or than this fraction of the face it underlines, for a large one. */
const THICKEST_EM = 0.09;

/** How much of a run's advance a rule must cover to be that run's own. */
const COVERS = 0.6;

/**
 * How far past the words it may run, in points, before it is a rule about
 * something else.
 *
 * An underline begins and ends with the words it underlines; a table's cell
 * border begins at the CELL, which is the text's own edge less the padding, and
 * ends at the cell's other edge however short the text stops. TAMReview.pdf's
 * tables are ruled 86.2 to 509.0 across a measure whose text stops well before
 * it, and a proportional allowance let every one of them through.
 */
const OVERHANG_PT = 2;

/** An underline sits within this fraction of the size below the baseline. */
const UNDER_EM = 0.35;

/** A strikeout crosses between these fractions of the size ABOVE it. */
const STRIKE_LOW = 0.15;
const STRIKE_HIGH = 0.5;

/** The runs with their drawn rules read onto them, and the rules so read. */
export interface DrawnRules {
  readonly runs: Array<TextRun>;
  /** The vectors the runs took over: painting them again would double them. */
  readonly consumed: ReadonlySet<PdfVector>;
}

/**
 * Read the underlines and strikeouts a page DREW onto the runs they mark.
 *
 * @param runs    The page's runs, placed on the shown page.
 * @param vectors The page's lifted paths, placed the same way.
 * @returns The runs, marked; and the paths that became those marks.
 */
export function markDrawnRules(
  runs: ReadonlyArray<TextRun>,
  vectors: ReadonlyArray<PdfVector>,
): DrawnRules {
  const marks = new Map<TextRun, { underline?: string; strike?: boolean }>();
  const consumed = new Set<PdfVector>();
  for (const v of vectors) {
    if (!isRule(v)) continue;
    const under = coveredRuns(runs, v, 'under');
    const through = under.length > 0 ? [] : coveredRuns(runs, v, 'through');
    const hit = under.length > 0 ? under : through;
    if (hit.length === 0) continue;
    // A rule that runs well past the words it would mark is a rule about the
    // page, not about them: a table's, a header's, a footer's.
    const left = Math.min(...hit.map((r) => Math.min(r.x, r.endX)));
    const right = Math.max(...hit.map((r) => Math.max(r.x, r.endX)));
    const span = right - left;
    if (!(span > 0)) continue;
    if (left - v.minX > OVERHANG_PT || v.maxX - right > OVERHANG_PT) continue;
    for (const run of hit) {
      const had = marks.get(run) ?? {};
      marks.set(
        run,
        under.length > 0 && v.fillHex !== undefined
          ? { ...had, underline: v.fillHex }
          : { ...had, strike: true },
      );
    }
    consumed.add(v);
  }
  if (marks.size === 0) return { runs: [...runs], consumed };
  return {
    runs: runs.map((run) => {
      const mark = marks.get(run);
      if (!mark) return run;
      return {
        ...run,
        markup: {
          ...run.markup,
          // An annotation that says the same thing said it first.
          ...(mark.underline !== undefined && run.markup?.underline === undefined
            ? { underline: 'single' as const, underlineHex: mark.underline }
            : {}),
          ...(mark.strike === true ? { strike: true } : {}),
        },
      };
    }),
    consumed,
  };
}

/** A thin filled bar, which is the only shape an underline is drawn as. */
function isRule(v: PdfVector): boolean {
  if (v.fillHex === undefined || v.strokeHex !== undefined || v.gradient !== undefined) {
    return false;
  }
  // Nothing is underlined in white: white over white paper is not a mark, and
  // where a page paints a white seam between two table cells it is hiding a
  // join, not marking the words above it.
  if (v.fillHex === 'FFFFFF') return false;
  const h = v.maxY - v.minY;
  const w = v.maxX - v.minX;
  return h >= THINNEST_PT && h <= THICKEST_PT && w > 2;
}

/**
 * The runs a rule marks: on one baseline, covered along their advance, and at
 * the offset from that baseline the mark is drawn at.
 */
function coveredRuns(
  runs: ReadonlyArray<TextRun>,
  v: PdfVector,
  where: 'under' | 'through',
): Array<TextRun> {
  const out: Array<TextRun> = [];
  const mid = (v.minY + v.maxY) / 2;
  for (const run of runs) {
    const size = run.fontSizePt;
    if (!(size > 0) || run.angleDeg !== undefined) continue;
    if (v.maxY - v.minY > Math.max(THICKEST_PT, size * THICKEST_EM)) continue;
    const below = run.y - mid;
    const fits =
      where === 'under'
        ? below > 0 && below < size * UNDER_EM
        : below < -size * STRIKE_LOW && below > -size * STRIKE_HIGH;
    if (!fits) continue;
    const left = Math.min(run.x, run.endX);
    const right = Math.max(run.x, run.endX);
    const advance = right - left;
    if (!(advance > 0)) continue;
    const shared = Math.min(right, v.maxX) - Math.max(left, v.minX);
    if (shared < advance * COVERS) continue;
    out.push(run);
  }
  return out;
}
