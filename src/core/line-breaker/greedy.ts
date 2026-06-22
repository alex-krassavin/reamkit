// Greedy first-fit line breaking (E-PARITY FP3).
//
// Word and LibreOffice break each line at the LAST box/glue/penalty breakpoint
// that still fits the target width, never looking ahead to balance the
// paragraph. This diverges from our Knuth-Plass total-fit (knuth-plass.ts) on
// purpose: under a renderer-compatibility `layoutProfile` we want the same line
// (and therefore page) count those tools produce, even where the typography is
// objectively worse. The default 'ream' profile keeps Knuth-Plass.
//
// The result is the same shape Knuth-Plass returns: indices into items[] where
// each line ends, the last being the paragraph-final forced-break sentinel —
// so the layout consumer (wrap) is identical for both breakers.

import type { Item } from '@/core/line-breaker/knuth-plass';
import { FORBIDDEN_BREAK, FORCED_BREAK } from '@/core/line-breaker/knuth-plass';

/**
 * Break a paragraph into lines greedily (first-fit, E-PARITY FP3): each line
 * ends at the last feasible breakpoint that still fits the target width, never
 * looking ahead to balance the paragraph — matching Word/LibreOffice line (and
 * page) counts under a renderer-compatibility `layoutProfile`. An over-long run
 * with no earlier break gets an emergency break so it cannot merge forward.
 *
 * @param items      The box/glue/penalty stream, sharing the {@link Item} shape
 *                   and feasibility convention of {@link breakLines}.
 * @param lineWidths Target width: a single number (uniform) or a per-line array
 *                   (0-based), reusing the last entry for trailing lines.
 * @returns Indices into `items` where each line ends — the same shape
 *          {@link breakLines} returns, so the layout consumer is identical.
 */
export function greedyBreakLines(
  items: ReadonlyArray<Item>,
  lineWidths: number | ReadonlyArray<number>,
): ReadonlyArray<number> {
  if (items.length === 0) return [];

  const widthFor = (line: number): number =>
    typeof lineWidths === 'number'
      ? lineWidths
      : lineWidths[Math.min(line, lineWidths.length - 1)]!;

  // Cumulative natural width before each item (a penalty contributes 0; it only
  // adds its own width when the break is actually taken there).
  const widthBefore = new Array<number>(items.length + 1);
  widthBefore[0] = 0;
  for (let i = 0; i < items.length; i++) {
    const it = items[i]!;
    widthBefore[i + 1] = widthBefore[i]! + (it.type === 'penalty' ? 0 : it.width);
  }

  // Natural width of the line [start, b): a penalty at b adds its own width (the
  // hyphen char); a glue at b is the breakpoint and is dropped (widthBefore[b]
  // already excludes item b).
  const lineWidth = (start: number, b: number): number => {
    let w = widthBefore[b]! - widthBefore[start]!;
    const item = items[b]!;
    if (item.type === 'penalty') w += item.width;
    return w;
  };

  // Same feasibility convention as Knuth-Plass: a non-forbidden penalty, or a
  // glue immediately following a box.
  const isFeasible = (i: number): boolean => {
    const item = items[i]!;
    if (item.type === 'penalty') return item.penalty < FORBIDDEN_BREAK;
    if (item.type === 'glue') {
      const prev = items[i - 1];
      return !!prev && prev.type === 'box';
    }
    return false;
  };

  const breaks: Array<number> = [];
  let lineStart = 0;
  let lineNo = 0;
  let lastGood = -1; // latest feasible breakpoint that still fit this line

  let i = 0;
  while (i < items.length) {
    if (!isFeasible(i)) {
      i++;
      continue;
    }
    const item = items[i]!;
    const forced = item.type === 'penalty' && item.penalty <= FORCED_BREAK;
    const fits = lineWidth(lineStart, i) <= widthFor(lineNo);

    // The content up to here overflows but an earlier break fit — close the line
    // at that break and re-examine i on the fresh line (also rescues an
    // overflowing forced/last line). Checked before `forced` so a too-wide final
    // word spills onto its own line instead of overrunning the box.
    if (!fits && lastGood >= 0) {
      breaks.push(lastGood);
      lineStart = lastGood + 1;
      lineNo++;
      lastGood = -1;
      continue; // re-process i against the new line (no i++)
    }

    if (forced) {
      breaks.push(i);
      lineStart = i + 1;
      lineNo++;
      lastGood = -1;
      i++;
      continue;
    }

    if (fits) {
      lastGood = i; // remember and keep extending the line
    } else {
      // Overflow with no earlier break: a single over-long run (e.g. a URL in a
      // thin column). Emit an emergency break here so it can't merge forward.
      breaks.push(i);
      lineStart = i + 1;
      lineNo++;
      lastGood = -1;
    }
    i++;
  }
  return breaks;
}
