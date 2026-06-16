// Knuth & Plass "Breaking Paragraphs into Lines" (Software—Practice and
// Experience, 1981). Total-fit dynamic programming over a stream of
// box/glue/penalty items.
//
// We deviate from the canonical paper in two minor ways:
//   - Only one fitness class transition penalty (γ) — matches TeX's defaults.
//   - Shrink ignored when stretch suffices; underset lines (r < -1) are
//     dropped from the active set rather than retried with shrink, since our
//     pipeline never asks for negative stretch.

export interface BoxItem {
  readonly type: 'box';
  readonly width: number;
}

export interface GlueItem {
  readonly type: 'glue';
  readonly width: number; // natural width
  readonly stretch: number; // how far it can expand (positive)
  readonly shrink: number; // how far it can contract (positive)
}

export interface PenaltyItem {
  readonly type: 'penalty';
  readonly width: number; // width if break taken here (e.g. hyphen char)
  readonly penalty: number; // negative = forced, +Infinity = forbidden
  readonly flagged: boolean; // true for hyphenation points — paired penalty avoided
}

export type Item = BoxItem | GlueItem | PenaltyItem;

export interface BreakResult {
  // Indices into items[] where the line breaks happen (last index is also
  // the final forced break / sentinel penalty).
  readonly breaks: ReadonlyArray<number>;
  // Per-line adjustment ratio in [-1, ∞). Used by the layout consumer to
  // distribute extra space across glues on each line.
  readonly ratios: ReadonlyArray<number>;
}

// Tuning knobs — match TeX's defaults closely enough for our needs.
// Exported: writers push the paragraph-final forced break with this value.
export const FORCED_BREAK = -10_000;
// A penalty ≥ this forbids a break at that point. Exported for the greedy
// breaker (greedy.ts), which shares the same feasibility convention.
export const FORBIDDEN_BREAK = 10_000;
const LINE_PENALTY = 10; // α — extra demerits per line
const FITNESS_PENALTY = 100; // γ — adjacent-line fitness mismatch
const FLAGGED_PENALTY = 100; // π — two flagged breakpoints in a row
// Max acceptable adjustment ratio for non-forced breaks. TeX's default is
// roughly equivalent to ratio ~ 1.3 (badness ≤ 200). We're more permissive
// because our paragraphs are typically short and stretch glue is generous.
const TOLERANCE_RATIO = 4;
// Cap on the number of simultaneously active nodes. The DP is O(items × active),
// so an unbounded active list degrades to O(n²) on long paragraphs. Empirically
// 200 is plenty — the optimal-fit path almost always lives near the
// lowest-demerit nodes, and worse nodes only matter when they win later via a
// large penalty discount (rare). When the list overflows, we drop the
// highest-demerit nodes first.
const MAX_ACTIVE_NODES = 200;

const FIT_VERY_TIGHT = 0;
const FIT_TIGHT = 1;
const FIT_LOOSE = 2;
const FIT_VERY_LOOSE = 3;

type FitnessClass = 0 | 1 | 2 | 3;

function fitnessClass(r: number): FitnessClass {
  if (r < -0.5) return FIT_VERY_TIGHT;
  if (r <= 0.5) return FIT_TIGHT;
  if (r <= 1) return FIT_LOOSE;
  return FIT_VERY_LOOSE;
}

interface ActiveNode {
  readonly itemIndex: number;
  readonly line: number;
  readonly fitness: FitnessClass;
  // Cumulative running sums up to (and not including) this node's break.
  readonly totalWidth: number;
  readonly totalStretch: number;
  readonly totalShrink: number;
  readonly totalDemerits: number;
  readonly previous: ActiveNode | null;
}

/**
 * Break a paragraph into lines using the Knuth-Plass total-fit algorithm.
 *
 * `lineWidths`: either a number (uniform width) or an array indexed by line
 * (0-based). When the array is shorter than the paragraph the last entry is
 * reused for trailing lines — this matches the common case of "first line
 * different, rest the same".
 */
export function breakLines(
  items: ReadonlyArray<Item>,
  lineWidths: number | ReadonlyArray<number>,
): BreakResult {
  if (items.length === 0) return { breaks: [], ratios: [] };

  const widthFor = (line: number): number => {
    if (typeof lineWidths === 'number') return lineWidths;
    if (line < lineWidths.length) return lineWidths[line]!;
    return lineWidths[lineWidths.length - 1]!;
  };

  const startNode: ActiveNode = {
    itemIndex: -1,
    line: 0,
    fitness: FIT_TIGHT,
    totalWidth: 0,
    totalStretch: 0,
    totalShrink: 0,
    totalDemerits: 0,
    previous: null,
  };

  let active: Array<ActiveNode> = [startNode];

  const isFeasibleBreakpoint = (i: number): boolean => {
    const item = items[i]!;
    if (item.type === 'penalty' && item.penalty < FORBIDDEN_BREAK) return true;
    if (item.type === 'glue') {
      // A glue can be a break point only if it follows a box.
      const prev = items[i - 1];
      return !!prev && prev.type === 'box';
    }
    return false;
  };

  // Snapshot cumulative widths up to (exclusive of) each item so we can
  // compute line widths between any two breakpoints in O(1).
  const widthBeforeItem: Array<number> = new Array(items.length + 1);
  const stretchBeforeItem: Array<number> = new Array(items.length + 1);
  const shrinkBeforeItem: Array<number> = new Array(items.length + 1);
  widthBeforeItem[0] = 0;
  stretchBeforeItem[0] = 0;
  shrinkBeforeItem[0] = 0;
  for (let i = 0; i < items.length; i++) {
    const it = items[i]!;
    const addW = it.type === 'penalty' ? 0 : it.width;
    const addSt = it.type === 'glue' ? it.stretch : 0;
    const addSh = it.type === 'glue' ? it.shrink : 0;
    widthBeforeItem[i + 1] = widthBeforeItem[i]! + addW;
    stretchBeforeItem[i + 1] = stretchBeforeItem[i]! + addSt;
    shrinkBeforeItem[i + 1] = shrinkBeforeItem[i]! + addSh;
  }

  // Line width for the run starting after node `a` and breaking at item `b`.
  // Per Knuth-Plass, a penalty at `b` contributes its own width (e.g. the
  // hyphen char); a glue at `b` is dropped (the line ends before it).
  const measureLine = (a: ActiveNode, b: number) => {
    // Start cumulative position is the item *after* the previous break.
    const startIdx = a.itemIndex + 1;
    let width = widthBeforeItem[b]! - widthBeforeItem[startIdx]!;
    const stretch = stretchBeforeItem[b]! - stretchBeforeItem[startIdx]!;
    const shrink = shrinkBeforeItem[b]! - shrinkBeforeItem[startIdx]!;
    const item = items[b]!;
    if (item.type === 'penalty') width += item.width;
    return { width, stretch, shrink };
  };

  const computeRatio = (
    line: { width: number; stretch: number; shrink: number },
    target: number,
  ): number => {
    if (line.width < target) {
      // No stretch to fill the gap → line is underfull and cannot be made to
      // fit. Return Infinity so the caller marks the break infeasible.
      return line.stretch > 0 ? (target - line.width) / line.stretch : Infinity;
    }
    if (line.width > target) {
      // No shrink → overfull line, infeasible (r < -1).
      return line.shrink > 0 ? (target - line.width) / line.shrink : -Infinity;
    }
    return 0;
  };

  const lineDemerits = (
    r: number,
    penaltyAtBreak: number,
    lastItemPenalty: PenaltyItem | null,
  ): number => {
    const badness = 100 * Math.pow(Math.abs(r), 3);
    let d: number;
    if (penaltyAtBreak >= 0 && penaltyAtBreak < FORBIDDEN_BREAK) {
      d = Math.pow(LINE_PENALTY + badness + penaltyAtBreak, 2);
    } else if (penaltyAtBreak > FORCED_BREAK && penaltyAtBreak < 0) {
      d = Math.pow(LINE_PENALTY + badness, 2) - penaltyAtBreak * penaltyAtBreak;
    } else {
      d = Math.pow(LINE_PENALTY + badness, 2);
    }
    void lastItemPenalty;
    return d;
  };

  for (let i = 0; i < items.length; i++) {
    if (!isFeasibleBreakpoint(i)) continue;
    const item = items[i]!;
    const penaltyAtBreak = item.type === 'penalty' ? item.penalty : 0;

    // For each active node, see if breaking at i is feasible.
    const survivingActive: Array<ActiveNode> = [];
    // Bucket of best new candidate per (line, fitness) pair.
    const candidates = new Map<string, ActiveNode>();

    for (const a of active) {
      const target = widthFor(a.line);
      const line = measureLine(a, i);
      const r = computeRatio(line, target);

      const isForced = penaltyAtBreak <= FORCED_BREAK;

      if (r < -1 && !isForced) {
        // Overfull line — adding more content only makes this worse, so the
        // node can never produce a feasible line ending later in the stream.
        // Deactivate (do not push to surviving).
        continue;
      }
      // A break within tolerance is feasible. A break that is underfull *beyond*
      // tolerance (r > TOLERANCE_RATIO) is too loose to be ideal, but a loose
      // line is still visually acceptable — unlike an overfull one, which spills
      // past the box and overlaps its neighbours. So we keep the node active for
      // a tighter break later AND record the loose break as a candidate. Its
      // badness (∝ r³, uncapped on the loose side) makes it a genuine last resort
      // that never wins when a within-tolerance break exists, but lets a narrow
      // column (where no tight break fits, e.g. long words in a thin table cell)
      // still wrap loosely instead of producing an overfull, overlapping line.
      // Forced breaks drain the active set per Knuth-Plass §3, so the node only
      // survives for non-forced breaks.
      if (!isForced) {
        survivingActive.push(a);
      }

      // Clamp r so demerits stay finite: overfull → -1; an underfull *forced*
      // break (a short, justified last line) keeps its historical tolerance clamp.
      // A non-forced too-loose break keeps its real (large) ratio so the looser
      // it is, the worse it scores.
      const effectiveR = r < -1 ? -1 : isForced && r > TOLERANCE_RATIO ? TOLERANCE_RATIO : r;
      const fc = fitnessClass(effectiveR);
      let demerits = lineDemerits(effectiveR, penaltyAtBreak, null);
      // Adjacent-line fitness mismatch penalty.
      if (Math.abs(fc - a.fitness) > 1) demerits += FITNESS_PENALTY;
      // Two consecutive hyphenated breaks add an extra π.
      if (item.type === 'penalty' && item.flagged) {
        const prevItem = a.itemIndex >= 0 ? items[a.itemIndex] : null;
        if (prevItem && prevItem.type === 'penalty' && prevItem.flagged)
          demerits += FLAGGED_PENALTY;
      }
      demerits += a.totalDemerits;

      const key = `${a.line + 1}:${fc}`;
      const existing = candidates.get(key);
      if (!existing || demerits < existing.totalDemerits) {
        candidates.set(key, {
          itemIndex: i,
          line: a.line + 1,
          fitness: fc,
          totalWidth: widthBeforeItem[i + 1]!,
          totalStretch: stretchBeforeItem[i + 1]!,
          totalShrink: shrinkBeforeItem[i + 1]!,
          totalDemerits: demerits,
          previous: a,
        });
      }
    }

    active = survivingActive;
    for (const c of candidates.values()) active.push(c);

    if (active.length === 0) {
      // Pathological: nothing fits anywhere. Fall back to inserting one
      // emergency break at the current item.
      active = [
        {
          itemIndex: i,
          line: 1,
          fitness: FIT_TIGHT,
          totalWidth: widthBeforeItem[i + 1]!,
          totalStretch: stretchBeforeItem[i + 1]!,
          totalShrink: shrinkBeforeItem[i + 1]!,
          totalDemerits: 0,
          previous: startNode,
        },
      ];
    } else if (active.length > MAX_ACTIVE_NODES) {
      // Trim the high-demerit tail; nodes with much worse cumulative cost are
      // unlikely to win the global minimum.
      active.sort((p, q) => p.totalDemerits - q.totalDemerits);
      active.length = MAX_ACTIVE_NODES;
    }
  }

  // Pick the active node with the lowest total demerits.
  let best: ActiveNode | null = null;
  for (const a of active) {
    if (!best || a.totalDemerits < best.totalDemerits) best = a;
  }
  if (!best) return { breaks: [], ratios: [] };

  // Walk back to reconstruct breakpoints in source order.
  const breaks: Array<number> = [];
  const ratios: Array<number> = [];
  let node: ActiveNode = best;
  while (node.previous) {
    breaks.unshift(node.itemIndex);
    const prev = node.previous;
    const target = widthFor(prev.line);
    const line = measureLine(prev, node.itemIndex);
    const r = computeRatio(line, target);
    ratios.unshift(r < -1 ? -1 : r);
    node = prev;
  }
  return { breaks, ratios };
}
