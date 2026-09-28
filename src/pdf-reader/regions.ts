// The regions of a column a flowing reading takes one after the other: the
// stretches that read straight down, and the bands where blocks stand side by
// side and each reads down on its own.
//
// A page is not always one text read line by line. An invoice sets its dates
// and numbers in a stack of labels and values beside the address it bills —
// "Invoice Date / Jun 3, 2013 / Invoice Number / INV-0046" next to "Orange
// Demo Inc. / 23 Main Street / Central City" — and the two stacks keep their
// own leading: a line of one stands a point or three off the line of the other.
// Read across, every line of one caught a line of the other into a row of its
// own, and bigboundingbox.pdf came back as "INVOICE → Jun 3, 2013 → 23 Main
// Street", its title in the middle of the stack beside it.
//
// Blocks whose lines DO share their baselines are a row-by-row setting — an
// address beside the one it is billed to, line for line — and stay one text:
// its lines are read across, on the stops they stand on.

import type { TextRun } from './content';

/** A band of blocks side by side, each a column of its own. */
export interface SideBySide {
  readonly kind: 'side';
  readonly cells: ReadonlyArray<{
    readonly runs: ReadonlyArray<TextRun>;
    /** Where the cell's own ink starts and ends across the page. */
    readonly from: number;
    readonly to: number;
  }>;
}

/** A stretch of the column read line by line, straight down. */
export interface Flowing {
  readonly kind: 'flow';
  readonly runs: ReadonlyArray<TextRun>;
}

export type Region = Flowing | SideBySide;

/**
 * The column's runs cut into its regions, top of the page first.
 *
 * @param runs The column's runs.
 * @returns The regions, in reading order; one `flow` region where the column
 *          has no side-by-side band.
 */
export function regionsOf(runs: ReadonlyArray<TextRun>): Array<Region> {
  const lines = clusterLines(runs);
  const out: Array<Region> = [];
  // The lines read so far that belong to no band. A band may reach back over
  // the last of them — its first line need not be the first that is off.
  let pending: Array<Cluster> = [];
  const flush = (flow: ReadonlyArray<Cluster>): void => {
    if (flow.length > 0) out.push({ kind: 'flow', runs: flow.flatMap((l) => l.runs) });
  };
  let floor = 0;
  for (let i = 0; i < lines.length; ) {
    const band = bandFrom(lines, i, floor);
    if (band === undefined) {
      pending.push(lines[i]!);
      i++;
      continue;
    }
    flush(pending.slice(0, pending.length - (i - band.start)));
    out.push(band.region);
    pending = [];
    i = band.end;
    floor = band.end;
  }
  flush(pending);
  return out;
}

/** Runs sharing a baseline, within half the size, top of the page first. */
interface Cluster {
  readonly y: number;
  readonly size: number;
  readonly runs: Array<TextRun>;
}

function clusterLines(runs: ReadonlyArray<TextRun>): Array<Cluster> {
  const sorted = [...runs].sort((a, b) => b.y - a.y || a.x - b.x);
  const out: Array<{ y: number; size: number; runs: Array<TextRun> }> = [];
  for (const run of sorted) {
    const last = out[out.length - 1];
    if (last && Math.abs(last.y - run.y) <= Math.max(1, (run.fontSizePt || 10) * 0.5)) {
      last.runs.push(run);
      last.size = Math.max(last.size, run.fontSizePt || 0);
    } else {
      out.push({ y: run.y, size: run.fontSizePt || 10, runs: [run] });
    }
  }
  return out;
}

/** What a run marks: nothing for a blank or for a glyph that names no character. */
const inked = (r: TextRun): boolean => r.text.replaceAll('�', '').trim() !== '';

/** A line's pieces: its runs cut where a gap only a column leaves opens. */
function piecesOf(line: Cluster): Array<{ from: number; to: number; y: number; size: number }> {
  const runs = line.runs.filter(inked).sort((a, b) => a.x - b.x);
  const out: Array<{ from: number; to: number; ys: Array<number>; size: number }> = [];
  for (const run of runs) {
    const last = out[out.length - 1];
    const size = run.fontSizePt || 10;
    if (last && run.x - last.to < Math.max(size, last.size) * PIECE_GAP_EM) {
      last.to = Math.max(last.to, run.endX);
      last.ys.push(run.y);
      last.size = Math.max(last.size, size);
    } else {
      out.push({ from: run.x, to: run.endX, ys: [run.y], size });
    }
  }
  return out.map((p) => ({ from: p.from, to: p.to, y: middle(p.ys), size: p.size }));
}

/**
 * Whether a line's pieces stand on different baselines — the mark of two
 * blocks set side by side with leading of their own, which the clustering has
 * run together.
 */
function misaligned(line: Cluster): boolean {
  const pieces = piecesOf(line);
  if (pieces.length < 2) return false;
  const ys = pieces.map((p) => p.y);
  // Judged by the SMALLER type: a title set three sizes up beside a stack of
  // labels is off the labels' lines by more than its own rounding could put it.
  const size = Math.min(...pieces.map((p) => p.size));
  return Math.max(...ys) - Math.min(...ys) > Math.max(ALIGN_SLACK_PT, size * ALIGN_SLACK_EM);
}

/**
 * The side-by-side band starting at line `i`, where one does: lines around a
 * misaligned one that keep clear of the same gutters, cut into the blocks
 * between them.
 */
function bandFrom(
  lines: ReadonlyArray<Cluster>,
  i: number,
  floor: number,
): { region: SideBySide; start: number; end: number } | undefined {
  if (!misaligned(lines[i]!)) return undefined;
  let from = i;
  let to = i + 1;
  const count = (a: number, b: number): number => guttersOf(lines.slice(a, b)).length;
  if (count(from, to) === 0) return undefined;
  const near = (a: Cluster, b: Cluster): boolean =>
    Math.abs(a.y - b.y) <= Math.max(a.size, b.size) * BAND_REACH_EM;
  // Down, then up: the band takes in every line close enough to be the same
  // setting that leaves it as many gutters as it had — a line across one of
  // them is not one of its blocks' lines.
  while (
    to < lines.length &&
    near(lines[to - 1]!, lines[to]!) &&
    count(from, to + 1) >= count(from, to)
  ) {
    to++;
  }
  while (
    from > floor &&
    near(lines[from]!, lines[from - 1]!) &&
    count(from - 1, to) >= count(from, to)
  ) {
    from--;
  }
  const band = lines.slice(from, to);
  const gutters = guttersOf(band);
  if (gutters.length === 0) return undefined;
  const edges = [-Infinity, ...gutters.map((g) => (g.from + g.to) / 2), Infinity];
  const cells = edges.slice(0, -1).map((left, k) => {
    const right = edges[k + 1]!;
    const runs = band.flatMap((l) =>
      l.runs.filter((r) => {
        const mid = (r.x + r.endX) / 2;
        return mid >= left && mid < right;
      }),
    );
    const ink = runs.filter(inked);
    return {
      runs,
      from: Math.min(...ink.map((r) => r.x)),
      to: Math.max(...ink.map((r) => r.endX)),
    };
  });
  if (cells.filter((c) => c.runs.some(inked)).length < 2) return undefined;
  return {
    region: { kind: 'side', cells: cells.filter((c) => c.runs.some(inked)) },
    start: from,
    end: to,
  };
}

/** The strips across a band that no line's ink crosses, and that are wide enough to be gutters. */
function guttersOf(band: ReadonlyArray<Cluster>): Array<{ from: number; to: number }> {
  const spans = band
    .flatMap((l) => piecesOf(l))
    .map((p) => [p.from, p.to] as const)
    .sort((a, b) => a[0] - b[0]);
  // The band's TEXT size, not its largest: a title set in 24pt over a stack of
  // 9pt labels is not a reason for the labels' gutter to be twice as wide.
  const size = middle(band.map((l) => l.size));
  const out: Array<{ from: number; to: number }> = [];
  let reach = spans[0]?.[1] ?? 0;
  for (const [from, to] of spans.slice(1)) {
    if (from - reach >= size * PIECE_GAP_EM) out.push({ from: reach, to: from });
    reach = Math.max(reach, to);
  }
  return out;
}

function middle(values: ReadonlyArray<number>): number {
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)] ?? 0;
}

/** A gap this wide, in ems, is one a column leaves, not a word space or a tab within a line. */
const PIECE_GAP_EM = 1.5;

/**
 * How far apart two pieces' baselines may stand and still be one line: past a
 * point and a half — or a sixth of the size — the two were set with leading of
 * their own, not on one line with a rounding in it.
 */
const ALIGN_SLACK_PT = 1.5;
const ALIGN_SLACK_EM = 1 / 6;

/** How far, in ems, a line may stand from the band and still belong to it. */
const BAND_REACH_EM = 2.5;
