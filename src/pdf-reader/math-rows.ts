// §22 (OfficeMathML) — the MATRIX a page drew, read back out of where it drew it.
//
// A PDF states no mathematics. A matrix reaches the page as numbers on two
// baselines with a pair of stretched brackets drawn between them, and read line
// by line — which is all a page says — bug1997343.pdf's
//
//     ( 1 2 ) ( 1 1 )   ( 1 3 )
//     ( 3 4 ) ( 0 1 ) = ( 3 7 )
//
// came back as three lines of prose: "1 2 1 1 1 3", "( )( ) = ( )", "3 4 0 1 3 7".
// The rows are there in the geometry and so are the brackets, so the matrix is
// rebuilt as the object it is — a delimiter around a grid of cells — which the
// layout then sets the way the file set it.
//
// What says "matrix" and not "three lines of text": the baselines stand CLOSER
// together than the page's own leading (they are sub-lines of one display), the
// middle one carries brackets, and the rows above and below fill the space
// between a bracket pair.

import type { MathNode } from '@/core/document-model';
import type { TextRun } from './content';

/** A matrix found on the page: what to draw, where, and which runs it used. */
export interface MathBlock {
  /** The math object — a row of delimiters and whatever stands between them. */
  readonly math: MathNode;
  /** The top baseline it was drawn on, for ordering the page's blocks. */
  readonly top: number;
  /** Where it stands across the page. */
  readonly x: number;
  readonly width: number;
  /** The runs it consumed, which the prose reading must not read again. */
  readonly used: ReadonlySet<TextRun>;
}

/** The brackets a matrix may be written in, opening → closing. */
const BRACKETS: ReadonlyMap<string, string> = new Map([
  ['(', ')'],
  ['[', ']'],
  ['{', '}'],
  ['⟨', '⟩'],
  ['|', '|'],
]);

/**
 * The matrices a page's rows hold (§22.1.2.68 `m:m`).
 *
 * @param runs     The column's runs.
 * @param fontSize The text's own size, which says how far apart the lines of
 *                 PROSE stand — a display's sub-lines stand closer.
 * @returns One entry per matrix found, in page order.
 */
export function matrixBlocks(runs: ReadonlyArray<TextRun>, fontSize: number): Array<MathBlock> {
  const out: Array<MathBlock> = [];
  for (const cluster of clusters(baselines(runs, fontSize), fontSize)) {
    const found = matrixFrom(cluster);
    if (found) out.push(found);
  }
  return out;
}

/** The baseline a row of runs stands on. */
const baselineOf = (row: ReadonlyArray<TextRun>): number => Math.max(...row.map((r) => r.y));

/**
 * The runs grouped by the baseline they were DRAWN on, top to bottom.
 *
 * Tighter than the reading's own rows, which allow a line half an em of drift
 * so that a footnote mark stays with its word: the sub-lines of a display stand
 * six points apart in a ten-point face, and grouped that loosely
 * bug1997343.pdf's brackets joined the row of numbers under them.
 */
function baselines(runs: ReadonlyArray<TextRun>, fontSize: number): Array<ReadonlyArray<TextRun>> {
  const tol = Math.max(1, fontSize * SAME_BASELINE);
  const out: Array<Array<TextRun>> = [];
  for (const run of [...runs].sort((a, b) => b.y - a.y || a.x - b.x)) {
    const last = out[out.length - 1];
    if (last && Math.abs(baselineOf(last) - run.y) <= tol) last.push(run);
    else out.push([run]);
  }
  return out;
}

/** How far off a baseline, in ems, a run may sit and still be ON it. */
const SAME_BASELINE = 0.15;

/**
 * The runs of a page grouped into the SUB-LINES of one display: rows standing
 * closer together than a line of prose does, three or more of them (two rows
 * and the brackets between).
 */
function clusters(
  rows: ReadonlyArray<ReadonlyArray<TextRun>>,
  fontSize: number,
): Array<Array<ReadonlyArray<TextRun>>> {
  const out: Array<Array<ReadonlyArray<TextRun>>> = [];
  let run: Array<ReadonlyArray<TextRun>> = [];
  for (const row of rows) {
    const last = run[run.length - 1];
    if (last && baselineOf(last) - baselineOf(row) >= fontSize * SUBLINE_GAP) {
      if (run.length >= MIN_SUBLINES) out.push(run);
      run = [];
    }
    run.push(row);
  }
  if (run.length >= MIN_SUBLINES) out.push(run);
  return out;
}

/** How far apart, in ems, two baselines have to stand to be separate LINES. */
const SUBLINE_GAP = 0.9;

/** A matrix is a row of cells above another, and the brackets between them. */
const MIN_SUBLINES = 3;

/** Build the matrix a cluster of sub-lines holds, or nothing where it holds none. */
function matrixFrom(cluster: ReadonlyArray<ReadonlyArray<TextRun>>): MathBlock | undefined {
  // The brackets stand on the row BETWEEN the rows of cells, which is where a
  // stretched bracket's own baseline falls.
  const middle = cluster.find((row) => pairsOf(row).length > 0);
  if (!middle) return undefined;
  const pairs = pairsOf(middle);
  const cells = cluster.filter((row) => row !== middle);
  if (cells.length === 0) return undefined;
  const children: Array<MathNode> = [];
  let used = new Set<TextRun>();
  let at = -Infinity;
  for (const pair of pairs) {
    // Whatever stands between the last bracket and this one — the `=` of an
    // equation, an operator — is text of the display, in the order it was set.
    for (const run of middle) {
      if (run.x >= at && run.endX <= pair.open.x && run.text.trim().length > 0) {
        children.push({ type: 'run', text: spaced(run.text) });
        used.add(run);
      }
    }
    const inside = matrixInside(cluster, middle, pair);
    if (!inside) return undefined;
    children.push({
      type: 'delimiter',
      begChr: pair.open.text.trim(),
      endChr: pair.close.text.trim(),
      children: [inside.matrix],
    });
    used = new Set([...used, ...inside.used, pair.open, pair.close]);
    at = pair.close.endX;
  }
  for (const run of middle) {
    if (run.x >= at && run.text.trim().length > 0) {
      children.push({ type: 'run', text: spaced(run.text) });
      used.add(run);
    }
  }
  if (children.length === 0) return undefined;
  // Every run of the cluster has to belong to the matrix. A display with a line
  // of prose beside it is not one object, and re-set as one the prose would
  // move: only a cluster the matrix accounts for entirely is read this way.
  for (const row of cluster) for (const run of row) if (!used.has(run)) return undefined;
  const all = [...used];
  return {
    math: { type: 'row', children },
    top: baselineOf(cells[0] ?? middle),
    x: Math.min(...all.map((r) => r.x)),
    width: Math.max(...all.map((r) => r.endX)) - Math.min(...all.map((r) => r.x)),
    used,
  };
}

/**
 * An operator standing between two matrices, with the air the page gave it.
 * The file sets `(1 2 / 3 4)(1 1 / 0 1) = (1 3 / 3 7)` with the equals sign
 * clear of both brackets; set tight against them it reads as one word.
 */
const spaced = (text: string): string => ` ${text.trim()} `;

/** An opening bracket and the closing one that answers it. */
interface Pair {
  readonly open: TextRun;
  readonly close: TextRun;
}

/** The bracket pairs a row holds, left to right and never nested. */
function pairsOf(row: ReadonlyArray<TextRun>): Array<Pair> {
  const out: Array<Pair> = [];
  let open: TextRun | undefined;
  for (const run of [...row].sort((a, b) => a.x - b.x)) {
    const text = run.text.trim();
    if (open === undefined) {
      if (BRACKETS.has(text)) open = run;
      continue;
    }
    if (text === BRACKETS.get(open.text.trim())) {
      out.push({ open, close: run });
      open = undefined;
    }
  }
  return out;
}

/**
 * The grid inside one bracket pair: every sub-line's runs that fall between the
 * brackets, clustered into columns by where they stand.
 */
function matrixInside(
  cluster: ReadonlyArray<ReadonlyArray<TextRun>>,
  middle: ReadonlyArray<TextRun>,
  pair: Pair,
): { matrix: MathNode; used: Set<TextRun> } | undefined {
  const used = new Set<TextRun>();
  const rows: Array<Array<TextRun>> = [];
  for (const row of cluster) {
    const inside = row.filter(
      (run) =>
        run !== pair.open &&
        run !== pair.close &&
        run.x >= pair.open.x &&
        run.endX <= pair.close.endX &&
        run.text.trim().length > 0,
    );
    if (inside.length === 0) continue;
    // The bracket row itself may carry the middle row of an odd-sized matrix.
    if (row === middle && inside.some((run) => BRACKETS.has(run.text.trim()))) return undefined;
    inside.sort((a, b) => a.x - b.x);
    for (const run of inside) used.add(run);
    rows.push(inside);
  }
  if (rows.length < 2) return undefined;
  // The columns: a cell of one row stands over a cell of the next, so the
  // centres line up. Anything else is a line of text that happens to be short.
  const size = Math.max(...rows.flat().map((r) => r.fontSizePt || 10));
  const centres: Array<number> = [];
  const columnOf = (run: TextRun): number => {
    const centre = (run.x + run.endX) / 2;
    const found = centres.findIndex((c) => Math.abs(c - centre) <= size / 2);
    if (found >= 0) return found;
    centres.push(centre);
    return centres.length - 1;
  };
  const grid = rows.map((row) => {
    const cells: Array<Array<string>> = [];
    for (const run of row) {
      const at = columnOf(run);
      (cells[at] ??= []).push(run.text.trim());
    }
    return cells;
  });
  const width = Math.max(...grid.map((row) => row.length));
  return {
    matrix: {
      type: 'matrix',
      rows: grid.map((row) =>
        Array.from(
          { length: width },
          (_, i): MathNode => ({ type: 'run', text: (row[i] ?? []).join('') }),
        ),
      ),
    },
    used,
  };
}
