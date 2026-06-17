// Evaluation context (E-SHEET W9) — everything a formula needs from the world
// outside its own syntax tree: the grid's cached cell values, the injected
// reference date, and the workbook date epoch. The context is the seam that
// keeps the evaluator pure and deterministic — it never reads the wall clock or
// recomputes a cell; it only looks up values Excel already cached.

import type { Rect, Scalar } from '@/excel/formula/value';

export interface EvalContext {
  // The cached value of the cell at (row, col) — both absolute, 0-indexed. An
  // absent or empty cell yields a blank scalar; the evaluator never recurses
  // into a referenced cell's own formula (there is none stored — we have only
  // the cached value), which is exactly why no recalculation engine is needed.
  readonly getCell: (row: number, col: number) => Scalar;
  // Visit every POPULATED cell within a rectangle (blanks are skipped — they
  // contribute nothing to SUM/COUNT/COUNTIF). Iterating the sparse cell set keeps
  // an aggregate over a whole-column range O(populated), not O(rows), mirroring
  // the colorScale extent scan. The order is unspecified.
  readonly eachCell: (rect: Rect, visit: (row: number, col: number, value: Scalar) => void) => void;
  // The reference "today" as an Excel serial day (options.now, converted with the
  // workbook epoch). TODAY()/NOW() and the timePeriod windows read it. undefined
  // ⇒ those clock-relative constructs yield #VALUE! / no-op, preserving
  // determinism when the caller supplies no date.
  readonly nowSerial: number | undefined;
  // false = 1900 epoch (the default), true = 1904 epoch. Date functions
  // (YEAR/MONTH/DAY/DATE/WEEKDAY/…) convert serials ↔ calendar with it.
  readonly date1904: boolean;
}
