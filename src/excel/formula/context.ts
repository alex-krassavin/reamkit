// Evaluation context (E-SHEET W9) — everything a formula needs from the world
// outside its own syntax tree: the grid's cached cell values, the injected
// reference date, and the workbook date epoch. The context is the seam that
// keeps the evaluator pure and deterministic — it never reads the wall clock or
// recomputes a cell; it only looks up values Excel already cached.

import type { FValue, Rect, Scalar } from '@/excel/formula/value';

/**
 * Everything a formula needs from the world outside its own syntax tree (E-SHEET
 * W9): the grid's cached cell values, the injected reference date, and the workbook
 * date epoch — optionally extended with cross-sheet / defined-name resolution. The
 * context is the seam that keeps the evaluator pure and deterministic: it never
 * reads the wall clock or recomputes a cell, only looking up values Excel already
 * cached.
 */
export interface EvalContext {
  /**
   * The cached value of the cell at `(row, col)` — both absolute, 0-indexed. An
   * absent or empty cell yields a blank scalar; the evaluator never recurses into a
   * referenced cell's own formula (there is none stored — we have only the cached
   * value), which is exactly why no recalculation engine is needed.
   */
  readonly getCell: (row: number, col: number) => Scalar;
  /**
   * Visit every POPULATED cell within a rectangle (blanks are skipped — they
   * contribute nothing to SUM/COUNT/COUNTIF). Iterating the sparse cell set keeps an
   * aggregate over a whole-column range O(populated), not O(rows), mirroring the
   * colorScale extent scan. The order is unspecified.
   */
  readonly eachCell: (rect: Rect, visit: (row: number, col: number, value: Scalar) => void) => void;
  /**
   * The reference "today" as an Excel serial day (`options.now`, converted with the
   * workbook epoch). `TODAY()`/`NOW()` and the `timePeriod` windows read it.
   * undefined ⇒ those clock-relative constructs yield `#VALUE!` / no-op, preserving
   * determinism when the caller supplies no date.
   */
  readonly nowSerial: number | undefined;
  /**
   * false = 1900 epoch (the default), true = 1904 epoch. Date functions
   * (YEAR/MONTH/DAY/DATE/WEEKDAY/…) convert serials ↔ calendar with it.
   */
  readonly date1904: boolean;
  // --- cross-workbook resolution (optional) — present only when the caller wired
  // the whole workbook in. Absent ⇒ a sheet-qualified reference / defined name
  // resolves to nothing (#REF!/#NAME?), exactly as before this was supported.
  /** A sheet name (case-insensitive) → its 0-based workbook index, or undefined. */
  readonly sheetIndex?: (name: string) => number | undefined;
  /** {@link EvalContext.getCell}, but on a specific sheet by index (cross-sheet references). */
  readonly getCellOn?: (sheet: number, row: number, col: number) => Scalar;
  /** {@link EvalContext.eachCell}, but on a specific sheet by index (cross-sheet references). */
  readonly eachCellOn?: (
    sheet: number,
    rect: Rect,
    visit: (row: number, col: number, value: Scalar) => void,
  ) => void;
  /**
   * Resolve a defined name (case-insensitive) to its target — a reference value (a
   * Rect, optionally on another sheet) or a literal scalar; undefined ⇒ `#NAME?`.
   */
  readonly resolveName?: (name: string) => FValue | undefined;
}
