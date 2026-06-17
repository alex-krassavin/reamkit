// Formula engine public surface (E-SHEET W9) — the deterministic expression
// evaluator that closes the last two Excel-render gaps: conditional-format
// `expression` rules (an arbitrary formula per cell) and `timePeriod` rules
// (clock-relative date windows). The engine evaluates a formula against the
// grid's CACHED cell values — no recalculation — with TODAY()/NOW() and the
// timePeriod windows reading an injected reference day, so the same bytes plus
// the same options.now always produce identical output and the wall clock is
// never read. A formula that uses a construct the engine does not model
// (sheet-qualified ref, an unknown function, a defined name) yields an error and
// the rule simply does not apply — Ream never misrenders.

import type { Ast } from '@/excel/formula/parser';
import type { EvalContext } from '@/excel/formula/context';
import type { Shift } from '@/excel/formula/eval';

import { parse } from '@/excel/formula/parser';
import { NO_SHIFT, evaluate } from '@/excel/formula/eval';
import { toBool } from '@/excel/formula/value';

export type { EvalContext } from '@/excel/formula/context';
export type { Shift } from '@/excel/formula/eval';
export type { Scalar, FValue, Rect, FErr } from '@/excel/formula/value';
export { num, str, bool, err, BLANK } from '@/excel/formula/value';
export { serialFromDate, serialToParts, timePeriodMatches } from '@/excel/formula/dates';
export { NO_SHIFT } from '@/excel/formula/eval';

// A parsed formula, ready to evaluate against many cells (the AST is compiled
// once per conditional-format rule, then evaluated per covered cell with a
// per-cell shift).
export interface CompiledFormula {
  readonly ast: Ast;
}

// Parse a formula string. Returns undefined (rather than throwing) when the
// formula cannot be parsed — the caller treats that as a rule that never
// applies, which is the correct graceful-loss behaviour for an unsupported
// construct.
export function compileFormula(src: string): CompiledFormula | undefined {
  try {
    return { ast: parse(src) };
  } catch {
    return undefined;
  }
}

// Evaluate a compiled formula in a cell context and reduce it to the rule's
// truth test: the conditional format applies iff the result is the logical TRUE
// or a non-zero number (Excel §18.3.1.10). Any error / blank / text / zero — and
// a multi-cell result — yields false, so the rule simply does not paint.
export function evaluateToBool(
  compiled: CompiledFormula,
  ctx: EvalContext,
  shift: Shift = NO_SHIFT,
): boolean {
  return toBool(evaluate(compiled.ast, ctx, shift), ctx) === true;
}
