// Formula value model (E-SHEET W9) — the runtime values a formula evaluates to,
// plus the Excel coercion rules between them. A spreadsheet formula yields one of
// five scalar kinds (number, text, logical, error, blank) or a reference to a
// rectangle of cells; the evaluator (eval.ts) and the function library
// (functions.ts) thread these through. Errors propagate: any operand that is an
// error short-circuits to that error (§18.17.2 error values).

import type { EvalContext } from '@/excel/formula/context';

// §18.17.2 — the seven error values a formula can carry. Stored as their literal
// display text so they round-trip and read naturally in diagnostics.
export type FErr = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A';

// A rectangle of cells (absolute, 0-indexed, inclusive) — the value an A1:B3
// reference evaluates to. A single cell is a 1×1 rect. Aggregate functions
// iterate it; scalar contexts dereference it (1×1 → its value, else #VALUE!).
export interface Rect {
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

export type Scalar =
  | { readonly t: 'num'; readonly v: number }
  | { readonly t: 'str'; readonly v: string }
  | { readonly t: 'bool'; readonly v: boolean }
  | { readonly t: 'err'; readonly v: FErr }
  | { readonly t: 'blank' };

export type FValue = Scalar | { readonly t: 'ref'; readonly rect: Rect };

export const BLANK: Scalar = { t: 'blank' };
export const TRUE: Scalar = { t: 'bool', v: true };
export const FALSE: Scalar = { t: 'bool', v: false };

export function num(v: number): Scalar {
  // A non-finite arithmetic result (overflow, 0/0 handled by callers) is #NUM!.
  return Number.isFinite(v) ? { t: 'num', v } : { t: 'err', v: '#NUM!' };
}
export function str(v: string): Scalar {
  return { t: 'str', v };
}
export function bool(v: boolean): Scalar {
  return v ? TRUE : FALSE;
}
export function err(v: FErr): Scalar {
  return { t: 'err', v };
}

export function isErr(v: FValue): v is { t: 'err'; v: FErr } {
  return v.t === 'err';
}

// Collapse a reference to a scalar: a 1×1 rect yields the cell's value (a blank
// cell → blank); any larger rect in a scalar context is #VALUE! (we do not
// implement implicit intersection). Non-refs pass through unchanged.
export function deref(v: FValue, ctx: EvalContext): Scalar {
  if (v.t !== 'ref') return v;
  const { r0, c0, r1, c1 } = v.rect;
  if (r0 === r1 && c0 === c1) return ctx.getCell(r0, c0);
  return err('#VALUE!');
}

// §18.17.3 — number coercion. blank → 0, logical → 1/0, a numeric string →
// its number (Excel parses "5" and " 5 " leniently); non-numeric text → #VALUE!.
// An error propagates. References are dereferenced first.
export function toNumber(v: FValue, ctx: EvalContext): number | FErr {
  const s = deref(v, ctx);
  switch (s.t) {
    case 'num':
      return s.v;
    case 'bool':
      return s.v ? 1 : 0;
    case 'blank':
      return 0;
    case 'err':
      return s.v;
    case 'str': {
      const trimmed = s.v.trim();
      if (trimmed === '') return '#VALUE!';
      const n = Number(trimmed);
      return Number.isFinite(n) ? n : '#VALUE!';
    }
  }
}

// §18.17.3 — text coercion. A number formats with the shortest round-trip
// representation, a logical as TRUE/FALSE, blank as the empty string. Errors
// propagate (returned as the FErr).
export function toText(v: FValue, ctx: EvalContext): string | FErr {
  const s = deref(v, ctx);
  switch (s.t) {
    case 'str':
      return s.v;
    case 'num':
      return numToText(s.v);
    case 'bool':
      return s.v ? 'TRUE' : 'FALSE';
    case 'blank':
      return '';
    case 'err':
      return s.v;
  }
}

// §18.17.3 — logical coercion for IF/AND/OR and a rule's final truth test. A
// number is true iff non-zero; a numeric string coerces then tests; blank →
// false; non-numeric text → #VALUE!. Errors propagate.
export function toBool(v: FValue, ctx: EvalContext): boolean | FErr {
  const s = deref(v, ctx);
  switch (s.t) {
    case 'bool':
      return s.v;
    case 'num':
      return s.v !== 0;
    case 'blank':
      return false;
    case 'err':
      return s.v;
    case 'str': {
      const u = s.v.trim().toUpperCase();
      if (u === 'TRUE') return true;
      if (u === 'FALSE') return false;
      const n = Number(s.v.trim());
      if (s.v.trim() !== '' && Number.isFinite(n)) return n !== 0;
      return '#VALUE!';
    }
  }
}

// The shortest text for a number that still round-trips — JS `String` already
// gives Excel-compatible output for the integer / short-decimal values that
// dominate cached cell values (Excel itself caps at 15 significant digits).
function numToText(n: number): string {
  if (Number.isInteger(n)) return String(n);
  return String(n);
}
