// Formula evaluator (E-SHEET W9) — walks the AST against an EvalContext and a
// per-cell shift (how far the current cell sits from the rule's origin, used to
// resolve a conditional-format expression's relative references). Cells and
// ranges evaluate to reference values; scalar operators dereference them.
// Operator semantics follow Excel: arithmetic coerces operands to numbers (blank
// → 0), `&` to text, comparisons order number < text < logical with
// case-insensitive text and #DIV/0!/#VALUE!/#NUM! where Excel raises them.

import type { Ast, BinOp, CellRef } from '@/excel/formula/parser';
import type { EvalContext } from '@/excel/formula/context';
import type { FErr, FValue, Rect, Scalar } from '@/excel/formula/value';

import { callFn } from '@/excel/formula/functions';
import { bool, deref, err, num, str, toNumber, toText } from '@/excel/formula/value';

// The cell's offset from the rule origin, added to UNANCHORED reference axes.
export interface Shift {
  readonly dRow: number;
  readonly dCol: number;
}

export const NO_SHIFT: Shift = { dRow: 0, dCol: 0 };

export function evaluate(ast: Ast, ctx: EvalContext, shift: Shift): FValue {
  switch (ast.k) {
    case 'num':
      return num(ast.v);
    case 'str':
      return str(ast.v);
    case 'bool':
      return bool(ast.v);
    case 'err':
      return err(ast.v);
    case 'cell': {
      const rect = cellRect(ast.ref, shift);
      return rect ?? err('#REF!');
    }
    case 'range': {
      const a = resolveRef(ast.a, shift);
      const b = resolveRef(ast.b, shift);
      if (!a || !b) return err('#REF!');
      return {
        t: 'ref',
        rect: {
          r0: Math.min(a.row, b.row),
          c0: Math.min(a.col, b.col),
          r1: Math.max(a.row, b.row),
          c1: Math.max(a.col, b.col),
        },
      };
    }
    case 'name':
      // A defined name (or any unrecognised bare word) is unsupported → #NAME?.
      return err('#NAME?');
    case 'unary':
      return evalUnary(ast.op, evaluate(ast.x, ctx, shift), ctx);
    case 'pct': {
      const n = toNumber(evaluate(ast.x, ctx, shift), ctx);
      return typeof n === 'string' ? err(n) : num(n / 100);
    }
    case 'bin':
      return evalBin(ast.op, ast.a, ast.b, ctx, shift);
    case 'call':
      return callFn(ast.name, ast.args, (a) => evaluate(a, ctx, shift), ctx);
  }
}

// Resolve a reference axis with the shift: an anchored ($) axis stays put, an
// unanchored axis moves by the cell's offset from the origin. Returns the
// absolute (row, col) or undefined when the shift pushes it off the grid (#REF!).
function resolveRef(ref: CellRef, shift: Shift): { row: number; col: number } | undefined {
  const row = ref.row.abs ? ref.row.index : ref.row.index + shift.dRow;
  const col = ref.col.abs ? ref.col.index : ref.col.index + shift.dCol;
  if (row < 0 || col < 0) return undefined;
  return { row, col };
}

function cellRect(ref: CellRef, shift: Shift): { t: 'ref'; rect: Rect } | undefined {
  const r = resolveRef(ref, shift);
  if (!r) return undefined;
  return { t: 'ref', rect: { r0: r.row, c0: r.col, r1: r.row, c1: r.col } };
}

function evalUnary(op: '-' | '+', x: FValue, ctx: EvalContext): FValue {
  const n = toNumber(x, ctx);
  if (typeof n === 'string') return err(n);
  return num(op === '-' ? -n : n);
}

function evalBin(op: BinOp, aAst: Ast, bAst: Ast, ctx: EvalContext, shift: Shift): FValue {
  const a = evaluate(aAst, ctx, shift);
  const b = evaluate(bAst, ctx, shift);
  if (op === '&') {
    const ta = toText(a, ctx);
    if (typeof ta !== 'string') return err(ta);
    const tb = toText(b, ctx);
    if (typeof tb !== 'string') return err(tb);
    return str(ta + tb);
  }
  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
    return compareOp(op, deref(a, ctx), deref(b, ctx));
  }
  // Arithmetic.
  const na = toNumber(a, ctx);
  if (typeof na === 'string') return err(na);
  const nb = toNumber(b, ctx);
  if (typeof nb === 'string') return err(nb);
  switch (op) {
    case '+':
      return num(na + nb);
    case '-':
      return num(na - nb);
    case '*':
      return num(na * nb);
    case '/':
      return nb === 0 ? err('#DIV/0!') : num(na / nb);
    case '^': {
      const r = Math.pow(na, nb);
      return Number.isNaN(r) ? err('#NUM!') : num(r);
    }
  }
}

function compareOp(op: BinOp, a: Scalar, b: Scalar): FValue {
  const c = compare(a, b);
  if (typeof c === 'string') return err(c);
  switch (op) {
    case '=':
      return bool(c === 0);
    case '<>':
      return bool(c !== 0);
    case '<':
      return bool(c < 0);
    case '>':
      return bool(c > 0);
    case '<=':
      return bool(c <= 0);
    case '>=':
      return bool(c >= 0);
    default:
      return err('#VALUE!');
  }
}

// Three-way compare two scalars per Excel ordering: number < text < logical.
// Text compares case-insensitively; a blank takes the neutral element of the
// other operand's category (0 vs a number, "" vs text, FALSE vs a logical).
// Either operand being an error short-circuits to that error.
function compare(a: Scalar, b: Scalar): number | FErr {
  if (a.t === 'err') return a.v;
  if (b.t === 'err') return b.v;
  const na = normalizeBlank(a, b);
  const nb = normalizeBlank(b, a);
  const ca = category(na);
  const cb = category(nb);
  if (ca !== cb) return ca < cb ? -1 : 1;
  if (na.t === 'str' && nb.t === 'str') {
    const la = na.v.toLowerCase();
    const lb = nb.v.toLowerCase();
    return la < lb ? -1 : la > lb ? 1 : 0;
  }
  const va = scalarNumber(na);
  const vb = scalarNumber(nb);
  return va < vb ? -1 : va > vb ? 1 : 0;
}

// A blank operand mirrors the other's category so 0 = blank-vs-number, "" =
// blank-vs-text, FALSE = blank-vs-logical. Two blanks compare equal (as 0).
function normalizeBlank(s: Scalar, other: Scalar): Scalar {
  if (s.t !== 'blank') return s;
  switch (other.t) {
    case 'str':
      return { t: 'str', v: '' };
    case 'bool':
      return { t: 'bool', v: false };
    default:
      return { t: 'num', v: 0 };
  }
}

function category(s: Scalar): number {
  switch (s.t) {
    case 'num':
      return 0;
    case 'str':
      return 1;
    case 'bool':
      return 2;
    default:
      return 0; // blank already normalised away by the caller
  }
}

function scalarNumber(s: Scalar): number {
  if (s.t === 'num') return s.v;
  if (s.t === 'bool') return s.v ? 1 : 0;
  return 0;
}
