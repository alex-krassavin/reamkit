// Formula evaluator (E-SHEET W9) — walks the AST against an EvalContext and a
// per-cell shift (how far the current cell sits from the rule's origin, used to
// resolve a conditional-format expression's relative references). Cells and
// ranges evaluate to reference values; scalar operators dereference them.
// Operator semantics follow Excel: arithmetic coerces operands to numbers (blank
// → 0), `&` to text, comparisons order number < text < logical with
// case-insensitive text and #DIV/0!/#VALUE!/#NUM! where Excel raises them.

import type { Ast, BinOp, CellRef } from '@/excel/formula/parser';
import type { EvalContext } from '@/excel/formula/context';
import type { FErr, FValue, Scalar } from '@/excel/formula/value';

import { callFn } from '@/excel/formula/functions';
import { bool, deref, err, num, str, toNumber, toText } from '@/excel/formula/value';

/**
 * The cell's offset from the rule origin, added to UNANCHORED reference axes.
 * `curRow`/`curCol` carry the ABSOLUTE current cell (origin + delta) so `ROW()`/
 * `COLUMN()` with no argument can report it; absent ({@link NO_SHIFT}) ⇒ those
 * no-arg forms have no cell to name and yield `#VALUE!`.
 */
export interface Shift {
  /** Row delta from the rule origin, added to an unanchored row axis. */
  readonly dRow: number;
  /** Column delta from the rule origin, added to an unanchored column axis. */
  readonly dCol: number;
  /** The absolute current row, for no-arg `ROW()`; absent ⇒ `#VALUE!`. */
  readonly curRow?: number;
  /** The absolute current column, for no-arg `COLUMN()`; absent ⇒ `#VALUE!`. */
  readonly curCol?: number;
}

/** The zero shift — no offset and no current cell (no-arg `ROW`/`COLUMN` → `#VALUE!`). */
export const NO_SHIFT: Shift = { dRow: 0, dCol: 0 };

/**
 * Evaluate an {@link Ast} against an `EvalContext` and a per-cell {@link Shift},
 * yielding an `FValue` (a scalar, a reference, or an array). Cells/ranges
 * evaluate to reference values; scalar operators dereference them.
 *
 * @param ast   The syntax tree to evaluate.
 * @param ctx   The evaluation context (grid access, defined names, date system).
 * @param shift The current cell's offset from the rule origin.
 * @returns The computed value.
 */
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
      const sheet = sheetIdx(ast.sheet, ctx);
      if (sheet === false) return err('#REF!');
      const r = resolveRef(ast.ref, shift);
      if (!r) return err('#REF!');
      return {
        t: 'ref',
        rect: { r0: r.row, c0: r.col, r1: r.row, c1: r.col },
        ...(sheet !== undefined ? { sheet } : {}),
      };
    }
    case 'range': {
      const sheet = sheetIdx(ast.sheet, ctx);
      if (sheet === false) return err('#REF!');
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
        ...(sheet !== undefined ? { sheet } : {}),
      };
    }
    case 'name':
      // A defined name resolves against the workbook (a reference or a literal);
      // an unknown name — or no workbook wired in — is #NAME?.
      return ctx.resolveName?.(ast.name) ?? err('#NAME?');
    case 'array':
      // An inline array constant → its scalar elements (each reduced now).
      return {
        t: 'arr',
        rows: ast.rows.map((row) => row.map((el) => deref(evaluate(el, ctx, shift), ctx))),
      };
    case 'unary':
      return evalUnary(ast.op, evaluate(ast.x, ctx, shift), ctx);
    case 'pct': {
      const n = toNumber(evaluate(ast.x, ctx, shift), ctx);
      return typeof n === 'string' ? err(n) : num(n / 100);
    }
    case 'bin':
      return evalBin(ast.op, ast.a, ast.b, ctx, shift);
    case 'call':
      return callFn(ast.name, ast.args, (a) => evaluate(a, ctx, shift), ctx, shift);
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

// Resolve a sheet-qualifier name to a workbook sheet index: undefined ⇒ the
// rule's own sheet (no qualifier); a number ⇒ that sheet; false ⇒ an unknown
// sheet (or no workbook wired in), which the caller turns into #REF!.
function sheetIdx(name: string | undefined, ctx: EvalContext): number | undefined | false {
  if (name === undefined) return undefined;
  const idx = ctx.sheetIndex?.(name);
  return idx === undefined ? false : idx;
}

function evalUnary(op: '-' | '+', x: FValue, ctx: EvalContext): FValue {
  const n = toNumber(x, ctx);
  if (typeof n === 'string') return err(n);
  return num(op === '-' ? -n : n);
}

function evalBin(op: BinOp, aAst: Ast, bAst: Ast, ctx: EvalContext, shift: Shift): FValue {
  const a = evaluate(aAst, ctx, shift);
  const b = evaluate(bAst, ctx, shift);
  // An array operand broadcasts the operator element-wise (e.g. A1={1,3,5}).
  if (a.t === 'arr' || b.t === 'arr') return broadcastBin(op, a, b, ctx);
  return applyBin(op, a, b, ctx);
}

// Apply a binary operator to a (scalar/reference) operand pair — the core scalar
// semantics, shared by the direct path and by array broadcasting.
function applyBin(op: BinOp, a: FValue, b: FValue, ctx: EvalContext): FValue {
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

// Broadcast a binary operator over an array operand: a scalar (or 1×1) pairs with
// every element; two equal-shaped arrays combine element-wise; a 1-row/1-col
// operand stretches along that axis. Mismatched shapes → #VALUE!.
function broadcastBin(op: BinOp, a: FValue, b: FValue, ctx: EvalContext): FValue {
  const ar = a.t === 'arr' ? a.rows : [[deref(a, ctx)]];
  const br = b.t === 'arr' ? b.rows : [[deref(b, ctx)]];
  const nRows = Math.max(ar.length, br.length);
  const out: Array<Array<Scalar>> = [];
  for (let r = 0; r < nRows; r++) {
    const arow = ar[ar.length === 1 ? 0 : r];
    const brow = br[br.length === 1 ? 0 : r];
    if (!arow || !brow) return err('#VALUE!');
    const nCols = Math.max(arow.length, brow.length);
    const row: Array<Scalar> = [];
    for (let c = 0; c < nCols; c++) {
      const ea = arow[arow.length === 1 ? 0 : c];
      const eb = brow[brow.length === 1 ? 0 : c];
      if (ea === undefined || eb === undefined) return err('#VALUE!');
      const res = applyBin(op, ea, eb, ctx);
      row.push(res.t === 'ref' || res.t === 'arr' ? err('#VALUE!') : res);
    }
    out.push(row);
  }
  return { t: 'arr', rows: out };
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
