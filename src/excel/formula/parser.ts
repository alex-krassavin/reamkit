// Formula parser (E-SHEET W9) — a precedence-climbing (Pratt) parser turning the
// token stream into an AST. Operator precedence follows ECMA-376 §18.17.2 /
// Excel: range `:` > unary `-`/`+` > postfix `%` > `^` > `*` `/` > `+` `-` >
// `&` > comparisons. A "word" token is classified here: a `(` after it makes it
// a function call; otherwise an A1-shaped word is a cell reference (its `$`
// anchors recorded per axis for the conditional-format relative shift), TRUE/
// FALSE a logical literal, anything else a defined name (unsupported → #NAME?).

import type { FErr } from '@/excel/formula/value';
import type { Token } from '@/excel/formula/lexer';

import { tokenize } from '@/excel/formula/lexer';

// One coordinate of a cell reference. `abs` records whether the axis was `$`
// anchored — an unanchored axis shifts by the cell's offset from the rule's
// origin when a conditional-format expression is evaluated per cell.
export interface Axis {
  readonly index: number; // 0-indexed row or column
  readonly abs: boolean;
}

export interface CellRef {
  readonly col: Axis;
  readonly row: Axis;
}

export type Ast =
  | { readonly k: 'num'; readonly v: number }
  | { readonly k: 'str'; readonly v: string }
  | { readonly k: 'bool'; readonly v: boolean }
  | { readonly k: 'err'; readonly v: FErr }
  | { readonly k: 'cell'; readonly ref: CellRef; readonly sheet?: string }
  | { readonly k: 'range'; readonly a: CellRef; readonly b: CellRef; readonly sheet?: string }
  | { readonly k: 'name'; readonly name: string }
  | { readonly k: 'array'; readonly rows: ReadonlyArray<ReadonlyArray<Ast>> }
  | { readonly k: 'unary'; readonly op: '-' | '+'; readonly x: Ast }
  | { readonly k: 'pct'; readonly x: Ast }
  | { readonly k: 'bin'; readonly op: BinOp; readonly a: Ast; readonly b: Ast }
  | { readonly k: 'call'; readonly name: string; readonly args: ReadonlyArray<Ast> };

export type BinOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=';

export class ParseError extends Error {}

// Precedence + associativity per binary operator. Higher binds tighter.
const PREC: Readonly<Record<BinOp, number>> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
};
const RIGHT_ASSOC: ReadonlySet<BinOp> = new Set<BinOp>(['^']);

const MAX_DEPTH = 64;
const MAX_ARGS = 255;

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;
const MAX_COL = 16383; // XFD
const MAX_ROW = 1048575;

// Parse a formula string into an AST, or throw ParseError. Callers (the CF
// compiler) catch and treat a parse failure as "rule does not apply" — a formula
// using a construct we do not model never misrenders, it just no-ops.
export function parse(src: string): Ast {
  const parser = new Parser(tokenize(src));
  const ast = parser.parseExpr(0);
  parser.expectEof();
  return ast;
}

class Parser {
  private pos = 0;
  private depth = 0;
  constructor(private readonly toks: ReadonlyArray<Token>) {}

  private peek(): Token {
    return this.toks[this.pos]!;
  }
  private next(): Token {
    return this.toks[this.pos++]!;
  }

  expectEof(): void {
    if (this.peek().kind !== 'eof') throw new ParseError(`trailing tokens at ${this.pos}`);
  }

  parseExpr(minBp: number): Ast {
    if (++this.depth > MAX_DEPTH) throw new ParseError('expression too deep');
    let left = this.parseUnary();
    for (;;) {
      const t = this.peek();
      if (t.kind !== 'op' || !(t.text in PREC)) break;
      const op = t.text as BinOp;
      const p = PREC[op];
      if (p < minBp) break;
      this.next();
      const right = this.parseExpr(RIGHT_ASSOC.has(op) ? p : p + 1);
      left = { k: 'bin', op, a: left, b: right };
    }
    this.depth--;
    return left;
  }

  // Unary +/- bind tighter than ^ (Excel: -2^2 = 4), so they live above the
  // binary loop; postfix % binds tighter still.
  private parseUnary(): Ast {
    const t = this.peek();
    if (t.kind === 'op' && (t.text === '-' || t.text === '+')) {
      this.next();
      return { k: 'unary', op: t.text, x: this.parseUnary() };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): Ast {
    let x = this.parsePrimary();
    while (this.peek().kind === 'op' && this.peek().text === '%') {
      this.next();
      x = { k: 'pct', x };
    }
    return x;
  }

  private parsePrimary(): Ast {
    const t = this.next();
    switch (t.kind) {
      case 'num': {
        const v = Number(t.text);
        if (!Number.isFinite(v)) throw new ParseError(`bad number ${t.text}`);
        return { k: 'num', v };
      }
      case 'str':
        return { k: 'str', v: t.text };
      case 'err':
        return { k: 'err', v: t.text as FErr };
      case 'word':
        // A word directly before `!` is a sheet qualifier (Sheet2!A1); otherwise
        // it is classified by parseWord (function / cell / TRUE/FALSE / name).
        if (this.peek().kind === 'op' && this.peek().text === '!') {
          this.next();
          return this.parseSheetCell(t.text);
        }
        return this.parseWord(t.text);
      case 'sheetq':
        // A 'quoted sheet name' is only valid immediately before `!` and a cell.
        this.expect('!');
        return this.parseSheetCell(t.text);
      case 'op':
        if (t.text === '(') {
          const e = this.parseExpr(0);
          this.expect(')');
          return e;
        }
        if (t.text === '{') return this.parseArray();
        throw new ParseError(`unexpected operator ${t.text}`);
      case 'eof':
        throw new ParseError('unexpected end of formula');
    }
  }

  private parseWord(word: string): Ast {
    // A `(` directly after the word makes it a function call, regardless of shape
    // (so LOG10( … ) is the function, bare LOG10 is the cell reference).
    if (this.peek().kind === 'op' && this.peek().text === '(') {
      this.next();
      const args = this.parseArgs();
      this.expect(')');
      return { k: 'call', name: word.toUpperCase(), args };
    }
    const upper = word.toUpperCase();
    if (upper === 'TRUE') return { k: 'bool', v: true };
    if (upper === 'FALSE') return { k: 'bool', v: false };
    const ref = parseCellRef(word);
    if (ref) {
      // A `:` between two cell references forms a range (A1:B3).
      if (this.peek().kind === 'op' && this.peek().text === ':') {
        const save = this.pos;
        this.next();
        const rhs = this.peek();
        if (rhs.kind === 'word') {
          const b = parseCellRef(rhs.text);
          if (b) {
            this.next();
            return { k: 'range', a: ref, b };
          }
        }
        this.pos = save; // not a range after all
      }
      return { k: 'cell', ref };
    }
    return { k: 'name', name: upper };
  }

  // A sheet-qualified cell or range — the `!` is already consumed; `sheet` is the
  // (unquoted) sheet name. `Sheet2!A1` or `Sheet2!A1:B3`. The evaluator resolves
  // the name against the workbook; an unknown sheet becomes #REF!.
  private parseSheetCell(sheet: string): Ast {
    const t = this.next();
    if (t.kind !== 'word') throw new ParseError('expected a cell after !');
    const a = parseCellRef(t.text);
    if (!a) throw new ParseError('expected a cell reference after !');
    if (this.peek().kind === 'op' && this.peek().text === ':') {
      this.next();
      const rhs = this.next();
      if (rhs.kind !== 'word') throw new ParseError('expected a cell after :');
      const b = parseCellRef(rhs.text);
      if (!b) throw new ParseError('bad range end after !');
      return { k: 'range', a, b, sheet };
    }
    return { k: 'cell', ref: a, sheet };
  }

  // An inline array constant {1,2,3} / {1,2;3,4}: rows separated by `;`, elements
  // by `,`. The opening `{` is already consumed. Elements are parsed as
  // expressions (so a signed literal like -1 works) and reduced to scalars at
  // eval; the total element count is capped against a crafted huge array.
  private parseArray(): Ast {
    const rows: Array<Array<Ast>> = [];
    let row: Array<Ast> = [];
    let count = 0;
    for (;;) {
      row.push(this.parseExpr(0));
      if (++count > MAX_ARGS) throw new ParseError('array constant too large');
      const t = this.peek();
      if (t.kind === 'op' && t.text === ',') {
        this.next();
        continue;
      }
      if (t.kind === 'op' && t.text === ';') {
        this.next();
        rows.push(row);
        row = [];
        continue;
      }
      break;
    }
    rows.push(row);
    this.expect('}');
    return { k: 'array', rows };
  }

  private parseArgs(): Array<Ast> {
    const args: Array<Ast> = [];
    if (this.peek().kind === 'op' && this.peek().text === ')') return args;
    for (;;) {
      args.push(this.parseExpr(0));
      if (args.length > MAX_ARGS) throw new ParseError('too many arguments');
      const t = this.peek();
      if (t.kind === 'op' && t.text === ',') {
        this.next();
        continue;
      }
      break;
    }
    return args;
  }

  private expect(op: string): void {
    const t = this.next();
    if (t.kind !== 'op' || t.text !== op) throw new ParseError(`expected ${op}`);
  }
}

// Classify + decode an A1-shaped word into a CellRef, or undefined if it is not
// a reference (then the caller treats it as a defined name). Out-of-range column
// or row also returns undefined (Excel would surface #REF!; we no-op the rule).
function parseCellRef(word: string): CellRef | undefined {
  const m = CELL_RE.exec(word);
  if (!m) return undefined;
  const colAbs = m[1] === '$';
  const col = columnToIndex(m[2]!);
  const rowAbs = m[3] === '$';
  const row = Number(m[4]) - 1;
  if (col > MAX_COL || row < 0 || row > MAX_ROW) return undefined;
  return { col: { index: col, abs: colAbs }, row: { index: row, abs: rowAbs } };
}

// Column letters → 0-indexed column (A→0, Z→25, AA→26, …).
function columnToIndex(letters: string): number {
  let idx = 0;
  const up = letters.toUpperCase();
  for (let i = 0; i < up.length; i++) {
    idx = idx * 26 + (up.charCodeAt(i) - 64); // 'A' = 65
  }
  return idx - 1;
}
