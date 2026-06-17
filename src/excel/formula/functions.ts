// Formula function library (E-SHEET W9) — a curated, deterministic subset of the
// Excel function set chosen for what conditional-format expressions actually use:
// logic (AND/OR/IF/IFERROR/IS*), math (ABS/MOD/ROUND…/SUM/AVERAGE/MIN/MAX/COUNT),
// text (LEN/LEFT/RIGHT/MID/SEARCH/EXACT/…), the range predicates COUNTIF/SUMIF,
// and the date functions (TODAY/NOW/YEAR/MONTH/DAY/WEEKDAY/DATE/EDATE/EOMONTH)
// that read the injected reference day. An unknown function returns #NAME? so the
// rule simply does not apply — Ream never guesses and never misrenders.

import type { Ast } from '@/excel/formula/parser';
import type { EvalContext } from '@/excel/formula/context';
import type { FErr, FValue, Rect, Scalar } from '@/excel/formula/value';

import { serialFromYmd, serialToParts } from '@/excel/formula/dates';
import { bool, deref, err, isErr, num, str, toBool, toNumber, toText } from '@/excel/formula/value';

type Ev = (a: Ast) => FValue;

export function callFn(name: string, args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  switch (name) {
    // --- logic ---------------------------------------------------------------
    case 'TRUE':
      return bool(true);
    case 'FALSE':
      return bool(false);
    case 'NOT': {
      if (args.length !== 1) return err('#VALUE!');
      const b = toBool(ev(args[0]!), ctx);
      return typeof b === 'string' ? err(b) : bool(!b);
    }
    case 'AND':
    case 'OR':
      return andOr(name, args, ev, ctx);
    case 'IF': {
      if (args.length < 2 || args.length > 3) return err('#VALUE!');
      const c = toBool(ev(args[0]!), ctx);
      if (typeof c === 'string') return err(c);
      if (c) return ev(args[1]!);
      return args.length === 3 ? ev(args[2]!) : bool(false);
    }
    case 'IFERROR': {
      if (args.length !== 2) return err('#VALUE!');
      const v = ev(args[0]!);
      return isErr(deref(v, ctx)) ? ev(args[1]!) : v;
    }
    case 'ISBLANK':
      return is1(args, ev, ctx, (s) => s.t === 'blank');
    case 'ISNUMBER':
      return is1(args, ev, ctx, (s) => s.t === 'num');
    case 'ISTEXT':
      return is1(args, ev, ctx, (s) => s.t === 'str');
    case 'ISLOGICAL':
      return is1(args, ev, ctx, (s) => s.t === 'bool');
    case 'ISERROR':
      return is1(args, ev, ctx, (s) => s.t === 'err');
    case 'ISERR':
      return is1(args, ev, ctx, (s) => s.t === 'err' && s.v !== '#N/A');
    case 'ISNA':
      return is1(args, ev, ctx, (s) => s.t === 'err' && s.v === '#N/A');
    case 'N': {
      if (args.length !== 1) return err('#VALUE!');
      const s = deref(ev(args[0]!), ctx);
      if (s.t === 'num') return s;
      if (s.t === 'bool') return num(s.v ? 1 : 0);
      if (s.t === 'err') return s;
      return num(0);
    }

    // --- math ----------------------------------------------------------------
    case 'ABS':
      return math1(args, ev, ctx, Math.abs);
    case 'SIGN':
      return math1(args, ev, ctx, Math.sign);
    case 'SQRT':
      return math1(args, ev, ctx, (x) => (x < 0 ? NaN : Math.sqrt(x)));
    case 'INT':
      return math1(args, ev, ctx, Math.floor);
    case 'MOD':
      return mod(args, ev, ctx);
    case 'POWER':
      return power(args, ev, ctx);
    case 'ROUND':
      return roundFn(args, ev, ctx, 'round');
    case 'ROUNDUP':
      return roundFn(args, ev, ctx, 'up');
    case 'ROUNDDOWN':
    case 'TRUNC':
      return roundFn(args, ev, ctx, 'down');
    case 'SUM':
      return aggregate(args, ev, ctx, (ns) => ns.reduce((a, b) => a + b, 0));
    case 'AVERAGE':
      return aggregate(
        args,
        ev,
        ctx,
        (ns) => (ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : NaN),
        true,
      );
    case 'MIN':
      return aggregate(args, ev, ctx, (ns) => (ns.length ? Math.min(...ns) : 0));
    case 'MAX':
      return aggregate(args, ev, ctx, (ns) => (ns.length ? Math.max(...ns) : 0));
    case 'COUNT':
      return countFn(args, ev, ctx, false);
    case 'COUNTA':
      return countFn(args, ev, ctx, true);
    case 'COUNTIF':
      return countSumIf(args, ev, ctx, false);
    case 'SUMIF':
      return countSumIf(args, ev, ctx, true);

    // --- text ----------------------------------------------------------------
    case 'LEN': {
      const t = textArg(args, ev, ctx, 0);
      return typeof t === 'string' ? num(t.length) : err(t.v);
    }
    case 'LEFT':
      return leftRight(args, ev, ctx, true);
    case 'RIGHT':
      return leftRight(args, ev, ctx, false);
    case 'MID':
      return mid(args, ev, ctx);
    case 'LOWER':
      return text1(args, ev, ctx, (s) => s.toLowerCase());
    case 'UPPER':
      return text1(args, ev, ctx, (s) => s.toUpperCase());
    case 'TRIM':
      return text1(args, ev, ctx, (s) => s.replace(/\s+/g, ' ').trim());
    case 'CONCATENATE':
      return concat(args, ev, ctx);
    case 'EXACT':
      return exact(args, ev, ctx);
    case 'SEARCH':
      return findSearch(args, ev, ctx, false);
    case 'FIND':
      return findSearch(args, ev, ctx, true);
    case 'VALUE': {
      const t = textArg(args, ev, ctx, 0);
      if (typeof t !== 'string') return err(t.v);
      const n = Number(t.trim());
      return t.trim() !== '' && Number.isFinite(n) ? num(n) : err('#VALUE!');
    }

    // --- date ----------------------------------------------------------------
    case 'TODAY':
    case 'NOW':
      // NOW carries no time-of-day here (we keep the integer day) so the engine
      // stays deterministic; both read the injected reference serial.
      return ctx.nowSerial === undefined ? err('#VALUE!') : num(ctx.nowSerial);
    case 'YEAR':
      return datePart(args, ev, ctx, (p) => p.year);
    case 'MONTH':
      return datePart(args, ev, ctx, (p) => p.month);
    case 'DAY':
      return datePart(args, ev, ctx, (p) => p.day);
    case 'WEEKDAY':
      return weekday(args, ev, ctx);
    case 'DATE':
      return dateFn(args, ev, ctx);
    case 'EDATE':
      return edate(args, ev, ctx, false);
    case 'EOMONTH':
      return edate(args, ev, ctx, true);

    default:
      return err('#NAME?');
  }
}

// --- helpers ---------------------------------------------------------------

function is1(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  test: (s: Scalar) => boolean,
): FValue {
  if (args.length !== 1) return err('#VALUE!');
  return bool(test(deref(ev(args[0]!), ctx)));
}

function numArg(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  i: number,
): number | { e: FErr } {
  const n = toNumber(ev(args[i]!), ctx);
  return typeof n === 'string' ? { e: n } : n;
}

function textArg(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  i: number,
): string | { v: FErr } {
  const t = toText(ev(args[i]!), ctx);
  return typeof t === 'string' ? t : { v: t };
}

function math1(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  f: (x: number) => number,
): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  const r = f(a);
  return Number.isNaN(r) ? err('#NUM!') : num(r);
}

function text1(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  f: (s: string) => string,
): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const t = textArg(args, ev, ctx, 0);
  return typeof t === 'string' ? str(f(t)) : err(t.v);
}

function mod(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  const b = numArg(args, ev, ctx, 1);
  if (typeof b !== 'number') return err(b.e);
  if (b === 0) return err('#DIV/0!');
  // Excel MOD takes the sign of the divisor: n - d*FLOOR(n/d).
  return num(a - b * Math.floor(a / b));
}

function power(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  const b = numArg(args, ev, ctx, 1);
  if (typeof b !== 'number') return err(b.e);
  const r = Math.pow(a, b);
  return Number.isNaN(r) ? err('#NUM!') : num(r);
}

function roundFn(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  mode: 'round' | 'up' | 'down',
): FValue {
  if (args.length < 1 || args.length > 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  let digits = 0;
  if (args.length === 2) {
    const d = numArg(args, ev, ctx, 1);
    if (typeof d !== 'number') return err(d.e);
    digits = Math.trunc(d);
  }
  const f = Math.pow(10, digits);
  const x = Math.abs(a) * f;
  const sign = Math.sign(a);
  // Excel rounds half away from zero; ROUNDUP/DOWN are ceil/trunc on magnitude.
  const m = mode === 'round' ? Math.round(x) : mode === 'up' ? Math.ceil(x) : Math.floor(x);
  return num((sign * m) / f);
}

function andOr(name: 'AND' | 'OR', args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  let any = false;
  let acc = name === 'AND';
  for (const a of args) {
    const v = ev(a);
    if (v.t === 'ref') {
      let e: FErr | undefined;
      ctx.eachCell(v.rect, (_r, _c, s) => {
        if (s.t === 'err') {
          e ??= s.v;
        } else if (s.t === 'num' || s.t === 'bool') {
          const b = s.t === 'bool' ? s.v : s.v !== 0;
          any = true;
          acc = name === 'AND' ? acc && b : acc || b;
        }
        // text / blank cells are ignored by AND/OR over a range
      });
      if (e) return err(e);
    } else {
      const b = toBool(v, ctx);
      if (typeof b === 'string') return err(b);
      any = true;
      acc = name === 'AND' ? acc && b : acc || b;
    }
  }
  return any ? bool(acc) : err('#VALUE!');
}

// Collect the numeric values across the arguments (range cells contribute their
// numbers; a direct logical counts as 1/0; text and blanks are ignored). An
// error anywhere short-circuits to that error. The caller decides whether an
// empty set is valid (AVERAGE turns it into #DIV/0!).
function gatherNumbers(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): Array<number> | FErr {
  const out: Array<number> = [];
  for (const a of args) {
    const v = ev(a);
    if (v.t === 'ref') {
      let e: FErr | undefined;
      ctx.eachCell(v.rect, (_r, _c, s) => {
        if (s.t === 'num') out.push(s.v);
        else if (s.t === 'err') e ??= s.v;
      });
      if (e) return e;
    } else {
      const s = deref(v, ctx);
      if (s.t === 'err') return s.v;
      if (s.t === 'num') out.push(s.v);
      else if (s.t === 'bool') out.push(s.v ? 1 : 0);
    }
  }
  return out;
}

function aggregate(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  reduce: (ns: ReadonlyArray<number>) => number,
  requireSome = false,
): FValue {
  const ns = gatherNumbers(args, ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  if (requireSome && ns.length === 0) return err('#DIV/0!');
  const r = reduce(ns);
  return Number.isNaN(r) ? err('#DIV/0!') : num(r);
}

function countFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, nonBlank: boolean): FValue {
  let count = 0;
  for (const a of args) {
    const v = ev(a);
    if (v.t === 'ref') {
      ctx.eachCell(v.rect, (_r, _c, s) => {
        if (nonBlank ? s.t !== 'blank' : s.t === 'num') count++;
      });
    } else {
      const s = deref(v, ctx);
      if (nonBlank ? s.t !== 'blank' : s.t === 'num') count++;
    }
  }
  return num(count);
}

function leftRight(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, left: boolean): FValue {
  if (args.length < 1 || args.length > 2) return err('#VALUE!');
  const t = textArg(args, ev, ctx, 0);
  if (typeof t !== 'string') return err(t.v);
  let n = 1;
  if (args.length === 2) {
    const d = numArg(args, ev, ctx, 1);
    if (typeof d !== 'number') return err(d.e);
    n = Math.trunc(d);
    if (n < 0) return err('#VALUE!');
  }
  return str(left ? t.slice(0, n) : n === 0 ? '' : t.slice(Math.max(0, t.length - n)));
}

function mid(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 3) return err('#VALUE!');
  const t = textArg(args, ev, ctx, 0);
  if (typeof t !== 'string') return err(t.v);
  const start = numArg(args, ev, ctx, 1);
  if (typeof start !== 'number') return err(start.e);
  const len = numArg(args, ev, ctx, 2);
  if (typeof len !== 'number') return err(len.e);
  if (start < 1 || len < 0) return err('#VALUE!');
  const s = Math.trunc(start) - 1;
  return str(t.slice(s, s + Math.trunc(len)));
}

function concat(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  let out = '';
  for (let i = 0; i < args.length; i++) {
    const t = textArg(args, ev, ctx, i);
    if (typeof t !== 'string') return err(t.v);
    out += t;
  }
  return str(out);
}

function exact(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const a = textArg(args, ev, ctx, 0);
  if (typeof a !== 'string') return err(a.v);
  const b = textArg(args, ev, ctx, 1);
  if (typeof b !== 'string') return err(b.v);
  return bool(a === b);
}

function findSearch(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  caseSensitive: boolean,
): FValue {
  if (args.length < 2 || args.length > 3) return err('#VALUE!');
  let needle = textArg(args, ev, ctx, 0);
  if (typeof needle !== 'string') return err(needle.v);
  let hay = textArg(args, ev, ctx, 1);
  if (typeof hay !== 'string') return err(hay.v);
  let start = 1;
  if (args.length === 3) {
    const s = numArg(args, ev, ctx, 2);
    if (typeof s !== 'number') return err(s.e);
    start = Math.trunc(s);
    if (start < 1) return err('#VALUE!');
  }
  if (!caseSensitive) {
    needle = needle.toLowerCase();
    hay = hay.toLowerCase();
  }
  const idx = hay.indexOf(needle, start - 1);
  return idx < 0 ? err('#VALUE!') : num(idx + 1);
}

function datePart(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  pick: (p: { year: number; month: number; day: number; dow: number }) => number,
): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  if (a < 0) return err('#NUM!');
  return num(pick(serialToParts(a, ctx.date1904)));
}

function weekday(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 1 || args.length > 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  if (a < 0) return err('#NUM!');
  let type = 1;
  if (args.length === 2) {
    const t = numArg(args, ev, ctx, 1);
    if (typeof t !== 'number') return err(t.e);
    type = Math.trunc(t);
  }
  const dow = serialToParts(a, ctx.date1904).dow; // 0 = Sun … 6 = Sat
  switch (type) {
    case 1:
      return num(dow + 1); // Sun=1 … Sat=7
    case 2:
      return num(((dow + 6) % 7) + 1); // Mon=1 … Sun=7
    case 3:
      return num((dow + 6) % 7); // Mon=0 … Sun=6
    default:
      return err('#NUM!');
  }
}

function dateFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 3) return err('#VALUE!');
  const y = numArg(args, ev, ctx, 0);
  if (typeof y !== 'number') return err(y.e);
  const m = numArg(args, ev, ctx, 1);
  if (typeof m !== 'number') return err(m.e);
  const d = numArg(args, ev, ctx, 2);
  if (typeof d !== 'number') return err(d.e);
  const serial = serialFromYmd(Math.trunc(y), Math.trunc(m), Math.trunc(d), ctx.date1904);
  return serial < 0 ? err('#NUM!') : num(serial);
}

function edate(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, endOfMonth: boolean): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  if (a < 0) return err('#NUM!');
  const m = numArg(args, ev, ctx, 1);
  if (typeof m !== 'number') return err(m.e);
  const p = serialToParts(a, ctx.date1904);
  const idx = p.year * 12 + (p.month - 1) + Math.trunc(m);
  const ty = Math.floor(idx / 12);
  const tm = (idx % 12) + 1; // 1-indexed
  const dim = daysInMonth(ty, tm);
  const day = endOfMonth ? dim : Math.min(p.day, dim);
  const serial = serialFromYmd(ty, tm, day, ctx.date1904);
  return serial < 0 ? err('#NUM!') : num(serial);
}

function daysInMonth(year: number, month1: number): number {
  return new Date(Date.UTC(year, month1, 0)).getUTCDate();
}

function countSumIf(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, sum: boolean): FValue {
  const minArgs = sum ? 2 : 2;
  const maxArgs = sum ? 3 : 2;
  if (args.length < minArgs || args.length > maxArgs) return err('#VALUE!');
  const range = ev(args[0]!);
  if (range.t !== 'ref') return err('#VALUE!');
  const crit = parseCriteria(deref(ev(args[1]!), ctx));
  if ('e' in crit) return err(crit.e);
  // SUMIF's optional sum-range is offset-aligned to the test range; we resolve
  // each matching test cell's position into the sum-range by the same delta.
  const sumRect = sum && args.length === 3 ? asRef(ev(args[2]!)) : undefined;
  if (sum && args.length === 3 && !sumRect) return err('#VALUE!');
  let count = 0;
  let total = 0;
  let e: FErr | undefined;
  ctx.eachCell(range.rect, (r, c, s) => {
    if (!matchCriteria(s, crit)) return;
    count++;
    if (!sum) return;
    let target = s;
    if (sumRect) {
      const sr = sumRect.r0 + (r - range.rect.r0);
      const sc = sumRect.c0 + (c - range.rect.c0);
      target = ctx.getCell(sr, sc);
    }
    if (target.t === 'num') total += target.v;
    else if (target.t === 'err') e ??= target.v;
  });
  if (e) return err(e);
  return num(sum ? total : count);
}

function asRef(v: FValue): Rect | undefined {
  return v.t === 'ref' ? v.rect : undefined;
}

interface NumCrit {
  readonly kind: 'num';
  readonly op: '=' | '<>' | '>' | '>=' | '<' | '<=';
  readonly n: number;
}
interface TextCrit {
  readonly kind: 'text';
  readonly negate: boolean;
  readonly re: RegExp;
}

// Parse a COUNTIF/SUMIF criteria scalar into a predicate spec. A bare number or
// logical is an equality test; a string may carry a leading comparison operator
// and (for equality) `*`/`?` wildcards. A blank criteria matches blanks/zero.
function parseCriteria(s: Scalar): NumCrit | TextCrit | { e: FErr } {
  if (s.t === 'err') return { e: s.v };
  if (s.t === 'num') return { kind: 'num', op: '=', n: s.v };
  if (s.t === 'bool') return { kind: 'num', op: '=', n: s.v ? 1 : 0 };
  if (s.t === 'blank') return { kind: 'num', op: '=', n: 0 };
  let text = s.v;
  let op: NumCrit['op'] = '=';
  const m = /^(<>|>=|<=|>|<|=)/.exec(text);
  if (m) {
    op = m[1] as NumCrit['op'];
    text = text.slice(m[1]!.length);
  }
  const n = Number(text.trim());
  if (text.trim() !== '' && Number.isFinite(n)) return { kind: 'num', op, n };
  return { kind: 'text', negate: op === '<>', re: wildcardToRegExp(text) };
}

function matchCriteria(cell: Scalar, crit: NumCrit | TextCrit): boolean {
  if (crit.kind === 'num') {
    const v = cell.t === 'num' ? cell.v : cell.t === 'bool' ? (cell.v ? 1 : 0) : undefined;
    if (v === undefined) return false;
    switch (crit.op) {
      case '=':
        return v === crit.n;
      case '<>':
        return v !== crit.n;
      case '>':
        return v > crit.n;
      case '>=':
        return v >= crit.n;
      case '<':
        return v < crit.n;
      case '<=':
        return v <= crit.n;
    }
  }
  const text = cell.t === 'str' ? cell.v : cell.t === 'blank' ? '' : undefined;
  if (text === undefined) return false;
  const hit = crit.re.test(text);
  return crit.negate ? !hit : hit;
}

// Excel wildcard text criteria: `*` = any run, `?` = one char, `~` escapes the
// next wildcard. Case-insensitive, anchored to the whole string.
function wildcardToRegExp(pattern: string): RegExp {
  let out = '^';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!;
    if (
      ch === '~' &&
      (pattern[i + 1] === '*' || pattern[i + 1] === '?' || pattern[i + 1] === '~')
    ) {
      out += escapeRe(pattern[i + 1]!);
      i++;
    } else if (ch === '*') {
      out += '.*';
    } else if (ch === '?') {
      out += '.';
    } else {
      out += escapeRe(ch);
    }
  }
  out += '$';
  return new RegExp(out, 'is');
}

function escapeRe(ch: string): string {
  return /[.*+?^${}()|[\]\\]/.test(ch) ? `\\${ch}` : ch;
}
