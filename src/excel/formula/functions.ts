// Formula function library (E-SHEET W9) — a deterministic subset of the Excel
// function set broad enough for what conditional-format expressions actually use:
//   logic / info: AND/OR/NOT/IF/IFERROR/IFNA/IFS/SWITCH/XOR/NA + IS* (incl.
//     ISEVEN/ISODD/ISNONTEXT/ISREF);
//   math: ABS/SIGN/SQRT/INT/MOD/POWER/ROUND…/TRUNC/PRODUCT/QUOTIENT/GCD/LCM/
//     MROUND/EVEN/ODD/CEILING/FLOOR + the aggregates SUM/AVERAGE/MIN/MAX/MEDIAN/
//     LARGE/SMALL/COUNT/COUNTA/COUNTBLANK/SUMPRODUCT;
//   range predicates: COUNTIF/SUMIF/AVERAGEIF + the multi-criteria COUNTIFS/
//     SUMIFS/AVERAGEIFS;
//   text: LEN/LEFT/RIGHT/MID/LOWER/UPPER/TRIM/PROPER/CONCATENATE/CONCAT/TEXTJOIN/
//     SUBSTITUTE/REPLACE/REPT/EXACT/SEARCH/FIND/CHAR/CODE/CLEAN/T/VALUE;
//   lookup: CHOOSE/MATCH/INDEX/VLOOKUP/HLOOKUP;
//   date / time: TODAY/NOW/YEAR/MONTH/DAY/WEEKDAY/DATE/EDATE/EOMONTH/DAYS/HOUR/
//     MINUTE/SECOND/TIME/WEEKNUM/ISOWEEKNUM — the clock-relative ones read the
//     injected reference day.
// Wave 4 broadens this toward the full surface a CF predicate can reach: the trig /
// exponential family (PI/EXP/LN/LOG/LOG10/SIN…/ASIN…/SINH…/DEGREES/RADIANS/SQRTPI/
// FACT/COMBIN/PERMUT/SUMSQ), the statistics (STDEV(P)/VAR(P) + .S/.P aliases, AVEDEV/
// DEVSQ/GEOMEAN/HARMEAN, AVERAGEA/MAXA/MINA, MODE/RANK + aliases, PERCENTILE/QUARTILE),
// the position functions (ROW/COLUMN — with or without a reference —, ROWS/COLUMNS),
// the info functions (TYPE/ERROR.TYPE), and UNICHAR/UNICODE.
// An unknown function (or a misused one) returns #NAME?/#VALUE! so the rule simply
// does not apply — Ream never guesses and never misrenders. Where an exact Excel
// semantic is ambiguous (e.g. CEILING with mixed-sign significance) the engine
// errs to #NUM! rather than risk a wrong truth value.

import type { Ast } from '@/excel/formula/parser';
import type { EvalContext } from '@/excel/formula/context';
import type { Shift } from '@/excel/formula/eval';
import type { FErr, FValue, Rect, Scalar } from '@/excel/formula/value';

import { serialFromYmd, serialToParts } from '@/excel/formula/dates';
import {
  bool,
  deref,
  err,
  isErr,
  num,
  refEach,
  refGet,
  str,
  toBool,
  toNumber,
  toText,
} from '@/excel/formula/value';

type Ev = (a: Ast) => FValue;

export function callFn(
  name: string,
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  shift: Shift,
): FValue {
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
    case 'DAYS':
      return daysFn(args, ev, ctx);
    case 'HOUR':
      return timePart(args, ev, ctx, 'h');
    case 'MINUTE':
      return timePart(args, ev, ctx, 'm');
    case 'SECOND':
      return timePart(args, ev, ctx, 's');
    case 'TIME':
      return timeFn(args, ev, ctx);
    case 'WEEKNUM':
      return weekNum(args, ev, ctx, false);
    case 'ISOWEEKNUM':
      return weekNum(args, ev, ctx, true);

    // --- extended logic / information ---------------------------------------
    case 'ISNONTEXT':
      return is1(args, ev, ctx, (s) => s.t !== 'str');
    case 'ISREF':
      return args.length === 1 ? bool(ev(args[0]!).t === 'ref') : err('#VALUE!');
    case 'ISEVEN':
      return isParity(args, ev, ctx, true);
    case 'ISODD':
      return isParity(args, ev, ctx, false);
    case 'NA':
      return err('#N/A');
    case 'IFNA': {
      if (args.length !== 2) return err('#VALUE!');
      const v = ev(args[0]!);
      const d = deref(v, ctx);
      return d.t === 'err' && d.v === '#N/A' ? ev(args[1]!) : v;
    }
    case 'XOR':
      return xorFn(args, ev, ctx);
    case 'IFS':
      return ifsFn(args, ev, ctx);
    case 'SWITCH':
      return switchFn(args, ev, ctx);

    // --- extended math ------------------------------------------------------
    case 'PRODUCT':
      return aggregate(args, ev, ctx, (ns) => (ns.length ? ns.reduce((a, b) => a * b, 1) : 0));
    case 'MEDIAN':
      return medianFn(args, ev, ctx);
    case 'LARGE':
      return largeSmall(args, ev, ctx, true);
    case 'SMALL':
      return largeSmall(args, ev, ctx, false);
    case 'QUOTIENT':
      return quotient(args, ev, ctx);
    case 'CEILING':
      return ceilingFloor(args, ev, ctx, 'ceil');
    case 'FLOOR':
      return ceilingFloor(args, ev, ctx, 'floor');
    case 'MROUND':
      return mround(args, ev, ctx);
    case 'EVEN':
      return evenOdd(args, ev, ctx, true);
    case 'ODD':
      return evenOdd(args, ev, ctx, false);
    case 'GCD':
      return gcdLcm(args, ev, ctx, 'gcd');
    case 'LCM':
      return gcdLcm(args, ev, ctx, 'lcm');
    case 'SUMPRODUCT':
      return sumProduct(args, ev, ctx);
    case 'COUNTBLANK':
      return countBlank(args, ev, ctx);
    case 'COUNTIFS':
      return countSumAvgIfs(args, ev, ctx, 'count');
    case 'SUMIFS':
      return countSumAvgIfs(args, ev, ctx, 'sum');
    case 'AVERAGEIFS':
      return countSumAvgIfs(args, ev, ctx, 'avg');
    case 'AVERAGEIF':
      return averageIf(args, ev, ctx);

    // --- extended text ------------------------------------------------------
    case 'CONCAT':
      return concatFlatten(args, ev, ctx);
    case 'TEXTJOIN':
      return textJoin(args, ev, ctx);
    case 'SUBSTITUTE':
      return substitute(args, ev, ctx);
    case 'REPLACE':
      return replaceFn(args, ev, ctx);
    case 'REPT':
      return rept(args, ev, ctx);
    case 'PROPER':
      return text1(args, ev, ctx, properCase);
    case 'CHAR':
      return charFn(args, ev, ctx);
    case 'CODE':
      return codeFn(args, ev, ctx);
    case 'CLEAN':
      return text1(args, ev, ctx, stripControl);
    case 'T':
      return tFn(args, ev, ctx);

    // --- lookup -------------------------------------------------------------
    case 'CHOOSE':
      return chooseFn(args, ev, ctx);
    case 'MATCH':
      return matchFn(args, ev, ctx);
    case 'INDEX':
      return indexFn(args, ev, ctx);
    case 'VLOOKUP':
      return lookupFn(args, ev, ctx, 'v');
    case 'HLOOKUP':
      return lookupFn(args, ev, ctx, 'h');

    // --- trig / exponential / extended math (Wave 4) ------------------------
    case 'PI':
      return args.length === 0 ? num(Math.PI) : err('#VALUE!');
    case 'EXP':
      return math1(args, ev, ctx, Math.exp);
    case 'LN':
      return math1(args, ev, ctx, (x) => (x <= 0 ? NaN : Math.log(x)));
    case 'LOG10':
      return math1(args, ev, ctx, (x) => (x <= 0 ? NaN : Math.log10(x)));
    case 'LOG':
      return logFn(args, ev, ctx);
    case 'SIN':
      return math1(args, ev, ctx, Math.sin);
    case 'COS':
      return math1(args, ev, ctx, Math.cos);
    case 'TAN':
      return math1(args, ev, ctx, Math.tan);
    case 'ASIN':
      return math1(args, ev, ctx, (x) => (Math.abs(x) > 1 ? NaN : Math.asin(x)));
    case 'ACOS':
      return math1(args, ev, ctx, (x) => (Math.abs(x) > 1 ? NaN : Math.acos(x)));
    case 'ATAN':
      return math1(args, ev, ctx, Math.atan);
    case 'ATAN2':
      // Excel ATAN2(x, y) is the angle of the point (x, y) — JS atan2 takes (y, x).
      return math2(args, ev, ctx, (x, y) => (x === 0 && y === 0 ? NaN : Math.atan2(y, x)));
    case 'SINH':
      return math1(args, ev, ctx, Math.sinh);
    case 'COSH':
      return math1(args, ev, ctx, Math.cosh);
    case 'TANH':
      return math1(args, ev, ctx, Math.tanh);
    case 'DEGREES':
      return math1(args, ev, ctx, (x) => (x * 180) / Math.PI);
    case 'RADIANS':
      return math1(args, ev, ctx, (x) => (x * Math.PI) / 180);
    case 'SQRTPI':
      return math1(args, ev, ctx, (x) => (x < 0 ? NaN : Math.sqrt(x * Math.PI)));
    case 'FACT':
      return factFn(args, ev, ctx);
    case 'COMBIN':
      return combinPermut(args, ev, ctx, false);
    case 'PERMUT':
      return combinPermut(args, ev, ctx, true);
    case 'SUMSQ':
      return aggregate(args, ev, ctx, (ns) => ns.reduce((a, b) => a + b * b, 0));

    // --- statistics (Wave 4) -------------------------------------------------
    case 'STDEV':
    case 'STDEV.S':
      return statSpread(args, ev, ctx, 'stdev', true);
    case 'STDEVP':
    case 'STDEV.P':
      return statSpread(args, ev, ctx, 'stdev', false);
    case 'VAR':
    case 'VAR.S':
      return statSpread(args, ev, ctx, 'var', true);
    case 'VARP':
    case 'VAR.P':
      return statSpread(args, ev, ctx, 'var', false);
    case 'AVEDEV':
      return statSpread(args, ev, ctx, 'avedev', false);
    case 'DEVSQ':
      return statSpread(args, ev, ctx, 'devsq', false);
    case 'GEOMEAN':
      return geoHarMean(args, ev, ctx, true);
    case 'HARMEAN':
      return geoHarMean(args, ev, ctx, false);
    case 'AVERAGEA':
      return aggregateA(args, ev, ctx, 'avg');
    case 'MAXA':
      return aggregateA(args, ev, ctx, 'max');
    case 'MINA':
      return aggregateA(args, ev, ctx, 'min');
    case 'MODE':
    case 'MODE.SNGL':
      return modeFn(args, ev, ctx);
    case 'RANK':
    case 'RANK.EQ':
      return rankFn(args, ev, ctx);
    case 'PERCENTILE':
    case 'PERCENTILE.INC':
      return percentileFn(args, ev, ctx);
    case 'QUARTILE':
    case 'QUARTILE.INC':
      return quartileFn(args, ev, ctx);

    // --- reference / information (Wave 4) -----------------------------------
    case 'ROW':
      return rowColFn(args, ev, shift, true);
    case 'COLUMN':
      return rowColFn(args, ev, shift, false);
    case 'ROWS':
      return rowsColsFn(args, ev, true);
    case 'COLUMNS':
      return rowsColsFn(args, ev, false);
    case 'TYPE':
      return typeFn(args, ev, ctx);
    case 'ERROR.TYPE':
      return errorTypeFn(args, ev, ctx);

    // --- text (Wave 4) -------------------------------------------------------
    case 'UNICHAR':
      return unicharFn(args, ev, ctx);
    case 'UNICODE':
      return unicodeFn(args, ev, ctx);

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
      refEach(v, ctx, (_r, _c, s) => {
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
      refEach(v, ctx, (_r, _c, s) => {
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
      refEach(v, ctx, (_r, _c, s) => {
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
  refEach(range, ctx, (r, c, s) => {
    if (!matchCriteria(s, crit)) return;
    count++;
    if (!sum) return;
    let target = s;
    if (sumRect) {
      const sr = sumRect.rect.r0 + (r - range.rect.r0);
      const sc = sumRect.rect.c0 + (c - range.rect.c0);
      target = refGet(sumRect.sheet, ctx, sr, sc);
    }
    if (target.t === 'num') total += target.v;
    else if (target.t === 'err') e ??= target.v;
  });
  if (e) return err(e);
  return num(sum ? total : count);
}

// A reference value reduced to its rectangle + (optional cross-sheet) sheet index;
// undefined when the value is not a reference. Range functions iterate it through
// refEach/refGet so a cross-sheet qualifier reads the right sheet.
type RefVal = { readonly rect: Rect; readonly sheet?: number };
function asRef(v: FValue): RefVal | undefined {
  return v.t === 'ref'
    ? { rect: v.rect, ...(v.sheet !== undefined ? { sheet: v.sheet } : {}) }
    : undefined;
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

// --- extended library helpers (broader CF function coverage) ---------------

// Cap a positional range scan (MATCH/INDEX/lookup/CONCAT/TEXTJOIN) so a
// whole-column reference cannot make a per-cell rule iterate millions of cells.
// A range past the cap yields a graceful error (the rule no-ops), never a hang.
const MAX_SCAN = 200_000;

function isParity(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, even: boolean): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  const odd = Math.abs(Math.trunc(a)) % 2 === 1;
  return bool(even ? !odd : odd);
}

function xorFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  let trues = 0;
  let any = false;
  for (const a of args) {
    const v = ev(a);
    if (v.t === 'ref') {
      let e: FErr | undefined;
      refEach(v, ctx, (_r, _c, s) => {
        if (s.t === 'err') e ??= s.v;
        else if (s.t === 'num' || s.t === 'bool') {
          any = true;
          if (s.t === 'bool' ? s.v : s.v !== 0) trues++;
        }
      });
      if (e) return err(e);
    } else {
      const b = toBool(v, ctx);
      if (typeof b === 'string') return err(b);
      any = true;
      if (b) trues++;
    }
  }
  return any ? bool(trues % 2 === 1) : err('#VALUE!');
}

function ifsFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 2 || args.length % 2 !== 0) return err('#VALUE!');
  for (let i = 0; i < args.length; i += 2) {
    const c = toBool(ev(args[i]!), ctx);
    if (typeof c === 'string') return err(c);
    if (c) return ev(args[i + 1]!);
  }
  return err('#N/A');
}

function switchFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 3) return err('#VALUE!');
  const target = deref(ev(args[0]!), ctx);
  if (target.t === 'err') return target;
  let i = 1;
  for (; i + 1 < args.length; i += 2) {
    const cand = deref(ev(args[i]!), ctx);
    if (cand.t === 'err') return cand;
    if (scalarEquals(target, cand)) return ev(args[i + 1]!);
  }
  // A trailing odd argument is the default value.
  return i < args.length ? ev(args[i]!) : err('#N/A');
}

// Exact-equality between two scalars (text case-insensitive), for SWITCH/MATCH.
function scalarEquals(a: Scalar, b: Scalar): boolean {
  if (a.t === 'num' && b.t === 'num') return a.v === b.v;
  if (a.t === 'str' && b.t === 'str') return a.v.toLowerCase() === b.v.toLowerCase();
  if (a.t === 'bool' && b.t === 'bool') return a.v === b.v;
  return a.t === 'blank' && b.t === 'blank';
}

// Three-way compare for MATCH/VLOOKUP approximate search; undefined ⇒ different
// types (incomparable, skipped). Text compares case-insensitively.
function cmpScalar(a: Scalar, b: Scalar): number | undefined {
  if (a.t === 'num' && b.t === 'num') return a.v < b.v ? -1 : a.v > b.v ? 1 : 0;
  if (a.t === 'str' && b.t === 'str') {
    const x = a.v.toLowerCase();
    const y = b.v.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (a.t === 'bool' && b.t === 'bool') return a.v === b.v ? 0 : a.v ? 1 : -1;
  return undefined;
}

function quotient(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  const b = numArg(args, ev, ctx, 1);
  if (typeof b !== 'number') return err(b.e);
  if (b === 0) return err('#DIV/0!');
  return num(Math.trunc(a / b));
}

// CEILING / FLOOR (legacy two-argument form): round the magnitude to a multiple
// of |significance|, away from zero (ceil) or toward zero (floor). A number and
// significance of opposite signs is #NUM! (we never guess the modern variant).
function ceilingFloor(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  mode: 'ceil' | 'floor',
): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  const sig = numArg(args, ev, ctx, 1);
  if (typeof sig !== 'number') return err(sig.e);
  if (sig === 0) return num(0);
  if (a !== 0 && a > 0 !== sig > 0) return err('#NUM!');
  const q = Math.abs(a) / Math.abs(sig);
  const snapped = Math.abs(q - Math.round(q)) < 1e-9 ? Math.round(q) : q;
  const mag = (mode === 'ceil' ? Math.ceil(snapped) : Math.floor(snapped)) * Math.abs(sig);
  return num(a < 0 ? -mag : mag);
}

// Round a non-negative ratio half-away-from-zero, snapping a value within a
// floating tolerance of *.5 up (1.3/0.2 = 6.4999… must round to 7, like Excel).
function roundHalfAway(q: number): number {
  const f = Math.floor(q);
  return Math.abs(q - f - 0.5) < 1e-9 ? f + 1 : Math.round(q);
}

function mround(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const n = numArg(args, ev, ctx, 0);
  if (typeof n !== 'number') return err(n.e);
  const mult = numArg(args, ev, ctx, 1);
  if (typeof mult !== 'number') return err(mult.e);
  if (mult === 0) return num(0);
  if (n !== 0 && n > 0 !== mult > 0) return err('#NUM!');
  const mag = roundHalfAway(Math.abs(n) / Math.abs(mult)) * Math.abs(mult);
  return num(n < 0 ? -mag : mag);
}

function evenOdd(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, even: boolean): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  let m = Math.ceil(Math.abs(a)); // round the magnitude away from zero to an int
  if (even) {
    if (m % 2 !== 0) m++;
  } else if (m % 2 === 0) {
    m++;
  }
  return num(a < 0 ? -m : m);
}

function gcd2(a: number, b: number): number {
  a = Math.abs(a);
  b = Math.abs(b);
  while (b) {
    const t = a % b;
    a = b;
    b = t;
  }
  return a;
}

function gcdLcm(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, mode: 'gcd' | 'lcm'): FValue {
  const ns = gatherNumbers(args, ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  if (ns.length === 0) return err('#VALUE!');
  const ints: Array<number> = [];
  for (const n of ns) {
    const t = Math.trunc(n);
    if (t < 0) return err('#NUM!');
    ints.push(t);
  }
  if (mode === 'gcd') return num(ints.reduce((a, b) => gcd2(a, b)));
  let l = ints[0]!;
  for (let i = 1; i < ints.length; i++) {
    const b = ints[i]!;
    if (l === 0 || b === 0) {
      l = 0;
      continue;
    }
    l = (l / gcd2(l, b)) * b;
    if (!Number.isSafeInteger(l)) return err('#NUM!');
  }
  return num(l);
}

function medianFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  const ns = gatherNumbers(args, ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  if (ns.length === 0) return err('#NUM!');
  const s = [...ns].sort((a, b) => a - b);
  const midIdx = Math.floor(s.length / 2);
  return num(s.length % 2 ? s[midIdx]! : (s[midIdx - 1]! + s[midIdx]!) / 2);
}

function largeSmall(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, large: boolean): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const ns = gatherNumbers([args[0]!], ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  const k = numArg(args, ev, ctx, 1);
  if (typeof k !== 'number') return err(k.e);
  const ki = Math.trunc(k);
  if (ki < 1 || ki > ns.length) return err('#NUM!');
  const s = [...ns].sort((a, b) => (large ? b - a : a - b));
  return num(s[ki - 1]!);
}

function countBlank(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const r = asRef(ev(args[0]!));
  if (!r) return err('#VALUE!');
  const area = (r.rect.r1 - r.rect.r0 + 1) * (r.rect.c1 - r.rect.c0 + 1);
  let filled = 0;
  refEach(r, ctx, (_x, _y, s) => {
    // An empty-string cell counts as blank for COUNTBLANK (like Excel).
    if (!(s.t === 'blank' || (s.t === 'str' && s.v === ''))) filled++;
  });
  return num(Math.max(0, area - filled));
}

// COUNTIFS / SUMIFS / AVERAGEIFS — multiple aligned (range, criteria) pairs, AND
// across them. The value range (sum/average) is the first argument for the IFS
// variants. Iterates the first range's POPULATED cells, aligning the others by
// offset — matching the existing COUNTIF/SUMIF semantics over the sparse grid.
function countSumAvgIfs(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  mode: 'count' | 'sum' | 'avg',
): FValue {
  const isCount = mode === 'count';
  let valueRef: RefVal | undefined;
  let pairStart = 0;
  if (!isCount) {
    const vr = asRef(ev(args[0]!));
    if (!vr) return err('#VALUE!');
    valueRef = vr;
    pairStart = 1;
  }
  const pairs = args.length - pairStart;
  if (pairs < 2 || pairs % 2 !== 0) return err('#VALUE!');
  const ranges: Array<RefVal> = [];
  const crits: Array<NumCrit | TextCrit> = [];
  for (let i = pairStart; i + 1 < args.length; i += 2) {
    const rng = asRef(ev(args[i]!));
    if (!rng) return err('#VALUE!');
    const crit = parseCriteria(deref(ev(args[i + 1]!), ctx));
    if ('e' in crit) return err(crit.e);
    ranges.push(rng);
    crits.push(crit);
  }
  const base = ranges[0]!.rect;
  const h = base.r1 - base.r0;
  const w = base.c1 - base.c0;
  for (const r of ranges)
    if (r.rect.r1 - r.rect.r0 !== h || r.rect.c1 - r.rect.c0 !== w) {
      return err('#VALUE!');
    }
  if (valueRef) {
    const vr = valueRef.rect;
    if (vr.r1 - vr.r0 !== h || vr.c1 - vr.c0 !== w) return err('#VALUE!');
  }
  let count = 0;
  let total = 0;
  let numCount = 0;
  let e: FErr | undefined;
  refEach(ranges[0]!, ctx, (r, c, s0) => {
    if (!matchCriteria(s0, crits[0]!)) return;
    for (let k = 1; k < ranges.length; k++) {
      const rk = ranges[k]!;
      const cell = refGet(rk.sheet, ctx, rk.rect.r0 + (r - base.r0), rk.rect.c0 + (c - base.c0));
      if (!matchCriteria(cell, crits[k]!)) return;
    }
    count++;
    if (isCount) return;
    const vref = valueRef!;
    const target = refGet(
      vref.sheet,
      ctx,
      vref.rect.r0 + (r - base.r0),
      vref.rect.c0 + (c - base.c0),
    );
    if (target.t === 'num') {
      total += target.v;
      numCount++;
    } else if (target.t === 'err') e ??= target.v;
  });
  if (e) return err(e);
  if (isCount) return num(count);
  if (mode === 'sum') return num(total);
  return numCount === 0 ? err('#DIV/0!') : num(total / numCount);
}

function averageIf(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 2 || args.length > 3) return err('#VALUE!');
  const range = asRef(ev(args[0]!));
  if (!range) return err('#VALUE!');
  const crit = parseCriteria(deref(ev(args[1]!), ctx));
  if ('e' in crit) return err(crit.e);
  const avgRef = args.length === 3 ? asRef(ev(args[2]!)) : range;
  if (!avgRef) return err('#VALUE!');
  const rr = range.rect;
  const ar = avgRef.rect;
  let total = 0;
  let count = 0;
  let e: FErr | undefined;
  refEach(range, ctx, (r, c, s) => {
    if (!matchCriteria(s, crit)) return;
    const target =
      avgRef === range ? s : refGet(avgRef.sheet, ctx, ar.r0 + (r - rr.r0), ar.c0 + (c - rr.c0));
    if (target.t === 'num') {
      total += target.v;
      count++;
    } else if (target.t === 'err') e ??= target.v;
  });
  if (e) return err(e);
  return count === 0 ? err('#DIV/0!') : num(total / count);
}

// SUMPRODUCT of equal-dimension ranges (the dot-product form). Iterates the first
// range's populated cells; a blank/text/zero factor at a position contributes a
// zero term. The array-condition idiom (SUMPRODUCT((A=x)*(B))) needs array
// semantics we do not model, so it gracefully #VALUE!s rather than misrender.
function sumProduct(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length === 0) return err('#VALUE!');
  const vals = args.map((a) => ev(a));
  const refs = vals.map((v) => asRef(v));
  if (refs.every((r) => r === undefined)) {
    let p = 1;
    for (const v of vals) {
      const n = toNumber(v, ctx);
      if (typeof n === 'string') return err(n);
      p *= n;
    }
    return num(p);
  }
  if (refs.some((r) => r === undefined)) return err('#VALUE!');
  const base = refs[0]!.rect;
  const h = base.r1 - base.r0;
  const w = base.c1 - base.c0;
  for (const r of refs)
    if (r!.rect.r1 - r!.rect.r0 !== h || r!.rect.c1 - r!.rect.c0 !== w) {
      return err('#VALUE!');
    }
  let total = 0;
  let e: FErr | undefined;
  refEach(refs[0]!, ctx, (r, c, s0) => {
    if (s0.t === 'err') {
      e ??= s0.v;
      return;
    }
    if (s0.t !== 'num') return;
    let term = s0.v;
    for (let k = 1; k < refs.length && term !== 0; k++) {
      const rk = refs[k]!;
      const cell = refGet(rk.sheet, ctx, rk.rect.r0 + (r - base.r0), rk.rect.c0 + (c - base.c0));
      if (cell.t === 'num') term *= cell.v;
      else if (cell.t === 'err') {
        e ??= cell.v;
        term = 0;
      } else term = 0;
    }
    total += term;
  });
  return e ? err(e) : num(total);
}

function concatFlatten(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  let out = '';
  for (const a of args) {
    const v = ev(a);
    if (v.t === 'ref') {
      const { r0, c0, r1, c1 } = v.rect;
      if ((r1 - r0 + 1) * (c1 - c0 + 1) > MAX_SCAN) return err('#VALUE!');
      for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
          const t = toText(refGet(v.sheet, ctx, r, c), ctx);
          if (typeof t !== 'string') return err(t);
          out += t;
          if (out.length > 32767) return err('#VALUE!');
        }
      }
    } else {
      const t = toText(v, ctx);
      if (typeof t !== 'string') return err(t);
      out += t;
    }
  }
  return str(out);
}

function textJoin(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 3) return err('#VALUE!');
  const delim = textArg(args, ev, ctx, 0);
  if (typeof delim !== 'string') return err(delim.v);
  const ignore = toBool(ev(args[1]!), ctx);
  if (typeof ignore === 'string') return err(ignore);
  const parts: Array<string> = [];
  for (let i = 2; i < args.length; i++) {
    const v = ev(args[i]!);
    if (v.t === 'ref') {
      const { r0, c0, r1, c1 } = v.rect;
      if ((r1 - r0 + 1) * (c1 - c0 + 1) > MAX_SCAN) return err('#VALUE!');
      for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
          const t = toText(refGet(v.sheet, ctx, r, c), ctx);
          if (typeof t !== 'string') return err(t);
          if (!(ignore && t === '')) parts.push(t);
        }
      }
    } else {
      const t = toText(v, ctx);
      if (typeof t !== 'string') return err(t);
      if (!(ignore && t === '')) parts.push(t);
    }
  }
  return str(parts.join(delim));
}

function substitute(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 3 || args.length > 4) return err('#VALUE!');
  const text = textArg(args, ev, ctx, 0);
  if (typeof text !== 'string') return err(text.v);
  const oldT = textArg(args, ev, ctx, 1);
  if (typeof oldT !== 'string') return err(oldT.v);
  const newT = textArg(args, ev, ctx, 2);
  if (typeof newT !== 'string') return err(newT.v);
  if (oldT === '') return str(text);
  if (args.length === 3) return str(text.split(oldT).join(newT));
  const inst = numArg(args, ev, ctx, 3);
  if (typeof inst !== 'number') return err(inst.e);
  const k = Math.trunc(inst);
  if (k < 1) return err('#VALUE!');
  let idx = -1;
  let count = 0;
  for (;;) {
    idx = text.indexOf(oldT, idx + 1);
    if (idx < 0) return str(text); // fewer than k occurrences ⇒ unchanged
    if (++count === k) return str(text.slice(0, idx) + newT + text.slice(idx + oldT.length));
  }
}

function replaceFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 4) return err('#VALUE!');
  const text = textArg(args, ev, ctx, 0);
  if (typeof text !== 'string') return err(text.v);
  const start = numArg(args, ev, ctx, 1);
  if (typeof start !== 'number') return err(start.e);
  const len = numArg(args, ev, ctx, 2);
  if (typeof len !== 'number') return err(len.e);
  const newT = textArg(args, ev, ctx, 3);
  if (typeof newT !== 'string') return err(newT.v);
  const s = Math.trunc(start);
  const l = Math.trunc(len);
  if (s < 1 || l < 0) return err('#VALUE!');
  return str(text.slice(0, s - 1) + newT + text.slice(s - 1 + l));
}

function rept(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const text = textArg(args, ev, ctx, 0);
  if (typeof text !== 'string') return err(text.v);
  const n = numArg(args, ev, ctx, 1);
  if (typeof n !== 'number') return err(n.e);
  const k = Math.trunc(n);
  if (k < 0) return err('#VALUE!');
  if (text.length * k > 32767) return err('#VALUE!'); // Excel's cell text cap
  return str(text.repeat(k));
}

// PROPER: capitalise the first letter of each run of letters, lowercase the rest.
function properCase(s: string): string {
  return s.replace(/[A-Za-z]+/g, (w) => w[0]!.toUpperCase() + w.slice(1).toLowerCase());
}

function charFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const n = numArg(args, ev, ctx, 0);
  if (typeof n !== 'number') return err(n.e);
  const code = Math.trunc(n);
  if (code < 1 || code > 255) return err('#VALUE!');
  return str(String.fromCharCode(code));
}

function codeFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  const t = textArg(args, ev, ctx, 0);
  if (typeof t !== 'string') return err(t.v);
  if (t.length === 0) return err('#VALUE!');
  return num(t.charCodeAt(0));
}

// CLEAN: drop the non-printing control characters (code < 32).
function stripControl(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) >= 32) out += s[i];
  return out;
}

function tFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const s = deref(ev(args[0]!), ctx);
  if (s.t === 'err') return s;
  return s.t === 'str' ? s : str('');
}

function daysFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const end = numArg(args, ev, ctx, 0);
  if (typeof end !== 'number') return err(end.e);
  const start = numArg(args, ev, ctx, 1);
  if (typeof start !== 'number') return err(start.e);
  return num(Math.trunc(end) - Math.trunc(start));
}

// HOUR/MINUTE/SECOND read the time-of-day from a serial's fractional part. A
// whole-day serial (the common cached value) has no time, so all three are 0;
// a TIME()-built fraction decodes back exactly.
function timePart(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  part: 'h' | 'm' | 's',
): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  if (a < 0) return err('#NUM!');
  let secs = Math.round((a - Math.floor(a)) * 86400);
  if (secs >= 86400) secs -= 86400;
  if (part === 'h') return num(Math.floor(secs / 3600));
  if (part === 'm') return num(Math.floor(secs / 60) % 60);
  return num(secs % 60);
}

function timeFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 3) return err('#VALUE!');
  const h = numArg(args, ev, ctx, 0);
  if (typeof h !== 'number') return err(h.e);
  const m = numArg(args, ev, ctx, 1);
  if (typeof m !== 'number') return err(m.e);
  const s = numArg(args, ev, ctx, 2);
  if (typeof s !== 'number') return err(s.e);
  const total = Math.trunc(h) * 3600 + Math.trunc(m) * 60 + Math.trunc(s);
  return num((((total % 86400) + 86400) % 86400) / 86400); // Excel wraps mod 24h
}

// The Excel WEEKNUM return_type → the weekday (0=Sun…6=Sat) that begins the week.
function weekStartDow(type: number): number | undefined {
  switch (type) {
    case 1:
    case 17:
      return 0; // week begins Sunday
    case 2:
    case 11:
      return 1; // week begins Monday
    case 12:
      return 2;
    case 13:
      return 3;
    case 14:
      return 4;
    case 15:
      return 5;
    case 16:
      return 6;
    default:
      return undefined;
  }
}

function weekNum(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, iso: boolean): FValue {
  if (args.length < 1 || args.length > 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  if (a < 0) return err('#NUM!');
  const serial = Math.floor(a);
  if (iso) return num(isoWeek(serial, ctx.date1904));
  let type = 1;
  if (args.length === 2) {
    const t = numArg(args, ev, ctx, 1);
    if (typeof t !== 'number') return err(t.e);
    type = Math.trunc(t);
  }
  if (type === 21) return num(isoWeek(serial, ctx.date1904));
  const firstDow = weekStartDow(type);
  if (firstDow === undefined) return err('#NUM!');
  const parts = serialToParts(serial, ctx.date1904);
  const jan1 = serialFromYmd(parts.year, 1, 1, ctx.date1904);
  const jan1Dow = serialToParts(jan1, ctx.date1904).dow;
  const offset = (jan1Dow - firstDow + 7) % 7; // days from the week-start to Jan 1
  const dayOfYear = serial - jan1 + 1; // 1-indexed
  return num(Math.floor((dayOfYear + offset - 1) / 7) + 1);
}

// ISO 8601 week number: weeks start Monday; week 1 holds the year's first
// Thursday. Computed from the Thursday of the target's week (which fixes the year).
function isoWeek(serial: number, date1904: boolean): number {
  const isoDow = (serialToParts(serial, date1904).dow + 6) % 7; // Mon=0…Sun=6
  const thursday = serial - isoDow + 3;
  const ty = serialToParts(thursday, date1904).year;
  const jan1 = serialFromYmd(ty, 1, 1, date1904);
  return Math.floor((thursday - jan1) / 7) + 1;
}

function chooseFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 2) return err('#VALUE!');
  const i = numArg(args, ev, ctx, 0);
  if (typeof i !== 'number') return err(i.e);
  const k = Math.trunc(i);
  if (k < 1 || k > args.length - 1) return err('#VALUE!');
  return ev(args[k]!);
}

function matchFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 2 || args.length > 3) return err('#VALUE!');
  const target = deref(ev(args[0]!), ctx);
  if (target.t === 'err') return target;
  const range = asRef(ev(args[1]!));
  if (!range) return err('#VALUE!');
  const rect = range.rect;
  let type = 1;
  if (args.length === 3) {
    const t = numArg(args, ev, ctx, 2);
    if (typeof t !== 'number') return err(t.e);
    type = Math.trunc(t);
  }
  const horizontal = rect.r0 === rect.r1;
  const vertical = rect.c0 === rect.c1;
  if (!horizontal && !vertical) return err('#N/A'); // MATCH wants a 1-D vector
  const len = horizontal ? rect.c1 - rect.c0 + 1 : rect.r1 - rect.r0 + 1;
  if (len > MAX_SCAN) return err('#N/A');
  const at = (i: number): Scalar =>
    horizontal
      ? refGet(range.sheet, ctx, rect.r0, rect.c0 + i)
      : refGet(range.sheet, ctx, rect.r0 + i, rect.c0);
  if (type === 0) {
    const re = target.t === 'str' ? wildcardToRegExp(target.v) : undefined;
    for (let i = 0; i < len; i++) {
      const cell = at(i);
      if (re ? cell.t === 'str' && re.test(cell.v) : scalarEquals(cell, target)) return num(i + 1);
    }
    return err('#N/A');
  }
  // Approximate: type 1 ⇒ largest ≤ target (ascending), -1 ⇒ smallest ≥ (descending).
  let best = -1;
  for (let i = 0; i < len; i++) {
    const cmp = cmpScalar(at(i), target);
    if (cmp === undefined) continue;
    if (type === 1 ? cmp <= 0 : cmp >= 0) best = i;
    else break;
  }
  return best < 0 ? err('#N/A') : num(best + 1);
}

function indexFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 2 || args.length > 3) return err('#VALUE!');
  const range = asRef(ev(args[0]!));
  if (!range) return err('#VALUE!');
  const rect = range.rect;
  const rn = numArg(args, ev, ctx, 1);
  if (typeof rn !== 'number') return err(rn.e);
  let r = Math.trunc(rn);
  let cn = 1;
  if (args.length === 3) {
    const c = numArg(args, ev, ctx, 2);
    if (typeof c !== 'number') return err(c.e);
    cn = Math.trunc(c);
  }
  const rows = rect.r1 - rect.r0 + 1;
  const cols = rect.c1 - rect.c0 + 1;
  // A single index into a one-row range selects the column.
  if (args.length === 2 && rows === 1 && cols > 1) {
    cn = r;
    r = 1;
  }
  if (r < 0 || cn < 0 || r > rows || cn > cols) return err('#REF!');
  if (r === 0 || cn === 0) return err('#VALUE!'); // whole row/column ⇒ an array
  return refGet(range.sheet, ctx, rect.r0 + r - 1, rect.c0 + cn - 1);
}

function lookupFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, vh: 'v' | 'h'): FValue {
  if (args.length < 3 || args.length > 4) return err('#VALUE!');
  const target = deref(ev(args[0]!), ctx);
  if (target.t === 'err') return target;
  const table = asRef(ev(args[1]!));
  if (!table) return err('#VALUE!');
  const rect = table.rect;
  const sh = table.sheet;
  const idx = numArg(args, ev, ctx, 2);
  if (typeof idx !== 'number') return err(idx.e);
  const line = Math.trunc(idx);
  let approx = true;
  if (args.length === 4) {
    const b = toBool(ev(args[3]!), ctx);
    if (typeof b === 'string') return err(b);
    approx = b;
  }
  const rows = rect.r1 - rect.r0 + 1;
  const cols = rect.c1 - rect.c0 + 1;
  const vertical = vh === 'v';
  const vecLen = vertical ? rows : cols;
  if (line < 1 || line > (vertical ? cols : rows)) return err('#REF!');
  if (vecLen > MAX_SCAN) return err('#N/A');
  const keyAt = (i: number): Scalar =>
    vertical ? refGet(sh, ctx, rect.r0 + i, rect.c0) : refGet(sh, ctx, rect.r0, rect.c0 + i);
  const resultAt = (i: number): Scalar =>
    vertical
      ? refGet(sh, ctx, rect.r0 + i, rect.c0 + line - 1)
      : refGet(sh, ctx, rect.r0 + line - 1, rect.c0 + i);
  if (!approx) {
    const re = target.t === 'str' ? wildcardToRegExp(target.v) : undefined;
    for (let i = 0; i < vecLen; i++) {
      const k = keyAt(i);
      if (re ? k.t === 'str' && re.test(k.v) : scalarEquals(k, target)) return resultAt(i);
    }
    return err('#N/A');
  }
  let best = -1;
  for (let i = 0; i < vecLen; i++) {
    const cmp = cmpScalar(keyAt(i), target);
    if (cmp === undefined) continue;
    if (cmp <= 0) best = i;
    else break;
  }
  return best < 0 ? err('#N/A') : resultAt(best);
}

// === Wave 4 helpers — broadening the library toward the full CF surface ======

// A two-argument numeric function (ATAN2, LOG-with-base style).
function math2(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  f: (a: number, b: number) => number,
): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const a = numArg(args, ev, ctx, 0);
  if (typeof a !== 'number') return err(a.e);
  const b = numArg(args, ev, ctx, 1);
  if (typeof b !== 'number') return err(b.e);
  const r = f(a, b);
  return Number.isNaN(r) ? err('#NUM!') : num(r);
}

// LOG(number, [base = 10]).
function logFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 1 || args.length > 2) return err('#VALUE!');
  const x = numArg(args, ev, ctx, 0);
  if (typeof x !== 'number') return err(x.e);
  let base = 10;
  if (args.length === 2) {
    const b = numArg(args, ev, ctx, 1);
    if (typeof b !== 'number') return err(b.e);
    base = b;
  }
  if (x <= 0 || base <= 0 || base === 1) return err('#NUM!');
  return num(Math.log(x) / Math.log(base));
}

// FACT(n) — n! over the truncated non-negative integer (capped to avoid overflow).
function factFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const n = numArg(args, ev, ctx, 0);
  if (typeof n !== 'number') return err(n.e);
  const k = Math.trunc(n);
  if (k < 0 || k > 170) return err('#NUM!');
  let r = 1;
  for (let i = 2; i <= k; i++) r *= i;
  return num(r);
}

// COMBIN(n, k) = C(n,k); PERMUT(n, k) = P(n,k) = n!/(n−k)!.
function combinPermut(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, perm: boolean): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const nN = numArg(args, ev, ctx, 0);
  if (typeof nN !== 'number') return err(nN.e);
  const kN = numArg(args, ev, ctx, 1);
  if (typeof kN !== 'number') return err(kN.e);
  const n = Math.trunc(nN);
  const k = Math.trunc(kN);
  if (n < 0 || k < 0 || k > n) return err('#NUM!');
  let p = 1;
  for (let i = 0; i < k; i++) p *= n - i;
  if (perm) return num(p);
  let kf = 1;
  for (let i = 2; i <= k; i++) kf *= i;
  return num(Math.round(p / kf));
}

// STDEV/VAR (sample → ÷(n−1), population → ÷n) and the deviation aggregates
// AVEDEV (mean |x−μ|) and DEVSQ (Σ(x−μ)²).
function statSpread(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  kind: 'stdev' | 'var' | 'avedev' | 'devsq',
  sample: boolean,
): FValue {
  const ns = gatherNumbers(args, ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  const n = ns.length;
  if (n === 0 || (sample && n < 2)) return err('#DIV/0!');
  const mean = ns.reduce((a, b) => a + b, 0) / n;
  if (kind === 'avedev') return num(ns.reduce((a, b) => a + Math.abs(b - mean), 0) / n);
  const ss = ns.reduce((a, b) => a + (b - mean) * (b - mean), 0);
  if (kind === 'devsq') return num(ss);
  const variance = ss / (sample ? n - 1 : n);
  return num(kind === 'var' ? variance : Math.sqrt(variance));
}

// GEOMEAN (all values must be > 0) / HARMEAN.
function geoHarMean(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext, geo: boolean): FValue {
  const ns = gatherNumbers(args, ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  if (ns.length === 0 || ns.some((x) => x <= 0)) return err('#NUM!');
  if (geo) return num(Math.exp(ns.reduce((a, b) => a + Math.log(b), 0) / ns.length));
  return num(ns.length / ns.reduce((a, b) => a + 1 / b, 0));
}

// AVERAGEA/MAXA/MINA — like AVERAGE/MAX/MIN but text counts as 0 and a logical as
// 1/0, so they enter the count / extent.
function aggregateA(
  args: ReadonlyArray<Ast>,
  ev: Ev,
  ctx: EvalContext,
  kind: 'avg' | 'max' | 'min',
): FValue {
  const ns: Array<number> = [];
  let e: FErr | undefined;
  const take = (s: Scalar): void => {
    if (s.t === 'num') ns.push(s.v);
    else if (s.t === 'bool') ns.push(s.v ? 1 : 0);
    else if (s.t === 'str') ns.push(0);
    else if (s.t === 'err') e ??= s.v;
  };
  for (const a of args) {
    const v = ev(a);
    if (v.t === 'ref') refEach(v, ctx, (_r, _c, s) => take(s));
    else take(deref(v, ctx));
  }
  if (e) return err(e);
  if (ns.length === 0) return kind === 'avg' ? err('#DIV/0!') : num(0);
  if (kind === 'avg') return num(ns.reduce((a, b) => a + b, 0) / ns.length);
  return num(kind === 'max' ? Math.max(...ns) : Math.min(...ns));
}

// MODE — the most frequent number, ties broken by earliest occurrence; #N/A when
// every value is unique.
function modeFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  const ns = gatherNumbers(args, ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  const counts = new Map<number, number>();
  for (const x of ns) counts.set(x, (counts.get(x) ?? 0) + 1);
  let maxCount = 1;
  for (const c of counts.values()) if (c > maxCount) maxCount = c;
  if (maxCount < 2) return err('#N/A');
  for (const x of ns) if (counts.get(x) === maxCount) return num(x);
  return err('#N/A');
}

// RANK(number, ref, [order]) — order 0/omitted ranks descending (largest = 1).
function rankFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length < 2 || args.length > 3) return err('#VALUE!');
  const x = numArg(args, ev, ctx, 0);
  if (typeof x !== 'number') return err(x.e);
  const ns = gatherNumbers([args[1]!], ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  let asc = false;
  if (args.length === 3) {
    const o = numArg(args, ev, ctx, 2);
    if (typeof o !== 'number') return err(o.e);
    asc = o !== 0;
  }
  if (!ns.includes(x)) return err('#N/A');
  let rank = 1;
  for (const v of ns) if (asc ? v < x : v > x) rank++;
  return num(rank);
}

// PERCENTILE(array, k∈[0,1]) — the inclusive (linear-interpolation) method.
function percentileFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const ns = gatherNumbers([args[0]!], ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  const k = numArg(args, ev, ctx, 1);
  if (typeof k !== 'number') return err(k.e);
  return percentile(ns, k);
}

// QUARTILE(array, q) — q 0..4 maps to the 0/25/50/75/100th percentile.
function quartileFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 2) return err('#VALUE!');
  const ns = gatherNumbers([args[0]!], ev, ctx);
  if (!Array.isArray(ns)) return err(ns);
  const q = numArg(args, ev, ctx, 1);
  if (typeof q !== 'number') return err(q.e);
  const qi = Math.trunc(q);
  if (qi < 0 || qi > 4) return err('#NUM!');
  return percentile(ns, qi / 4);
}

function percentile(ns: ReadonlyArray<number>, k: number): FValue {
  if (ns.length === 0 || k < 0 || k > 1) return err('#NUM!');
  const sorted = [...ns].sort((a, b) => a - b);
  const pos = k * (sorted.length - 1);
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return num(sorted[lo]! + (sorted[hi]! - sorted[lo]!) * (pos - lo));
}

// The reference an argument evaluates to (ROW/COLUMN/ROWS/COLUMNS), else undefined.
function rectArg(arg: Ast, ev: Ev): Rect | undefined {
  const v = ev(arg);
  return v.t === 'ref' ? v.rect : undefined;
}

// ROW([ref]) / COLUMN([ref]): with a reference, the top row/left column (1-based)
// of its rectangle; with no argument, the current cell from the per-cell shift.
function rowColFn(args: ReadonlyArray<Ast>, ev: Ev, shift: Shift, isRow: boolean): FValue {
  if (args.length === 0) {
    const cur = isRow ? shift.curRow : shift.curCol;
    return cur === undefined ? err('#VALUE!') : num(cur + 1);
  }
  if (args.length !== 1) return err('#VALUE!');
  const rect = rectArg(args[0]!, ev);
  return rect ? num((isRow ? rect.r0 : rect.c0) + 1) : err('#VALUE!');
}

// ROWS(ref) / COLUMNS(ref): the height / width of the reference's rectangle.
function rowsColsFn(args: ReadonlyArray<Ast>, ev: Ev, isRows: boolean): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const rect = rectArg(args[0]!, ev);
  if (!rect) return err('#VALUE!');
  return num(isRows ? rect.r1 - rect.r0 + 1 : rect.c1 - rect.c0 + 1);
}

// TYPE: 1 number (and blank), 2 text, 4 logical, 16 error.
function typeFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const s = deref(ev(args[0]!), ctx);
  if (s.t === 'str') return num(2);
  if (s.t === 'bool') return num(4);
  if (s.t === 'err') return num(16);
  return num(1);
}

// §18.17.2 ERROR.TYPE — the 1..7 code for an error value; #N/A for a non-error.
const ERROR_TYPE_CODES: Readonly<Record<FErr, number>> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
};
function errorTypeFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const s = deref(ev(args[0]!), ctx);
  return s.t === 'err' ? num(ERROR_TYPE_CODES[s.v]) : err('#N/A');
}

// UNICHAR(n) → the character for a Unicode code point; UNICODE(text) → the code
// point of the first character.
function unicharFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  if (args.length !== 1) return err('#VALUE!');
  const n = numArg(args, ev, ctx, 0);
  if (typeof n !== 'number') return err(n.e);
  const cp = Math.trunc(n);
  if (cp < 1 || cp > 0x10ffff) return err('#VALUE!');
  return str(String.fromCodePoint(cp));
}
function unicodeFn(args: ReadonlyArray<Ast>, ev: Ev, ctx: EvalContext): FValue {
  const t = textArg(args, ev, ctx, 0);
  if (typeof t !== 'string') return err(t.v);
  if (args.length !== 1 || t.length === 0) return err('#VALUE!');
  return num(t.codePointAt(0)!);
}
