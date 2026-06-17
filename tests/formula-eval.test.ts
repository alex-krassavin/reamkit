// E-SHEET W9 — the deterministic formula engine that powers conditional-format
// `expression` and `timePeriod` rules. These exercise the engine directly
// (lexer → parser → evaluator → function library + date helpers) against a
// hand-built grid context, phrasing every check as a boolean so it runs through
// the same evaluateToBool path a real CF rule uses. The engine reads only cached
// cell values and an injected reference day — no recalculation, no wall clock.

import { describe, expect, it } from 'vitest';

import type { EvalContext, Rect, Scalar, Shift } from '@/excel/formula';
import {
  BLANK,
  bool,
  compileFormula,
  evaluateToBool,
  num,
  serialFromDate,
  serialToParts,
  str,
  timePeriodMatches,
} from '@/excel/formula';

// "A1" / "$B$2" → [row, col] 0-indexed (ignores the $ anchors for the helper).
function a1(ref: string): [number, number] {
  const m = /^\$?([A-Z]+)\$?([0-9]+)$/.exec(ref)!;
  let col = 0;
  for (const ch of m[1]!) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]) - 1, col - 1];
}

function makeCtx(
  cells: Record<string, number | string | boolean | Scalar>,
  nowSerial?: number,
  date1904 = false,
): EvalContext {
  const COLS = 16384;
  const byKey = new Map<number, Scalar>();
  const list: Array<{ row: number; col: number; value: Scalar }> = [];
  for (const [ref, v] of Object.entries(cells)) {
    const [row, col] = a1(ref);
    const s: Scalar =
      typeof v === 'number'
        ? num(v)
        : typeof v === 'boolean'
          ? bool(v)
          : typeof v === 'string'
            ? str(v)
            : v;
    byKey.set(row * COLS + col, s);
    list.push({ row, col, value: s });
  }
  return {
    nowSerial,
    date1904,
    getCell: (row, col) => byKey.get(row * COLS + col) ?? BLANK,
    eachCell: (rect: Rect, visit) => {
      for (const c of list) {
        if (c.row >= rect.r0 && c.row <= rect.r1 && c.col >= rect.c0 && c.col <= rect.c1) {
          visit(c.row, c.col, c.value);
        }
      }
    },
  };
}

// Compile + evaluate a formula (sans leading '='), returning the boolean truth
// test. A formula that fails to parse compiles to undefined → false (no-op).
function ev(formula: string, ctx: EvalContext = makeCtx({}), shift?: Shift): boolean {
  const c = compileFormula(formula);
  return c ? evaluateToBool(c, ctx, shift) : false;
}

describe('formula engine — literals & operators (W9)', () => {
  it('evaluates arithmetic with Excel precedence', () => {
    expect(ev('1+2*3=7')).toBe(true);
    expect(ev('(1+2)*3=9')).toBe(true);
    expect(ev('2^3=8')).toBe(true);
    expect(ev('-2^2=4')).toBe(true); // unary minus binds tighter than ^ in Excel
    expect(ev('2^-2=0.25')).toBe(true);
    expect(ev('10%=0.1')).toBe(true);
    expect(ev('7-3-2=2')).toBe(true); // left-associative
  });

  it('compares numbers, text (case-insensitive) and across types', () => {
    expect(ev('1<2')).toBe(true);
    expect(ev('2<=2')).toBe(true);
    expect(ev('3<>4')).toBe(true);
    expect(ev('"abc"="ABC"')).toBe(true); // text compare folds case
    expect(ev('"apple"<"banana"')).toBe(true);
    expect(ev('5<"5"')).toBe(true); // number < text across categories
    expect(ev('"x"<TRUE')).toBe(true); // text < logical
  });

  it('concatenates with & and coerces operands to text', () => {
    expect(ev('"a"&"b"="ab"')).toBe(true);
    expect(ev('"x"&5="x5"')).toBe(true);
    expect(ev('1&2="12"')).toBe(true);
  });

  it('propagates errors to a non-match (the rule never paints)', () => {
    expect(ev('1/0')).toBe(false); // #DIV/0!
    expect(ev('NOPE()')).toBe(false); // #NAME?
    expect(ev('1/0=0')).toBe(false); // error in a comparison stays an error
    expect(ev('#REF!')).toBe(false);
  });

  it('treats a non-zero number or TRUE as a match; zero/blank/text as no match', () => {
    expect(ev('5')).toBe(true);
    expect(ev('0')).toBe(false);
    expect(ev('TRUE')).toBe(true);
    expect(ev('FALSE')).toBe(false);
    expect(ev('""')).toBe(false);
  });
});

describe('formula engine — references & relative shift (W9)', () => {
  const ctx = makeCtx({ A1: 3, A2: 10, B1: 'hi', C5: 7 });

  it('reads cell values and ranges', () => {
    expect(ev('A1=3', ctx)).toBe(true);
    expect(ev('A2>A1', ctx)).toBe(true);
    expect(ev('B1="HI"', ctx)).toBe(true);
    expect(ev('SUM(A1:A2)=13', ctx)).toBe(true);
  });

  it('shifts unanchored references by the cell offset from the rule origin', () => {
    // The formula "A1>5" written for the origin, evaluated at a cell two rows
    // down, reads A3 — but with the value at A2 via a one-row shift.
    expect(ev('A1>5', ctx, { dRow: 1, dCol: 0 })).toBe(true); // reads A2 = 10
    expect(ev('A1>5', ctx, { dRow: 0, dCol: 0 })).toBe(false); // reads A1 = 3
  });

  it('keeps $-anchored axes fixed under a shift', () => {
    expect(ev('$A$1=3', ctx, { dRow: 5, dCol: 5 })).toBe(true); // still A1
    expect(ev('$A1>5', ctx, { dRow: 1, dCol: 3 })).toBe(true); // col fixed, row shifts → A2
  });

  it('a shift off the grid is #REF! → no match', () => {
    expect(ev('A1', ctx, { dRow: -5, dCol: 0 })).toBe(false);
  });
});

describe('formula engine — function library (W9)', () => {
  const ctx = makeCtx({ A1: 1, A2: 2, A3: 2, A4: 3, B1: 'Hello', C1: -4 });

  it('logic: AND / OR / NOT / IF / IFERROR', () => {
    expect(ev('AND(1=1,2=2)', ctx)).toBe(true);
    expect(ev('AND(1=1,2=3)', ctx)).toBe(false);
    expect(ev('OR(1=2,3=3)', ctx)).toBe(true);
    expect(ev('NOT(1=2)', ctx)).toBe(true);
    expect(ev('IF(A1=1,"y","n")="y"', ctx)).toBe(true);
    expect(ev('IFERROR(1/0,42)=42', ctx)).toBe(true);
  });

  it('IS* predicates', () => {
    expect(ev('ISNUMBER(A1)', ctx)).toBe(true);
    expect(ev('ISTEXT(B1)', ctx)).toBe(true);
    expect(ev('ISBLANK(Z9)', ctx)).toBe(true);
    expect(ev('ISERROR(1/0)', ctx)).toBe(true);
    expect(ev('ISNUMBER(B1)', ctx)).toBe(false);
  });

  it('math: ABS / MOD / INT / ROUND family / POWER / SQRT', () => {
    expect(ev('ABS(C1)=4', ctx)).toBe(true);
    expect(ev('MOD(7,3)=1', ctx)).toBe(true);
    expect(ev('MOD(-1,3)=2', ctx)).toBe(true); // sign of divisor
    expect(ev('INT(2.9)=2', ctx)).toBe(true);
    expect(ev('ROUND(2.5,0)=3', ctx)).toBe(true); // half away from zero
    expect(ev('ROUND(-2.5,0)=-3', ctx)).toBe(true);
    expect(ev('ROUNDUP(2.1,0)=3', ctx)).toBe(true);
    expect(ev('ROUNDDOWN(2.9,0)=2', ctx)).toBe(true);
    expect(ev('POWER(2,10)=1024', ctx)).toBe(true);
    expect(ev('SQRT(9)=3', ctx)).toBe(true);
  });

  it('aggregates: SUM / AVERAGE / MIN / MAX / COUNT / COUNTA', () => {
    expect(ev('SUM(A1:A4)=8', ctx)).toBe(true);
    expect(ev('AVERAGE(A1:A4)=2', ctx)).toBe(true);
    expect(ev('MIN(A1:A4)=1', ctx)).toBe(true);
    expect(ev('MAX(A1:A4)=3', ctx)).toBe(true);
    expect(ev('COUNT(A1:B4)=4', ctx)).toBe(true); // A1:A4 numbers; B1 text → not counted
    expect(ev('COUNTA(A1:B4)=5', ctx)).toBe(true); // B1 (text) counted as non-blank
  });

  it('COUNTIF / SUMIF with comparison and wildcard criteria', () => {
    const c = makeCtx({ A1: 2, A2: 2, A3: 5, A4: 9, B1: 'apple', B2: 'apricot', B3: 'cherry' });
    expect(ev('COUNTIF(A1:A4,2)=2', c)).toBe(true);
    expect(ev('COUNTIF(A1:A4,">3")=2', c)).toBe(true);
    expect(ev('COUNTIF(B1:B3,"ap*")=2', c)).toBe(true);
    expect(ev('SUMIF(A1:A4,">3")=14', c)).toBe(true);
  });

  it('text: LEN / LEFT / RIGHT / MID / SEARCH / EXACT / CONCATENATE / VALUE', () => {
    expect(ev('LEN(B1)=5', ctx)).toBe(true);
    expect(ev('LEFT(B1,2)="He"', ctx)).toBe(true);
    expect(ev('RIGHT(B1,2)="lo"', ctx)).toBe(true);
    expect(ev('MID(B1,2,3)="ell"', ctx)).toBe(true);
    expect(ev('SEARCH("L",B1)=3', ctx)).toBe(true); // case-insensitive
    expect(ev('EXACT(B1,"hello")', ctx)).toBe(false); // case-sensitive
    expect(ev('CONCATENATE(B1," ",B1)="Hello Hello"', ctx)).toBe(true);
    expect(ev('VALUE("42")=42', ctx)).toBe(true);
  });
});

describe('formula engine — dates against an injected reference (W9)', () => {
  const today = serialFromDate(new Date(Date.UTC(2026, 5, 17)), false); // 2026-06-17
  const ctx = makeCtx({ A1: today }, today);

  it('TODAY/NOW read the injected serial; absent ⇒ no match', () => {
    expect(ev('TODAY()=A1', ctx)).toBe(true);
    expect(ev('NOW()=A1', ctx)).toBe(true);
    expect(ev('TODAY()=A1', makeCtx({ A1: today }))).toBe(false); // no nowSerial
  });

  it('YEAR / MONTH / DAY / WEEKDAY decompose a serial', () => {
    expect(ev('YEAR(A1)=2026', ctx)).toBe(true);
    expect(ev('MONTH(A1)=6', ctx)).toBe(true);
    expect(ev('DAY(A1)=17', ctx)).toBe(true);
    const dow = new Date(Date.UTC(2026, 5, 17)).getUTCDay(); // 0=Sun
    expect(ev(`WEEKDAY(A1)=${dow + 1}`, ctx)).toBe(true);
  });

  it('DATE / EDATE / EOMONTH build and shift serials', () => {
    expect(ev('DATE(2026,6,17)=A1', ctx)).toBe(true);
    expect(ev('YEAR(EDATE(A1,1))=2026', ctx)).toBe(true);
    expect(ev('MONTH(EDATE(A1,1))=7', ctx)).toBe(true);
    expect(ev('DAY(EOMONTH(A1,0))=30', ctx)).toBe(true); // June has 30 days
  });
});

describe('timePeriod windows (W9)', () => {
  const today = serialFromDate(new Date(Date.UTC(2026, 5, 17)), false);

  it('matches day-relative windows', () => {
    expect(timePeriodMatches('today', today, today, false)).toBe(true);
    expect(timePeriodMatches('yesterday', today - 1, today, false)).toBe(true);
    expect(timePeriodMatches('tomorrow', today + 1, today, false)).toBe(true);
    expect(timePeriodMatches('today', today - 1, today, false)).toBe(false);
  });

  it('matches the last-7-days window (today and the previous six)', () => {
    expect(timePeriodMatches('last7Days', today, today, false)).toBe(true);
    expect(timePeriodMatches('last7Days', today - 6, today, false)).toBe(true);
    expect(timePeriodMatches('last7Days', today - 7, today, false)).toBe(false);
    expect(timePeriodMatches('last7Days', today + 1, today, false)).toBe(false);
  });

  it('matches week windows (Sunday-start)', () => {
    const dow = serialToParts(today, false).dow;
    const weekStart = today - dow;
    expect(timePeriodMatches('thisWeek', weekStart, today, false)).toBe(true);
    expect(timePeriodMatches('thisWeek', weekStart + 6, today, false)).toBe(true);
    expect(timePeriodMatches('lastWeek', weekStart - 1, today, false)).toBe(true);
    expect(timePeriodMatches('nextWeek', weekStart + 7, today, false)).toBe(true);
  });

  it('matches month windows across a year boundary', () => {
    const dec = serialFromDate(new Date(Date.UTC(2025, 11, 15)), false);
    const jan = serialFromDate(new Date(Date.UTC(2026, 0, 15)), false);
    const feb = serialFromDate(new Date(Date.UTC(2026, 1, 15)), false);
    expect(timePeriodMatches('thisMonth', jan, jan, false)).toBe(true);
    expect(timePeriodMatches('lastMonth', dec, jan, false)).toBe(true); // Dec is last month of Jan
    expect(timePeriodMatches('nextMonth', feb, jan, false)).toBe(true);
    expect(timePeriodMatches('thisMonth', feb, jan, false)).toBe(false);
  });
});
