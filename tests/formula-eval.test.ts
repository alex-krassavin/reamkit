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

describe('formula engine — extended function library', () => {
  // A1:A5 = 1..5, B1:B5 = 10..50, C1:C3 = fruit; E1:H1 / E2:H2 a horizontal table.
  const g = makeCtx({
    A1: 1,
    A2: 2,
    A3: 3,
    A4: 4,
    A5: 5,
    B1: 10,
    B2: 20,
    B3: 30,
    B4: 40,
    B5: 50,
    C1: 'apple',
    C2: 'banana',
    C3: 'cherry',
    E1: 1,
    F1: 2,
    G1: 3,
    H1: 4,
    E2: 10,
    F2: 20,
    G2: 30,
    H2: 40,
  });

  it('extended logic / information: ISEVEN/ISODD/ISNONTEXT/XOR/IFS/SWITCH/IFNA/NA', () => {
    expect(ev('ISEVEN(4)', g)).toBe(true);
    expect(ev('ISODD(3)', g)).toBe(true);
    expect(ev('ISNONTEXT(5)', g)).toBe(true);
    expect(ev('ISNONTEXT("x")', g)).toBe(false);
    expect(ev('XOR(TRUE,FALSE)', g)).toBe(true);
    expect(ev('XOR(TRUE,TRUE)', g)).toBe(false);
    expect(ev('XOR(TRUE,TRUE,TRUE)', g)).toBe(true);
    expect(ev('IFS(FALSE,1,TRUE,2)=2', g)).toBe(true);
    expect(ev('ISNA(IFS(FALSE,1,FALSE,2))', g)).toBe(true); // no match ⇒ #N/A
    expect(ev('SWITCH(2,1,"a",2,"b","z")="b"', g)).toBe(true);
    expect(ev('SWITCH(9,1,"a","def")="def"', g)).toBe(true); // trailing default
    expect(ev('IFNA(NA(),5)=5', g)).toBe(true);
    expect(ev('IFNA(3,5)=3', g)).toBe(true);
  });

  it('extended math: PRODUCT/QUOTIENT/GCD/LCM/MROUND/EVEN/ODD/CEILING/FLOOR', () => {
    expect(ev('PRODUCT(A1:A3)=6', g)).toBe(true);
    expect(ev('QUOTIENT(7,2)=3', g)).toBe(true);
    expect(ev('ISERROR(QUOTIENT(1,0))', g)).toBe(true);
    expect(ev('GCD(12,18)=6', g)).toBe(true);
    expect(ev('LCM(4,6)=12', g)).toBe(true);
    expect(ev('MROUND(10,3)=9', g)).toBe(true);
    expect(ev('MROUND(13,2)=14', g)).toBe(true); // 6.5 rounds half away from zero
    expect(ev('EVEN(3)=4', g)).toBe(true);
    expect(ev('EVEN(-1)=-2', g)).toBe(true);
    expect(ev('ODD(2)=3', g)).toBe(true);
    expect(ev('ODD(-2)=-3', g)).toBe(true);
    expect(ev('CEILING(2.5,1)=3', g)).toBe(true);
    expect(ev('FLOOR(2.5,1)=2', g)).toBe(true);
    expect(ev('CEILING(-2.5,-1)=-3', g)).toBe(true);
    expect(ev('ISERROR(CEILING(-2.5,1))', g)).toBe(true); // mixed sign ⇒ #NUM!
  });

  it('statistics / IFS aggregates: MEDIAN/LARGE/SMALL/COUNTBLANK/SUMPRODUCT/COUNTIFS/SUMIFS/AVERAGEIF(S)', () => {
    expect(ev('MEDIAN(A1:A5)=3', g)).toBe(true);
    expect(ev('MEDIAN(A1:A4)=2.5', g)).toBe(true);
    expect(ev('LARGE(A1:A5,2)=4', g)).toBe(true);
    expect(ev('SMALL(A1:A5,2)=2', g)).toBe(true);
    expect(ev('COUNTBLANK(A1:A10)=5', g)).toBe(true); // A6:A10 are empty
    expect(ev('SUMPRODUCT(A1:A3,B1:B3)=140', g)).toBe(true); // 1·10+2·20+3·30
    expect(ev('COUNTIFS(A1:A5,">2",B1:B5,"<50")=2', g)).toBe(true);
    expect(ev('SUMIFS(B1:B5,A1:A5,">2")=120', g)).toBe(true);
    expect(ev('AVERAGEIF(A1:A5,">2")=4', g)).toBe(true);
    expect(ev('AVERAGEIFS(B1:B5,A1:A5,">2")=40', g)).toBe(true);
  });

  it('extended text: SUBSTITUTE/REPLACE/REPT/PROPER/CHAR/CODE/CLEAN/T/CONCAT/TEXTJOIN', () => {
    expect(ev('SUBSTITUTE("a-b-c","-","+")="a+b+c"', g)).toBe(true);
    expect(ev('SUBSTITUTE("a-b-c","-","+",2)="a-b+c"', g)).toBe(true);
    expect(ev('REPLACE("abcdef",2,3,"XY")="aXYef"', g)).toBe(true);
    expect(ev('REPT("ab",3)="ababab"', g)).toBe(true);
    expect(ev('PROPER("hello world")="Hello World"', g)).toBe(true);
    expect(ev('CHAR(65)="A"', g)).toBe(true);
    expect(ev('CODE("A")=65', g)).toBe(true);
    expect(ev(`CLEAN("a${String.fromCharCode(7)}b")="ab"`, g)).toBe(true); // strips a BEL
    expect(ev('T("hi")="hi"', g)).toBe(true);
    expect(ev('T(B1)=""', g)).toBe(true); // a number ⇒ empty text
    expect(ev('CONCAT(C1:C3)="applebananacherry"', g)).toBe(true);
    expect(ev('TEXTJOIN("-",TRUE,C1:C3)="apple-banana-cherry"', g)).toBe(true);
    expect(ev('TEXTJOIN(",",TRUE,C1,"",C2)="apple,banana"', g)).toBe(true); // skips empty
  });

  it('date / time: DAYS / TIME with HOUR-MINUTE-SECOND / WEEKNUM / ISOWEEKNUM', () => {
    expect(ev('DAYS(45,40)=5', g)).toBe(true);
    expect(ev('HOUR(TIME(13,30,15))=13', g)).toBe(true);
    expect(ev('MINUTE(TIME(13,30,15))=30', g)).toBe(true);
    expect(ev('SECOND(TIME(13,30,15))=15', g)).toBe(true);
    expect(ev('WEEKNUM(DATE(2023,1,1))=1', g)).toBe(true); // Jan 1 2023 is a Sunday
    expect(ev('WEEKNUM(DATE(2023,1,8))=2', g)).toBe(true);
    expect(ev('ISOWEEKNUM(DATE(2023,1,1))=52', g)).toBe(true); // ISO: belongs to 2022 W52
    expect(ev('ISOWEEKNUM(DATE(2023,1,2))=1', g)).toBe(true); // the Monday starts ISO W1
  });

  it('lookup: CHOOSE / MATCH / INDEX / VLOOKUP / HLOOKUP', () => {
    expect(ev('CHOOSE(2,"a","b","c")="b"', g)).toBe(true);
    expect(ev('MATCH(3,A1:A5,0)=3', g)).toBe(true); // exact
    expect(ev('MATCH(3.5,A1:A5,1)=3', g)).toBe(true); // largest ≤ 3.5 (ascending)
    expect(ev('MATCH("banana",C1:C3,0)=2', g)).toBe(true);
    expect(ev('INDEX(B1:B5,3)=30', g)).toBe(true);
    expect(ev('INDEX(A1:A5,4)=4', g)).toBe(true);
    expect(ev('VLOOKUP(3,A1:B5,2,FALSE)=30', g)).toBe(true);
    expect(ev('VLOOKUP(3.5,A1:B5,2,TRUE)=30', g)).toBe(true); // approximate
    expect(ev('ISNA(VLOOKUP(99,A1:B5,2,FALSE))', g)).toBe(true);
    expect(ev('HLOOKUP(3,E1:H2,2,FALSE)=30', g)).toBe(true);
  });

  it('an unsupported function still no-ops gracefully (#NAME?)', () => {
    expect(ev('XLOOKUP(1,A1:A5,B1:B5)=10', g)).toBe(false); // not in the library
  });
});

describe('formula engine — Wave 4 (trig / exp / extended math)', () => {
  it('evaluates the exponential and logarithm family', () => {
    expect(ev('ROUND(EXP(1),5)=2.71828')).toBe(true);
    expect(ev('LN(EXP(3))=3')).toBe(true);
    expect(ev('LOG10(1000)=3')).toBe(true);
    expect(ev('LOG(8,2)=3')).toBe(true);
    expect(ev('LOG(100)=2')).toBe(true); // default base 10
    expect(ev('LN(0)')).toBe(false); // domain error → #NUM! → no match
  });

  it('evaluates trig in radians plus DEGREES/RADIANS/PI', () => {
    expect(ev('SIN(0)=0')).toBe(true);
    expect(ev('ROUND(COS(0),5)=1')).toBe(true);
    expect(ev('ROUND(PI(),5)=3.14159')).toBe(true);
    expect(ev('DEGREES(PI())=180')).toBe(true);
    expect(ev('ROUND(RADIANS(180),5)=ROUND(PI(),5)')).toBe(true);
    expect(ev('ROUND(ATAN2(1,1),5)=ROUND(PI()/4,5)')).toBe(true);
    expect(ev('ASIN(2)')).toBe(false); // out of domain → #NUM!
  });

  it('evaluates FACT, COMBIN, PERMUT, SUMSQ', () => {
    expect(ev('FACT(5)=120')).toBe(true);
    expect(ev('COMBIN(5,2)=10')).toBe(true);
    expect(ev('PERMUT(5,2)=20')).toBe(true);
    expect(ev('SUMSQ(3,4)=25')).toBe(true);
    expect(ev('FACT(-1)')).toBe(false); // #NUM!
  });
});

describe('formula engine — Wave 4 (statistics)', () => {
  const g = makeCtx({ A1: 2, A2: 4, A3: 4, A4: 4, A5: 5, A6: 5, A7: 7, A8: 9 });

  it('computes sample/population spread', () => {
    expect(ev('STDEVP(A1:A8)=2', g)).toBe(true); // population stdev of the classic set
    expect(ev('VARP(A1:A8)=4', g)).toBe(true);
    expect(ev('ROUND(STDEV(A1:A8),4)=2.1381', g)).toBe(true); // sample stdev (÷ n−1)
    expect(ev('ROUND(VAR(A1:A8),4)=4.5714', g)).toBe(true);
    expect(ev('STDEV(5)')).toBe(false); // < 2 points → #DIV/0!
  });

  it('computes deviation, mean and modal aggregates', () => {
    expect(ev('DEVSQ(2,4,6)=8')).toBe(true);
    expect(ev('AVEDEV(2,4,6)*3=4')).toBe(true); // mean |dev| = 4/3, tested exactly
    expect(ev('GEOMEAN(4,9)=6')).toBe(true);
    expect(ev('HARMEAN(2,4,4)=3')).toBe(true);
    expect(ev('MODE(A1:A8)=4', g)).toBe(true);
    expect(ev('ISNA(MODE(1,2,3))')).toBe(true); // all unique → #N/A
  });

  it('AVERAGEA/MAXA/MINA count text as 0 and RANK orders a value', () => {
    const t = makeCtx({ A1: 4, A2: 'x', A3: 8 });
    expect(ev('AVERAGEA(A1:A3)=4', t)).toBe(true); // (4 + 0 + 8) / 3
    expect(ev('MINA(A1:A3)=0', t)).toBe(true);
    expect(ev('MAXA(A1:A3)=8', t)).toBe(true);
    expect(ev('RANK(7,A1:A8)=2', g)).toBe(true); // 7 is the 2nd-largest
    expect(ev('RANK(7,A1:A8,1)=7', g)).toBe(true); // ascending order
  });

  it('PERCENTILE/QUARTILE interpolate', () => {
    const p = makeCtx({ A1: 1, A2: 2, A3: 3, A4: 4 });
    expect(ev('PERCENTILE(A1:A4,0.5)=2.5', p)).toBe(true);
    expect(ev('QUARTILE(A1:A4,2)=2.5', p)).toBe(true);
    expect(ev('QUARTILE(A1:A4,0)=1', p)).toBe(true);
  });
});

describe('formula engine — Wave 4 (reference / information)', () => {
  const g = makeCtx({ A1: 1, A2: 2, A3: 3 });
  const at = (row: number, col: number): Shift => ({ dRow: 0, dCol: 0, curRow: row, curCol: col });

  it('ROW/COLUMN report the reference or the current cell', () => {
    expect(ev('ROW(A5)=5', g)).toBe(true);
    expect(ev('COLUMN(C1)=3', g)).toBe(true);
    expect(ev('ROW()=8', g, at(7, 2))).toBe(true); // current cell C8
    expect(ev('COLUMN()=3', g, at(7, 2))).toBe(true);
    expect(ev('MOD(ROW(),2)=0', g, at(3, 0))).toBe(true); // row-banding idiom (row 4)
  });

  it('ROWS/COLUMNS measure a range', () => {
    expect(ev('ROWS(A1:A3)=3', g)).toBe(true);
    expect(ev('COLUMNS(A1:C1)=3', g)).toBe(true);
  });

  it('TYPE / ERROR.TYPE classify a value', () => {
    expect(ev('TYPE(1)=1')).toBe(true);
    expect(ev('TYPE("x")=2')).toBe(true);
    expect(ev('TYPE(TRUE)=4')).toBe(true);
    expect(ev('TYPE(1/0)=16')).toBe(true);
    expect(ev('ERROR.TYPE(1/0)=2')).toBe(true);
    expect(ev('ISNA(ERROR.TYPE(5))')).toBe(true); // not an error → #N/A
  });

  it('UNICHAR / UNICODE round-trip a code point', () => {
    expect(ev('UNICHAR(65)="A"')).toBe(true);
    expect(ev('UNICODE("A")=65')).toBe(true);
    expect(ev('UNICODE(UNICHAR(8364))=8364')).toBe(true); // €
  });
});

describe('formula engine — Wave 4 (array constants)', () => {
  it('flattens an array constant into an aggregate', () => {
    expect(ev('SUM({1,2,3})=6')).toBe(true);
    expect(ev('SUM({1,2;3,4})=10')).toBe(true); // 2×2
    expect(ev('SUM({-1,2,-3})=-2')).toBe(true); // signed elements
    expect(ev('COUNT({1,2,3,4})=4')).toBe(true);
    expect(ev('MAX({3,9,5})=9')).toBe(true);
  });

  it('broadcasts an operator element-wise against a scalar', () => {
    expect(ev('SUM({1,2,3}*2)=12')).toBe(true);
    expect(ev('SUM({1,2,3}+10)=36')).toBe(true);
  });

  it('supports the OR(cell={…}) membership idiom', () => {
    const g = makeCtx({ A1: 3 });
    expect(ev('OR(A1={1,3,5})', g)).toBe(true); // 3 is in the set
    expect(ev('OR(A1={2,4,6})', g)).toBe(false); // 3 is not
    expect(ev('OR(5={5,6,7})')).toBe(true);
    expect(ev('AND({1,1,1})')).toBe(true);
    expect(ev('AND({1,0,1})')).toBe(false);
  });

  it('collapses to the top-left element in a scalar context', () => {
    expect(ev('{9,8,7}+0=9')).toBe(true); // row array → first element
    expect(ev('{5;6;7}+0=5')).toBe(true); // column array → first element
    expect(ev('IF({1,0},"y","n")="y"')).toBe(true); // condition uses the top-left
  });
});
