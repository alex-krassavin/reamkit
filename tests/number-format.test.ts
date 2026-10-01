import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';
import { applyNumberFormat, generalToWidth, numberFormatColorHex } from '@/excel';

const noCustom = new Map<number, string>();

describe('applyNumberFormat — numbers', () => {
  it('preserves integers under General (numFmtId 0)', () => {
    expect(applyNumberFormat('42', 0, noCustom)).toBe('42');
    expect(applyNumberFormat('-13', 0, noCustom)).toBe('-13');
  });

  it('formats with thousands separator (built-in 3 = #,##0)', () => {
    expect(applyNumberFormat('1234567', 3, noCustom)).toBe('1,234,567');
    expect(applyNumberFormat('500', 3, noCustom)).toBe('500');
    expect(applyNumberFormat('-1234567', 3, noCustom)).toBe('-1,234,567');
  });

  it('formats with two decimals (built-in 4 = #,##0.00)', () => {
    expect(applyNumberFormat('1234.5', 4, noCustom)).toBe('1,234.50');
    expect(applyNumberFormat('1234567.891', 4, noCustom)).toBe('1,234,567.89');
  });

  it('formats as percent (built-in 9 = 0%)', () => {
    expect(applyNumberFormat('0.42', 9, noCustom)).toBe('42%');
    expect(applyNumberFormat('1', 9, noCustom)).toBe('100%');
  });

  it('passes text-typed cells (numFmtId 49 = @) through verbatim', () => {
    expect(applyNumberFormat('hello', 49, noCustom)).toBe('hello');
  });

  it('applies custom number format with literal text', () => {
    const fmt = new Map<number, string>([[164, '"$"#,##0.00']]);
    expect(applyNumberFormat('1234.5', 164, fmt)).toBe('$1,234.50');
  });

  it('uses the negative section for negative values (built-in 40), no leak', () => {
    // #,##0.00_);[Red](#,##0.00) — negatives render in parentheses.
    expect(applyNumberFormat('-1234.5', 40, noCustom)).toBe('(1,234.50)');
    expect(applyNumberFormat('1234.5', 40, noCustom)).toBe('1,234.50 ');
  });

  it('pads a short integer part to its 0 placeholders with zeros', () => {
    const fmt = new Map<number, string>([
      [164, '000'],
      [165, '0,000'],
    ]);
    expect(applyNumberFormat('5', 164, fmt)).toBe('005');
    expect(applyNumberFormat('1234', 164, fmt)).toBe('1234');
    expect(applyNumberFormat('5', 165, fmt)).toBe('0,005');
  });

  it('shows a space for each insignificant zero under a ? placeholder', () => {
    // §18.8.31 — `?` lines up decimal points: it is `0` for a digit that
    // counts and a space for one that does not. Excel's own Accounting format
    // writes its zero section as `"-"??`, which reads " -  ", not "-??".
    const fmt = new Map<number, string>([
      [164, '0.0??'],
      [165, '??'],
      [166, '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-'],
    ]);
    expect(applyNumberFormat('1.5', 164, fmt)).toBe('1.5  ');
    expect(applyNumberFormat('1.234', 164, fmt)).toBe('1.234');
    expect(applyNumberFormat('5', 165, fmt)).toBe(' 5');
    expect(applyNumberFormat('0', 166, fmt)).toBe(' -   ');
    expect(applyNumberFormat('1234.5', 166, fmt)).toBe(' 1,234.50 ');
  });
});

describe('applyNumberFormat — scientific notation (§18.8.31)', () => {
  it('formats built-in 11 (0.00E+00)', () => {
    expect(applyNumberFormat('12345.678', 11, noCustom)).toBe('1.23E+04');
    expect(applyNumberFormat('0.000678', 11, noCustom)).toBe('6.78E-04');
    expect(applyNumberFormat('0', 11, noCustom)).toBe('0.00E+00');
  });

  it('keeps the sign for negative mantissa', () => {
    expect(applyNumberFormat('-12345.678', 11, noCustom)).toBe('-1.23E+04');
  });

  it('groups the exponent to multiples of intDigits (engineering ##0.0E+0)', () => {
    const fmt = new Map<number, string>([[165, '##0.0E+0']]);
    expect(applyNumberFormat('12345', 165, fmt)).toBe('12.3E+3');
    expect(applyNumberFormat('1234567', 165, fmt)).toBe('1.2E+6');
  });

  it('preserves lowercase e and renormalises a carried mantissa', () => {
    const fmt = new Map<number, string>([[166, '0.0e+00']]);
    expect(applyNumberFormat('999', 166, fmt)).toBe('1.0e+03'); // 9.99e2 → 1.0e3
  });
});

describe('applyNumberFormat — dates', () => {
  // 2024-01-01 in Excel serial — 1899-12-30 + 45292 days.
  const JAN_1_2024 = '45292';
  const JAN_15_2024_NOON = String(45292 + 14 + 0.5); // 2024-01-15 12:00 UTC

  it('formats built-in 14 (m/d/yyyy)', () => {
    expect(applyNumberFormat(JAN_1_2024, 14, noCustom)).toBe('1/1/2024');
  });

  it('formats built-in 15 (d-mmm-yy)', () => {
    expect(applyNumberFormat(JAN_1_2024, 15, noCustom)).toBe('1-Jan-24');
  });

  it('formats built-in 22 (m/d/yyyy h:mm) preserving the time portion', () => {
    expect(applyNumberFormat(JAN_15_2024_NOON, 22, noCustom)).toBe('1/15/2024 12:00');
  });

  it('formats built-in 20 (h:mm)', () => {
    expect(applyNumberFormat(JAN_15_2024_NOON, 20, noCustom)).toBe('12:00');
  });

  it('respects AM/PM marker (built-in 18 = h:mm AM/PM)', () => {
    expect(applyNumberFormat(JAN_15_2024_NOON, 18, noCustom)).toBe('12:00 PM');
    const sevenAm = String(45292 + 14 + 7 / 24);
    expect(applyNumberFormat(sevenAm, 18, noCustom)).toBe('7:00 AM');
  });

  it('applies a custom date format dd.mm.yyyy', () => {
    const fmt = new Map<number, string>([[170, 'dd.mm.yyyy']]);
    expect(applyNumberFormat(JAN_1_2024, 170, fmt)).toBe('01.01.2024');
  });

  it('applies a custom format with month name (d mmmm yyyy)', () => {
    const fmt = new Map<number, string>([[171, 'd mmmm yyyy']]);
    expect(applyNumberFormat(JAN_1_2024, 171, fmt)).toBe('1 January 2024');
  });

  it('disambiguates m as minutes when adjacent to h: / :s', () => {
    const fmt = new Map<number, string>([[172, 'hh:mm:ss']]);
    expect(applyNumberFormat(JAN_15_2024_NOON, 172, fmt)).toBe('12:00:00');
  });

  it('disambiguates m as month when standalone in a date context', () => {
    const fmt = new Map<number, string>([[173, 'm/d']]);
    expect(applyNumberFormat(JAN_1_2024, 173, fmt)).toBe('1/1');
  });

  it('applies the 1904 epoch when date1904 is true', () => {
    // 1904 mode: serial 0 = 1904-01-01.
    expect(applyNumberFormat('0', 14, noCustom, true)).toBe('1/1/1904');
    // The 2024 calendar date sits 1462 days earlier in 1904-mode storage.
    expect(applyNumberFormat(String(45292 - 1462), 14, noCustom, true)).toBe('1/1/2024');
  });

  it('defaults to 1900 epoch when date1904 is omitted', () => {
    expect(applyNumberFormat('0', 14, noCustom)).toBe('12/30/1899');
    expect(applyNumberFormat('45292', 14, noCustom)).toBe('1/1/2024');
  });
});

describe('applyNumberFormat — codes real producers write', () => {
  it('reads the format name in any case (bug-fixes.xlsx writes GENERAL)', () => {
    // Matched exactly, "GENERAL" carries no digit placeholder, so the whole
    // word became a literal prefix and every cell read "GENERAL1", "GENERAL2".
    const fmt = new Map<number, string>([[164, 'GENERAL']]);
    expect(applyNumberFormat('1', 164, fmt)).toBe('1');
    expect(applyNumberFormat('0.5', 164, fmt)).toBe('0.5');
  });

  it('blanks a zero written only with # (tdf171828 hides a column that way)', () => {
    // §18.8.31: `#` is an optional digit. A `0` or `?` anywhere in the integer
    // part forces it back.
    const fmt = new Map<number, string>([
      [165, '#'],
      [166, '#.##'],
      [167, '#.00'],
    ]);
    expect(applyNumberFormat('0', 165, fmt)).toBe('');
    expect(applyNumberFormat('7', 165, fmt)).toBe('7');
    expect(applyNumberFormat('0.5', 166, fmt)).toBe('.5');
    expect(applyNumberFormat('0', 167, fmt)).toBe('.00');
    expect(applyNumberFormat('0', 3, noCustom)).toBe('0');
  });

  it('counts elapsed time in [h] / [mm] / [ss] instead of wrapping', () => {
    const fmt = new Map<number, string>([
      [165, '[ss].00'],
      [166, '[mm]:ss'],
      [167, '[h]:mm'],
    ]);
    // seconds-without-truncate-and-decimals.xlsx: 3.14159270833 days.
    expect(applyNumberFormat('3.14159270833', 165, fmt)).toBe('271433.61');
    expect(applyNumberFormat('0.0424', 166, fmt)).toBe('61:03');
    expect(applyNumberFormat('3.14159270833', 167, fmt)).toBe('75:23');
  });

  it('renders the currency symbol out of a [$SYMBOL-LOCALE] tag', () => {
    // Dropped with the bracket — as a colour or a locale is — `[$$-409]#,##0`
    // silently lost its "$" and formats.xlsx printed a bare 12,345.00.
    const fmt = new Map<number, string>([
      [164, '[$$-409]#,##0;[RED]\\-[$$-409]#,##0'],
      [165, '#,##0.00\\ [$USD]'],
      [166, '[$-409]#,##0'],
    ]);
    expect(applyNumberFormat('12345', 164, fmt)).toBe('$12,345');
    expect(applyNumberFormat('-1234', 164, fmt)).toBe('-$1,234');
    expect(applyNumberFormat('1234.5', 165, fmt)).toBe('1,234.50 USD');
    // A locale with no symbol contributes nothing.
    expect(applyNumberFormat('12345', 166, fmt)).toBe('12,345');
  });

  it('approximates fractions to the denominator the format allows', () => {
    const fmt = new Map<number, string>([
      [164, '# ??/??'],
      [165, '# ?/?'],
      [166, '?/?'],
      [167, '# ?/16'],
    ]);
    expect(applyNumberFormat('25.378', 164, fmt)).toBe('25 31/82');
    expect(applyNumberFormat('2.55', 165, fmt)).toBe('2 5/9');
    expect(applyNumberFormat('0.3889', 166, fmt)).toBe('2/5');
    expect(applyNumberFormat('3.3', 167, fmt)).toBe('3 5/16');
    // A whole number keeps its integer and drops the fraction entirely.
    expect(applyNumberFormat('4', 164, fmt)).toBe('4');
  });

  it('knows the built-in currency and fraction ids (§18.8.30)', () => {
    // A gap in the table is invisible: the id resolves to no code at all, the
    // cell falls through to the General rendering, and 49273.xlsx printed 0.5
    // where its `# ?/?` cell reads " 1/2". The engines were all there.
    expect(applyNumberFormat('1234.5', 5, noCustom)).toBe('$1,235 ');
    expect(applyNumberFormat('-1234.5', 6, noCustom)).toBe('($1,235)');
    expect(applyNumberFormat('1234.5', 7, noCustom)).toBe('$1,234.50 ');
    expect(applyNumberFormat('-1234.5', 8, noCustom)).toBe('($1,234.50)');
    expect(numberFormatColorHex('-1234.5', 8, noCustom)).toBe('FF0000');
    expect(applyNumberFormat('0.5', 12, noCustom)).toBe('1/2');
    expect(applyNumberFormat('2.25', 13, noCustom)).toBe('2 1/4');
    expect(applyNumberFormat('12345', 48, noCustom)).toBe('12.3E+3');
  });

  it('reads General through the brackets in front of it', () => {
    // `[DBNum1][$-804]General` is General with a numeral system and a locale on
    // it. Tested against the whole code, it missed — and the placeholder
    // grammar, finding no digit placeholder, printed the keyword: 49273.xlsx
    // read "General12323". Choosing the numerals is a separate feature; this is
    // about not writing the word out.
    const fmt = new Map<number, string>([
      [176, '[DBNum1][$-804]General'],
      [177, '[$-409]GENERAL'],
    ]);
    expect(applyNumberFormat('12323', 176, fmt)).toBe('12323');
    expect(applyNumberFormat('0.5', 177, fmt)).toBe('0.5');
  });

  it('reports the colour the applying section names', () => {
    const fmt = new Map<number, string>([[164, '[$$-409]#,##0;[RED]\\-[$$-409]#,##0']]);
    expect(numberFormatColorHex('12345', 164, fmt)).toBeUndefined();
    expect(numberFormatColorHex('-1234', 164, fmt)).toBe('FF0000');
    // Built-in 40 is #,##0.00_);[Red](#,##0.00).
    expect(numberFormatColorHex('-5', 40, noCustom)).toBe('FF0000');
    expect(numberFormatColorHex('5', 40, noCustom)).toBeUndefined();
    // [Color n] indexes the §18.8.27 palette, numbered from 1 here.
    expect(numberFormatColorHex('1', 164, new Map([[164, '[Color 5]0']]))).toBe('0000FF');
  });

  it('separates the whole number from a fraction with a QUOTED literal too', () => {
    // The integer part is the last placeholder run before the fraction, and
    // what separates them is a literal that need not be a bare space:
    // formats.xlsx writes `#"  "?/?`. Anchored at the end of the head, the run
    // went unfound there — so 1.2 came out as the improper `#  6/5`, the `#`
    // printed as itself and the whole number folded into the numerator.
    const fmt = new Map<number, string>([
      [300, '#"  "?/?'],
      [301, '# ?/?'],
      [302, '?/?'],
    ]);
    expect(applyNumberFormat('1.2', 300, fmt)).toBe('1  1/5');
    expect(applyNumberFormat('-1.2', 300, fmt)).toBe('-1  1/5');
    expect(applyNumberFormat('0.75', 300, fmt)).toBe('3/4');
    // The plain spellings are unchanged, improper form included.
    expect(applyNumberFormat('1.2', 301, fmt)).toBe('1 1/5');
    expect(applyNumberFormat('1.2', 302, fmt)).toBe('6/5');
  });

  it('rounds a General number to the decimals its column has room for', () => {
    // General is not a fixed format: a spreadsheet shows as many decimal places
    // as fit and ROUNDS to that. Rendering every stored digit and letting the
    // cell clip it turns 4.3900875881221957 into "4.390087" where every other
    // reader shows "4.390088" — off by one in the last place shown, with
    // nothing to say a digit was cut (Sparklines.xlsx).
    const upTo =
      (n: number) =>
      (t: string): boolean =>
        t.length <= n;
    expect(generalToWidth('4.3900875881221957', upTo(8))).toBe('4.390088');
    expect(generalToWidth('-6.1052278206732389', upTo(9))).toBe('-6.105228');
    // Room for everything is room for General's eleven characters: Excel's own
    // PDF writes 4.390087588 in a column forty wide.
    expect(generalToWidth('4.3900875881221957', upTo(30))).toBe('4.390087588');
    // An integer that will not fit goes scientific rather than reading as a
    // number a thousand times smaller (escape-unicode.xlsx printed "1161014"
    // for 1161014163).
    expect(generalToWidth('1161014163', upTo(9))).toBe('1.161E+09');
    expect(generalToWidth('1161014163', upTo(30))).toBe('1161014163');
    // Too narrow even for that: the cell says so its own way, with hashes.
    expect(generalToWidth('1161014163', upTo(3))).toBe('1161014163');
  });

  it('writes General in a cell as Excel does: eleven characters, no zeros at the end', () => {
    // Each a cell of Excel's own PDF (2026-10-02), its column's room counted in
    // characters. However wide the column, General writes eleven of them, its
    // sign apart: we wrote every stored digit.
    const upTo =
      (n: number) =>
      (t: string): boolean =>
        t.length <= n;
    const wide: ReadonlyArray<readonly [string, string]> = [
      ['10000000000', '10000000000'],
      ['-10000000000', '-10000000000'],
      ['2500000000000', '2.5E+12'],
      ['123456789012345', '1.23457E+14'],
      ['0.123456789012345', '0.123456789'],
      ['0.30000000000000004', '0.3'],
      ['1.2345E-05', '0.000012345'],
      ['1E-07', '0.0000001'],
      ['3.5E-07', '0.00000035'],
      ['-1.2345678E-05', '-1.23457E-05'],
      ['1.23456789E-10', '1.23457E-10'],
      ['1E+21', '1E+21'],
    ];
    expect(wide.map(([v]) => generalToWidth(v, upTo(99)))).toEqual(wide.map(([, t]) => t));
    // Narrower, the finest that fits: rounded, or scientific where that keeps
    // a finer last digit, never with the zeros a mantissa ends in — we wrote
    // 1.000E+10 — and rounded on the decimal (1.2345E-05 is 1.234…E-05 in
    // binary, and toExponential wrote 1.234E-05).
    const narrow: ReadonlyArray<readonly [string, number, string]> = [
      ['10000000000', 9, '1E+10'],
      ['1234567890', 9, '1.235E+09'],
      ['1.2345E-05', 9, '1.235E-05'],
      ['3.5E-07', 9, '3.5E-07'],
      ['2500000000000', 7, '2.5E+12'],
      ['99999999999', 7, '1E+11'],
      ['1E-07', 7, '1E-07'],
      ['2500000000000', 6, '3E+12'],
      ['0.0004', 5, '4E-04'],
      ['0.00123456', 5, '0.001'],
      ['123456', 5, '1E+05'],
      ['99.96', 4, '100'],
      ['1.2345E-05', 4, '0'],
      ['-1.2345678E-05', 5, '-0'],
    ];
    expect(narrow.map(([v, room]) => generalToWidth(v, upTo(room)))).toEqual(
      narrow.map(([, , t]) => t),
    );
    // Of two as fine, the plain one down to 0.0001 and the scientific below it.
    expect(generalToWidth('0.000123456', upTo(6))).toBe('0.0001');
    expect(generalToWidth('1.00049E-05', upTo(9))).toBe('1E-05');
    expect(generalToWidth('0.000099996', upTo(9))).toBe('1E-04');
    // …unless the number fits as it is.
    expect(generalToWidth('1E-05', upTo(9))).toBe('0.00001');
  });

  it('rounds on the decimal number, half away from zero', () => {
    // `toFixed` rounds the BINARY double, which is not the number the file
    // means. A rate stored as 0.0095 is 0.00949999999999999… in binary; times
    // 100 that is 0.9499999999999998 and toFixed(1) answers "0.9" where every
    // other reader shows 1.0%. Eleven of AverageTaxRates.xlsx's percentages
    // were a tenth low for exactly this reason.
    const fmt = new Map<number, string>([
      [164, '0.0%'],
      [165, '0.00'],
      [166, '#,##0'],
    ]);
    expect(applyNumberFormat('0.0095', 164, fmt)).toBe('1.0%');
    expect(applyNumberFormat('-0.0095', 164, fmt)).toBe('-1.0%');
    expect(applyNumberFormat('0.18449999999999999', 164, fmt)).toBe('18.5%');
    // The classic binary traps: both of these round DOWN under toFixed.
    expect(applyNumberFormat('1.005', 165, fmt)).toBe('1.01');
    expect(applyNumberFormat('2.675', 165, fmt)).toBe('2.68');
    // Half away from zero, not towards +∞ — and not to even either.
    expect(applyNumberFormat('-1.005', 165, fmt)).toBe('-1.01');
    expect(applyNumberFormat('0.125', 165, fmt)).toBe('0.13');
    expect(applyNumberFormat('1234.5', 166, fmt)).toBe('1,235');
  });

  it('decodes a section that is all literal (the Accounting zero)', () => {
    // `_(* "-"_)` carries no digit placeholder, so the grammar took a shortcut
    // and returned the section RAW — a balance row that should read "-" read
    // `_(* "-"_)`, the format code printing itself (49156.xlsx).
    const fmt = new Map<number, string>([[164, '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)']]);
    expect(applyNumberFormat('0', 164, fmt).trim()).toBe('-');
    expect(applyNumberFormat('1234', 164, fmt).trim()).toBe('1,234');
    expect(applyNumberFormat('-1234', 164, fmt).trim()).toBe('(1,234)');
  });

  it('reads .0 after a seconds token as its decimals (built-in 47)', () => {
    // Read as a literal dot and a literal zero, mm:ss.0 printed ".0" for every
    // value. 0.5208333 days = 12:30:00 to a tenth of a second.
    expect(applyNumberFormat('0.5208333', 47, noCustom)).toBe('30:00.0');
  });
});

describe('General spells an exponent Excel’s way, not JavaScript’s', () => {
  it('gives an upper-case E, an explicit sign and two exponent digits', () => {
    // General bottoms out in ECMAScript's Number-to-String, which writes
    // `3e-104`. §18.8.31 spells scientific notation `E+`/`E-` — built-in 11 is
    // `0.00E+00` — and every other reader prints `3E-104`. 57236.xlsx stores
    // three such values and we printed all three with a lower-case e.
    expect(applyNumberFormat('3.0000000000000002E-104', 0, noCustom)).toBe('3E-104');
    expect(applyNumberFormat('4.9999999999999998E-106', 0, noCustom)).toBe('5E-106');
    expect(applyNumberFormat('1e21', 0, noCustom)).toBe('1E+21');
    // A single-digit exponent is padded, as the built-in codes are.
    expect(applyNumberFormat('0.0000001', 0, noCustom)).toBe('1E-07');
    // Everything that is not exponential is untouched.
    expect(applyNumberFormat('42', 0, noCustom)).toBe('42');
    expect(applyNumberFormat('0.5', 0, noCustom)).toBe('0.5');
  });
});

describe('a comma against the last placeholder is a scale, not punctuation', () => {
  it('shows the value in thousands, one thousand per comma (§18.8.31)', () => {
    const custom = new Map<number, string>([[164, '#,##0,,']]);
    // bug69812.xlsx: one cell, 25 396 277 490 under `#,##0,,`. Both references
    // print "25,396"; the commas printed themselves.
    expect(applyNumberFormat('25396277490', 164, custom)).toBe('25,396');
    expect(applyNumberFormat('25396277490', 164, new Map([[164, '#,##0,']]))).toBe('25,396,277');
    expect(applyNumberFormat('12345678', 164, new Map([[164, '#,##0.0,']]))).toBe('12,345.7');
    // A suffix after the scale still prints.
    expect(applyNumberFormat('25396277490', 164, new Map([[164, '0.0,,"M"']]))).toBe('25396.3M');
  });

  it('leaves a comma that does not touch a placeholder alone', () => {
    // Quoted, so a literal — and the grouping comma inside the digits is not a
    // scale either.
    expect(applyNumberFormat('1234', 164, new Map([[164, '#,##0","']]))).toBe('1,234,');
    expect(applyNumberFormat('1234567', 164, new Map([[164, '#,##0']]))).toBe('1,234,567');
  });
});

describe('a section may state a CONDITION, and then it is not the sign that picks it', () => {
  const fmt = (code: string, v: string): string =>
    applyNumberFormat(v, 164, new Map([[164, code]]));

  it('takes the first section whose test the value passes', () => {
    // FormatKM.xlsx carries its own expected values beside every case.
    const km = '[>999999]#,,"M";[>999]#,"K";#';
    expect(fmt(km, '1.02')).toBe('1');
    expect(fmt(km, '102')).toBe('102');
    expect(fmt(km, '1021.02')).toBe('1K');
    expect(fmt(km, '102102.102')).toBe('102K');
    expect(fmt(km, '1021021.02')).toBe('1M');
    expect(fmt(km, '1021021021.02')).toBe('1021M');
  });

  it('knows every comparison, spaces and all', () => {
    expect(fmt('[<10]#" Wow"', '1.5')).toBe('2 Wow');
    expect(fmt('[>10]#" Big"', '11')).toBe('11 Big');
    expect(fmt('[<=10]#" Wow"', '10')).toBe('10 Wow');
    expect(fmt('[>=10]#" Big"', '10')).toBe('10 Big');
    expect(fmt('[=10]#" Wow"', '10')).toBe('10 Wow');
    expect(fmt('[<>10]#" Wow"', '11')).toBe('11 Wow');
    expect(fmt('[<   10]#" Wow"', '1')).toBe('1 Wow');
  });

  it('prints the number plainly when it satisfies no condition', () => {
    expect(fmt('[<10]#" Wow"', '11')).toBe('11');
    expect(fmt('[>10]#" Big"', '10')).toBe('10');
  });

  it('leaves an unconditional format on the sign', () => {
    expect(fmt('#,##0;[Red]-#,##0', '-1234')).toBe('-1,234');
    expect(fmt('0.0;(0.0);"zero"', '0')).toBe('zero');
  });
});

describe('applyNumberFormat — a format code built to stall it', () => {
  // The code is the workbook's own, read again in every cell that names it.
  // Each of these took the expressions that read it time quadratic in the
  // code's length: seconds at a hundred thousand characters, in each cell.
  function timed(code: string, value: string): { out: string; ms: number } {
    const fmt = new Map([[164, code]]);
    const start = performance.now();
    const out = applyNumberFormat(value, 164, fmt);
    numberFormatColorHex(value, 164, fmt);
    return { out, ms: performance.now() - start };
  }
  /** Digits with a comma before every third from the right. */
  const grouped = (digits: string): string => {
    const parts: Array<string> = [];
    for (let end = digits.length; end > 0; end -= 3) {
      parts.unshift(digits.slice(Math.max(0, end - 3), end));
    }
    return parts.join(',');
  };

  it('groups the zeros a format pads with', () => {
    expect(applyNumberFormat('5', 164, new Map([[164, '000,000']]))).toBe('000,005');
    expect(applyNumberFormat('1234567', 164, new Map([[164, '#,##0']]))).toBe('1,234,567');
  });

  it('groups a hundred thousand of them in time linear in the code', () => {
    const { out, ms } = timed(`#,${'0'.repeat(99_999)}`, '1234.5');
    expect(out).toBe(grouped(`${'0'.repeat(99_995)}1235`));
    expect(ms).toBeLessThan(1000);
  });

  it('finds a fraction after a long head in time linear in it', () => {
    const { out, ms } = timed(`${'0'.repeat(99_998)}x0 ?/?`, '1.5');
    expect(out).toBe(`${'0'.repeat(99_998)}x1 1/2`);
    expect(ms).toBeLessThan(1000);
  });

  it('gives up on a scientific code with a line end in time linear in it', () => {
    const code = `0.0${'E+'.repeat(50_000)}\n`;
    const { out, ms } = timed(code, '1234.5');
    expect(out).toBe(code);
    expect(ms).toBeLessThan(1000);
  });

  it('reads past brackets that never close in time linear in them', () => {
    const { out, ms } = timed(`${'['.repeat(100_000)}0`, '1234.5');
    expect(out).toBe(`${'['.repeat(100_000)}1235`);
    expect(ms).toBeLessThan(1000);
  });
});

describe('applyNumberFormat — the digits Excel shows', () => {
  // Each expectation is what Excel itself printed for the value and the code,
  // in a PDF of a probe workbook (2026-10-01): a double shown as its first 15
  // significant digits, rounded half away from zero, and zeros after them.
  const show = (value: string, code: string): string =>
    applyNumberFormat(value, 164, new Map([[164, code]]));
  const places = (n: number): string => `0.${'0'.repeat(n)}`;

  it('writes zeros past fifteen significant digits, not the binary expansion', () => {
    expect(show('0.1', places(20))).toBe('0.10000000000000000000');
    expect(show(String(1 / 3), places(20))).toBe('0.33333333333333300000');
    expect(show(String(2 / 3), places(17))).toBe('0.66666666666666700');
    expect(show('123456.789012345678', places(15))).toBe('123456.789012346000000');
    expect(show('9876543210.12345', places(15))).toBe('9876543210.123450000000000');
  });

  it('writes a number of 1E+21 and more out in full', () => {
    expect(show('1E+21', '0.00')).toBe('1000000000000000000000.00');
    expect(show('-1E+21', '0.00')).toBe('-1000000000000000000000.00');
    expect(show('1.23456789012345E+22', '0')).toBe('12345678901234500000000');
    expect(show('1E+21', '#,##0')).toBe('1,000,000,000,000,000,000,000');
  });

  it('rounds half away from zero on the decimal, however small the number', () => {
    expect(show('9.995', '0.00')).toBe('10.00');
    expect(show('-0.125', '0.00')).toBe('-0.13');
    expect(show('8.0945E-12', places(15))).toBe('0.000000000008095');
    expect(show('5E-16', places(15))).toBe('0.000000000000001');
    expect(show('2.5E-08', places(8))).toBe('0.00000003');
    expect(show('-2.5E-08', places(8))).toBe('-0.00000003');
  });

  it('rounds a scientific mantissa on the decimal too', () => {
    expect(show('1234.5', '0.000E+00')).toBe('1.235E+03');
    expect(show('0.00012345', '0.000E+00')).toBe('1.235E-04');
    expect(show('9.9995', '0.000E+00')).toBe('1.000E+01');
    expect(show('1.45', '0.0E+00')).toBe('1.5E+00');
    expect(show('0.35', '0E+00')).toBe('4E-01');
    expect(show('999.95', '##0.0E+0')).toBe('1.0E+3');
    expect(show('0.00012345', '##0.0E+0')).toBe('123.5E-6');
    expect(show(String(1 / 3), `${places(18)}E+00`)).toBe('3.333333333333330000E-01');
  });

  it('shows as many places as a code asks for, past a hundred', () => {
    // toFixed stops at a hundred places, and one cell whose code asked for more
    // threw out of the whole conversion.
    expect(show('0.1', places(101))).toBe(`0.1${'0'.repeat(100)}`);
    expect(show('1234.5', `${places(101)}E+00`)).toBe(`1.2345${'0'.repeat(97)}E+03`);
    expect(show('0.5000123456', `hh:mm:ss.${'0'.repeat(101)}`)).toMatch(/^12:00:01\.0666\d{97}$/);
  });

  it('converts a workbook whose cell asks for more than a hundred places', async () => {
    const styles =
      `<numFmts count="1"><numFmt numFmtId="164" formatCode="${places(101)}"/></numFmts>` +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border/></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>';
    const xlsx = buildXlsx({ rows: [[{ value: 1.5, styleIndex: 1 }]], stylesXml: styles });
    const html = new TextDecoder().decode(await Ream.parse(xlsx).convert('html'));
    expect(html).toContain(`1.5${'0'.repeat(100)}`);
  });
});
