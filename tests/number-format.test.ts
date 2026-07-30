import { describe, expect, it } from 'vitest';

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
    // Room for everything ⇒ everything.
    expect(generalToWidth('4.3900875881221957', upTo(30))).toBe('4.390087588122196');
    // An integer that will not fit goes scientific rather than reading as a
    // number a thousand times smaller (escape-unicode.xlsx printed "1161014"
    // for 1161014163).
    expect(generalToWidth('1161014163', upTo(9))).toBe('1.161E+09');
    expect(generalToWidth('1161014163', upTo(30))).toBe('1161014163');
    // Too narrow even for that: the cell says so its own way, with hashes.
    expect(generalToWidth('1161014163', upTo(3))).toBe('1161014163');
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
