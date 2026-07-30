// ECMA-376 Part 1 §18.8.30 — Built-in number formats and §18.8.31 — custom
// format codes. Implements the subset typical business spreadsheets use:
//   - Built-in: 0..4, 9..10, 37..40, 49 (General / integer / decimal /
//     thousands / percent / accounting basics / text)
//   - Custom format strings with `0`, `#`, `,`, `.`, `%`, quoted literals,
//     and `[colour]` / `[locale]` codes (the latter two stripped silently)
// Dates (m/d/yyyy etc.) are deferred — the cell value is shown verbatim.

import { INDEXED_COLORS } from '@/core/indexed-colors';

const BUILTIN: ReadonlyMap<number, string> = new Map([
  [0, 'General'],
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [37, '#,##0_);(#,##0)'],
  [38, '#,##0_);[Red](#,##0)'],
  [39, '#,##0.00_);(#,##0.00)'],
  [40, '#,##0.00_);[Red](#,##0.00)'],
  [49, '@'],
]);

const BUILTIN_DATE_FORMATS: ReadonlyMap<number, string> = new Map([
  [14, 'm/d/yyyy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yyyy h:mm'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mm:ss.0'],
]);

/**
 * Render a cell's raw stored value through its number format (§18.8.30 built-ins
 * and §18.8.31 custom codes). Handles `General`, text (`@`), the built-in date /
 * time formats, custom date codes, and the numeric placeholder grammar (`0` `#`
 * `,` `.` `%`, scientific notation, quoted literals, `;`-separated sections). An
 * unknown id, or a value that does not parse as a number, falls back to a plain
 * render of `rawValue`.
 *
 * @param rawValue      The cell's raw stored value (a serial for dates).
 * @param numFmtId      The number-format id from the cell's `cellXf`.
 * @param customFormats Custom format codes by id (from `<numFmts>`).
 * @param date1904      The workbook date epoch (1904 vs the 1900 default).
 * @returns The formatted display string.
 */
export function applyNumberFormat(
  rawValue: string,
  numFmtId: number,
  customFormats: ReadonlyMap<number, string>,
  date1904: boolean = false,
): string {
  if (rawValue.length === 0) return '';
  if (numFmtId === 0) return defaultNumberRender(rawValue);
  if (numFmtId === 49) return rawValue;

  const builtinDate = BUILTIN_DATE_FORMATS.get(numFmtId);
  if (builtinDate !== undefined) return formatExcelDate(rawValue, builtinDate, date1904);

  const custom = customFormats.get(numFmtId);
  if (custom !== undefined && (isDateFormat(custom) || hasElapsedToken(custom)))
    return formatExcelDate(rawValue, custom, date1904);

  const format = custom ?? BUILTIN.get(numFmtId);
  if (!format) return defaultNumberRender(rawValue);
  // Excel accepts the name in any case, and writers use every spelling of it:
  // bug-fixes.xlsx declares `<numFmt formatCode="GENERAL"/>`. Matched exactly,
  // the code falls through to the placeholder grammar, which finds no digit
  // placeholder in it and renders the whole word as a literal prefix — "GENERAL1".
  if (format.trim().toLowerCase() === 'general') return defaultNumberRender(rawValue);

  return applyFormatString(rawValue, format);
}

/**
 * The colour a number format paints its cell in (§18.8.31), as a 6-digit RRGGBB,
 * or undefined when the section that applies declares none.
 *
 * A `[Red]` at the head of a section is part of the format, not decoration: the
 * accounting codes write the negative section as `[Red]-#,##0.00`, and rendering
 * it in the cell's own font colour loses the one signal the format exists for.
 * `[Color 12]` indexes the same legacy palette a `<color indexed>` does.
 *
 * @param rawValue      The cell's raw stored value — it selects the section.
 * @param numFmtId      The number-format id from the cell's `cellXf`.
 * @param customFormats Custom format codes by id (from `<numFmts>`).
 */
export function numberFormatColorHex(
  rawValue: string,
  numFmtId: number,
  customFormats: ReadonlyMap<number, string>,
): string | undefined {
  const format = customFormats.get(numFmtId) ?? BUILTIN.get(numFmtId);
  if (format === undefined || !format.includes('[')) return undefined;
  const sections = splitSections(format);
  const n = Number(rawValue);
  const section = Number.isFinite(n)
    ? (sections[sectionIndexFor(n, sections)] ?? sections[0]!)
    : (sections[3] ?? sections[0]!);
  return colorOfSection(section);
}

const FORMAT_COLORS: ReadonlyMap<string, string> = new Map([
  ['black', '000000'],
  ['blue', '0000FF'],
  ['cyan', '00FFFF'],
  ['green', '00FF00'],
  ['magenta', 'FF00FF'],
  ['red', 'FF0000'],
  ['white', 'FFFFFF'],
  ['yellow', 'FFFF00'],
]);

function colorOfSection(section: string): string | undefined {
  for (const m of section.matchAll(/\[([^\]]*)\]/g)) {
    const body = m[1]!.trim().toLowerCase();
    const named = FORMAT_COLORS.get(body);
    if (named !== undefined) return named;
    const indexed = /^color\s*(\d+)$/.exec(body);
    if (indexed) {
      // §18.8.31 numbers the palette from 1; §18.8.27 indexes it from 0.
      const i = Number(indexed[1]) - 1;
      if (i >= 0 && i < INDEXED_COLORS.length) return INDEXED_COLORS[i];
    }
  }
  return undefined;
}

function sectionIndexFor(n: number, sections: ReadonlyArray<string>): number {
  if (n > 0) return 0;
  if (n < 0) return sections.length > 1 ? 1 : 0;
  return sections.length > 2 ? 2 : 0;
}

/**
 * Excel serial date → JS `Date`. The default 1900 epoch uses 1899-12-30 as day 0
 * (so serial 1 is 1900-01-01) and inherits the Lotus 1-2-3 leap-year bug: serial 60
 * is considered "1900-02-29" which never existed. For serial ≥ 61 the simple
 * formula is exact; values < 60 are vanishingly rare in business sheets (and our
 * render is approximate anyway).
 *
 * The 1904 epoch (legacy Mac Excel) uses 1904-01-01 as day 0 and has no leap bug.
 * Files saved with `<workbookPr date1904="1"/>` store dates offset by exactly 1462
 * days from the 1900-epoch interpretation.
 */
export function excelSerialToDate(serial: number, date1904: boolean): Date {
  const ms = serial * 86400 * 1000;
  if (date1904) return new Date(Date.UTC(1904, 0, 1) + ms);
  return new Date(Date.UTC(1899, 11, 30) + ms);
}

/**
 * The inverse over a UTC calendar date: `(year, month0, day)` → the integer Excel
 * serial day, using the same epoch. Round-trips {@link excelSerialToDate} exactly
 * for an integer serial (the time-of-day is zero), so a serial → parts → serial
 * loop is stable. Used by the formula engine's date functions and the `timePeriod`
 * windows (E-SHEET W9) to map an injected reference date into serial space.
 *
 * @param month0 The 0-indexed month (0 = January).
 */
export function excelSerialFromUtcParts(
  year: number,
  month0: number,
  day: number,
  date1904: boolean,
): number {
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return Math.round((Date.UTC(year, month0, day) - epoch) / 86400000);
}

const MONTH_ABBR = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];
const MONTH_FULL = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];
const DAY_ABBR = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const DAY_FULL = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

/** An elapsed h/m/s count, truncated and zero-padded to its placeholder width. */
function elapsed(value: number, width: number): string {
  const whole = Math.floor(value);
  const digits = String(Math.abs(whole)).padStart(width, '0');
  return whole < 0 ? `-${digits}` : digits;
}

/**
 * `[h]`, `[mm]`, `[ss]` — the elapsed-time brackets (§18.8.31), which suppress
 * the wrap at 24 h / 60 min / 60 s and count from serial 0 instead.
 *
 * They carry no date letter outside the bracket, so a code like `[ss].00` reads
 * as a plain numeric format to {@link isDateFormat} and lands in the placeholder
 * grammar, where `[ss]` is stripped as if it were a colour and `.00` renders the
 * serial as ".3".
 */
function hasElapsedToken(code: string): boolean {
  return /\[(?:h+|m+|s+)\]/i.test(code.replace(/"[^"]*"/g, ''));
}

function isDateFormat(code: string): boolean {
  // Strip quoted literals and [] codes, then look for any date token. The
  // "m" letter alone is ambiguous (month vs minute) so it can't be a sole
  // signal, but its presence alongside d/y/h/s already implies dates.
  const cleaned = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  return /[dyhs]|m+/i.test(cleaned) && /[dyhs]/i.test(cleaned);
}

interface DateToken {
  readonly kind:
    | 'lit'
    | 'y'
    | 'M'
    | 'd'
    | 'h'
    | 'm'
    | 's'
    | 'ampm'
    | 'elapsed-h'
    | 'elapsed-m'
    | 'elapsed-s'
    | 'subsec';
  readonly text: string;
}

function tokenizeDateFormat(format: string): Array<DateToken> {
  const tokens: Array<DateToken> = [];
  let i = 0;
  while (i < format.length) {
    const ch = format[i]!;
    if (ch === '"') {
      let lit = '';
      i++;
      while (i < format.length && format[i] !== '"') {
        lit += format[i]!;
        i++;
      }
      i++;
      tokens.push({ kind: 'lit', text: lit });
      continue;
    }
    if (ch === '\\' && i + 1 < format.length) {
      tokens.push({ kind: 'lit', text: format[i + 1]! });
      i += 2;
      continue;
    }
    if (ch === '[') {
      let body = '';
      i++;
      while (i < format.length && format[i] !== ']') {
        body += format[i]!;
        i++;
      }
      i++;
      if (/^h+$/i.test(body)) tokens.push({ kind: 'elapsed-h', text: body });
      else if (/^m+$/i.test(body)) tokens.push({ kind: 'elapsed-m', text: body });
      else if (/^s+$/i.test(body)) tokens.push({ kind: 'elapsed-s', text: body });
      // Anything else in brackets (colors, locales) is ignored.
      continue;
    }
    if (ch === '.') {
      // `ss.00` / `[ss].0` — the decimals of a seconds token, not a literal dot
      // followed by literal zeros (which is how built-in 47, `mm:ss.0`, read).
      const last = tokens[tokens.length - 1];
      const run = /^\.(0+)/.exec(format.substring(i));
      if (run && (last?.kind === 's' || last?.kind === 'elapsed-s')) {
        tokens.push({ kind: 'subsec', text: run[1]! });
        i += run[0].length;
        continue;
      }
    }
    const lower = ch.toLowerCase();
    if (lower === 'y' || lower === 'm' || lower === 'd' || lower === 'h' || lower === 's') {
      let run = ch;
      i++;
      while (i < format.length && format[i]!.toLowerCase() === lower) {
        run += format[i]!;
        i++;
      }
      if (lower === 'y') tokens.push({ kind: 'y', text: run });
      else if (lower === 'd') tokens.push({ kind: 'd', text: run });
      else if (lower === 'h') tokens.push({ kind: 'h', text: run });
      else if (lower === 's') tokens.push({ kind: 's', text: run });
      else tokens.push({ kind: 'm', text: run });
      continue;
    }
    if (format.startsWith('AM/PM', i) || format.startsWith('A/P', i)) {
      const len = format.startsWith('AM/PM', i) ? 5 : 3;
      tokens.push({ kind: 'ampm', text: format.substring(i, i + len) });
      i += len;
      continue;
    }
    tokens.push({ kind: 'lit', text: ch });
    i++;
  }
  return tokens;
}

function resolveMonthVsMinute(tokens: Array<DateToken>): Array<DateToken> {
  // §18.8.31: "m" or "mm" immediately after h/hh or immediately before s/ss
  // is minutes; otherwise month.
  const resolved: Array<DateToken> = tokens.map((t) => ({ ...t }));
  for (let i = 0; i < resolved.length; i++) {
    const t = resolved[i]!;
    if (t.kind !== 'm') continue;
    let prev = i - 1;
    while (prev >= 0 && resolved[prev]!.kind === 'lit') prev--;
    let next = i + 1;
    while (next < resolved.length && resolved[next]!.kind === 'lit') next++;
    const prevKind = prev >= 0 ? resolved[prev]!.kind : undefined;
    const prevIsHour = prevKind === 'h' || prevKind === 'elapsed-h';
    const nextIsSec = next < resolved.length && resolved[next]!.kind === 's';
    if (prevIsHour || nextIsSec) {
      // Keep kind 'm' but tag as minutes via length-only check at render time.
      continue;
    }
    // Otherwise it's a month token — relabel as 'M'.
    resolved[i] = { kind: 'M', text: t.text };
  }
  return resolved;
}

function formatExcelDate(rawValue: string, format: string, date1904: boolean): string {
  const serial = Number(rawValue);
  if (!Number.isFinite(serial)) return rawValue;

  // A format may carry up to four `;`-separated sections
  // (positive;negative;zero;text). For a date/number value only the first
  // section applies — e.g. `mmm-yy;@` or `m/d/yyyy;@` (the `@` text-section is
  // for string cells). Render solely the first section so the `;@` tail and any
  // negative/zero/text sub-formats don't leak verbatim into the output.
  format = splitSections(format)[0] ?? format;

  const tokens = resolveMonthVsMinute(tokenizeDateFormat(format));
  // Elapsed h/m/s count from serial 0 rather than wrapping at the next unit, so
  // they are read off the serial itself. Fractional seconds are rounded before
  // the split so the two halves agree — 59.996 s under `.00` is 60.00, not 59
  // with a ".00" tail. Only round when decimals are actually asked for; whole
  // seconds keep truncating, as every other token here does.
  const subsecDigits = tokens.find((t) => t.kind === 'subsec')?.text.length ?? 0;
  const scale = 10 ** subsecDigits;
  const totalSeconds =
    subsecDigits > 0 ? Math.round(serial * 86400 * scale) / scale : serial * 86400;

  const date = excelSerialToDate(subsecDigits > 0 ? totalSeconds / 86400 : serial, date1904);
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1; // 1-12
  const day = date.getUTCDate();
  const weekday = date.getUTCDay();
  const hour24 = date.getUTCHours();
  const minute = date.getUTCMinutes();
  const second = date.getUTCSeconds();

  const has12hr = tokens.some((t) => t.kind === 'ampm');
  const hourValue = has12hr ? (hour24 % 12 === 0 ? 12 : hour24 % 12) : hour24;
  const ampmLabel = (template: string): string => {
    const isUpper = template[0] === 'A';
    const isShort = template.length <= 3;
    if (hour24 < 12) return isShort ? (isUpper ? 'A' : 'a') : isUpper ? 'AM' : 'am';
    return isShort ? (isUpper ? 'P' : 'p') : isUpper ? 'PM' : 'pm';
  };

  let out = '';
  for (const t of tokens) {
    switch (t.kind) {
      case 'lit':
        out += t.text;
        break;
      case 'y':
        out += t.text.length <= 2 ? pad2(year % 100) : String(year).padStart(4, '0');
        break;
      case 'M': {
        const len = t.text.length;
        if (len === 5) out += MONTH_FULL[month - 1]![0]!;
        else if (len === 4) out += MONTH_FULL[month - 1]!;
        else if (len === 3) out += MONTH_ABBR[month - 1]!;
        else if (len === 2) out += pad2(month);
        else out += String(month);
        break;
      }
      case 'd': {
        const len = t.text.length;
        if (len >= 4) out += DAY_FULL[weekday]!;
        else if (len === 3) out += DAY_ABBR[weekday]!;
        else if (len === 2) out += pad2(day);
        else out += String(day);
        break;
      }
      case 'h':
        out += t.text.length >= 2 ? pad2(hourValue) : String(hourValue);
        break;
      case 'elapsed-h':
        out += elapsed(totalSeconds / 3600, t.text.length);
        break;
      case 'elapsed-m':
        out += elapsed(totalSeconds / 60, t.text.length);
        break;
      case 'elapsed-s':
        out += elapsed(totalSeconds, t.text.length);
        break;
      case 'subsec':
        out += (totalSeconds - Math.floor(totalSeconds)).toFixed(subsecDigits).substring(1);
        break;
      case 'm':
        out += t.text.length >= 2 ? pad2(minute) : String(minute);
        break;
      case 's':
        out += t.text.length >= 2 ? pad2(second) : String(second);
        break;
      case 'ampm':
        out += ampmLabel(t.text);
        break;
    }
  }
  return out;
}

function defaultNumberRender(rawValue: string): string {
  const n = Number(rawValue);
  if (!Number.isFinite(n)) return rawValue;
  // Don't reformat — just trim trailing zeros after a stored decimal.
  if (Number.isInteger(n)) return String(n);
  return String(n);
}

/**
 * A General number rounded to the decimals that fit `maxChars` characters.
 *
 * General is not a fixed format: a spreadsheet shows as many decimal places as
 * the column has room for and ROUNDS to that, so 4.3900875881221957 in a
 * column eight characters wide reads 4.390088. We rendered every stored digit
 * and let the cell clip it, which turns the same value into 4.390087 — off by
 * one in the last place shown, with nothing to say a digit was cut.
 *
 * The integer part is never dropped: a number too wide even without decimals
 * keeps them all and the cell says so its own way (see hashOnOverflow).
 *
 * @param rawValue The cell's stored value.
 * @param maxChars How many characters the column has room for.
 */
export function generalToWidth(rawValue: string, maxChars: number): string {
  const n = Number(rawValue);
  if (!Number.isFinite(n) || Number.isInteger(n)) return defaultNumberRender(rawValue);
  const full = defaultNumberRender(rawValue);
  if (full.length <= maxChars) return full;
  const dot = full.indexOf('.');
  if (dot < 0) return full;
  // Room left for decimals once the sign, the integer part and the point are in.
  const room = Math.floor(maxChars) - dot - 1;
  if (room < 1) return full;
  const decimals = Math.min(room, full.length - dot - 1);
  const shifted = Number(`${Number(n.toPrecision(15))}e${decimals}`);
  const rounded = shifted < 0 ? -Math.round(-shifted) : Math.round(shifted);
  const back = Number(`${rounded}e${-decimals}`);
  return Number.isFinite(back) ? defaultNumberRender(String(back)) : full;
}

function applyFormatString(rawValue: string, format: string): string {
  const n = Number(rawValue);
  const sections = splitSections(format);

  if (!Number.isFinite(n)) {
    const textSection = sections[3];
    if (textSection !== undefined) {
      return textSection.replace(/@/g, rawValue);
    }
    return rawValue;
  }

  const sectionIdx = sectionIndexFor(n, sections);
  const section = sections[sectionIdx] ?? sections[0]!;
  return applyNumericSection(n, section, sectionIdx === 1);
}

// Split on top-level ';' — ignore ';' inside quoted strings and brackets.
function splitSections(format: string): Array<string> {
  const out: Array<string> = [];
  let current = '';
  let inQuotes = false;
  let bracketDepth = 0;
  for (let i = 0; i < format.length; i++) {
    const ch = format[i]!;
    if (ch === '"') {
      inQuotes = !inQuotes;
      current += ch;
      continue;
    }
    if (!inQuotes) {
      if (ch === '[') bracketDepth++;
      else if (ch === ']') bracketDepth = Math.max(0, bracketDepth - 1);
      else if (ch === '\\' && i + 1 < format.length) {
        current += ch + format[i + 1]!;
        i++;
        continue;
      } else if (ch === ';' && bracketDepth === 0) {
        out.push(current);
        current = '';
        continue;
      }
    }
    current += ch;
  }
  out.push(current);
  return out;
}

function countDigitPlaceholders(s: string): number {
  let n = 0;
  for (const c of s) if (c === '0' || c === '#' || c === '?') n++;
  return n;
}

// §18.8.31 scientific notation: `0.00E+00`, `##0.0E+0` (engineering), `0.0e-0`.
// The mantissa is normalised so its integer part holds `intDigits` significant
// figures (1 for `0.00E+00`, 3 for `##0.0E+0` → exponent snaps to a multiple of
// 3); the exponent is zero-padded to the placeholder count and carries a sign
// (always for `E+`, only when negative for `E-`). The `E`/`e` case is preserved.
/**
 * `value` rounded to `decimals` places the way a spreadsheet rounds: on the
 * DECIMAL number, half away from zero.
 *
 * `toFixed` rounds the binary double, which is not the number the file means.
 * A rate stored as 0.0095 is 0.009499999999999999… in binary; times 100 that is
 * 0.9499999999999998, and `toFixed(1)` dutifully answers "0.9" where Excel and
 * every other reader show 1.0%. AverageTaxRates.xlsx had eleven of its
 * percentages a tenth low for exactly this reason.
 *
 * Rounding to 15 significant digits first collapses the binary noise back to
 * the decimal the author typed; shifting through a string exponent then keeps
 * the scaling exact, so the final round sees 9.5 and not 9.499999999999998.
 */
function toFixedDecimal(value: number, decimals: number): string {
  if (!Number.isFinite(value)) return value.toFixed(decimals);
  const decimal = Number(value.toPrecision(15));
  const shifted = Number(`${decimal}e${decimals}`);
  if (!Number.isFinite(shifted)) return value.toFixed(decimals);
  // Excel rounds a half away from zero; Math.round takes it towards +∞.
  const rounded = shifted < 0 ? -Math.round(-shifted) : Math.round(shifted);
  const back = Number(`${rounded}e${-decimals}`);
  return Number.isFinite(back) ? back.toFixed(decimals) : value.toFixed(decimals);
}

function formatScientific(value: number, cleaned: string, negativeSection: boolean): string {
  const m = /^(.*?)([eE])([+-])(.*)$/.exec(cleaned);
  if (!m) return cleaned;
  const mantissaFmt = m[1]!;
  const eChar = m[2]!;
  const expSignFmt = m[3]!;
  const expFmt = m[4]!;

  const dot = mantissaFmt.indexOf('.');
  const intDigits = Math.max(
    1,
    countDigitPlaceholders(dot >= 0 ? mantissaFmt.slice(0, dot) : mantissaFmt),
  );
  const decimals = dot >= 0 ? countDigitPlaceholders(mantissaFmt.slice(dot + 1)) : 0;

  let exp = 0;
  let mant = Math.abs(value);
  if (mant !== 0) {
    exp = Math.floor(Math.log10(mant));
    exp -= ((exp % intDigits) + intDigits) % intDigits; // engineering grouping
    mant = mant / Math.pow(10, exp);
    // Rounding the mantissa can carry it up to 10^intDigits — renormalise.
    if (Number(mant.toFixed(decimals)) >= Math.pow(10, intDigits)) {
      exp += intDigits;
      mant = mant / Math.pow(10, intDigits);
    }
  }

  const mantStr = mant.toFixed(decimals);
  const expDigits = countDigitPlaceholders(expFmt) || 2;
  const expStr = String(Math.abs(exp)).padStart(expDigits, '0');
  const expSign = exp < 0 ? '-' : expSignFmt === '+' ? '+' : '';
  const signPrefix = value < 0 && !negativeSection ? '-' : '';
  return `${signPrefix}${mantStr}${eChar}${expSign}${expStr}`;
}

/**
 * Resolve the `[...]` codes a numeric section may carry (§18.8.31).
 *
 * `[$SYMBOL-LOCALE]` is a currency tag and its SYMBOL is *rendered* — dropping
 * the whole bracket, as a colour or a locale is dropped, silently loses the `$`
 * from `[$$-409]#,##0`. The symbol becomes a quoted literal so the placeholder
 * grammar treats it as text. Everything else in brackets (colours, locales,
 * conditions) is decoration here and goes.
 */
function resolveBracketCodes(format: string): string {
  return format.replace(/\[([^\]]*)\]/g, (_, body: string) => {
    const currency = /^\$([^-]*)(?:-.*)?$/.exec(body);
    const symbol = currency?.[1]?.replace(/"/g, '') ?? '';
    return symbol.length > 0 ? `"${symbol}"` : '';
  });
}

/**
 * §18.8.31 fraction formats — `# ?/?`, `# ??/??`, `?/?`, `# ?/16`.
 *
 * `?` is a digit placeholder like `#` but padded to width with spaces, and the
 * count of them on the right of the slash bounds the denominator: `??` admits
 * any denominator up to 99, so 25.378 renders as "25 31/82". A literal number
 * there instead (`?/16`) fixes the denominator. Returns undefined when the
 * section is not a fraction, leaving the ordinary grammar to run.
 */
function formatFraction(
  value: number,
  cleaned: string,
  negativeSection: boolean,
): string | undefined {
  const m = /^(.*?)([0#?]+)\s*\/\s*([0#?]+|\d+)(.*)$/.exec(cleaned);
  if (!m) return undefined;
  const head = m[1]!;
  const denFmt = m[3]!;
  const tail = m[4]!;

  // The integer part, if any, is the placeholder run at the end of the head.
  const intMatch = /[0#?]+\s*$/.exec(head);
  const literalPrefix = unquoteLiteral(head.substring(0, intMatch ? intMatch.index : head.length));
  const hasInteger = intMatch !== null;

  const magnitude = Math.abs(value);
  const whole = hasInteger ? Math.floor(magnitude) : 0;
  const rest = magnitude - whole;

  const fixedDen = /^\d+$/.test(denFmt) ? Number(denFmt) : undefined;
  const maxDen = fixedDen ?? Math.pow(10, denFmt.length) - 1;
  if (!Number.isFinite(maxDen) || maxDen < 1) return undefined;

  let num: number;
  let den: number;
  if (fixedDen !== undefined) {
    den = fixedDen;
    num = Math.round(rest * fixedDen);
  } else {
    [num, den] = bestRational(rest, maxDen);
  }
  const pieces: Array<string> = [];
  if (hasInteger) {
    // Rounding can carry the fraction to a whole unit; without an integer field
    // to carry into (`?/?` is an improper fraction) it stays in the numerator.
    let integer = whole;
    if (den > 0 && num >= den) {
      integer += Math.floor(num / den);
      num %= den;
    }
    if (integer !== 0 || num === 0) pieces.push(String(integer));
  }
  if (num !== 0 || !hasInteger) pieces.push(`${num}/${den}`);

  const sign = value < 0 && !negativeSection ? '-' : '';
  return `${literalPrefix}${sign}${pieces.join(' ')}${unquoteLiteral(tail)}`;
}

/**
 * The closest `n/d` to `value` with `d <= maxDen`, by continued fractions —
 * the same approximation Excel and LibreOffice show for a `?/?` format.
 */
function bestRational(value: number, maxDen: number): [number, number] {
  if (value === 0) return [0, 1];
  let [prevN, prevD] = [0, 1];
  let [n, d] = [1, 0];
  let x = value;
  for (let i = 0; i < 64; i++) {
    const a = Math.floor(x);
    const nextN = a * n + prevN;
    const nextD = a * d + prevD;
    if (nextD > maxDen) break;
    [prevN, prevD] = [n, d];
    [n, d] = [nextN, nextD];
    const frac = x - a;
    if (frac < 1e-12) break;
    x = 1 / frac;
  }
  if (d === 0) return [Math.round(value * maxDen), maxDen];
  return [n, d];
}

function applyNumericSection(value: number, format: string, negativeSection: boolean): string {
  const cleaned = resolveBracketCodes(format);

  const fraction = formatFraction(value, cleaned, negativeSection);
  if (fraction !== undefined) return fraction;

  if (/[eE][+-]/.test(cleaned)) return formatScientific(value, cleaned, negativeSection);

  const isPercent = cleaned.includes('%');
  let magnitude = Math.abs(value);
  if (isPercent) magnitude *= 100;

  const { intFormat, decFormat, literalPrefix, literalSuffix } = splitNumberFormat(cleaned);
  let decimals = 0;
  for (const c of decFormat) if (c === '0' || c === '#') decimals++;

  const fixed = toFixedDecimal(magnitude, decimals);
  const [intRaw, decRaw] = fixed.split('.');
  const useThousands = /,(?=\d)/.test(intFormat) || /[0#],[0#]/.test(intFormat);
  const intStr = useThousands ? intRaw!.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : intRaw!;

  // §18.8.31: `#` is an *optional* digit, so an integer part of zero written
  // only with `#` renders as nothing at all — format `#` blanks a zero cell
  // (tdf171828 uses that to hide a whole column of them) and `#.##` shows ".5"
  // rather than "0.5". A `0` or `?` anywhere in the integer part forces it.
  let numberPart = intRaw === '0' && !/[0?]/.test(intFormat) ? '' : intStr;
  if (decimals > 0) {
    let dec = decRaw ?? ''.padEnd(decimals, '0');
    if (dec.length < decimals) dec = dec.padEnd(decimals, '0');
    // '0' forces its digit, '#' drops it when the whole tail from there on is
    // zeros — so `#.##` renders 0.5 as ".5", 0.55 as ".55" and 0.05 as ".05".
    // The decision is per position and reads rightwards, which is why the tail
    // is trimmed after the fact rather than skipped as it is built.
    let kept = '';
    for (let i = 0; i < decFormat.length; i++) {
      const placeholder = decFormat[i]!;
      if (placeholder === '0' || placeholder === '#') kept += dec[i] ?? '0';
    }
    for (let i = kept.length - 1; i >= 0 && kept[i] === '0' && decFormat[i] === '#'; i--) {
      kept = kept.substring(0, i);
    }
    // Strip trailing # of empty content.
    if (kept.length > 0) numberPart += '.' + kept;
    else if (decFormat.includes('0')) numberPart += '.' + dec;
  }

  let signPrefix = '';
  if (value < 0 && !negativeSection) signPrefix = '-';

  return `${literalPrefix}${signPrefix}${numberPart}${isPercent ? '%' : ''}${literalSuffix}`;
}

interface SplitNumberFormat {
  literalPrefix: string;
  intFormat: string;
  decFormat: string;
  literalSuffix: string;
}

function splitNumberFormat(cleaned: string): SplitNumberFormat {
  // Find the leftmost and rightmost stretches of digit placeholders.
  // Anything outside that range is literal text.
  let firstDigit = -1;
  let lastDigit = -1;
  for (let i = 0; i < cleaned.length; i++) {
    const ch = cleaned[i]!;
    if (ch === '0' || ch === '#') {
      if (firstDigit < 0) firstDigit = i;
      lastDigit = i;
    }
  }
  if (firstDigit < 0) {
    // A section with no digit placeholder is ALL literal — and a literal still
    // has to be decoded. Returning it raw printed the format code itself:
    // Excel's Accounting format writes its zero as `_(* "-"_)`, and a balance
    // row that should read "-" read `_(* "-"_)` instead (49156.xlsx).
    return {
      literalPrefix: unquoteLiteral(cleaned),
      intFormat: '',
      decFormat: '',
      literalSuffix: '',
    };
  }
  const literalPrefix = unquoteLiteral(cleaned.substring(0, firstDigit));
  const literalSuffix = unquoteLiteral(cleaned.substring(lastDigit + 1));
  const digitRange = cleaned.substring(firstDigit, lastDigit + 1);
  // The '%' inside digit range is handled by isPercent at the caller; strip
  // it from the digit range so the dot parser doesn't get confused.
  const dotIdx = digitRange.indexOf('.');
  const intFormat = dotIdx >= 0 ? digitRange.substring(0, dotIdx) : digitRange;
  const decFormat = dotIdx >= 0 ? digitRange.substring(dotIdx + 1) : '';
  return { literalPrefix, intFormat, decFormat, literalSuffix };
}

function unquoteLiteral(s: string): string {
  // "...": preserve content; \x: keep x; _x: width placeholder, treat as space.
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (ch === '"') {
      i++;
      while (i < s.length && s[i] !== '"') {
        out += s[i]!;
        i++;
      }
      continue;
    }
    if (ch === '\\' && i + 1 < s.length) {
      out += s[i + 1]!;
      i++;
      continue;
    }
    if (ch === '_' && i + 1 < s.length) {
      out += ' ';
      i++;
      continue;
    }
    if (ch === '*' && i + 1 < s.length) {
      // §18.8.31 `*x` repeats x until the cell is full — a width filler, not a
      // literal. The accounting formats all carry `_-* ` and printed a literal
      // asterisk in front of every value ("* 210,896"). Repeating it would need
      // the laid-out cell width, which this layer does not have; the character
      // is a space in every format Excel itself writes, so emit nothing.
      i++;
      continue;
    }
    if (ch === '%') {
      // handled by isPercent at caller
      continue;
    }
    out += ch;
  }
  return out;
}
