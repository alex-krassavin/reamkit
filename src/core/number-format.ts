// ECMA-376 Part 1 §18.8.30 — Built-in number formats and §18.8.31 — custom
// format codes. Implements the subset typical business spreadsheets use:
//   - Built-in: 0..4, 9..10, 37..40, 49 (General / integer / decimal /
//     thousands / percent / accounting basics / text)
//   - Custom format strings with `0`, `#`, `,`, `.`, `%`, quoted literals,
//     and `[colour]` / `[locale]` codes (the latter two stripped silently)
//   - Dates: the §18.8.30 built-ins and custom day/month/year codes, off the
//     workbook's 1900 or 1904 epoch.

import { INDEXED_COLORS } from '@/core/indexed-colors';

const BUILTIN: ReadonlyMap<number, string> = new Map([
  [0, 'General'],
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  // §18.8.30 — the currency and fraction builtins. Absent from this table a
  // cell falls through to `defaultNumberRender`, so 49273.xlsx's `# ?/?` cell
  // printed 0.5 where every reader shows ` 1/2`; the fraction engine below was
  // there all along, nothing ever asked it.
  [5, '$#,##0_);($#,##0)'],
  [6, '$#,##0_);[Red]($#,##0)'],
  [7, '$#,##0.00_);($#,##0.00)'],
  [8, '$#,##0.00_);[Red]($#,##0.00)'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [37, '#,##0_);(#,##0)'],
  [38, '#,##0_);[Red](#,##0)'],
  [39, '#,##0.00_);(#,##0.00)'],
  [40, '#,##0.00_);[Red](#,##0.00)'],
  [48, '##0.0E+0'],
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
 * Render a TEXT cell's string through its number format (§18.8.31). Only the
 * text section applies — the fourth, or a format's only section when it holds
 * `@` — with `@` standing for the text and the literals around it written as
 * they are: Excel shows "abc" under `0;-0;0;"["@"]"` as [abc], under `@" kg"`
 * as "abc kg" and under `_(@_)` as " abc " (2026-11-20). Any other format,
 * number or date, leaves the text as it is, as Excel does under `0.00`.
 *
 * @param text          The cell's string.
 * @param numFmtId      The number-format id from the cell's `cellXf`.
 * @param customFormats Custom format codes by id (from `<numFmts>`).
 * @returns The displayed text.
 */
export function applyTextFormat(
  text: string,
  numFmtId: number,
  customFormats: ReadonlyMap<number, string>,
): string {
  if (numFmtId === 0 || numFmtId === 49) return text;
  const format = customFormats.get(numFmtId) ?? BUILTIN.get(numFmtId);
  if (format === undefined) return text;
  return formatText(text, splitSections(format));
}

/**
 * `text` under the text section of a format, or as it is when the format has
 * none. A fourth section that is empty hides the text: `;;;` shows nothing.
 * `@` may stand more than once (`@"x"@` reads abcxabc) and `\@` is an at sign.
 */
function formatText(text: string, sections: ReadonlyArray<string>): string {
  const only = sections.length === 1 ? tokenizeSection(resolveBracketCodes(sections[0]!)) : [];
  const tokens =
    sections.length >= 4
      ? tokenizeSection(resolveBracketCodes(sections[3]!))
      : only.some((t) => t.kind === 'text')
        ? only
        : undefined;
  if (tokens === undefined) return text;
  return tokens.map((t) => (t.kind === 'text' ? text : tokenText(t))).join('');
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
  const idx = Number.isFinite(n) ? sectionFor(n, sections).index : 3;
  // A value that satisfies no condition is not formatted by any section, so it
  // takes no colour from one either.
  if (idx < 0) return undefined;
  const section = sections[idx] ?? sections[0]!;
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

// The `[…]` codes of a format, sought only as far as its last `]`. None starts
// after that, and an expression looking for a `]` that is not there runs to
// the end of the format from every `[` — quadratic in a code of brackets that
// never close.
const BRACKET_CODE = /\[([^\]]*)\]/g;

function bracketCodes(format: string): IterableIterator<RegExpMatchArray> {
  return format.slice(0, format.lastIndexOf(']') + 1).matchAll(BRACKET_CODE);
}

/** The format with each `[…]` code replaced by what `fn` makes of its inside. */
function replaceBracketCodes(format: string, fn: (body: string) => string): string {
  const end = format.lastIndexOf(']') + 1;
  return (
    format.slice(0, end).replace(BRACKET_CODE, (_, body: string) => fn(body)) + format.slice(end)
  );
}

function colorOfSection(section: string): string | undefined {
  for (const m of bracketCodes(section)) {
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

/**
 * §18.8.31 — a section may carry a CONDITION in brackets (`[>999]`), and then
 * the sections are no longer positive/negative/zero: the first whose test the
 * value passes applies, and a section without one is the else. Returns the
 * comparison, or undefined for a section that states none.
 *
 * The operator is what tells this bracket from the others a section may carry —
 * `[Red]`, `[$€-407]`, `[h]` — none of which start with one.
 */
function conditionOf(section: string): { op: string; rhs: number } | undefined {
  const m = /\[\s*(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*\]/.exec(section);
  if (!m) return undefined;
  const rhs = Number(m[2]);
  return Number.isFinite(rhs) ? { op: m[1]!, rhs } : undefined;
}

function conditionHolds(n: number, c: { op: string; rhs: number }): boolean {
  switch (c.op) {
    case '<':
      return n < c.rhs;
    case '<=':
      return n <= c.rhs;
    case '>':
      return n > c.rhs;
    case '>=':
      return n >= c.rhs;
    case '=':
      return n === c.rhs;
    default:
      return n !== c.rhs;
  }
}

/** The section a number is shown in, and whether it is shown without its sign. */
interface SectionChoice {
  /** The section's index, or -1 when the value passes none of the conditions. */
  readonly index: number;
  /** Whether the section writes the magnitude only, as a negative section does. */
  readonly unsigned: boolean;
}

/**
 * The section that applies to `n`, and whether it prints the minus sign.
 *
 * Without conditions the sections are positive;negative;zero, and the negative
 * one writes its own sign or none. With conditions, Excel's own `string value`
 * of a probe (2026-11-20, 83 cells) shows:
 *  - an unconditioned section takes the place a sign would have given it: in a
 *    format of three, the first is `[>0]` and the second `[<0]`, and only the
 *    third is "anything else" — `[>100]"a"0;"b"0;"c"0` reads 5 as c5 and -5
 *    as b5; in a format of two, the first is `[>0]` and the second is the
 *    else. We took every unconditioned section for the else: b5 at 5;
 *  - a section whose own condition admits only negative numbers (`[<0]`,
 *    `[<-1]`, `[<=-1]`, `[=-5]`, the implied `[<0]`) writes no minus sign:
 *    `[<0]"neg "0` reads -5 as "neg 5", but `[<=0]`, `[<1]` and `[<>0]` keep
 *    it, "-n5";
 *  - the else of a two-section format writes none when the first's limit is
 *    zero or below and it is not an equality — `[>0]`, `[>=0]`, `[>-10]`,
 *    `[<-10]`, `[<>-5]` — and keeps it after `[>5]`, `[>0.5]` or any `[=…]`;
 *    the else of three always keeps it.
 * We printed the value's own sign under every condition.
 */
function sectionFor(n: number, sections: ReadonlyArray<string>): SectionChoice {
  // At most three sections take conditions; a fourth is the text section and is
  // never a number's.
  const numeric = Math.min(sections.length, 3);
  const conditions = sections.slice(0, numeric).map(conditionOf);
  if (!conditions.some((c) => c !== undefined)) {
    if (n > 0) return { index: 0, unsigned: false };
    if (n < 0) return { index: sections.length > 1 ? 1 : 0, unsigned: sections.length > 1 };
    return { index: sections.length > 2 ? 2 : 0, unsigned: false };
  }
  for (let i = 0; i < numeric; i++) {
    const c = conditions[i] ?? impliedCondition(i, numeric);
    if (c === undefined) {
      const first = conditions[0];
      const unsigned = numeric === 2 && first !== undefined && first.op !== '=' && first.rhs <= 0;
      return { index: i, unsigned };
    }
    if (conditionHolds(n, c)) return { index: i, unsigned: onlyNegatives(c) };
  }
  return { index: -1, unsigned: false };
}

/** The test an unconditioned section stands for in a format that has conditions. */
function impliedCondition(
  index: number,
  sections: number,
): { op: string; rhs: number } | undefined {
  if (index === 0) return { op: '>', rhs: 0 };
  if (index === 1 && sections === 3) return { op: '<', rhs: 0 };
  return undefined;
}

/** Whether every number that passes the condition is negative. */
function onlyNegatives(c: { op: string; rhs: number }): boolean {
  return (c.op === '<' && c.rhs <= 0) || ((c.op === '<=' || c.op === '=') && c.rhs < 0);
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
  const cleaned = replaceBracketCodes(code.replace(/"[^"]*"/g, ''), () => '');
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
      case 'subsec': {
        // toFixed stops at a hundred places, and a format asking for more is the
        // file's to write; the places past those are zeros.
        const places = Math.min(subsecDigits, 100);
        const fraction = (totalSeconds - Math.floor(totalSeconds)).toFixed(places);
        out += fraction.substring(1) + '0'.repeat(subsecDigits - places);
        break;
      }
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
  // Don't reformat — just trim trailing zeros after a stored decimal…
  // …but a value ECMAScript renders in exponential form comes back spelled
  // JavaScript's way, `3e-104`, and General is Excel's format: an upper-case E
  // and a signed exponent of at least two digits. 57236.xlsx stores three such
  // values and we printed all three with a lower-case e where every other
  // reader prints `3E-104`. A cell in a column writes them through
  // `generalRenderings` below, spelled the same way.
  return excelExponent(String(n));
}

/** Respell ECMAScript's `3e-104` as Excel's `3E-104` (§18.8.31). */
function excelExponent(rendered: string): string {
  const m = /^(-?[\d.]+)e([+-]?)(\d+)$/i.exec(rendered);
  if (!m) return rendered;
  return `${m[1]!}E${m[2] === '-' ? '-' : '+'}${m[3]!.padStart(2, '0')}`;
}

/**
 * The most characters General writes in a cell, its sign apart: however wide
 * the column, 123456789012345 is 1.23457E+14 and 0.123456789012345 is
 * 0.123456789 — Excel's own PDF (2026-10-02).
 */
const GENERAL_CELL_CHARS = 11;

/**
 * A General number in its column: the most precise of the ways General may
 * write it ({@link generalRenderings}) that `fits`.
 *
 * General is not a fixed format: a spreadsheet shows as many places as the
 * column has room for and ROUNDS to that, so 4.3900875881221957 in a column
 * eight characters wide reads 4.390088. We rendered every stored digit and let
 * the cell clip it, which turns the same value into 4.390087 — off by one in
 * the last place shown, with nothing to say a digit was cut. A number too wide
 * for every way of writing it comes back in its finest, and the cell fills
 * with `#` (CellProperties.hashOnOverflow); one under 1 always fits as `0` or
 * `-0`, which is what Excel prints in a column too narrow for anything else.
 *
 * `fits` decides whether a rendering has room. The print model asks it in the
 * document's own terms — the column's width less the padding it carries, in
 * the font's widest digits — because that is the room Excel gives up places
 * against, whatever face the page is drawn in.
 *
 * @param rawValue The cell's stored value.
 * @param fits     Whether a rendering will fit the cell.
 */
export function generalToWidth(rawValue: string, fits: (text: string) => boolean): string {
  const n = Number(rawValue);
  if (!Number.isFinite(n)) return defaultNumberRender(rawValue);
  let widest: string | undefined;
  for (const text of generalRenderings(n, GENERAL_CELL_CHARS)) {
    if (fits(text)) return text;
    widest ??= text;
  }
  return widest ?? defaultNumberRender(rawValue);
}

/**
 * Every way General may write `value` in `chars` characters, the most precise
 * first. One rule, measured in Excel's own PDF on chart axes, which give it
 * nine characters with the sign among them, and in cells, eleven with the
 * sign apart (2026-10-01/02):
 *
 *  - the number as it is, its first 15 significant digits, where that fits:
 *    0.000012345, 0.0000003, 12345678901;
 *  - else rounded to fewer places or written in scientific notation, whichever
 *    keeps the finer last digit — 0.123456789, not 1.23457E-01; 1.23457E-10,
 *    not 0; 4E-04, not 0, in a column five characters wide — and of two as
 *    fine, the plain one down to 0.0001 and the scientific one below it: 0.0001
 *    for 0.000123456 but 1E-05 for 1.00049E-05;
 *  - a whole part longer than the room only in scientific notation, and never
 *    with the zeros a mantissa or a fraction ends in: 1E+10, 2.5E+12.
 *
 * The digits are the 15 of {@link decimalOf}, rounded half away from zero, so
 * 1.2345E-05 is 1.235E-05, where `toExponential` rounded the binary double
 * down to 1.234E-05.
 *
 * @param value      The number.
 * @param chars      How many characters General has.
 * @param signCounts Whether a minus sign takes one of them.
 * @returns The renderings, sign included, finest first.
 */
export function* generalRenderings(
  value: number,
  chars: number,
  signCounts = false,
): Generator<string, void, undefined> {
  if (value === 0 || !Number.isFinite(value)) {
    yield value === 0 ? '0' : String(value);
    return;
  }
  const sign = value < 0 ? '-' : '';
  const room = chars - (signCounts ? sign.length : 0);
  const { digits, exponent } = decimalOf(value);
  let shown = digits.length;
  while (shown > 1 && digits[shown - 1] === '0') shown--;
  const exactPlaces = Math.max(0, shown - 1 - exponent);
  const exact = placeDigits(digits, exponent, exactPlaces);
  if (exact.length <= room) yield sign + exact;
  const forms: Array<{ text: string; place: number; plain: boolean }> = [];
  if (exponent < room) {
    const most = Math.min(exactPlaces, room - Math.max(exponent, 0) - 2);
    for (let places = Math.max(0, most); places >= 0; places--) {
      const text = withoutTrailingZeros(placeDigits(digits, exponent, places));
      if (text.length <= room) forms.push({ text, place: -decimalsIn(text), plain: true });
    }
  }
  const tail = (e: number): string =>
    `E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  for (let places = Math.max(0, room - tail(exponent).length - 2); places >= 0; places--) {
    let e = exponent;
    let mantissa = placeDigits(digits, 0, places);
    // Rounding can carry into a new leading digit, and so into the exponent:
    // 9.9996E+09 is 1E+10.
    if (mantissa.length > 1 && mantissa[1] !== '.') {
      e += 1;
      mantissa = placeDigits('1', 0, places);
    }
    const stripped = withoutTrailingZeros(mantissa);
    const text = stripped + tail(e);
    if (text.length <= room) forms.push({ text, place: e - decimalsIn(stripped), plain: false });
  }
  const plainFirst = exponent >= -4;
  forms.sort(
    (a, b) => a.place - b.place || (a.plain === b.plain ? 0 : a.plain === plainFirst ? -1 : 1),
  );
  const seen = new Set<string>(exact.length <= room ? [exact] : []);
  for (const form of forms) {
    if (seen.has(form.text)) continue;
    seen.add(form.text);
    yield sign + form.text;
  }
}

/** A decimal without the zeros that end its fraction, nor a point left bare. */
function withoutTrailingZeros(fixed: string): string {
  if (!fixed.includes('.')) return fixed;
  let end = fixed.length;
  while (fixed[end - 1] === '0') end--;
  return fixed.slice(0, fixed[end - 1] === '.' ? end - 1 : end);
}

/** How many places a plain decimal shows after its point. */
const decimalsIn = (fixed: string): number => {
  const dot = fixed.indexOf('.');
  return dot < 0 ? 0 : fixed.length - dot - 1;
};

function applyFormatString(rawValue: string, format: string): string {
  const n = Number(rawValue);
  const sections = splitSections(format);

  if (!Number.isFinite(n)) return formatText(rawValue, sections);

  const choice = sectionFor(n, sections);
  if (choice.index < 0) {
    // A value no condition admits: a format of one section shows it as General
    // without its sign — Excel reads -5 under `[>0]"a"0` as 5 and 5.5 under
    // `[<0]"a"0.00` as 5.5 (2026-11-20). A format of two or three fills the
    // cell with `#`, which needs the cell's width; we write the number.
    return Math.min(sections.length, 3) === 1
      ? defaultNumberRender(String(Math.abs(n)))
      : defaultNumberRender(rawValue);
  }
  const section = sections[choice.index] ?? sections[0]!;
  return applyNumericSection(n, section, choice.unsigned);
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

/**
 * The decimal a spreadsheet takes a double for: its first 15 significant
 * digits, and the power of ten the first of them stands at.
 *
 * Excel shows a number as those 15 digits and zeros after them, however many
 * places its format asks for: 0.1 under thirty places is 0.1 and twenty-nine
 * zeros, 1/3 is fifteen threes and zeros, 1E+21 under `0.00` is a 1, twenty-one
 * zeros and `.00` — its own PDF of each (2026-10-01).
 */
function decimalOf(value: number): { digits: string; exponent: number } {
  const [mantissa, exponent] = Math.abs(value).toExponential(14).split('e');
  return { digits: mantissa!.replace('.', ''), exponent: Number(exponent) };
}

/**
 * Those digits at `decimals` places — the first standing at `exponent` —
 * rounded half away from zero, as plain digits with a point and no sign.
 */
function placeDigits(digits: string, exponent: number, decimals: number): string {
  // How many of the digits fall before the cut; the rest decide the rounding.
  const kept = exponent + 1 + decimals;
  let units: string;
  if (kept >= digits.length) units = digits + '0'.repeat(kept - digits.length);
  else if (kept < 0) units = '';
  else units = digits[kept]! >= '5' ? incremented(digits.slice(0, kept)) : digits.slice(0, kept);
  const padded = units.replace(/^0+/, '').padStart(decimals + 1, '0');
  const whole = padded.slice(0, padded.length - decimals);
  return decimals > 0 ? `${whole}.${padded.slice(whole.length)}` : whole;
}

/** A string of decimal digits plus one. */
function incremented(digits: string): string {
  let at = digits.length - 1;
  while (at >= 0 && digits[at] === '9') at--;
  if (at < 0) return `1${'0'.repeat(digits.length)}`;
  return `${digits.slice(0, at)}${String(Number(digits[at]) + 1)}${'0'.repeat(digits.length - at - 1)}`;
}

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
 * The digits are the 15 of {@link decimalOf} and zeros after them, placed by
 * hand: past seventeen significant digits `toFixed` printed the binary
 * expansion (0.1 under twenty places as 0.10000000000000000555), from 1E+21 on
 * it wrote the exponent (`1e+21.00`), and past a hundred places it threw.
 */
function toFixedDecimal(value: number, decimals: number): string {
  if (!Number.isFinite(value)) return String(value);
  const { digits, exponent } = decimalOf(value);
  const placed = placeDigits(digits, exponent, decimals);
  return value < 0 && /[1-9]/.test(placed) ? `-${placed}` : placed;
}

/**
 * §18.8.31 scientific notation: `0.00E+00`, `##0.0E+0` (engineering), `0.0e-0`.
 * The mantissa is normalised so its integer part holds `intDigits` significant
 * figures (1 for `0.00E+00`, 3 for `##0.0E+0` → exponent snaps to a multiple of
 * 3); the exponent is zero-padded to the placeholder count and carries a sign
 * (always for `E+`, only when negative for `E-`). The `E`/`e` case is preserved.
 *
 * The mantissa is laid out by the same placeholders as any number, and the
 * text around it stays: Excel's own `string value` of a probe (2026-11-20)
 * reads `00.0E+00` at 12345 as 01.2E+04, `"x"0.0E+0` at 123 as x1.2E+2,
 * `0.0E+0" m"` at 1234 as 1.2E+3 m and `$0.0E+0` at -123 as -$1.2E+2 — where
 * we wrote 1.2E+04, dropped the literals and put the minus after the `$`.
 */
function formatScientific(
  value: number,
  tokens: ReadonlyArray<NumberToken>,
  at: number,
  negativeSection: boolean,
): string {
  const mantissa = tokens.slice(0, at);
  const marker = tokens[at] as ExponentToken;
  const after = tokens.slice(at + 1);
  let expPlaces = 0;
  while (after[expPlaces]?.kind === 'digit') expPlaces++;

  const point = mantissa.findIndex((t) => t.kind === 'point');
  const split = point < 0 ? mantissa.length : point;
  let intPlaces = 0;
  let decimals = 0;
  mantissa.forEach((t, i) => {
    if (t.kind !== 'digit') return;
    if (i < split) intPlaces++;
    else decimals++;
  });
  const intDigits = Math.max(1, intPlaces);

  // The mantissa is the decimal's own digits, rounded half away from zero like
  // any other number shown: Excel prints 1234.5 under `0.000E+00` as 1.235E+03
  // and 1.45 under `0.0E+00` as 1.5E+00, where the binary mantissa's `toFixed`
  // gave 1.234 and 1.4.
  let exp = 0;
  let mantStr = placeDigits('0', 0, decimals);
  if (value !== 0) {
    const { digits, exponent } = decimalOf(value);
    exp = exponent - (((exponent % intDigits) + intDigits) % intDigits); // engineering grouping
    mantStr = placeDigits(digits, exponent - exp, decimals);
    // Rounding the mantissa can carry it up to 10^intDigits — renormalise.
    if (mantStr.split('.')[0]!.length > intDigits) {
      exp += intDigits;
      mantStr = placeDigits(digits, exponent - exp, decimals);
    }
  }

  const [mantInt, mantDec] = mantStr.split('.');
  const body = layOut(
    mantissa,
    commaRoles(mantissa),
    mantInt === '0' ? '' : mantInt!,
    mantDec ?? '',
  );
  const expStr = String(Math.abs(exp)).padStart(expPlaces || 2, '0');
  const expSign = exp < 0 ? '-' : marker.sign === '+' ? '+' : '';
  const signPrefix = value < 0 && !negativeSection ? '-' : '';
  return `${signPrefix}${body}${marker.e}${expSign}${expStr}${tokensAsText(after.slice(expPlaces))}`;
}

/**
 * Resolve the `[...]` codes a numeric section may carry (§18.8.31).
 *
 * `[$SYMBOL-LOCALE]` is a currency tag and its SYMBOL is *rendered* — dropping
 * the whole bracket, as a colour or a locale is dropped, silently loses the `$`
 * from `[$$-409]#,##0`. The symbol becomes a quoted literal so the placeholder
 * grammar treats it as text. Everything else in brackets (colours, locales,
 * conditions) is decoration here and goes.
 *
 * A bracket inside quotes or after a backslash is text, not a code: Excel
 * reads 5 under `0" [kg]"` as "5 [kg]", under `\[0\]` as [5] and "abc" under
 * `0;-0;0;"["@"]"` as [abc] (2026-11-20). Taken for codes, they went — the
 * last of them with the `@` inside, leaving nothing at all.
 */
function resolveBracketCodes(format: string): string {
  // No code opens after the last `]`, so a `[` past it is text, and is not
  // searched for its end again and again.
  const lastClose = format.lastIndexOf(']');
  let out = '';
  for (let i = 0; i < format.length; i++) {
    const ch = format[i]!;
    if (ch === '"') {
      const close = format.indexOf('"', i + 1);
      const end = close < 0 ? format.length - 1 : close;
      out += format.slice(i, end + 1);
      i = end;
    } else if (ch === '\\' && i + 1 < format.length) {
      out += format.slice(i, i + 2);
      i++;
    } else if (ch === '[' && i < lastClose) {
      const close = format.indexOf(']', i);
      const currency = /^\$([^-]*)(?:-.*)?$/.exec(format.slice(i + 1, close));
      const symbol = currency?.[1]?.replace(/"/g, '') ?? '';
      if (symbol.length > 0) out += `"${symbol}"`;
      i = close;
    } else {
      out += ch;
    }
  }
  return out;
}

const isPlaceholder = (c: string | undefined): boolean => c === '0' || c === '#' || c === '?';
const isAsciiDigit = (c: string | undefined): boolean => c !== undefined && c >= '0' && c <= '9';
const isSpace = (c: string | undefined): boolean => c !== undefined && /\s/.test(c);

// The four line ends `.` does not match.
const isLineEnd = (c: number): boolean => c === 0x0a || c === 0x0d || c === 0x2028 || c === 0x2029;

function hasLineEnd(s: string): boolean {
  for (let i = 0; i < s.length; i++) if (isLineEnd(s.charCodeAt(i))) return true;
  return false;
}

/**
 * A fraction section cut into head, numerator, denominator and tail — what
 * `/^(.*?)([0#?]+)\s*\/\s*([0#?]+|\d+)(.*)$/` makes of it, in one pass. The
 * expression, lazy at the front, took every position of a run of placeholders
 * in turn as the start of the numerator and read the rest of the run from
 * each: quadratic in a long run no slash follows. A start inside a run sees
 * the same rest as the run's first, so only the first is tried. The head
 * cannot cross a line end, nor the tail hold one.
 */
function fractionParts(
  s: string,
): { head: string; num: string; den: string; tail: string } | undefined {
  let firstEnd = s.length;
  let lastEnd = -1;
  for (let i = 0; i < s.length; i++) {
    if (!isLineEnd(s.charCodeAt(i))) continue;
    firstEnd = Math.min(firstEnd, i);
    lastEnd = i;
  }
  for (let start = 0; start < firstEnd; ) {
    if (!isPlaceholder(s[start])) {
      start++;
      continue;
    }
    let numEnd = start;
    while (isPlaceholder(s[numEnd])) numEnd++;
    let slash = numEnd;
    while (isSpace(s[slash])) slash++;
    if (s[slash] === '/') {
      let den = slash + 1;
      while (isSpace(s[den])) den++;
      let denEnd = den;
      const inDen = isPlaceholder(s[den]) ? isPlaceholder : isAsciiDigit;
      while (inDen(s[denEnd])) denEnd++;
      if (denEnd > den && lastEnd < denEnd) {
        return {
          head: s.slice(0, start),
          num: s.slice(start, numEnd),
          den: s.slice(den, denEnd),
          tail: s.slice(denEnd),
        };
      }
    }
    start = numEnd;
  }
  return undefined;
}

/**
 * The last run of placeholders in the head and what follows it — what
 * `/([0#?]+)([^0#?]*)$/` matches, read back from the end: that expression
 * retried every placeholder of an earlier run, and is quadratic in a long one.
 */
function lastPlaceholderRun(head: string): { at: number; after: string } | undefined {
  let end = head.length;
  while (end > 0 && !isPlaceholder(head[end - 1])) end--;
  if (end === 0) return undefined;
  let at = end - 1;
  while (at > 0 && isPlaceholder(head[at - 1])) at--;
  return { at, after: head.slice(end) };
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
  const m = fractionParts(cleaned);
  if (!m) return undefined;
  const { head, num: numFmt, den: denFmt, tail } = m;

  // The integer part, if any, is the LAST placeholder run in the head — whatever
  // separates it from the fraction is a literal, and need not be a bare space:
  // formats.xlsx writes `#"  "?/?`, quoting two spaces. Anchored at the end of
  // the head, the run went unfound there, so 1.2 came out as the improper
  // `#  6/5` — the `#` printed as itself and the whole number folded into the
  // numerator — where every reader shows `1  1/5`.
  const intMatch = lastPlaceholderRun(head);
  const hasInteger = intMatch !== undefined;
  const literalPrefix = unquoteLiteral(head.substring(0, intMatch ? intMatch.at : head.length));
  const intFmt = intMatch ? head.slice(intMatch.at, head.length - intMatch.after.length) : '';
  const separator = intMatch ? unquoteLiteral(intMatch.after) : ' ';

  const magnitude = Math.abs(value);
  const whole = hasInteger ? Math.floor(magnitude) : 0;
  const rest = magnitude - whole;

  // A denominator spelled in digits is fixed (`?/16`) — but only from a digit
  // that is not a placeholder: `00/00` is two placeholders a side, and Excel
  // reads 1.5 under `# 00/00` as 1 01/02 (2026-11-20), where taking it for a
  // denominator of 0 dropped the whole fraction.
  const fixedDen = /^[1-9]\d*$/.test(denFmt) ? Number(denFmt) : undefined;
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
  // Rounding can carry the fraction to a whole unit; without an integer field
  // to carry into (`?/?` is an improper fraction) it stays in the numerator.
  let integer = whole;
  if (hasInteger && den > 0 && num >= den) {
    integer += Math.floor(num / den);
    num %= den;
  }
  // Every placeholder keeps its width, as it does in any number: Excel's own
  // `string value` of a probe (2026-11-20) reads `# ??/??` at 3.14159 as
  // "3  1/7 " — the numerator right-aligned in its two places, the denominator
  // left-aligned in its two — `# ?/?` at 0.25 as " 1/4", the integer blank but
  // its separator kept, and at 3 or at 0 as "3    " and "0    ": a fraction
  // that rounds away leaves blanks the width of what it would have printed,
  // and the integer is shown even at zero. `# 00/00` pads with zeros, 1 01/02.
  // We wrote 3 1/7, 1/4, 3 and 0, so a column of them no longer lined up.
  const showFraction = !hasInteger || num !== 0;
  const intText = hasInteger
    ? padPlaceholders(intFmt, showFraction && integer === 0 ? '' : String(integer))
    : '';
  const gap = separator.length > 0 ? separator : ' ';
  const fractionText = showFraction
    ? `${hasInteger ? gap : ''}${padPlaceholders(numFmt, String(num))}/${
        fixedDen !== undefined ? denFmt : padDenominator(denFmt, String(den))
      }`
    : ' '.repeat(gap.length + numFmt.length + 1 + denFmt.length);

  // The minus leads the whole text, literals and all: -$1 1/2 for `"$"# ?/?`.
  // With an integer field it is there even when the value rounds to nothing
  // (`# ?/?` at -0.01 reads "-0    "); without one, only when a numerator is
  // (`?/?` at -0.01 reads "0/1").
  const sign = value < 0 && !negativeSection && (hasInteger || num !== 0) ? '-' : '';
  return `${sign}${literalPrefix}${intText}${fractionText}${unquoteLiteral(tail)}`;
}

/**
 * Digits right-aligned in a run of placeholders: a place the digits leave over
 * shows `0` under `0`, a space under `?` and nothing under `#`; digits the run
 * cannot hold go in front of it.
 */
function padPlaceholders(placeholders: string, digits: string): string {
  let pad = '';
  for (let i = 0; i < placeholders.length - digits.length; i++) {
    const ch = placeholders[i];
    if (ch === '0') pad += '0';
    else if (ch === '?') pad += ' ';
  }
  return pad + digits;
}

/**
 * A fraction's denominator in its placeholders. Unlike every other digit run it
 * is LEFT-aligned under `?` — Excel writes 1/7 under `??/??` as " 1/7 " — while
 * `0` still pads in front of it, 1/2 under `00/00` reading 01/02 (2026-11-20).
 */
function padDenominator(placeholders: string, digits: string): string {
  let zeros = '';
  let spaces = '';
  for (let i = digits.length; i < placeholders.length; i++) {
    if (placeholders[i] === '0') zeros += '0';
    else if (placeholders[i] === '?') spaces += ' ';
  }
  return zeros + digits + spaces;
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
  // …and General survives being bracketed. `[DBNum1][$-804]General` is General
  // with a locale and a numeral system on it; matched only against the whole
  // code, it fell through to the placeholder grammar, which finds no digit
  // placeholder and renders the keyword itself — 49273.xlsx printed
  // "General12323". The numeral system is a separate, much larger feature; the
  // bug is printing the word.
  if (cleaned.trim().toLowerCase() === 'general') {
    return defaultNumberRender(String(negativeSection ? Math.abs(value) : value));
  }

  const fraction = formatFraction(value, cleaned, negativeSection);
  if (fraction !== undefined) return fraction;

  const tokens = tokenizeSection(cleaned);
  // A minus sign leads the whole text, in front of every literal the format
  // writes — Excel's own `string value` of a probe (2026-11-20) reads `$#,##0`
  // at -5 as -$5, `"Total: "#,##0` as -Total: 1,234, `(###) ###-####` as
  // -(555) 123-4567 and `_(0_)` as "- 5 ". We wrote the literals first: $-5.
  const signed = value < 0 && !negativeSection;

  // General with text around it is still General: `General" kg"` reads 5 kg,
  // `"$"General` at -5 reads -$5. Matched only as the whole code, the keyword
  // printed itself.
  if (tokens.some((t) => t.kind === 'general')) {
    const general = defaultNumberRender(String(Math.abs(value)));
    const text = tokens.map((t) => (t.kind === 'general' ? general : tokenText(t))).join('');
    return `${signed && value !== 0 ? '-' : ''}${text}`;
  }
  // A text section (`"a"@"b"`) given a number shows it as General, its
  // literals and all dropped: Excel reads 5 there as 5 and -5 under `"a"@` as -5.
  if (tokens.some((t) => t.kind === 'text')) {
    return `${signed ? '-' : ''}${defaultNumberRender(String(Math.abs(value)))}`;
  }

  const exponent = tokens.findIndex((t) => t.kind === 'exponent');
  if (exponent >= 0) {
    // A code with a line end in it is no code Excel writes, and we print it
    // back as it stands, as we always have.
    if (hasLineEnd(cleaned)) return cleaned;
    return formatScientific(value, tokens, exponent, negativeSection);
  }

  const roles = commaRoles(tokens);
  // §18.8.31 — `%` multiplies by a hundred, once for each: Excel reads `0%%` at
  // 0.5 as 5000%%. A QUOTED or escaped percent sign is text and scales nothing:
  // `0"%"` and `0\%` read 1%, where we multiplied for any `%` in the code and
  // printed 50%.
  let magnitude = Math.abs(value);
  for (const t of tokens) if (t.kind === 'percent') magnitude *= 100;
  // §18.8.31 — a comma against the last placeholder is not punctuation, it is a
  // scale: the value is shown in thousands (or millions, per comma).
  if (roles.scale > 0) magnitude /= 1000 ** roles.scale;

  let decimals = 0;
  const point = tokens.findIndex((t) => t.kind === 'point');
  if (point >= 0)
    for (let i = point + 1; i < tokens.length; i++) if (tokens[i]!.kind === 'digit') decimals++;
  const fixed = toFixedDecimal(magnitude, decimals);
  const [intRaw, decRaw] = fixed.split('.');
  // §18.8.31: `#` is an *optional* digit, so an integer part of zero written
  // only with `#` renders as nothing at all — format `#` blanks a zero cell
  // (tdf171828 uses that to hide a whole column of them) and `#.##` shows ".5"
  // rather than "0.5". The placeholders decide what a zero shows.
  const whole = intRaw === '0' ? '' : intRaw!;
  // …and a value that rounds to zero shows no sign: Excel reads -0.4 under `0`
  // as 0, -0.04 under `0.0` as 0.0 and -0.004 under `0%` as 0%.
  const minus = signed && /[1-9]/.test(fixed) ? '-' : '';
  return minus + layOut(tokens, roles, whole, decRaw ?? '');
}

/** The `E+` / `e-` of a scientific code. */
interface ExponentToken {
  readonly kind: 'exponent';
  readonly e: string;
  readonly sign: string;
}

/** One element of a number format's section (§18.8.31). */
type NumberToken =
  | { readonly kind: 'digit'; readonly ch: string }
  | { readonly kind: 'point' }
  | { readonly kind: 'comma' }
  | { readonly kind: 'percent' }
  | ExponentToken
  | { readonly kind: 'general' }
  | { readonly kind: 'text' }
  | { readonly kind: 'literal'; readonly text: string };

/**
 * A section cut into its placeholders and its text, in order: quoted text and
 * `\x` are literal, `_x` is a space the width of x, `*x` (fill the cell with x)
 * is nothing — repeating it needs the laid-out width this layer does not have.
 * Only the first `.` is the decimal point.
 */
function tokenizeSection(code: string): Array<NumberToken> {
  const tokens: Array<NumberToken> = [];
  let point = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i]!;
    const next = code[i + 1];
    if (ch === '"') {
      const close = code.indexOf('"', i + 1);
      const end = close < 0 ? code.length : close;
      tokens.push({ kind: 'literal', text: code.slice(i + 1, end) });
      i = end;
    } else if (ch === '\\' && next !== undefined) {
      tokens.push({ kind: 'literal', text: next });
      i++;
    } else if (ch === '_' && next !== undefined) {
      tokens.push({ kind: 'literal', text: ' ' });
      i++;
    } else if (ch === '*' && next !== undefined) {
      i++;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      tokens.push({ kind: 'digit', ch });
    } else if (ch === '.' && !point) {
      point = true;
      tokens.push({ kind: 'point' });
    } else if (ch === ',') {
      tokens.push({ kind: 'comma' });
    } else if (ch === '%') {
      tokens.push({ kind: 'percent' });
    } else if ((ch === 'E' || ch === 'e') && (next === '+' || next === '-')) {
      tokens.push({ kind: 'exponent', e: ch, sign: next });
      i++;
    } else if ((ch === 'G' || ch === 'g') && code.slice(i, i + 7).toLowerCase() === 'general') {
      tokens.push({ kind: 'general' });
      i += 6;
    } else if (ch === '@') {
      tokens.push({ kind: 'text' });
    } else {
      tokens.push({ kind: 'literal', text: ch });
    }
  }
  return tokens;
}

/** A token as the text it prints when it is not a number's own. */
function tokenText(t: NumberToken): string {
  switch (t.kind) {
    case 'literal':
      return t.text;
    case 'digit':
      return t.ch;
    case 'point':
      return '.';
    case 'comma':
      return ',';
    case 'percent':
      return '%';
    case 'exponent':
      return t.e + t.sign;
    default:
      return '';
  }
}

function tokensAsText(tokens: ReadonlyArray<NumberToken>): string {
  return tokens.map(tokenText).join('');
}

/** What each comma of a section does, and what they come to. */
interface CommaRoles {
  readonly roles: ReadonlyArray<'group' | 'scale' | 'drop' | 'literal' | undefined>;
  /** Whether the integer digits are grouped in thousands. */
  readonly group: boolean;
  /** How many thousands the value is divided by. */
  readonly scale: number;
}

/**
 * §18.8.31 gives a comma two meanings; Excel's own `string value` of a probe
 * (2026-11-20) shows four:
 *  - between two placeholders of the integer it GROUPS the whole integer in
 *    thousands, wherever it stands: `#,#` reads 1234567 as 1,234,567;
 *  - a run of them right after the last placeholder of its part SCALES by a
 *    thousand each, before the point (`0,.0` reads 12345 as 12.3), after the
 *    decimals (`0.0,` reads 1234 as 1.2) or before text (`0, "K"` reads 12 K);
 *  - right after a placeholder that more of its part follow, it does nothing:
 *    `0,"-"0` reads 1234567 as 123456-7 and `0.0,0` reads 1234.5 as 1234.50;
 *  - anywhere else it is a comma: `0 ,` reads 12345 as "12345 ,".
 */
function commaRoles(tokens: ReadonlyArray<NumberToken>): CommaRoles {
  const point = tokens.findIndex((t) => t.kind === 'point');
  const split = point < 0 ? tokens.length : point;
  let lastIntDigit = -1;
  let lastDecDigit = -1;
  tokens.forEach((t, i) => {
    if (t.kind !== 'digit') return;
    if (i < split) lastIntDigit = i;
    else lastDecDigit = i;
  });
  const roles: Array<'group' | 'scale' | 'drop' | 'literal' | undefined> = [];
  let group = false;
  let scale = 0;
  for (let a = 0; a < tokens.length; a++) {
    if (tokens[a]!.kind !== 'comma') continue;
    let b = a;
    while (tokens[b + 1]?.kind === 'comma') b++;
    const inInteger = b < split;
    let role: 'group' | 'scale' | 'drop' | 'literal';
    if (tokens[a - 1]?.kind !== 'digit') role = 'literal';
    else if (tokens[b + 1]?.kind === 'digit') role = inInteger ? 'group' : 'drop';
    else if ((inInteger ? lastIntDigit : lastDecDigit) < a) role = 'scale';
    else role = 'drop';
    for (let i = a; i <= b; i++) roles[i] = role;
    if (role === 'group') group = true;
    if (role === 'scale') scale += b - a + 1;
    a = b;
  }
  return { roles, group, scale };
}

/**
 * Write a number into a section's tokens. The integer digits go right to left
 * over the placeholders before the point, whatever text stands between them,
 * and any the placeholders cannot hold go in front of the first; the decimals
 * go left to right over those after it. A place the digits leave over shows
 * `0` under `0`, a space under `?` and nothing under `#` (§18.8.31) — `?` is
 * there so a column lines up: Excel's Accounting format writes its zero as
 * `_-* "-"??_-`, and with `?` read as a literal every such zero printed "-??".
 *
 * Excel's own `string value` of a probe (2026-11-20) is the evidence:
 * `000-00-0000` reads 123456789 as 123-45-6789 and 1234567890 as 1234-56-7890,
 * `(###) ###-####` reads 1234567 as "() 123-4567", `0.0"x"0` reads 1.23 as
 * 1.2x3. Only the outermost text survived before: (5551234567.
 *
 * Thousands are counted over the whole integer, so the separator after the
 * fourth digit from the right lands in front of text that follows it:
 * `#,##0"-"000` reads 1234567 as 1,234,-567. A separator beside a place left
 * blank is blank too, a space beside a `?`: `?,???` reads 5 as "    5".
 *
 * The decimal point is always written, even with nothing after it: `#.##`
 * reads 1 as "1." and 0 as ".", `0.###` reads 2 as "2." — we dropped it.
 *
 * @param whole    The integer's significant digits, '' for a zero.
 * @param decimals Exactly as many digits as there are decimal placeholders.
 */
function layOut(
  tokens: ReadonlyArray<NumberToken>,
  { roles, group }: CommaRoles,
  whole: string,
  decimals: string,
): string {
  const point = tokens.findIndex((t) => t.kind === 'point');
  const split = point < 0 ? tokens.length : point;
  const intSlots: Array<number> = [];
  const decSlots: Array<number> = [];
  tokens.forEach((t, i) => {
    if (t.kind === 'digit') (i < split ? intSlots : decSlots).push(i);
  });
  const out = tokens.map((t, i) =>
    t.kind === 'comma' ? (roles[i] === 'literal' ? ',' : '') : tokenText(t),
  );
  // What follows a place that shows `shown`, `pos` places from the right.
  const separator = (pos: number, shown: string): string =>
    !group || pos === 0 || pos % 3 !== 0 ? '' : shown === ' ' ? ' ' : shown === '' ? '' : ',';

  const places = intSlots.length;
  for (let k = 0; k < places; k++) {
    const pos = places - 1 - k;
    let text = '';
    if (k === 0) {
      for (let j = 0; j < whole.length - places; j++) {
        text += whole[j]! + separator(whole.length - 1 - j, whole[j]!);
      }
    }
    const ch = (tokens[intSlots[k]!] as { ch: string }).ch;
    const shown =
      pos < whole.length
        ? whole[whole.length - 1 - pos]!
        : ch === '0'
          ? '0'
          : ch === '?'
            ? ' '
            : '';
    out[intSlots[k]!] = text + shown + separator(pos, shown);
  }
  // No placeholder before the point: the integer is still written, in front
  // of it — `.00` reads 12.5 as 12.50.
  if (places === 0 && point >= 0) out[point] = `${whole}.`;

  // '0' forces its digit, '#' drops it when the whole tail from there on is
  // zeros — so `#.##` renders 0.5 as ".5", 0.55 as ".55" and 0.05 as ".05" —
  // and '?' turns such a zero into a space. The decision is per position and
  // reads rightwards, which is why the tail is settled after the fact.
  const shown = decSlots.map((_, i) => decimals[i] ?? '0');
  for (let i = shown.length - 1; i >= 0 && shown[i] === '0'; i--) {
    const ch = (tokens[decSlots[i]!] as { ch: string }).ch;
    if (ch === '#') shown[i] = '';
    else if (ch === '?') shown[i] = ' ';
    else break;
  }
  decSlots.forEach((slot, i) => (out[slot] = shown[i]!));
  return out.join('');
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
