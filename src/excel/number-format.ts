// ECMA-376 Part 1 §18.8.30 — Built-in number formats and §18.8.31 — custom
// format codes. Implements the subset typical business spreadsheets use:
//   - Built-in: 0..4, 9..10, 37..40, 49 (General / integer / decimal /
//     thousands / percent / accounting basics / text)
//   - Custom format strings with `0`, `#`, `,`, `.`, `%`, quoted literals,
//     and `[colour]` / `[locale]` codes (the latter two stripped silently)
// Dates (m/d/yyyy etc.) are deferred — the cell value is shown verbatim.

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
  if (custom !== undefined && isDateFormat(custom))
    return formatExcelDate(rawValue, custom, date1904);

  const format = custom ?? BUILTIN.get(numFmtId);
  if (!format) return defaultNumberRender(rawValue);
  if (format === 'General') return defaultNumberRender(rawValue);

  return applyFormatString(rawValue, format);
}

// Excel serial date → JS Date. The default 1900 epoch uses 1899-12-30 as day 0
// (so serial 1 is 1900-01-01) and inherits the Lotus 1-2-3 leap-year bug:
// serial 60 is considered "1900-02-29" which never existed. For serial ≥ 61
// the simple formula is exact; values < 60 are vanishingly rare in business
// sheets (and our render is approximate anyway).
//
// The 1904 epoch (legacy Mac Excel) uses 1904-01-01 as day 0 and has no leap
// bug. Files saved with <workbookPr date1904="1"/> store dates offset by
// exactly 1462 days from the 1900-epoch interpretation.
export function excelSerialToDate(serial: number, date1904: boolean): Date {
  const ms = serial * 86400 * 1000;
  if (date1904) return new Date(Date.UTC(1904, 0, 1) + ms);
  return new Date(Date.UTC(1899, 11, 30) + ms);
}

// The inverse over a UTC calendar date: (year, month0, day) → the integer Excel
// serial day, using the same epoch. Round-trips excelSerialToDate exactly for an
// integer serial (the time-of-day is zero), so a serial → parts → serial loop is
// stable. Used by the formula engine's date functions and the timePeriod windows
// (E-SHEET W9) to map an injected reference date into serial space.
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

function isDateFormat(code: string): boolean {
  // Strip quoted literals and [] codes, then look for any date token. The
  // "m" letter alone is ambiguous (month vs minute) so it can't be a sole
  // signal, but its presence alongside d/y/h/s already implies dates.
  const cleaned = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  return /[dyhs]|m+/i.test(cleaned) && /[dyhs]/i.test(cleaned);
}

interface DateToken {
  readonly kind: 'lit' | 'y' | 'M' | 'd' | 'h' | 'm' | 's' | 'ampm' | 'elapsed-h';
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
      // Anything else in brackets (colors, locales) is ignored.
      continue;
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
    const prevIsHour = prev >= 0 && resolved[prev]!.kind === 'h';
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

  const date = excelSerialToDate(serial, date1904);
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth() + 1; // 1-12
  const day = date.getUTCDate();
  const weekday = date.getUTCDay();
  const hour24 = date.getUTCHours();
  const minute = date.getUTCMinutes();
  const second = date.getUTCSeconds();

  const tokens = resolveMonthVsMinute(tokenizeDateFormat(format));
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
        // [h] = elapsed hours since serial 0; approximate as serial*24.
        out += String(Math.floor(serial * 24));
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

  let sectionIdx: number;
  if (n > 0) sectionIdx = 0;
  else if (n < 0) sectionIdx = sections.length > 1 ? 1 : 0;
  else sectionIdx = sections.length > 2 ? 2 : 0;

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

function applyNumericSection(value: number, format: string, negativeSection: boolean): string {
  // Strip [colour] / [locale] codes.
  const cleaned = format.replace(/\[[^\]]*\]/g, '');

  if (/[eE][+-]/.test(cleaned)) return formatScientific(value, cleaned, negativeSection);

  const isPercent = cleaned.includes('%');
  let magnitude = Math.abs(value);
  if (isPercent) magnitude *= 100;

  const { intFormat, decFormat, literalPrefix, literalSuffix } = splitNumberFormat(cleaned);
  let decimals = 0;
  for (const c of decFormat) if (c === '0' || c === '#') decimals++;

  const fixed = magnitude.toFixed(decimals);
  const [intRaw, decRaw] = fixed.split('.');
  const useThousands = /,(?=\d)/.test(intFormat) || /[0#],[0#]/.test(intFormat);
  const intStr = useThousands ? intRaw!.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : intRaw!;

  let numberPart = intStr;
  if (decimals > 0) {
    let dec = decRaw ?? ''.padEnd(decimals, '0');
    if (dec.length < decimals) dec = dec.padEnd(decimals, '0');
    // Trim '#' placeholders → optional trailing zeros stripped, '0' kept.
    let kept = '';
    for (let i = 0; i < decFormat.length; i++) {
      const placeholder = decFormat[i]!;
      const digit = dec[i] ?? '0';
      if (placeholder === '0') kept += digit;
      else if (placeholder === '#') {
        if (digit !== '0' || kept.length > 0) kept += digit;
      }
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
    return { literalPrefix: cleaned, intFormat: '', decFormat: '', literalSuffix: '' };
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
    if (ch === '%') {
      // handled by isPercent at caller
      continue;
    }
    out += ch;
  }
  return out;
}
