// Date helpers (E-SHEET W9) — serial ↔ calendar conversions for the formula
// engine's date functions (TODAY/YEAR/MONTH/DAY/WEEKDAY/DATE/EDATE/EOMONTH) and
// the conditional-format `timePeriod` windows. All conversions go through the
// shared epoch math in number-format.ts so cell serials and the injected
// reference date live in exactly one serial space (1900 or 1904). Everything is
// UTC-based and deterministic — no host time zone, no wall clock.

import type { TimePeriodKind } from '@/core/spreadsheet-model';

import { excelSerialFromUtcParts, excelSerialToDate } from '@/excel/number-format';

/** The UTC calendar fields an Excel serial decomposes into. */
export interface DateParts {
  readonly year: number;
  /** 1-indexed month (January = 1). */
  readonly month: number;
  readonly day: number;
  /** Day of week, 0 = Sunday … 6 = Saturday (Excel WEEKDAY type 1 minus 1). */
  readonly dow: number;
}

/**
 * Decompose an Excel serial into UTC calendar parts. A fractional serial is
 * floored to its day first (the calendar date is time-of-day independent).
 */
export function serialToParts(serial: number, date1904: boolean): DateParts {
  const d = excelSerialToDate(Math.floor(serial), date1904);
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    dow: d.getUTCDay(),
  };
}

/**
 * A UTC Date's calendar day → integer Excel serial. Reads the Date's UTC date
 * parts (not its instant), so a caller passing `new Date('2026-06-17')`
 * (UTC midnight) maps to that day regardless of the host time zone.
 */
export function serialFromDate(d: Date, date1904: boolean): number {
  return excelSerialFromUtcParts(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), date1904);
}

/**
 * `(year, month1, day)` → integer serial. Excel's `DATE` rolls out-of-range
 * months and days (`DATE(2024,13,1)` = 2025-01-01); `Date.UTC` already
 * normalises, so this inherits the same behaviour.
 */
export function serialFromYmd(
  year: number,
  month1: number,
  day: number,
  date1904: boolean,
): number {
  return excelSerialFromUtcParts(year, month1 - 1, day, date1904);
}

/**
 * §18.3.1.10 `timePeriod` — does a cell's date fall in the window relative to
 * `nowSerial` (the injected reference day)? Windows match Excel's built-in
 * rules: the week runs Sunday..Saturday; "last 7 days" is today and the previous
 * six. Both serials are floored to whole days before comparison.
 */
export function timePeriodMatches(
  period: TimePeriodKind,
  cellSerial: number,
  nowSerial: number,
  date1904: boolean,
): boolean {
  const cell = Math.floor(cellSerial);
  const today = Math.floor(nowSerial);
  switch (period) {
    case 'today':
      return cell === today;
    case 'yesterday':
      return cell === today - 1;
    case 'tomorrow':
      return cell === today + 1;
    case 'last7Days':
      return today - cell >= 0 && today - cell <= 6;
    case 'thisWeek':
    case 'lastWeek':
    case 'nextWeek': {
      const weekStart = today - serialToParts(today, date1904).dow;
      const shift = period === 'lastWeek' ? -7 : period === 'nextWeek' ? 7 : 0;
      const lo = weekStart + shift;
      return cell >= lo && cell <= lo + 6;
    }
    case 'thisMonth':
    case 'lastMonth':
    case 'nextMonth': {
      const t = serialToParts(today, date1904);
      const delta = period === 'lastMonth' ? -1 : period === 'nextMonth' ? 1 : 0;
      // Normalise the target month across a year boundary (0-indexed arithmetic).
      const idx = t.year * 12 + (t.month - 1) + delta;
      const targetYear = Math.floor(idx / 12);
      const targetMonth = (idx % 12) + 1;
      const c = serialToParts(cell, date1904);
      return c.year === targetYear && c.month === targetMonth;
    }
  }
}
