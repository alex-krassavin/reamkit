// §17.6.12 — the page numbers a PDF prints in its running head or foot, read
// back as the numbering they are.
//
// A page number is not text: it is the number of the page it stands on, in the
// numerals its section prints them in. A thesis numbers its front matter i, ii,
// iii and starts its body again at 1; bug793632.pdf is exactly that, four pages
// of it. Read as one decimal count, its footer printed 1, 2, 3, 4 — or, taken
// as the text of its first page, "i" on every one.

import type { NumberingFormat } from '@/core/document-model';

/** A page's number as its band prints it. */
export interface PageNumber {
  /** The numeral as the page shows it: "iv", "47". */
  readonly text: string;
  readonly value: number;
  readonly format: NumberingFormat;
}

/** A stretch of pages counted in one sequence (§17.6.12). */
export interface NumberingRun {
  /** The first page of the stretch. */
  readonly from: number;
  readonly format: NumberingFormat;
  /** The number the stretch's first page carries. */
  readonly start: number;
}

/** What the pages' bands say about their numbering. */
export interface PageNumbering {
  /** Each page's number, where its band shows one. */
  readonly numbers: ReadonlyArray<PageNumber | undefined>;
  /** The sequences the pages are counted in, first page first. */
  readonly runs: ReadonlyArray<NumberingRun>;
}

/** A numeral standing as a word of its own: digits, or roman in one case. */
const NUMERAL = /(?<=^|\s)(\d{1,4}|[ivxlcdm]{1,9}|[IVXLCDM]{1,9})(?=\s|$)/gu;

/**
 * More sequences than this and the "numbers" are something else that changes
 * from page to page — a date, a figure count — and the pages are left counted
 * the ordinary way.
 */
const MOST_RUNS = 4;

/**
 * …unless the sequences run long: a book leaves its blank pages out and skips
 * their numbers, and freeculture.pdf's count jumps two at five of its chapters
 * — seven sequences over three hundred and thirty-five numbered pages, and
 * read as none its footer printed no number at all. A date or a figure count
 * starts a sequence of its own on nearly every page.
 */
const LEAST_RUN_PAGES = 10;

/**
 * The numerals a band's text holds, in order.
 *
 * @param text The band's text on one page.
 * @returns Its numerals; a roman one only where it is a numeral as written.
 */
export function numeralsIn(text: string): Array<PageNumber> {
  const out: Array<PageNumber> = [];
  for (const m of text.matchAll(NUMERAL)) {
    const token = m[1]!;
    if (/^\d+$/u.test(token)) {
      out.push({ text: token, value: Number(token), format: 'decimal' });
      continue;
    }
    const value = romanValue(token);
    if (value === undefined) continue;
    out.push({
      text: token,
      value,
      format: token === token.toLowerCase() ? 'lowerRoman' : 'upperRoman',
    });
  }
  return out;
}

/**
 * A roman numeral's value, where the letters ARE one as a numeral is written
 * — "iv", "xiv", "MCMXC" — and not merely letters a numeral uses: "ic", "vx"
 * and "mix" are words or nothing.
 */
function romanValue(token: string): number | undefined {
  const upper = token.toUpperCase();
  const worth: Readonly<Record<string, number>> = {
    I: 1,
    V: 5,
    X: 10,
    L: 50,
    C: 100,
    D: 500,
    M: 1000,
  };
  let value = 0;
  for (let i = 0; i < upper.length; i++) {
    const here = worth[upper[i]!]!;
    const next = worth[upper[i + 1] ?? ''] ?? 0;
    value += here < next ? -here : here;
  }
  return value > 0 && value < 4000 && toRoman(value) === upper ? value : undefined;
}

/** The canonical roman numeral for a value, which is how a valid one reads. */
function toRoman(value: number): string {
  const steps: ReadonlyArray<readonly [number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let out = '';
  let left = value;
  for (const [n, s] of steps) {
    while (left >= n) {
      out += s;
      left -= n;
    }
  }
  return out;
}

/**
 * The page numbers a document's running band prints, and the sequences they
 * run in.
 *
 * The number is the numeral that CHANGES from page to page: "Page 3 of 10" is
 * counted by its first, "Chapter I — 47" by its last. A sequence goes on while
 * each page's number is one more than the last in the same numerals, and
 * starts again where it does not — i, ii, iii and then 1 are two sequences.
 * A page whose band shows no number is counted on.
 *
 * @param bands Each page's band text, or undefined where the page has none.
 * @returns The numbering, or undefined where the bands show no page number
 *          worth reading as one.
 */
export function pageNumberingOf(
  bands: ReadonlyArray<string | undefined>,
): PageNumbering | undefined {
  const found = bands.map((text) => (text === undefined ? [] : numeralsIn(text)));
  // Which numeral is the page's: the first position at which two pages that
  // both show one disagree.
  const counted = found.filter((n) => n.length > 0);
  if (counted.length < 2) return undefined;
  const width = Math.min(...counted.map((n) => n.length));
  let at = -1;
  for (let k = 0; k < width && at < 0; k++) {
    if (new Set(counted.map((n) => n[k]!.text)).size > 1) at = k;
  }
  if (at < 0) return undefined;
  const numbers = found.map((n) => n[at]);
  const runs: Array<NumberingRun> = [];
  numbers.forEach((number, page) => {
    if (number === undefined) return;
    const run = runs[runs.length - 1];
    if (run !== undefined) {
      const expected = run.start + (page - run.from);
      if (number.format === run.format && number.value === expected) return;
      runs.push({ from: page, format: number.format, start: number.value });
      return;
    }
    // The pages before the first number are counted with it where they can
    // be — a title page is page 1 of a body whose second page says 2 — and
    // are a count of their own where they cannot.
    const start = number.value - page;
    if (start >= 1) runs.push({ from: 0, format: number.format, start });
    else
      runs.push(
        { from: 0, format: 'decimal', start: 1 },
        { from: page, format: number.format, start: number.value },
      );
  });
  if (runs.length === 0 || runs.length > Math.max(MOST_RUNS, counted.length / LEAST_RUN_PAGES)) {
    return undefined;
  }
  return { numbers, runs };
}

/**
 * The sequence a page is counted in.
 *
 * @param runs The document's sequences, first page first.
 * @param page The page.
 * @returns The run that page belongs to.
 */
export function runOf(runs: ReadonlyArray<NumberingRun>, page: number): NumberingRun | undefined {
  let found: NumberingRun | undefined;
  for (const run of runs) if (run.from <= page) found = run;
  return found;
}
