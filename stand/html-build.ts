// The HTML stand: a workbook made into a web page twice — by LibreOffice,
// which is the gold, and by Ream — both pages drawn by the same headless
// Chrome a sheet at a time, beside LibreOffice's own picture of each sheet and
// a map of where the two pages disagree.
//
//   the workbook → LibreOffice → gold.html → Chrome → a picture a sheet   (the gold)
//   the workbook → Ream        → ours.html → Chrome → a picture a sheet   (what we make)
//   the workbook → LibreOffice → a PDF page a sheet → poppler            (the sheet as it looks)
//
// The same browser on both sides, so what differs is the HTML. LibreOffice's
// files are kept per workbook (keyed by its bytes) and only ours is redone;
// both pages are drawn again each time, which takes a second.
//
// What a page is held to is the sheet as Excel shows it on screen: the whole
// grid at once, no pages. LibreOffice's HTML is that with faults of its own —
// no gridlines, a picture set above the grid rather than on it — and its
// picture of the sheet, the source column, is the one to believe when the two
// pages disagree.
//
// Usage:
//   npm run stand:html:build -- book.xlsx                copy into stand/files-html, build it
//   npm run stand:html:build -- stand/files-html/x.xlsx  rebuild one
//   npm run stand:html:build                             rebuild every workbook there
//   npm run stand:html                                   the page: http://localhost:4478

import { createHash } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { BLANK, diff, pageCount, pages, readPicture, soffice } from './build';
import { Chrome } from './chrome';
import type { PageScore } from './build';
import type { Box } from './chrome';
import type { Ppm } from '../scripts/corpus/lib';
import { Ream } from '@/core/converter/ream';

const here = dirname(fileURLToPath(import.meta.url));

/** Where the workbooks on this stand live (not in git: they are people's documents). */
export const FILES_DIR = resolve(here, 'files-html');
/** Where every build writes its pages, pictures and report (not in git). */
export const OUT_DIR = resolve(here, 'out-html');

/** What a file on this stand is. */
export const WORKBOOK = /\.(?:xlsx|xlsm|xls)$/iu;

/** The window both pages are laid out in, before a wide sheet widens it. */
const WINDOW = { width: 1280, height: 900 };
/** Past this width a page is not widened further: what is drawn is narrower still. */
const WIDEST_WINDOW = 8000;
/** The most of a sheet drawn, from its top left; a larger sheet is cut. */
const MOST_DRAWN = { width: 2400, height: 3600 };
/** How many sheets are drawn. Every one is counted. */
const MOST_SHEETS = 12;
/** CSS's own: a point is 4/3 of a pixel on the page and on the source alike. */
const DPI = 96;
/** Bumped when LibreOffice's side is made differently, so a kept one is redone. */
const GOLD_RECIPE = 1;

/** A sheet's extent on a page, in CSS pixels. */
export interface Size {
  readonly width: number;
  readonly height: number;
}

/** One sheet of the report: how big it came out on each side, and how far the two disagree. */
export interface SheetRow {
  readonly name: string;
  readonly gold: Size | null;
  readonly ours: Size | null;
  /** Whether each side's picture was drawn (see {@link MOST_SHEETS}). */
  readonly drawn: { readonly gold: boolean; readonly ours: boolean };
  readonly score: PageScore | null;
}

/** What one build leaves in `report.json`. */
export interface HtmlReport {
  readonly name: string;
  readonly source: string;
  readonly hash: string;
  readonly builtAt: number;
  /**
   * Whether the sheets were drawn one by one. When our page does not say
   * where a sheet begins, each side is drawn whole instead, as one row.
   */
  readonly perSheet: boolean;
  readonly sheets: ReadonlyArray<SheetRow>;
  /** LibreOffice's PDF of the workbook, a page a sheet: counted, and drawn. */
  readonly sourcePages: { readonly count: number; readonly drawn: number };
  /** The drawn rows' scores, in order — what the page sorts by. */
  readonly scores: ReadonlyArray<PageScore>;
  readonly losses: ReadonlyArray<{
    readonly severity: string;
    readonly feature: string;
    readonly detail: string;
  }>;
  readonly errors: ReadonlyArray<string>;
  readonly ms: { readonly ours: number; readonly render: number; readonly total: number };
}

interface GoldStamp {
  readonly hash: string;
  readonly recipe: number;
  readonly sourcePages: number;
  readonly drawnSource: number;
  readonly errors: ReadonlyArray<string>;
}

/** The name a file goes by on the stand — its own, minus the extension. */
export function stemOf(file: string): string {
  return basename(file).replace(WORKBOOK, '');
}

// ---- LibreOffice's side, kept ----

/**
 * LibreOffice's web page of the workbook and its picture of each sheet —
 * redone only when the file changes.
 */
function goldSide(book: string, dir: string, hash: string): GoldStamp {
  const stampFile = resolve(dir, 'gold.json');
  if (existsSync(stampFile)) {
    const kept = JSON.parse(readFileSync(stampFile, 'utf8')) as GoldStamp;
    if (kept.hash === hash && kept.recipe === GOLD_RECIPE) return kept;
  }
  const errors: Array<string> = [];
  const work = resolve(dir, '.work');
  rmSync(work, { recursive: true, force: true });
  mkdirSync(work, { recursive: true });
  // A plain name, so LibreOffice's output is where it is looked for; the
  // extension kept, since it is what picks the reader.
  const input = resolve(work, `gold${/\.[^.]+$/u.exec(book)?.[0] ?? '.xlsx'}`);
  copyFileSync(book, input);

  // The page, with its pictures beside it under names of LibreOffice's own.
  const page = resolve(dir, 'gold');
  rmSync(page, { recursive: true, force: true });
  try {
    const made = resolve(work, 'html');
    soffice(['--convert-to', 'html', '--outdir', made, input]);
    if (!existsSync(resolve(made, 'gold.html'))) throw new Error('LibreOffice wrote no page');
    renameSync(made, page);
  } catch (e) {
    errors.push(`gold: ${(e as Error).message.split('\n')[0] ?? ''}`);
  }

  // Each sheet whole on a page of its own size (Calc's "whole sheet export"),
  // drawn at the page's own scale.
  let sourcePages = 0;
  let drawnSource = 0;
  try {
    const made = resolve(work, 'pdf');
    soffice([
      '--convert-to',
      'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}',
      '--outdir',
      made,
      input,
    ]);
    const pdf = resolve(made, 'gold.pdf');
    if (!existsSync(pdf)) throw new Error('LibreOffice wrote no PDF');
    sourcePages = pageCount(pdf);
    drawnSource = pages(pdf, dir, 'source', {
      dpi: DPI,
      maxPages: MOST_SHEETS,
      maxPx: MOST_DRAWN,
    }).length;
    renameSync(pdf, resolve(dir, 'source.pdf'));
  } catch (e) {
    errors.push(`source: ${(e as Error).message.split('\n')[0] ?? ''}`);
  }
  rmSync(work, { recursive: true, force: true });

  const stamp: GoldStamp = { hash, recipe: GOLD_RECIPE, sourcePages, drawnSource, errors };
  writeFileSync(stampFile, `${JSON.stringify(stamp, null, 1)}\n`);
  return stamp;
}

// ---- the pages, drawn ----

interface Found {
  readonly name: string;
  readonly box: Box | null;
}

interface Layout {
  readonly sheets: ReadonlyArray<Found>;
  readonly whole: Box | null;
}

/**
 * Where each sheet is on a page, asked of the page itself. LibreOffice heads
 * every sheet with an anchor named `tableN`, and the sheet's pictures and grid
 * follow it up to a rule; ours names a sheet on the element that holds it, its
 * grid and drawings on a surface inside. A rectangle takes in the tables and
 * pictures inside, which run past a box narrower than they are; an empty sheet
 * has none.
 */
const FIND_SHEETS = String.raw`(() => {
  const box = (els, origin) => {
    let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
    const take = (e) => {
      const q = e.getBoundingClientRect();
      if (q.width === 0 || q.height === 0) return;
      l = Math.min(l, q.left); t = Math.min(t, q.top);
      r = Math.max(r, q.right); b = Math.max(b, q.bottom);
    };
    for (const el of els) {
      take(el);
      for (const e of el.querySelectorAll('table,img,svg,figure,canvas,[style*="position"]')) take(e);
    }
    if (l === Infinity) return null;
    if (origin) {
      const o = origin.getBoundingClientRect();
      l = Math.min(l, o.left); t = Math.min(t, o.top);
    }
    const x = Math.max(0, Math.floor(l + scrollX)), y = Math.max(0, Math.floor(t + scrollY));
    return { x, y, width: Math.ceil(r + scrollX) - x, height: Math.ceil(b + scrollY) - y };
  };
  const sheets = [];
  const anchors = [...document.querySelectorAll('a[name^="table"]')];
  if (anchors.length > 0) {
    const all = [];
    for (const a of anchors) {
      const em = a.querySelector('em');
      const els = [];
      for (let el = a.nextElementSibling; el && el.tagName !== 'HR' && !(el.tagName === 'A' && el.hasAttribute('name')); el = el.nextElementSibling) els.push(el);
      sheets.push({ name: ((em || a).textContent || '').trim(), box: box(els) });
      all.push(a, ...els);
    }
    return { sheets, whole: box(all) };
  }
  for (const el of document.querySelectorAll('[data-sheet]')) {
    // What the sheet holds, not the section: its heading is our own, and the
    // section is as wide as the page however narrow the grid in it. From the
    // surface's corner, which is the first cell's, so a sheet of drawings alone
    // keeps the room above and beside them.
    const surface = el.querySelector('.surface');
    sheets.push({
      name: el.getAttribute('data-sheet') || '',
      box: surface ? box([...surface.children], surface) : box([el]),
    });
  }
  const root = document.querySelector('article') || document.body;
  return { sheets, whole: box([...root.children]) };
})()`;

/** How wide the page has laid itself out — wider than the window when a sheet is. */
const PAGE_WIDTH = `Math.max(document.documentElement.scrollWidth, document.body ? document.body.scrollWidth : 0)`;

/**
 * Open a page and find its sheets. The window is first widened to the page:
 * Chrome draws past the window by growing it, and a page that centres itself
 * would move under the rectangles measured before.
 */
async function lay(chrome: Chrome, page: string): Promise<Layout> {
  await chrome.size(WINDOW.width, WINDOW.height);
  await chrome.open(page);
  let width = WINDOW.width;
  for (let i = 0; i < 3; i++) {
    const wants = Math.min(WIDEST_WINDOW, Number(await chrome.evaluate(PAGE_WIDTH)));
    if (wants <= width) break;
    width = wants;
    await chrome.size(width, WINDOW.height);
  }
  return (await chrome.evaluate(FIND_SHEETS)) as Layout;
}

/** A sheet's rectangle as much of it as is drawn. */
function drawnPart(box: Box): Box {
  return {
    x: box.x,
    y: box.y,
    width: Math.max(1, Math.min(box.width, MOST_DRAWN.width)),
    height: Math.max(1, Math.min(box.height, MOST_DRAWN.height)),
  };
}

/** A rectangle of a page, drawn. */
interface Shot {
  readonly box: Box;
  readonly png: Uint8Array;
}

/**
 * A page's sheets as the rows to compare: those it names, or the page whole as
 * one — which is also how LibreOffice writes a workbook of a single sheet,
 * with no anchor to name it by.
 */
function sheetsOf(layout: Layout, whole: boolean): Array<Found> {
  if (!whole && layout.sheets.length > 0) return [...layout.sheets];
  return [{ name: '', box: layout.whole }];
}

/** Each sheet of a page drawn — the first {@link MOST_SHEETS} of them. */
async function shoot(chrome: Chrome, sheets: ReadonlyArray<Found>): Promise<Array<Shot | null>> {
  const out: Array<Shot | null> = [];
  for (const [i, s] of sheets.entries()) {
    out.push(
      i < MOST_SHEETS && s.box ? { box: s.box, png: await chrome.shoot(drawnPart(s.box)) } : null,
    );
  }
  return out;
}

/**
 * The rows of the report: the two pages' sheets paired by name where both
 * name them alike, the rest in order — a sheet nameless on one side is the
 * next one unpaired on the other — in the gold's order, then any only ours has.
 */
function pair(
  gold: ReadonlyArray<Found>,
  goldShots: ReadonlyArray<Shot | null>,
  ours: ReadonlyArray<Found>,
  oursShots: ReadonlyArray<Shot | null>,
): Array<Row> {
  const taken = new Set<number>();
  const match = gold.map((g) => {
    const i = ours.findIndex((o, k) => !taken.has(k) && o.name !== '' && o.name === g.name);
    if (i >= 0) taken.add(i);
    return i;
  });
  let next = 0;
  for (const [gi, oi] of match.entries()) {
    if (oi >= 0) continue;
    while (next < ours.length && taken.has(next)) next++;
    if (next >= ours.length) break;
    match[gi] = next;
    taken.add(next);
  }
  const rows: Array<Row> = gold.map((g, gi) => {
    const oi = match[gi]!;
    const o = oi >= 0 ? ours[oi] : undefined;
    return {
      name: g.name || o?.name || 'the whole page',
      gold: goldShots[gi] ?? null,
      ours: oi >= 0 ? (oursShots[oi] ?? null) : null,
      goldBox: g.box,
      oursBox: o?.box ?? null,
    };
  });
  for (const [oi, o] of ours.entries()) {
    if (taken.has(oi)) continue;
    rows.push({
      name: o.name || 'the whole page',
      gold: null,
      ours: oursShots[oi] ?? null,
      goldBox: null,
      oursBox: o.box,
    });
  }
  return rows;
}

interface Row {
  readonly name: string;
  readonly gold: Shot | null;
  readonly ours: Shot | null;
  /** Whether the sheet is there at all on each side, drawn or not. */
  readonly goldBox: Box | null;
  readonly oursBox: Box | null;
}

// ---- one file ----

/** Build one workbook on the stand: LibreOffice's side (kept), ours (always), the diffs and the report. */
export async function build(book: string, chrome: Chrome): Promise<HtmlReport> {
  const started = Date.now();
  const name = stemOf(book);
  const dir = resolve(OUT_DIR, name);
  mkdirSync(dir, { recursive: true });
  const bytes = new Uint8Array(readFileSync(book));
  const hash = createHash('sha1').update(bytes).digest('hex').slice(0, 12);
  const stamp = goldSide(book, dir, hash);
  const errors = [...stamp.errors];

  let losses: HtmlReport['losses'] = [];
  let oursMs = 0;
  const oursPage = resolve(dir, 'ours.html');
  rmSync(oursPage, { force: true });
  try {
    const t0 = Date.now();
    const made = await Ream.parse(bytes).convertWithReport('html');
    writeFileSync(oursPage, made.bytes);
    losses = made.losses.map((l) => ({
      severity: l.severity,
      feature: l.feature,
      detail: l.detail,
    }));
    oursMs = Date.now() - t0;
  } catch (e) {
    errors.push(`ours: ${(e as Error).message.split('\n')[0] ?? ''}`);
  }

  // Ours first: whether OUR page says where its sheets are decides whether
  // the two are compared a sheet at a time or each whole.
  const t1 = Date.now();
  let perSheet = false;
  let ours: Array<Found> = [];
  let oursShots: Array<Shot | null> = [];
  try {
    if (existsSync(oursPage)) {
      const layout = await lay(chrome, oursPage);
      perSheet = layout.sheets.length > 0;
      ours = sheetsOf(layout, !perSheet);
      oursShots = await shoot(chrome, ours);
    }
  } catch (e) {
    errors.push(`our page: ${(e as Error).message}`);
  }
  const goldPage = resolve(dir, 'gold', 'gold.html');
  let gold: Array<Found> = [];
  let goldShots: Array<Shot | null> = [];
  try {
    if (existsSync(goldPage)) {
      const layout = await lay(chrome, goldPage);
      gold = sheetsOf(layout, !perSheet);
      goldShots = await shoot(chrome, gold);
    }
  } catch (e) {
    errors.push(`gold page: ${(e as Error).message}`);
  }
  const renderMs = Date.now() - t1;
  const rows = pair(gold, goldShots, ours, oursShots);

  for (const f of readdirSync(dir)) {
    if (/^(?:gold|ours|diff)-\d+\.png$/u.test(f)) unlinkSync(resolve(dir, f));
  }
  const scores: Array<PageScore> = [];
  const sheets: Array<SheetRow> = [];
  for (const [i, r] of rows.entries()) {
    const at = String(i + 1);
    const read = async (side: 'gold' | 'ours', shot: Shot | null): Promise<Ppm | null> => {
      if (!shot) return null;
      const file = resolve(dir, `${side}-${at}.png`);
      writeFileSync(file, shot.png);
      return readPicture(file);
    };
    const g = await read('gold', r.gold);
    const o = await read('ours', r.ours);
    let score: PageScore | null = null;
    if (g || o) {
      const made = diff(o ?? BLANK, g ?? BLANK, i + 1);
      writeFileSync(resolve(dir, `diff-${at}.png`), made.png);
      score = made.score;
      scores.push(score);
    }
    const size = (b: Box | null): Size | null => (b ? { width: b.width, height: b.height } : null);
    sheets.push({
      name: r.name,
      gold: size(r.goldBox),
      ours: size(r.oursBox),
      drawn: { gold: g !== null, ours: o !== null },
      score,
    });
  }

  const report: HtmlReport = {
    name,
    source: book,
    hash,
    builtAt: Date.now(),
    perSheet,
    sheets,
    sourcePages: { count: stamp.sourcePages, drawn: stamp.drawnSource },
    scores,
    losses,
    errors,
    ms: { ours: oursMs, render: renderMs, total: Date.now() - started },
  };
  writeFileSync(resolve(dir, 'report.json'), `${JSON.stringify(report, null, 1)}\n`);
  return report;
}

/** A path given on the command line, brought onto the stand if it is not there yet. */
function onStand(file: string): string {
  const full = resolve(file);
  if (!existsSync(full)) throw new Error(`not found: ${file}`);
  if (dirname(full) === FILES_DIR) return full;
  mkdirSync(FILES_DIR, { recursive: true });
  const kept = resolve(FILES_DIR, basename(full));
  copyFileSync(full, kept);
  return kept;
}

async function main(): Promise<void> {
  const given = process.argv.slice(2).filter((a) => !a.startsWith('--'));
  mkdirSync(FILES_DIR, { recursive: true });
  const files =
    given.length > 0
      ? given.map(onStand)
      : readdirSync(FILES_DIR)
          .filter((f) => WORKBOOK.test(f))
          .sort()
          .map((f) => resolve(FILES_DIR, f));
  // `--missing` builds only the workbooks with no report yet.
  const missing = process.argv.includes('--missing');
  const chrome = await Chrome.launch(WINDOW);
  try {
    for (const file of files) {
      if (missing && existsSync(resolve(OUT_DIR, stemOf(file), 'report.json'))) continue;
      let r: HtmlReport;
      try {
        r = await build(file, chrome);
      } catch (e) {
        process.stdout.write(`${stemOf(file).padEnd(44)} !! ${(e as Error).message}\n`);
        continue;
      }
      const worst = r.scores.reduce((m, s) => Math.max(m, s.missing, s.extra), 0);
      const errorsNote = r.errors.length > 0 ? `  !! ${r.errors.join(' | ')}` : '';
      process.stdout.write(
        `${r.name.padEnd(44)} worst ${worst.toFixed(3)}  sheets ${String(r.sheets.length)}${r.perSheet ? '' : ' (whole)'}  ${String(r.ms.total)} ms${errorsNote}\n`,
      );
    }
  } finally {
    chrome.close();
  }
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
