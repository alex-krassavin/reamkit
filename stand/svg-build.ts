// The SVG stand: a workbook drawn as images of its sheets — each sheet whole on
// a page of its own — by LibreOffice, which is the gold, and by Ream, side by
// side a sheet at a time with a map of where the two disagree.
//
//   the workbook → LibreOffice → a PDF page a sheet → poppler   (the gold)
//   the workbook → Ream        → ours.svg            → Chrome    (what we make)
//
// Both are drawn at twice their size (144 dpi), a point being two pixels on
// each side. LibreOffice's PDF is the one the HTML stand draws its source
// column from ("whole sheet export"): where that stand has already made it for
// the same bytes it is taken from there, and otherwise made here and kept.
//
// Usage:
//   npm run stand:svg:build -- book.xlsx              copy into stand/files-html, build it
//   npm run stand:svg:build                           rebuild every workbook there
//   npm run stand:svg                                 the page: http://localhost:4479

import { createHash } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { corpusFontOptions } from '../scripts/corpus/fonts';
import { BLANK, diff, pageCount, pages, readPicture, soffice } from './build';
import { Chrome } from './chrome';
import { OUT_DIR as HTML_OUT_DIR } from './html-build';
import type { PageScore } from './build';
import type { Box } from './chrome';
import type { HtmlReport, SheetRow } from './html-build';
import type { Ppm } from '../scripts/corpus/lib';
import { Ream } from '@/core/converter/ream';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

const here = dirname(fileURLToPath(import.meta.url));

/** The workbooks — the HTML stand's own, so the two compare the same files. */
export const FILES_DIR = resolve(here, 'files-html');
/** Where every build writes its pages, pictures and report (not in git). */
export const OUT_DIR = resolve(here, 'out-svg');

/** What a file on this stand is. */
export const WORKBOOK = /\.(?:xlsx|xlsm|xls)$/iu;

/** Twice the page's size: a point is two pixels on both sides. */
const SCALE = 2;
const DPI = 72 * SCALE;
/** The most of a sheet drawn, in pixels from its top left; a larger sheet is cut. */
const MOST_DRAWN = { width: 3200, height: 4800 };
/** How many sheets are drawn. Every one is counted. */
const MOST_SHEETS = 12;
/**
 * Bumped when the gold side is made differently, so a kept one is redone. 2:
 * LibreOffice's pages are drawn as `lo-N.png`, in its own order, and each build
 * stands them beside ours as `gold-k.png`.
 */
const GOLD_RECIPE = 2;

interface GoldStamp {
  readonly hash: string;
  readonly recipe: number;
  readonly pages: number;
  readonly drawn: number;
  readonly errors: ReadonlyArray<string>;
}

/**
 * The stamp an earlier build left beside its gold, or undefined where there is
 * none to read — or one that no longer parses. Read in one go rather than
 * checked first, so the answer is about the file read.
 */
function readStamp(file: string): GoldStamp | undefined {
  try {
    return JSON.parse(readFileSync(file, 'utf8')) as GoldStamp;
  } catch {
    return undefined;
  }
}

/** The name a file goes by on the stand — its own, minus the extension. */
export function stemOf(file: string): string {
  return basename(file).replace(WORKBOOK, '');
}

/** LibreOffice's whole-sheet PDF of the workbook, drawn a page a sheet — redone only when the file changes. */
function goldSide(book: string, dir: string, hash: string): GoldStamp {
  const stampFile = resolve(dir, 'gold.json');
  const kept = readStamp(stampFile);
  if (kept?.hash === hash && kept.recipe === GOLD_RECIPE) return kept;
  // The same file, drawn another way: its PDF is still good.
  const keptPdf = kept?.hash === hash && existsSync(resolve(dir, 'gold.pdf'));
  const errors: Array<string> = [];
  const work = resolve(dir, '.work');
  rmSync(work, { recursive: true, force: true });
  mkdirSync(work, { recursive: true });
  let pdf = resolve(dir, 'gold.pdf');
  // The HTML stand's source PDF is this very file, where it was made from the
  // same bytes.
  const html = resolve(HTML_OUT_DIR, stemOf(book));
  const htmlStamp = resolve(html, 'gold.json');
  const shared =
    existsSync(htmlStamp) &&
    (JSON.parse(readFileSync(htmlStamp, 'utf8')) as { hash?: string }).hash === hash &&
    existsSync(resolve(html, 'source.pdf'));
  try {
    if (keptPdf) {
      // Drawn again below from the PDF already made.
    } else if (shared) {
      copyFileSync(resolve(html, 'source.pdf'), pdf);
    } else {
      const input = resolve(work, `gold${/\.[^.]+$/u.exec(book)?.[0] ?? '.xlsx'}`);
      copyFileSync(book, input);
      soffice([
        '--convert-to',
        'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}',
        '--outdir',
        work,
        input,
      ]);
      const made = resolve(work, 'gold.pdf');
      if (!existsSync(made)) throw new Error('LibreOffice wrote no PDF');
      copyFileSync(made, pdf);
    }
  } catch (e) {
    errors.push(`gold: ${(e as Error).message.split('\n')[0] ?? ''}`);
    pdf = '';
  }
  let count = 0;
  let drawn = 0;
  if (pdf) {
    count = pageCount(pdf);
    drawn = pages(pdf, dir, 'lo', { dpi: DPI, maxPages: MOST_SHEETS, maxPx: MOST_DRAWN }).length;
  }
  rmSync(work, { recursive: true, force: true });
  const stamp: GoldStamp = { hash, recipe: GOLD_RECIPE, pages: count, drawn, errors };
  writeFileSync(stampFile, `${JSON.stringify(stamp, null, 1)}\n`);
  return stamp;
}

/** Luminance under which a pixel is ink — the diff's own threshold. */
const INK_LUMA = 200;

/** Where a picture's ink begins: its leftmost and its topmost inked pixel. */
function inkOrigin(p: Ppm): { x: number; y: number } {
  let x = p.width;
  let y = p.height;
  for (let row = 0; row < p.height; row++) {
    for (let col = 0; col < p.width; col++) {
      const i = (row * p.width + col) * 3;
      const luma = 0.299 * p.rgb[i]! + 0.587 * p.rgb[i + 1]! + 0.114 * p.rgb[i + 2]!;
      if (luma >= INK_LUMA) continue;
      if (col < x) x = col;
      if (row < y) y = row;
    }
  }
  return x === p.width ? { x: 0, y: 0 } : { x, y };
}

/** A picture moved by `(dx, dy)`, on white paper of its own size. */
function shifted(p: Ppm, dx: number, dy: number): Ppm {
  const rgb = new Uint8Array(p.width * p.height * 3).fill(255);
  for (let row = 0; row < p.height; row++) {
    const from = row - dy;
    if (from < 0 || from >= p.height) continue;
    for (let col = 0; col < p.width; col++) {
      const fc = col - dx;
      if (fc < 0 || fc >= p.width) continue;
      const i = (row * p.width + col) * 3;
      const j = (from * p.width + fc) * 3;
      rgb[i] = p.rgb[j]!;
      rgb[i + 1] = p.rgb[j + 1]!;
      rgb[i + 2] = p.rgb[j + 2]!;
    }
  }
  return { width: p.width, height: p.height, rgb };
}

/**
 * Which of LibreOffice's pages each of ours stands beside. LibreOffice exports
 * EVERY sheet, the hidden ones too, and we draw only those a window shows — so
 * our k-th page is the k-th visible sheet, at that sheet's place in the book.
 */
function goldPageOf(bytes: Uint8Array): (ours: number) => number {
  let visible: Array<number> | undefined;
  try {
    visible = readXlsxToSheetDoc(bytes).sheets.flatMap((s, i) => (s.hidden ? [] : [i]));
  } catch {
    visible = undefined;
  }
  return (ours) => visible?.[ours] ?? ours;
}

/** Where each page of our SVG is, in CSS pixels — its `<g data-page>` boxes. */
const FIND_PAGES = `[...document.querySelectorAll('g[data-page]')].map((g) => {
  const r = g.getBoundingClientRect();
  return { x: Math.floor(r.left + scrollX), y: Math.floor(r.top + scrollY), width: Math.ceil(r.width), height: Math.ceil(r.height) };
})`;

/** Build one workbook on the stand: LibreOffice's side (kept), ours (always), the diffs and the report. */
export async function build(book: string, chrome: Chrome): Promise<HtmlReport> {
  const started = Date.now();
  const name = stemOf(book);
  const dir = resolve(OUT_DIR, name);
  mkdirSync(dir, { recursive: true });
  const bytes = new Uint8Array(readFileSync(book));
  const hash = createHash('sha1').update(bytes).digest('hex').slice(0, 12);
  const gold = goldSide(book, dir, hash);
  const errors = [...gold.errors];

  let losses: HtmlReport['losses'] = [];
  let oursMs = 0;
  const oursFile = resolve(dir, 'ours.svg');
  rmSync(oursFile, { force: true });
  try {
    const t0 = Date.now();
    const made = await Ream.parse(bytes).convertWithReport('svg', corpusFontOptions());
    writeFileSync(oursFile, made.bytes);
    losses = made.losses.map((l) => ({
      severity: l.severity,
      feature: l.feature,
      detail: l.detail,
    }));
    oursMs = Date.now() - t0;
  } catch (e) {
    errors.push(`ours: ${(e as Error).message.split('\n')[0] ?? ''}`);
  }

  // Our pages, each cut out of the SVG as Chrome draws it.
  const t1 = Date.now();
  for (const f of readdirSync(dir)) {
    if (/^(?:ours|diff)-\d+\.png$/u.test(f)) unlinkSync(resolve(dir, f));
  }
  let boxes: Array<Box> = [];
  try {
    if (existsSync(oursFile)) {
      await chrome.open(oursFile);
      boxes = (await chrome.evaluate(FIND_PAGES)) as Array<Box>;
      for (const [i, b] of boxes.slice(0, MOST_SHEETS).entries()) {
        const png = await chrome.shoot({
          x: b.x,
          y: b.y,
          width: Math.max(1, Math.min(b.width, MOST_DRAWN.width / SCALE)),
          height: Math.max(1, Math.min(b.height, MOST_DRAWN.height / SCALE)),
        });
        writeFileSync(resolve(dir, `ours-${String(i + 1)}.png`), png);
      }
    }
  } catch (e) {
    errors.push(`our page: ${(e as Error).message}`);
  }
  const renderMs = Date.now() - t1;

  // Each of our pages beside LibreOffice's page of the same sheet, its
  // picture copied to stand at our number; the hidden sheets' come last.
  const pageOf = goldPageOf(bytes);
  const order: Array<number> = boxes.map((_, i) => pageOf(i));
  for (let p = 0; p < gold.pages; p++) if (!order.includes(p)) order.push(p);
  for (const f of readdirSync(dir)) {
    if (/^gold-\d+\.png$/u.test(f)) unlinkSync(resolve(dir, f));
  }
  const scores: Array<PageScore> = [];
  const sheets: Array<SheetRow> = [];
  for (const [i, page] of order.entries()) {
    const at = String(i + 1);
    const loFile = resolve(dir, `lo-${String(page + 1)}.png`);
    let g: Ppm | null = null;
    if (page < gold.drawn && existsSync(loFile)) {
      copyFileSync(loFile, resolve(dir, `gold-${at}.png`));
      g = await readPicture(loFile);
    }
    const oursPng = resolve(dir, `ours-${at}.png`);
    const o = i < boxes.length && existsSync(oursPng) ? await readPicture(oursPng) : null;
    let score: PageScore | null = null;
    if (g || o) {
      // Compared from where each one's ink begins: LibreOffice's page starts
      // at the first used cell, ours at A1, as the window does.
      let aligned = o ?? BLANK;
      if (o && g) {
        const go = inkOrigin(g);
        const oo = inkOrigin(o);
        aligned = shifted(o, go.x - oo.x, go.y - oo.y);
      }
      const made = diff(aligned, g ?? BLANK, i + 1);
      writeFileSync(resolve(dir, `diff-${at}.png`), made.png);
      score = made.score;
      scores.push(score);
    }
    const box = boxes[i];
    sheets.push({
      name: i < boxes.length ? `Sheet ${at}` : `Hidden · LibreOffice's page ${String(page + 1)}`,
      gold: g ? { width: g.width / SCALE, height: g.height / SCALE } : null,
      ours: box ? { width: box.width, height: box.height } : null,
      drawn: { gold: g !== null, ours: o !== null },
      score,
    });
  }

  const report: HtmlReport = {
    name,
    source: book,
    hash,
    builtAt: Date.now(),
    perSheet: true,
    sheets,
    sourcePages: { count: 0, drawn: 0 },
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
  const missing = process.argv.includes('--missing');
  const chrome = await Chrome.launch({ width: 1280, height: 900, scale: SCALE });
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
        `${r.name.padEnd(44)} worst ${worst.toFixed(3)}  sheets ${String(r.sheets.length)}  ${String(r.ms.total)} ms${errorsNote}\n`,
      );
    }
  } finally {
    chrome.close();
  }
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
