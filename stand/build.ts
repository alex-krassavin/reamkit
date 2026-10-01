// The stand: a PDF turned into a .docx twice — by LibreOffice, which is the
// gold, and by Ream — and both packages drawn by the same LibreOffice, beside
// the source page and a map of where the two disagree.
//
//   the file → LibreOffice → gold.docx → LibreOffice → pages   (the gold)
//   the file → Ream        → ours.docx → LibreOffice → pages   (what we make)
//
// Same renderer on both sides, so what differs is the .docx. The gold side and
// the source pages are kept per file (keyed by its bytes) and only ours is
// redone, which is what makes a change to the reader a few seconds to see.
//
// Usage:
//   npm run stand:build -- Invoice.pdf           copy into stand/files, build it
//   npm run stand:build -- stand/files/x.pdf     rebuild one
//   npm run stand:build                          rebuild every file in stand/files
//   npm run stand                                the page: http://localhost:4477

import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import {
  closeSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  openSync,
  readFileSync,
  readdirSync,
  renameSync,
  rmSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { brotliDecompressSync } from 'node:zlib';
import { tmpdir } from 'node:os';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { parsePpm } from '../scripts/corpus/lib';
import type { Ppm } from '../scripts/corpus/lib';
import { Ream } from '@/core/converter/ream';
import { encodePng } from '@/core/png-encode';

const here = dirname(fileURLToPath(import.meta.url));

// sharp is the docs site's dependency, not the library's: the stand reads its
// kept pages back with it, and the library must not grow a native image
// dependency for a dev-only bench.
interface Sharp {
  removeAlpha: () => Sharp;
  raw: () => Sharp;
  toBuffer: (o: { resolveWithObject: true }) => Promise<{
    data: Buffer;
    info: { width: number; height: number };
  }>;
}
const sharp = createRequire(resolve(here, '../docs/package.json'))('sharp') as (
  input: string,
) => Sharp;
/** Where the PDFs on the stand live (not in git: they are people's documents). */
export const FILES_DIR = resolve(here, 'files');
/** Where every build writes its packages, pages and report (not in git). */
export const OUT_DIR = resolve(here, 'out');

/** One resolution for every picture, so a pixel means the same on all four. */
const DPI = 110;
const SOFFICE_TIMEOUT_MS = 180_000;
/** Bumped when the gold side is made differently, so a kept one is redone. */
const GOLD_RECIPE = 3;
/**
 * Bumped when the SOURCE page is drawn differently — redone on its own,
 * without asking LibreOffice again. 2: drawn to the crop box, which is what a
 * viewer shows (§14.11.2); drawn to the media box, a file cropped to a corner
 * of its sheet showed the whole sheet beside our page of the corner.
 */
const SOURCE_RECIPE = 2;
/**
 * How many pages of each side are DRAWN. Every page is counted — a page too
 * many or too few is the first thing to know — but a corpus file can run to a
 * hundred pages, and the look of the first few is what a pass over the corpus
 * needs.
 */
const MAX_PAGES = 6;

/** Per channel, how far two pixels may differ and still be the same ink. */
const SAME_TOLERANCE = 40;
/** How far, in pixels, a pixel may look for its match (anti-aliasing, hinting). */
const SLACK_PX = 1;
/** Luminance under which a pixel is ink rather than paper. */
const INK_LUMA = 200;

/** What a page of the report says. */
export interface PageScore {
  readonly page: number;
  /** Share of the gold's ink that has no match in ours. */
  readonly missing: number;
  /** Share of our ink that has no match in the gold. */
  readonly extra: number;
}

/** What one build leaves in `report.json`. */
export interface Report {
  readonly name: string;
  readonly source: string;
  readonly hash: string;
  readonly builtAt: number;
  readonly pages: { readonly source: number; readonly gold: number; readonly ours: number };
  /** How many of those pages are drawn (see {@link MAX_PAGES}). */
  readonly drawn: { readonly source: number; readonly gold: number; readonly ours: number };
  readonly scores: ReadonlyArray<PageScore>;
  readonly errors: ReadonlyArray<string>;
  readonly ms: { readonly ours: number; readonly render: number; readonly total: number };
}

interface GoldStamp {
  readonly hash: string;
  readonly recipe: number;
  /** How the source page was drawn (see {@link SOURCE_RECIPE}); absent before there was a choice. */
  readonly sourceRecipe?: number;
  /** Pages counted, and pages drawn. */
  readonly source: number;
  readonly gold: number;
  readonly drawnSource: number;
  readonly drawnGold: number;
  readonly errors: ReadonlyArray<string>;
}

/** The name a file goes by on the stand — its own, minus the extension. */
export function stemOf(file: string): string {
  return basename(file).replace(/\.pdf$/iu, '');
}

// ---- LibreOffice, one at a time ----

const PROFILE_NAME = 'ream-stand-lo';
const PROFILE = pathToFileURL(resolve(tmpdir(), PROFILE_NAME)).href;
const LOCK = resolve(tmpdir(), 'ream-stand-lo.lock');

/**
 * Run `soffice` under a lock.
 *
 * One profile is one running instance: a second `soffice` started while the
 * first converts hands its job over and exits with nothing written. The stand
 * keeps ONE profile (a fresh one costs seconds to set up on every call) and
 * takes turns on it instead.
 */
export function soffice(args: ReadonlyArray<string>): void {
  const held = lock();
  try {
    execFileSync(SOFFICE, [`-env:UserInstallation=${PROFILE}`, '--headless', ...args], {
      stdio: 'ignore',
      timeout: SOFFICE_TIMEOUT_MS,
    });
  } catch (e) {
    // A conversion past its time leaves LibreOffice running on the profile,
    // and every call after it hands its job to that instance and waits: one
    // slow file stalled the whole pass. Only this bench's instances go.
    try {
      execFileSync('pkill', ['-f', PROFILE_NAME], { stdio: 'ignore' });
    } catch {
      // Nothing was left running.
    }
    throw e;
  } finally {
    closeSync(held);
    rmSync(LOCK, { force: true });
  }
}

/** LibreOffice itself rather than the shell script in front of it, so a timeout stops IT. */
const SOFFICE = existsSync('/Applications/LibreOffice.app/Contents/MacOS/soffice')
  ? '/Applications/LibreOffice.app/Contents/MacOS/soffice'
  : 'soffice';

function lock(): number {
  for (;;) {
    try {
      const fd = openSync(LOCK, 'wx');
      writeFileSync(fd, String(process.pid));
      return fd;
    } catch {
      // Held — by a live build, or left behind by one that died.
      const owner = Number(readFileSync(LOCK, 'utf8'));
      if (!alive(owner)) {
        rmSync(LOCK, { force: true });
        continue;
      }
      sleep(250);
    }
  }
}

function alive(pid: number): boolean {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function sleep(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

/** A .docx drawn by LibreOffice: the pages as a PDF beside it. */
function renderDocx(docx: string, work: string): string {
  const pdf = resolve(work, `${basename(docx).replace(/\.docx$/iu, '')}.pdf`);
  // Once more when nothing came of it: an instance still closing down takes
  // the job and drops it, and the file is not at fault.
  for (let attempt = 0; attempt < 2 && !existsSync(pdf); attempt++) {
    soffice(['--convert-to', 'pdf', '--outdir', work, docx]);
  }
  if (!existsSync(pdf)) throw new Error(`LibreOffice would not open ${basename(docx)}`);
  return pdf;
}

// ---- pictures ----

/** How many pages a PDF has, by poppler's count. */
export function pageCount(pdf: string): number {
  try {
    const info = execFileSync('pdfinfo', [pdf], { encoding: 'utf8', timeout: 60_000 });
    return Number(/^Pages:\s+(\d+)/mu.exec(info)?.[1] ?? 0);
  } catch {
    return 0;
  }
}

/** How a PDF is drawn into pictures: the resolution, how many pages, and how much of each. */
export interface Drawing {
  readonly dpi?: number;
  readonly maxPages?: number;
  /** The most of a page drawn, in pixels from its top left; a page past it is cut. */
  readonly maxPx?: { readonly width: number; readonly height: number };
}

/** Every page of a PDF, drawn by poppler, as PPM files in page order. */
function rasterize(pdf: string, prefix: string, how: Drawing): Array<string> {
  const cut = how.maxPx
    ? ['-x', '0', '-y', '0', '-W', String(how.maxPx.width), '-H', String(how.maxPx.height)]
    : [];
  execFileSync(
    'pdftoppm',
    [
      '-cropbox',
      '-r',
      String(how.dpi ?? DPI),
      ...cut,
      '-f',
      '1',
      '-l',
      String(how.maxPages ?? MAX_PAGES),
      pdf,
      prefix,
    ],
    { stdio: 'ignore', timeout: 120_000 },
  );
  const dir = dirname(prefix);
  const stem = basename(prefix);
  const number = (f: string): number => Number(/-(\d+)\.ppm$/u.exec(f)?.[1] ?? 0);
  return readdirSync(dir)
    .filter((f) => f.startsWith(`${stem}-`) && f.endsWith('.ppm'))
    .sort((a, b) => number(a) - number(b))
    .map((f) => resolve(dir, f));
}

/** Draw a PDF into `<tag>-N.png` pages in `dir`, keeping the PPMs for the diff. */
export function pages(pdf: string, dir: string, tag: string, how: Drawing = {}): Array<Ppm> {
  const raw = resolve(dir, '.raw');
  rmSync(raw, { recursive: true, force: true });
  mkdirSync(raw, { recursive: true });
  for (const f of readdirSync(dir)) {
    if (new RegExp(`^${tag}-\\d+\\.png$`, 'u').test(f)) unlinkSync(resolve(dir, f));
  }
  const out: Array<Ppm> = [];
  rasterize(pdf, resolve(raw, tag), how).forEach((file, i) => {
    const ppm = parsePpm(new Uint8Array(readFileSync(file)));
    out.push(ppm);
    writeFileSync(
      resolve(dir, `${tag}-${String(i + 1)}.png`),
      encodePng(ppm.width, ppm.height, 'rgb', ppm.rgb),
    );
  });
  // A page as raw pixels is three bytes a dot: kept, a corpus on the stand
  // would be gigabytes. The PNG is what is kept, and read back when needed.
  rmSync(raw, { recursive: true, force: true });
  return out;
}

/** The kept pages of one side, read back from the PNGs `pages` wrote. */
export async function keptPages(dir: string, tag: string): Promise<Array<Ppm>> {
  const number = (f: string): number => Number(/-(\d+)\.png$/u.exec(f)?.[1] ?? 0);
  const files = readdirSync(dir)
    .filter((f) => new RegExp(`^${tag}-\\d+\\.png$`, 'u').test(f))
    .sort((a, b) => number(a) - number(b));
  const out: Array<Ppm> = [];
  for (const f of files) out.push(await readPicture(resolve(dir, f)));
  return out;
}

/** A PNG read back as raw pixels, its alpha dropped. */
export async function readPicture(file: string): Promise<Ppm> {
  const { data, info } = await sharp(file)
    .removeAlpha()
    .raw()
    .toBuffer({ resolveWithObject: true });
  return { width: info.width, height: info.height, rgb: new Uint8Array(data) };
}

const luma = (p: Ppm, i: number): number =>
  0.299 * p.rgb[i]! + 0.587 * p.rgb[i + 1]! + 0.114 * p.rgb[i + 2]!;

/** Whether `p` has, within the slack around (x, y), a pixel `other` would call the same. */
function matchNear(p: Ppm, x: number, y: number, other: Ppm, oi: number): boolean {
  for (let dy = -SLACK_PX; dy <= SLACK_PX; dy++) {
    const ny = y + dy;
    if (ny < 0 || ny >= p.height) continue;
    for (let dx = -SLACK_PX; dx <= SLACK_PX; dx++) {
      const nx = x + dx;
      if (nx < 0 || nx >= p.width) continue;
      const pi = (ny * p.width + nx) * 3;
      if (
        Math.abs(p.rgb[pi]! - other.rgb[oi]!) <= SAME_TOLERANCE &&
        Math.abs(p.rgb[pi + 1]! - other.rgb[oi + 1]!) <= SAME_TOLERANCE &&
        Math.abs(p.rgb[pi + 2]! - other.rgb[oi + 2]!) <= SAME_TOLERANCE
      ) {
        return true;
      }
    }
  }
  return false;
}

/** No page at all: what a side that has fewer pages is diffed as. */
export const BLANK: Ppm = { width: 0, height: 0, rgb: new Uint8Array(0) };

/** `p` on white paper of the given size — a missing page is a blank one. */
function onPaper(p: Ppm, width: number, height: number): Ppm {
  if (p.width === width && p.height === height) return p;
  const rgb = new Uint8Array(width * height * 3).fill(255);
  for (let y = 0; y < Math.min(p.height, height); y++) {
    const row = p.rgb.subarray(y * p.width * 3, (y * p.width + Math.min(p.width, width)) * 3);
    rgb.set(row, y * width * 3);
  }
  return { width, height, rgb };
}

/**
 * Where ours and the gold disagree, as a picture: agreement is the gold page
 * faded, RED is gold ink ours does not have there, BLUE is ours where the gold
 * has none. A line set a few points off shows as a red and a blue copy of it.
 */
export function diff(
  oursPage: Ppm,
  goldPage: Ppm,
  page: number,
): { png: Uint8Array; score: PageScore } {
  const width = Math.max(oursPage.width, goldPage.width);
  const height = Math.max(oursPage.height, goldPage.height);
  const ours = onPaper(oursPage, width, height);
  const gold = onPaper(goldPage, width, height);
  const rgb = new Uint8Array(width * height * 3).fill(255);
  let goldInk = 0;
  let oursInk = 0;
  let missing = 0;
  let extra = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 3;
      const gInk = luma(gold, i) < INK_LUMA;
      const oInk = luma(ours, i) < INK_LUMA;
      if (gInk) goldInk++;
      if (oInk) oursInk++;
      const goldFound = matchNear(ours, x, y, gold, i);
      const oursFound = matchNear(gold, x, y, ours, i);
      if (gInk && !goldFound) {
        missing++;
        rgb.set([222, 38, 38], i);
      } else if (oInk && !oursFound) {
        extra++;
        rgb.set([24, 112, 232], i);
      } else if (!goldFound || !oursFound) {
        // A tint that moved — a fill or a shade drawn lighter or darker.
        rgb.set(luma(gold, i) < luma(ours, i) ? [246, 170, 170] : [168, 200, 246], i);
      } else {
        // Agreement: the gold, faded, so the page is still readable.
        for (let c = 0; c < 3; c++) rgb[i + c] = 255 - Math.round((255 - gold.rgb[i + c]!) * 0.3);
      }
    }
  }
  return {
    png: encodePng(width, height, 'rgb', rgb),
    score: {
      page,
      missing: goldInk > 0 ? missing / goldInk : 0,
      extra: oursInk > 0 ? extra / oursInk : 0,
    },
  };
}

// ---- one file ----

/** LibreOffice's own PDF → .docx, and the source page: redone only when the file changes. */
function goldSide(pdf: string, dir: string, hash: string): GoldStamp {
  const stampFile = resolve(dir, 'gold.json');
  if (existsSync(stampFile)) {
    const kept = JSON.parse(readFileSync(stampFile, 'utf8')) as GoldStamp;
    if (kept.hash === hash && kept.recipe === GOLD_RECIPE) {
      if (kept.sourceRecipe === SOURCE_RECIPE) return kept;
      let drawnSource = 0;
      try {
        drawnSource = pages(pdf, dir, 'source').length;
      } catch {
        drawnSource = 0;
      }
      const redrawn: GoldStamp = { ...kept, drawnSource, sourceRecipe: SOURCE_RECIPE };
      writeFileSync(stampFile, `${JSON.stringify(redrawn, null, 1)}\n`);
      return redrawn;
    }
  }
  const errors: Array<string> = [];
  const source = pageCount(pdf);
  let drawnSource = 0;
  try {
    drawnSource = pages(pdf, dir, 'source').length;
  } catch (e) {
    // poppler reads what it reads: a Brotli stream is beyond it, and the
    // source side is then the one missing picture.
    errors.push(`source: ${(e as Error).message.split('\n')[0] ?? ''}`);
  }
  let goldPages = 0;
  let drawnGold = 0;
  const work = resolve(dir, '.work');
  rmSync(work, { recursive: true, force: true });
  mkdirSync(work, { recursive: true });
  try {
    // A plain name, so LibreOffice's output is where it is looked for.
    const input = resolve(work, 'gold.pdf');
    copyFileSync(pdf, input);
    // The picture is LibreOffice's reading of the file drawn straight from its
    // own model. Its .docx of that reading does not survive being opened again:
    // the white rectangle a page is painted with comes back IN FRONT of every
    // text box (Invoice-6VOBWUGP-0010 reopens as a blank page), so drawing the
    // package would hold ours to a page with nothing on it.
    const drawn = resolve(work, 'drawn');
    soffice(['--infilter=writer_pdf_import', '--convert-to', 'pdf', '--outdir', drawn, input]);
    const page = resolve(drawn, 'gold.pdf');
    if (!existsSync(page)) throw new Error('LibreOffice could not read the file');
    goldPages = pageCount(page);
    drawnGold = pages(page, dir, 'gold').length;
    // The package itself, kept to open beside ours.
    soffice(['--infilter=writer_pdf_import', '--convert-to', 'docx', '--outdir', work, input]);
    const docx = resolve(work, 'gold.docx');
    if (existsSync(docx)) renameSync(docx, resolve(dir, 'gold.docx'));
  } catch (e) {
    errors.push(`gold: ${(e as Error).message}`);
  }
  rmSync(work, { recursive: true, force: true });
  const stamp: GoldStamp = {
    hash,
    recipe: GOLD_RECIPE,
    sourceRecipe: SOURCE_RECIPE,
    source,
    gold: goldPages,
    drawnSource,
    drawnGold,
    errors,
  };
  writeFileSync(stampFile, `${JSON.stringify(stamp, null, 1)}\n`);
  return stamp;
}

/** Build one file on the stand: gold (kept), ours (always), the diff and the report. */
export async function build(pdf: string): Promise<Report> {
  const started = Date.now();
  const name = stemOf(pdf);
  const dir = resolve(OUT_DIR, name);
  mkdirSync(dir, { recursive: true });
  const bytes = new Uint8Array(readFileSync(pdf));
  const hash = createHash('sha1').update(bytes).digest('hex').slice(0, 12);
  const stamp = goldSide(pdf, dir, hash);
  const errors = [...stamp.errors];

  let oursPages: Array<Ppm> = [];
  let oursCount = 0;
  let oursMs = 0;
  let renderMs = 0;
  const work = resolve(dir, '.work');
  rmSync(work, { recursive: true, force: true });
  mkdirSync(work, { recursive: true });
  try {
    const t0 = Date.now();
    const doc = Ream.parse(bytes, { filters: { BrotliDecode: brotli } });
    writeFileSync(resolve(dir, 'ours.docx'), await doc.convert('docx'));
    oursMs = Date.now() - t0;
    const t1 = Date.now();
    copyFileSync(resolve(dir, 'ours.docx'), resolve(work, 'ours.docx'));
    const drawn = renderDocx(resolve(work, 'ours.docx'), work);
    oursCount = pageCount(drawn);
    oursPages = pages(drawn, dir, 'ours');
    renderMs = Date.now() - t1;
  } catch (e) {
    errors.push(`ours: ${(e as Error).message}`);
  }
  rmSync(work, { recursive: true, force: true });

  const goldPages = await keptPages(dir, 'gold');
  for (const f of readdirSync(dir)) if (/^diff-\d+\.png$/u.test(f)) unlinkSync(resolve(dir, f));
  const scores: Array<PageScore> = [];
  const count = Math.max(goldPages.length, oursPages.length);
  for (let i = 0; i < count; i++) {
    const made = diff(oursPages[i] ?? BLANK, goldPages[i] ?? BLANK, i + 1);
    writeFileSync(resolve(dir, `diff-${String(i + 1)}.png`), made.png);
    scores.push(made.score);
  }

  const report: Report = {
    name,
    source: pdf,
    hash,
    builtAt: Date.now(),
    pages: { source: stamp.source, gold: stamp.gold, ours: oursCount },
    drawn: { source: stamp.drawnSource, gold: goldPages.length, ours: oursPages.length },
    scores,
    errors,
    ms: { ours: oursMs, render: renderMs, total: Date.now() - started },
  };
  writeFileSync(resolve(dir, 'report.json'), `${JSON.stringify(report, null, 1)}\n`);
  return report;
}

/**
 * §7.4 `/BrotliDecode`, from the runtime: the library carries no Brotli and
 * takes one from its caller, and the stand is a caller with `node:zlib`.
 */
function brotli(input: Uint8Array): Uint8Array {
  return new Uint8Array(brotliDecompressSync(input));
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
          .filter((f) => /\.pdf$/iu.test(f))
          .sort()
          .map((f) => resolve(FILES_DIR, f));
  // `--missing` builds only the files with no report yet: a pass over a corpus
  // picks up where the last one stopped.
  const missing = process.argv.includes('--missing');
  for (const file of files) {
    if (missing && existsSync(resolve(OUT_DIR, stemOf(file), 'report.json'))) continue;
    let r: Report;
    try {
      r = await build(file);
    } catch (e) {
      process.stdout.write(`${stemOf(file).padEnd(40)} !! ${(e as Error).message}\n`);
      continue;
    }
    const worst = r.scores.reduce((m, s) => Math.max(m, s.missing, s.extra), 0);
    const pagesNote = `pages gold ${String(r.pages.gold)} ours ${String(r.pages.ours)}`;
    const errorsNote = r.errors.length > 0 ? `  !! ${r.errors.join(' | ')}` : '';
    process.stdout.write(
      `${r.name.padEnd(40)} worst ${worst.toFixed(3)}  ${pagesNote}  ${String(r.ms.total)} ms${errorsNote}\n`,
    );
  }
}

if (process.argv[1] !== undefined && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await main();
}
