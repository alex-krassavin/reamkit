// Rank a corpus directory by how different the PAGE looks from LibreOffice's.
//
// `scout.ts` ranks by the text layer — how many characters and pages each side
// produced. That is the right first question and the wrong second one: it is
// blind to everything the eye actually catches. A cell painted plum instead of
// grey, a ranking printed upside down, a logo that never drew, an axis rule in
// the wrong colour — every one of those scores ZERO there, and every one of
// them was a real defect found by opening the picture.
//
// So this renders both sides to pixels and sorts by the fraction that differ.
// Read it as a pointer, not a verdict: a page LibreOffice paginates differently
// scores near 1 without a single wrong pixel on it, and our own font is not the
// workbook's, so text-heavy pages carry a permanent few percent. What the list
// buys is the other end — a file at 0.000 is one nobody needs to open.
//
// LibreOffice's side is cached under corpus/.lo-cache (keyed by the source's
// bytes), because its output cannot change and it is four fifths of the wall
// clock: a re-run after a fix compares against the same reference for free.
//
// Usage:
//   npx tsx scripts/corpus/pixel-scout.ts corpus/external/lo-xlsx [limit] [offset]
//   npx tsx scripts/corpus/pixel-scout.ts corpus/external/poi-xlsx 500 0 --dpi 72
//   npx tsx scripts/corpus/pixel-scout.ts corpus/external/lo-docx-import
//   npx tsx scripts/corpus/pixel-scout.ts corpus/external/lo-pptx

import { createHash } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { basename, resolve } from 'node:path';

import { colorDiff, parsePpm, rasterize, referenceToPdf, visualDiff } from './lib';
import { corpusFontOptions } from './fonts';
import { Ream } from '@/core/converter/ream';

// The document's OWN families, substituted the way LibreOffice substitutes them
// (see ./fonts) — not one pinned face, which used to skip the font pipeline
// whole and charge the difference to layout.
const FONT_OPTIONS = corpusFontOptions();

// Which line breaker to render with. Ream's own default is Knuth-Plass
// total-fit; Word and LibreOffice break greedily, so a tight box can break one
// word differently for a reason that is not a fidelity gap. Set
// CORPUS_LAYOUT_PROFILE=libreoffice to take the reference's breaker and measure
// only what is left.
const LAYOUT_PROFILE = ((): 'ream' | 'word' | 'libreoffice' | undefined => {
  const v = process.env['CORPUS_LAYOUT_PROFILE'];
  return v === 'ream' || v === 'word' || v === 'libreoffice' ? v : undefined;
})();

// Enough to see a missing block, a wrong fill or a shifted column; not enough
// to rank on anti-aliasing. A4 at 60 dpi is ~500×700.
const DEFAULT_DPI = 60;
// A 278-page workbook, or a 60-slide deck, does not need that many comparisons
// to say whether it differs — see samplePages for which ones it gets.
const MAX_PAGES = 8;

const CACHE_DIR = resolve('corpus/.lo-cache');

/** LibreOffice's PDF for `src`, converting once and keeping it by content hash. */
function cachedReference(src: string, work: string): string {
  // A PDF is its own reference: asking LibreOffice to "convert" one runs it
  // through Draw, which redraws the page rather than reporting it. What the
  // comparison wants is the file as its author wrote it.
  if (/\.pdf$/iu.test(src)) return src;
  mkdirSync(CACHE_DIR, { recursive: true });
  const key = createHash('sha256').update(readFileSync(src)).digest('hex').slice(0, 16);
  const cached = resolve(CACHE_DIR, `${key}.pdf`);
  if (existsSync(cached)) return cached;
  const produced = referenceToPdf(src, work);
  copyFileSync(produced, cached);
  return cached;
}

/**
 * Which page indices to compare out of `count`, at most {@link MAX_PAGES} of
 * them, spread evenly from the first to the last. Everything shorter than the
 * cap is compared whole, which is nearly every document in the corpus.
 */
function samplePages(count: number): Array<number> {
  if (count <= MAX_PAGES) return Array.from({ length: count }, (_, i) => i);
  const step = (count - 1) / (MAX_PAGES - 1);
  return Array.from({ length: MAX_PAGES }, (_, i) => Math.round(i * step));
}

interface Row {
  readonly score: number;
  readonly line: string;
  readonly name: string;
  /**
   * Whether the two sides paginate the same. When they do not, page N is not
   * page N and the pixel score is comparing two different sheets — the number
   * is meaningless and the file belongs in its own list.
   */
  readonly aligned: boolean;
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const flagAt = args.findIndex((a) => a.startsWith('--'));
  const positional = flagAt < 0 ? args : args.slice(0, flagAt);
  const dpiFlag = args.indexOf('--dpi');
  const dpi = dpiFlag >= 0 ? Number(args[dpiFlag + 1]) : DEFAULT_DPI;
  const dir = positional[0];
  if (dir === undefined) throw new Error('usage: pixel-scout.ts <corpus-dir> [limit] [offset]');
  const limit = Number(positional[1] ?? 1000);
  const offset = Number(positional[2] ?? 0);

  const work = resolve(`corpus/.pixel-scout-${String(process.pid)}`);
  rmSync(work, { recursive: true, force: true });
  mkdirSync(work, { recursive: true });

  const files = readdirSync(dir)
    // Every input the converter reads. A deck ranks like a document — one slide
    // is one page on both sides — and so do the legacy binaries and a PDF read
    // back out: 46 files of the corpus are those, and none of them had ever
    // been compared to a picture.
    .filter((f) => /\.(?:xlsx|docx|pptx|doc|xls|ppt|pdf)$/iu.test(f))
    .sort()
    .slice(offset, offset + limit);

  const rows: Array<Row> = [];
  for (const name of files) {
    const src = resolve(dir, name);
    const oursPdf = resolve(work, 'ours.pdf');
    try {
      writeFileSync(
        oursPdf,
        // LibreOffice heads a page with the file's NAME when the sheet asks for
        // `&F`; a byte-oriented API has to be told it.
        await Ream.parse(new Uint8Array(readFileSync(src))).convert('pdf', {
          fileName: basename(src),
          ...FONT_OPTIONS,
          ...(LAYOUT_PROFILE !== undefined ? { layoutProfile: LAYOUT_PROFILE } : {}),
        }),
      );
    } catch (e) {
      rows.push({
        score: 99,
        name,
        aligned: false,
        line: `  threw   ${name} — ${(e as Error).message.slice(0, 60)}`,
      });
      continue;
    }
    let refPdf: string;
    try {
      refPdf = cachedReference(src, work);
    } catch {
      rows.push({
        score: 98,
        name,
        aligned: false,
        line: `  ref!    ${name} — soffice produced no PDF`,
      });
      continue;
    }

    rmSync(resolve(work, 'px'), { recursive: true, force: true });
    mkdirSync(resolve(work, 'px'), { recursive: true });
    let ourPages: Array<string>;
    let refPages: Array<string>;
    try {
      ourPages = rasterize(oursPdf, resolve(work, 'px/o%d.ppm'), dpi);
      refPages = rasterize(refPdf, resolve(work, 'px/r%d.ppm'), dpi);
    } catch (e) {
      rows.push({
        score: 97,
        name,
        aligned: false,
        line: `  raster! ${name} — ${(e as Error).message.slice(0, 50)}`,
      });
      continue;
    }

    // The worst page carries the file: one wrong page among twenty is exactly
    // the case worth opening, and averaging hides it. Which pages get looked at
    // is spread over the whole document rather than taken off the front — a
    // deck opens on a title slide and keeps its charts and tables in the
    // middle, so the first eight of forty are the eight least worth ranking on.
    let worst = 0;
    let worstPage = 0;
    let worstColor = 0;
    for (const i of samplePages(Math.min(ourPages.length, refPages.length))) {
      const ourPpm = parsePpm(new Uint8Array(readFileSync(ourPages[i]!)));
      const refPpm = parsePpm(new Uint8Array(readFileSync(refPages[i]!)));
      // Two pixels of slack: enough that a glyph landing a hair to the left is
      // not a difference, not enough to hide one that moved a column.
      const d = visualDiff(ourPpm, refPpm, 24, 2);
      const c = colorDiff(ourPpm, refPpm);
      if (c > worstColor) worstColor = c;
      if (d.mismatchRatio > worst) {
        worst = d.mismatchRatio;
        worstPage = i + 1;
      }
    }
    const aligned = ourPages.length === refPages.length;
    // Rank on the colour distance: it is the one that answers "is something
    // painted wrong or missing", which is what a look at the page finds. The
    // pixel ratio rides along because it says whether the page MOVED.
    rows.push({
      score: worstColor,
      name,
      aligned,
      line:
        `${worstColor.toFixed(3).padStart(8)} ${worst.toFixed(3).padStart(7)}px  ` +
        `${name.padEnd(46)} p${String(worstPage).padStart(2)}  ` +
        `pages ${ourPages.length}/${refPages.length}`,
    });
  }

  rmSync(work, { recursive: true, force: true });
  const sorted = rows.sort((a, b) => b.score - a.score);
  // The files whose pages line up: here the number means what it says, and the
  // top of this list is where to look next.
  console.log('──  colour  pixels  file ──');
  for (const r of sorted.filter((x) => x.aligned)) console.log(r.line);
  // …and the rest, where page N is not page N. A pagination gap is its own
  // question and usually its own cause (a band we trim, a row we fit).
  console.log('\n── pagination differs (pixel score not comparable) ──');
  for (const r of sorted.filter((x) => !x.aligned)) console.log(r.line);
  // One ranking per corpus directory: a sweep of the decks must not overwrite
  // the record of the documents, and the name says which set it ranked.
  writeFileSync(
    resolve(`corpus/.pixel-scout-${basename(resolve(dir))}.json`),
    JSON.stringify(
      sorted.map((r) => ({ name: r.name, score: Number(r.score.toFixed(4)), aligned: r.aligned })),
      null,
      1,
    ),
  );
}

await main();
