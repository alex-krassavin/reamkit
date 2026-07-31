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

import { parsePpm, rasterize, referenceToPdf, visualDiff } from './lib';
import type { FontBytesByVariant } from '@/core/font';
import { Ream } from '@/core/converter/ream';

const FONTS: FontBytesByVariant = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

// Enough to see a missing block, a wrong fill or a shifted column; not enough
// to rank on anti-aliasing. A4 at 60 dpi is ~500×700.
const DEFAULT_DPI = 60;
// A 278-page workbook does not need 278 comparisons to say whether it differs.
const MAX_PAGES = 8;

const CACHE_DIR = resolve('corpus/.lo-cache');

/** LibreOffice's PDF for `src`, converting once and keeping it by content hash. */
function cachedReference(src: string, work: string): string {
  mkdirSync(CACHE_DIR, { recursive: true });
  const key = createHash('sha256').update(readFileSync(src)).digest('hex').slice(0, 16);
  const cached = resolve(CACHE_DIR, `${key}.pdf`);
  if (existsSync(cached)) return cached;
  const produced = referenceToPdf(src, work);
  copyFileSync(produced, cached);
  return cached;
}

interface Row {
  readonly score: number;
  readonly line: string;
  readonly name: string;
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
    .filter((f) => f.toLowerCase().endsWith('.xlsx'))
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
          fonts: FONTS,
        }),
      );
    } catch (e) {
      rows.push({ score: 99, name, line: `  threw   ${name} — ${(e as Error).message.slice(0, 60)}` });
      continue;
    }
    let refPdf: string;
    try {
      refPdf = cachedReference(src, work);
    } catch {
      rows.push({ score: 98, name, line: `  ref!    ${name} — soffice produced no PDF` });
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
      rows.push({ score: 97, name, line: `  raster! ${name} — ${(e as Error).message.slice(0, 50)}` });
      continue;
    }

    // The worst page carries the file: one wrong page among twenty is exactly
    // the case worth opening, and averaging hides it.
    const compared = Math.min(ourPages.length, refPages.length, MAX_PAGES);
    let worst = 0;
    let worstPage = 0;
    for (let i = 0; i < compared; i++) {
      const d = visualDiff(
        parsePpm(new Uint8Array(readFileSync(ourPages[i]!))),
        parsePpm(new Uint8Array(readFileSync(refPages[i]!))),
      );
      if (d.mismatchRatio > worst) {
        worst = d.mismatchRatio;
        worstPage = i + 1;
      }
    }
    // A page count that differs is a difference in itself, and the pages past
    // the shorter side were never compared.
    const pageGap = Math.abs(ourPages.length - refPages.length);
    const score = worst + (pageGap > 0 ? 0.05 : 0);
    rows.push({
      score,
      name,
      line:
        `${score.toFixed(3).padStart(8)}  ${name.padEnd(46)} ` +
        `p${String(worstPage).padStart(2)}  pages ${ourPages.length}/${refPages.length}`,
    });
  }

  rmSync(work, { recursive: true, force: true });
  const sorted = rows.sort((a, b) => b.score - a.score);
  for (const r of sorted) console.log(r.line);
  writeFileSync(
    resolve('corpus/.pixel-scout.json'),
    JSON.stringify(
      sorted.map((r) => ({ name: r.name, score: Number(r.score.toFixed(4)) })),
      null,
      1,
    ),
  );
}

await main();
