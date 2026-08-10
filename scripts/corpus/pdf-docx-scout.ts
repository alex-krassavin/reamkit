// Rank a PDF corpus by how little the WORD DOCUMENT made from it looks like it.
//
// `pixel-scout.ts` measures the reader: it renders a PDF through Ream's own
// page emitter and compares that to the file. That answers "did we understand
// the page", and it answered it well — the pdf.js corpus sits near zero. It
// says nothing at all about the question a user actually asks, which is what
// happens when they open the .docx in Word.
//
// This measures the whole chain instead:
//
//   the file            →  raster            (the reference: the page as written)
//   the file → Ream → .docx → LibreOffice → raster   (what a reader is handed)
//
// LibreOffice stands in for Word because Word cannot be driven four hundred
// times. It is a weaker oracle than Word — it opened every document Word
// refused for a schema violation — so treat a low score as necessary and not
// sufficient, and open the worst files in Word by hand.
//
// Both readings are measured, because a PDF has two right answers and which one
// is right depends on the document. `flow` reads it as prose and repaginates;
// `positional` keeps every line where its glyphs stand. A report wants the
// first, a form the second, and a converter that only does one of them is wrong
// half the time.
//
// Usage:
//   npx tsx scripts/corpus/pdf-docx-scout.ts corpus/external/pdfjs [limit] [offset]
//   npx tsx scripts/corpus/pdf-docx-scout.ts corpus/external/pdfjs 60 0 --dpi 72
//   npx tsx scripts/corpus/pdf-docx-scout.ts corpus/external/pdfjs --mode positional

import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

import { colorDiff, parsePpm, rasterize, referenceToPdf, visualDiff } from './lib';
import { corpusFontOptions } from './fonts';
import { Ream } from '@/core/converter/ream';

const FONT_OPTIONS = corpusFontOptions();
const DEFAULT_DPI = 60;
const MAX_PAGES = 6;

type Mode = 'flow' | 'positional';

/** One reading of one file: what it cost, and how far the result landed. */
interface Reading {
  /** 0–1 where the pages line up; a sentinel above 1 where the chain broke. */
  readonly score: number;
  /** Why it broke, when it did. */
  readonly note: string;
  readonly ourPages: number;
  readonly refPages: number;
}

interface Row {
  readonly name: string;
  readonly best: number;
  readonly readings: ReadonlyMap<Mode, Reading>;
}

function samplePages(count: number): Array<number> {
  if (count <= MAX_PAGES) return Array.from({ length: count }, (_, i) => i);
  const step = (count - 1) / (MAX_PAGES - 1);
  return Array.from({ length: MAX_PAGES }, (_, i) => Math.round(i * step));
}

/**
 * Run one file through one reading and measure what came back.
 *
 * @param src  The source PDF.
 * @param mode Which reading of it to take.
 * @param work A scratch directory this call owns.
 * @param dpi  Raster resolution for the comparison.
 */
async function measure(src: string, mode: Mode, work: string, dpi: number): Promise<Reading> {
  const none = { ourPages: 0, refPages: 0 };
  const bytes = new Uint8Array(readFileSync(src));
  const docx = resolve(work, 'ours.docx');
  try {
    const doc = Ream.parse(bytes, mode === 'positional' ? { pdfLayout: 'positional' } : {});
    writeFileSync(docx, await doc.convert('docx', FONT_OPTIONS));
  } catch (e) {
    return { score: 99, note: `convert: ${(e as Error).message.slice(0, 48)}`, ...none };
  }

  // The .docx through a Word-family renderer — the step that tells us whether
  // what we wrote is a document at all.
  let oursPdf: string;
  try {
    oursPdf = referenceToPdf(docx, work);
  } catch {
    return { score: 98, note: 'soffice would not open our .docx', ...none };
  }

  const px = resolve(work, 'px');
  rmSync(px, { recursive: true, force: true });
  mkdirSync(px, { recursive: true });
  let ours: Array<string>;
  let refs: Array<string>;
  try {
    ours = rasterize(oursPdf, resolve(px, 'o%d.ppm'), dpi);
    refs = rasterize(src, resolve(px, 'r%d.ppm'), dpi);
  } catch (e) {
    return { score: 97, note: `raster: ${(e as Error).message.slice(0, 44)}`, ...none };
  }
  if (refs.length === 0) return { score: 96, note: 'the source rendered no pages', ...none };
  if (ours.length === 0) {
    return { score: 95, note: 'our document has no pages', ourPages: 0, refPages: refs.length };
  }

  // The worst compared page carries the file: one ruined page among ten is
  // exactly the case worth opening, and an average hides it.
  let worst = 0;
  for (const i of samplePages(Math.min(ours.length, refs.length))) {
    const o = parsePpm(new Uint8Array(readFileSync(ours[i]!)));
    const r = parsePpm(new Uint8Array(readFileSync(refs[i]!)));
    worst = Math.max(worst, Math.max(visualDiff(o, r, 24, 2).mismatchRatio, colorDiff(o, r)));
  }
  const pageNote = ours.length === refs.length ? '' : `pages ${ours.length}≠${refs.length}`;
  return { score: worst, note: pageNote, ourPages: ours.length, refPages: refs.length };
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const flagAt = args.findIndex((a) => a.startsWith('--'));
  const positional = flagAt < 0 ? args : args.slice(0, flagAt);
  const dpiFlag = args.indexOf('--dpi');
  const dpi = dpiFlag >= 0 ? Number(args[dpiFlag + 1]) : DEFAULT_DPI;
  const modeFlag = args.indexOf('--mode');
  const modes: Array<Mode> = modeFlag >= 0 ? [args[modeFlag + 1] as Mode] : ['flow', 'positional'];

  const dir = positional[0];
  if (dir === undefined) throw new Error('usage: pdf-docx-scout.ts <corpus-dir> [limit] [offset]');
  const limit = Number(positional[1] ?? 1000);
  const offset = Number(positional[2] ?? 0);

  // Outside the repo on purpose. A sweep takes long enough that the tree gets
  // worked on while it runs, and a work directory under `corpus/` is one
  // `rm -rf` away from every remaining file reporting ENOENT as if the
  // converter had failed. This one cost a 400-file run.
  const work = mkdtempSync(join(tmpdir(), 'ream-pdf-docx-'));

  const files = readdirSync(dir)
    .filter((f) => /\.pdf$/iu.test(f))
    .sort()
    .slice(offset, offset + limit);

  const rows: Array<Row> = [];
  for (const name of files) {
    const readings = new Map<Mode, Reading>();
    for (const mode of modes)
      readings.set(mode, await measure(resolve(dir, name), mode, work, dpi));
    const best = Math.min(...[...readings.values()].map((r) => r.score));
    rows.push({ name, best, readings });
    const cells = [...readings]
      .map(([m, r]) => `${m[0]!}=${r.score > 1 ? r.note : r.score.toFixed(3)}`)
      .join('  ');
    process.stdout.write(`${best > 1 ? '  !' : '   '} ${best.toFixed(3)}  ${cells}  ${name}\n`);
  }

  rows.sort((a, b) => b.best - a.best);
  const ok = rows.filter((r) => r.best <= 1);
  const broke = rows.filter((r) => r.best > 1);
  process.stdout.write(`\n=== ${String(rows.length)} files, worst first\n`);
  for (const r of rows.slice(0, 40)) {
    const cells = [...r.readings]
      .map(([m, x]) => `${m}=${x.score > 1 ? x.note : x.score.toFixed(3)}`)
      .join('  ')
      .padEnd(46);
    process.stdout.write(`  ${r.best > 1 ? '  !  ' : r.best.toFixed(3)}  ${cells}  ${r.name}\n`);
  }
  const sum = ok.reduce((a, r) => a + r.best, 0);
  process.stdout.write(
    `\n${String(broke.length)} produced nothing comparable; ` +
      `${String(ok.length)} compared, summing ${sum.toFixed(3)} ` +
      `(mean ${(sum / Math.max(ok.length, 1)).toFixed(3)})\n`,
  );

  writeFileSync(
    resolve('corpus/.pdf-docx-scout.json'),
    `${JSON.stringify(
      rows.map((r) => ({
        name: r.name,
        best: Number(r.best.toFixed(4)),
        ...Object.fromEntries(
          [...r.readings].map(([m, x]) => [
            m,
            { score: Number(x.score.toFixed(4)), note: x.note, pages: [x.ourPages, x.refPages] },
          ]),
        ),
      })),
      null,
      1,
    )}\n`,
  );
  rmSync(work, { recursive: true, force: true });
}

await main();
