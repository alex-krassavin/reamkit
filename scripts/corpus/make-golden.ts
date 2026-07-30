// Generate the golden page-geometry summaries for tests/fixtures/real/.
//
// The corpus harness (run.ts) needs LibreOffice, a raster pass and a judgement
// call per document; nothing about it can run inside `npm test`. This script
// front-loads that work: LibreOffice renders each adopted document ONCE, here,
// and what gets committed is a small text summary — page count, page size,
// per-page text. The test then runs offline against the summary.
//
// What the summary is FOR matters. Page size and orientation come straight from
// the print model, so LibreOffice and Ream must agree exactly and the test says
// so. Pagination and placement depend on typesetting, where we do not match
// LibreOffice and do not claim to; those numbers are recorded (both sides) so
// the distance is visible and any drift in it fails the test. A golden file is
// a measurement, not an aspiration.
//
// Usage: npx tsx scripts/corpus/make-golden.ts [file.xlsx ...]

import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { basename, dirname, resolve } from 'node:path';

import { referenceToPdf, stext } from './lib';
import type { FontBytesByVariant } from '@/core/font';
import { Ream } from '@/core/converter/ream';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
const fixtureDir = resolve(root, 'tests/fixtures/real');
const goldenDir = resolve(fixtureDir, 'golden');
const workDir = resolve(root, 'corpus/.work-golden');

const FONTS: FontBytesByVariant = {
  regular: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Bold.ttf'))),
  italic: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Italic.ttf'))),
  boldItalic: new Uint8Array(
    readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
  ),
};

/** One page's geometry, rounded to whole points — sub-point noise is not signal. */
export interface GoldenPage {
  readonly widthPt: number;
  readonly heightPt: number;
}

/** A renderer's view of one document. */
export interface GoldenRender {
  readonly pages: number;
  /** Page geometry, deduplicated: most documents are uniform, and a repeated
   *  row per page would bury a real mid-document size change in noise. */
  readonly pageSizes: ReadonlyArray<GoldenPage>;
  /** Extracted characters — a coarse content-coverage figure. */
  readonly chars: number;
  /** First line of text per page, so a wholesale reordering cannot hide. */
  readonly firstLines: ReadonlyArray<string>;
}

export interface Golden {
  readonly file: string;
  /**
   * Whether the workbook names its paper size (§18.3.1.63 `pageSetup@paperSize`).
   *
   * When it does, both renderers must honour it and the test compares sizes
   * strictly. When it does not, there is no right answer in the file: Excel
   * picks by locale and printer, LibreOffice picks by locale (Letter on the
   * machine that generated these), and we pick a deterministic A4. Comparing
   * sizes there would be measuring the golden machine's locale, not our code.
   */
  readonly paperDeclared: boolean;
  /** The reference: what LibreOffice makes of this document. */
  readonly libreOffice: GoldenRender;
  /** What we make of it, at the revision that generated this file. */
  readonly ream: GoldenRender;
}

/** Group characters into lines by rounded baseline, then read them out in order. */
function summarize(pdfPath: string, tag: string): GoldenRender {
  const pages = stext(pdfPath, resolve(workDir, `${tag}.xml`));
  const sizes = new Map<string, GoldenPage>();
  const firstLines: Array<string> = [];
  let chars = 0;
  for (const page of pages) {
    const size = { widthPt: Math.round(page.width), heightPt: Math.round(page.height) };
    sizes.set(`${size.widthPt}x${size.heightPt}`, size);
    chars += page.chars.length;

    const byBaseline = new Map<number, Array<{ x: number; c: string }>>();
    for (const box of page.chars) {
      const key = Math.round(box.y);
      let line = byBaseline.get(key);
      if (!line) byBaseline.set(key, (line = []));
      line.push({ x: box.x, c: box.c });
    }
    const topmost = [...byBaseline.entries()].sort((a, b) => a[0] - b[0])[0];
    const text = topmost
      ? topmost[1]
          .sort((a, b) => a.x - b.x)
          .map((e) => e.c)
          .join('')
          .trim()
      : '';
    firstLines.push(text.slice(0, 60));
  }
  return { pages: pages.length, pageSizes: [...sizes.values()], chars, firstLines };
}

async function main(): Promise<void> {
  const only = process.argv.slice(2);
  const files = readdirSync(fixtureDir)
    .filter((f) => /\.xlsx$/i.test(f))
    .filter((f) => only.length === 0 || only.includes(f))
    .sort();

  rmSync(workDir, { recursive: true, force: true });
  mkdirSync(workDir, { recursive: true });
  mkdirSync(goldenDir, { recursive: true });

  for (const file of files) {
    const input = resolve(fixtureDir, file);
    const tag = basename(file, '.xlsx');
    try {
      const ourPath = resolve(workDir, `${tag}.ream.pdf`);
      writeFileSync(
        ourPath,
        await Ream.parse(new Uint8Array(readFileSync(input))).convert('pdf', {
          fonts: FONTS,
          fileName: basename(input),
        }),
      );
      const golden: Golden = {
        file,
        paperDeclared: readXlsxToSheetDoc(new Uint8Array(readFileSync(input))).sheets.some(
          (s) => s.grid.pageSetup?.paperSize !== undefined,
        ),
        libreOffice: summarize(referenceToPdf(input, workDir), `${tag}.lo`),
        ream: summarize(ourPath, `${tag}.ream`),
      };
      writeFileSync(resolve(goldenDir, `${tag}.json`), JSON.stringify(golden, null, 2) + '\n');
      const lo = golden.libreOffice;
      const us = golden.ream;
      const sizeMatch = JSON.stringify(lo.pageSizes) === JSON.stringify(us.pageSizes);
      const agree = sizeMatch ? '✅' : golden.paperDeclared ? '❌' : '·';
      console.log(
        `${agree} ${file.padEnd(42)} pages ${String(us.pages).padStart(3)}/${String(lo.pages).padEnd(3)}` +
          ` chars ${String(us.chars).padStart(6)}/${String(lo.chars).padEnd(6)}` +
          ` size ${us.pageSizes.map((s) => `${s.widthPt}x${s.heightPt}`).join(',')}` +
          ` vs ${lo.pageSizes.map((s) => `${s.widthPt}x${s.heightPt}`).join(',')}`,
      );
    } catch (err) {
      console.error(`✗ ${file}: ${(err as Error).message.slice(0, 100)}`);
    }
  }
  rmSync(workDir, { recursive: true, force: true });
}

await main();
