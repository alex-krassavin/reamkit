// Rank a PDF corpus by how far OUR .docx is from the one LibreOffice makes of
// the same file.
//
// `pdf-docx-scout.ts` compares our .docx to the SOURCE PAGE, and that comparison
// has a floor nobody can reach: a reflowed document is not the page it came
// from, so every honest conversion scores badly and the number never says
// whether the result is any good. It also never noticed a package that would
// not open at all, because a file that never becomes a page has no pages to
// compare.
//
// This compares two conversions of the same file instead:
//
//   the file → LibreOffice → .docx → LibreOffice → raster   (what a converter does)
//   the file → Ream        → .docx → LibreOffice → raster   (what we do)
//
// Same medium, same renderer, same question — so the difference is ours, and
// "as good as LibreOffice" is a target that can actually be hit. LibreOffice
// imports a PDF through Draw (`writer_pdf_import`), which is a real conversion
// with real decisions in it, and its output is a document a person would accept.
//
// Usage:
//   npx tsx scripts/corpus/pdf-docx-vs-lo.ts corpus/external/pdfjs [limit] [offset]
//   npx tsx scripts/corpus/pdf-docx-vs-lo.ts corpus/external/pdfjs 40 0 --dpi 72

import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

import { colorDiff, parsePpm, rasterize, referenceToPdf, visualDiff } from './lib';
import { Ream } from '@/core/converter/ream';

const DEFAULT_DPI = 60;
const MAX_PAGES = 6;

/** LibreOffice's own PDF → .docx, which is the thing we are trying to match. */
function loConvertToDocx(pdf: string, outDir: string): string {
  const profile = pathToFileURL(resolve(tmpdir(), `ream-lo-p2d-${String(process.pid)}`)).href;
  execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${profile}`,
      '--headless',
      '--infilter=writer_pdf_import',
      '--convert-to',
      'docx',
      '--outdir',
      outDir,
      pdf,
    ],
    { stdio: 'ignore', timeout: 180_000 },
  );
  const out = resolve(outDir, `${basename(pdf).replace(/\.pdf$/iu, '')}.docx`);
  if (!existsSync(out)) throw new Error('no docx');
  return out;
}

function samplePages(count: number): Array<number> {
  if (count <= MAX_PAGES) return Array.from({ length: count }, (_, i) => i);
  const step = (count - 1) / (MAX_PAGES - 1);
  return Array.from({ length: MAX_PAGES }, (_, i) => Math.round(i * step));
}

interface Row {
  readonly name: string;
  readonly score: number;
  readonly note: string;
  readonly pages: readonly [number, number];
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const flagAt = args.findIndex((a) => a.startsWith('--'));
  const positional = flagAt < 0 ? args : args.slice(0, flagAt);
  const dpiFlag = args.indexOf('--dpi');
  const dpi = dpiFlag >= 0 ? Number(args[dpiFlag + 1]) : DEFAULT_DPI;
  const dir = positional[0];
  if (dir === undefined) throw new Error('usage: pdf-docx-vs-lo.ts <corpus-dir> [limit] [offset]');
  const limit = Number(positional[1] ?? 1000);
  const offset = Number(positional[2] ?? 0);

  const work = mkdtempSync(join(tmpdir(), 'ream-vs-lo-'));
  const files = readdirSync(dir)
    .filter((f) => /\.pdf$/iu.test(f))
    .sort()
    .slice(offset, offset + limit);

  const rows: Array<Row> = [];
  for (const name of files) {
    const src = resolve(dir, name);
    let theirs: string;
    try {
      theirs = loConvertToDocx(src, work);
    } catch {
      // LibreOffice could not convert it either, so there is nothing to match.
      rows.push({ score: -1, name, note: 'LibreOffice made no .docx', pages: [0, 0] });
      continue;
    }
    const ours = resolve(work, 'ours.docx');
    try {
      writeFileSync(ours, await Ream.parse(new Uint8Array(readFileSync(src))).convert('docx'));
    } catch (e) {
      rows.push({
        score: 99,
        name,
        note: `convert: ${(e as Error).message.slice(0, 44)}`,
        pages: [0, 0],
      });
      continue;
    }
    let theirPdf: string;
    let ourPdf: string;
    try {
      theirPdf = referenceToPdf(theirs, work);
    } catch {
      rows.push({ score: -1, name, note: 'their .docx would not render', pages: [0, 0] });
      continue;
    }
    try {
      ourPdf = referenceToPdf(ours, work);
    } catch {
      // The loudest failure there is: we wrote something no reader will open.
      rows.push({ score: 98, name, note: 'OUR .docx would not open', pages: [0, 0] });
      continue;
    }
    const px = resolve(work, 'px');
    rmSync(px, { recursive: true, force: true });
    mkdirSync(px, { recursive: true });
    let a: Array<string>;
    let b: Array<string>;
    try {
      a = rasterize(ourPdf, resolve(px, 'o%d.ppm'), dpi);
      b = rasterize(theirPdf, resolve(px, 't%d.ppm'), dpi);
    } catch (e) {
      rows.push({
        score: 97,
        name,
        note: `raster: ${(e as Error).message.slice(0, 40)}`,
        pages: [0, 0],
      });
      continue;
    }
    if (a.length === 0 || b.length === 0) {
      rows.push({ score: 96, name, note: 'no pages', pages: [a.length, b.length] });
      continue;
    }
    let worst = 0;
    for (const i of samplePages(Math.min(a.length, b.length))) {
      const o = parsePpm(new Uint8Array(readFileSync(a[i]!)));
      const t = parsePpm(new Uint8Array(readFileSync(b[i]!)));
      worst = Math.max(worst, Math.max(visualDiff(o, t, 24, 2).mismatchRatio, colorDiff(o, t)));
    }
    const note = a.length === b.length ? '' : `pages ${a.length}≠${b.length}`;
    rows.push({ score: worst, name, note, pages: [a.length, b.length] });
    process.stdout.write(`   ${worst.toFixed(3)}  ${note}  ${name}\n`);
  }

  const ok = rows.filter((r) => r.score >= 0 && r.score <= 1);
  const broke = rows.filter((r) => r.score > 1);
  const skipped = rows.filter((r) => r.score < 0);
  rows.sort((x, y) => y.score - x.score);
  process.stdout.write('\n=== worst first\n');
  for (const r of rows.filter((r) => r.score >= 0).slice(0, 30)) {
    process.stdout.write(
      `  ${r.score > 1 ? '  !  ' : r.score.toFixed(3)}  ${r.note.padEnd(30)}  ${r.name}\n`,
    );
  }
  const sum = ok.reduce((a, r) => a + r.score, 0);
  process.stdout.write(
    `\n${String(skipped.length)} LibreOffice could not convert either; ` +
      `${String(broke.length)} of ours produced nothing comparable; ` +
      `${String(ok.length)} compared, summing ${sum.toFixed(2)} ` +
      `(mean ${(sum / Math.max(ok.length, 1)).toFixed(3)})\n`,
  );
  writeFileSync(
    resolve('corpus/.pdf-docx-vs-lo.json'),
    `${JSON.stringify(
      rows.map((r) => ({
        name: r.name,
        score: Number(r.score.toFixed(4)),
        note: r.note,
        pages: r.pages,
      })),
      null,
      1,
    )}\n`,
  );
  rmSync(work, { recursive: true, force: true });
}

await main();
