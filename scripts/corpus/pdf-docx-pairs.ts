// Put OUR .docx beside the one LibreOffice makes of the same PDF, page by page,
// as one picture per page — the thing to actually look at.
//
// `pdf-docx-vs-lo.ts` scores that same pair; this one draws it. A score says a
// page differs, and every page differs; only the picture says whether what is
// left is a document.
//
// Usage:
//   npx tsx scripts/corpus/pdf-docx-pairs.ts corpus/external/pdfjs 10 0
//   npx tsx scripts/corpus/pdf-docx-pairs.ts corpus/external/pdfjs 10 10 --pages 2

import { execFileSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

import { parsePpm, rasterize } from './lib';
import type { Ppm } from './lib';
import { Ream } from '@/core/converter/ream';
import { encodePng } from '@/core/png-encode';

const OUT_DIR = resolve('corpus/.pairs');

/**
 * One `soffice` run with a profile of its OWN.
 *
 * A shared profile is one running instance: a second call while the first is
 * busy hands its job to that instance and exits, so one file that hangs makes
 * every file after it produce nothing. A profile per call cannot do that.
 */
function soffice(args: ReadonlyArray<string>, tag: string): void {
  const profile = pathToFileURL(resolve(tmpdir(), `ream-pair-${tag}`)).href;
  execFileSync('soffice', [`-env:UserInstallation=${profile}`, '--headless', ...args], {
    stdio: 'ignore',
    timeout: 120_000,
  });
}

/** Lay two rasters side by side, with a rule between them. */
function pair(left: Ppm, right: Ppm): { png: Uint8Array; width: number; height: number } {
  const gap = 8;
  const width = left.width + gap + right.width;
  const height = Math.max(left.height, right.height);
  const rgb = new Uint8Array(width * height * 3).fill(0xf0);
  const blit = (src: Ppm, x0: number): void => {
    for (let y = 0; y < src.height; y++) {
      for (let x = 0; x < src.width; x++) {
        const s = (y * src.width + x) * 3;
        const d = (y * width + x0 + x) * 3;
        rgb[d] = src.rgb[s] ?? 0xff;
        rgb[d + 1] = src.rgb[s + 1] ?? 0xff;
        rgb[d + 2] = src.rgb[s + 2] ?? 0xff;
      }
    }
  };
  blit(left, 0);
  blit(right, left.width + gap);
  return { png: encodePng(width, height, 'rgb', rgb), width, height };
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const flagAt = args.findIndex((a) => a.startsWith('--'));
  const positional = flagAt < 0 ? args : args.slice(0, flagAt);
  const pagesFlag = args.indexOf('--pages');
  const wanted = pagesFlag >= 0 ? Number(args[pagesFlag + 1]) : 1;
  const dpiFlag = args.indexOf('--dpi');
  const dpi = dpiFlag >= 0 ? Number(args[dpiFlag + 1]) : 70;
  const dir = positional[0];
  if (dir === undefined) throw new Error('usage: pdf-docx-pairs.ts <corpus-dir> [limit] [offset]');
  const limit = Number(positional[1] ?? 10);
  const offset = Number(positional[2] ?? 0);

  rmSync(OUT_DIR, { recursive: true, force: true });
  mkdirSync(OUT_DIR, { recursive: true });
  const work = mkdtempSync(join(tmpdir(), 'ream-pairs-'));

  const files = readdirSync(dir)
    .filter((f) => /\.pdf$/iu.test(f))
    .sort()
    .slice(offset, offset + limit);

  for (const name of files) {
    const stem = name.replace(/\.pdf$/iu, '');
    const src = resolve(dir, name);
    const box = resolve(work, stem);
    mkdirSync(box, { recursive: true });
    let note = '';
    try {
      soffice(['--infilter=writer_pdf_import', '--convert-to', 'docx', '--outdir', box, src], stem);
      const theirs = resolve(box, `${stem}.docx`);
      if (!existsSync(theirs)) throw new Error('LibreOffice made no .docx');
      const ours = resolve(box, 'ours.docx');
      writeFileSync(ours, await Ream.parse(new Uint8Array(readFileSync(src))).convert('docx'));
      soffice(['--convert-to', 'pdf', '--outdir', box, theirs], `${stem}-t`);
      soffice(['--convert-to', 'pdf', '--outdir', box, ours], `${stem}-o`);
      const theirPdf = resolve(box, `${stem}.pdf`);
      const ourPdf = resolve(box, 'ours.pdf');
      if (!existsSync(ourPdf)) throw new Error('OUR .docx would not open');
      if (!existsSync(theirPdf)) throw new Error('their .docx would not render');
      const a = rasterize(ourPdf, resolve(box, 'o%d.ppm'), dpi);
      const b = rasterize(theirPdf, resolve(box, 't%d.ppm'), dpi);
      const pages = Math.min(wanted, a.length, b.length);
      for (let i = 0; i < pages; i++) {
        const composed = pair(
          parsePpm(new Uint8Array(readFileSync(a[i]!))),
          parsePpm(new Uint8Array(readFileSync(b[i]!))),
        );
        writeFileSync(resolve(OUT_DIR, `${stem}-p${String(i + 1)}.png`), composed.png);
      }
      note = `${String(a.length)} pages ours, ${String(b.length)} theirs`;
    } catch (e) {
      note = `!! ${(e as Error).message.slice(0, 60)}`;
    }
    process.stdout.write(`  ${basename(name).padEnd(44)} ${note}\n`);
  }
  rmSync(work, { recursive: true, force: true });
  process.stdout.write(`\nOURS on the left, LibreOffice on the right → ${OUT_DIR}\n`);
}

await main();
