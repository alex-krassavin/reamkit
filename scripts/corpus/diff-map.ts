// Where two renders disagree, as a picture.
//
// The scout's two numbers say HOW MUCH a page differs, never WHERE: `colorDiff`
// compares colour histograms and knows nothing of position, and the pixel score
// is one fraction for the whole sheet. On a dense drawing that is not enough to
// work from — Brotli-Prototype-FileA.pdf lost its entire title block and the
// colour score did not move at all.
//
// This marks every pixel the two renders disagree on, in red, over a faded copy
// of the reference. What is left is a map of what to look at next.
//
// Usage:
//   npx tsx scripts/corpus/diff-map.ts corpus/external/pdfjs/160F-2019.pdf
//   npx tsx scripts/corpus/diff-map.ts <file> --page 3 --dpi 144

import { createRequire } from 'node:module';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { parseOptions, parsePpm, rasterize, referenceToPdf } from './lib';
import { corpusFontOptions } from './fonts';
import { Ream } from '@/core/converter/ream';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
const outDir = resolve(root, 'corpus/.visual');
const workDir = resolve(root, `corpus/.diff-work-${String(process.pid)}`);

const requireFromDocs = createRequire(resolve(root, 'docs/package.json'));

const sharp = requireFromDocs('sharp');

const arg = (flag: string): string | undefined => {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : undefined;
};

/** How far apart two colours must be, per channel summed, to count as different. */
const TOLERANCE = 24 * 3;

async function main(): Promise<void> {
  const input = process.argv[2];
  if (!input) throw new Error('usage: diff-map.ts <file> [--page N] [--dpi N]');
  const page = Number(arg('--page') ?? '1');
  const dpi = Number(arg('--dpi') ?? '72');

  rmSync(workDir, { recursive: true, force: true });
  mkdirSync(workDir, { recursive: true });
  mkdirSync(outDir, { recursive: true });

  const ourPdf = resolve(workDir, 'ours.pdf');
  writeFileSync(
    ourPdf,
    await Ream.parse(new Uint8Array(readFileSync(input)), parseOptions(input)).convert('pdf', {
      ...corpusFontOptions(),
      fileName: basename(input),
    }),
  );
  // A PDF is its own reference; anything else goes through LibreOffice.
  const refPdf = /\.pdf$/iu.test(input) ? input : referenceToPdf(input, workDir);
  const ours = rasterize(ourPdf, resolve(workDir, 'ours-%d.ppm'), dpi);
  const refs = rasterize(refPdf, resolve(workDir, 'ref-%d.ppm'), dpi);
  const a = ours[page - 1];
  const b = refs[page - 1];
  if (!a || !b)
    throw new Error(`page ${String(page)} missing (ours ${ours.length}, ref ${refs.length})`);

  const our = parsePpm(new Uint8Array(readFileSync(a)));
  const ref = parsePpm(new Uint8Array(readFileSync(b)));
  const w = Math.min(our.width, ref.width);
  const h = Math.min(our.height, ref.height);
  const rgb = new Uint8Array(w * h * 3);
  let differing = 0;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const oi = (y * our.width + x) * 3;
      const ri = (y * ref.width + x) * 3;
      const d =
        Math.abs(our.rgb[oi]! - ref.rgb[ri]!) +
        Math.abs(our.rgb[oi + 1]! - ref.rgb[ri + 1]!) +
        Math.abs(our.rgb[oi + 2]! - ref.rgb[ri + 2]!);
      const o = (y * w + x) * 3;
      if (d > TOLERANCE) {
        differing++;
        // Red where only WE painted, blue where only the reference did.
        const ourInk = 765 - (our.rgb[oi]! + our.rgb[oi + 1]! + our.rgb[oi + 2]!);
        const refInk = 765 - (ref.rgb[ri]! + ref.rgb[ri + 1]! + ref.rgb[ri + 2]!);
        if (ourInk > refInk) {
          rgb[o] = 255;
        } else {
          rgb[o + 2] = 255;
        }
      } else {
        // The reference, faded, so the marks have something to sit on.
        const grey = 200 + Math.round((ref.rgb[ri]! * 55) / 255);
        rgb[o] = grey;
        rgb[o + 1] = grey;
        rgb[o + 2] = grey;
      }
    }
  }
  const out = resolve(outDir, `${basename(input)}-diff-p${String(page).padStart(2, '0')}.png`);
  await sharp(Buffer.from(rgb), { raw: { width: w, height: h, channels: 3 } }).toFile(out);
  rmSync(workDir, { recursive: true, force: true });
  console.log(
    `${String(differing)} of ${String(w * h)} px differ (${((differing / (w * h)) * 100).toFixed(2)}%)`,
  );
  console.log('red = ours only, blue = reference only');
  console.log(out);
}

await main();
