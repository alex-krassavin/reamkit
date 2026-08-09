// Render a document side by side with the reference, as an image.
//
// Every other metric in this harness is a number — characters extracted, pages
// emitted, a similarity ratio. Numbers are what make a regression diffable, but
// they are blind in a specific way: AverageTaxRates.xlsx rendered a hidden
// column and seven hidden rows that no other reader shows, and the metrics
// called it "446 characters more than LibreOffice" — technically true, easy to
// wave through, and obvious the instant anyone looked at the page.
//
// This produces that look: one PNG per page, ours on the left, the reference on
// the right, so the two can be compared without opening anything.
//
// Usage:
//   npx tsx scripts/corpus/visual-diff.ts AverageTaxRates
//   npx tsx scripts/corpus/visual-diff.ts corpus/external/lo-xlsx/tdf123353.xlsx
//   npx tsx scripts/corpus/visual-diff.ts tdf58243 --pages 1,3 --dpi 130

import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { basename, dirname, resolve } from 'node:path';

import { parseOptions, referenceToPdf } from './lib';
import { corpusFontOptions } from './fonts';
import { Ream } from '@/core/converter/ream';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
const outDir = resolve(root, 'corpus/.visual');
// Per-process scratch. One shared directory raced when two runs overlapped:
// each wipes it on entry and then lists it to collect the pages it just
// rasterised, so a run could pick up the OTHER document's PNGs and produce a
// side-by-side of two different files. It did — a comparison of one fixture
// showed a completely different spreadsheet in the "ours" pane, which is the
// worst possible failure for a harness whose whole job is to be believed.
const workDir = resolve(root, `corpus/.visual-work-${String(process.pid)}`);

// sharp is the docs site's dependency, not the library's — this is a dev-only
// viewer and the library must not grow a native image dependency for it.
const requireFromDocs = createRequire(resolve(root, 'docs/package.json'));
type Sharp = (input: Buffer | string | { create: unknown }) => {
  composite: (items: ReadonlyArray<Record<string, unknown>>) => { toFile: (p: string) => unknown };
  metadata: () => Promise<{ width?: number; height?: number }>;
};
const sharp = requireFromDocs('sharp') as Sharp;

const arg = (flag: string): string | undefined => {
  const i = process.argv.indexOf(flag);
  return i >= 0 ? process.argv[i + 1] : undefined;
};
const DPI = Number(arg('--dpi') ?? 100);
const LABEL_H = 26;
const GAP = 12;

/** The corpus directories a bare fixture name is looked for in, in order. */
const CORPUS_DIRS: ReadonlyArray<string> = [
  'tests/fixtures/real',
  'corpus/inputs',
  'corpus/external/lo-xlsx',
  'corpus/external/poi-xlsx',
  'corpus/external/lo-docx-import',
  'corpus/external/lo-docx-export',
  'corpus/external/poi-docx',
  'corpus/external/lo-pptx',
  'corpus/external/poi-pptx',
];

/** The extensions a bare name may carry. Every reader converts to PDF the same way. */
const CORPUS_EXTS: ReadonlyArray<string> = ['.xlsx', '.docx', '.pptx'];

/** Resolve a bare fixture name, a corpus-relative path, or an absolute one. */
function resolveInput(nameOrPath: string): string {
  const candidates = [
    nameOrPath,
    resolve(root, nameOrPath),
    ...CORPUS_DIRS.flatMap((dir) =>
      CORPUS_EXTS.map((ext) => resolve(root, dir, `${nameOrPath}${ext}`)),
    ),
  ];
  const hit = candidates.find((c) => existsSync(c));
  if (!hit) throw new Error(`not found: ${nameOrPath}`);
  return hit;
}

/** Rasterise every page of a PDF; returns the page PNG paths in order. */
function rasterize(pdf: string, tag: string): Array<string> {
  execFileSync(
    'mutool',
    ['draw', '-r', String(DPI), '-o', resolve(workDir, `${tag}-%d.png`), pdf],
    {
      stdio: 'ignore',
      timeout: 120_000,
    },
  );
  return readdirSync(workDir)
    .filter((f) => f.startsWith(`${tag}-`) && f.endsWith('.png'))
    .sort((a, b) => pageNo(a) - pageNo(b))
    .map((f) => resolve(workDir, f));
}

const pageNo = (f: string): number => Number(/-(\d+)\.png$/.exec(f)?.[1] ?? 0);

/** A label strip, so a page cannot be mistaken for the other side's. */
function label(text: string, width: number): Buffer {
  const svg =
    `<svg width="${width}" height="${LABEL_H}" xmlns="http://www.w3.org/2000/svg">` +
    `<rect width="${width}" height="${LABEL_H}" fill="#1f2933"/>` +
    `<text x="8" y="18" font-family="monospace" font-size="13" fill="#ffffff">${text}</text>` +
    `</svg>`;
  return Buffer.from(svg);
}

async function main(): Promise<void> {
  const input = resolveInput(process.argv[2] ?? '');
  const name = basename(input).replace(/\.(?:xlsx|docx)$/iu, '');
  const only = arg('--pages')
    ?.split(',')
    .map((n) => Number(n.trim()));

  rmSync(workDir, { recursive: true, force: true });
  mkdirSync(workDir, { recursive: true });
  mkdirSync(outDir, { recursive: true });

  const ourPdf = resolve(workDir, 'ours.pdf');
  writeFileSync(
    ourPdf,
    await Ream.parse(new Uint8Array(readFileSync(input)), parseOptions(input)).convert('pdf', {
      // The document's own families (see ./fonts). CJK included: the library
      // fetches a face per SCRIPT now, which is also what pixel-scout measures —
      // a host face substituted here instead showed a Japanese sheet in a Korean
      // one, and read as missing glyphs the score never saw.
      ...corpusFontOptions(),
      fileName: basename(input),
    }),
  );
  const ours = rasterize(ourPdf, 'ours');
  // A PDF is its own reference — the same oracle pixel-scout uses. Asking
  // LibreOffice to "convert" one runs it through Draw, which REDRAWS the page
  // rather than reporting it: Brotli-Prototype-FileA.pdf came back as one page
  // of twenty-five, and 160F-2019.pdf came back without the button it draws.
  // Read against that, the picture and the number disagree about the same file.
  const isPdf = /\.pdf$/iu.test(input);
  const refs = rasterize(isPdf ? input : referenceToPdf(input, workDir), 'ref');

  const count = Math.max(ours.length, refs.length);
  const wanted = only ?? Array.from({ length: count }, (_, i) => i + 1);
  const written: Array<string> = [];

  for (const p of wanted) {
    if (p < 1 || p > count) continue;
    const a = ours[p - 1];
    const b = refs[p - 1];
    // A page present on one side only still gets written — a page-count
    // difference is exactly the kind of thing worth seeing.
    const dims = async (f: string | undefined) =>
      f ? await sharp(f).metadata() : { width: 0, height: 0 };
    const da = await dims(a);
    const db = await dims(b);
    const wa = da.width ?? 0;
    const wb = db.width ?? 0;
    const h = Math.max(da.height ?? 0, db.height ?? 0) + LABEL_H;
    const w = wa + GAP + wb;

    const layers: Array<Record<string, unknown>> = [
      { input: label(`OURS  ${name}  page ${p}/${ours.length}`, wa || 1), top: 0, left: 0 },
      {
        input: label(
          `${isPdf ? 'THE FILE ITSELF' : 'LIBREOFFICE'}  page ${p}/${refs.length}`,
          wb || 1,
        ),
        top: 0,
        left: wa + GAP,
      },
    ];
    if (a) layers.push({ input: a, top: LABEL_H, left: 0 });
    if (b) layers.push({ input: b, top: LABEL_H, left: wa + GAP });

    const out = resolve(outDir, `${name}-p${String(p).padStart(2, '0')}.png`);
    await sharp({
      create: { width: w, height: h, channels: 3, background: { r: 210, g: 214, b: 220 } },
    })
      .composite(layers)
      .toFile(out);
    written.push(out);
  }

  rmSync(workDir, { recursive: true, force: true });
  console.log(`ours ${ours.length} page(s) · libreoffice ${refs.length} page(s)`);
  for (const f of written) console.log(f);
}

await main();
