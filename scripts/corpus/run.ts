// Corpus validation harness.
//
// For each document in corpus/inputs/:
//   1. our PDF      = convertDocxToPdfSync / convertXlsxToPdfSync
//   2. reference    = soffice --convert-to pdf  (LibreOffice "gold standard")
//   3. rasterise both (mutool, RGB PPM) and extract structured text (stext)
//   4. report structural diff (text similarity + baseline drift) and visual
//      diff (pixel mismatch ratio) per document.
//
// Output: a markdown regression table to stdout and corpus/report.md.
//
// Usage: npx tsx scripts/corpus/run.ts [--dpi 100] [--keep]

import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { basename, dirname, resolve } from 'node:path';

import {
  listCorpus,
  parsePpm,
  rasterize,
  referenceToPdf,
  stext,
  structuralDiff,
  visualDiff,
} from './lib';
import type { FontBytesByVariant } from '@/core/font';
import { convertDocxToPdf, convertDocxToPdfSync, convertXlsxToPdfSync } from '@/core/converter';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '../..');
// Which document set to validate. Default = our trusted synthetic fixtures;
// point CORPUS_DIR at corpus/external/* for fetched real documents (and set
// CORPUS_SANDBOX=docker for those — they're untrusted).
const corpusDir = process.env.CORPUS_DIR
  ? resolve(root, process.env.CORPUS_DIR)
  : resolve(root, 'corpus/inputs');
const workDir = resolve(root, 'corpus/.work');

const dpiArg = process.argv.indexOf('--dpi');
const DPI = dpiArg >= 0 ? Number(process.argv[dpiArg + 1]) : 100;
const KEEP = process.argv.includes('--keep');

const FONTS: FontBytesByVariant = {
  regular: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Bold.ttf'))),
  italic: new Uint8Array(readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-Italic.ttf'))),
  boldItalic: new Uint8Array(
    readFileSync(resolve(root, 'tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
  ),
};

// CORPUS_SANDBOX=docker sandboxes the LibreOffice reference render (see lib.ts)
// — the main external risk. CORPUS_ISOLATE_OURS=1 ALSO runs our own parser in a
// child process (wall-clock timeout + heap cap), for genuinely hostile input;
// it's separate because the per-doc child startup is slow over a big batch and
// OpcPackage.open already caps decompression. Default: both off (fast in-process
// path for our trusted fixtures).
const ISOLATE_OURS = process.env.CORPUS_ISOLATE_OURS === '1';
// CORPUS_AUTOFONT=1 renders docx with the real async font auto-substitution
// (sans→Roboto / serif→Tinos / mono→Cousine by the document's declared family)
// instead of a fixed Roboto set — so the visual metric reflects layout fidelity
// rather than a font mismatch. In-process only (the per-URL font cache must
// persist across docs); ignored under CORPUS_ISOLATE_OURS.
const AUTOFONT = process.env.CORPUS_AUTOFONT === '1';
const OUR_TIMEOUT_MS = 60_000;

// Convert with our library, writing the PDF to `outPath`. When isolating, spawn
// a child so a hostile/pathological doc can't hang or OOM the runner.
async function ourConvert(input: string, isXlsx: boolean, outPath: string): Promise<void> {
  if (ISOLATE_OURS) {
    execFileSync('npx', ['tsx', resolve(here, 'convert-one.ts'), input, outPath], {
      stdio: 'ignore',
      timeout: OUR_TIMEOUT_MS,
      env: { ...process.env, NODE_OPTIONS: '--max-old-space-size=512' },
    });
    return;
  }
  const bytes = new Uint8Array(readFileSync(input));
  const pdf =
    AUTOFONT && !isXlsx
      ? await convertDocxToPdf(bytes)
      : isXlsx
        ? convertXlsxToPdfSync(bytes, { fonts: FONTS })
        : convertDocxToPdfSync(bytes, { fonts: FONTS });
  writeFileSync(outPath, pdf);
}

interface Row {
  readonly name: string;
  readonly status: string;
  readonly pages: string;
  readonly textSim: string;
  readonly geom: string;
  readonly drift: string;
  readonly visual: string;
  readonly note: string;
}

function pct(n: number): string {
  return (n * 100).toFixed(1) + '%';
}

async function main(): Promise<void> {
  // Start from a clean work dir so stale rasters from a prior --keep run can't
  // contaminate this one.
  rmSync(workDir, { recursive: true, force: true });
  mkdirSync(workDir, { recursive: true });
  const inputs = listCorpus(corpusDir);
  if (inputs.length === 0) {
    console.error('No corpus inputs. Run: npx tsx scripts/corpus/build-corpus.ts');
    process.exit(1);
  }
  console.error(`Validating ${inputs.length} document(s) against LibreOffice...\n`);

  const rows: Array<Row> = [];

  for (const input of inputs) {
    const name = basename(input);
    const isXlsx = name.endsWith('.xlsx') || name.endsWith('.xlsm');
    try {
      // 1. Our PDF (isolated in a child process under sandbox mode).
      const ourPdfPath = resolve(workDir, name + '.our.pdf');
      await ourConvert(input, isXlsx, ourPdfPath);

      // 2. Reference PDF (sandboxed LibreOffice under sandbox mode).
      const refPdfPath = referenceToPdf(input, workDir);

      // 3. Rasterise + stext.
      const ourPpms = rasterize(ourPdfPath, resolve(workDir, name + '.our-%d.ppm'), DPI);
      const refPpms = rasterize(refPdfPath, resolve(workDir, name + '.ref-%d.ppm'), DPI);
      const ourStext = stext(ourPdfPath, resolve(workDir, name + '.our.xml'));
      const refStext = stext(refPdfPath, resolve(workDir, name + '.ref.xml'));

      // 4. Diffs.
      const sd = structuralDiff(ourStext, refStext);
      const pageMatch = ourPpms.length === refPpms.length;
      let worstVisual = 0;
      let dimNote = '';
      const pageCount = Math.min(ourPpms.length, refPpms.length);
      for (let i = 0; i < pageCount; i++) {
        const our = parsePpm(new Uint8Array(readFileSync(ourPpms[i]!)));
        const ref = parsePpm(new Uint8Array(readFileSync(refPpms[i]!)));
        const vd = visualDiff(our, ref);
        if (vd.mismatchRatio > worstVisual) worstVisual = vd.mismatchRatio;
        if (!vd.dimsMatch && !dimNote) dimNote = `dims ${vd.ourDims} vs ${vd.refDims}`;
      }

      rows.push({
        name,
        status: sd.textSimilarity > 0.95 && worstVisual < 0.1 ? '✅' : '⚠️',
        pages: pageMatch ? String(ourPpms.length) : `${ourPpms.length}≠${refPpms.length}`,
        textSim: pct(sd.textSimilarity),
        geom: pct(sd.geometrySimilarity),
        drift: sd.medianBaselineDriftPt.toFixed(1) + 'pt',
        visual: pct(worstVisual),
        note: dimNote || `${sd.ourChars}/${sd.refChars} chars`,
      });
      console.error(`✓ ${name}`);
    } catch (err) {
      rows.push({
        name,
        status: '❌',
        pages: '-',
        textSim: '-',
        geom: '-',
        drift: '-',
        visual: '-',
        note: (err as Error).message.slice(0, 60),
      });
      console.error(`✗ ${name}: ${(err as Error).message}`);
    }
  }

  const report = renderReport(rows, DPI);
  writeFileSync(resolve(root, 'corpus/report.md'), report);
  console.log('\n' + report);

  if (!KEEP) rmSync(workDir, { recursive: true, force: true });
}

function renderReport(rows: Array<Row>, dpi: number): string {
  // Surface problems first: ❌ then ⚠️ then ✅, alphabetical within a group.
  const rank: Record<string, number> = { '❌': 0, '⚠️': 1, '✅': 2 };
  const sorted = [...rows].sort(
    (a, b) => (rank[a.status] ?? 9) - (rank[b.status] ?? 9) || a.name.localeCompare(b.name),
  );
  const counts = { ok: 0, warn: 0, fail: 0 };
  for (const r of rows) {
    if (r.status === '✅') counts.ok++;
    else if (r.status === '⚠️') counts.warn++;
    else counts.fail++;
  }

  const lines: Array<string> = [];
  lines.push(`# Corpus validation report`);
  lines.push('');
  lines.push(
    `Reference: LibreOffice \`soffice\`. Raster DPI: ${dpi}. ` +
      `Visual = worst-page pixel mismatch ratio (lower is better). ` +
      `TextSim = LCS char similarity vs reference (higher is better). ` +
      `Drift = median baseline-y delta.`,
  );
  lines.push('');
  lines.push(
    `**${rows.length} docs — ✅ ${counts.ok} clean · ⚠️ ${counts.warn} divergent · ❌ ${counts.fail} failed.**`,
  );
  lines.push('');
  lines.push('| Doc | St | Pages | TextSim | Geom | Drift | Visual | Note |');
  lines.push('|---|---|---|---|---|---|---|---|');
  for (const r of sorted) {
    lines.push(
      `| ${r.name} | ${r.status} | ${r.pages} | ${r.textSim} | ${r.geom} | ${r.drift} | ${r.visual} | ${r.note} |`,
    );
  }
  return lines.join('\n');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
