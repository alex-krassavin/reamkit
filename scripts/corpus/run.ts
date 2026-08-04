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
import { EXPLICIT_FONTS, corpusFontOptions } from './fonts';
import { Ream } from '@/core/converter/ream';

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

// CORPUS_SANDBOX=docker sandboxes the LibreOffice reference render (see lib.ts)
// — the main external risk. CORPUS_ISOLATE_OURS=1 ALSO runs our own parser in a
// child process (wall-clock timeout + heap cap), for genuinely hostile input;
// it's separate because the per-doc child startup is slow over a big batch and
// OpcPackage.open already caps decompression. Default: both off (fast in-process
// path for our trusted fixtures).
const ISOLATE_OURS = process.env.CORPUS_ISOLATE_OURS === '1';
// Fonts: the document's OWN families, substituted the way LibreOffice
// substitutes them, off a disk cache (see ./fonts). CORPUS_FONTS=roboto pins the
// old single-family render for a side-by-side. The child process reads the same
// cache, so isolation no longer costs a download per document.
// CORPUS_PROFILE selects a renderer-compatibility layoutProfile for OUR render
// (E-PARITY): 'libreoffice' should track the LibreOffice golden most closely,
// 'word' targets Word. Unset = the default 'ream' typesetter.
const PROFILE = process.env.CORPUS_PROFILE as 'ream' | 'word' | 'libreoffice' | undefined;
const OUR_TIMEOUT_MS = 60_000;

// Convert with our library, writing the PDF to `outPath`. When isolating, spawn
// a child so a hostile/pathological doc can't hang or OOM the runner.
async function ourConvert(input: string, outPath: string): Promise<void> {
  if (ISOLATE_OURS) {
    execFileSync('npx', ['tsx', resolve(here, 'convert-one.ts'), input, outPath], {
      stdio: 'ignore',
      timeout: OUR_TIMEOUT_MS,
      env: { ...process.env, NODE_OPTIONS: '--max-old-space-size=512' },
    });
    return;
  }
  const bytes = new Uint8Array(readFileSync(input));
  const profileOpt = PROFILE ? { layoutProfile: PROFILE } : {};
  // The Ream facade sniffs the format from the bytes and dispatches to the right
  // reader, so one path covers all seven inputs (docx/xlsx/pptx/pdf + the legacy
  // doc/xls/ppt).
  const pdf = await Ream.parse(bytes).convert('pdf', { ...corpusFontOptions(), ...profileOpt });
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

interface VerdictInput {
  readonly textSimilarity: number;
  readonly geometrySimilarity: number;
  readonly worstVisual: number;
  readonly pageMatch: boolean;
  readonly dimsMatch: boolean;
  readonly ourChars: number;
  readonly refChars: number;
}

/**
 * The per-document verdict.
 *
 * This used to be `textSimilarity > 0.95 && worstVisual < 0.1`, which passed
 * documents no reader would call passing:
 *
 *   - a page-count mismatch was not consulted at all, so a sheet that spilled
 *     onto a second page LibreOffice fits on one was green;
 *   - two empty extractions score a similarity of 1.0, so documents where
 *     BOTH sides rendered no text at all were green — the most alarming
 *     result the harness can produce, reported as its best;
 *   - a page-size mismatch reached the report only as free text in the note
 *     column, where nothing aggregated it;
 *   - geometrySimilarity was computed, printed, and never used.
 *
 * `🈳` (vacuous) is deliberately its own state rather than a failure: an empty
 * document legitimately renders empty on both sides. What it must never be is
 * a pass, because it is evidence of nothing.
 */
function verdict(v: VerdictInput): string {
  if (v.ourChars === 0 && v.refChars === 0) return '🈳';
  if (!v.dimsMatch || !v.pageMatch) return '⚠️';
  if (v.textSimilarity <= 0.95 || v.worstVisual >= 0.1) return '⚠️';
  // Below this the pages hold the same characters in visibly different places.
  if (v.geometrySimilarity < 0.5) return '⚠️';
  return '✅';
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
    const isPdf = /\.pdf$/i.test(name);
    try {
      // 1. Our PDF (isolated in a child process under sandbox mode).
      const ourPdfPath = resolve(workDir, name + '.our.pdf');
      await ourConvert(input, ourPdfPath);

      // 2. Reference PDF. For Office formats this is LibreOffice's golden render;
      //    for a PDF *source* the reference is the original file, so the diff
      //    measures our read→render roundtrip fidelity (LibreOffice's own PDF
      //    import is too lossy to be a golden).
      const refPdfPath = isPdf ? input : referenceToPdf(input, workDir);

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
        status: verdict({
          textSimilarity: sd.textSimilarity,
          geometrySimilarity: sd.geometrySimilarity,
          worstVisual,
          pageMatch,
          dimsMatch: dimNote === '',
          ourChars: sd.ourChars,
          refChars: sd.refChars,
        }),
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
  // Surface problems first: ❌ then ⚠️ then 🈳 then ✅, alphabetical within a group.
  const rank: Record<string, number> = { '❌': 0, '⚠️': 1, '🈳': 2, '✅': 3 };
  const sorted = [...rows].sort(
    (a, b) => (rank[a.status] ?? 9) - (rank[b.status] ?? 9) || a.name.localeCompare(b.name),
  );
  const counts = { ok: 0, warn: 0, vacuous: 0, fail: 0 };
  for (const r of rows) {
    if (r.status === '✅') counts.ok++;
    else if (r.status === '⚠️') counts.warn++;
    else if (r.status === '🈳') counts.vacuous++;
    else counts.fail++;
  }

  const lines: Array<string> = [];
  lines.push(`# Corpus validation report`);
  lines.push('');
  lines.push(
    `Reference: LibreOffice \`soffice\` (PDF sources: the original file — a read→render roundtrip). Our profile: \`${PROFILE ?? 'ream'}\`` +
      `${EXPLICIT_FONTS ? ' (pinned Roboto)' : ' (substituted fonts)'}. Raster DPI: ${dpi}. ` +
      `Visual = worst-page pixel mismatch ratio (lower is better). ` +
      `TextSim = LCS char similarity vs reference (higher is better). ` +
      `Drift = median baseline-y delta.`,
  );
  lines.push('');
  lines.push(
    `**${rows.length} docs — ✅ ${counts.ok} clean · ⚠️ ${counts.warn} divergent · ` +
      `🈳 ${counts.vacuous} vacuous (both sides empty — evidence of nothing) · ❌ ${counts.fail} failed.**`,
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
