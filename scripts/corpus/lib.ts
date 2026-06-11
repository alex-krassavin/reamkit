// Corpus-validation primitives: rasterise + structured-text extraction via
// mutool, reference rendering via soffice, and diff metrics. No external npm
// deps — we parse mutool's P6 PPM and stext XML ourselves.

import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { basename, dirname, resolve } from 'node:path';

// ---- external tools ----

export interface Ppm {
  readonly width: number;
  readonly height: number;
  readonly rgb: Uint8Array; // width*height*3
}

// Parse a binary P6 PPM (header "P6\n<w> <h>\n<max>\n" then RGB bytes).
export function parsePpm(bytes: Uint8Array): Ppm {
  let pos = 0;
  const token = (): string => {
    // Skip whitespace and comments.
    while (pos < bytes.length) {
      const c = bytes[pos]!;
      if (c === 0x23) {
        while (pos < bytes.length && bytes[pos] !== 0x0a) pos++;
      } else if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) {
        pos++;
      } else break;
    }
    let s = '';
    while (pos < bytes.length) {
      const c = bytes[pos]!;
      if (c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d) break;
      s += String.fromCharCode(c);
      pos++;
    }
    return s;
  };
  const magic = token();
  if (magic !== 'P6') throw new Error(`Expected P6 PPM, got ${magic}`);
  const width = Number(token());
  const height = Number(token());
  Number(token()); // maxval
  pos++; // single whitespace after maxval
  const rgb = bytes.subarray(pos, pos + width * height * 3);
  return { width, height, rgb };
}

export interface CharBox {
  readonly c: string;
  readonly x: number;
  readonly y: number;
}

export interface StextPage {
  readonly width: number;
  readonly height: number;
  readonly chars: Array<CharBox>;
}

// Extract per-character positions from mutool stext XML. We avoid a full XML
// parse: stext is line-oriented and regular, so targeted regexes are robust.
export function parseStext(xml: string): Array<StextPage> {
  const pages: Array<StextPage> = [];
  const pageRe = /<page id="[^"]*" width="([\d.]+)" height="([\d.]+)">([\s\S]*?)<\/page>/g;
  let pm: RegExpExecArray | null;
  while ((pm = pageRe.exec(xml)) !== null) {
    const width = Number(pm[1]);
    const height = Number(pm[2]);
    const body = pm[3]!;
    const chars: Array<CharBox> = [];
    const charRe = /<char [^>]*x="([\d.-]+)" y="([\d.-]+)"[^>]* c="([^"]*)"/g;
    let cm: RegExpExecArray | null;
    while ((cm = charRe.exec(body)) !== null) {
      chars.push({ x: Number(cm[1]), y: Number(cm[2]), c: decodeEntity(cm[3]!) });
    }
    pages.push({ width, height, chars });
  }
  return pages;
}

function decodeEntity(s: string): string {
  return s
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)));
}

const SOFFICE_TIMEOUT_MS = 180_000;
const MUTOOL_TIMEOUT_MS = 60_000;

const SANDBOX_IMAGE = process.env.CORPUS_SANDBOX_IMAGE ?? 'docgen-losandbox:latest';

// Produce the reference ("golden") PDF for an input document. Routes to the
// Docker sandbox when CORPUS_SANDBOX=docker — use that for UNTRUSTED inputs
// (e.g. real-world / GovDocs1 documents); the plain path runs the host's
// LibreOffice and is only for inputs you trust (our own synthetic fixtures).
export function referenceToPdf(input: string, outDir: string): string {
  return process.env.CORPUS_SANDBOX === 'docker'
    ? sofficeToPdfSandboxed(input, outDir)
    : sofficeToPdf(input, outDir);
}

export function sofficeToPdf(input: string, outDir: string): string {
  // LibreOffice writes <basename>.pdf into outDir.
  execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', outDir, input], {
    stdio: 'ignore',
    timeout: SOFFICE_TIMEOUT_MS,
  });
  return expectPdf(input, outDir);
}

// Convert via LibreOffice inside a locked-down Docker container: no network,
// all capabilities dropped, no-new-privileges, read-only rootfs (profile + tmp
// on tmpfs), and CPU/memory/PID limits. A hostile document can neither escape
// the container nor exhaust the host. The input dir is mounted read-only.
export function sofficeToPdfSandboxed(input: string, outDir: string): string {
  const inDir = dirname(resolve(input));
  const base = basename(input);
  execFileSync(
    'docker',
    [
      'run',
      '--rm',
      '--network',
      'none',
      '--memory',
      '1g',
      '--cpus',
      '1',
      '--pids-limit',
      '256',
      '--cap-drop',
      'ALL',
      '--security-opt',
      'no-new-privileges',
      '--read-only',
      '--tmpfs',
      '/tmp:size=512m',
      // HOME on tmpfs, owned by the sandbox user (uid 1000) so LibreOffice can
      // write its dconf/fontconfig caches under a read-only rootfs.
      '--tmpfs',
      '/home/sandbox:size=64m,uid=1000,gid=1000',
      '-v',
      `${inDir}:/in:ro`,
      '-v',
      `${resolve(outDir)}:/out`,
      SANDBOX_IMAGE,
      `/in/${base}`,
    ],
    { stdio: 'ignore', timeout: SOFFICE_TIMEOUT_MS },
  );
  return expectPdf(input, outDir);
}

function expectPdf(input: string, outDir: string): string {
  const base = basename(input).replace(/\.[^.]+$/, '');
  const out = resolve(outDir, `${base}.pdf`);
  if (!existsSync(out)) {
    throw new Error('soffice produced no PDF (conversion rejected the file)');
  }
  return out;
}

export function rasterize(pdf: string, outPattern: string, dpi: number): Array<string> {
  execFileSync(
    'mutool',
    ['draw', '-c', 'rgb', '-F', 'pnm', '-o', outPattern, '-r', String(dpi), pdf],
    {
      stdio: 'ignore',
      timeout: MUTOOL_TIMEOUT_MS,
    },
  );
  // outPattern uses %d; collect the produced files.
  const dir = outPattern.substring(0, outPattern.lastIndexOf('/'));
  const prefix = outPattern.substring(outPattern.lastIndexOf('/') + 1).split('%d')[0]!;
  return readdirSync(dir)
    .filter((f) => f.startsWith(prefix) && f.endsWith('.ppm'))
    .sort((a, b) => pageNum(a) - pageNum(b))
    .map((f) => resolve(dir, f));
}

function pageNum(f: string): number {
  const m = f.match(/(\d+)\.ppm$/);
  return m ? Number(m[1]) : 0;
}

export function stext(pdf: string, outFile: string): Array<StextPage> {
  execFileSync('mutool', ['draw', '-F', 'stext', '-o', outFile, pdf], {
    stdio: 'ignore',
    timeout: MUTOOL_TIMEOUT_MS,
  });
  return parseStext(readFileSync(outFile, 'latin1'));
}

// ---- diff metrics ----

export interface VisualDiff {
  readonly dimsMatch: boolean;
  readonly ourDims: string;
  readonly refDims: string;
  readonly mismatchRatio: number; // fraction of pixels differing beyond threshold
}

// Compare two PPMs. Pixels differing by more than `tol` (per channel, 0..255)
// count as mismatches. Returns the mismatch ratio over the overlap region.
export function visualDiff(our: Ppm, ref: Ppm, tol = 24): VisualDiff {
  const dimsMatch = our.width === ref.width && our.height === ref.height;
  const w = Math.min(our.width, ref.width);
  const h = Math.min(our.height, ref.height);
  let mismatches = 0;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const oi = (y * our.width + x) * 3;
      const ri = (y * ref.width + x) * 3;
      const d =
        Math.abs(our.rgb[oi]! - ref.rgb[ri]!) +
        Math.abs(our.rgb[oi + 1]! - ref.rgb[ri + 1]!) +
        Math.abs(our.rgb[oi + 2]! - ref.rgb[ri + 2]!);
      if (d > tol * 3) mismatches++;
    }
  }
  return {
    dimsMatch,
    ourDims: `${our.width}x${our.height}`,
    refDims: `${ref.width}x${ref.height}`,
    mismatchRatio: w * h > 0 ? mismatches / (w * h) : 1,
  };
}

export interface StructuralDiff {
  readonly ourChars: number;
  readonly refChars: number;
  readonly textSimilarity: number; // 0..1 over normalised text
  readonly medianBaselineDriftPt: number; // median |Δy| of matched leading chars
  // Font-agnostic geometry: of the words whose TEXT matches in reading order,
  // the share whose position agrees within GEOM_TOL_PT on both axes. Catches
  // layout faithfulness even when a substitute font changes every advance.
  readonly geometrySimilarity: number; // 0..1, 1 when nothing matched
  readonly matchedWords: number;
}

// Normalise text for content comparison: collapse whitespace, drop it entirely
// (so line-break differences don't penalise content equality).
function normalize(chars: Array<CharBox>): string {
  return chars
    .map((c) => c.c)
    .join('')
    .replace(/\s+/g, '');
}

// Longest-common-subsequence ratio over two strings, capped for performance.
function lcsRatio(a: string, b: string): number {
  const n = Math.min(a.length, 4000);
  const m = Math.min(b.length, 4000);
  if (n === 0 && m === 0) return 1;
  if (n === 0 || m === 0) return 0;
  const dp = new Uint16Array((n + 1) * (m + 1));
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      const idx = i * (m + 1) + j;
      if (a[i - 1] === b[j - 1]) dp[idx] = dp[(i - 1) * (m + 1) + (j - 1)]! + 1;
      else dp[idx] = Math.max(dp[(i - 1) * (m + 1) + j]!, dp[i * (m + 1) + (j - 1)]!);
    }
  }
  const lcs = dp[n * (m + 1) + m]!;
  return (2 * lcs) / (n + m);
}

export function structuralDiff(
  ourPages: Array<StextPage>,
  refPages: Array<StextPage>,
): StructuralDiff {
  const ourChars = ourPages.reduce((s, p) => s + p.chars.length, 0);
  const refChars = refPages.reduce((s, p) => s + p.chars.length, 0);
  const ourText = ourPages.map((p) => normalize(p.chars)).join('');
  const refText = refPages.map((p) => normalize(p.chars)).join('');
  const textSimilarity = lcsRatio(ourText, refText);

  // Baseline drift: match the first N distinct y-positions (line baselines)
  // across pages and compare. Robust to font-substitution x-advance changes.
  const ourYs = leadingBaselines(ourPages);
  const refYs = leadingBaselines(refPages);
  const k = Math.min(ourYs.length, refYs.length);
  const drifts: Array<number> = [];
  for (let i = 0; i < k; i++) drifts.push(Math.abs(ourYs[i]! - refYs[i]!));
  drifts.sort((a, b) => a - b);
  const medianBaselineDriftPt = drifts.length > 0 ? drifts[Math.floor(drifts.length / 2)]! : 0;

  const geom = geometrySimilarity(ourPages, refPages);

  return {
    ourChars,
    refChars,
    textSimilarity,
    medianBaselineDriftPt,
    geometrySimilarity: geom.similarity,
    matchedWords: geom.matched,
  };
}

const GEOM_TOL_PT = 6;

interface WordBox {
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly page: number;
}

// Group character boxes into words: a word breaks on whitespace or a large
// horizontal jump (column gaps, tabs).
function wordBoxes(pages: Array<StextPage>): Array<WordBox> {
  const out: Array<WordBox> = [];
  pages.forEach((p, pageIdx) => {
    let text = '';
    let x = 0;
    let y = 0;
    let lastX = 0;
    const flush = () => {
      if (text.length > 0) out.push({ text, x, y, page: pageIdx });
      text = '';
    };
    for (const c of p.chars) {
      if (/\s/.test(c.c)) {
        flush();
        continue;
      }
      const jump = text.length > 0 && (Math.abs(c.y - y) > 2 || c.x - lastX > 18);
      if (jump) flush();
      if (text.length === 0) {
        x = c.x;
        y = c.y;
      }
      text += c.c;
      lastX = c.x;
    }
    flush();
  });
  return out;
}

// Greedy in-order matching of equal word texts (a windowed LCS stand-in),
// then the share of matches whose positions agree within tolerance.
function geometrySimilarity(
  ourPages: Array<StextPage>,
  refPages: Array<StextPage>,
): { similarity: number; matched: number } {
  const ours = wordBoxes(ourPages);
  const refs = wordBoxes(refPages);
  const WINDOW = 40;
  let i = 0;
  let matched = 0;
  let close = 0;
  for (const ref of refs) {
    const limit = Math.min(ours.length, i + WINDOW);
    for (let j = i; j < limit; j++) {
      if (ours[j]!.text !== ref.text) continue;
      matched++;
      if (
        ours[j]!.page === ref.page &&
        Math.abs(ours[j]!.x - ref.x) <= GEOM_TOL_PT &&
        Math.abs(ours[j]!.y - ref.y) <= GEOM_TOL_PT
      ) {
        close++;
      }
      i = j + 1;
      break;
    }
  }
  return { similarity: matched > 0 ? close / matched : 1, matched };
}

// Distinct baseline y-positions in reading order (one per text line).
function leadingBaselines(pages: Array<StextPage>): Array<number> {
  const ys: Array<number> = [];
  for (const p of pages) {
    let lastY = -1e9;
    for (const c of p.chars) {
      if (Math.abs(c.y - lastY) > 1) {
        ys.push(c.y);
        lastY = c.y;
      }
    }
  }
  return ys;
}

export function listCorpus(dir: string): Array<string> {
  return readdirSync(dir)
    .filter((f) => /\.(docx|docm|xlsx|xlsm)$/i.test(f) && !f.startsWith('~$'))
    .sort()
    .map((f) => resolve(dir, f));
}
