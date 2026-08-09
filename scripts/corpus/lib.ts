// Corpus-validation primitives: rasterise + structured-text extraction via
// mutool, reference rendering via soffice, and diff metrics. No external npm
// deps — we parse mutool's P6 PPM and stext XML ourselves.

import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

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
  //
  // Its user profile is the reason for the -env flag. One profile is one
  // running instance: a second `soffice` started while the first is converting
  // does not convert, it hands the job to the instance that owns the profile
  // and exits — and the caller finds no PDF and reports a document that
  // "threw". A per-process profile lets the scout sweep in the background while
  // a visual diff runs in the foreground, which is exactly how both get used.
  const profile = pathToFileURL(resolve(tmpdir(), `ream-lo-${String(process.pid)}`)).href;
  execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${profile}`,
      '--headless',
      '--convert-to',
      'pdf',
      '--outdir',
      outDir,
      input,
    ],
    {
      stdio: 'ignore',
      timeout: SOFFICE_TIMEOUT_MS,
    },
  );
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
//
// `slack` is how far a pixel may LOOK for its match, in pixels. Two renderers
// never put a glyph on the same pixel — our font is not the workbook's and the
// hinting differs — so a page of text scores 16 % against a reference it is
// character-for-character identical to, and the number says nothing about
// whether the page is right. Searching a small neighbourhood before counting a
// mismatch throws that away and keeps what matters: a missing block, a wrong
// fill, a shifted column. Zero (the default) compares pixel against pixel.
export function visualDiff(our: Ppm, ref: Ppm, tol = 24, slack = 0): VisualDiff {
  const dimsMatch = our.width === ref.width && our.height === ref.height;
  const w = Math.min(our.width, ref.width);
  const h = Math.min(our.height, ref.height);
  const at = (p: Ppm, x: number, y: number): number => (y * p.width + x) * 3;
  const diffAt = (oi: number, ri: number): number =>
    Math.abs(our.rgb[oi]! - ref.rgb[ri]!) +
    Math.abs(our.rgb[oi + 1]! - ref.rgb[ri + 1]!) +
    Math.abs(our.rgb[oi + 2]! - ref.rgb[ri + 2]!);
  let mismatches = 0;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const oi = at(our, x, y);
      if (diffAt(oi, at(ref, x, y)) <= tol * 3) continue;
      let matched = false;
      for (let dy = -slack; dy <= slack && !matched; dy++) {
        for (let dx = -slack; dx <= slack && !matched; dx++) {
          const nx = x + dx;
          const ny = y + dy;
          if (nx < 0 || ny < 0 || nx >= w || ny >= h) continue;
          if (diffAt(oi, at(ref, nx, ny)) <= tol * 3) matched = true;
        }
      }
      if (!matched) mismatches++;
    }
  }
  return {
    dimsMatch,
    ourDims: `${our.width}x${our.height}`,
    refDims: `${ref.width}x${ref.height}`,
    mismatchRatio: w * h > 0 ? mismatches / (w * h) : 1,
  };
}

/**
 * How differently the two pages are COLOURED, ignoring where the colour sits:
 * the L1 distance between their quantised colour histograms, over the pixel
 * count. 0 is the same palette in the same proportions.
 *
 * The pixel comparison above answers "is this the same page" and is dominated,
 * on any page with text, by the fact that our font is not the workbook's — a
 * character-for-character identical page scores several percent because every
 * glyph lands a hair to one side. This answers a narrower question that survives
 * that: is anything painted a colour the reference does not paint, or in a
 * quantity it does not paint it. A fill resolved from the wrong palette, a chart
 * that did not draw, an image dropped for an unsupported PNG type, a bar chart
 * with the wrong bars — every one of those moves this, and a page of shifted
 * text does not.
 *
 * Colours that land in NEIGHBOURING buckets are treated as the same colour.
 * Two rasterisations of the same flat fill differ by a unit or so — ours wrote
 * `0 0.615686 0.941176 rg` and came back (0,156,239) against the reference's
 * (0,157,240) — and a hard bucket edge between them scored a slide painted
 * exactly right as a total mismatch, which put half a dozen perfect pages at
 * the top of the ranking. So surplus in one bucket cancels against a deficit
 * beside it before anything is counted: a colour that moved less than a
 * quantisation step is the rasteriser rounding, not a difference.
 *
 * @param our  Our raster.
 * @param ref  The reference raster.
 * @param bits Bits kept per channel (5 ⇒ a 32 768-bucket histogram, one step
 * being 8 of 255 — the neighbourhood that cancels is one step wide).
 * @returns 0…1, the share of pixels whose colour has no counterpart.
 */
export function colorDiff(our: Ppm, ref: Ppm, bits = 5): number {
  const shift = 8 - bits;
  const buckets = 1 << (bits * 3);
  const hist = (p: Ppm): { counts: Float64Array; total: number } => {
    const counts = new Float64Array(buckets);
    const total = p.width * p.height;
    for (let i = 0; i < total; i++) {
      const o = i * 3;
      const key =
        ((p.rgb[o]! >> shift) << (bits * 2)) |
        ((p.rgb[o + 1]! >> shift) << bits) |
        (p.rgb[o + 2]! >> shift);
      counts[key]! += 1;
    }
    return { counts, total };
  };
  const a = hist(our);
  const b = hist(ref);
  if (a.total === 0 || b.total === 0) return 1;
  // Normalised to shares, so two pages of different pixel dimensions still
  // compare (a page is a page whatever the paper).
  const diff = new Float64Array(buckets);
  for (let i = 0; i < buckets; i++) diff[i] = a.counts[i]! / a.total - b.counts[i]! / b.total;

  // One hop of transport: a surplus pays off the deficits touching it in the
  // colour cube. What is left has moved further than the grid can blame.
  const side = 1 << bits;
  const at = (r: number, g: number, bl: number): number => (r << (bits * 2)) | (g << bits) | bl;
  for (let r = 0; r < side; r++) {
    for (let g = 0; g < side; g++) {
      for (let bl = 0; bl < side; bl++) {
        const i = at(r, g, bl);
        if (diff[i]! <= 0) continue;
        for (let dr = -1; dr <= 1 && diff[i]! > 0; dr++) {
          for (let dg = -1; dg <= 1 && diff[i]! > 0; dg++) {
            for (let db = -1; db <= 1 && diff[i]! > 0; db++) {
              const [nr, ng, nb] = [r + dr, g + dg, bl + db];
              if (nr < 0 || ng < 0 || nb < 0 || nr >= side || ng >= side || nb >= side) continue;
              const j = at(nr, ng, nb);
              if (diff[j]! >= 0) continue;
              const moved = Math.min(diff[i]!, -diff[j]!);
              diff[i]! -= moved;
              diff[j]! += moved;
            }
          }
        }
      }
    }
  }
  let sum = 0;
  for (let i = 0; i < buckets; i++) sum += Math.abs(diff[i]!);
  return sum / 2;
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
    .filter(
      (f) => /\.(docx|docm|xlsx|xlsm|pptx|pptm|doc|xls|ppt|pdf)$/i.test(f) && !f.startsWith('~$'),
    )
    .sort()
    .map((f) => resolve(dir, f));
}

/**
 * How the corpus tools read a source before re-rendering it.
 *
 * The measurement is a PAGE against a page: our render beside the reference's,
 * pixel for pixel. So a PDF is read as the page it is — every line where its
 * glyphs stand, beside the rules and fills the page paints — and not as the
 * re-flowable document the default reading recovers for a docx or a markdown
 * conversion. Reading a form as prose and then measuring the picture asked the
 * scores to answer for a difference nobody was trying to remove.
 *
 * @param input Path to the source file.
 * @returns The parse options for it (empty for everything that is not a PDF).
 */
export function parseOptions(input: string): { pdfLayout?: 'flow' | 'positional' } {
  return /\.pdf$/i.test(input) ? { pdfLayout: 'positional' } : {};
}
