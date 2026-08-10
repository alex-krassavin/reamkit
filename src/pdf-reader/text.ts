// E-PDF EP2/EP8/EP13 — page text extraction. Interprets the page content (and,
// recursively, the Form XObjects it paints — EP13, whose text would otherwise be
// missed) into positioned runs, then tags any run whose origin falls inside a
// /Link annotation's /Rect with that link's URI (EP8) so hyperlinks survive.

import { IDENTITY, interpretContent, multiply } from './content';
import { buildContentFont } from './font';
import { collectPageAppearances } from './annots';
import { textMarkupOf } from './annot-draw';
import { patternTint } from './pattern-tint';
import type { Quad, TextMarkupAnnot } from './annot-draw';
import type { ContentFont, Matrix, TextRun } from './content';
import type { PdfDict } from '@/pdf/objects';
import type { PdfFile, PdfPage, Rectangle } from './document';

import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

const MAX_FORM_DEPTH = 8;

/**
 * Extract a page's text as positioned {@link TextRun}s (E-PDF EP2/EP8/EP13).
 * Interprets the page content stream and, recursively, the Form XObjects it
 * paints (EP13 — their text would otherwise be missed), then tags any run whose
 * origin falls inside a `/Link` annotation's `/Rect` with that link's URI (EP8)
 * so hyperlinks survive.
 *
 * @param file The owning {@link PdfFile}.
 * @param page The page to extract.
 * @returns The page's runs, each carrying an `href` when it sits under a link.
 */
export function extractPageText(file: PdfFile, page: PdfPage): Array<TextRun> {
  const runs: Array<TextRun> = [];
  collectRuns(file, page.resources, file.pageContent(page), IDENTITY, 0, new Set(), runs);
  // §12.5.5 — an annotation draws in its own appearance stream, and the words
  // it draws are the page's words too: a field's value, a button's caption.
  // Only its ARTWORK was being lifted, so 160F-2019.pdf's reset button arrived
  // as a tinted rectangle with nothing written on it.
  for (const appearance of collectPageAppearances(file, page)) {
    collectRuns(
      file,
      appearance.resources ?? page.resources,
      file.streamData(appearance.stream),
      appearance.ctm,
      1,
      new Set([appearance.stream]),
      runs,
    );
  }
  const links = collectLinks(file, page);
  const marks = collectTextMarkup(file, page);
  if (links.length === 0 && marks.length === 0) return runs;
  return runs.map((run) => {
    const link = links.find((l) => inRect(run.x, run.y, l.rect));
    const marked = marks.find((m) => m.quads.some((q) => covers(q, run)));
    return {
      ...run,
      ...(link ? { href: link.href } : {}),
      ...(marked ? { markup: marked.mark } : {}),
    };
  });
}

/**
 * §12.5.6.10 — whether a marked quad covers this run.
 *
 * The quad is a box round a run of text and the run is a baseline with an
 * advance, so the test is the baseline falling inside the box's height while
 * the two overlap horizontally by more than a hair. A quad drawn round one word
 * of a line must not claim the whole line.
 */
function covers(q: Quad, run: TextRun): boolean {
  if (run.y < q.y0 || run.y > q.y0 + q.h) return false;
  const left = Math.min(run.x, run.endX);
  const right = Math.max(run.x, run.endX);
  const shared = Math.min(right, q.x1) - Math.max(left, q.x0);
  return shared > Math.min(2, (right - left) / 2);
}

/** Every text-markup annotation on the page, with the boxes it marks. */
function collectTextMarkup(file: PdfFile, page: PdfPage): Array<TextMarkupAnnot> {
  const annots = file.get(page.dict, 'Annots');
  if (!Array.isArray(annots)) return [];
  const out: Array<TextMarkupAnnot> = [];
  for (const entry of annots) {
    const annot = file.resolve(entry);
    if (!(annot instanceof Map)) continue;
    const mark = textMarkupOf(file, annot);
    if (mark) out.push(mark);
  }
  return out;
}

// Interpret one content stream (a page or a Form XObject) into runs, then recurse
// into the Form XObjects it paints — each composing its /Matrix onto the
// placement CTM and using its own /Resources fonts.
function collectRuns(
  file: PdfFile,
  resources: PdfDict | undefined,
  content: Uint8Array,
  baseCtm: Matrix,
  depth: number,
  visiting: Set<PdfStream>,
  out: Array<TextRun>,
): void {
  const result = interpretContent(content, buildFonts(file, resources), baseCtm);
  out.push(...result.texts.map((r) => withPatternColour(file, resources, r, visiting)));
  if (depth >= MAX_FORM_DEPTH) return;
  // §9.6.5 — a Type 3 glyph's procedure may show text of its own, and it is
  // text the page shows. ContentStreamCycleType3insideType3.pdf sets a word
  // inside the glyph of another word.
  for (const glyph of result.glyphs) {
    if (visiting.has(glyph.stream)) continue;
    visiting.add(glyph.stream);
    collectRuns(
      file,
      glyph.resources ?? resources,
      file.streamData(glyph.stream),
      glyph.ctm,
      depth + 1,
      visiting,
      out,
    );
    visiting.delete(glyph.stream);
  }
  if (!resources) return;
  const xobjects = file.get(resources, 'XObject');
  if (!(xobjects instanceof Map)) return;
  for (const placement of result.images) {
    const stream = file.resolve(xobjects.get(placement.name) ?? PDF_NULL);
    if (!(stream instanceof PdfStream) || visiting.has(stream)) continue;
    const sub = file.get(stream.dict, 'Subtype');
    if (!(sub instanceof PdfName) || sub.value !== 'Form') continue;
    visiting.add(stream);
    const formRes = file.get(stream.dict, 'Resources');
    collectRuns(
      file,
      formRes instanceof Map ? formRes : resources,
      file.streamData(stream),
      multiply(matrixOf(file, stream.dict), placement.ctm),
      depth + 1,
      visiting,
      out,
    );
    visiting.delete(stream);
  }
}

/**
 * §8.6.6.2 — a run whose glyphs the page fills with a tiling PATTERN, given the
 * pattern's own colour.
 *
 * A pattern is a content stream, not a colour, and type cannot be filled with
 * one here: what a run carries is a single colour. Painting it in whatever was
 * set before the pattern was is simply wrong — it comes out black where the
 * page shows magenta. The pattern's first mark says what colour the page meant,
 * and the glyphs take that. A documented approximation: the shape of the
 * pattern is lost, its colour is not.
 *
 * @param file      The owning file.
 * @param resources The resources the run was drawn with.
 * @param run       The run, which may name a fill pattern.
 * @param visiting  The streams already being interpreted, against a cycle.
 * @returns The run, its colour taken from the pattern where there is one.
 */
function withPatternColour(
  file: PdfFile,
  resources: PdfDict | undefined,
  run: TextRun,
  visiting: Set<PdfStream>,
): TextRun {
  const name = run.fillPatternName;
  if (name === undefined || !resources) return run;
  const patterns = file.get(resources, 'Pattern');
  if (!(patterns instanceof Map)) return run;
  const stream = file.resolve(patterns.get(name) ?? PDF_NULL);
  if (!(stream instanceof PdfStream) || visiting.has(stream)) return run;
  visiting.add(stream);
  try {
    const tint = patternTint(file, resources, name);
    // The pattern's colour at the pattern's own strength: a tile that covers a
    // third of its cell reads as a third-strength tint, and painting it solid
    // is as wrong in the other direction as painting it black was.
    return tint ? { ...run, colorHex: tintedHex(tint.colorHex, tint.coverage) } : run;
  } catch {
    return run; // A pattern the reader cannot run says nothing about colour.
  } finally {
    visiting.delete(stream);
  }
}

/** A colour laid over white paper at `coverage` strength, as a 6-hex string. */
function tintedHex(colorHex: string, coverage: number): string {
  const k = Math.min(1, Math.max(0, coverage));
  if (k >= 1) return colorHex;
  const channel = (at: number): string => {
    const c = Number.parseInt(colorHex.slice(at, at + 2), 16);
    const mixed = Math.round(255 - (255 - (Number.isFinite(c) ? c : 0)) * k);
    return mixed.toString(16).toUpperCase().padStart(2, '0');
  };
  return `${channel(0)}${channel(2)}${channel(4)}`;
}

/**
 * The `/Font` resources of one dictionary, built into interpreter fonts. Shared
 * with the path and picture walks, which need them for one thing only: a Type 3
 * font's glyphs are content streams (§9.6.5), and a walk that interprets with
 * no fonts at all cannot see that there is anything to run.
 *
 * @param file      The owning file.
 * @param resources The resource dictionary to read `/Font` from.
 * @returns Resource name → font, skipping any the reader cannot build.
 */
export function buildFonts(
  file: PdfFile,
  resources: PdfDict | undefined,
): Map<string, ContentFont> {
  const fonts = new Map<string, ContentFont>();
  if (!resources) return fonts;
  const fontContainer = file.get(resources, 'Font');
  if (!(fontContainer instanceof Map)) return fonts;
  for (const [fontName, fontRef] of fontContainer) {
    const fontDict = file.resolve(fontRef);
    if (fontDict instanceof Map) {
      try {
        fonts.set(fontName, buildContentFont(file, fontDict));
      } catch {
        // A malformed font is skipped — its text falls back to Latin-1.
      }
    }
  }
  return fonts;
}

function matrixOf(file: PdfFile, dict: PdfDict): Matrix {
  const m = file.resolve(dict.get('Matrix') ?? PDF_NULL);
  if (Array.isArray(m) && m.length >= 6 && m.every((v) => typeof v === 'number')) {
    return [m[0], m[1], m[2], m[3], m[4], m[5]] as Matrix;
  }
  return IDENTITY;
}

interface LinkRect {
  readonly rect: Rectangle; // [x0, y0, x1, y1], normalised
  readonly href: string;
}

// §12.5.6.5 — /Link annotations with a §12.6.4.7 URI action. Internal GoTo links
// (named destinations) are skipped: the reconstruction has no page anchors.
function collectLinks(file: PdfFile, page: PdfPage): Array<LinkRect> {
  const annots = file.get(page.dict, 'Annots');
  if (!Array.isArray(annots)) return [];
  const out: Array<LinkRect> = [];
  for (const a of annots) {
    const annot = file.resolve(a);
    if (!(annot instanceof Map)) continue;
    const sub = file.get(annot, 'Subtype');
    if (!(sub instanceof PdfName) || sub.value !== 'Link') continue;
    const rect = normRect(file.get(annot, 'Rect'));
    if (!rect) continue;
    const action = file.get(annot, 'A');
    if (!(action instanceof Map)) continue;
    const s = file.get(action, 'S');
    if (!(s instanceof PdfName) || s.value !== 'URI') continue;
    const uri = file.get(action, 'URI');
    if (typeof uri === 'string' && uri.length > 0) out.push({ rect, href: uri });
  }
  return out;
}

function normRect(v: unknown): Rectangle | undefined {
  if (!Array.isArray(v) || v.length < 4) return undefined;
  const n = v.slice(0, 4).map((x) => (typeof x === 'number' ? x : NaN));
  if (n.some((x) => !Number.isFinite(x))) return undefined;
  return [
    Math.min(n[0]!, n[2]!),
    Math.min(n[1]!, n[3]!),
    Math.max(n[0]!, n[2]!),
    Math.max(n[1]!, n[3]!),
  ];
}

// A run's glyph origin (baseline) lies within the rect (small tolerances absorb
// the baseline-to-rect-bottom gap).
function inRect(x: number, y: number, r: Rectangle): boolean {
  return x >= r[0] - 1 && x <= r[2] + 1 && y >= r[1] - 2 && y <= r[3] + 2;
}
