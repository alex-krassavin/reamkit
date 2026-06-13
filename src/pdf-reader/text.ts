// E-PDF EP2/EP8/EP13 — page text extraction. Interprets the page content (and,
// recursively, the Form XObjects it paints — EP13, whose text would otherwise be
// missed) into positioned runs, then tags any run whose origin falls inside a
// /Link annotation's /Rect with that link's URI (EP8) so hyperlinks survive.

import { IDENTITY, interpretContent, multiply } from './content';
import { buildContentFont } from './font';
import type { ContentFont, Matrix, TextRun } from './content';
import type { PdfDict } from '@/pdf/objects';
import type { PdfFile, PdfPage, Rectangle } from './document';

import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

const MAX_FORM_DEPTH = 8;

export function extractPageText(file: PdfFile, page: PdfPage): Array<TextRun> {
  const runs: Array<TextRun> = [];
  collectRuns(file, page.resources, file.pageContent(page), IDENTITY, 0, new Set(), runs);
  const links = collectLinks(file, page);
  if (links.length === 0) return runs;
  return runs.map((run) => {
    const link = links.find((l) => inRect(run.x, run.y, l.rect));
    return link ? { ...run, href: link.href } : run;
  });
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
  out.push(...result.texts);
  if (depth >= MAX_FORM_DEPTH || !resources) return;
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

function buildFonts(file: PdfFile, resources: PdfDict | undefined): Map<string, ContentFont> {
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
