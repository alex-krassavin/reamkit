// E-PDF EP2/EP8 — page text extraction. Builds the page's font map from its
// /Resources /Font dictionary, runs the content-stream interpreter to get the
// positioned text runs, then tags any run whose origin falls inside a /Link
// annotation's /Rect with that link's URI (EP8) so hyperlinks survive.

import { interpretContent } from './content';
import { buildContentFont } from './font';
import type { ContentFont, TextRun } from './content';
import type { PdfFile, PdfPage, Rectangle } from './document';

import { PdfName } from '@/pdf/objects';

export function extractPageText(file: PdfFile, page: PdfPage): Array<TextRun> {
  const fonts = new Map<string, ContentFont>();
  if (page.resources) {
    const fontContainer = file.get(page.resources, 'Font');
    if (fontContainer instanceof Map) {
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
    }
  }
  const runs = interpretContent(file.pageContent(page), fonts).texts;
  const links = collectLinks(file, page);
  if (links.length === 0) return runs;
  return runs.map((run) => {
    const link = links.find((l) => inRect(run.x, run.y, l.rect));
    return link ? { ...run, href: link.href } : run;
  });
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
