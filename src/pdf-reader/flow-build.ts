// E-PDF — shared FlowDoc construction for the two reconstruction paths (the
// tagged fast-path EP3 and the heuristic layout path EP4). A reconstructed PDF
// carries a body of paragraphs/tables/images over the empty style sheet, with
// any lifted image bytes (EP6) in the resource store the writers embed from.

import type {
  BodyElement,
  CustomPathCmd,
  ParagraphProperties,
  SectionProperties,
  ShapeFill,
  ShapeLine,
} from '@/core/document-model';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';

import type { PdfImage } from './images';
import type { PdfPage } from './document';
import type { PdfVector } from './vector';
import { ResourceStore, pt } from '@/core/ir';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';

// A reconstruction's document plus the losses incurred reading it (e.g. an
// undecodable image colour space) — surfaced through the reader's LossReport.
export interface Reconstruction {
  readonly doc: FlowDoc;
  readonly losses: ReadonlyArray<Loss>;
}

export function paragraphBlock(text: string, outlineLevel?: number): BodyElement {
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
  return {
    kind: 'paragraph',
    paragraph: { properties, runs: text.length > 0 ? [{ text, properties: {} }] : [] },
  };
}

// One piece of reconstructed text, carrying any hyperlink (E-PDF EP8).
export interface TextSpan {
  readonly text: string;
  readonly href?: string;
}

// Build a paragraph from positioned spans, coalescing consecutive spans that
// share an href into one run (so a link survives as its own run) and squashing
// whitespace. With no hrefs this collapses to a single run — the same shape
// `paragraphBlock` produces.
export function paragraphFromRuns(
  spans: ReadonlyArray<TextSpan>,
  outlineLevel?: number,
): BodyElement {
  const merged: Array<{ text: string; href?: string }> = [];
  for (const s of spans) {
    const last = merged[merged.length - 1];
    if (last && last.href === s.href) last.text += s.text;
    else if (s.href !== undefined) merged.push({ text: s.text, href: s.href });
    else merged.push({ text: s.text });
  }
  const runs = merged
    .map((m) => ({ text: m.text.replace(/\s+/g, ' '), href: m.href }))
    .filter((m) => m.text.length > 0);
  // Trim the paragraph's outer whitespace.
  if (runs.length > 0) {
    runs[0]!.text = runs[0]!.text.replace(/^ /, '');
    runs[runs.length - 1]!.text = runs[runs.length - 1]!.text.replace(/ $/, '');
  }
  const properties: ParagraphProperties = outlineLevel !== undefined ? { outlineLevel } : {};
  return {
    kind: 'paragraph',
    paragraph: {
      properties,
      runs: runs
        .filter((r) => r.text.length > 0)
        .map((r) => ({ text: r.text, properties: {}, ...(r.href ? { href: r.href } : {}) })),
    },
  };
}

// Store the image bytes (content-addressed dedup) and build the block that
// references them, sized in points from the placement CTM.
export function imageBlock(image: PdfImage, resources: ResourceStore, alt?: string): BodyElement {
  const resource = resources.put(image.bytes);
  return {
    kind: 'image',
    image: {
      resource,
      width: pt(image.widthPt),
      height: pt(image.heightPt),
      paragraphProperties: {},
      ...(alt ? { altText: alt } : {}),
    },
  };
}

// Collapse losses sharing a message (the same colour space dropped on many pages).
export function dedupeLosses(losses: ReadonlyArray<Loss>): Array<Loss> {
  const byDetail = new Map<string, Loss>();
  for (const loss of losses) if (!byDetail.has(loss.detail)) byDetail.set(loss.detail, loss);
  return [...byDetail.values()];
}

// A lifted path (filled EP10 / stroked EP11) as a custom-geometry shape.
// Page-space points (y-up) become path-space (bbox-relative, y-down); the shape
// is sized from the bounding box (plus the stroke thickness) and placed in flow
// order by the caller. A fill becomes a solid fill, a stroke becomes the outline.
export function shapeBlock(v: PdfVector): BodyElement {
  const w = v.maxX - v.minX;
  const h = v.maxY - v.minY;
  const fx = (x: number): number => x - v.minX;
  const fy = (y: number): number => v.maxY - y; // flip to top-left origin
  const commands: Array<CustomPathCmd> = v.segs.map((s): CustomPathCmd => {
    switch (s.op) {
      case 'move':
        return { cmd: 'move', x: fx(s.x), y: fy(s.y) };
      case 'line':
        return { cmd: 'line', x: fx(s.x), y: fy(s.y) };
      case 'cubic':
        return {
          cmd: 'cubic',
          x1: fx(s.x1),
          y1: fy(s.y1),
          x2: fx(s.x2),
          y2: fy(s.y2),
          x: fx(s.x),
          y: fy(s.y),
        };
      case 'close':
        return { cmd: 'close' };
    }
  });
  // A stroked line can be geometrically flat (a horizontal rule has h≈0); give
  // the shape box at least the stroke thickness so the line has room to draw.
  const thick = v.strokeHex !== undefined ? Math.max(v.lineWidth ?? 0.75, 0.5) : 0;
  const fill: ShapeFill =
    v.gradient !== undefined
      ? { kind: 'gradient', gradient: v.gradient }
      : v.fillHex !== undefined
        ? { kind: 'solid', colorHex: v.fillHex }
        : { kind: 'none' };
  const line: ShapeLine | undefined =
    v.strokeHex !== undefined
      ? { width: pt(thick), colorHex: v.strokeHex, fill: 'solid' }
      : undefined;
  return {
    kind: 'shape',
    shape: {
      width: pt(Math.max(w, thick)),
      height: pt(Math.max(h, thick)),
      geometry: { kind: 'custom', custom: { pathWidth: w, pathHeight: h, commands } },
      fill,
      ...(line ? { line } : {}),
      paragraphProperties: {},
    },
  };
}

// Derive the FlowDoc section geometry from the source pages so a reconstructed
// PDF re-renders at its real page size and orientation rather than the layout
// engine's A4 default — without it an A3 source paginates onto several A4
// pages, a wide/landscape page is letterboxed, and so on. PDF user-space units
// are points, matching Pt, so the MediaBox extents map straight to the page
// size. Uses the first page's box (the near-universal uniform-size case); a PDF
// whose pages differ in size reflows to this single geometry — a known
// approximation, still far better than a fixed A4.
export function sectionFromPdfPages(pages: ReadonlyArray<PdfPage>): SectionProperties | undefined {
  const box = pages[0]?.mediaBox;
  if (!box) return undefined;
  const width = Math.abs(box[2] - box[0]);
  const height = Math.abs(box[3] - box[1]);
  if (!(width > 0 && height > 0)) return undefined;
  return {
    pageSize: {
      width: pt(width),
      height: pt(height),
      orientation: width > height ? 'landscape' : 'portrait',
    },
    // A PDF page has no margin model — text is positioned absolutely anywhere on
    // the MediaBox — so the page box is the content box. Zero margins keep the
    // reflow from inventing a 1-inch inset the source never had, never clip a
    // small page (the layout default of 1 inch can exceed a tiny MediaBox), and
    // reduce spurious over-pagination.
    margins: { top: pt(0), right: pt(0), bottom: pt(0), left: pt(0) },
    headers: [],
    footers: [],
  };
}

export function buildFlowDoc(
  body: ReadonlyArray<BodyElement>,
  resources: ResourceStore = new ResourceStore(),
  section?: SectionProperties,
): FlowDoc {
  return {
    kind: 'flow',
    body: resolveBodyStyles([...body], EMPTY_STYLE_SHEET),
    sections: [],
    ...(section ? { section } : {}),
    styles: EMPTY_STYLE_SHEET,
    resources,
  };
}
