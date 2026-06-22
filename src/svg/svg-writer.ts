// SVG writer (ir-design §7 / stage 6): the third adapter, written purely
// against the PageDoc schema (LaidOutDocument/PageItem). It deliberately
// knows nothing about OOXML or the PDF writer.
//
// PageItem coordinates are top-left/y-down (the frame the schema froze on at
// stage 6.4) — exactly SVG's own frame, so this writer emits them verbatim;
// it is the PDF emitter that converts into PDF's y-up frame at emission.
//
// Pages stack vertically in one <svg>, separated by a gap — a faithful,
// dependency-free preview of the laid-out document.

import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { Loss } from '@/core/ir';
import type { LaidOutDocument, LaidOutPage, TextLineItem } from '@/layout/page-doc';
import type { PathSegment, VectorShape } from '@/core/vector';
import { svgPathData } from '@/core/vector';

import { FEATURES } from '@/core/ir';
import { toBase64 } from '@/core/bytes';
import { gradientSvgDef } from '@/core/drawingml/shape-render';
import { paintPlan } from '@/layout/page-doc';

const PAGE_GAP = 12;

/** Options for {@link writeSvg}. */
export interface SvgWriteOptions {
  /** Gap between stacked pages, in px/pt (default 12). */
  readonly pageGap?: number;
}

/**
 * Render a {@link LaidOutDocument} to a single SVG preview (ir-design §7 /
 * stage 6). Written purely against the PageDoc schema; knows nothing about OOXML
 * or the PDF writer. PageItem coordinates are already top-left/y-down — SVG's own
 * frame — so they are emitted verbatim. Pages stack vertically in one `<svg>`,
 * separated by {@link SvgWriteOptions.pageGap}, each on a white outlined rect so
 * they read as pages.
 *
 * @param laid The laid-out, paginated document from the layout engine.
 * @param opts Optional page-gap override.
 * @returns The encoded SVG bytes plus the recorded {@link Loss} list.
 */
export function writeSvg(laid: LaidOutDocument, opts: SvgWriteOptions = {}): WriteResult {
  const gap = opts.pageGap ?? PAGE_GAP;
  const losses: Array<Loss> = [];
  const width = Math.max(1, ...laid.pages.map((p) => p.width));
  const height =
    laid.pages.reduce((s, p) => s + p.height, 0) + gap * Math.max(0, laid.pages.length - 1);

  const parts: Array<string> = [];
  parts.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(width)}" height="${fmt(height)}" viewBox="0 0 ${fmt(width)} ${fmt(height)}">`,
  );
  let yOffset = 0;
  const idc = { n: 0 }; // unique gradient-id counter across the whole document
  laid.pages.forEach((page, i) => {
    parts.push(`<g transform="translate(0 ${fmt(yOffset)})" data-page="${i + 1}">`);
    // Page background + outline so stacked pages read as pages.
    parts.push(
      `<rect x="0" y="0" width="${fmt(page.width)}" height="${fmt(page.height)}" fill="#ffffff" stroke="#cccccc"/>`,
    );
    emitPage(parts, page, laid, losses, idc);
    parts.push('</g>');
    yOffset += page.height + gap;
  });
  parts.push('</svg>');

  return { bytes: new TextEncoder().encode(parts.join('\n')), losses };
}

/**
 * The page-medium {@link DocumentWriter} adapter (id `'svg'`), wrapping
 * {@link writeSvg}, with the set of {@link FEATURES} it renders.
 */
export const svgWriter: DocumentWriter<LaidOutDocument> = {
  id: 'svg',
  consumes: 'page',
  supports: new Set([FEATURES.text, FEATURES.tables, FEATURES.images, FEATURES.shapes]),
  write: (doc, opts) => writeSvg(doc, opts ?? {}),
};

function emitPage(
  out: Array<string>,
  page: LaidOutPage,
  laid: LaidOutDocument,
  losses: Array<Loss>,
  idc: { n: number },
): void {
  // The shared canonical paint order — one owner for every writer. PageItem
  // coordinates are already top-left/y-down: SVG's native frame.
  const plan = paintPlan(page.commands);

  for (const f of plan.fills) {
    out.push(
      `<rect x="${fmt(f.x)}" y="${fmt(f.y)}" width="${fmt(f.width)}" height="${fmt(f.height)}" fill="#${f.fillColorHex}"/>`,
    );
  }

  for (const img of plan.images) {
    const href = imageHref(img.imageResourceName, laid);
    if (!href) continue;
    out.push(
      `<image x="${fmt(img.x)}" y="${fmt(img.y)}" width="${fmt(img.width)}" height="${fmt(img.height)}" href="${href}" preserveAspectRatio="none"/>`,
    );
  }

  for (const b of plan.borders) {
    const x2 = b.x + b.width;
    const yTop = b.y;
    const yBottom = b.y + b.height;
    const [ax, ay, bx, by] =
      b.side === 'top'
        ? [b.x, yTop, x2, yTop]
        : b.side === 'bottom'
          ? [b.x, yBottom, x2, yBottom]
          : b.side === 'left'
            ? [b.x, yTop, b.x, yBottom]
            : [x2, yTop, x2, yBottom];
    out.push(
      `<line x1="${fmt(ax)}" y1="${fmt(ay)}" x2="${fmt(bx)}" y2="${fmt(by)}" stroke="#${b.borderColorHex}" stroke-width="${fmt(b.borderSizePt)}"/>`,
    );
  }

  for (const sh of plan.shapes) {
    emitShape(out, sh.shape, idc);
  }

  for (const t of plan.lines) {
    emitTextLine(out, t, losses);
  }
}

function emitTextLine(out: Array<string>, item: TextLineItem, losses: Array<Loss>): void {
  const y = item.baselineY;
  let x: number = item.originX;
  for (const tok of item.line.tokens) {
    if (tok.kind === 'image') {
      x += tok.widthPt; // inline image boxes reserve space; not rendered in v0
      continue;
    }
    if (tok.kind === 'math') {
      losses.push({
        severity: 'dropped',
        feature: FEATURES.math,
        detail: 'inline math box not rendered by the SVG writer (v0)',
      });
      x += tok.widthPt;
      continue;
    }
    if (tok.text.trim().length > 0) {
      out.push(
        `<text x="${fmt(x)}" y="${fmt(y)}" font-family="sans-serif" font-size="${fmt(tok.fontSizePt)}" fill="#${tok.resolvedRun.colorHex}">${escapeXml(tok.text)}</text>`,
      );
    }
    x += tok.widthPt;
  }
}

function emitShape(out: Array<string>, shape: VectorShape, idc: { n: number }): void {
  const [a, b, c, d, e, f] = shape.transform;
  // The stored CTM maps the shape's local y-up frame straight into the
  // top-left page frame — SVG's matrix() convention verbatim.
  const transform = `matrix(${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ${fmt(e)} ${fmt(f)})`;
  let fill: string;
  if (shape.fillGradient) {
    const id = `grad${idc.n++}`;
    out.push(gradientSvgDef(id, shape.fillGradient));
    fill = `url(#${id})`;
  } else {
    fill = shape.fillColorHex ? `#${shape.fillColorHex}` : 'none';
  }
  const stroke = shape.stroke
    ? ` stroke="#${shape.stroke.colorHex}" stroke-width="${fmt(shape.stroke.widthPt)}"`
    : '';
  for (const path of shape.paths) {
    const d2 = pathData(path.segments);
    const rule = path.fillRule === 'evenodd' ? ' fill-rule="evenodd"' : '';
    out.push(`<path d="${d2}" fill="${fill}"${rule}${stroke} transform="${transform}"/>`);
  }
}

// Local path coordinates are y-up; the transform (page flip composed in by
// layout) maps them into the y-down page frame. Emit the raw coordinates.
function pathData(segments: ReadonlyArray<PathSegment>): string {
  return svgPathData(segments, fmt);
}

function imageHref(resourceName: string, laid: LaidOutDocument): string | undefined {
  if (!resourceName) return undefined;
  // resourceName (ImN) → ResourceId → bytes from the content-addressed store;
  // the mime comes from the layout-time prepare (the image expert) instead of
  // re-sniffing the bytes here.
  for (const [resourceId, res] of laid.imageResources) {
    if (res.resourceName !== resourceName) continue;
    const bytes = laid.resources.get(resourceId);
    if (!bytes) return undefined;
    return `data:${res.prepared.mimeType};base64,${toBase64(bytes)}`;
  }
  return undefined;
}

function escapeXml(s: string): string {
  return s
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function fmt(n: number): string {
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
}
