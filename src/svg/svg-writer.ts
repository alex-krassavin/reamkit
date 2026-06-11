// SVG writer (ir-design §7 / stage 6): the third adapter, written purely
// against the PageDoc draft (LaidOutDocument/PageItem) — its job is to
// crash-test that contract before the schema freezes. It deliberately knows
// nothing about OOXML or the PDF writer.
//
// PageItem coordinates are currently in the PDF page frame (y-up, bottom-left
// origin); SVG is y-down/top-left, so this writer flips — exactly mirroring
// how the PDF writer owns its own frame. The flip moving INTO the schema
// (top-left as canonical) is the planned stabilization step.
//
// Pages stack vertically in one <svg>, separated by a gap — a faithful,
// dependency-free preview of the laid-out document.

import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { Loss } from '@/core/ir';
import type {
  LaidOutDocument,
  LaidOutPage,
  PageItem,
  TextLineItem,
} from '@/pdf/styled-page-renderer';
import type { PathSegment, VectorShape } from '@/pdf/vector-graphics';

import { FEATURES } from '@/core/ir';
import { paintPlan } from '@/pdf/styled-page-renderer';

const PAGE_GAP = 12;

export interface SvgWriteOptions {
  /** Gap between stacked pages, in px/pt (default 12). */
  readonly pageGap?: number;
}

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
  laid.pages.forEach((page, i) => {
    parts.push(`<g transform="translate(0 ${fmt(yOffset)})" data-page="${i + 1}">`);
    // Page background + outline so stacked pages read as pages.
    parts.push(
      `<rect x="0" y="0" width="${fmt(page.width)}" height="${fmt(page.height)}" fill="#ffffff" stroke="#cccccc"/>`,
    );
    emitPage(parts, page, laid, losses);
    parts.push('</g>');
    yOffset += page.height + gap;
  });
  parts.push('</svg>');

  return { bytes: new TextEncoder().encode(parts.join('\n')), losses };
}

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
): void {
  const H = page.height; // y-flip: svgY = H - pdfY
  // The shared canonical paint order — one owner for every writer.
  const plan = paintPlan(page.commands);

  for (const f of plan.fills) {
    out.push(
      `<rect x="${fmt(f.x)}" y="${fmt(H - f.y - f.height)}" width="${fmt(f.width)}" height="${fmt(f.height)}" fill="#${f.fillColorHex}"/>`,
    );
  }

  for (const img of plan.images) {
    const href = imageHref(img.imageResourceName, laid);
    if (!href) continue;
    out.push(
      `<image x="${fmt(img.x)}" y="${fmt(H - img.y - img.height)}" width="${fmt(img.width)}" height="${fmt(img.height)}" href="${href}" preserveAspectRatio="none"/>`,
    );
  }

  for (const b of plan.borders) {
    const x2 = b.x + b.width;
    const yTop = H - b.y - b.height;
    const yBottom = H - b.y;
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
    emitShape(out, sh.shape, H);
  }

  for (const t of plan.lines) {
    emitTextLine(out, t, H, losses);
  }
}

function emitTextLine(
  out: Array<string>,
  item: TextLineItem,
  pageHeight: number,
  losses: Array<Loss>,
): void {
  const y = pageHeight - item.baselineY;
  let x = item.originX;
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

function emitShape(out: Array<string>, shape: VectorShape, pageHeight: number): void {
  const [a, b, c, d, e, f] = shape.transform;
  // PDF CTM maps local → page (y-up). Compose with the page flip so the SVG
  // group lands where the PDF painted: svg = flip ∘ ctm.
  const transform = `matrix(${fmt(a)} ${fmt(-b)} ${fmt(-c)} ${fmt(d)} ${fmt(e)} ${fmt(pageHeight - f)})`;
  const fill = shape.fillColorHex ? `#${shape.fillColorHex}` : 'none';
  const stroke = shape.stroke
    ? ` stroke="#${shape.stroke.colorHex}" stroke-width="${fmt(shape.stroke.widthPt)}"`
    : '';
  for (const path of shape.paths) {
    const d2 = pathData(path.segments);
    const rule = path.fillRule === 'evenodd' ? ' fill-rule="evenodd"' : '';
    out.push(`<path d="${d2}" fill="${fill}"${rule}${stroke} transform="${transform}"/>`);
  }
}

// Local path coordinates are y-up; the group transform's negative b/c terms
// flip them. Emit the raw coordinates.
function pathData(segments: ReadonlyArray<PathSegment>): string {
  const parts: Array<string> = [];
  for (const s of segments) {
    if (s.op === 'move') parts.push(`M ${fmt(s.x)} ${fmt(-s.y)}`);
    else if (s.op === 'line') parts.push(`L ${fmt(s.x)} ${fmt(-s.y)}`);
    else if (s.op === 'cubic')
      parts.push(
        `C ${fmt(s.x1)} ${fmt(-s.y1)} ${fmt(s.x2)} ${fmt(-s.y2)} ${fmt(s.x)} ${fmt(-s.y)}`,
      );
    else parts.push('Z');
  }
  return parts.join(' ');
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

function toBase64(bytes: Uint8Array): string {
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  // btoa is available in browsers/workers/Node 16+.
  return btoa(bin);
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
