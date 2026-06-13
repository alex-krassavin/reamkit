// E-PDF EP10/EP11/EP16c — lift painted vector paths off a page. Runs the content
// interpreter for its fill (EP10), stroke (EP11) and shading-pattern (EP16c)
// placements and keeps the ones that read as real graphics: it drops
// hairline/degenerate fills, invisible white paint, short stroke specks, and the
// near-full-page background, so a reconstructed document gains genuine coloured
// shapes, lines and gradients without the dot / page-background clutter. Clips
// and the bare `sh` operator are not captured (a documented loss).

import { IDENTITY, interpretContent } from './content';
import { buildShadingMap } from './shading';
import type { ContentFont, PathSeg } from './content';
import type { ShapeGradient } from '@/core/vector';

import type { PdfFile, PdfPage } from './document';

export interface PdfVector {
  readonly segs: ReadonlyArray<PathSeg>;
  readonly fillHex?: string; // present iff a qualifying solid fill survived (EP10)
  readonly gradient?: ShapeGradient; // present iff a shading-pattern fill survived (EP16c)
  readonly strokeHex?: string; // present iff a qualifying stroke survived (EP11)
  readonly lineWidth?: number; // stroke width in page-space points (EP11)
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
  readonly mcid?: number;
}

const NO_FONTS: ReadonlyMap<string, ContentFont> = new Map();
const MIN_SIDE = 2; // pt — skip thin filled rules
const MIN_AREA = 16; // pt² — skip dots / hairlines
const MIN_STROKE_LEN = 6; // pt — skip stroke specks (tick marks, dots)
const MAX_VECTORS = 2000; // per-page DoS guard

export function collectPageVectors(file: PdfFile, page: PdfPage): Array<PdfVector> {
  const [px0, py0, px1, py1] = page.mediaBox;
  const pageArea = Math.max(1, Math.abs((px1 - px0) * (py1 - py0)));
  const shadings = buildShadingMap(file, page);
  const out: Array<PdfVector> = [];
  for (const v of interpretContent(file.pageContent(page), NO_FONTS, IDENTITY, shadings).vectors) {
    if (out.length >= MAX_VECTORS) break;
    const b = bbox(v.segs);
    if (!b) continue;
    const w = b.maxX - b.minX;
    const h = b.maxY - b.minY;
    const area = w * h;
    // A fill (solid or gradient) must be a non-white area larger than a hairline
    // and smaller than a page background; a stroke must be a non-white line
    // longer than a speck.
    const solidFill = v.fillHex !== undefined && v.fillHex !== 'FFFFFF';
    const filled =
      (v.gradient !== undefined || solidFill) &&
      w >= MIN_SIDE &&
      h >= MIN_SIDE &&
      area >= MIN_AREA &&
      area <= 0.85 * pageArea;
    const stroked =
      v.strokeHex !== undefined &&
      v.strokeHex !== 'FFFFFF' &&
      Math.max(w, h) >= MIN_STROKE_LEN &&
      area <= 0.85 * pageArea;
    if (!filled && !stroked) continue;
    out.push({
      segs: v.segs,
      ...(filled
        ? v.gradient
          ? { gradient: v.gradient }
          : v.fillHex !== undefined
            ? { fillHex: v.fillHex }
            : {}
        : {}),
      ...(stroked
        ? {
            strokeHex: v.strokeHex,
            ...(v.lineWidth !== undefined ? { lineWidth: v.lineWidth } : {}),
          }
        : {}),
      ...b,
      ...(v.mcid !== undefined ? { mcid: v.mcid } : {}),
    });
  }
  return out;
}

function bbox(
  segs: ReadonlyArray<PathSeg>,
): { minX: number; minY: number; maxX: number; maxY: number } | undefined {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  const add = (x: number, y: number): void => {
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x);
    maxY = Math.max(maxY, y);
  };
  for (const s of segs) {
    if (s.op === 'move' || s.op === 'line') add(s.x, s.y);
    else if (s.op === 'cubic') {
      add(s.x1, s.y1);
      add(s.x2, s.y2);
      add(s.x, s.y);
    }
  }
  return Number.isFinite(minX) ? { minX, minY, maxX, maxY } : undefined;
}
