// E-PDF EP10 — lift FILLED vector paths off a page. Runs the content
// interpreter for its fill placements (EP10) and keeps the ones that read as
// real graphics: it drops hairline/degenerate fills, invisible white fills, and
// the near-full-page background rectangle, so a reconstructed document gains
// genuine coloured shapes without the table-rule / page-background clutter.
// Strokes, clips, shadings and gradients are not captured (a documented loss).

import { interpretContent } from './content';
import type { ContentFont, PathSeg } from './content';

import type { PdfFile, PdfPage } from './document';

export interface PdfVector {
  readonly segs: ReadonlyArray<PathSeg>;
  readonly fillHex: string;
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
  readonly mcid?: number;
}

const NO_FONTS: ReadonlyMap<string, ContentFont> = new Map();
const MIN_SIDE = 2; // pt — skip thin rules
const MIN_AREA = 16; // pt² — skip dots / hairlines
const MAX_VECTORS = 2000; // per-page DoS guard

export function collectPageVectors(file: PdfFile, page: PdfPage): Array<PdfVector> {
  const [px0, py0, px1, py1] = page.mediaBox;
  const pageArea = Math.max(1, Math.abs((px1 - px0) * (py1 - py0)));
  const out: Array<PdfVector> = [];
  for (const v of interpretContent(file.pageContent(page), NO_FONTS).vectors) {
    if (out.length >= MAX_VECTORS) break;
    if (v.fillHex === 'FFFFFF') continue; // invisible on white paper
    const b = bbox(v.segs);
    if (!b) continue;
    const w = b.maxX - b.minX;
    const h = b.maxY - b.minY;
    if (w < MIN_SIDE || h < MIN_SIDE || w * h < MIN_AREA) continue;
    if (w * h > 0.85 * pageArea) continue; // a page-background fill
    out.push({
      segs: v.segs,
      fillHex: v.fillHex,
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
