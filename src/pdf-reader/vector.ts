// E-PDF EP10/EP11/EP16c — lift painted vector paths off a page. Runs the content
// interpreter for its fill (EP10), stroke (EP11) and shading-pattern (EP16c)
// placements and keeps the ones that read as real graphics: it drops
// hairline/degenerate fills, invisible white paint, short stroke specks, and the
// near-full-page background, so a reconstructed document gains genuine coloured
// shapes, lines and gradients without the dot / page-background clutter. Clips
// and the bare `sh` operator are not captured (a documented loss).

import { IDENTITY, interpretContent, multiply } from './content';
import { buildShadingMap } from './shading';
import { collectPageAppearances } from './annots';
import type { ContentFont, ImagePlacement, Matrix, PathSeg, VectorPlacement } from './content';
import type { ShapeGradient } from '@/core/vector';
import type { PdfDict } from '@/pdf/objects';

import type { PdfFile, PdfPage } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

/**
 * Every vector the page paints, its FORM XOBJECTS included (§8.8).
 *
 * A `Do` of a form is a call: its content stream draws in the caller's space
 * through the form's own `/Matrix`. Interpreting the page stream alone reads
 * only what the page drew directly, and a document that puts its artwork in
 * forms — as every CAD and drawing producer does — comes back with none of it.
 * 22060_A1_01_Plans.pdf holds eleven, and its floor plans were simply absent.
 *
 * The same walk `collectPageImages` already makes, with the same depth and
 * cycle guards, collecting paths instead of pictures.
 */
function paintedVectors(
  file: PdfFile,
  page: PdfPage,
  shadings: ReadonlyMap<string, ShapeGradient>,
): Array<VectorPlacement & { orderKey: ReadonlyArray<number> }> {
  const out: Array<VectorPlacement & { orderKey: ReadonlyArray<number> }> = [];
  const visiting = new Set<PdfStream>();
  const walk = (
    resources: PdfDict | undefined,
    content: Uint8Array,
    baseCtm: Matrix,
    depth: number,
    prefix: ReadonlyArray<number>,
  ): void => {
    if (out.length >= MAX_VECTORS) return;
    const xobjects = resources ? file.get(resources, 'XObject') : PDF_NULL;
    const xobjDict = xobjects instanceof Map ? xobjects : undefined;
    const result = interpretContent(content, NO_FONTS, baseCtm, shadings);

    // §8.5.3 — later marks cover earlier ones, and a form is drawn where its
    // `Do` stands, not after everything around it. Walking the stream first and
    // its forms afterwards puts every form on top: 22060_A1_01_Plans.pdf backs
    // its legend with a white box inside a form, and hoisted to the end that box
    // covered the legend's own words.
    const events: Array<{ order: number; vector?: VectorPlacement; xobject?: ImagePlacement }> = [
      ...result.vectors.map((vector) => ({ order: vector.order, vector })),
      ...result.images.map((xobject) => ({ order: xobject.order, xobject })),
    ].sort((a, b) => a.order - b.order);

    for (const event of events) {
      if (out.length >= MAX_VECTORS) return;
      if (event.vector) {
        out.push({ ...event.vector, orderKey: [...prefix, event.order] });
        continue;
      }
      const placement = event.xobject!;
      if (depth >= MAX_FORM_DEPTH) continue;
      const stream = xobjDict ? file.resolve(xobjDict.get(placement.name) ?? PDF_NULL) : PDF_NULL;
      if (!(stream instanceof PdfStream) || visiting.has(stream)) continue;
      const subtype = file.get(stream.dict, 'Subtype');
      if (!(subtype instanceof PdfName) || subtype.value !== 'Form') continue;
      visiting.add(stream);
      const formRes = file.get(stream.dict, 'Resources');
      walk(
        formRes instanceof Map ? formRes : resources,
        file.streamData(stream),
        multiply(formMatrix(file, stream.dict), placement.ctm),
        depth + 1,
        [...prefix, placement.order],
      );
      visiting.delete(stream);
    }
  };
  walk(page.resources, file.pageContent(page), IDENTITY, 0, []);
  // §12.5.5 — an annotation's appearance paints OVER the page it sits on, so
  // its marks sort after every mark of the content stream.
  collectPageAppearances(file, page).forEach((appearance, index) => {
    walk(
      appearance.resources ?? page.resources,
      file.streamData(appearance.stream),
      appearance.ctm,
      1,
      [Number.MAX_SAFE_INTEGER, index],
    );
  });
  return out;
}

/** §8.10.2 `/Matrix` — the form's own space, composed onto the placement CTM. */
function formMatrix(file: PdfFile, dict: PdfDict): Matrix {
  const m = file.resolve(dict.get('Matrix') ?? PDF_NULL);
  if (!Array.isArray(m) || m.length !== 6) return IDENTITY;
  const n = m.map((v) => (typeof file.resolve(v) === 'number' ? (file.resolve(v) as number) : 0));
  return [n[0]!, n[1]!, n[2]!, n[3]!, n[4]!, n[5]!];
}

/**
 * One painted vector path lifted off a page (E-PDF EP10/EP11/EP16c): the path
 * segments, whichever of solid fill / gradient fill / stroke survived the
 * de-cluttering filter, and the path's page-space bounding box plus enclosing
 * marked-content id.
 */
export interface PdfVector {
  /**
   * §8.5.3 — where this was painted, as the chain of positions leading to it:
   * `[4]` is the fifth mark of the page, `[4, 2]` the third mark of the form
   * that mark called. Compared element by element it is the page's painting
   * order across forms and patterns alike, which is the only thing that says
   * what covers what.
   */
  readonly orderKey: ReadonlyArray<number>;
  readonly segs: ReadonlyArray<PathSeg>;
  /** Present iff a qualifying solid fill survived (EP10). */
  readonly fillHex?: string;
  /** Present iff a shading-pattern fill survived (EP16c). */
  readonly gradient?: ShapeGradient;
  /** Present iff a qualifying stroke survived (EP11). */
  readonly strokeHex?: string;
  /** Stroke width in page-space points (EP11). */
  readonly lineWidth?: number;
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
const MIN_RULE_LEN = 6; // pt — a filled rule shorter than this is a speck
const MAX_VECTORS = 2000; // per-page DoS guard
const MAX_FORM_DEPTH = 12;

/**
 * Lift the painted vector paths off a page (E-PDF EP10/EP11/EP16c). Runs the
 * content interpreter for its fill (EP10), stroke (EP11) and shading-pattern
 * (EP16c) placements and keeps only the ones that read as real graphics:
 * hairline/degenerate fills, invisible white paint, short stroke specks and the
 * near-full-page background are all dropped, so a reconstructed document gains
 * genuine coloured shapes, lines and gradients without the dot / page-background
 * clutter. Clips and the bare `sh` operator are not captured (a documented loss).
 */
export function collectPageVectors(
  file: PdfFile,
  page: PdfPage,
  occupied: ReadonlyArray<Box> = [],
): Array<PdfVector> {
  const [px0, py0, px1, py1] = page.mediaBox;
  const pageArea = Math.max(1, Math.abs((px1 - px0) * (py1 - py0)));
  const shadings = buildShadingMap(file, page);
  const out: Array<PdfVector> = [];
  // What the page has painted so far. White paint is invisible only over white:
  // over anything else it is the thing that HIDES it, and the caller seeds this
  // with the pictures it has already placed.
  const painted: Array<Box> = [...occupied];
  for (const raw of paintedVectors(file, page, shadings)) {
    if (out.length >= MAX_VECTORS) break;
    const v = clipped(raw);
    if (!v) continue;
    const b = bbox(v.segs);
    if (!b) continue;
    const w = b.maxX - b.minX;
    const h = b.maxY - b.minY;
    const area = w * h;
    // A fill (solid or gradient) must be a non-white area larger than a hairline
    // and smaller than a page background; a stroke must be a non-white line
    // longer than a speck.
    // §8.7.3 — a path filled with a TILING pattern shows the pattern, not a
    // colour: `collectPageImages` walks into it and lifts what it draws. The
    // `fillHex` still standing on the placement is whatever colour was set
    // before the pattern was, and painting it covered 22060_A1_01_Plans.pdf's
    // four floor plans with four black rectangles.
    //
    // White is dropped as paint that shows nothing — except where something is
    // already painted under it, which is the one place white is not invisible
    // but OPAQUE. 22060_A1_01_Plans.pdf backs its legend with a white box over
    // a floor plan, and dropped it the plan and the title block read straight
    // through the legend's text.
    const white = v.fillHex === 'FFFFFF';
    const solidFill =
      v.patternName === undefined &&
      v.fillHex !== undefined &&
      (!white || painted.some((box) => overlaps(box, b)));
    // A fill is a BOX when both sides have some size to them, and a RULE when
    // one side is long and the other barely there. Only boxes were kept, and a
    // form's lines are not strokes at all: 160F-2019.pdf draws every rule as a
    // filled rectangle a half point high, so the whole grid — every box, every
    // line of the certificate — was thrown out as hairline clutter, and its
    // text arrived with no form under it.
    const long = Math.max(w, h);
    const short = Math.min(w, h);
    const isBox = short >= MIN_SIDE && area >= MIN_AREA;
    const isRule = long >= MIN_RULE_LEN && short > 0;
    const filled =
      (v.gradient !== undefined || solidFill) && (isBox || isRule) && area <= 0.85 * pageArea;
    const stroked =
      v.strokeHex !== undefined &&
      v.strokeHex !== 'FFFFFF' &&
      Math.max(w, h) >= MIN_STROKE_LEN &&
      area <= 0.85 * pageArea;
    if (!filled && !stroked) continue;
    painted.push(b);
    out.push({
      orderKey: v.orderKey,
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

/** A page-space rectangle: what something covers. */
interface Box {
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
}

/** Whether two boxes share any area at all. */
function overlaps(a: Box, b: Box): boolean {
  return a.minX < b.maxX && b.minX < a.maxX && a.minY < b.maxY && b.minY < a.maxY;
}

/**
 * §8.5.4 — what a painted path actually MARKS, once its clip is honoured.
 *
 * The mark is the intersection of the path with the clipping region, and one
 * observation decides it without intersecting anything: where the path covers
 * the clip, the intersection IS the clip. That is the stencil every drawing
 * producer writes — 22060_A1_01_Plans.pdf paints four black rectangles of
 * 397×421pt through clips cut to its floor plans, and read as rectangles they
 * flooded two thirds of an A3 sheet in black.
 *
 * So the smaller region stands: a path inside its clip is itself, a path around
 * its clip becomes the clip in the paint's own colour. Neither is a general
 * path intersection, and where the two merely overlap the answer is the smaller
 * of them — bounded, never larger than the truth.
 *
 * @param v The painted path as the interpreter saw it.
 * @returns The path to draw, or `undefined` when the clip leaves nothing.
 */
function clipped<T extends VectorPlacement>(v: T): T | undefined {
  const clip = v.clip;
  if (!clip) return v;
  const b = bbox(v.segs);
  if (!b) return v;
  // Disjoint: the clip lets none of it through.
  if (clip.minX >= b.maxX || clip.maxX <= b.minX || clip.minY >= b.maxY || clip.maxY <= b.minY) {
    return undefined;
  }
  const pathArea = Math.max(0, b.maxX - b.minX) * Math.max(0, b.maxY - b.minY);
  const clipArea = Math.max(0, clip.maxX - clip.minX) * Math.max(0, clip.maxY - clip.minY);
  if (clipArea >= pathArea) return v;
  return { ...v, clip: undefined, segs: clip.segs };
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
