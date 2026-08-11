// E-PDF EP10/EP11/EP16c — lift painted vector paths off a page. Runs the content
// interpreter for its fill (EP10), stroke (EP11) and shading-pattern (EP16c)
// placements and keeps the ones that read as real graphics: it drops
// hairline/degenerate fills, invisible white paint, short stroke specks, and the
// near-full-page background, so a reconstructed document gains genuine coloured
// shapes, lines and gradients without the dot / page-background clutter. Clips
// and the bare `sh` operator are not captured (a documented loss).

import { IDENTITY, interpretContent, multiply } from './content';
import {
  buildAlphaMap,
  buildColorSpaceMap,
  buildShadingMap,
  gradientShading,
  shadingTypeOf,
} from './shading';
import { collectPageAppearances } from './annots';
import { hiddenProperties, hiddenXObject } from './optional-content';
import { buildFonts } from './text';
import type { ColorSpaceInfo, GsPaint } from './shading';
import type {
  ContentFont,
  ImagePlacement,
  Matrix,
  PathSeg,
  Type3Call,
  VectorPlacement,
} from './content';
import type { ShapeGradient } from '@/core/vector';
import type { Loss } from '@/core/ir';
import type { PdfDict } from '@/pdf/objects';

import type { PdfFile, PdfPage } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';
import { FEATURES } from '@/core/ir';

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
/** The name-keyed state a resource dictionary supplies to the interpreter. */
interface ResourceMaps {
  readonly shadings: ReadonlyMap<string, ShapeGradient>;
  readonly alphas: ReadonlyMap<string, GsPaint>;
  readonly spaces: ReadonlyMap<string, ColorSpaceInfo>;
}

function paintedVectors(
  file: PdfFile,
  page: PdfPage,
): {
  placements: Array<VectorPlacement & { orderKey: ReadonlyArray<number> }>;
  /** §8.7.4.3 — how many `sh` regions the page painted that this does not lift. */
  bareShadings: number;
} {
  const out: Array<VectorPlacement & { orderKey: ReadonlyArray<number> }> = [];
  let bareShadings = 0;
  const visiting = new Set<PdfStream>();
  // §7.8.3 — a name resolves against the resources IN FORCE, which is the
  // form's or the appearance's own where it has one. annotation-highlight.pdf
  // keeps its `/Multiply` in the appearance's `/ExtGState` and nowhere else,
  // and read against the page's it was never found at all. Cached per
  // dictionary, since a page's forms nearly all share one.
  const stateCache = new Map<PdfDict | undefined, ResourceMaps>();
  const mapsOf = (resources: PdfDict | undefined): ResourceMaps => {
    const had = stateCache.get(resources);
    if (had) return had;
    const made: ResourceMaps = {
      shadings: buildShadingMap(file, resources),
      alphas: buildAlphaMap(file, resources),
      spaces: buildColorSpaceMap(file, resources),
    };
    stateCache.set(resources, made);
    return made;
  };
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
    const maps = mapsOf(resources);
    const result = interpretContent(
      content,
      buildFonts(file, resources),
      baseCtm,
      maps.shadings,
      maps.alphas,
      maps.spaces,
      hiddenProperties(file, resources),
    );
    // §8.7.4.3 — a bare `sh` paints the CLIP with a shading. Where that shading
    // is axial or radial it is a gradient, and the region it fills is the clip:
    // 131 of the corpus's `sh` paints are axial, ten times as many as the
    // function-based ones, and every one of them was dropped.
    for (const paint of result.shadings) {
      if (out.length >= MAX_VECTORS) break;
      const dict = resources ? file.get(resources, 'Shading') : PDF_NULL;
      const found = dict instanceof Map ? file.resolve(dict.get(paint.name) ?? PDF_NULL) : PDF_NULL;
      const sh = found instanceof PdfStream ? found.dict : found instanceof Map ? found : undefined;
      // §8.7.4.5.3 — a function-based shading is a PICTURE, and the image pass
      // draws it. Counted here it was a loss the file did not have.
      if (sh && shadingTypeOf(file, sh) === 1) continue;
      const gradient = sh ? gradientShading(file, sh) : undefined;
      // §11.6.5 — under a soft mask the clip is not the extent: the MASK is,
      // and nothing here applies one. bug1721218_reduced.pdf fades a shadow out
      // under the router it draws, and painted to its clip that shadow arrived
      // as a solid black blob beside the picture.
      if (paint.masked) {
        bareShadings++;
        continue;
      }
      if (!sh || !gradient || !paint.clip) {
        bareShadings++;
        continue;
      }
      // The clip is the whole of what the paint covers, so the clip's own path
      // IS the shape. Without one the region is the page, which is a guess this
      // does not make.
      out.push({
        order: paint.order,
        segs: paint.clip.segs,
        gradient,
        // §11.6.4.4 — alphatrans.pdf paints its gradient at half opacity over
        // three coloured squares, and drawn solid it buried all three.
        ...(paint.alpha !== undefined ? { alpha: paint.alpha } : {}),
        ...(paint.darkens ? { darkens: true } : {}),
        orderKey: [...prefix, paint.order],
      });
    }

    // §8.5.3 — later marks cover earlier ones, and a form is drawn where its
    // `Do` stands, not after everything around it. Walking the stream first and
    // its forms afterwards puts every form on top: 22060_A1_01_Plans.pdf backs
    // its legend with a white box inside a form, and hoisted to the end that box
    // covered the legend's own words.
    const events: Array<{
      order: number;
      vector?: VectorPlacement;
      xobject?: ImagePlacement;
      glyph?: Type3Call;
    }> = [
      ...result.vectors.map((vector) => ({ order: vector.order, vector })),
      ...result.images.map((xobject) => ({ order: xobject.order, xobject })),
      ...result.glyphs.map((glyph) => ({ order: glyph.order, glyph })),
    ].sort((a, b) => a.order - b.order);

    for (const event of events) {
      if (out.length >= MAX_VECTORS) return;
      if (event.vector) {
        out.push({ ...event.vector, orderKey: [...prefix, event.order] });
        continue;
      }
      // §9.6.5 — a Type 3 glyph is a content stream, and what it paints is
      // what the page shows where that character stands.
      if (event.glyph) {
        const call = event.glyph;
        if (depth >= MAX_FORM_DEPTH || visiting.has(call.stream)) continue;
        visiting.add(call.stream);
        walk(call.resources ?? resources, file.streamData(call.stream), call.ctm, depth + 1, [
          ...prefix,
          call.order,
        ]);
        visiting.delete(call.stream);
        continue;
      }
      const placement = event.xobject!;
      if (depth >= MAX_FORM_DEPTH) continue;
      const stream = xobjDict ? file.resolve(xobjDict.get(placement.name) ?? PDF_NULL) : PDF_NULL;
      if (!(stream instanceof PdfStream) || visiting.has(stream)) continue;
      // §8.11.3.1 — a form may carry its own `/OC`, and a hidden one paints
      // nothing at all.
      if (hiddenXObject(file, stream)) continue;
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
  return { placements: out, bareShadings };
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
  /** §11.6.4.4 — how opaque the fill is, when the page asked for less than all. */
  readonly alpha?: number;
  /**
   * §11.3.5 — the fill only DARKENS what it covers, so the marks under it read
   * through. A highlighter is this, and nothing on a page reads it as paint.
   */
  readonly darkens?: boolean;
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
// What counts as a mark rather than a speck.
//
// These were set to drop dots and hairlines, and they dropped LETTERS: a
// program that draws its lettering as outlines writes an `I` at six points as a
// mark 0.8pt wide and 4pt tall, which was under every one of the old bounds
// (2pt, 16pt², 6pt). Brotli-Prototype-FileA.pdf labels its floor plans that way
// and they came back reading "V G ROOM" and "D N NG" — the wide letters kept,
// the narrow ones thrown out as clutter.
const MIN_SIDE = 0.5; // pt — thinner than a hairline rule
const MIN_AREA = 2; // pt² — smaller than the smallest letter
const MIN_STROKE_LEN = 6; // pt — skip stroke specks (tick marks, dots)
const MIN_RULE_LEN = 2; // pt — a filled rule shorter than this is a speck
/**
 * How many painted paths a page may hand over, as a guard against a file built
 * to exhaust a reader.
 *
 * This was two thousand, and it was not a guard, it was a silent truncation:
 * the busiest page of Brotli-Prototype-FileA.pdf keeps 1994 paths AFTER the
 * de-cluttering filter, so the cap cut the page off partway and took its whole
 * title block, the vegetation of its perspective and every hatch in its legend
 * with it — with nothing in the report to say a thing was missing. Twenty
 * thousand leaves room for a drawing and still bounds the work (that file's
 * twenty-five sheets read in 420ms); reaching it is now reported.
 */
const MAX_VECTORS = 20000;
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
): PageVectors {
  const [px0, py0, px1, py1] = page.cropBox;
  const pageArea = Math.max(1, Math.abs((px1 - px0) * (py1 - py0)));
  const out: Array<PdfVector> = [];
  // What the page has painted so far. White paint is invisible only over white:
  // over anything else it is the thing that HIDES it, and the caller seeds this
  // with the pictures it has already placed.
  const painted: Array<Box> = [...occupied];
  const lifted = paintedVectors(file, page);
  const raws = lifted.placements;
  // §11.6.5 — a mask that fades the paint from place to place, which no shape
  // downstream has. bug852992_reduced.pdf fades both its green ground and the
  // orange box on it toward the edges, and both came back flat with nothing
  // said. §11.3.5 — and a blend rule nothing here performs, likewise.
  const losses: Array<Loss> = [];
  const bareShadings = lifted.bareShadings;
  const asked = new Set<string>();
  for (const raw of raws) {
    if (raw.masked === true) {
      asked.add(
        'PDF soft mask (/SMask in the graphics state) is not applied; the shape is drawn at full opacity throughout',
      );
    }
    if (raw.blend !== undefined) {
      asked.add(
        `PDF blend mode /${raw.blend} is not performed; the shape is drawn over what it was to blend with`,
      );
    }
  }
  for (const detail of asked)
    losses.push({ severity: 'degraded', feature: FEATURES.images, detail });
  for (const raw of raws) {
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
    // No cap on the AREA. A fill the size of the page is a page background, and
    // a page that is pale blue is pale blue: filled-background.pdf is nothing
    // but that fill, and it came back a blank sheet. bug1755507.pdf and
    // bug946506.pdf are a page of blue with a card on it, XiaoBiaoSong.pdf and
    // SimFang-variant.pdf a page of grey — eight files across the corpus, every
    // one of them better for it and not one worse. What keeps a background from
    // burying the words is not its size but where it sits: the flowing reading
    // puts every mark behind the text, and the placed one keeps the page's own
    // painting order.
    const filled = (v.gradient !== undefined || solidFill) && (isBox || isRule);
    // White paint is invisible only over white — the same rule the fill above
    // follows, which the stroke did not. 160F-2019.pdf's every form field is a
    // tinted box with a WHITE one-point border stroked inside it, and dropping
    // the border left each field a point wider on every side than the file
    // draws it, seventy-six times over.
    const stroked =
      v.strokeHex !== undefined &&
      (v.strokeHex !== 'FFFFFF' || painted.some((box) => overlaps(box, b))) &&
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
      ...(filled && v.alpha !== undefined ? { alpha: v.alpha } : {}),
      ...(filled && v.darkens === true ? { darkens: true } : {}),
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
  // A cap that says nothing is a cap that reads as "the page had no more".
  const cut = out.length >= MAX_VECTORS || raws.length >= MAX_VECTORS;
  if (cut) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.shapes,
      detail: `page carries more than ${String(MAX_VECTORS)} painted paths; the rest were not read`,
    });
  }
  // §8.7.4.3 — a bare `sh` paints the clip region rather than filling a path,
  // and nothing here lifts it. Said where it HAPPENED: reported for every
  // document instead, it fired on all four hundred files of the pdf.js corpus,
  // most of which contain no `sh` at all, and a loss report that cries wolf on
  // every document tells a reader nothing.
  if (bareShadings > 0) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.images,
      detail: `${String(bareShadings)} bare-shading (sh) region${bareShadings === 1 ? '' : 's'} painted the clip rather than a path, and ${bareShadings === 1 ? 'was' : 'were'} not reconstructed`,
    });
  }
  return { vectors: out, losses };
}

/** The paths lifted off one page, plus a loss if the page ran past the cap. */
export interface PageVectors {
  readonly vectors: Array<PdfVector>;
  readonly losses: ReadonlyArray<Loss>;
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
