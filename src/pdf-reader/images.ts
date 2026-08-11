// E-PDF EP6 — lift the raster images off a page. Runs the content interpreter
// (EP2) for its `Do` placements, resolves each name against the page's
// /Resources /XObject, and either decodes an /Image (image-decode.ts) or
// recurses into a /Form XObject (composing its /Matrix onto the placement CTM,
// depth-guarded). Each surviving image carries its page-space rectangle (from
// the CTM that maps the unit square) and the enclosing structure id, so the
// tagged path can attach it to a /Figure and the heuristic path can order it by
// position. Unsupported images become losses rather than broken pictures.

import { interpretContent, multiply } from './content';
import { decodePdfImage } from './image-decode';
import { collectPageAppearances } from './annots';
import { buildFonts } from './text';
import { hiddenProperties, hiddenXObject } from './optional-content';
import { buildAlphaMap, sampledShading } from './shading';
import type { ClipRegion, ContentFont, Matrix } from './content';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { ImageCrop } from '@/core/document-model/types';
import type { Loss } from '@/core/ir';

import type { PdfFile, PdfPage } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';
import { encodePng } from '@/core/png-encode';
import { FEATURES } from '@/core/ir';

/**
 * One raster image lifted off a page (E-PDF EP6): the standalone image file
 * (`png`/`jpeg`/`jpeg2000`), its page-space rectangle (computed from the CTM
 * that maps the unit square) and the enclosing marked-content id so the tagged
 * path can attach it to a `/Figure`.
 */
export interface PdfImage {
  readonly bytes: Uint8Array;
  readonly format: 'png' | 'jpeg' | 'jpeg2000';
  /** Display size in page points (from the CTM). */
  readonly widthPt: number;
  readonly heightPt: number;
  /** Page-space lower-left corner (points, y-up). */
  readonly x: number;
  readonly y: number;
  /**
   * §8.9.5 — how far the CTM turns the picture, in degrees counter-clockwise.
   * The box above is the picture's own, unturned; a turn spins it about its
   * centre, which is what every downstream format does with one.
   */
  readonly rotationDeg?: number;
  /**
   * §8.5.4 — the fraction of each of the picture's OWN edges a clip cut away,
   * where one bounded it: `a:srcRect` in DrawingML terms. Absent is whole.
   */
  readonly crop?: ImageCrop;
  /** Enclosing marked-content id, if the placement was inside a `/Figure`. */
  readonly mcid?: number;
  /**
   * §8.5.3 — where this was painted, as the chain of positions leading to it.
   * The same key a lifted path carries, so the two can be ordered against each
   * other: a picture drawn over a filled box has the larger key.
   */
  readonly orderKey: ReadonlyArray<number>;
}

/** The images lifted off one page plus any losses for images that could not be reconstructed. */
export interface PageImages {
  readonly images: Array<PdfImage>;
  readonly losses: Array<Loss>;
}

const MAX_FORM_DEPTH = 12;
const MAX_IMAGES = 4096; // per-page DoS guard

/**
 * Lift the raster images off a page (E-PDF EP6). Runs the content interpreter
 * (EP2) for its `Do` placements, resolves each name against the page's
 * `/Resources` `/XObject`, and either decodes an `/Image` (via `decodePdfImage`)
 * or recurses into a `/Form` XObject — composing the form's `/Matrix` onto the
 * placement CTM, depth-guarded against cyclic forms. Each surviving image
 * carries its page-space rectangle and enclosing structure id. Unsupported
 * images become {@link Loss} entries rather than broken pictures.
 */
export function collectPageImages(file: PdfFile, page: PdfPage): PageImages {
  const images: Array<PdfImage> = [];
  const lossByDetail = new Map<string, Loss>();
  const visiting = new Set<PdfStream>();
  const alphaCache = new Map<PdfDict | undefined, ReturnType<typeof buildAlphaMap>>();
  // §9.6.5 — the fonts have to be built here, empty though this pass's interest
  // in text is: without them the interpreter cannot know a face is Type 3, and
  // a Type 3 glyph is a content stream that may paint a PICTURE. Passed
  // `NO_FONTS`, this pass saw no glyph calls at all and a page set in a bitmap
  // Type 3 font came back with nothing on it. Cached per resource dictionary,
  // as the paint states are.
  const fontCache = new Map<PdfDict | undefined, ReadonlyMap<string, ContentFont>>();
  const fontsOf = (resources: PdfDict | undefined): ReadonlyMap<string, ContentFont> => {
    const had = fontCache.get(resources);
    if (had) return had;
    const made = buildFonts(file, resources);
    fontCache.set(resources, made);
    return made;
  };

  const addLoss = (severity: 'dropped' | 'degraded', detail: string): void => {
    if (!lossByDetail.has(detail)) {
      lossByDetail.set(detail, { severity, feature: FEATURES.images, detail });
    }
  };

  const walk = (
    resources: PdfDict | undefined,
    content: Uint8Array,
    baseCtm: Matrix,
    depth: number,
    inheritedMcid: number | undefined,
    prefix: ReadonlyArray<number>,
  ): void => {
    const xobjects = resources ? file.get(resources, 'XObject') : PDF_NULL;
    const xobjDict = xobjects instanceof Map ? xobjects : undefined;
    // §11.3.5 — `gs` resolves against the resources IN FORCE, and the only
    // thing wanted from it here is whether the page asked for a blend nothing
    // downstream can perform. Cached per dictionary, as the paths are.
    let paints = alphaCache.get(resources);
    if (!paints) {
      paints = buildAlphaMap(file, resources);
      alphaCache.set(resources, paints);
    }
    const result = interpretContent(
      content,
      fontsOf(resources),
      baseCtm,
      undefined,
      paints,
      undefined,
      hiddenProperties(file, resources),
    );

    // §8.7.3.1 — a path filled with a TILING pattern shows that pattern's own
    // content stream, drawn through the pattern's `/Matrix`. It is a call like
    // a form's, so it is walked like one: 22060_A1_01_Plans.pdf draws all four
    // of its floor plans this way, as one JPEG apiece inside a pattern, and
    // nothing else on the page refers to those images at all.
    const patterns = resources ? file.get(resources, 'Pattern') : PDF_NULL;
    const patternDict = patterns instanceof Map ? patterns : undefined;
    for (const vector of result.vectors) {
      if (vector.patternName === undefined || depth >= MAX_FORM_DEPTH) continue;
      const stream = patternDict
        ? file.resolve(patternDict.get(vector.patternName) ?? PDF_NULL)
        : PDF_NULL;
      if (!(stream instanceof PdfStream) || visiting.has(stream)) continue;
      // Type 1 only: a type-2 (shading) pattern is a gradient, and the vector
      // path already carries it.
      if (file.get(stream.dict, 'PatternType') !== 1) continue;
      visiting.add(stream);
      const patternRes = file.get(stream.dict, 'Resources');
      // A pattern paints where its FILL stands, so its marks take the fill's
      // place in the order.
      walk(
        patternRes instanceof Map ? patternRes : resources,
        file.streamData(stream),
        multiply(matrixOf(file, stream.dict), baseCtm),
        depth + 1,
        inheritedMcid,
        [...prefix, vector.order],
      );
      visiting.delete(stream);
    }

    // §9.6.5 — a Type 3 glyph is a content stream, and a bitmap font's glyph
    // is a PICTURE: an inline stencil mask, or a `Do` of one image apiece. The
    // vector pass has walked these since it learned about Type 3; this one
    // never did, so a page set in such a font came back with nothing on it at
    // all. french_diacritics.pdf draws each accented letter as a 40×59 stencil
    // inside its glyph, and bug1011159.pdf as one XObject per glyph.
    for (const call of result.glyphs) {
      if (depth >= MAX_FORM_DEPTH || visiting.has(call.stream)) continue;
      visiting.add(call.stream);
      walk(
        call.resources ?? resources,
        file.streamData(call.stream),
        call.ctm,
        depth + 1,
        inheritedMcid,
        [...prefix, call.order],
      );
      visiting.delete(call.stream);
    }

    // §8.7.4.3 / §8.7.4.5.3 — a bare `sh` paints the clip with a shading, and a
    // FUNCTION-BASED one is a function of two variables over a rectangle: no
    // gradient stands for it, but a picture does exactly.
    // function_based_shading.pdf is nine such squares and the whole page.
    for (const paint of result.shadings) {
      if (images.length >= MAX_IMAGES) return;
      const dict = resources ? file.get(resources, 'Shading') : PDF_NULL;
      const shading =
        dict instanceof Map ? file.resolve(dict.get(paint.name) ?? PDF_NULL) : PDF_NULL;
      const sh =
        shading instanceof PdfStream ? shading.dict : shading instanceof Map ? shading : undefined;
      if (!sh) continue;
      const sampled = sampledShading(file, sh);
      if (!sampled) continue;
      // §8.7.4.5.3 — the `/Matrix` on the shading maps its domain onto the
      // space the `sh` was painted in, and the CTM carries that to the page.
      const placed = multiply(matrixOf(file, sh), paint.ctm);
      const [dx0, dx1, dy0, dy1] = sampled.domain;
      const unit: Matrix = [dx1 - dx0, 0, 0, dy1 - dy0, dx0, dy0];
      images.push({
        ...geometry(
          multiply(unit, placed),
          {
            bytes: encodePng(sampled.size, sampled.size, 'rgb', sampled.rgb),
            format: 'png',
          },
          inheritedMcid,
          paint.clip,
        ),
        orderKey: [...prefix, paint.order],
      });
    }

    for (const placement of result.images) {
      if (images.length >= MAX_IMAGES) return;
      // §8.9.7 — an inline image names no resource: it IS the resource, written
      // into the stream. Wrapped as one, it decodes down the same path.
      if (placement.inline) {
        const decoded = decodePdfImage(
          file,
          new PdfStream(placement.inline.dict, placement.inline.data),
          placement.fillHex,
        );
        if (decoded.ok) {
          images.push({
            ...geometry(placement.ctm, decoded, placement.mcid ?? inheritedMcid, placement.clip),
            orderKey: [...prefix, placement.order],
          });
          if (decoded.degraded) addLoss('degraded', decoded.degraded);
        } else {
          addLoss(decoded.severity, decoded.detail);
        }
        continue;
      }
      const stream = xobjDict ? file.resolve(xobjDict.get(placement.name) ?? PDF_NULL) : PDF_NULL;
      if (!(stream instanceof PdfStream)) continue;
      // §8.11.3.1 — an XObject may carry its own `/OC`, and a hidden one is
      // not painted at all.
      if (hiddenXObject(file, stream)) continue;
      const subtype = nameOf(file.get(stream.dict, 'Subtype'));
      const mcid = placement.mcid ?? inheritedMcid;
      if (subtype === 'Image') {
        const decoded = decodePdfImage(file, stream, placement.fillHex);
        if (decoded.ok) {
          images.push({
            ...geometry(placement.ctm, decoded, mcid, placement.clip),
            orderKey: [...prefix, placement.order],
          });
          if (decoded.degraded) addLoss('degraded', decoded.degraded);
          // §11.3.5 — the picture is here, but the RULE that was to mix it with
          // what lies under it is not: no anchored picture blends, in any
          // format this writes. blendmode.pdf lays a second photograph over a
          // dog in each of sixteen cells, one blend to a cell, and every one of
          // them came back as the dog alone with nothing said about it.
          if (placement.blend !== undefined) {
            addLoss(
              'degraded',
              `PDF blend mode /${placement.blend} is not performed; the picture is drawn over what it was to blend with`,
            );
          }
          // §11.6.5 — a mask that fades the paint from place to place. No
          // anchored picture has one, so the picture is drawn at full strength.
          if (placement.masked === true) {
            addLoss(
              'degraded',
              'PDF soft mask (/SMask in the graphics state) is not applied; the picture is drawn at full opacity throughout',
            );
          }
        } else {
          addLoss(decoded.severity, decoded.detail);
        }
      } else if (subtype === 'Form' && depth < MAX_FORM_DEPTH && !visiting.has(stream)) {
        visiting.add(stream);
        const formRes = file.get(stream.dict, 'Resources');
        walk(
          formRes instanceof Map ? formRes : resources,
          file.streamData(stream),
          multiply(matrixOf(file, stream.dict), placement.ctm),
          depth + 1,
          mcid,
          [...prefix, placement.order],
        );
        visiting.delete(stream);
      }
    }
  };

  walk(page.resources, file.pageContent(page), [1, 0, 0, 1, 0, 0], 0, undefined, []);
  // §12.5.5 — the same for a widget's appearance: a scanned signature or a
  // field's icon lives there and nowhere in the page's own stream.
  collectPageAppearances(file, page).forEach((appearance, index) => {
    walk(
      appearance.resources ?? page.resources,
      file.streamData(appearance.stream),
      appearance.ctm,
      1,
      undefined,
      [Number.MAX_SAFE_INTEGER, index],
    );
  });
  return { images, losses: [...lossByDetail.values()] };
}

function geometry(
  ctm: Matrix,
  decoded: { bytes: Uint8Array; format: 'png' | 'jpeg' | 'jpeg2000' },
  mcid: number | undefined,
  clip: ClipRegion | undefined,
): Omit<PdfImage, 'orderKey'> {
  const widthPt = Math.hypot(ctm[0], ctm[1]) || 1;
  const heightPt = Math.hypot(ctm[2], ctm[3]) || 1;
  // §8.9.5 — a picture is placed by the CTM, and the CTM may TURN it. Taking
  // only the column lengths and the translation threw the turn away and put
  // the picture square at the origin of its own space:
  // image-rotated-black-white-ratio.pdf sets its picture at forty degrees in
  // the middle of the page and it came back upright in the corner.
  const angle = (Math.atan2(ctm[1], ctm[0]) * 180) / Math.PI;
  const box = clippedUnitBox(ctm, clip);
  // The centre is what a turn leaves in place, so the box is measured off it:
  // the shown part of the unit square, through the matrix.
  const u = (box.u0 + box.u1) / 2;
  const v = (box.v0 + box.v1) / 2;
  const cx = ctm[0] * u + ctm[2] * v + ctm[4];
  const cy = ctm[1] * u + ctm[3] * v + ctm[5];
  const shownW = widthPt * (box.u1 - box.u0);
  const shownH = heightPt * (box.v1 - box.v0);
  // §20.1.8.55 — what the clip cut away, as the fraction of each edge. The unit
  // square's v runs UP and an image's rows run DOWN from its first, so the top
  // is what is above v1.
  const crop =
    box.u0 > 0 || box.v0 > 0 || box.u1 < 1 || box.v1 < 1
      ? { left: box.u0, right: 1 - box.u1, top: 1 - box.v1, bottom: box.v0 }
      : undefined;
  return {
    bytes: decoded.bytes,
    format: decoded.format,
    widthPt: shownW,
    heightPt: shownH,
    x: cx - shownW / 2,
    y: cy - shownH / 2,
    ...(crop ? { crop } : {}),
    ...(Math.abs(angle) > 0.5 ? { rotationDeg: angle } : {}),
    ...(mcid !== undefined ? { mcid } : {}),
  };
}

/** Below this the clip took a sliver off an edge and is not worth a crop. */
const CROPS = 0.01;

/**
 * §8.5.4 — the part of the unit square a clip leaves showing, in the PICTURE's
 * own axes.
 *
 * A clip is stated in the page's coordinates and a picture may be turned within
 * them, so intersecting the two boxes as they stand crops along the wrong axes.
 * Carried back through the placement matrix the clip lands in the unit square
 * the image is drawn into, where a crop is what `a:srcRect` means: the fraction
 * off each of the picture's own edges. image-rotated-black-white-ratio.pdf
 * turns picture and clip together by thirty-one degrees, and in that space the
 * clip is square on and takes the middle 53% of both sides.
 *
 * Where the clip is turned differently from the picture this bounds it rather
 * than cutting it exactly — a rectangle is all `a:srcRect` can say.
 */
function clippedUnitBox(
  ctm: Matrix,
  clip: ClipRegion | undefined,
): { u0: number; v0: number; u1: number; v1: number } {
  const whole = { u0: 0, v0: 0, u1: 1, v1: 1 };
  if (!clip) return whole;
  const det = ctm[0] * ctm[3] - ctm[1] * ctm[2];
  if (!Number.isFinite(det) || Math.abs(det) < 1e-9) return whole;
  let u0 = Infinity;
  let v0 = Infinity;
  let u1 = -Infinity;
  let v1 = -Infinity;
  const add = (x: number, y: number): void => {
    const dx = x - ctm[4];
    const dy = y - ctm[5];
    const u = (ctm[3] * dx - ctm[2] * dy) / det;
    const v = (ctm[0] * dy - ctm[1] * dx) / det;
    u0 = Math.min(u0, u);
    v0 = Math.min(v0, v);
    u1 = Math.max(u1, u);
    v1 = Math.max(v1, v);
  };
  for (const seg of clip.segs) {
    if (seg.op === 'close') continue;
    if (seg.op === 'cubic') {
      // The hull bounds the curve, which is what a bounding crop wants.
      add(seg.x1, seg.y1);
      add(seg.x2, seg.y2);
    }
    add(seg.x, seg.y);
  }
  if (!Number.isFinite(u0) || !Number.isFinite(v0)) return whole;
  const cut = {
    u0: Math.min(1, Math.max(0, u0)),
    v0: Math.min(1, Math.max(0, v0)),
    u1: Math.min(1, Math.max(0, u1)),
    v1: Math.min(1, Math.max(0, v1)),
  };
  // A clip that leaves nothing is a clip this has read wrong — draw it whole
  // rather than drop it to a hairline.
  if (!(cut.u1 - cut.u0 > CROPS) || !(cut.v1 - cut.v0 > CROPS)) return whole;
  if (cut.u1 - cut.u0 > 1 - CROPS && cut.v1 - cut.v0 > 1 - CROPS) return whole;
  return cut;
}

function matrixOf(file: PdfFile, dict: PdfDict): Matrix {
  const m = file.resolve(dict.get('Matrix') ?? PDF_NULL);
  if (Array.isArray(m) && m.length >= 6 && m.every((v) => typeof v === 'number')) {
    return [m[0], m[1], m[2], m[3], m[4], m[5]] as Matrix;
  }
  return [1, 0, 0, 1, 0, 0];
}

function nameOf(v: PdfValue | undefined): string {
  return v instanceof PdfName ? v.value : '';
}
