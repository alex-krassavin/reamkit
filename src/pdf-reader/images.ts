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
import type { ContentFont, Matrix } from './content';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { Loss } from '@/core/ir';

import type { PdfFile, PdfPage } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';
import { FEATURES } from '@/core/ir';

export interface PdfImage {
  readonly bytes: Uint8Array;
  readonly format: 'png' | 'jpeg' | 'jpeg2000';
  readonly widthPt: number; // display size in page points (from the CTM)
  readonly heightPt: number;
  readonly x: number; // page-space lower-left corner (points, y-up)
  readonly y: number;
  readonly mcid?: number;
}

export interface PageImages {
  readonly images: Array<PdfImage>;
  readonly losses: Array<Loss>;
}

const NO_FONTS: ReadonlyMap<string, ContentFont> = new Map();
const MAX_FORM_DEPTH = 12;
const MAX_IMAGES = 4096; // per-page DoS guard

export function collectPageImages(file: PdfFile, page: PdfPage): PageImages {
  const images: Array<PdfImage> = [];
  const lossByDetail = new Map<string, Loss>();
  const visiting = new Set<PdfStream>();

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
  ): void => {
    const xobjects = resources ? file.get(resources, 'XObject') : PDF_NULL;
    const xobjDict = xobjects instanceof Map ? xobjects : undefined;
    for (const placement of interpretContent(content, NO_FONTS, baseCtm).images) {
      if (images.length >= MAX_IMAGES) return;
      const stream = xobjDict ? file.resolve(xobjDict.get(placement.name) ?? PDF_NULL) : PDF_NULL;
      if (!(stream instanceof PdfStream)) continue;
      const subtype = nameOf(file.get(stream.dict, 'Subtype'));
      const mcid = placement.mcid ?? inheritedMcid;
      if (subtype === 'Image') {
        const decoded = decodePdfImage(file, stream);
        if (decoded.ok) {
          images.push(geometry(placement.ctm, decoded, mcid));
          if (decoded.degraded) addLoss('degraded', decoded.degraded);
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
        );
        visiting.delete(stream);
      }
    }
  };

  walk(page.resources, file.pageContent(page), [1, 0, 0, 1, 0, 0], 0, undefined);
  return { images, losses: [...lossByDetail.values()] };
}

function geometry(
  ctm: Matrix,
  decoded: { bytes: Uint8Array; format: 'png' | 'jpeg' | 'jpeg2000' },
  mcid: number | undefined,
): PdfImage {
  return {
    bytes: decoded.bytes,
    format: decoded.format,
    widthPt: Math.hypot(ctm[0], ctm[1]) || 1,
    heightPt: Math.hypot(ctm[2], ctm[3]) || 1,
    x: ctm[4],
    y: ctm[5],
    ...(mcid !== undefined ? { mcid } : {}),
  };
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
