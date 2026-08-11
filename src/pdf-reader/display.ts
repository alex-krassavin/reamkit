// §14.11.1 `/Rotate` — the page as it is SHOWN, not as it is stored.
//
// A page may be authored one way up and turned another for display: a landscape
// sheet drawn sideways inside a portrait `/MediaBox`, with `/Rotate 270` to
// stand it up again. Brotli-Prototype-FileA.pdf is twenty-five of them, and read
// without the turn every one came back portrait with its words running down the
// page.
//
// The turn belongs at the edge of the reader, not threaded through everything
// downstream: the marks are mapped into the page's own upright frame as they are
// lifted, and the rest of the reconstruction sees a page that was never turned.

import type { PathSeg, TextRun } from './content';
import type { PdfImage } from './images';
import type { PdfPage } from './document';
import type { PdfVector } from './vector';

/** A page's shown geometry: how big it is, and how to place a mark on it. */
export interface Display {
  /** The shown page's width in points — the `/MediaBox`'s, swapped on a quarter turn. */
  readonly width: number;
  /** The shown page's height in points. */
  readonly height: number;
  /** How far every mark turns with the page, degrees counter-clockwise. */
  readonly turnDeg: number;
  /** Map a `/MediaBox`-space point into the shown page's own y-up frame. */
  readonly place: (x: number, y: number) => { x: number; y: number };
}

/**
 * The {@link Display} a page's shown box and `/Rotate` describe.
 *
 * `/Rotate` turns the page CLOCKWISE when shown, so the content turns with it —
 * counter-clockwise by the same amount as seen from the content's own frame,
 * which is what `turnDeg` states.
 *
 * @param page The page whose shown geometry is wanted.
 * @returns Its size, its turn, and the map onto it.
 */
export function displayOf(page: PdfPage): Display {
  // §14.11.2 — what a viewer SHOWS is the crop box, which is the media box
  // where the page states no other.
  const [x0, y0, x1, y1] = page.cropBox;
  const left = Math.min(x0, x1);
  const bottom = Math.min(y0, y1);
  const w = Math.abs(x1 - x0);
  const h = Math.abs(y1 - y0);
  const quarter = page.rotate === 90 || page.rotate === 270;
  const place = (x: number, y: number): { x: number; y: number } => {
    const mx = x - left;
    const my = y - bottom;
    switch (page.rotate) {
      case 90:
        return { x: my, y: w - mx };
      case 180:
        return { x: w - mx, y: h - my };
      case 270:
        return { x: h - my, y: mx };
      default:
        return { x: mx, y: my };
    }
  };
  return {
    width: quarter ? h : w,
    height: quarter ? w : h,
    turnDeg: page.rotate === 90 ? -90 : page.rotate === 180 ? 180 : page.rotate === 270 ? 90 : 0,
    place,
  };
}

/**
 * The same runs, placed on the shown page.
 *
 * @param runs The runs as the content stream drew them.
 * @param d    The page's shown geometry.
 * @returns The runs in the shown page's frame, each carrying the page's turn.
 */
export function placeRuns(runs: ReadonlyArray<TextRun>, d: Display): Array<TextRun> {
  return runs.map((r) => {
    const origin = d.place(r.x, r.y);
    const end = d.place(r.endX, r.endY);
    const angle = (((r.angleDeg ?? 0) + d.turnDeg) % 360) + 0;
    const { angleDeg: _was, ...rest } = r;
    return {
      ...rest,
      x: origin.x,
      y: origin.y,
      endX: end.x,
      endY: end.y,
      ...(Math.abs(angle) > 0.5 ? { angleDeg: angle } : {}),
    };
  });
}

/**
 * The same pictures, placed on the shown page. A quarter turn swaps a picture's
 * width and height and leaves it standing on its side, which the caller carries
 * as the picture's own rotation.
 *
 * @param images The pictures as the content stream placed them.
 * @param d      The page's shown geometry.
 * @returns The pictures in the shown page's frame.
 */
export function placeImages(images: ReadonlyArray<PdfImage>, d: Display): Array<PdfImage> {
  return images.map((img) => {
    // The picture's own rectangle, mapped corner to corner: the corners are what
    // the turn moves, and the box is what is left of them.
    const corners = [
      d.place(img.x, img.y),
      d.place(img.x + img.widthPt, img.y),
      d.place(img.x + img.widthPt, img.y + img.heightPt),
      d.place(img.x, img.y + img.heightPt),
    ];
    const xs = corners.map((c) => c.x);
    const ys = corners.map((c) => c.y);
    const minX = Math.min(...xs);
    const minY = Math.min(...ys);
    // The page's own turn adds to the picture's.
    const turned = (img.rotationDeg ?? 0) + d.turnDeg;
    return {
      ...img,
      x: minX,
      y: minY,
      widthPt: Math.max(...xs) - minX,
      heightPt: Math.max(...ys) - minY,
      ...(Math.abs(turned) > 0.5 ? { rotationDeg: turned } : {}),
    };
  });
}

/**
 * The same paths, placed on the shown page. A path is geometry, so every point
 * of it moves and the bounding box is taken again from what comes out.
 *
 * @param vectors The paths as the content stream painted them.
 * @param d       The page's shown geometry.
 * @returns The paths in the shown page's frame.
 */
export function placeVectors(vectors: ReadonlyArray<PdfVector>, d: Display): Array<PdfVector> {
  return vectors.map((v) => {
    const segs: Array<PathSeg> = v.segs.map((s): PathSeg => {
      switch (s.op) {
        case 'move':
        case 'line': {
          const p = d.place(s.x, s.y);
          return { op: s.op, x: p.x, y: p.y };
        }
        case 'cubic': {
          const c1 = d.place(s.x1, s.y1);
          const c2 = d.place(s.x2, s.y2);
          const p = d.place(s.x, s.y);
          return { op: 'cubic', x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: p.x, y: p.y };
        }
        case 'close':
          return s;
      }
    });
    const xs: Array<number> = [];
    const ys: Array<number> = [];
    for (const s of segs) {
      if (s.op === 'close') continue;
      xs.push(s.x);
      ys.push(s.y);
      if (s.op === 'cubic') (xs.push(s.x1, s.x2), ys.push(s.y1, s.y2));
    }
    if (xs.length === 0) return { ...v, segs };
    return {
      ...v,
      segs,
      minX: Math.min(...xs),
      minY: Math.min(...ys),
      maxX: Math.max(...xs),
      maxY: Math.max(...ys),
    };
  });
}
