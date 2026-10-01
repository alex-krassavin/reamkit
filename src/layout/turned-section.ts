// §17.6.20 `w:textDirection` — a section whose lines run DOWN the sheet.
//
// `tbRl` turns the section's text a quarter clockwise: each line runs top to
// bottom and the next stands to the left of it. It is how a viewer shows a page
// turned by `/Rotate 90` whose words stood upright in its box, and how Word lays
// a section out for "rotate all text 90°" — for `btLr` as well, which it sets
// the same way.
//
// Word turns the TEXT and nothing else: the header and the footer stay across
// the top and the foot of the sheet, and a drawing anchored to the page stands
// where its offsets put it on the sheet, upright. So the body is laid out in the
// frame the text reads in — the sheet turned back a quarter, as wide as the
// sheet is tall — and every item of it is then turned onto the sheet, while the
// bands are laid out on the sheet as they always are. A drawing's anchor is
// given the opposite turn before the layout sees it, so that turning the body
// puts it back where Word does.

import type { BodyElement, FloatAnchor } from '@/core/document-model';
import type { PageItem } from './page-doc';
import { pt } from '@/core/ir';

/** The sheet a turned section prints on, and the margins it states for it. */
export interface Sheet {
  readonly width: number;
  readonly height: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
}

/** A quarter turn in the 1/60000° a DrawingML rotation is stated in (§20.1.7.6). */
const QUARTER_60K = 90 * 60000;

/**
 * One item of a turned section's body, carried from the frame the text reads
 * in onto the sheet: a point `(x, y)` of the frame (top-left origin, y down)
 * lands at `(sheetWidth − y, x)` — the frame's top edge is the sheet's right
 * edge, and its left edge the sheet's top.
 *
 * @param item       The item as the frame's layout placed it.
 * @param sheetWidth The sheet's width, which is the frame's height.
 * @returns The same item on the sheet, turned a quarter clockwise.
 */
export function turnOntoSheet(item: PageItem, sheetWidth: number): PageItem {
  const box = (x: number, y: number, w: number, h: number) => ({
    x: pt(sheetWidth - y - h),
    y: pt(x),
    width: pt(h),
    height: pt(w),
  });
  switch (item.type) {
    case 'line': {
      // The line turns about its own origin: counter-clockwise is how the
      // item states it, and a quarter clockwise is -90 of that.
      const quarter = withinTurn((item.rotationDeg ?? 0) - 90);
      const turned = quarter > 180 ? quarter - 360 : quarter;
      const { rotationDeg: _flat, clip, ...rest } = item;
      return {
        ...rest,
        originX: pt(sheetWidth - item.baselineY),
        baselineY: pt(item.originX),
        ...(turned !== 0 ? { rotationDeg: turned } : {}),
        ...(clip ? { clip: box(clip.x, clip.y, clip.width, clip.height) } : {}),
      };
    }
    case 'image': {
      // §20.1.7.6 — a picture turns about its own centre, clockwise, and its
      // box is the one it has before the turn.
      const cx = item.x + item.width / 2;
      const cy = item.y + item.height / 2;
      const turned = withinTurn((item.rotationDeg ?? 0) + 90);
      const { rotationDeg: _flat, clip, ...rest } = item;
      return {
        ...rest,
        x: pt(sheetWidth - cy - item.width / 2),
        y: pt(cx - item.height / 2),
        ...(turned !== 0 ? { rotationDeg: turned } : {}),
        ...(clip
          ? { clip: { paths: clip.paths, transform: turnMatrix(clip.transform, sheetWidth) } }
          : {}),
      };
    }
    case 'shape': {
      const shadow = item.shape.shadow;
      return {
        ...item,
        shape: {
          ...item.shape,
          transform: turnMatrix(item.shape.transform, sheetWidth),
          // A shadow's offset is measured in the page's own frame.
          ...(shadow ? { shadow: { ...shadow, dxPt: -shadow.dyPt, dyPt: shadow.dxPt } } : {}),
        },
      };
    }
    case 'border':
      // The box is the CELL's, and the side names an edge of it: the frame's
      // top edge is the sheet's right one, and so on round.
      return {
        ...item,
        ...box(item.x, item.y, item.width, item.height),
        side: TURNED_SIDE[item.side],
      };
    case 'fill':
      return { ...item, ...box(item.x, item.y, item.width, item.height) };
  }
}

/** An angle in degrees brought into one full turn, `[0, 360)`, with the float noise off. */
function withinTurn(deg: number): number {
  const t = ((deg % 360) + 360) % 360;
  return Math.abs(t) < 1e-9 || Math.abs(t - 360) < 1e-9 ? 0 : t;
}

/** Which edge of a box each edge becomes when the box turns a quarter clockwise. */
const TURNED_SIDE = {
  top: 'right',
  right: 'bottom',
  bottom: 'left',
  left: 'top',
} as const;

/**
 * A local→page matrix `[a b c d e f]` (x' = a·u + c·v + e, y' = b·u + d·v + f,
 * page top-left, y down) followed by the turn onto the sheet.
 */
function turnMatrix(
  m: readonly [number, number, number, number, number, number],
  sheetWidth: number,
): [number, number, number, number, number, number] {
  return [-m[1], m[0], -m[3], m[2], sheetWidth - m[5], m[4]];
}

/**
 * A drawing anchored in a turned section, given the place in the text's frame
 * from which turning the body carries it to where Word stands it: on the sheet,
 * at the offsets its anchor states, upright.
 *
 * The anchor is resolved against the SHEET — its page, its margins and the
 * bands beside them. What is measured from the text itself (a paragraph, a
 * line, a column) has no place on the sheet before the text is laid out, and is
 * measured from the margin instead. Only pictures and shapes are carried; a
 * chart or a floating table turns with the text.
 *
 * @param el    The body element, anchored or not.
 * @param sheet The sheet the section prints on.
 * @returns The element, its anchor restated in the frame and its turn backed off.
 */
export function floatIntoFrame(el: BodyElement, sheet: Sheet): BodyElement {
  if (el.kind === 'image' && el.image.float !== undefined) {
    const { image } = el;
    const float = image.float!;
    const placed = frameAnchor(float, image.width, image.height, sheet);
    return {
      ...el,
      image: { ...image, float: placed, rotation60k: backedOff(image.rotation60k) },
    };
  }
  if (el.kind === 'shape' && el.shape.float !== undefined) {
    const { shape } = el;
    const float = shape.float!;
    const placed = frameAnchor(float, shape.width, shape.height, sheet);
    return {
      ...el,
      shape: {
        ...shape,
        float: placed,
        transform: { ...shape.transform, rotation60k: backedOff(shape.transform?.rotation60k) },
      },
    };
  }
  return el;
}

/** A clockwise turn with a quarter taken off, kept in the one full turn. */
function backedOff(rotation60k: number | undefined): number {
  const full = 360 * 60000;
  return ((((rotation60k ?? 0) - QUARTER_60K) % full) + full) % full;
}

/**
 * The anchor a drawing of `w × h` needs in the frame for the body's turn to put
 * it where `float` stands it on the sheet.
 */
function frameAnchor(float: FloatAnchor, w: number, h: number, sheet: Sheet): FloatAnchor {
  // Where the anchor puts the drawing's box on the sheet (top-left, y down).
  const x = sheetX(float.posH, w, sheet);
  const y = sheetY(float.posV, h, sheet);
  // Its centre, carried back into the frame: the inverse of `turnOntoSheet`
  // takes a sheet point (X, Y) to (Y, sheetWidth − X).
  const cx = y + h / 2;
  const cy = sheet.width - (x + w / 2);
  const { posH: _h, posV: _v, ...rest } = float;
  return {
    ...rest,
    posH: { relativeFrom: 'page', offsetPt: pt(cx - w / 2) },
    posV: { relativeFrom: 'page', offsetPt: pt(cy - h / 2) },
  };
}

/** §20.4.3.3 — the box's left edge on the sheet. */
function sheetX(h: FloatAnchor['posH'], w: number, sheet: Sheet): number {
  const textRight = sheet.width - sheet.marginRight;
  const [base, span] =
    h?.relativeFrom === 'page'
      ? [0, sheet.width]
      : h?.relativeFrom === 'leftMargin'
        ? [0, sheet.marginLeft]
        : h?.relativeFrom === 'rightMargin'
          ? [textRight, sheet.marginRight]
          : [sheet.marginLeft, textRight - sheet.marginLeft];
  if (h?.align === 'center') return base + (span - w) / 2;
  if (h?.align === 'right') return base + span - w;
  return base + (h?.offsetPt ?? 0);
}

/** §20.4.3.4 — the box's top edge on the sheet. */
function sheetY(v: FloatAnchor['posV'], h: number, sheet: Sheet): number {
  const textBottom = sheet.height - sheet.marginBottom;
  const [base, span] =
    v?.relativeFrom === 'page'
      ? [0, sheet.height]
      : v?.relativeFrom === 'topMargin'
        ? [0, sheet.marginTop]
        : v?.relativeFrom === 'bottomMargin'
          ? [textBottom, sheet.marginBottom]
          : [sheet.marginTop, textBottom - sheet.marginTop];
  if (v?.align === 'center') return base + (span - h) / 2;
  if (v?.align === 'bottom') return base + span - h;
  return base + (v?.offsetPt ?? 0);
}
