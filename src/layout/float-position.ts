// §20.4.2.10/§20.4.2.11 — where an anchored drawing stands on its page.
//
// One answer for every place a drawing can hang off: the body, a header, a
// footer, a table cell, a section whose text runs down the sheet. Each caller
// describes the boxes its anchor can name (a {@link FloatFrame}) and the rules
// below place the drawing in them. A header kept rules of its own once, and
// they knew no `relativeFrom` at all: #143's page-wide header picture, aligned
// right against the PAGE, was aligned right against the text instead and left
// a white strip as wide as the right margin down the side of the sheet.
//
// The rules are Word's, measured (Word for Mac 16, PDF export, our own probes:
// 48 pictures a page, every `relativeFrom` × every alignment and an offset, in
// the body, a header and a footer, on odd and even pages, with and without
// mirrored margins). Word places a header's drawing exactly as it places the
// body's; only the anchoring paragraph differs.

import type { FloatAnchor } from '@/core/document-model';

/**
 * §20.4.3.4 ST_RelFromH / §20.4.3.5 ST_RelFromV — the boxes an anchor's
 * `relativeFrom` names, in page coordinates (y-up). Normally they are the
 * page, its margins and the text column; for a drawing anchored inside a table
 * cell they are all that cell, which is what `@layoutInCell` asks for.
 */
export interface FloatFrame {
  /** The `page` box. */
  readonly pageLeft: number;
  readonly pageWidth: number;
  readonly pageTopYUp: number;
  readonly pageBottomYUp: number;
  /**
   * The `margin` box — the text area of THIS page. Where a band is taller
   * than its margin, the text area ends where the band begins, and so does
   * the box: under a 52pt footer on a page whose bottom margin is 43pt, Word
   * aligns a picture's foot to the bottom of the `margin` 52pt up the page.
   */
  readonly marginLeft: number;
  readonly contentWidth: number;
  readonly marginTopYUp: number;
  readonly marginBottomYUp: number;
  /** The `column` box. A header or footer has one column: its own text area. */
  readonly columnLeft: number;
  readonly columnWidth: number;
  /**
   * The anchoring paragraph's top, y-up — above its space before: Word hangs
   * a drawing 0pt below its `paragraph` 24pt above the paragraph's first line
   * when the paragraph has 24pt of space before it.
   */
  readonly anchorY: number;
  /**
   * The height of the line the anchor sits in, counted from `anchorY`, for the
   * `line` box an alignment is measured in. Absent ⇒ a line of no height.
   */
  readonly anchorLineHeightPt?: number;
  /**
   * Where the anchor character stands across the page, for `character`.
   * Absent ⇒ the column's left edge.
   */
  readonly anchorX?: number;
  /**
   * §20.4.3.1/§20.4.3.2 — whether the page is an EVEN one, where `inside`
   * means right (and bottom) and `outside` left (and top). It is the sheet's
   * own count: a section numbered from 1 on the second sheet still puts its
   * first page's `inside` picture on the right.
   */
  readonly evenPage?: boolean;
  /**
   * How far in from the page's top and bottom edges a picture aligned
   * `inside` or `outside` the PAGE (vertically) stands. See {@link floatTopYUpIn}.
   * Absent ⇒ 0, the edges themselves.
   */
  readonly pageInsetTopPt?: number;
  readonly pageInsetBottomPt?: number;
}

/**
 * How far in from the page's top and bottom edges a header's or footer's
 * picture aligned `inside` or `outside` the page (vertically) stands: 18pt,
 * whatever the band's own distance from the edge. Measured with the header
 * 5pt, 14.2pt, 36pt and 72pt from the edge, and the footer 10pt to 72pt:
 * 18pt each time. (The body's is half the band distance — see the frame.)
 */
export const BAND_PAGE_INSET_PT = 18;

/** A box across the page: its left edge and its width. */
type Span = readonly [left: number, width: number];

/**
 * §20.4.2.10 — the drawing's left edge on the page.
 *
 * - The box is the one `relativeFrom` names (§20.4.3.4). `leftMargin` and
 *   `rightMargin` are the BANDS between the page edge and the text area, which
 *   is where a marginal note is placed: tdf103573.docx centres one box in each.
 *   `insideMargin` is the left band on an odd page and the right one on an even
 *   page, `outsideMargin` the other way round — with or without mirrored
 *   margins. `character` is a box of no width where the anchor stands, so a
 *   picture aligned `center` to it straddles the anchor and one aligned `right`
 *   ends at it.
 * - §20.4.3.1 — `inside` is `left` on an odd page and `right` on an even one,
 *   `outside` the opposite; that too holds without mirrored margins.
 * - An offset is always measured from the box's LEFT edge, on even pages too.
 *
 * @param frame   The boxes the anchor can name.
 * @param f       The anchor.
 * @param widthPt The drawing's width.
 * @returns The left edge, in page points.
 */
export function floatXIn(frame: FloatFrame, f: FloatAnchor, widthPt: number): number {
  const h = f.posH;
  if (!h) return frame.columnLeft;
  const even = frame.evenPage === true;
  const textRight = frame.marginLeft + frame.contentWidth;
  const pageRight = frame.pageLeft + frame.pageWidth;
  const leftBand: Span = [frame.pageLeft, frame.marginLeft - frame.pageLeft];
  const rightBand: Span = [textRight, pageRight - textRight];
  const [base, span]: Span =
    h.relativeFrom === 'page'
      ? [frame.pageLeft, frame.pageWidth]
      : h.relativeFrom === 'column'
        ? [frame.columnLeft, frame.columnWidth]
        : h.relativeFrom === 'character'
          ? [frame.anchorX ?? frame.columnLeft, 0]
          : h.relativeFrom === 'leftMargin'
            ? leftBand
            : h.relativeFrom === 'rightMargin'
              ? rightBand
              : h.relativeFrom === 'insideMargin'
                ? even
                  ? rightBand
                  : leftBand
                : h.relativeFrom === 'outsideMargin'
                  ? even
                    ? leftBand
                    : rightBand
                  : [frame.marginLeft, frame.contentWidth];
  const align =
    h.align === 'inside'
      ? even
        ? 'right'
        : 'left'
      : h.align === 'outside'
        ? even
          ? 'left'
          : 'right'
        : h.align;
  if (align === 'center') return base + (span - widthPt) / 2;
  if (align === 'right') return base + span - widthPt;
  if (align === 'left') return base;
  return base + (h.offsetPt ?? 0);
}

/**
 * §20.4.2.11 — the drawing's TOP on the page, y-up.
 *
 * - The box is the one `relativeFrom` names (§20.4.3.5). `topMargin` runs from
 *   the page's top edge to the text area, `bottomMargin` from the text area to
 *   the page's foot; `insideMargin` is the top one on an odd page and the
 *   bottom one on an even page, `outsideMargin` the other way round. `line` is
 *   the anchoring line, from the top of the paragraph's space before.
 * - §20.4.3.2 — `inside` is `top` on an odd page and `bottom` on an even one,
 *   `outside` the opposite.
 * - A `paragraph` anchor takes no alignment at all: top, centre, bottom,
 *   inside and outside all put the picture's top on the paragraph's.
 * - `inside`/`outside` of the PAGE itself stops short of the edge: in the body
 *   by half the header's distance from the top edge (7.1pt for a header 14.2pt
 *   down, 18pt for 36pt, 2.5pt for 5pt) and half the footer's from the bottom
 *   (14.2pt for 28.35pt, 22.5pt for 45pt); in a header or footer by
 *   {@link BAND_PAGE_INSET_PT}. The frame carries those insets.
 *
 * @param frame    The boxes the anchor can name.
 * @param f        The anchor.
 * @param heightPt The drawing's height.
 * @returns The top edge, y-up.
 */
export function floatTopYUpIn(frame: FloatFrame, f: FloatAnchor, heightPt = 0): number {
  const v = f.posV;
  if (!v) return frame.anchorY;
  const even = frame.evenPage === true;
  // [top, bottom], both y-up.
  const topBand = [frame.pageTopYUp, frame.marginTopYUp] as const;
  const bottomBand = [frame.marginBottomYUp, frame.pageBottomYUp] as const;
  const [top, bottom] =
    v.relativeFrom === 'page'
      ? [frame.pageTopYUp, frame.pageBottomYUp]
      : v.relativeFrom === 'margin'
        ? [frame.marginTopYUp, frame.marginBottomYUp]
        : v.relativeFrom === 'topMargin'
          ? topBand
          : v.relativeFrom === 'bottomMargin'
            ? bottomBand
            : v.relativeFrom === 'insideMargin'
              ? even
                ? bottomBand
                : topBand
              : v.relativeFrom === 'outsideMargin'
                ? even
                  ? topBand
                  : bottomBand
                : v.relativeFrom === 'line'
                  ? [frame.anchorY, frame.anchorY - (frame.anchorLineHeightPt ?? 0)]
                  : [frame.anchorY, frame.anchorY];
  if (v.align === undefined || v.relativeFrom === 'paragraph') return top - (v.offsetPt ?? 0);
  const toTop = v.align === 'top' || v.align === (even ? 'outside' : 'inside');
  const toBottom = v.align === 'bottom' || v.align === (even ? 'inside' : 'outside');
  const edge = v.relativeFrom === 'page' && (v.align === 'inside' || v.align === 'outside');
  if (toTop) return top - (edge ? (frame.pageInsetTopPt ?? 0) : 0);
  if (toBottom) return bottom + heightPt + (edge ? (frame.pageInsetBottomPt ?? 0) : 0);
  return bottom + (top - bottom + heightPt) / 2;
}
