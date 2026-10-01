// A page cut down to what is drawn on it — for a document whose page is its
// content rather than a sheet of paper: a worksheet drawn as an image of
// itself, laid out on a page with room for any sheet and then trimmed.

import type { LaidOutDocument, LaidOutPage, PageItem } from '@/layout/page-doc';
import type { VectorShape } from '@/core/vector';

import { pt } from '@/core/ir';

/**
 * Each page cut down to the right and bottom edges of what is drawn on it,
 * plus `pad`. The top-left corner stays where it is — the content's origin is
 * the page's, and the layout already kept its margin there.
 *
 * @param laid The laid-out document, its pages larger than their content.
 * @param pad  The room kept past the content on the right and at the bottom.
 * @returns The same document on pages no larger than what they hold.
 */
export function fitPagesToContent(laid: LaidOutDocument, pad: number): LaidOutDocument {
  return { ...laid, pages: laid.pages.map((page) => fitPage(page, pad)) };
}

function fitPage(page: LaidOutPage, pad: number): LaidOutPage {
  let right = 0;
  let bottom = 0;
  for (const item of page.commands) {
    const reach = reachOf(item);
    right = Math.max(right, reach.right);
    bottom = Math.max(bottom, reach.bottom);
  }
  return {
    ...page,
    width: pt(Math.min(page.width, Math.ceil(right + pad))),
    height: pt(Math.min(page.height, Math.ceil(bottom + pad))),
  };
}

/** How far right and down one item paints. */
function reachOf(item: PageItem): { right: number; bottom: number } {
  switch (item.type) {
    case 'fill':
    case 'image':
      return { right: item.x + item.width, bottom: item.y + item.height };
    case 'border': {
      // A rule is centred on its edge: half of it lies past the box.
      const half = item.borderSizePt / 2;
      return { right: item.x + item.width + half, bottom: item.y + item.height + half };
    }
    case 'line': {
      // A line cut to its cell paints no further than the cut.
      if (item.clip) {
        return { right: item.clip.x + item.clip.width, bottom: item.clip.y + item.clip.height };
      }
      const descent = item.line.metricDescentPt ?? item.line.maxFontSizePt * 0.25;
      return {
        right: item.originX + item.line.contentWidthPt,
        bottom: item.baselineY + descent,
      };
    }
    case 'shape':
      return shapeReach(item.shape);
  }
}

/** The far corner of a shape's paths on the page, its stroke included (control points bound a curve). */
function shapeReach(shape: VectorShape): { right: number; bottom: number } {
  const [a, b, c, d, e, f] = shape.transform;
  const half = (shape.stroke?.widthPt ?? 0) / 2;
  let right = 0;
  let bottom = 0;
  const take = (x: number, y: number): void => {
    right = Math.max(right, a * x + c * y + e + half);
    bottom = Math.max(bottom, b * x + d * y + f + half);
  };
  for (const path of shape.paths) {
    for (const s of path.segments) {
      if (s.op === 'close') continue;
      if (s.op === 'cubic') {
        take(s.x1, s.y1);
        take(s.x2, s.y2);
      }
      take(s.x, s.y);
    }
  }
  return { right, bottom };
}
