// PageDoc — the laid-out document schema (ir-design §6, FROZEN at stage 6.4).
//
// The second IR tree: what layoutStyledDocument produces and every writer
// consumes. Deliberately importing nothing outside core/ — a writer can
// depend on this file without dragging in any other format module.
//
// Page-frame geometry is branded Pt with a TOP-LEFT origin, y growing
// downward (like CSS/SVG); the PDF emitter converts into PDF's native y-up
// frame at emission. Offsets measured from a text baseline (math/inline-image
// boxes inside a Line) and styling magnitudes (font/stroke sizes) stay plain
// numbers — they are not page-frame coordinates.

import type { Pt, ResourceId, ResourceStore } from '@/core/ir';
import type { FontMeasure, ParsedTtf } from '@/core/font';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type { PathSegment, VectorShape } from '@/core/vector';
import type { PreparedImage } from '@/core/images';

export interface FontResource {
  readonly resourceName: string;
  readonly parsed: ParsedTtf;
  // Pure measurement/encoding (no PdfDocument): layout measures with this and
  // emit encodes with it; the PDF font objects are created in the emit phase.
  readonly measure: FontMeasure;
  // Glyphs collected by the usage walk — the emit phase subsets to these.
  readonly gids: ReadonlySet<number>;
}

export interface TextToken {
  readonly kind: 'text';
  readonly text: string;
  readonly isSpace: boolean;
  // External hyperlink target inherited from the source run (rels-resolved).
  // Writers MUST sanitize the scheme (core/links) before emitting anything
  // clickable.
  readonly href?: string;
  // §17.11.14 — the token renders a footnote reference number; pagination
  // reserves the note's block at the bottom of the page it lands on.
  readonly footnoteRef?: string;
  // §17.16.22 — internal link target: a bookmark name in this document (not
  // a URL; resolves to a GoTo destination, never through the scheme list).
  readonly anchor?: string;
  // List-item marker glyphs ("1.", "•") — tagged PDF brackets them in a Lbl
  // element separate from the item body.
  readonly listMarker?: true;
  // The token falls inside a comment range (commentRangeRefs): the emitter
  // fills a soft highlight behind it (E-COMMENTS CM2c).
  readonly highlight?: true;
  readonly resolvedRun: ResolvedRunProperties;
  readonly font: FontResource;
  readonly fontSizePt: number;
  readonly widthPt: number;
  // UAX #9 embedding level of this token's characters (0 for pure-LTR docs).
  // Even = LTR, odd = RTL. Used for visual reordering at emit time.
  readonly bidiLevel: number;
}

export interface ImageToken {
  readonly kind: 'image';
  readonly imageResourceName: string;
  readonly widthPt: number;
  readonly heightPt: number;
  // Constants kept to satisfy Token consumers — they never read these for images.
  readonly isSpace: false;
  readonly bidiLevel: number;
}

// A math draw item with its font already resolved (math-layout emits a variant;
// the tokenizer maps it to a FontResource so the emit phase needs no fonts).
export type ResolvedMathItem =
  | {
      readonly kind: 'glyph';
      readonly x: number;
      readonly y: number;
      readonly text: string;
      readonly font: FontResource;
      readonly sizePt: number;
    }
  | {
      readonly kind: 'rule';
      readonly x: number;
      readonly y: number;
      readonly w: number;
      readonly h: number;
    }
  | {
      readonly kind: 'path';
      readonly segments: ReadonlyArray<PathSegment>;
      readonly strokeWidthPt?: number;
      readonly fill?: boolean;
    };

// An inline OfficeMath object — an atomic box straddling the baseline (its own
// ascent/descent extend the line height).
export interface MathToken {
  readonly kind: 'math';
  readonly items: ReadonlyArray<ResolvedMathItem>;
  readonly widthPt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
  readonly isSpace: false;
  readonly bidiLevel: number;
}

export type Token = TextToken | ImageToken | MathToken;

export interface Line {
  readonly tokens: ReadonlyArray<Token>;
  readonly contentWidthPt: number;
  readonly maxFontSizePt: number;
  readonly availableWidthPt: number;
  readonly firstLine: boolean;
  readonly resolved: ResolvedParagraphProperties;
  isLastInParagraph: boolean;
  // Max ascent/descent contributed by math tokens (0 when none) — they straddle
  // the baseline, so the line height/descent must grow to fit them.
  readonly mathAscentPt?: number;
  readonly mathDescentPt?: number;
  // E-PARITY: metric-derived single-line height and descent (Pt), the max over
  // the line's text-token fonts under a non-default layoutProfile. Absent under
  // 'ream', where leading stays the flat 1.2×/0.2 model (byte-identical).
  readonly metricHeightPt?: number;
  readonly metricDescentPt?: number;
}

export interface ImageResource {
  readonly resourceName: string;
  // Decoded/validated at layout time (the probe); the emit phase replays it
  // without touching the source bytes again.
  readonly prepared: PreparedImage;
}

// PageDoc items (ir-design §6, frozen at stage 6.4): the positioned,
// layout-output vocabulary a page is made of. Page-frame geometry is branded
// `Pt` (PostScript points) with a TOP-LEFT origin, y growing downward (like
// CSS/SVG); the PDF emitter converts into PDF's native y-up frame at
// emission. Offsets measured from a text baseline (math/inline-image boxes
// inside a Line) and styling magnitudes (font/stroke sizes) stay plain
// numbers — they are not page-frame coordinates.
export interface PageItemBase {
  // Tagged PDF (§14.8): the logical structure node this item's content belongs
  // to. Set only on body content in tagged mode; undefined text in the line
  // pass is treated as an artifact. Ignored when not tagging.
  readonly structId?: number;
  // Tagged PDF: explicitly mark this item as a pagination artifact (running
  // header/footer, §14.8.2.2.2). Distinguishes header/footer text from
  // not-yet-tagged body content so it is typed /Artifact /Pagination, never a P.
  readonly artifact?: 'pagination';
}

// A laid-out line of text (tokens carry their fonts/sizes/positions).
export interface TextLineItem extends PageItemBase {
  readonly type: 'line';
  readonly line: Line;
  readonly originX: Pt;
  // Distance from the page TOP down to the text baseline.
  readonly baselineY: Pt;
}

// One edge of a table-cell frame; (x, y) is the cell box's top-left corner.
export interface BorderItem extends PageItemBase {
  readonly type: 'border';
  readonly side: 'top' | 'right' | 'bottom' | 'left';
  readonly x: Pt;
  readonly y: Pt;
  readonly width: Pt;
  readonly height: Pt;
  readonly borderSizePt: number;
  readonly borderColorHex: string;
}

// A filled rectangle (cell shading); (x, y) is its top-left corner.
export interface FillItem extends PageItemBase {
  readonly type: 'fill';
  readonly x: Pt;
  readonly y: Pt;
  readonly width: Pt;
  readonly height: Pt;
  readonly fillColorHex: string;
}

// A placed raster image (resource name binds to the page XObject dict);
// (x, y) is its top-left corner.
export interface ImageItem extends PageItemBase {
  readonly type: 'image';
  readonly x: Pt;
  readonly y: Pt;
  readonly width: Pt;
  readonly height: Pt;
  readonly imageResourceName: string;
}

// DrawingML vector geometry. `shape.transform` maps the shape's local y-up
// frame into the top-left page frame (see flipTransform).
export interface ShapeItem extends PageItemBase {
  readonly type: 'shape';
  readonly shape: VectorShape;
}

export type PageItem = TextLineItem | BorderItem | FillItem | ImageItem | ShapeItem;
// Canonical paint order of a laid-out page (oop-design §3.2): fills under
// everything, then images, borders, vector shapes, finally text. ONE owner
// for every writer (pdf emit, svg, future canvas) — and the switch is
// exhaustive, so a new PageItem kind refuses to compile until each group
// has a home. Order within a group is the layout's emission order.
export interface PagePaintPlan {
  readonly fills: ReadonlyArray<FillItem>;
  readonly images: ReadonlyArray<ImageItem>;
  readonly borders: ReadonlyArray<BorderItem>;
  readonly shapes: ReadonlyArray<ShapeItem>;
  readonly lines: ReadonlyArray<TextLineItem>;
}

export function paintPlan(commands: ReadonlyArray<PageItem>): PagePaintPlan {
  const fills: Array<FillItem> = [];
  const images: Array<ImageItem> = [];
  const borders: Array<BorderItem> = [];
  const shapes: Array<ShapeItem> = [];
  const lines: Array<TextLineItem> = [];
  for (const c of commands) {
    switch (c.type) {
      case 'fill':
        fills.push(c);
        break;
      case 'image':
        images.push(c);
        break;
      case 'border':
        borders.push(c);
        break;
      case 'shape':
        shapes.push(c);
        break;
      case 'line':
        lines.push(c);
        break;
      default:
        assertNeverPageItem(c);
    }
  }
  return { fills, images, borders, shapes, lines };
}

function assertNeverPageItem(item: never): never {
  throw new Error(`Unhandled PageItem kind: ${String((item as PageItem).type)}`);
}

// PageDoc (ir-design §6, frozen at stage 6.4): the laid-out document a writer
// consumes — positioned PageItems per page plus the font/image resources the
// items reference. Format-neutral: the SVG writer renders exactly this; PDF
// needs the PdfLayoutAux companion on top.
export interface LaidOutDocument {
  readonly pages: ReadonlyArray<LaidOutPage>;
  // Content-addressed binaries the items reference (images) — ir-design §6.
  readonly resources: ResourceStore;
  readonly fontResources: Map<string, FontResource>;
  readonly imageResources: Map<ResourceId, ImageResource>;
}

export interface LaidOutPage {
  readonly commands: Array<PageItem>;
  readonly width: Pt;
  readonly height: Pt;
}
