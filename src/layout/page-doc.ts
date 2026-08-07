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

import type { BorderStyle, ImageCrop, PictureOutline, ShapeShadow } from '@/core/document-model';
import type { Pt, ResourceId, ResourceStore } from '@/core/ir';
import type { FontMeasure, ParsedTtf } from '@/core/font';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type { PathSegment, StrokeStyle, VectorPath, VectorShape } from '@/core/vector';
import type { PreparedImage } from '@/core/images';
import type { MetaPicture } from '@/core/metafile/picture';

/** A font bound into a {@link LaidOutDocument}: the parsed face plus what layout/emit need from it. */
export interface FontResource {
  readonly resourceName: string;
  readonly parsed: ParsedTtf;
  /**
   * Pure measurement/encoding (no `PdfDocument`): layout measures with this and
   * emit encodes with it; the PDF font objects are created in the emit phase.
   */
  readonly measure: FontMeasure;
  /** Glyphs collected by the usage walk — the emit phase subsets to these. */
  readonly gids: ReadonlySet<number>;
}

/**
 * The face a run asked for and the registry could not give — what the page has
 * to draw itself. A font set carries only `regular` by contract, so a caller
 * who supplies one file still asks for bold headings and italic quotes.
 */
export interface SyntheticFace {
  /** No bold cut: the glyphs are stroked as well as filled. */
  readonly bold?: true;
  /** No italic cut: the glyphs are sheared. */
  readonly italic?: true;
  /**
   * No condensed cut: the glyphs are set at this fraction of their advance
   * (`Arial Narrow` is 82 % of Arial). Layout measures through it too — the
   * squeeze decides where the line breaks, not just how it looks.
   */
  readonly widthScale?: number;
}

/** A run of text on a {@link Line}: the string plus its resolved font, size, width and link/tagging state. */
export interface TextToken {
  readonly kind: 'text';
  readonly text: string;
  readonly isSpace: boolean;
  /**
   * External hyperlink target inherited from the source run (rels-resolved).
   * Writers MUST sanitize the scheme (`core/links`) before emitting anything
   * clickable.
   */
  readonly href?: string;
  /**
   * §17.11.14 — the token renders a footnote reference number; pagination
   * reserves the note's block at the bottom of the page it lands on.
   */
  readonly footnoteRef?: string;
  /**
   * §17.16.22 — internal link target: a bookmark name in this document (not a
   * URL; resolves to a GoTo destination, never through the scheme list).
   */
  readonly anchor?: string;
  /**
   * List-item marker glyphs (`"1."`, `"•"`) — tagged PDF brackets them in a Lbl
   * element separate from the item body.
   */
  readonly listMarker?: true;
  /**
   * §17.3.3.1 — the token is the break of a `w:br w:type="column"`: what
   * follows it belongs in the next column.
   */
  readonly columnBreak?: true;
  /**
   * §17.3.1.38 — the token IS a tab: its width is the distance to the stop it
   * advances to, resolved once the line is known, and its text is whatever
   * leader fills that gap.
   */
  readonly tab?: true;
  /**
   * The token falls inside a comment range (`commentRangeRefs`): the emitter
   * fills a soft highlight behind it (E-COMMENTS CM2c).
   */
  readonly highlight?: true;
  readonly resolvedRun: ResolvedRunProperties;
  readonly font: FontResource;
  /** What the chosen face lacks and the emitter must fake (see {@link SyntheticFace}). */
  readonly synthetic?: SyntheticFace;
  readonly fontSizePt: number;
  /**
   * §17.3.2.42 / §18.4.2 `vertAlign` — how far off the baseline this token
   * draws: positive for a superscript, negative for a subscript, absent on the
   * baseline. The line's height is unchanged; only the glyphs move.
   */
  readonly risePt?: number;
  readonly widthPt: number;
  /**
   * UAX #9 embedding level of this token's characters (0 for pure-LTR docs).
   * Even = LTR, odd = RTL. Used for visual reordering at emit time.
   */
  readonly bidiLevel: number;
}

/** An inline image on a {@link Line}: its sized box plus the resource name binding to the page XObject. */
export interface ImageToken {
  readonly kind: 'image';
  readonly imageResourceName: string;
  readonly widthPt: number;
  readonly heightPt: number;
  /** §20.1.2.2.24 — the frame the picture is drawn with, when it has one. */
  readonly outline?: PictureOutline;
  /** §20.1.8.40 — the drop shadow under the picture, when it casts one. */
  readonly shadow?: ShapeShadow;
  /** §14.1.2.10 — the contrast/brightness wash the picture is drawn through. */
  readonly wash?: { readonly gain: number; readonly black: number };
  /**
   * MS-EMF / MS-WMF — the picture to DRAW in this token's box, for an inline
   * metafile: it has no raster to place, so the emitter plays its primitives
   * where the token stands.
   */
  readonly metafile?: MetafileDrawing;
  /** §20.1.8.55 `a:srcRect` — the part of the source the frame shows. */
  readonly crop?: ImageCrop;
  /** §20.1.7.6 — degrees clockwise about the box's centre. */
  readonly rotationDeg?: number;
  /** §20.1.7.6 — the picture drawn mirrored in its box. */
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /**
   * §20.4.2.6 — where the picture itself sits inside the reserved box when the
   * drawing asked for an effect extent: offsets from the box's left edge and
   * from the baseline, plus the drawn size. Absent ⇒ the picture fills the box.
   */
  readonly drawBox?: {
    readonly dxPt: number;
    readonly dyPt: number;
    readonly widthPt: number;
    readonly heightPt: number;
  };
  /** Constants kept to satisfy {@link Token} consumers — they never read these for images. */
  readonly isSpace: false;
  readonly bidiLevel: number;
}

/**
 * A math draw item with its font already resolved (`math-layout` emits a
 * variant; the tokenizer maps it to a {@link FontResource} so the emit phase
 * needs no fonts).
 */
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

/**
 * An inline OfficeMath object — an atomic box straddling the baseline (its own
 * ascent/descent extend the line height).
 */
export interface MathToken {
  readonly kind: 'math';
  readonly items: ReadonlyArray<ResolvedMathItem>;
  readonly widthPt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
  readonly isSpace: false;
  readonly bidiLevel: number;
}

/** A positioned atom on a {@link Line}: a text run, an inline image, or an inline math box. */
export type Token = TextToken | ImageToken | MathToken;

/** A laid-out line of a paragraph: its broken tokens plus the geometry pagination needs. */
export interface Line {
  readonly tokens: ReadonlyArray<Token>;
  readonly contentWidthPt: number;
  readonly maxFontSizePt: number;
  readonly availableWidthPt: number;
  readonly firstLine: boolean;
  readonly resolved: ResolvedParagraphProperties;
  isLastInParagraph: boolean;
  /**
   * ECMA-376 §17.15.1.35 — the line ends at a soft line break in a document
   * that asked for `w:doNotExpandShiftReturn`, so justification leaves it at
   * its natural width (as the paragraph's own last line is left).
   */
  readonly noJustify?: boolean;
  /**
   * Max ascent/descent contributed by math tokens (0 when none) — they straddle
   * the baseline, so the line height/descent must grow to fit them.
   */
  readonly mathAscentPt?: number;
  readonly mathDescentPt?: number;
  /**
   * E-PARITY: metric-derived single-line height and descent (Pt), the max over
   * the line's text-token fonts under a non-default `layoutProfile`. Absent under
   * `'ream'`, where leading stays the flat 1.2×/0.2 model (byte-identical).
   */
  readonly metricHeightPt?: number;
  readonly metricDescentPt?: number;
}

/** An image bound into a {@link LaidOutDocument}: its resource name plus the decoded/validated bytes. */
/**
 * MS-EMF / MS-WMF — a metafile picture, ready to draw: paths and words in a
 * local y-up frame whose origin is the picture box's bottom-left corner. The
 * same vocabulary a chart's layout carries, because a metafile draws the same
 * two things.
 */
export interface MetafileDrawing {
  readonly shapes: ReadonlyArray<{
    readonly paths: ReadonlyArray<VectorPath>;
    readonly fillColorHex?: string;
    readonly stroke?: StrokeStyle;
  }>;
  readonly texts: ReadonlyArray<{
    readonly line: Line;
    readonly x: number;
    readonly y: number;
    readonly rotationDeg?: number;
  }>;
  /**
   * MS-EMF §2.3.1 — the bitmaps the picture blits into itself, each already a
   * page resource of its own. `x`/`y` are the box's BOTTOM-left corner, in the
   * same local frame as the shapes.
   */
  readonly images?: ReadonlyArray<{
    readonly resourceName: string;
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
    readonly rotationDeg?: number;
  }>;
}

export interface ImageResource {
  readonly resourceName: string;
  /**
   * Decoded/validated at layout time (the probe); the emit phase replays it
   * without touching the source bytes again. Absent for a METAFILE, which has
   * no raster to embed: it is drawn as the primitives `metafile` holds.
   */
  readonly prepared?: PreparedImage;
  /**
   * MS-EMF / MS-WMF — the picture read out of the resource, in its own logical
   * units. A metafile is a little drawing program, not a raster; the layout
   * turns these into the same primitives a chart is made of.
   */
  readonly metafile?: MetaPicture;
  /**
   * The resource name of each bitmap `metafile` blits, in the order its
   * primitives hold them. A metafile draws its own pictures, and each is an
   * ordinary image resource — entered under a key of its own, so the emitters
   * embed and name it without knowing where it came from.
   */
  readonly metafileImages?: ReadonlyArray<string>;
}

/**
 * The fields every {@link PageItem} shares (ir-design §6, frozen at stage 6.4):
 * the positioned, layout-output vocabulary a page is made of. Page-frame
 * geometry is branded `Pt` (PostScript points) with a TOP-LEFT origin, y growing
 * downward (like CSS/SVG); the PDF emitter converts into PDF's native y-up frame
 * at emission. Offsets measured from a text baseline (math/inline-image boxes
 * inside a {@link Line}) and styling magnitudes (font/stroke sizes) stay plain
 * numbers — they are not page-frame coordinates.
 */
export interface PageItemBase {
  /**
   * §20.4.2.3 — the float's `relativeHeight`. Items with one are sorted by it
   * inside their layer before the page is written; the rest keep their order.
   */
  readonly z?: number;
  /**
   * Tagged PDF (§14.8): the logical structure node this item's content belongs
   * to. Set only on body content in tagged mode; undefined text in the line pass
   * is treated as an artifact. Ignored when not tagging.
   */
  readonly structId?: number;
  /**
   * §20.4.2.3 `@behindDoc` — the item is BEHIND the page's content: a
   * watermark, a slide's backdrop. Such items paint before everything else, in
   * their own order. Without that they rode the ordinary passes, where every
   * shape paints after every image — so a slide's white backdrop covered the
   * photograph the slide is made of (corpus: tdf156808, tdf157635, tdf156856).
   */
  readonly behind?: boolean;
  /**
   * The picture this item belongs to, when it is part of one. A metafile is a
   * list of drawing orders, and text among them is BOTH over what came before
   * and under what comes after: an embedded diagram writes a label, lays a
   * panel over it and writes the label again, offset, as its own drop shadow.
   * Painted in the ordinary passes — every shape, then every line — the buried
   * copy would show through the panel. Items sharing this id are held out of
   * those passes and painted together, in list order.
   */
  readonly pictureId?: number;
  /**
   * Tagged PDF: explicitly mark this item as a pagination artifact (running
   * header/footer, §14.8.2.2.2). Distinguishes header/footer text from
   * not-yet-tagged body content so it is typed `/Artifact /Pagination`, never a P.
   */
  readonly artifact?: 'pagination';
}

/** A laid-out line of text (tokens carry their fonts/sizes/positions). */
export interface TextLineItem extends PageItemBase {
  readonly type: 'line';
  readonly line: Line;
  readonly originX: Pt;
  /** Distance from the page TOP down to the text baseline. */
  readonly baselineY: Pt;
  readonly clip?: { readonly x: Pt; readonly y: Pt; readonly width: Pt; readonly height: Pt };
  /**
   * Counter-clockwise rotation about `(originX, baselineY)`, in degrees. A
   * value-axis title reads bottom-to-top in every reader (§21.2.2.216
   * `c:title/c:tx/…/a:bodyPr@rot` = -5400000, i.e. 90°); drawn flat it sat over
   * the plot's own tick labels. Absent ⇒ horizontal.
   */
  readonly rotationDeg?: number;
  /**
   * §20.1.9.10 `a:prstTxWarp` — the line is WordArt: each glyph is placed on
   * its own, bent through the preset's curve and stretched onto the shape's
   * box. Absent ⇒ the line is set flat, one text matrix for the whole of it.
   */
  readonly warp?: TextLineWarp;
}

/**
 * What placing a warped line needs: the preset, the shape box the warped block
 * is stretched onto, and the box the UN-warped block occupies — both in the
 * page's top-left frame, so a glyph's position in the second maps straight into
 * the first.
 */
export interface TextLineWarp {
  readonly preset: string;
  readonly adjust?: number;
  readonly boxX: Pt;
  readonly boxY: Pt;
  readonly boxWidth: Pt;
  readonly boxHeight: Pt;
  /** The un-warped block's left edge and width (its widest line). */
  readonly srcX: Pt;
  readonly srcWidth: Pt;
  /** The un-warped block's ink top and total ink height. */
  readonly srcTop: Pt;
  readonly srcHeight: Pt;
}

/** One edge of a table-cell frame; `(x, y)` is the cell box's top-left corner. */
export interface BorderItem extends PageItemBase {
  readonly type: 'border';
  readonly side: 'top' | 'right' | 'bottom' | 'left';
  readonly x: Pt;
  readonly y: Pt;
  readonly width: Pt;
  readonly height: Pt;
  readonly borderSizePt: number;
  readonly borderColorHex: string;
  /**
   * §18.18.3 `ST_BorderStyle` — the rule's PATTERN, when it is not solid.
   * `dashed`/`dotted` are line patterns, not weights, and dropping them here
   * painted 57423.xlsx's dashed band as twelve continuous rules.
   */
  readonly borderStyle?: BorderStyle;
}

/** A filled rectangle (cell shading); `(x, y)` is its top-left corner. */
export interface FillItem extends PageItemBase {
  readonly type: 'fill';
  readonly x: Pt;
  readonly y: Pt;
  readonly width: Pt;
  readonly height: Pt;
  readonly fillColorHex: string;
}

/**
 * A placed raster image (resource name binds to the page XObject dict); `(x, y)`
 * is its top-left corner.
 */
export interface ImageItem extends PageItemBase {
  readonly type: 'image';
  readonly x: Pt;
  readonly y: Pt;
  readonly width: Pt;
  readonly height: Pt;
  readonly imageResourceName: string;
  /** §20.1.8.55 `a:srcRect` — the part of the source the frame shows. */
  readonly crop?: ImageCrop;
  /** §14.1.2.10 — the contrast/brightness wash the picture is drawn through. */
  readonly wash?: { readonly gain: number; readonly black: number };
  /**
   * §20.1.8.4 `a:alphaModFix` — how opaque the picture is drawn, 0…1. A slide
   * backed by a photograph at 70 % shows a pale wash of it, not the photograph.
   */
  readonly alpha?: number;
  /**
   * §20.1.8.23 `a:duotone` — the picture painted between two colours instead of
   * its own: its dark end is the first, its light end the second.
   */
  readonly duotone?: { readonly shadowHex: string; readonly highlightHex: string };
  /** §20.1.7.6 — degrees clockwise about the box's centre. */
  readonly rotationDeg?: number;
  /** §20.1.7.6 — the picture drawn mirrored in its box. */
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /**
   * §20.1.8.14 — a picture FILL is the shape's outline painted with a picture,
   * so it is clipped to that outline: the paths in the shape's local frame and
   * the matrix that maps them onto the page.
   */
  readonly clip?: {
    readonly paths: ReadonlyArray<VectorPath>;
    readonly transform: readonly [number, number, number, number, number, number];
  };
}

/**
 * DrawingML vector geometry. `shape.transform` maps the shape's local y-up frame
 * into the top-left page frame (see `flipTransform`).
 */
export interface ShapeItem extends PageItemBase {
  readonly type: 'shape';
  readonly shape: VectorShape;
}

/** A positioned item on a laid-out page: a text line, border, fill, image, or vector shape. */
export type PageItem = TextLineItem | BorderItem | FillItem | ImageItem | ShapeItem;
// Canonical paint order of a laid-out page (oop-design §3.2): fills under
// everything, then images, borders, vector shapes, finally text. ONE owner
// for every writer (pdf emit, svg, future canvas) — and the switch is
// exhaustive, so a new PageItem kind refuses to compile until each group
// has a home. Order within a group is the layout's emission order.
export interface PagePaintPlan {
  /** What sits behind the page's content, in its own order — painted first. */
  readonly behind: ReadonlyArray<PageItem>;
  readonly fills: ReadonlyArray<FillItem>;
  readonly images: ReadonlyArray<ImageItem>;
  readonly borders: ReadonlyArray<BorderItem>;
  readonly shapes: ReadonlyArray<ShapeItem>;
  readonly lines: ReadonlyArray<TextLineItem>;
  /**
   * Pictures, each as its own list in the order the picture draws — one entry
   * per {@link PageItemBase.pictureId}. They paint where the shape pass does.
   */
  readonly pictures: ReadonlyArray<ReadonlyArray<PageItem>>;
}

export function paintPlan(commands: ReadonlyArray<PageItem>): PagePaintPlan {
  const fills: Array<FillItem> = [];
  const images: Array<ImageItem> = [];
  const borders: Array<BorderItem> = [];
  const shapes: Array<ShapeItem> = [];
  const lines: Array<TextLineItem> = [];
  // A PICTURE paints as one thing, in its own order (see `pictureId`), so its
  // items leave the passes below and travel together. So does everything the
  // page puts BEHIND its content, which paints first and in its own order.
  const pictures = new Map<number, Array<PageItem>>();
  const behind: Array<PageItem> = [];
  for (const c of commands) {
    if (c.behind === true) {
      behind.push(c);
      continue;
    }
    if (c.pictureId !== undefined) {
      const group = pictures.get(c.pictureId);
      if (group) group.push(c);
      else pictures.set(c.pictureId, [c]);
      continue;
    }
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
  return { behind, fills, images, borders, shapes, lines, pictures: [...pictures.values()] };
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
