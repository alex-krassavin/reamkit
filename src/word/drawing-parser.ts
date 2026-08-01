// ECMA-376 Part 1 §20 (DrawingML) + Part 3 (Markup Compatibility).
//
// Parses a <w:drawing> into either an embedded picture reference (the
// pre-existing behaviour) or a DrawingML shape (<wps:wsp>). Markup
// Compatibility <mc:AlternateContent> is resolved here, preferring the modern
// wps Choice over the legacy VML Fallback (which we cannot render).

import type {
  BodyElement,
  CustomPathCmd,
  FloatAnchor,
  ImageCrop,
  InlineImage,
  LineEnd,
  RelativeSize,
  ShapeDash,
  ShapeFill,
  ShapeGeometry,
  ShapeGroupChild,
  ShapeLine,
  ShapeShadow,
  ShapeTextBody,
  ShapeTransform,
} from '@/core/document-model';
import type { ColorMod, ColorResolver } from '@/core/drawingml/colors';
import type { PoNode } from '@/core/po-helpers';
import type { Pt, ResourceId } from '@/core/ir';
import type { GradientStop, ShapeGradient } from '@/core/vector';
import { readColorMods, resolveColorNode } from '@/core/drawingml/colors';
import { emuToPt, pt } from '@/core/ir';
import {
  poAttr,
  poAttrLocal,
  poChildren,
  poFindDescendant,
  poIntAttr,
  poIs,
  poIsLocal,
  poTag,
  poText,
} from '@/core/po-helpers';

const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const WPG_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
// §20.3 — a canvas of shapes, written in the plain DrawingML namespace. Word
// exports a pasted PowerPoint group this way.
const LOCKED_CANVAS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas';
// Word's own drawing canvas: the same members (`pic:pic`, `wps:wsp`, `wpg:wgp`)
// placed by their own `a:xfrm` inside the frame `wp:extent` gives.
const WPC_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas';

// A group's own "geometry": no commands, so its box paints nothing.
const EMPTY_GEOMETRY = { pathWidth: 0, pathHeight: 0, commands: [] };

// Namespaces whose <mc:Choice> we can render. 'wps' = wordprocessingShape,
// 'wpg' = wordprocessingGroup. The VML in the Fallback is a different geometry
// language altogether, so a Choice we can read is always the better branch.
// The `mc:Choice` namespaces we can actually read. `wpc` joined them once the
// canvas below was understood — until then fdo65833.docx fell through to its
// VML fallback, which drew the picture at the whole canvas's size.
const UNDERSTOOD_NS = new Set(['wps', 'wpg', 'wpc']);

/**
 * A parsed DrawingML shape without the owning paragraph's properties (attached by
 * the caller, mirroring how the image branch returns size + id and the caller
 * adds the `pPr`).
 */
export interface ShapeData {
  readonly width: Pt;
  readonly height: Pt;
  /** §20.5.2.17 — the shapes a `wpg:wgp` group holds, placed in its own box. */
  readonly children?: ReadonlyArray<ShapeGroupChild>;
  readonly geometry: ShapeGeometry;
  readonly fill: ShapeFill;
  readonly line?: ShapeLine;
  readonly transform?: ShapeTransform;
  /** `wp14:sizeRelH/V` — a size stated as a share of the page or margins. */
  readonly relativeSize?: RelativeSize;
  /** The shape's text body (a `wps:txbx`), when it carries one. */
  readonly text?: ShapeTextBody;
}

/**
 * Parses the body elements of a `w:txbxContent`. Injected by the caller to avoid
 * a module cycle with `document-parser` (which imports this module).
 */
export type ParseBody = (children: ReadonlyArray<PoNode>) => Array<BodyElement>;

/**
 * The result of parsing a `<w:drawing>` (or legacy VML picture): an embedded
 * picture, a DrawingML shape, a chart reference, or a SmartArt diagram. Each
 * variant carries optional alternate text and a float anchor.
 */
export type DrawingContent =
  | {
      readonly kind: 'image';
      readonly imageId: string;
      readonly width: Pt;
      readonly height: Pt;
      /** §20.1.8.55 `a:srcRect` — the part of the source the frame shows. */
      readonly crop?: ImageCrop;
      /** §20.1.7.6 `a:xfrm @rot` — the picture's rotation (1/60000°, clockwise). */
      readonly rotation60k?: number;
      /** `wp14:sizeRelH/V` — a size stated as a share of the page or margins. */
      readonly relativeSize?: RelativeSize;
      /** §20.4.2.6 `wp:effectExtent` on an INLINE drawing: space reserved around it. */
      readonly effectExtent?: InlineImage['effectExtent'];
      /** `wp:docPr` `@descr`/`@title` — alternate text for the tagged-PDF Figure. */
      readonly altText?: string;
      readonly float?: FloatAnchor;
    }
  | {
      readonly kind: 'shape';
      readonly data: ShapeData;
      readonly altText?: string;
      readonly float?: FloatAnchor;
    }
  | {
      readonly kind: 'chart';
      /** The `c:chart` `@r:id` relationship id to the chart part. */
      readonly chartRelId: string;
      readonly width: Pt;
      readonly height: Pt;
      readonly altText?: string;
      readonly float?: FloatAnchor;
    }
  | {
      readonly kind: 'diagram';
      /** SmartArt data-part relationship id (`dgm:relIds` `@r:dm`); the reader resolves the drawing override. */
      readonly dmRelId: string;
      /** Frame width in EMU. */
      readonly widthEmu: number;
      /** Frame height in EMU. */
      readonly heightEmu: number;
      readonly altText?: string;
      readonly float?: FloatAnchor;
    };

// §20.4.2.3 — the anchor's placement: position children + wrap mode.
function parseFloatAnchor(anchor: PoNode): FloatAnchor | undefined {
  if (!poIs(anchor, 'wp:anchor')) return undefined;
  const behindRaw = poAttr(anchor, 'behindDoc');
  const behind = behindRaw === '1' || behindRaw === 'true';
  let wrap: FloatAnchor['wrap'] = 'none';
  for (const child of poChildren(anchor)) {
    if (poIs(child, 'wp:wrapSquare')) wrap = 'square';
    else if (poIs(child, 'wp:wrapTight')) wrap = 'tight';
    else if (poIs(child, 'wp:wrapThrough')) wrap = 'through';
    else if (poIs(child, 'wp:wrapTopAndBottom')) wrap = 'topAndBottom';
  }
  const zOrder = poIntAttr(anchor, 'relativeHeight');
  const posH = parseAnchorPos(anchor, 'wp:positionH', ['margin', 'page', 'column']);
  const posV = parseAnchorPos(anchor, 'wp:positionV', ['margin', 'page', 'paragraph', 'line']);
  // §20.4.2.3 — the stand-off the wrapped text keeps from each edge. Ignoring
  // it let effect-extent-line-width.docx fill its paragraph right down to the
  // text box's top edge, where Word and LibreOffice both stop a line earlier
  // and carry the rest below the box.
  const dist = (name: string): Pt => emuToPt(poIntAttr(anchor, name) ?? 0);
  const wrapDist = {
    topPt: dist('distT'),
    bottomPt: dist('distB'),
    leftPt: dist('distL'),
    rightPt: dist('distR'),
  };
  const anyDist = Object.values(wrapDist).some((v) => v > 0);
  return {
    wrap,
    ...(behind ? { behind: true } : {}),
    ...(zOrder !== undefined ? { zOrder } : {}),
    ...(anyDist ? { wrapDist } : {}),
    ...(posH ? { posH: posH as NonNullable<FloatAnchor['posH']> } : {}),
    ...(posV ? { posV: posV as NonNullable<FloatAnchor['posV']> } : {}),
  };
}

const ANCHOR_ALIGNS = new Set(['left', 'center', 'right']);

/**
 * `wp14:sizeRelH` / `wp14:sizeRelV` — the drawing's width and height as
 * hundredths of a percent of the page or the margins. Word writes the resolved
 * extent beside them for readers that do not know the namespace, which is what
 * we drew: dml-shape-relsize.docx asks for 40% of the margin and got the
 * fallback's 90pt.
 *
 * @param anchor The `wp:anchor` node.
 * @returns The relative size, or `undefined` when the anchor states none.
 */
function parseRelativeSize(anchor: PoNode): RelativeSize | undefined {
  const read = (
    tag: string,
    pctTag: string,
  ): { pct: number; from: 'margin' | 'page' } | undefined => {
    const node = expandMcChildren(poChildren(anchor)).find((c) => poIs(c, tag));
    if (!node) return undefined;
    const pctNode = poChildren(node).find((c) => poIs(c, pctTag));
    const raw = pctNode ? Number(poText(pctNode).trim()) : NaN;
    if (!Number.isFinite(raw) || raw <= 0) return undefined;
    return { pct: raw / 100000, from: poAttr(node, 'relativeFrom') === 'page' ? 'page' : 'margin' };
  };
  const h = read('wp14:sizeRelH', 'wp14:pctWidth');
  const v = read('wp14:sizeRelV', 'wp14:pctHeight');
  if (!h && !v) return undefined;
  return {
    ...(h ? { widthPct: h.pct, widthFrom: h.from } : {}),
    ...(v ? { heightPct: v.pct, heightFrom: v.from } : {}),
  };
}

/**
 * The first ELEMENT among a node's children. `poTag` gives a text node no tag
 * at all, and a "not #text" test let the whitespace between pretty-printed
 * elements win: dashed_line_custdash_percentage.docx indents its `a:lnRef`, so
 * the colour inside it went unread and its blue rule came out black.
 *
 * @param node The parent node.
 * @returns The first child that is an element, or `undefined`.
 */
function firstElementChild(node: PoNode | undefined): PoNode | undefined {
  return poChildren(node).find((c) => poTag(c) !== undefined);
}

function parseAnchorPos(
  anchor: PoNode,
  tag: 'wp:positionH' | 'wp:positionV',
  allowed: ReadonlyArray<string>,
): { relativeFrom: string; offsetPt?: number; align?: string } | undefined {
  // §17.17.2 — the position may sit inside an `mc:AlternateContent`: Word
  // writes the wp14 percentage offset as the Choice and the plain EMU offset as
  // the Fallback. Read as plain children the anchor looked positionless and
  // content-control-shape.docx's rule, 97% across the page, was drawn at its
  // left edge.
  const pos = expandMcChildren(poChildren(anchor)).find((c) => poIs(c, tag));
  if (!pos) return undefined;
  const relRaw = poAttr(pos, 'relativeFrom') ?? 'margin';
  // Unsupported bases (character, inside/outsideMargin…) degrade to the
  // nearest supported one.
  const relativeFrom = allowed.includes(relRaw) ? relRaw : allowed[0]!;
  const offsetNode = poChildren(pos).find((c) => poIs(c, 'wp:posOffset'));
  const offsetRaw = offsetNode ? Number(poText(offsetNode).trim()) : NaN;
  const alignNode = poChildren(pos).find((c) => poIs(c, 'wp:align'));
  const alignRaw = alignNode ? poText(alignNode).trim() : '';
  return {
    relativeFrom,
    ...(Number.isFinite(offsetRaw) ? { offsetPt: emuToPt(offsetRaw) } : {}),
    ...(tag === 'wp:positionH' && ANCHOR_ALIGNS.has(alignRaw) ? { align: alignRaw } : {}),
  };
}

/**
 * ECMA-376 Part 3 (Markup Compatibility) — resolve an `<mc:AlternateContent>` to
 * the children of the first `<mc:Choice>` whose `Requires` lists only namespaces
 * we understand, else the `<mc:Fallback>` children, else nothing. (`Requires`
 * holds space-separated namespace prefixes as declared in the document.)
 *
 * @param altContent The `mc:AlternateContent` node.
 * @returns The chosen branch's children.
 */
export function resolveMc(altContent: PoNode): ReadonlyArray<PoNode> {
  for (const choice of poChildren(altContent)) {
    if (!poIs(choice, 'mc:Choice')) continue;
    const requires = (poAttr(choice, 'Requires') ?? '').split(/\s+/).filter(Boolean);
    if (requires.length > 0 && requires.every((r) => UNDERSTOOD_NS.has(r))) {
      return poChildren(choice);
    }
  }
  const fallback = poChildren(altContent).find((c) => poIs(c, 'mc:Fallback'));
  return fallback ? poChildren(fallback) : [];
}

/**
 * Flatten a children list, expanding any `<mc:AlternateContent>` to its chosen
 * branch (via {@link resolveMc}) so downstream scanning sees plain elements (a
 * `<w:drawing>`, or the VML we ignore). Used both at run level and inside
 * `a:graphicData`.
 *
 * @param children The raw child list.
 * @returns The flattened children.
 */
export function expandMcChildren(children: ReadonlyArray<PoNode>): Array<PoNode> {
  const out: Array<PoNode> = [];
  for (const c of children) {
    if (poIs(c, 'mc:AlternateContent')) out.push(...resolveMc(c));
    else out.push(c);
  }
  return out;
}

/**
 * Parse a `<w:drawing>` (ECMA-376 Part 1 §20) into a {@link DrawingContent}. The
 * `a:graphicData` `@uri` selects the branch: a `wps:wsp` shape, a chart, a
 * SmartArt diagram, or — falling through — an embedded picture from
 * `a:blip @r:embed`.
 *
 * @param drawing      The `w:drawing` node.
 * @param resolveColor Resolver for theme/scheme colours used by shape fills/lines.
 * @param parseBody    Optional body parser for a shape's text box (omitted ⇒ no text).
 * @returns The parsed content, or `null` when no anchor / recognizable graphic is found.
 */
export function parseDrawing(
  drawing: PoNode,
  resolveColor: ColorResolver,
  parseBody?: ParseBody,
  // §20.1.8.14 — a group may hold a `pic:pic`, which is a shape with a picture
  // fill; resolving it needs the package's image resolver.
  resolveImage?: (relId: string) => ResourceId | undefined,
  // §21.2 a chart's r:id → its part path, the key the reader files charts
  // under; rel ids alone collide between the body and a header/footer.
  resolveChartPart?: (relId: string) => string | undefined,
  // §20.1.4.2.19 — `a:lnStyleLst` widths, indexed by a gallery style's `a:lnRef`.
  themeLineWidths?: ReadonlyArray<number>,
): DrawingContent | null {
  const anchor =
    poChildren(drawing).find((c) => poIs(c, 'wp:inline')) ??
    poChildren(drawing).find((c) => poIs(c, 'wp:anchor'));
  if (!anchor) return null;

  const extent = poFindDescendant(anchor, 'wp:extent');
  const extentCx = extent ? poIntAttr(extent, 'cx') : undefined;
  const extentCy = extent ? poIntAttr(extent, 'cy') : undefined;

  // wp:docPr (a direct child of wp:inline/wp:anchor) carries the drawing's
  // alternate text: @descr preferred, then @title. Used for the Figure /Alt.
  const docPr = poChildren(anchor).find((c) => poIs(c, 'wp:docPr'));
  const descr = docPr ? poAttr(docPr, 'descr') : undefined;
  const title = docPr ? poAttr(docPr, 'title') : undefined;
  const altText = (descr ?? title)?.trim() || undefined;
  const float = parseFloatAnchor(anchor);
  const alt = {
    ...(altText ? { altText } : {}),
    ...(float ? { float } : {}),
  };

  // `wp14:sizeRelH/V` — the drawing's size as a share of the page or margins.
  // dml-shape-relsize.docx asks for 40% of the margin width and we drew the
  // fallback extent, less than half of it.
  const relativeSize = parseRelativeSize(anchor);
  const graphicData = poFindDescendant(anchor, 'a:graphicData');
  const uri = graphicData ? poAttr(graphicData, 'uri') : undefined;

  if (graphicData && uri === WPS_URI) {
    const data = parseWsp(
      graphicData,
      extentCx,
      extentCy,
      resolveColor,
      parseBody,
      resolveImage,
      themeLineWidths,
    );
    if (!data) return null;
    return { kind: 'shape', data: { ...data, ...(relativeSize ? { relativeSize } : {}) }, ...alt };
  }

  if (graphicData && uri === WPG_URI) {
    const data = parseWgp(
      graphicData,
      extentCx,
      extentCy,
      resolveColor,
      parseBody,
      resolveImage,
      themeLineWidths,
    );
    if (!data) return null;
    return { kind: 'shape', data: { ...data, ...(relativeSize ? { relativeSize } : {}) }, ...alt };
  }

  // §20.3 `lc:lockedCanvas` — a group by another name: the same members, in the
  // `a:` namespace. fdo43641.docx draws its rectangle and arrow inside one and
  // we rendered an empty page.
  if (graphicData && uri === LOCKED_CANVAS_URI) {
    const data = parseLockedCanvas(
      graphicData,
      extentCx,
      extentCy,
      resolveColor,
      parseBody,
      resolveImage,
      themeLineWidths,
    );
    if (!data) return null;
    return { kind: 'shape', data: { ...data, ...(relativeSize ? { relativeSize } : {}) }, ...alt };
  }

  // A Word drawing canvas — a group whose members carry their own positions.
  // Unread, fdo65833.docx's canvas fell through to the picture path and drew
  // the screenshot inside it at the whole canvas's size.
  if (graphicData && uri === WPC_URI) {
    const canvas = expandMcChildren(poChildren(graphicData)).find((c) => poIsLocal(c, 'wpc'));
    if (!canvas || extentCx === undefined || extentCy === undefined) return null;
    const children = groupChildren(canvas, resolveColor, parseBody, resolveImage, themeLineWidths);
    const data: ShapeData = {
      width: emuToPt(extentCx),
      height: emuToPt(extentCy),
      ...(children.length > 0 ? { children } : {}),
      geometry: { kind: 'custom', custom: EMPTY_GEOMETRY },
      fill: { kind: 'none' },
    };
    return { kind: 'shape', data: { ...data, ...(relativeSize ? { relativeSize } : {}) }, ...alt };
  }

  if (graphicData && uri === CHART_URI) {
    const cChart = poFindDescendant(graphicData, 'c:chart');
    const chartRelId = cChart ? poAttr(cChart, 'id') : undefined; // r:id
    if (chartRelId && extentCx !== undefined && extentCy !== undefined) {
      return {
        kind: 'chart',
        chartRelId: resolveChartPart?.(chartRelId) ?? chartRelId,
        width: emuToPt(extentCx),
        height: emuToPt(extentCy),
        ...alt,
      };
    }
    return null;
  }

  // SmartArt diagram: keep the data-part rel id; the reader resolves the drawing
  // override and renders its shapes (E-SMARTART SA2).
  if (graphicData && uri === DIAGRAM_URI) {
    const relIds = poFindDescendant(graphicData, 'dgm:relIds');
    const dmRelId = relIds ? poAttr(relIds, 'dm') : undefined; // r:dm → data part
    if (dmRelId && extentCx !== undefined && extentCy !== undefined) {
      return { kind: 'diagram', dmRelId, widthEmu: extentCx, heightEmu: extentCy, ...alt };
    }
    return null;
  }

  // Picture path: a:blip r:embed + extent.
  const blip = poFindDescendant(anchor, 'a:blip');
  const rId = blip ? poAttr(blip, 'embed') : undefined;
  if (extentCx !== undefined && extentCy !== undefined && rId) {
    const crop = parseSrcRect(poFindDescendant(anchor, 'a:srcRect'));
    // §20.1.7.6 — a picture may be turned in its frame. crop-pixel.docx tilts
    // its cover by 10.7° and we set it square.
    const xfrm = poFindDescendant(anchor, 'a:xfrm');
    const rot = xfrm ? poIntAttr(xfrm, 'rot') : undefined;
    // §20.4.2.6 — an inline drawing reserves its effect extent on the line:
    // effect-extent-inline.docx turns its cover 40° and states the 46pt the
    // corners need on each side, without which the picture sat 48pt to the
    // left of where Word and LibreOffice draw it, up into the top margin.
    const effect = poIs(anchor, 'wp:inline')
      ? poChildren(anchor).find((c) => poIs(c, 'wp:effectExtent'))
      : undefined;
    const side = (name: string): Pt =>
      emuToPt(effect ? Math.max(0, poIntAttr(effect, name) ?? 0) : 0);
    const box = { leftPt: side('l'), topPt: side('t'), rightPt: side('r'), bottomPt: side('b') };
    const effectExtent = Object.values(box).some((v) => v > 0) ? box : undefined;
    return {
      kind: 'image',
      imageId: rId,
      width: emuToPt(extentCx),
      height: emuToPt(extentCy),
      ...(crop ? { crop } : {}),
      ...(rot ? { rotation60k: rot } : {}),
      ...(relativeSize ? { relativeSize } : {}),
      ...(effectExtent ? { effectExtent } : {}),
      ...alt,
    };
  }
  return null;
}

/**
 * Parse a legacy `<w:pict>`/`<w:object>` VML picture (ISO/IEC 29500-1 §14, VML
 * transitional) into an `image` {@link DrawingContent}. Modern files use
 * `<w:drawing>` ({@link parseDrawing}); VML still shows up in headers, OLE-object
 * previews (`@o:ole`) and documents last saved by older Word. A VML shape carries
 * an `<v:imagedata r:id>` pointing at the media part and a CSS-like `@style`
 * (`"width:75.6pt;height:49.2pt"`) giving its box; just enough is read to recover
 * the relationship id, the size and the `@alt` text.
 *
 * @param node The `w:pict` / `w:object` node.
 * @returns The picture, or `null` when there is no embedded `v:imagedata` or no usable size.
 */
export function parseVmlPicture(node: PoNode, parseBody?: ParseBody): DrawingContent | null {
  // §14.1.2 — a `v:group` is a drawing of its own, whatever it holds. Reaching
  // for the first `v:imagedata` in the tree first collapsed a whole canvas to
  // the picture inside it: fdo61343.docx groups a metafile with the text boxes
  // that label it, and the labels went with the group.
  const grouped = poChildren(node).some((c) => poIs(c, 'v:group'));
  const imagedata = grouped ? undefined : poFindDescendant(node, 'v:imagedata');
  if (!imagedata) return parseVmlWordArt(node) ?? parseVmlShape(node, parseBody);
  // @r:id binds the embedded picture. An external @o:href/@r:href link we do
  // not embed, and an empty placeholder frame (a <v:shape> with no imagedata),
  // both leave this undefined and are skipped.
  const imageId = poAttr(imagedata, 'id');
  if (imageId === undefined) return null;
  // The owning shape holds the @style box and @alt text — v:shape for pictures,
  // or a drawn primitive (v:rect/v:oval) when an image fills a shape.
  const shape =
    poFindDescendant(node, 'v:shape') ??
    poFindDescendant(node, 'v:rect') ??
    poFindDescendant(node, 'v:oval');
  const width = vmlStyleLength(shape, 'width');
  const height = vmlStyleLength(shape, 'height');
  if (width === undefined || height === undefined) return null;
  const altText =
    (shape ? poAttr(shape, 'alt') : undefined)?.trim() ||
    poAttr(imagedata, 'title')?.trim() || // o:title
    undefined;
  // §14.1.2 — a positioned VML picture floats exactly as a drawn VML shape
  // does: drawinglayer-pic-pos.docx hangs its photo two inches down the page
  // and we set it inline, over the frame above it.
  const float = shape
    ? vmlFloat(poAttr(shape, 'style') ?? '', poIntAttr(shape, 'z-index'))
    : undefined;
  return {
    kind: 'image',
    imageId,
    width,
    height,
    ...(float ? { float } : {}),
    ...(altText ? { altText } : {}),
  };
}

/**
 * §14.1.2 — a legacy VML shape drawn as the DrawingML one it corresponds to:
 * `v:rect`, `v:roundrect`, `v:oval`, `v:line` and a `v:shape` we cannot resolve
 * a path for (drawn as its box). 113 corpus documents still hold one — older
 * Word wrote every drawing this way — and read only for its `v:imagedata`,
 * every one of them lost the shape, its fill, its outline and the text in it:
 * drawinglayer-pic-pos.docx frames its title in a `v:rect` and we printed the
 * page without it.
 *
 * @param node       The `w:pict` / `w:object` node.
 * @param parseBody  Body parser for a `v:textbox`'s content.
 * @returns The shape, or `null` when nothing drawable is found.
 */
function parseVmlShape(node: PoNode, parseBody?: ParseBody): DrawingContent | null {
  const shape = poChildren(node).find((c) => VML_SHAPE_TAGS.has(poTag(c) ?? ''));
  if (!shape) return null;
  if (poIs(shape, 'v:group')) {
    const data = parseVmlGroup(shape, parseBody, vmlShapeTypes(node));
    if (!data) return null;
    const groupFloat = vmlFloat(poAttr(shape, 'style') ?? '', poIntAttr(shape, 'z-index'));
    return {
      kind: 'shape',
      data,
      ...(groupFloat ? { float: groupFloat } : {}),
      ...(poAttr(shape, 'alt')?.trim() ? { altText: poAttr(shape, 'alt')!.trim() } : {}),
    };
  }
  const tag = poTag(shape) ?? '';
  const style = poAttr(shape, 'style') ?? '';
  const width = vmlStyleLength(shape, 'width');
  const height = vmlStyleLength(shape, 'height');
  // A line states its ends rather than a box; everything else needs one.
  const from = tag === 'v:line' ? vmlPoint(poAttr(shape, 'from')) : undefined;
  const to = tag === 'v:line' ? vmlPoint(poAttr(shape, 'to')) : undefined;
  const boxW = width ?? (from && to ? Math.abs(to.x - from.x) : undefined);
  const boxH = height ?? (from && to ? Math.abs(to.y - from.y) : undefined);
  if (boxW === undefined || boxH === undefined) return null;

  const data = vmlShapeData(
    shape,
    tag,
    pt(Math.max(1, boxW)),
    pt(Math.max(1, boxH)),
    parseBody,
    vmlShapeTypes(node),
  );
  if (!data) return null;
  const float = vmlFloat(style, poIntAttr(shape, 'z-index'));
  return {
    kind: 'shape',
    data,
    ...(float ? { float } : {}),
    ...(poAttr(shape, 'alt')?.trim() ? { altText: poAttr(shape, 'alt')!.trim() } : {}),
  };
}

/**
 * One VML primitive as {@link ShapeData}: its preset geometry, fill, outline
 * and the words in its text box.
 *
 * @param shape     The `v:rect` / `v:oval` / … node.
 * @param tag       Its tag, which picks the preset.
 * @param width     The box's width, already in points.
 * @param height    Its height.
 * @param parseBody Body parser for a `v:textbox`.
 * @returns The shape, or `null` when it draws nothing and says nothing.
 */
/**
 * §14.1.2.19 `v:shapetype` — the reusable shape definitions declared alongside
 * the shapes that reference them by `@type="#id"`, keyed by id.
 *
 * @param node The `w:pict` / `w:object` node.
 * @returns The shapetypes found anywhere within it.
 */
function vmlShapeTypes(node: PoNode): ReadonlyMap<string, PoNode> {
  const out = new Map<string, PoNode>();
  const walk = (n: PoNode): void => {
    for (const c of poChildren(n)) {
      if (poIs(c, 'v:shapetype')) {
        const id = poAttr(c, 'id');
        if (id !== undefined) out.set(id, c);
      }
      walk(c);
    }
  };
  walk(node);
  return out;
}

function vmlShapeData(
  shape: PoNode,
  tag: string,
  width: Pt,
  height: Pt,
  parseBody?: ParseBody,
  shapeTypes?: ReadonlyMap<string, PoNode>,
): ShapeData | null {
  const typeRef = poAttr(shape, 'type')?.replace(/^#/u, '');
  const shapeType = typeRef !== undefined ? shapeTypes?.get(typeRef) : undefined;
  const fill = vmlFill(shape, shapeType);
  const line = vmlLine(shape, shapeType);
  const text = parseBody ? vmlTextBox(shape, parseBody) : undefined;
  // Nothing to draw and nothing to say: a spacer, not a shape.
  if (fill.kind === 'none' && !line && !text) return null;
  // §14.1.2.19 `style="rotation:N"` — VML turns a shape in whole degrees,
  // clockwise, the way `a:xfrm @rot` does in sixtieths of a thousandth.
  // fdo70838.docx stacks four rectangles at 75°, 105°, 255° and 285°, and
  // drawn square they came out as one wide box.
  const rotationDeg = vmlStyleNumber(shape, 'rotation');
  const transform =
    rotationDeg !== undefined && rotationDeg % 360 !== 0
      ? { rotation60k: Math.round(rotationDeg * 60000) }
      : undefined;
  return {
    width,
    height,
    geometry: {
      kind: 'preset',
      // §14.1.2.19 — a `v:shape` is whatever its shapetype draws. The straight
      // connector (`o:spt="32"`) is a LINE, and drawn as the rectangle every
      // other `v:shape` degrades to, fdo67737.docx's arrow came out as a long
      // thin box with the arrowhead on the wrong corner.
      preset:
        (shapeType && VML_SPT_PRESETS[poAttrLocal(shapeType, 'spt') ?? '']) ??
        VML_PRESETS[tag] ??
        'rect',
      adjust: new Map(),
    },
    fill,
    ...(line ? { line } : {}),
    ...(transform ? { transform } : {}),
    ...(text ? { text } : {}),
  };
}

/**
 * §14.1.2.7 `v:group` — a VML group: its `@style` gives the box on the page,
 * its `@coordsize`/`@coordorigin` the space its children are positioned in.
 * dml-textshape.docx draws its whole diagram inside one, and read as a single
 * shape it drew nothing at all.
 *
 * @param group     The `v:group` node.
 * @param parseBody Body parser for the members' text boxes.
 * @returns The group as a shape with members, or `null` when it has no box.
 */
function parseVmlGroup(
  group: PoNode,
  parseBody?: ParseBody,
  shapeTypes?: ReadonlyMap<string, PoNode>,
): ShapeData | null {
  const width = vmlStyleLength(group, 'width');
  const height = vmlStyleLength(group, 'height');
  if (width === undefined || height === undefined) return null;
  const size = vmlPair(poAttr(group, 'coordsize')) ?? { x: width, y: height };
  const origin = vmlPair(poAttr(group, 'coordorigin')) ?? { x: 0, y: 0 };
  const sx = size.x > 0 ? width / size.x : 1;
  const sy = size.y > 0 ? height / size.y : 1;

  const children: Array<ShapeGroupChild> = [];
  for (const child of poChildren(group)) {
    const tag = poTag(child) ?? '';
    if (!VML_SHAPE_TAGS.has(tag)) continue;
    // Inside a group the child's style is in the group's own coordinate units.
    const left = vmlStyleNumber(child, 'left') ?? 0;
    const top = vmlStyleNumber(child, 'top') ?? 0;
    const w = vmlStyleNumber(child, 'width');
    const h = vmlStyleNumber(child, 'height');
    if (w === undefined || h === undefined) continue;
    const data =
      tag === 'v:group'
        ? parseVmlGroup(child, parseBody, shapeTypes)
        : vmlShapeData(
            child,
            tag,
            pt(Math.max(1, w * sx)),
            pt(Math.max(1, h * sy)),
            parseBody,
            shapeTypes,
          );
    if (!data) continue;
    children.push({
      shape: { ...data, paragraphProperties: {} },
      xPt: pt((left - origin.x) * sx),
      yPt: pt((top - origin.y) * sy),
    });
  }
  if (children.length === 0) return null;
  return {
    width: pt(width),
    height: pt(height),
    children,
    geometry: { kind: 'custom', custom: EMPTY_GEOMETRY },
    fill: { kind: 'none' },
  };
}

const VML_SHAPE_TAGS = new Set(['v:rect', 'v:roundrect', 'v:oval', 'v:line', 'v:shape', 'v:group']);

/** A bare `x,y` attribute pair (`coordsize`, `coordorigin`). */
function vmlPair(raw: string | undefined): { x: number; y: number } | undefined {
  if (!raw) return undefined;
  const [a, b] = raw.split(',');
  const x = Number(a);
  const y = Number(b);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : undefined;
}

/** A unitless `@style` number (a child inside a group is in coord units). */
function vmlStyleNumber(shape: PoNode, prop: string): number | undefined {
  const style = poAttr(shape, 'style');
  if (style === undefined) return undefined;
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(-?[0-9.]+)`, 'iu').exec(style);
  if (!m) return undefined;
  const v = Number.parseFloat(m[1]!);
  return Number.isFinite(v) ? v : undefined;
}

// §14.1.2 — the VML primitives, as the preset geometry each corresponds to.
// A `v:shape` states its path in a `v:shapetype` of formulas we do not
// evaluate, so it is drawn as the box it occupies.
// §14.1.2.19 `o:spt` — the shapetype's own kind. VML states the geometry as a
// `path` of formulas we do not evaluate, but the primitives are numbered, and
// the number says which shape it is. fdo76016.docx draws an up arrow this way
// and we drew the rectangle every unrecognised `v:shape` degrades to.
const VML_SPT_PRESETS: Readonly<Record<string, string>> = {
  '1': 'rect',
  '2': 'roundRect',
  '3': 'ellipse',
  '4': 'diamond',
  '5': 'triangle', // isoceles
  '6': 'rtTriangle',
  '7': 'parallelogram',
  '8': 'trapezoid',
  '9': 'hexagon',
  '12': 'star5',
  '13': 'rightArrow',
  '20': 'line', // straight connector
  '32': 'line', // straight ARROW connector
  '56': 'pentagon',
  '58': 'star8',
  '59': 'star16',
  '60': 'star32',
  '66': 'leftArrow',
  '67': 'downArrow',
  '68': 'upArrow',
};

const VML_PRESETS: Readonly<Record<string, string>> = {
  'v:rect': 'rect',
  'v:roundrect': 'roundRect',
  'v:oval': 'ellipse',
  'v:line': 'line',
  'v:shape': 'rect',
};

/** A VML `from`/`to` point ("0,0" or "10pt,20pt"), in points. */
function vmlPoint(raw: string | undefined): { x: number; y: number } | undefined {
  if (!raw) return undefined;
  const parts = raw.split(',');
  if (parts.length !== 2) return undefined;
  const num = (t: string): number => {
    const m = /(-?[0-9.]+)\s*(pt|in|px|cm|mm|pc)?/iu.exec(t.trim());
    if (!m) return NaN;
    return Number.parseFloat(m[1]!) * (VML_UNIT_TO_PT[(m[2] ?? 'px').toLowerCase()] ?? 1);
  };
  const x = num(parts[0]!);
  const y = num(parts[1]!);
  return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : undefined;
}

/** §14.1.2.11 — a VML colour: `#rrggbb`, a named colour, or a system one. */
function vmlColor(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const v = raw.trim().replace(/^#/u, '').split(' ')[0] ?? '';
  if (/^[0-9A-Fa-f]{6}$/u.test(v)) return v.toUpperCase();
  if (/^[0-9A-Fa-f]{3}$/u.test(v)) {
    return [...v]
      .map((c) => c + c)
      .join('')
      .toUpperCase();
  }
  return VML_NAMED_COLORS[v.toLowerCase()];
}

const VML_NAMED_COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  silver: 'C0C0C0',
  gray: '808080',
  grey: '808080',
  white: 'FFFFFF',
  maroon: '800000',
  red: 'FF0000',
  purple: '800080',
  fuchsia: 'FF00FF',
  green: '008000',
  lime: '00FF00',
  olive: '808000',
  yellow: 'FFFF00',
  navy: '000080',
  blue: '0000FF',
  teal: '008080',
  aqua: '00FFFF',
  window: 'FFFFFF',
  windowtext: '000000',
};

// §14.1.2.5 — `@filled="f"` or a `<v:fill type="none">` says the shape is not
// filled; otherwise `@fillcolor`, defaulting to VML's own white.
function vmlFill(shape: PoNode, shapeType?: PoNode): ShapeFill {
  const filled = poAttr(shape, 'filled') ?? (shapeType ? poAttr(shapeType, 'filled') : undefined);
  if (filled === 'f' || filled === 'false') {
    return { kind: 'none' };
  }
  const fillEl = poChildren(shape).find((c) => poIs(c, 'v:fill'));
  if (fillEl && (poAttr(fillEl, 'type') === 'none' || poAttr(fillEl, 'on') === 'f')) {
    return { kind: 'none' };
  }
  const colorHex =
    vmlColor(poAttr(shape, 'fillcolor')) ??
    (fillEl ? vmlColor(poAttr(fillEl, 'color')) : undefined) ??
    vmlColor(shapeType ? poAttr(shapeType, 'fillcolor') : undefined);
  // §14.1.2.5 — a filled shape that names no colour is WHITE, not transparent.
  // fdo73215.docx draws its diagram as plain `v:rect`s inside a yellow one and
  // every unstated box let the yellow through.
  return { kind: 'solid', colorHex: colorHex ?? 'FFFFFF' };
}

// §14.1.2.21 — `@stroked="f"` says no outline; otherwise `@strokecolor` and
// `@strokeweight`, defaulting to VML's own hairline black.
function vmlLine(shape: PoNode, shapeType?: PoNode): ShapeLine | undefined {
  // §14.1.2.19 — a shape takes what its `v:shapetype` declares for anything it
  // does not state itself. The picture type `_x0000_t75` is `filled="f"
  // stroked="f"`, and read without it fdo61343.docx drew a black frame around
  // every picture in its canvas.
  const stroked =
    poAttr(shape, 'stroked') ?? (shapeType ? poAttr(shapeType, 'stroked') : undefined);
  if (stroked === 'f' || stroked === 'false') return undefined;
  const strokeEl = poChildren(shape).find((c) => poIs(c, 'v:stroke'));
  if (strokeEl && poAttr(strokeEl, 'on') === 'f') return undefined;
  const colorHex =
    vmlColor(poAttr(shape, 'strokecolor')) ??
    (strokeEl ? vmlColor(poAttr(strokeEl, 'color')) : undefined) ??
    '000000';
  // §14.1.2.21 `@startarrow`/`@endarrow` — VML's own arrowheads, spelled with
  // the same five shapes DrawingML names. fdo67737.docx ends its connector
  // with an open arrow and we drew a bare line.
  const arrow = (kind: 'start' | 'end'): LineEnd | undefined => {
    const type = VML_ARROWS.get(poAttr(strokeEl, `${kind}arrow`) ?? '');
    if (type === undefined) return undefined;
    const w = VML_ARROW_SIZES.get(poAttr(strokeEl, `${kind}arrowwidth`) ?? '');
    const len = VML_ARROW_SIZES.get(poAttr(strokeEl, `${kind}arrowlength`) ?? '');
    return { type, ...(w ? { width: w } : {}), ...(len ? { length: len } : {}) };
  };
  const headEnd = arrow('start');
  const tailEnd = arrow('end');
  const weight = poAttr(shape, 'strokeweight');
  const m = weight ? /(-?[0-9.]+)\s*(pt|in|px|cm|mm|pc)?/iu.exec(weight) : null;
  const widthPt = m
    ? Number.parseFloat(m[1]!) * (VML_UNIT_TO_PT[(m[2] ?? 'px').toLowerCase()] ?? 1)
    : 0.75;
  return {
    fill: 'solid',
    colorHex,
    width: pt(Number.isFinite(widthPt) ? widthPt : 0.75),
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  };
}

// §14.1.2.21 ST_StrokeArrowType → the DrawingML end this reader already draws.
const VML_ARROWS: ReadonlyMap<string, LineEnd['type']> = new Map([
  ['block', 'triangle'],
  ['classic', 'stealth'],
  ['diamond', 'diamond'],
  ['oval', 'oval'],
  ['open', 'arrow'],
]);

// §14.1.2.21 ST_StrokeArrowWidth / ST_StrokeArrowLength.
const VML_ARROW_SIZES: ReadonlyMap<string, 'sm' | 'med' | 'lg'> = new Map([
  ['narrow', 'sm'],
  ['short', 'sm'],
  ['medium', 'med'],
  ['wide', 'lg'],
  ['long', 'lg'],
]);

// §14.1.2.20 `v:textbox` — the shape's own words, in a `w:txbxContent`.
function vmlTextBox(shape: PoNode, parseBody: ParseBody): ShapeTextBody | undefined {
  const box = poChildren(shape).find((c) => poIs(c, 'v:textbox'));
  const content = box ? poFindDescendant(box, 'w:txbxContent') : undefined;
  if (!content) return undefined;
  const blocks = parseBody(poChildren(content));
  return blocks.length > 0 ? { content: blocks } : undefined;
}

/**
 * The `@style` of a positioned VML shape as a float anchor: `position:absolute`
 * with `margin-left`/`margin-top` (or `left`/`top`) against whatever
 * `mso-position-*-relative` names.
 *
 * @param style  The shape's `@style`.
 * @param zIndex The shape's z-index, when it has one.
 * @returns The anchor, or `undefined` for an inline shape.
 */
function vmlFloat(style: string, zIndex: number | undefined): FloatAnchor | undefined {
  if (!/position\s*:\s*absolute/iu.test(style)) return undefined;
  const prop = (name: string): number | undefined => {
    const m = new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*(-?[0-9.]+)(pt|in|px|cm|mm|pc)?`, 'iu').exec(
      style,
    );
    if (!m) return undefined;
    const v = Number.parseFloat(m[1]!) * (VML_UNIT_TO_PT[(m[2] ?? 'px').toLowerCase()] ?? 1);
    return Number.isFinite(v) ? v : undefined;
  };
  const rel = (name: string): string | undefined =>
    new RegExp(`mso-position-${name}-relative\\s*:\\s*([a-z-]+)`, 'iu').exec(style)?.[1];
  const hRaw = rel('horizontal');
  const vRaw = rel('vertical');
  // …and the position itself may be a KEYWORD rather than an offset:
  // docxopenhyperlinkbox.docx centres its box that way and we set it flush
  // left.
  const align = /mso-position-horizontal\s*:\s*(center|right|left)/iu.exec(style)?.[1] as
    | 'center'
    | 'right'
    | 'left'
    | undefined;
  const x = prop('margin-left') ?? prop('left') ?? 0;
  const y = prop('margin-top') ?? prop('top') ?? 0;
  return {
    wrap: 'square',
    ...(zIndex !== undefined && zIndex < 0 ? { behind: true } : {}),
    ...(zIndex !== undefined ? { zOrder: Math.abs(zIndex) } : {}),
    posH: {
      relativeFrom: hRaw === 'page' ? 'page' : hRaw === 'margin' ? 'margin' : 'column',
      ...(align ? { align } : { offsetPt: pt(x) }),
    },
    posV: {
      relativeFrom:
        vRaw === 'page'
          ? 'page'
          : vRaw === 'margin'
            ? 'margin'
            : vRaw === 'line'
              ? 'line'
              : 'paragraph',
      offsetPt: pt(y),
    },
  };
}

/**
 * §14.1.2.22 `v:textpath` — legacy WordArt: text bent along a preset path, whose
 * STRING lives in an attribute rather than in the document body. The path is a
 * `v:shapetype` of formulas we do not evaluate, so the words are set flat in
 * the shape's box — which is the whole of what the page says. Read nowhere,
 * WordArt.docx printed an empty page.
 *
 * @param node The `w:pict` / `w:object` node.
 * @returns The shape, or `null` when there is no textpath or no usable size.
 */
function parseVmlWordArt(node: PoNode): DrawingContent | null {
  // The `v:shapetype` template beside the shape carries a `v:textpath` of its
  // own, with no string on it — the words live on the SHAPE's.
  const shape = poFindDescendant(node, 'v:shape');
  const textpath = shape ? poFindDescendant(shape, 'v:textpath') : undefined;
  const string = textpath ? poAttr(textpath, 'string') : undefined;
  if (string === undefined || string === '') return null;
  const width = vmlStyleLength(shape, 'width');
  const height = vmlStyleLength(shape, 'height');
  if (width === undefined || height === undefined) return null;
  const fill = poAttr(shape, 'fillcolor');
  const colorHex =
    fill && /^#?[0-9A-Fa-f]{6}$/u.test(fill) ? fill.replace('#', '').toUpperCase() : undefined;
  // WordArt fills its box, and the parser has no font metrics to fit it with.
  // Half the height per line is close for a line of capitals; the width has
  // the last word, since a size that overflows it wraps and "WORD-ART" comes
  // out as "WORD-A / RT". A bold capital averages about 0.62em wide.
  const lines = string.split(/\r\n|[\r\n]/u);
  const longest = Math.max(1, ...lines.map((l) => l.length));
  const fontSizePt = pt(
    Math.max(1, Math.min((height / lines.length) * 0.5, width / (longest * 0.62))),
  );
  return {
    kind: 'shape',
    data: {
      width,
      height,
      geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
      fill: { kind: 'none' },
      text: {
        content: lines.map((line) => ({
          kind: 'paragraph' as const,
          paragraph: {
            properties: { alignment: 'center' as const },
            runs: [
              {
                text: line,
                properties: {
                  fontSizePt,
                  bold: true,
                  ...(colorHex ? { colorHex } : {}),
                },
              },
            ],
          },
        })),
        anchor: 'ctr' as const,
        insetLeft: pt(0),
        insetRight: pt(0),
        insetTop: pt(0),
        insetBottom: pt(0),
      },
    },
    ...(poAttr(shape, 'alt')?.trim() ? { altText: poAttr(shape, 'alt')!.trim() } : {}),
  };
}

// A VML @style length, normalised to points. VML inherits CSS units; Word
// emits pt, but in/px/cm/mm/pc all occur in the wild, and a bare number is a
// pixel count (the VML default).
const VML_UNIT_TO_PT: Readonly<Record<string, number>> = {
  pt: 1,
  in: 72,
  px: 0.75,
  cm: 72 / 2.54,
  mm: 72 / 25.4,
  pc: 12,
};
function vmlStyleLength(shape: PoNode | undefined, prop: 'width' | 'height'): Pt | undefined {
  const style = shape ? poAttr(shape, 'style') : undefined;
  if (style === undefined) return undefined;
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(-?[0-9.]+)(pt|in|px|cm|mm|pc)?`, 'i').exec(
    style,
  );
  if (!m) return undefined;
  const value = Number.parseFloat(m[1]!);
  if (!Number.isFinite(value) || value <= 0) return undefined;
  return pt(value * (VML_UNIT_TO_PT[(m[2] ?? 'px').toLowerCase()] ?? 1));
}

/**
 * §20.1.8.55 `a:srcRect` — the picture's own edges, cut away before it is
 * fitted to its frame. Each is a percentage in thousandths (ST_Percentage), so
 * `l="14711"` drops the left 14.711%.
 *
 * @param node The `a:srcRect` element, or `undefined` when the fill declares none.
 * @returns The crop, or `undefined` when nothing is cut away.
 */
function parseSrcRect(node: PoNode | undefined): ImageCrop | undefined {
  if (!node) return undefined;
  const edge = (name: string): number => {
    const v = Number(poAttr(node, name));
    // A negative srcRect pads rather than crops (the frame shows more than the
    // picture); nothing here draws that, so it reads as no cut on that edge.
    return Number.isFinite(v) && v > 0 ? v / 100000 : 0;
  };
  const crop = { left: edge('l'), top: edge('t'), right: edge('r'), bottom: edge('b') };
  if (crop.left + crop.top + crop.right + crop.bottom === 0) return undefined;
  // A crop that leaves nothing of either axis would divide by zero below; a
  // file that asks for it is asking for an empty frame, which is what it gets.
  if (crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) return undefined;
  return crop;
}

/**
 * §20.5.2.17 `wpg:wgp` — a drawing group: a box holding shapes of its own, in
 * its own coordinate space. `a:xfrm` gives the box (`a:off`/`a:ext`) and the
 * space its members are written in (`a:chOff`/`a:chExt`); a member at (x, y) in
 * that space lands at `off + (x − chOff) · ext / chExt` in the group's.
 *
 * The group draws nothing itself. Left unread, Tdf147485.docx — whose whole
 * picture is one — printed an empty page: the `mc:Fallback` beside the group is
 * VML, a geometry language of its own that we do not read either.
 *
 * @returns The group as a shape with children, or `null` when it states no size.
 */
/**
 * §20.3 — a `lc:lockedCanvas`: the same content model as a group, so it lays
 * out as one. Its own `a:grpSpPr/a:xfrm` states the child coordinate space the
 * members are placed in; the drawing's `wp:extent` states the box that space
 * is drawn into.
 */
function parseLockedCanvas(
  graphicData: PoNode,
  extentCx: number | undefined,
  extentCy: number | undefined,
  resolveColor: ColorResolver,
  parseBody?: ParseBody,
  resolveImage?: (relId: string) => ResourceId | undefined,
  themeLineWidths?: ReadonlyArray<number>,
): ShapeData | null {
  const canvas = expandMcChildren(poChildren(graphicData)).find((c) =>
    poIsLocal(c, 'lockedCanvas'),
  );
  if (!canvas) return null;
  const raw = groupChildren(canvas, resolveColor, parseBody, resolveImage, themeLineWidths);
  if (extentCx === undefined || extentCy === undefined) return null;
  // The members keep the offsets of the document they were copied from, and a
  // canvas states no child space to map them out of, so its content is drawn
  // from its OWN corner: what Word and LibreOffice both do with fdo43641.docx,
  // whose rectangle would otherwise sit 22pt above the frame it belongs in.
  const origin = contentOrigin(raw);
  const children = raw.map((c) => ({ ...c, xPt: pt(c.xPt - origin.x), yPt: pt(c.yPt - origin.y) }));
  return {
    width: emuToPt(extentCx),
    height: emuToPt(extentCy),
    ...(children.length > 0 ? { children } : {}),
    geometry: { kind: 'custom', custom: EMPTY_GEOMETRY },
    fill: { kind: 'none' },
  };
}

function parseWgp(
  graphicData: PoNode,
  extentCx: number | undefined,
  extentCy: number | undefined,
  resolveColor: ColorResolver,
  parseBody?: ParseBody,
  resolveImage?: (relId: string) => ResourceId | undefined,
  themeLineWidths?: ReadonlyArray<number>,
): ShapeData | null {
  const wgp = expandMcChildren(poChildren(graphicData)).find((c) => poIs(c, 'wpg:wgp'));
  if (!wgp) return null;
  const children = groupChildren(wgp, resolveColor, parseBody, resolveImage, themeLineWidths);
  const grpSpPr = poChildren(wgp).find((c) => poIs(c, 'wpg:grpSpPr'));
  const ext = grpSpPr
    ? poChildren(poChildren(grpSpPr).find((c) => poIs(c, 'a:xfrm')) ?? grpSpPr).find((c) =>
        poIs(c, 'a:ext'),
      )
    : undefined;
  const widthEmu = extentCx ?? (ext ? poIntAttr(ext, 'cx') : undefined);
  const heightEmu = extentCy ?? (ext ? poIntAttr(ext, 'cy') : undefined);
  if (widthEmu === undefined || heightEmu === undefined) return null;
  return {
    width: emuToPt(widthEmu),
    height: emuToPt(heightEmu),
    ...(children.length > 0 ? { children } : {}),
    // A group is a container: no geometry of its own — an empty command list,
    // so nothing is painted for the box itself.
    geometry: { kind: 'custom', custom: EMPTY_GEOMETRY },
    fill: { kind: 'none' },
  };
}

// The group's members, each mapped out of the child coordinate space into the
// group's own box. Nested groups recurse; a member without a size is skipped.
function groupChildren(
  wgp: PoNode,
  resolveColor: ColorResolver,
  parseBody: ParseBody | undefined,
  resolveImage?: (relId: string) => ResourceId | undefined,
  themeLineWidths?: ReadonlyArray<number>,
): Array<ShapeGroupChild> {
  // A locked canvas (§20.3) spells its children in the plain DrawingML
  // namespace — `a:sp`/`a:grpSp`/`a:pic` beside a WordprocessingGroup's
  // `wps:wsp`/`wpg:grpSp`/`pic:pic` — so members are matched by local name.
  const grpSpPr = poChildren(wgp).find((c) => poIsLocal(c, 'grpSpPr'));
  const xfrm = grpSpPr ? poChildren(grpSpPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  const at = (
    tag: string,
    ax: 'x' | 'cx',
    ay: 'y' | 'cy',
  ): { x: number; y: number } | undefined => {
    const n = xfrm ? poChildren(xfrm).find((c) => poIs(c, tag)) : undefined;
    const x = n ? poIntAttr(n, ax) : undefined;
    const y = n ? poIntAttr(n, ay) : undefined;
    return x !== undefined && y !== undefined ? { x, y } : undefined;
  };
  const ext = at('a:ext', 'cx', 'cy');
  const chExt = at('a:chExt', 'cx', 'cy');
  // With no child space declared the two are the same space, scale 1.
  const sx = ext && chExt && chExt.x > 0 ? ext.x / chExt.x : 1;
  const sy = ext && chExt && chExt.y > 0 ? ext.y / chExt.y : 1;
  // A canvas whose own transform is all zeros (fdo43641.docx writes
  // `a:ext`/`a:chExt` of 0×0) states no child space at all, and its members
  // keep the absolute offsets they had in the document they were copied from.
  // Word and LibreOffice both draw such a canvas from its content's own corner,
  // which is what taking the topmost-leftmost member as the origin does.
  const declared = at('a:chOff', 'x', 'y');
  const degenerate = xfrm !== undefined && (!chExt || chExt.x <= 0 || chExt.y <= 0);
  const chOff = degenerate
    ? memberOrigin(wgp, declared ?? { x: 0, y: 0 })
    : (declared ?? { x: 0, y: 0 });

  const out: Array<ShapeGroupChild> = [];
  for (const child of expandMcChildren(poChildren(wgp))) {
    const nested = poIsLocal(child, 'grpSp');
    const picture = poIsLocal(child, 'pic');
    if (!isGroupMember(child)) continue;
    const spPr = poChildren(child).find((c) => poIsLocal(c, nested ? 'grpSpPr' : 'spPr'));
    const childXfrm = spPr ? poChildren(spPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
    const cOff = childXfrm ? poChildren(childXfrm).find((c) => poIs(c, 'a:off')) : undefined;
    const cExt = childXfrm ? poChildren(childXfrm).find((c) => poIs(c, 'a:ext')) : undefined;
    const cx = cExt ? poIntAttr(cExt, 'cx') : undefined;
    const cy = cExt ? poIntAttr(cExt, 'cy') : undefined;
    if (cx === undefined || cy === undefined) continue;
    const data = nested
      ? parseNestedGroup(child, cx, cy, resolveColor, parseBody, resolveImage, themeLineWidths)
      : picture
        ? parseGroupPicture(child, cx, cy, resolveImage)
        : parseWspNode(child, cx, cy, resolveColor, parseBody, resolveImage, themeLineWidths);
    if (!data) continue;
    const x = (poIntAttr(cOff, 'x') ?? 0) - chOff.x;
    const y = (poIntAttr(cOff, 'y') ?? 0) - chOff.y;
    out.push({
      // Relative to the group's own corner, so `off` cancels out.
      shape: {
        ...data,
        width: emuToPt(cx * sx),
        height: emuToPt(cy * sy),
        paragraphProperties: {},
      },
      xPt: emuToPt(x * sx),
      yPt: emuToPt(y * sy),
    });
  }
  return out;
}

/**
 * The top-left corner the drawn content actually occupies, in the group's own
 * coordinates — nested members counted at their absolute position.
 *
 * @param children The group's members.
 * @returns The minimum x/y over the whole tree (0,0 when it is empty).
 */
function contentOrigin(children: ReadonlyArray<ShapeGroupChild>): { x: number; y: number } {
  let minX = Infinity;
  let minY = Infinity;
  const walk = (list: ReadonlyArray<ShapeGroupChild>, ox: number, oy: number): void => {
    for (const c of list) {
      const x = ox + c.xPt;
      const y = oy + c.yPt;
      // A group paints nothing itself — only what its members occupy counts.
      if (c.shape.children && c.shape.children.length > 0) {
        walk(c.shape.children, x, y);
        continue;
      }
      if (x < minX) minX = x;
      if (y < minY) minY = y;
    }
  };
  walk(children, 0, 0);
  return { x: Number.isFinite(minX) ? minX : 0, y: Number.isFinite(minY) ? minY : 0 };
}

/** A node that is a group MEMBER: a shape, a connector, a picture, or a nested group. */
function isGroupMember(node: PoNode): boolean {
  return (
    poIsLocal(node, 'wsp') ||
    poIsLocal(node, 'sp') ||
    // `a:cxnSp` is a connector: a shape whose geometry is a line.
    poIsLocal(node, 'cxnSp') ||
    poIsLocal(node, 'grpSp') ||
    poIsLocal(node, 'pic')
  );
}

/**
 * The top-left corner of a group's members, in the child coordinate space —
 * the origin to use when the group declares no child space of its own.
 *
 * @param wgp      The group / canvas node.
 * @param fallback What to return when no member states an offset.
 * @returns The minimum `a:off` over the members.
 */
function memberOrigin(wgp: PoNode, fallback: { x: number; y: number }): { x: number; y: number } {
  let minX = Infinity;
  let minY = Infinity;
  for (const child of expandMcChildren(poChildren(wgp))) {
    // Members only — the group's OWN `grpSpPr` sits among them and its
    // transform is the one we are trying to work around.
    if (!isGroupMember(child)) continue;
    const spPr = poChildren(child).find((c) => poIsLocal(c, 'spPr') || poIsLocal(c, 'grpSpPr'));
    const xfrm = poChildren(spPr).find((c) => poIs(c, 'a:xfrm'));
    const off = xfrm ? poChildren(xfrm).find((c) => poIs(c, 'a:off')) : undefined;
    const x = poIntAttr(off, 'x');
    const y = poIntAttr(off, 'y');
    if (x !== undefined && x < minX) minX = x;
    if (y !== undefined && y < minY) minY = y;
  }
  return {
    x: Number.isFinite(minX) ? minX : fallback.x,
    y: Number.isFinite(minY) ? minY : fallback.y,
  };
}

// A `wpg:grpSp` — a group inside a group. Same shape as the top-level one, but
// its box comes from its own a:xfrm rather than the drawing's wp:extent.
function parseNestedGroup(
  grpSp: PoNode,
  widthEmu: number,
  heightEmu: number,
  resolveColor: ColorResolver,
  parseBody: ParseBody | undefined,
  resolveImage?: (relId: string) => ResourceId | undefined,
  themeLineWidths?: ReadonlyArray<number>,
): ShapeData {
  const children = groupChildren(grpSp, resolveColor, parseBody, resolveImage, themeLineWidths);
  return {
    width: emuToPt(widthEmu),
    height: emuToPt(heightEmu),
    ...(children.length > 0 ? { children } : {}),
    geometry: { kind: 'custom', custom: EMPTY_GEOMETRY },
    fill: { kind: 'none' },
  };
}

// §20.1.8.14 — a `pic:pic` group member: a rectangle whose fill IS the picture.
// Left unread, WPGbodyPr.docx's group drew its three circles and lost the image
// standing inside them.
function parseGroupPicture(
  pic: PoNode,
  widthEmu: number,
  heightEmu: number,
  resolveImage: ((relId: string) => ResourceId | undefined) | undefined,
): ShapeData | null {
  const blip = poFindDescendant(pic, 'a:blip');
  const relId = blip ? poAttr(blip, 'embed') : undefined;
  const resource = relId !== undefined ? resolveImage?.(relId) : undefined;
  if (resource === undefined) return null;
  return {
    width: emuToPt(widthEmu),
    height: emuToPt(heightEmu),
    geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
    fill: { kind: 'picture', imageResource: resource },
  };
}

function parseWsp(
  graphicData: PoNode,
  extentCx: number | undefined,
  extentCy: number | undefined,
  resolveColor: ColorResolver,
  parseBody?: ParseBody,
  resolveImage?: (relId: string) => ResourceId | undefined,
  themeLineWidths?: ReadonlyArray<number>,
): ShapeData | null {
  // wps:wsp is normally a direct child, but a nested mc:AlternateContent can
  // wrap it — expand first so either layout is found.
  const wsp = expandMcChildren(poChildren(graphicData)).find((c) => poIs(c, 'wps:wsp'));
  if (!wsp) return null;
  return parseWspNode(
    wsp,
    extentCx,
    extentCy,
    resolveColor,
    parseBody,
    resolveImage,
    themeLineWidths,
  );
}

// One `wps:wsp`, given its box in EMU (from the drawing's wp:extent, or from
// the group member's own a:ext).
function parseWspNode(
  wsp: PoNode,
  extentCx: number | undefined,
  extentCy: number | undefined,
  resolveColor: ColorResolver,
  parseBody?: ParseBody,
  resolveImage?: (relId: string) => ResourceId | undefined,
  themeLineWidths?: ReadonlyArray<number>,
): ShapeData | null {
  const spPr = poChildren(wsp).find((c) => poIsLocal(c, 'spPr'));

  let geometry: ShapeGeometry = { kind: 'preset', preset: 'rect', adjust: new Map() };
  let fill: ShapeFill = { kind: 'none' };
  let line: ShapeLine | undefined;
  let transform: ShapeTransform | undefined;
  let widthEmu = extentCx;
  let heightEmu = extentCy;

  if (spPr) {
    const xfrm = poChildren(spPr).find((c) => poIs(c, 'a:xfrm'));
    if (xfrm) {
      transform = parseXfrm(xfrm);
      if (widthEmu === undefined || heightEmu === undefined) {
        const ext = poChildren(xfrm).find((c) => poIs(c, 'a:ext'));
        if (ext) {
          widthEmu = widthEmu ?? poIntAttr(ext, 'cx');
          heightEmu = heightEmu ?? poIntAttr(ext, 'cy');
        }
      }
    }
    const prst = poChildren(spPr).find((c) => poIs(c, 'a:prstGeom'));
    const cust = poChildren(spPr).find((c) => poIs(c, 'a:custGeom'));
    if (prst) geometry = parsePrstGeom(prst);
    else if (cust) geometry = parseCustGeom(cust);
    fill = parseFill(spPr, resolveColor, resolveImage);
    line = parseLine(spPr, resolveColor);
  }

  const text = parseBody ? parseTextBox(wsp, parseBody) : undefined;
  // §20.1.4.2.19/20.1.4.2.10 — a shape drawn from a gallery style keeps its
  // fill and outline in `<wps:style>` and carries none in `spPr` at all; read
  // alone, spPr says the shape has neither. TextEffects_Groupshapes.docx's
  // rectangle asks for accent1 that way and we drew its caption on white.
  const style = poChildren(wsp).find((c) => poIsLocal(c, 'style'));
  if (style && fill.kind === 'none') fill = styleRefFill(style, resolveColor);
  // §20.1.4.2.19 — a shape may state a WIDTH or a dash of its own and take its
  // COLOUR from the gallery style: dashed_line_custdash_percentage.docx rules a
  // 4.5pt accent-blue line that way and we drew a black hairline.
  const styleLine = style ? styleRefLine(style, resolveColor, themeLineWidths) : undefined;
  if (styleLine) {
    const own = line;
    line = own
      ? {
          ...styleLine,
          ...Object.fromEntries(Object.entries(own).filter(([, v]) => v !== undefined)),
        }
      : styleLine;
  }
  // §20.1.4.2.14 `<a:fontRef>` carries the colour the shape's text takes when
  // its own runs name none — on a gallery-drawn shape that is where the colour
  // IS. LineStyle_DashType.docx asks for `lt1` on seven blue rectangles and we
  // drew black on blue.
  const styled = style && text ? withStyleFontColor(text, style, resolveColor) : text;

  if (widthEmu === undefined || heightEmu === undefined) return null;
  return {
    width: emuToPt(widthEmu),
    height: emuToPt(heightEmu),
    geometry,
    fill,
    ...(line ? { line } : {}),
    ...(transform ? { transform } : {}),
    ...(styled ? { text: styled } : {}),
  };
}

// wps:txbx/w:txbxContent (the text body) + wps:bodyPr (insets + vertical
// anchor). Returns undefined when the shape carries no text.
// §20.1.4.2.13 `<a:fillRef>` — the fill a gallery style names. The theme's own
// `a:fillStyleLst` slot (which could make it a gradient) is out of reach here;
// the colour the reference names is what both references draw.
function styleRefFill(style: PoNode, resolveColor: ColorResolver): ShapeFill {
  const ref = poChildren(style).find((c) => poIs(c, 'a:fillRef'));
  if (!ref || poAttr(ref, 'idx') === '0') return { kind: 'none' };
  const child = firstElementChild(ref);
  const colorHex = child ? resolveColorNode(child, resolveColor) : undefined;
  return colorHex === undefined ? { kind: 'none' } : { kind: 'solid', colorHex };
}

// §20.1.4.2.19 `<a:lnRef>` — the outline a gallery style names. Its width lives
// in the theme's `a:lnStyleLst`, which is not reachable from here; the hairline
// below is what a shape with no stated width already draws.
function styleRefLine(
  style: PoNode,
  resolveColor: ColorResolver,
  themeLineWidths?: ReadonlyArray<number>,
): ShapeLine | undefined {
  const ref = poChildren(style).find((c) => poIs(c, 'a:lnRef'));
  if (!ref || poAttr(ref, 'idx') === '0') return undefined;
  const child = firstElementChild(ref);
  const colorHex = child ? resolveColorNode(child, resolveColor) : undefined;
  if (colorHex === undefined) return undefined;
  // §20.1.4.2.19 — `@idx` is a 1-based index into the theme's `a:lnStyleLst`,
  // which is where the WIDTH lives; the reference itself carries only the
  // colour. fdo66929.docx asks for the second style — 2pt in the standard
  // theme — and we drew every gallery outline as a 0.75pt hairline.
  const idx = poIntAttr(ref, 'idx');
  const width = idx !== undefined && idx > 0 ? themeLineWidths?.[idx - 1] : undefined;
  return { fill: 'solid', colorHex, width: pt(width ?? 0.75) };
}

// §20.1.4.2.14 — give every run that names no colour of its own the one the
// gallery style's `a:fontRef` names. The theme's colour is the FLOOR of the
// cascade (§17.7.2), so a run that could inherit one — through its own
// character style or its paragraph's — keeps what the style sheet gives it:
// ColorOverwritten.docx writes its arrow's two lines in a "red" and a "green"
// paragraph style, and stamping the theme's white over them left the shape
// blank.
function withStyleFontColor(
  text: ShapeTextBody,
  style: PoNode,
  resolveColor: ColorResolver,
): ShapeTextBody {
  const ref = poChildren(style).find((c) => poIs(c, 'a:fontRef'));
  const child = firstElementChild(ref);
  const colorHex = child ? resolveColorNode(child, resolveColor) : undefined;
  if (colorHex === undefined) return text;
  return {
    ...text,
    content: text.content.map((block) =>
      block.kind === 'paragraph'
        ? {
            ...block,
            paragraph: {
              ...block.paragraph,
              runs: block.paragraph.runs.map((run) =>
                run.properties.colorHex === undefined &&
                run.properties.styleId === undefined &&
                block.paragraph.properties.styleId === undefined
                  ? { ...run, properties: { ...run.properties, colorHex } }
                  : run,
              ),
            },
          }
        : block,
    ),
  };
}

function parseTextBox(wsp: PoNode, parseBody: ParseBody): ShapeTextBody | undefined {
  const txbx = poChildren(wsp).find((c) => poIs(c, 'wps:txbx'));
  if (!txbx) return undefined;
  const txContent = poChildren(txbx).find((c) => poIs(c, 'w:txbxContent'));
  if (!txContent) return undefined;
  const content = parseBody(poChildren(txContent));
  if (content.length === 0) return undefined;

  const bodyPr = poChildren(wsp).find((c) => poIs(c, 'wps:bodyPr'));
  const lIns = bodyPr ? poIntAttr(bodyPr, 'lIns') : undefined;
  const tIns = bodyPr ? poIntAttr(bodyPr, 'tIns') : undefined;
  const rIns = bodyPr ? poIntAttr(bodyPr, 'rIns') : undefined;
  const bIns = bodyPr ? poIntAttr(bodyPr, 'bIns') : undefined;
  const a = bodyPr ? poAttr(bodyPr, 'anchor') : undefined;
  const anchor: ShapeTextBody['anchor'] | undefined =
    a === 'ctr' ? 'ctr' : a === 'b' ? 'b' : a === 't' ? 't' : undefined;
  // §20.1.10.83 — the East-Asian stacked modes read top-to-bottom like `vert`
  // does; only the quarter turn differs, and `vert` is the closer of the two.
  const v = bodyPr ? poAttr(bodyPr, 'vert') : undefined;
  const vertical: ShapeTextBody['vertical'] | undefined =
    v === 'vert270' ? 'vert270' : v !== undefined && v !== 'horz' ? 'vert' : undefined;

  // §20.1.10.28 — the shape follows its text rather than the stated box.
  const autoFit = bodyPr !== undefined && poChildren(bodyPr).some((c) => poIs(c, 'a:spAutoFit'));

  return {
    content,
    ...(vertical ? { vertical } : {}),
    ...(autoFit ? { autoFit: true } : {}),
    ...(lIns !== undefined ? { insetLeft: emuToPt(lIns) } : {}),
    ...(tIns !== undefined ? { insetTop: emuToPt(tIns) } : {}),
    ...(rIns !== undefined ? { insetRight: emuToPt(rIns) } : {}),
    ...(bIns !== undefined ? { insetBottom: emuToPt(bIns) } : {}),
    ...(anchor ? { anchor } : {}),
  };
}

const isTrue = (v: string | undefined): boolean => v === '1' || v === 'true' || v === 'on';

/**
 * §20.1.7.6 `a:xfrm` — how a shape sits in its box: rotated about its centre by
 * `@rot` (sixtieth-thousandths of a degree) and mirrored by `@flipH`/`@flipV`.
 * The box itself comes from elsewhere — the anchor on a sheet, the extent in a
 * document — and holds the shape UNROTATED, which is what makes this a separate
 * transform rather than a different rectangle.
 *
 * @param xfrm The `a:xfrm` element.
 * @returns The transform; empty when it states neither rotation nor a flip.
 */
export function parseXfrm(xfrm: PoNode): ShapeTransform {
  const rot = poIntAttr(xfrm, 'rot');
  const flipH = isTrue(poAttr(xfrm, 'flipH'));
  const flipV = isTrue(poAttr(xfrm, 'flipV'));
  return {
    ...(rot !== undefined ? { rotation60k: rot } : {}),
    ...(flipH ? { flipH: true } : {}),
    ...(flipV ? { flipV: true } : {}),
  };
}

/**
 * Parse an `a:prstGeom` (§20.1.9.18) into a preset {@link ShapeGeometry}: the
 * `@prst` preset name plus the `a:avLst` adjust values (each `a:gd`'s `val …`
 * formula). Defaults to the `rect` preset when `@prst` is absent.
 */
export function parsePrstGeom(prst: PoNode): ShapeGeometry {
  const preset = poAttr(prst, 'prst') ?? 'rect';
  const adjust = new Map<string, number>();
  const avLst = poChildren(prst).find((c) => poIs(c, 'a:avLst'));
  if (avLst) {
    for (const gd of poChildren(avLst)) {
      if (!poIs(gd, 'a:gd')) continue;
      const nm = poAttr(gd, 'name');
      const fmla = poAttr(gd, 'fmla');
      if (!nm || !fmla) continue;
      const m = /^val\s+(-?\d+)/.exec(fmla);
      if (m) adjust.set(nm, Number(m[1]));
    }
  }
  return { kind: 'preset', preset, adjust };
}

/**
 * Parse an `a:custGeom` (ECMA-376 §20.1.9.11) → its first `<a:path>`
 * (§20.1.9.15) into a custom {@link ShapeGeometry}. Coordinates stay in
 * path-space (the geometry layer scales + y-flips them). Multiple subpaths with
 * differing `w`/`h` are a follow-up; falls back to a `rect` preset when the path
 * is empty or has no usable size.
 */
export function parseCustGeom(cust: PoNode): ShapeGeometry {
  const pathLst = poChildren(cust).find((c) => poIs(c, 'a:pathLst'));
  const path = pathLst ? poChildren(pathLst).find((c) => poIs(c, 'a:path')) : undefined;
  if (!path) return { kind: 'preset', preset: 'rect', adjust: new Map() };

  const pathWidth = poIntAttr(path, 'w') ?? 0;
  const pathHeight = poIntAttr(path, 'h') ?? 0;
  const commands: Array<CustomPathCmd> = [];
  for (const node of poChildren(path)) {
    switch (poTag(node)) {
      case 'a:moveTo': {
        const p = firstPt(node);
        if (p) commands.push({ cmd: 'move', x: p.x, y: p.y });
        break;
      }
      case 'a:lnTo': {
        const p = firstPt(node);
        if (p) commands.push({ cmd: 'line', x: p.x, y: p.y });
        break;
      }
      case 'a:cubicBezTo': {
        const p = pts(node);
        if (p.length >= 3)
          commands.push({
            cmd: 'cubic',
            x1: p[0]!.x,
            y1: p[0]!.y,
            x2: p[1]!.x,
            y2: p[1]!.y,
            x: p[2]!.x,
            y: p[2]!.y,
          });
        break;
      }
      case 'a:quadBezTo': {
        const p = pts(node);
        if (p.length >= 2)
          commands.push({ cmd: 'quad', x1: p[0]!.x, y1: p[0]!.y, x: p[1]!.x, y: p[1]!.y });
        break;
      }
      case 'a:arcTo':
        commands.push({
          cmd: 'arc',
          wR: poIntAttr(node, 'wR') ?? 0,
          hR: poIntAttr(node, 'hR') ?? 0,
          stAng: poIntAttr(node, 'stAng') ?? 0,
          swAng: poIntAttr(node, 'swAng') ?? 0,
        });
        break;
      case 'a:close':
        commands.push({ cmd: 'close' });
        break;
    }
  }

  if (pathWidth <= 0 || pathHeight <= 0 || commands.length === 0) {
    return { kind: 'preset', preset: 'rect', adjust: new Map() };
  }
  return { kind: 'custom', custom: { pathWidth, pathHeight, commands } };
}

// <a:pt x= y=> children of a path command.
function pts(node: PoNode): Array<{ x: number; y: number }> {
  const out: Array<{ x: number; y: number }> = [];
  for (const c of poChildren(node)) {
    if (!poIs(c, 'a:pt')) continue;
    const x = poIntAttr(c, 'x');
    const y = poIntAttr(c, 'y');
    if (x !== undefined && y !== undefined) out.push({ x, y });
  }
  return out;
}

const firstPt = (node: PoNode): { x: number; y: number } | undefined => pts(node)[0];

/**
 * Parse a shape's fill from its `a:spPr`: the first of `a:noFill`, `a:solidFill`
 * or `a:gradFill` wins. An unresolvable colour degrades to `{ kind: 'none' }`.
 *
 * @param spPr         The `wps:spPr` node.
 * @param resolveColor Resolver for theme/scheme colours.
 */
export function parseFill(
  spPr: PoNode,
  resolveColor: ColorResolver,
  resolveImage?: (relId: string) => ResourceId | undefined,
): ShapeFill {
  for (const child of poChildren(spPr)) {
    if (poIs(child, 'a:noFill')) return { kind: 'none' };
    // §20.1.8.14 — the shape is filled with a PICTURE. Unread, the shape fell
    // through to its gallery style's colour: crop-roundtrip.docx's photo came
    // out as a plain accent-orange rectangle.
    if (poIs(child, 'a:blipFill')) {
      const blip = poFindDescendant(child, 'a:blip');
      const relId = blip ? poAttr(blip, 'embed') : undefined;
      const resource = relId !== undefined ? resolveImage?.(relId) : undefined;
      if (resource === undefined) return { kind: 'none' };
      const crop = parseSrcRect(poFindDescendant(child, 'a:srcRect')) ?? fillRectCrop(child);
      return {
        kind: 'picture',
        imageResource: resource,
        ...(crop ? { imageCrop: crop } : {}),
      };
    }
    if (poIs(child, 'a:solidFill')) {
      const hex = colorFromContainer(child, resolveColor);
      return hex ? { kind: 'solid', colorHex: hex } : { kind: 'none' };
    }
    // §20.1.8.37 `a:pattFill` — a hatch of the foreground colour over the
    // background one. Drawn as the two blended by how much ink the pattern
    // lays down: the tile itself is beyond a vector fill, and an unfilled
    // shape is further from either reference than a tint of the right colour.
    // dml-shape-fillpattern.docx rules twelve rectangles this way and we drew
    // twelve empty boxes; dkvert.docx's cover bar came out blank.
    if (poIs(child, 'a:pattFill')) {
      const fg = poChildren(child).find((c) => poIs(c, 'a:fgClr'));
      const bg = poChildren(child).find((c) => poIs(c, 'a:bgClr'));
      const fgHex = fg ? colorFromContainer(fg, resolveColor) : undefined;
      const bgHex = bg ? colorFromContainer(bg, resolveColor) : 'FFFFFF';
      if (!fgHex) return { kind: 'none' };
      return {
        kind: 'solid',
        colorHex: blendHex(fgHex, bgHex ?? 'FFFFFF', patternCoverage(poAttr(child, 'prst'))),
      };
    }
    if (poIs(child, 'a:gradFill')) {
      const gradient = parseGradient(child, resolveColor);
      return gradient ? { kind: 'gradient', gradient } : { kind: 'none' };
    }
  }
  return { kind: 'none' };
}

/**
 * §20.1.8.37 — roughly how much of a tile the preset pattern inks, which is
 * what decides how strongly the foreground colour shows through. The families
 * carry the answer: `pctNN` states it outright, `lt*` is sparse, `dk*` dense,
 * and the woven ones cover about half.
 *
 * @param prst The `a:pattFill` `@prst` value.
 * @returns The ink coverage, 0..1.
 */
function patternCoverage(prst: string | undefined): number {
  if (!prst) return 0.25;
  const pct = /^pct(\d+)$/u.exec(prst);
  if (pct) return Math.min(1, Number(pct[1]) / 100);
  if (prst.startsWith('lt')) return 0.15;
  if (prst.startsWith('dk')) return 0.5;
  if (prst.startsWith('wd')) return 0.2;
  if (prst.startsWith('nar')) return 0.35;
  if (prst.endsWith('Grid') || prst === 'divot' || prst === 'wave') return 0.15;
  if (DENSE_PATTERNS.has(prst)) return 0.5;
  return 0.25;
}

const DENSE_PATTERNS = new Set([
  'trellis',
  'weave',
  'plaid',
  'smCheck',
  'lgCheck',
  'solidDmnd',
  'sphere',
  'dkCross',
  'dkDiagCross',
]);

/** `a` mixed over `b` at `t` (0 = all `b`, 1 = all `a`), as 6-hex. */
function blendHex(a: string, b: string, t: number): string {
  const ch = (hex: string, i: number): number => parseInt(hex.slice(i, i + 2), 16);
  const mix = (i: number): string =>
    Math.round(ch(a, i) * t + ch(b, i) * (1 - t))
      .toString(16)
      .padStart(2, '0');
  return (mix(0) + mix(2) + mix(4)).toUpperCase();
}

/**
 * §20.1.8.30 `a:stretch/a:fillRect` — where the picture's edges land relative
 * to the shape's box, in 1000ths of a percent of that box. A NEGATIVE inset
 * pushes the edge outside the box, which is how a fill zooms in: what remains
 * inside is the part the box shows, so the insets convert straight to the crop
 * fractions of the SOURCE. crop-roundtrip.docx frames the middle of its photo
 * that way and we drew the whole of it.
 *
 * @param blipFill The `a:blipFill` node.
 * @returns The crop, or `undefined` when the fill states no `a:fillRect`.
 */
function fillRectCrop(blipFill: PoNode): ImageCrop | undefined {
  const stretch = poChildren(blipFill).find((c) => poIs(c, 'a:stretch'));
  const rect = stretch ? poChildren(stretch).find((c) => poIs(c, 'a:fillRect')) : undefined;
  if (!rect) return undefined;
  const side = (name: string): number => (poIntAttr(rect, name) ?? 0) / 100000;
  const l = side('l');
  const t = side('t');
  const r = side('r');
  const b = side('b');
  if (l === 0 && t === 0 && r === 0 && b === 0) return undefined;
  const spanX = 1 - l - r;
  const spanY = 1 - t - b;
  if (spanX <= 0 || spanY <= 0) return undefined;
  const clamp = (v: number): number => Math.min(0.99, Math.max(0, v));
  return {
    left: clamp(-l / spanX),
    top: clamp(-t / spanY),
    right: clamp(-r / spanX),
    bottom: clamp(-b / spanY),
  };
}

/**
 * Parse a shape's drop shadow from its `a:spPr` (§20.1.8.40 `a:effectLst` →
 * `a:outerShdw`).
 *
 * The spec states the displacement in polar form: `dist` in EMU and `dir` in
 * 60 000ths of a degree, measured clockwise from due east in a frame whose y
 * grows DOWNWARD — so the standard 2 700 000 (45°) puts the shadow down and to
 * the right, which is where every reader draws it.
 *
 * @param spPr         The shape's `spPr` node.
 * @param resolveColor Resolver for theme/scheme colours.
 * @returns The shadow, or undefined when the shape declares none.
 */
export function parseShadow(spPr: PoNode, resolveColor: ColorResolver): ShapeShadow | undefined {
  const list = poChildren(spPr).find((c) => poIs(c, 'a:effectLst'));
  const shdw = list ? poChildren(list).find((c) => poIs(c, 'a:outerShdw')) : undefined;
  return shdw ? shadowFromOuterShdw(shdw, resolveColor) : undefined;
}

/**
 * Build a {@link ShapeShadow} from an `a:outerShdw` node.
 *
 * @param shdw         The `a:outerShdw` node.
 * @param resolveColor Resolver for theme/scheme colours.
 * @returns The shadow, or undefined when its colour will not resolve.
 */
export function shadowFromOuterShdw(
  shdw: PoNode,
  resolveColor: ColorResolver,
): ShapeShadow | undefined {
  const colorNode = poChildren(shdw).find((c) => resolveColorNode(c, resolveColor) !== undefined);
  const colorHex = colorNode ? resolveColorNode(colorNode, resolveColor) : undefined;
  if (!colorHex) return undefined;
  const alphaMod = colorNode ? readColorMods(colorNode).find((m) => m.kind === 'alpha') : undefined;
  const dist = emuToPt(poIntAttr(shdw, 'dist') ?? 0);
  // §20.1.10.13 ST_PositiveFixedAngle — 60 000ths of a degree.
  const dirDeg = (poIntAttr(shdw, 'dir') ?? 0) / 60000;
  const rad = (dirDeg * Math.PI) / 180;
  return {
    dxPt: dist * Math.cos(rad),
    dyPt: dist * Math.sin(rad),
    blurPt: emuToPt(poIntAttr(shdw, 'blurRad') ?? 0),
    colorHex,
    alpha: alphaMod ? alphaMod.val : 1,
  };
}

/**
 * Parse a shape's outline (`a:ln`) from its `a:spPr` into a {@link ShapeLine}:
 * width, cap, solid colour, dash pattern, and an explicit `a:noFill` (an unstroked
 * outline). Returns `undefined` when the shape has no `a:ln`.
 *
 * @param spPr         The `wps:spPr` node.
 * @param resolveColor Resolver for theme/scheme colours.
 */
export function parseLine(spPr: PoNode, resolveColor: ColorResolver): ShapeLine | undefined {
  const ln = poChildren(spPr).find((c) => poIs(c, 'a:ln'));
  if (!ln) return undefined;
  const widthEmu = poIntAttr(ln, 'w');
  // a:ln @cap: flat | rnd | sq (§20.1.10.31).
  const capRaw = poAttr(ln, 'cap');
  const cap: ShapeLine['cap'] | undefined =
    capRaw === 'rnd'
      ? 'round'
      : capRaw === 'sq'
        ? 'square'
        : capRaw === 'flat'
          ? 'flat'
          : undefined;
  let noFill = false;
  let colorHex: string | undefined;
  let dash: ShapeDash | undefined;
  // §20.1.8.21 — the author's own pattern, each length a percentage of the
  // line's width. dashed_line_custdash_percentage.docx rules a dash-dot-dot
  // line that way and we drew it solid.
  let customDash: Array<number> | undefined;
  let headEnd: LineEnd | undefined;
  let tailEnd: LineEnd | undefined;
  for (const c of poChildren(ln)) {
    if (poIs(c, 'a:noFill')) noFill = true;
    else if (poIs(c, 'a:solidFill')) colorHex = colorFromContainer(c, resolveColor);
    else if (poIs(c, 'a:prstDash')) dash = normalizeDash(poAttr(c, 'val'));
    else if (poIs(c, 'a:custDash')) customDash = parseCustDash(c);
    else if (poIs(c, 'a:headEnd')) headEnd = parseLineEnd(c);
    else if (poIs(c, 'a:tailEnd')) tailEnd = parseLineEnd(c);
  }
  return {
    ...(widthEmu !== undefined ? { width: emuToPt(widthEmu) } : {}),
    ...(colorHex ? { colorHex } : {}),
    ...(dash ? { dash } : {}),
    ...(customDash && customDash.length > 0 ? { customDash } : {}),
    ...(cap ? { cap } : {}),
    ...(noFill ? { fill: 'none' as const } : {}),
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  };
}

const LINE_END_TYPES = new Set(['triangle', 'stealth', 'diamond', 'oval', 'arrow']);
const LINE_END_SIZES = new Set(['sm', 'med', 'lg']);

/**
 * §20.1.8.24 / §20.1.8.42 — one end decoration: `@type` plus the `@w`/`@len`
 * size steps. `none` (and an unknown type) leaves the end bare.
 *
 * @param node The `a:headEnd` / `a:tailEnd` node.
 * @returns The parsed end, or `undefined` when nothing is drawn there.
 */
function parseLineEnd(node: PoNode): LineEnd | undefined {
  const type = poAttr(node, 'type');
  if (type === undefined || !LINE_END_TYPES.has(type)) return undefined;
  const w = poAttr(node, 'w');
  const len = poAttr(node, 'len');
  const size = (v: string | undefined): 'sm' | 'med' | 'lg' | undefined =>
    v !== undefined && LINE_END_SIZES.has(v) ? (v as 'sm' | 'med' | 'lg') : undefined;
  const width = size(w);
  const length = size(len);
  return {
    type: type as LineEnd['type'],
    ...(width ? { width } : {}),
    ...(length ? { length } : {}),
  };
}

/**
 * §20.1.8.21 `a:custDash` — the dash/space lengths of the author's own
 * pattern. Each `a:ds` states them in 1000ths of a percent of the line's
 * width, so the numbers here are plain multiples of that width.
 *
 * @param custDash The `a:custDash` node.
 * @returns The pattern as [dash, space, dash, space, …], or `undefined`.
 */
function parseCustDash(custDash: PoNode): Array<number> | undefined {
  const out: Array<number> = [];
  for (const ds of poChildren(custDash)) {
    if (!poIs(ds, 'a:ds')) continue;
    const d = (poIntAttr(ds, 'd') ?? 0) / 100000;
    const sp = (poIntAttr(ds, 'sp') ?? 0) / 100000;
    if (d <= 0 && sp <= 0) continue;
    out.push(Math.max(0, d), Math.max(0, sp));
  }
  return out.length > 0 ? out : undefined;
}

const DASH_VALUES = new Set<ShapeDash>([
  'solid',
  'dot',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
  'sysDash',
  'sysDot',
]);

// Map a:prstDash @val to a supported ShapeDash, folding rarer variants onto
// their nearest supported pattern.
function normalizeDash(v: string | undefined): ShapeDash | undefined {
  if (!v) return undefined;
  if (DASH_VALUES.has(v as ShapeDash)) return v as ShapeDash;
  if (v === 'lgDashDotDot') return 'lgDashDot';
  if (v === 'sysDashDot' || v === 'sysDashDotDot') return 'sysDash';
  return undefined;
}

// First a:srgbClr / a:schemeClr child → resolved hex (with colour transforms).
function colorFromContainer(parent: PoNode, resolveColor: ColorResolver): string | undefined {
  for (const c of poChildren(parent)) {
    const hex = resolveColorNode(c, resolveColor);
    if (hex) return hex;
  }
  return undefined;
}

// a:gradFill → a gradient fill (EP16). Reads the a:gsLst/a:gs stops (each with a
// @pos in 1000ths of a percent and a colour child), and the direction from a:lin
// (@ang in 60000ths of a degree, clockwise) or a:path (a radial/path gradient).
function parseGradient(grad: PoNode, resolveColor: ColorResolver): ShapeGradient | undefined {
  const gsLst = poChildren(grad).find((c) => poIs(c, 'a:gsLst'));
  if (!gsLst) return undefined;
  const stops: Array<GradientStop> = [];
  for (const gs of poChildren(gsLst)) {
    if (!poIs(gs, 'a:gs')) continue;
    let hex: string | undefined;
    for (const c of poChildren(gs)) {
      hex = resolveColorNode(c, resolveColor);
      if (hex) break;
    }
    if (!hex) continue;
    const pos = poIntAttr(gs, 'pos');
    const offset = pos !== undefined ? clampUnit(pos / 100000) : stops.length === 0 ? 0 : 1;
    stops.push({ offset, colorHex: hex });
  }
  if (stops.length === 0) return undefined;
  stops.sort((a, b) => a.offset - b.offset);
  if (poChildren(grad).some((c) => poIs(c, 'a:path'))) return { kind: 'radial', stops };
  const lin = poChildren(grad).find((c) => poIs(c, 'a:lin'));
  const ang = lin ? poIntAttr(lin, 'ang') : undefined;
  const angle = ang !== undefined ? (ang / 60000) % 360 : 0;
  return { kind: 'linear', angle, stops };
}

function clampUnit(x: number): number {
  return x < 0 ? 0 : x > 1 ? 1 : x;
}
