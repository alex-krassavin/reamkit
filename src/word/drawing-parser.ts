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
  PictureOutline,
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
import { buildStroke } from '@/core/drawingml/shape-render';
import { placeholderColors, readColorMods, resolveColorNode } from '@/core/drawingml/colors';
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
// ISO 29500 STRICT writes the same DrawingML part URIs under purl.oclc.org,
// which is why strict.docx's chart and SmartArt reached the page as nothing at
// all while its picture — matched by element name, not by URI — came through.
const DRAWINGML_URI_BASES = [
  'http://schemas.openxmlformats.org/drawingml/2006/',
  'http://purl.oclc.org/ooxml/drawingml/',
] as const;

function isDrawingMlUri(uri: string | undefined, name: string): boolean {
  return uri !== undefined && DRAWINGML_URI_BASES.some((base) => uri === base + name);
}
const WPG_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
// §20.3 — a canvas of shapes, written in the plain DrawingML namespace. Word
// exports a pasted PowerPoint group this way.
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
 * §21.1.2 `a:txBody` — the DrawingML text body a shape may carry instead of a
 * `wps:txbx`. Supplied by the caller (the reader owns the PresentationML text
 * reader that knows how to walk it), so this module stays free of it.
 */
export type ParseDrawingText = (
  txBody: PoNode,
  resolveColor: ColorResolver,
) => ShapeTextBody | undefined;

/**
 * §20.1.4.1.14/§20.1.4.1.15 — the theme's format scheme, as the raw nodes a
 * `<wps:style>` reference indexes: the fill styles (`a:fillStyleLst`), the
 * BACKGROUND fill styles (`a:bgFillStyleLst`, which an `a:fillRef` reaches with
 * an index past 1000) and the effect styles (`a:effectStyleLst`).
 */
export interface ThemeStyles {
  readonly fills?: ReadonlyArray<PoNode>;
  readonly bgFills?: ReadonlyArray<PoNode>;
  readonly effects?: ReadonlyArray<PoNode>;
}

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
      /** §20.1.2.2.24 `a:ln` on `pic:spPr` / VML `@stroked` — the picture's frame. */
      readonly outline?: PictureOutline;
      /** §20.1.8.40 `a:outerShdw` on the same `pic:spPr` — the picture's shadow. */
      readonly shadow?: ShapeShadow;
      /** §14.1.2.10 `@gain`/`@blacklevel` — the wash the picture is drawn through. */
      readonly wash?: { readonly gain: number; readonly black: number };
      /** §20.1.8.55 `a:srcRect` — the part of the source the frame shows. */
      readonly crop?: ImageCrop;
      /** §20.1.7.6 `a:xfrm @rot` — the picture's rotation (1/60000°, clockwise). */
      readonly rotation60k?: number;
      /** §20.1.7.6 `a:xfrm @flipH/@flipV` — the picture drawn mirrored. */
      readonly flipH?: boolean;
      readonly flipV?: boolean;
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
  // §20.4.2.3 `@wrapText` — which side(s) the text may stand on. `bothSides`
  // is the default, and the one that fills the gap on EACH side of a drawing.
  let wrapSide: FloatAnchor['wrapSide'];
  for (const child of poChildren(anchor)) {
    if (poIs(child, 'wp:wrapSquare')) wrap = 'square';
    else if (poIs(child, 'wp:wrapTight')) wrap = 'tight';
    else if (poIs(child, 'wp:wrapThrough')) wrap = 'through';
    else if (poIs(child, 'wp:wrapTopAndBottom')) wrap = 'topAndBottom';
    else continue;
    const side = poAttr(child, 'wrapText');
    wrapSide = side === 'left' || side === 'right' || side === 'largest' ? side : 'bothSides';
  }
  const zOrder = poIntAttr(anchor, 'relativeHeight');
  // §20.4.3.3 — `leftMargin`/`rightMargin` name the margin band on that side,
  // the place a marginal note belongs.
  const posH = parseAnchorPos(anchor, 'wp:positionH', [
    'margin',
    'page',
    'column',
    'leftMargin',
    'rightMargin',
  ]);
  // §20.4.3.4 — `insideMargin`/`outsideMargin` are the top/bottom margin on a
  // page that is not part of a facing pair, which is how Word writes them.
  const posV = parseAnchorPos(anchor, 'wp:positionV', [
    'margin',
    'page',
    'paragraph',
    'line',
    'topMargin',
    'bottomMargin',
  ]);
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
  // §20.4.2.3 `@layoutInCell` — inside a table cell, the cell is the frame the
  // position is measured in. Turned off (tdf129888dml.docx) the object reaches
  // past the table to the page it names, which is how a page number ends up in
  // the corner of the paper rather than the corner of a cell.
  const inCellRaw = poAttr(anchor, 'layoutInCell');
  const inCell = !(inCellRaw === '0' || inCellRaw === 'false');
  return {
    wrap,
    ...(wrapSide ? { wrapSide } : {}),
    ...(inCell ? {} : { inCell: false }),
    ...(behind ? { behind: true } : {}),
    ...(zOrder !== undefined ? { zOrder } : {}),
    ...(anyDist ? { wrapDist } : {}),
    ...(posH ? { posH: posH as NonNullable<FloatAnchor['posH']> } : {}),
    ...(posV ? { posV: posV as NonNullable<FloatAnchor['posV']> } : {}),
  };
}

const ANCHOR_ALIGNS = new Set(['left', 'center', 'right']);
const ANCHOR_V_ALIGNS = new Set(['top', 'center', 'bottom']);

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
  const read = (tag: string, pctTag: string): { pct: number; from: string } | undefined => {
    const node = expandMcChildren(poChildren(anchor)).find((c) => poIs(c, tag));
    if (!node) return undefined;
    const pctNode = poChildren(node).find((c) => poIs(c, pctTag));
    const raw = pctNode ? Number(poText(pctNode).trim()) : NaN;
    if (!Number.isFinite(raw) || raw <= 0) return undefined;
    // §20.4.3.6/§20.4.3.7 — `insideMargin`/`outsideMargin` are the left/top
    // margin on a page that is not part of a facing pair, which is how Word
    // writes them.
    const from = poAttr(node, 'relativeFrom') ?? 'margin';
    return { pct: raw / 100000, from };
  };
  const h = read('wp14:sizeRelH', 'wp14:pctWidth');
  const v = read('wp14:sizeRelV', 'wp14:pctHeight');
  if (!h && !v) return undefined;
  return {
    ...(h ? { widthPct: h.pct, widthFrom: relFromH(h.from) } : {}),
    ...(v ? { heightPct: v.pct, heightFrom: relFromV(v.from) } : {}),
  };
}

// §20.4.3.6 ST_SizeRelFromH — the bases we place a width against; anything
// else (a character, a gutter) falls back to the text area.
function relFromH(raw: string): NonNullable<RelativeSize['widthFrom']> {
  if (raw === 'page') return 'page';
  if (raw === 'leftMargin' || raw === 'insideMargin') return 'leftMargin';
  if (raw === 'rightMargin' || raw === 'outsideMargin') return 'rightMargin';
  return 'margin';
}

// §20.4.3.7 ST_SizeRelFromV — the same for a height.
function relFromV(raw: string): NonNullable<RelativeSize['heightFrom']> {
  if (raw === 'page') return 'page';
  if (raw === 'topMargin' || raw === 'insideMargin') return 'topMargin';
  if (raw === 'bottomMargin' || raw === 'outsideMargin') return 'bottomMargin';
  return 'margin';
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
    // §20.4.3.1 — the vertical keywords are their own set.
    ...(tag === 'wp:positionV' && ANCHOR_V_ALIGNS.has(alignRaw) ? { align: alignRaw } : {}),
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
  themeStyles?: ThemeStyles,
  // §21.1.2 — the DrawingML text reader, for a shape that carries an `a:txBody`.
  parseDrawingText?: ParseDrawingText,
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
      themeStyles,
      parseDrawingText,
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
      themeStyles,
      parseDrawingText,
    );
    if (!data) return null;
    return { kind: 'shape', data: { ...data, ...(relativeSize ? { relativeSize } : {}) }, ...alt };
  }

  // §20.3 `lc:lockedCanvas` — a group by another name: the same members, in the
  // `a:` namespace. fdo43641.docx draws its rectangle and arrow inside one and
  // we rendered an empty page.
  if (graphicData && isDrawingMlUri(uri, 'lockedCanvas')) {
    const data = parseLockedCanvas(
      graphicData,
      extentCx,
      extentCy,
      resolveColor,
      parseBody,
      resolveImage,
      themeLineWidths,
      themeStyles,
      parseDrawingText,
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
    const children = groupChildren(
      canvas,
      resolveColor,
      parseBody,
      resolveImage,
      themeLineWidths,
      themeStyles,
      parseDrawingText,
    );
    const data: ShapeData = {
      width: emuToPt(extentCx),
      height: emuToPt(extentCy),
      ...(children.length > 0 ? { children } : {}),
      geometry: { kind: 'custom', custom: EMPTY_GEOMETRY },
      fill: { kind: 'none' },
    };
    return { kind: 'shape', data: { ...data, ...(relativeSize ? { relativeSize } : {}) }, ...alt };
  }

  if (graphicData && isDrawingMlUri(uri, 'chart')) {
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
  if (graphicData && isDrawingMlUri(uri, 'diagram')) {
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
    // §20.1.7.6 — …and mirrored in its frame. graphic-object-fliph.docx turns
    // its folded corner to the other side and we drew it unflipped.
    const on = (name: string): boolean =>
      xfrm !== undefined && (poAttr(xfrm, name) === '1' || poAttr(xfrm, name) === 'true');
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
    // §20.1.2.2.24 — Word's "Picture Border" is an `a:ln` on the picture's own
    // `pic:spPr`. tdf125657.docx frames its screenshot that way and we drew the
    // picture bare.
    const spPr = poFindDescendant(anchor, 'pic:spPr');
    const outline = spPr ? pictureOutline(parseLine(spPr, resolveColor)) : undefined;
    // §20.1.8.40 — and the shadow on that same `pic:spPr`, which is how Word
    // writes the drop shadow of a pasted screenshot (imgshadow.docx).
    const shadow = spPr ? parseShadow(spPr, resolveColor) : undefined;
    return {
      kind: 'image',
      imageId: rId,
      width: emuToPt(extentCx),
      height: emuToPt(extentCy),
      ...(outline ? { outline } : {}),
      ...(shadow ? { shadow } : {}),
      ...(crop ? { crop } : {}),
      ...(rot ? { rotation60k: rot } : {}),
      ...(on('flipH') ? { flipH: true } : {}),
      ...(on('flipV') ? { flipV: true } : {}),
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
    ? vmlFloat(poAttr(shape, 'style') ?? '', poIntAttr(shape, 'z-index'), shape)
    : undefined;
  // §14.1.2.21 — a picture is framed only when it SAYS so. The picture
  // shapetype `_x0000_t75` is `stroked="f"`, and a `v:shape` that names it
  // without the package declaring it (fdo81031.docx) has nothing to inherit
  // from — so the frame is drawn on an explicit `stroked="t"` alone, which is
  // what fdo79915.docx's diagram states and what we drew it without.
  const typeRef = shape ? poAttr(shape, 'type')?.replace(/^#/u, '') : undefined;
  const shapeType = typeRef !== undefined ? vmlShapeTypes(node).get(typeRef) : undefined;
  const stroked = shape
    ? (poAttr(shape, 'stroked') ?? (shapeType ? poAttr(shapeType, 'stroked') : undefined))
    : undefined;
  const outline =
    shape && (stroked === 't' || stroked === 'true')
      ? pictureOutline(vmlLine(shape, shapeType))
      : undefined;
  // §14.1.2.10 — Word washes a watermark out with `@gain` (contrast) and
  // `@blacklevel` (brightness), each a fraction, a percentage or the
  // 1/65536ths it writes. pictureWatermark.docx prints its penguins at 30%
  // contrast lifted 35%, and drawn at full strength the photograph sat behind
  // the text instead of a ghost of it.
  const wash = vmlWash(imagedata);
  return {
    kind: 'image',
    imageId,
    width,
    height,
    ...(outline ? { outline } : {}),
    ...(wash ? { wash } : {}),
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
    const groupFloat = vmlFloat(poAttr(shape, 'style') ?? '', poIntAttr(shape, 'z-index'), shape);
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
  const float = vmlFloat(style, poIntAttr(shape, 'z-index'), shape);
  // §14.1.2.13 — a LINE states its ends, and those ends are where it is: its
  // style carries no `left`/`top` to place it by. tdf129888vml.docx rules a
  // margin line from 191pt across, and placed at the anchor's own zero it ran
  // down the very edge of the paper.
  const placed = float && from && to ? vmlLinePlaced(float, from, to) : float;
  return {
    kind: 'shape',
    data,
    ...(placed ? { float: placed } : {}),
    ...(poAttr(shape, 'alt')?.trim() ? { altText: poAttr(shape, 'alt')!.trim() } : {}),
  };
}

// The anchor a `v:line` is really at: its top-left corner, which its ends give
// and its style does not.
function vmlLinePlaced(
  float: FloatAnchor,
  from: { x: number; y: number },
  to: { x: number; y: number },
): FloatAnchor {
  const x = Math.min(from.x, to.x);
  const y = Math.min(from.y, to.y);
  return {
    ...float,
    ...(float.posH?.align === undefined
      ? { posH: { ...(float.posH ?? { relativeFrom: 'column' as const }), offsetPt: pt(x) } }
      : {}),
    ...(float.posV?.align === undefined
      ? { posV: { ...(float.posV ?? { relativeFrom: 'paragraph' as const }), offsetPt: pt(y) } }
      : {}),
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
  // §14.1.2.19 `style="flip:x|y|xy"` — and mirrors it about one axis or both,
  // which is how a connector is made to run the other way:
  // groupshape-child-rotation.docx turns its bent connector 180° and flips it
  // back in y, and drawn without the flip it ran from the wrong corner.
  const flip = vmlStyleValue(shape, 'flip')?.toLowerCase() ?? '';
  const flipH = flip.includes('x');
  const flipV = flip.includes('y');
  const spun = rotationDeg !== undefined && rotationDeg % 360 !== 0 ? rotationDeg : undefined;
  const transform =
    spun !== undefined || flipH || flipV
      ? {
          ...(spun !== undefined ? { rotation60k: Math.round(spun * 60000) } : {}),
          ...(flipH ? { flipH: true } : {}),
          ...(flipV ? { flipV: true } : {}),
        }
      : undefined;
  // §14.1.2.22 — a shape with a `v:textpath` is WordArt, wherever it sits: its
  // words are what it draws, not the box they were written in. FDO78590.docx
  // puts two inside its landscape group and we filled two black rectangles
  // over the picture.
  const wordArt = vmlWordArtBody(shape, width, height);
  if (wordArt) {
    return {
      width,
      height,
      geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
      fill: { kind: 'none' },
      text: wordArt,
      ...(transform ? { transform } : {}),
    };
  }
  // §14.1.4 — the shape may state its OUTLINE outright, as a `@path` of plain
  // coordinates in its `@coordsize` space. FDO78590.docx draws a whole
  // landscape as 97 such shapes, and every one of them came out as the black
  // rectangle an unrecognised shape degrades to.
  const custom = vmlCustomGeometry(shape, shapeType);
  return {
    width,
    height,
    geometry: custom ?? {
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
 * §14.1.4 `@path` — a VML shape's own outline, in the coordinate space its
 * `@coordsize` declares (1000×1000 by default) with `@coordorigin` at the
 * corner. Word writes freeform shapes this way, as plain numbers; a path that
 * names a FORMULA (`@n`) or an arc command is left to the preset fallback.
 *
 * @param shape     The `v:shape`/`v:rect`/… node.
 * @param shapeType Its `v:shapetype`, which may carry the path instead.
 * @returns The geometry, or `undefined` when there is no path we can read.
 */
function vmlCustomGeometry(shape: PoNode, shapeType?: PoNode): ShapeGeometry | undefined {
  const raw = poAttr(shape, 'path') ?? (shapeType ? poAttr(shapeType, 'path') : undefined);
  if (raw === undefined || raw === '' || raw.includes('@')) return undefined;
  const size =
    poAttr(shape, 'coordsize') ?? (shapeType ? poAttr(shapeType, 'coordsize') : undefined);
  const origin =
    poAttr(shape, 'coordorigin') ?? (shapeType ? poAttr(shapeType, 'coordorigin') : undefined);
  const [sx, sy] = pairOf(size) ?? [1000, 1000];
  const [ox, oy] = pairOf(origin) ?? [0, 0];
  if (sx <= 0 || sy <= 0) return undefined;
  const commands = vmlPathCommands(raw, ox, oy);
  if (commands === undefined || commands.length === 0) return undefined;
  return { kind: 'custom', custom: { pathWidth: sx, pathHeight: sy, commands } };
}

// "x,y" (or "x y") → the pair, or undefined.
function pairOf(v: string | undefined): [number, number] | undefined {
  const m = v !== undefined ? /^\s*(-?\d+)[ ,]+(-?\d+)\s*$/u.exec(v) : null;
  return m ? [Number(m[1]), Number(m[2])] : undefined;
}

// §14.1.4 — the commands a VML path may carry, of which these are the ones
// with a straight reading: move/line/curve, absolute and relative, and close.
// Anything else (the arcs, `ae`, `qb`) gives up on the whole path rather than
// draw half of it. Coordinates are comma-separated and MAY BE OMITTED, which
// means zero: Word writes `m,l3224,r,587` for "from the corner, right 3224,
// then down 587".
// The two-letter VML path commands (§14.1.4): the fill/stroke modifiers and
// the arcs. `nf`/`ns` draw nothing of their own; the arcs we do not draw.
const VML_PATH_PAIRS = new Set(['nf', 'ns', 'qx', 'qy', 'qb', 'at', 'wa', 'ar', 'ae']);

function vmlPathCommands(
  raw: string,
  originX: number,
  originY: number,
): Array<CustomPathCmd> | undefined {
  const out: Array<CustomPathCmd> = [];
  let x = 0;
  let y = 0;
  let i = 0;
  const isLetter = (c: string): boolean => /[a-z]/iu.test(c);
  while (i < raw.length) {
    const c = raw[i]!;
    if (!isLetter(c)) {
      // Separators between commands are fine; a stray number is not.
      if (/[\s,]/u.test(c)) {
        i++;
        continue;
      }
      return undefined;
    }
    // One letter, or two for the handful of two-letter commands. Only those:
    // a path ends `xe` — close, then end — and read greedily that is one
    // command nobody has heard of, and the whole outline is given up on.
    let cmd = c.toLowerCase();
    i++;
    if (i < raw.length && isLetter(raw[i]!)) {
      const pair = cmd + raw[i]!.toLowerCase();
      if (VML_PATH_PAIRS.has(pair)) {
        cmd = pair;
        i++;
      }
    }
    // Everything up to the next letter is this command's coordinate list.
    const start = i;
    while (i < raw.length && !isLetter(raw[i]!)) i++;
    const args = raw.slice(start, i);
    if (cmd === 'e' || cmd === 'n' || cmd === 'nf' || cmd === 'ns') continue;
    if (cmd === 'x') {
      out.push({ cmd: 'close' });
      continue;
    }
    const nums = vmlPathNumbers(args);
    if (nums === undefined) return undefined;
    const rel = cmd === 't' || cmd === 'r' || cmd === 'v';
    if (cmd === 'm' || cmd === 'l' || cmd === 't' || cmd === 'r') {
      if (nums.length < 2 || nums.length % 2 !== 0) return undefined;
      for (let k = 0; k < nums.length; k += 2) {
        x = rel ? x + nums[k]! : nums[k]!;
        y = rel ? y + nums[k + 1]! : nums[k + 1]!;
        const move = (cmd === 'm' || cmd === 't') && k === 0;
        out.push({ cmd: move ? 'move' : 'line', x: x - originX, y: y - originY });
      }
      continue;
    }
    if (cmd === 'c' || cmd === 'v') {
      if (nums.length < 6 || nums.length % 6 !== 0) return undefined;
      for (let k = 0; k < nums.length; k += 6) {
        const x1 = rel ? x + nums[k]! : nums[k]!;
        const y1 = rel ? y + nums[k + 1]! : nums[k + 1]!;
        const x2 = rel ? x + nums[k + 2]! : nums[k + 2]!;
        const y2 = rel ? y + nums[k + 3]! : nums[k + 3]!;
        x = rel ? x + nums[k + 4]! : nums[k + 4]!;
        y = rel ? y + nums[k + 5]! : nums[k + 5]!;
        out.push({
          cmd: 'cubic',
          x1: x1 - originX,
          y1: y1 - originY,
          x2: x2 - originX,
          y2: y2 - originY,
          x: x - originX,
          y: y - originY,
        });
      }
      continue;
    }
    return undefined; // an arc or something else we do not draw
  }
  return out;
}

// A VML coordinate list: comma-separated, and an omitted field is zero.
function vmlPathNumbers(args: string): Array<number> | undefined {
  const fields = args.trim().split(/\s*,\s*|\s+/u);
  const out: Array<number> = [];
  for (const f of fields) {
    if (f === '') {
      out.push(0);
      continue;
    }
    if (!/^-?\d+$/u.test(f)) return undefined;
    out.push(Number(f));
  }
  return out;
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
  // A NESTED group's `@style` width/height are its parent's coordinate units,
  // not a CSS length: the caller has already scaled them to points.
  sizePt?: { readonly w: number; readonly h: number },
): ShapeData | null {
  const width = sizePt?.w ?? vmlStyleLength(group, 'width');
  const height = sizePt?.h ?? vmlStyleLength(group, 'height');
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
    // §14.1.2.10 — a LINE states its ends instead of a box, and inside a group
    // it usually states nothing else at all: groupshape-line.docx nests one two
    // groups deep with only `from`/`to`, and read for a width it was dropped.
    // …and its ends are plain numbers in that space, not the CSS lengths a
    // top-level line writes: read as points they came out three quarters the
    // size and three quarters of the way along.
    const ends =
      tag === 'v:line'
        ? { from: vmlPair(poAttr(child, 'from')), to: vmlPair(poAttr(child, 'to')) }
        : undefined;
    const span = ends?.from && ends.to ? ends : undefined;
    const left = span ? Math.min(span.from!.x, span.to!.x) : (vmlStyleNumber(child, 'left') ?? 0);
    const top = span ? Math.min(span.from!.y, span.to!.y) : (vmlStyleNumber(child, 'top') ?? 0);
    const w = span ? Math.abs(span.to!.x - span.from!.x) : vmlStyleNumber(child, 'width');
    const h = span ? Math.abs(span.to!.y - span.from!.y) : vmlStyleNumber(child, 'height');
    if (w === undefined || h === undefined) continue;
    const data =
      tag === 'v:group'
        ? parseVmlGroup(child, parseBody, shapeTypes, {
            w: Math.max(1, w * sx),
            h: Math.max(1, h * sy),
          })
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
/** §14.1.2.19 — a CSS-like `@style` property's raw value ("flip:xy" → "xy"). */
function vmlStyleValue(shape: PoNode, prop: string): string | undefined {
  const style = poAttr(shape, 'style');
  if (style === undefined) return undefined;
  const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, 'iu').exec(style);
  return m?.[1]?.trim();
}

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
  '33': 'bentConnector2',
  '34': 'bentConnector3',
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

/**
 * §14.1.2.5 — `@filled="f"` or a `<v:fill type="none">` says the shape is not
 * filled; otherwise `@fillcolor`, defaulting to VML's own white. Exported
 * because §17.2.1's page background is a `v:background` with exactly this fill
 * on it — the same gradients and the same pictures.
 *
 * @param shape       The VML element carrying the fill.
 * @param shapeType   Its `v:shapetype`, for the attributes it does not state.
 * @param resolveImage Resolver for a `v:fill` that paints a PICTURE (`@r:id`).
 * @returns The fill.
 */
export function vmlFill(
  shape: PoNode,
  shapeType?: PoNode,
  resolveImage?: (id: string) => ResourceId | undefined,
): ShapeFill {
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
  const base = colorHex ?? 'FFFFFF';
  // §14.1.2.5 `@type` — the fill may be a GRADIENT between `fillcolor` and the
  // fill's own `@color2`, running top to bottom unless `@angle` says otherwise.
  // fdo76016.docx shades its arrow that way, and we painted it flat.
  // §14.1.2.5 `@type="frame"` — the fill is a PICTURE stretched over the box,
  // named by the relationship on the fill itself. tdf126533_pageBitmap.docx
  // papers its page with one and we painted the flat fallback colour.
  // §14.1.2.5 `@type="tile"` — the same picture REPEATED at its own size, which
  // is how a texture is laid: bib-chernigovka…docx papers its pages with a
  // 128-pixel parchment. `pattern` is NOT either of those — a two-colour tile
  // at its own tiny scale, which stretched over the shape is a black slab
  // (fdo77725.docx's 5 % dotted rectangle came out solid black).
  const fillType = fillEl ? poAttr(fillEl, 'type') : undefined;
  const fillRelId =
    fillType === 'frame' || fillType === 'tile' ? poAttr(fillEl, 'r:id') : undefined;
  const picture = fillRelId !== undefined ? resolveImage?.(fillRelId) : undefined;
  if (picture !== undefined) {
    return {
      kind: 'picture',
      imageResource: picture,
      ...(fillType === 'tile' ? { tiled: true } : {}),
    };
  }
  const gradient = fillEl ? vmlGradient(fillEl, base) : undefined;
  if (gradient) return { kind: 'gradient', gradient };
  // §14.1.2.5 `@opacity` — VML's own transparency, written as a fraction or a
  // percentage.
  const opacity = vmlOpacity(
    poAttr(shape, 'opacity') ?? (fillEl ? poAttr(fillEl, 'opacity') : undefined),
  );
  return { kind: 'solid', colorHex: base, ...(opacity !== undefined ? { alpha: opacity } : {}) };
}

// §14.1.2.5 — a `v:fill` of type `gradient` / `gradientRadial`. VML measures
// `@angle` in degrees from the vertical, so an absent one is the top-to-bottom
// sweep Word writes by default; our own angle counts clockwise from left-to-
// right, which is that same sweep at 90.
function vmlGradient(fillEl: PoNode, base: string): ShapeGradient | undefined {
  const type = poAttr(fillEl, 'type');
  if (type !== 'gradient' && type !== 'gradientRadial') return undefined;
  const color2 = vmlColor(poAttr(fillEl, 'color2'));
  if (color2 === undefined) return undefined;
  const raw = Number.parseFloat(poAttr(fillEl, 'angle') ?? '');
  const angle = (90 + (Number.isFinite(raw) ? raw : 0)) % 360;
  const center =
    type === 'gradientRadial' ? vmlRadialCenter(raw, poAttr(fillEl, 'focus')) : undefined;
  return {
    kind: type === 'gradientRadial' ? 'radial' : 'linear',
    // §14.1.2.5 — VML's radial sweep grows RECTANGLES out to the box, not
    // circles: fill.docx's corner sweep reads as a square-cornered wash in
    // both references and we drew a disc.
    ...(type === 'gradientRadial' ? { sweep: 'rect' as const } : { angle }),
    ...(center ? { center } : {}),
    stops:
      vmlStops(poAttr(fillEl, 'colors')) ?? vmlFocusStops(poAttr(fillEl, 'focus'), base, color2),
  };
}

// §14.1.2.5 — a RADIAL fill at full focus starts in a CORNER, not the middle:
// `@angle` says which one, counting clockwise from twelve o'clock. fill.docx
// sweeps out of its top-left corner (-135°) and, centred, our page ran navy
// through the middle where both references keep it in that corner alone.
function vmlRadialCenter(
  angleRaw: number,
  focus: string | undefined,
): { x: number; y: number } | undefined {
  const pct = Number.parseFloat((focus ?? '').replace('%', ''));
  if (!Number.isFinite(pct) || Math.abs(pct) < 99) return undefined;
  const deg = Number.isFinite(angleRaw) ? angleRaw : 0;
  const rad = (deg * Math.PI) / 180;
  // Clockwise from up, in the box's own y-DOWN fractions: at 0° the sweep
  // starts at the top edge, at 180° at the bottom.
  const dx = Math.sin(rad);
  const dy = Math.cos(rad);
  const snap = (v: number): number => (v > 0.5 ? 1 : v < -0.5 ? 0 : 0.5);
  return { x: snap(dx), y: snap(dy) };
}

// §14.1.2.5 `@focus` — where the SECOND colour sits along the sweep. At 0 the
// gradient runs plainly from one colour to the other; anywhere between, the
// second colour is a band inside and the first stands at BOTH ends, which is
// the axial gradient Word's own dialog calls it. tdf126533_axialAngle.docx
// asks for 50% — fuchsia, lime, fuchsia — and we drew a plain fuchsia→lime
// sweep with no way back.
function vmlFocusStops(
  raw: string | undefined,
  base: string,
  color2: string,
): ShapeGradient['stops'] {
  const pct = Number.parseFloat((raw ?? '').replace('%', ''));
  const f = Number.isFinite(pct) ? Math.abs(pct) / 100 : 0;
  if (f <= 0.01 || f >= 0.99) {
    return [
      { offset: 0, colorHex: base },
      { offset: 1, colorHex: color2 },
    ];
  }
  return [
    { offset: 0, colorHex: base },
    { offset: f, colorHex: color2 },
    { offset: 1, colorHex: base },
  ];
}

// §14.1.2.10 `@gain`/`@blacklevel` — the contrast and brightness a picture is
// drawn through, as `out = (in - 0.5) * gain + 0.5 + black`.
function vmlWash(imagedata: PoNode): { gain: number; black: number } | undefined {
  const gain = vmlFraction(poAttr(imagedata, 'gain'));
  const black = vmlFraction(poAttr(imagedata, 'blacklevel'));
  if (gain === undefined && black === undefined) return undefined;
  const g = Math.max(0, gain ?? 1);
  const b = black ?? 0;
  return Math.abs(g - 1) < 0.01 && Math.abs(b) < 0.01 ? undefined : { gain: g, black: b };
}

// A VML fraction: "0.5", "50%", or the 1/65536ths Word writes ("19661f").
function vmlFraction(raw: string | undefined): number | undefined {
  if (raw === undefined || raw.trim() === '') return undefined;
  const t = raw.trim();
  const pct = t.endsWith('%');
  const n = Number.parseFloat(pct ? t.slice(0, -1) : t.replace(/f$/iu, ''));
  if (!Number.isFinite(n)) return undefined;
  return pct ? n / 100 : Math.abs(n) > 1 ? n / 65536 : n;
}

// §14.1.2.5 `@opacity` — "0.5" or "50%" or the 1/65536ths Word writes.
function vmlOpacity(raw: string | undefined): number | undefined {
  if (raw === undefined || raw.trim() === '') return undefined;
  const pct = raw.trim().endsWith('%');
  const n = Number.parseFloat(pct ? raw.trim().slice(0, -1) : raw.replace(/f$/iu, ''));
  if (!Number.isFinite(n)) return undefined;
  const v = pct ? n / 100 : n > 1 ? n / 65536 : n;
  return v >= 1 ? undefined : Math.max(0, v);
}

// §14.1.2.5 `@colors` — the stops BETWEEN the two the fill names, as
// "position colour" pairs. A position is a fraction, a percentage, or the
// 1/65536ths Word writes ("19661f"). fill.docx runs navy → purple → magenta →
// red → orange in five of them and, read as the two endpoints alone, its page
// lost every colour in the middle.
function vmlStops(raw: string | undefined): ShapeGradient['stops'] | undefined {
  if (raw === undefined || raw.trim() === '') return undefined;
  const stops: Array<{ offset: number; colorHex: string }> = [];
  for (const part of raw.split(';')) {
    const [posRaw, colorRaw] = part.trim().split(/\s+/u);
    const colorHex = vmlColor(colorRaw);
    if (posRaw === undefined || colorHex === undefined) continue;
    const pct = posRaw.endsWith('%');
    const n = Number.parseFloat(pct ? posRaw.slice(0, -1) : posRaw.replace(/f$/iu, ''));
    if (!Number.isFinite(n)) continue;
    const offset = pct ? n / 100 : n > 1 ? n / 65536 : n;
    stops.push({ offset: Math.min(1, Math.max(0, offset)), colorHex });
  }
  if (stops.length < 2) return undefined;
  return stops.sort((a, b) => a.offset - b.offset);
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

// §14.1.2.22 — the WordArt body a shape's `v:textpath` describes: the string
// it carries, set at the size that fills the box (the layout does the fitting,
// which is where the font metrics are). Undefined when the shape has no
// textpath string of its own.
function vmlWordArtBody(shape: PoNode, width: Pt, height: Pt): ShapeTextBody | undefined {
  const textpath = poChildren(shape).find((c) => poIs(c, 'v:textpath'));
  const string = textpath ? poAttr(textpath, 'string') : undefined;
  if (string === undefined || string === '') return undefined;
  const fill = poAttr(shape, 'fillcolor');
  const colorHex =
    fill && /^#?[0-9A-Fa-f]{6}$/u.test(fill) ? fill.replace('#', '').toUpperCase() : undefined;
  const lines = string.split(/\r\n|[\r\n]/u);
  const fontSizePt = pt(Math.max(1, height / lines.length));
  return {
    content: lines.map((line) => ({
      kind: 'paragraph' as const,
      paragraph: {
        // WordArt is glyphs in a box: no paragraph spacing above or below, and
        // single line spacing, or the fit measures the box against a line
        // taller than the letters it holds.
        properties: {
          alignment: 'center' as const,
          spacingBefore: pt(0),
          spacingAfter: pt(0),
          spacingLine: pt(12),
          spacingLineRule: 'auto' as const,
        },
        runs: [
          { text: line, properties: { fontSizePt, bold: true, ...(colorHex ? { colorHex } : {}) } },
        ],
      },
    })),
    anchor: 'ctr' as const,
    fitToBox: true,
    insetLeft: pt(0),
    insetRight: pt(0),
    insetTop: pt(0),
    insetBottom: pt(0),
  };
}

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
// §14.3.2 — the wrap a `w10:wrap` names, or none when the shape states none.
function vmlWrap(shape: PoNode | undefined): FloatAnchor['wrap'] {
  const w = shape ? poChildren(shape).find((c) => poIsLocal(c, 'wrap')) : undefined;
  if (!w) return 'none';
  const type = poAttr(w, 'type');
  if (type === 'topAndBottom') return 'topAndBottom';
  if (type === 'none') return 'none';
  return 'square';
}

function vmlFloat(
  style: string,
  zIndexAttr: number | undefined,
  shape?: PoNode,
): FloatAnchor | undefined {
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
  // …and so may the vertical one: a Word watermark is a header picture centred
  // in the page this way, and pinned to the band's cursor instead it printed in
  // the top corner (pictureWatermark.docx).
  const vAlign = /mso-position-vertical\s*:\s*(center|top|bottom)/iu.exec(style)?.[1] as
    | 'center'
    | 'top'
    | 'bottom'
    | undefined;
  const x = prop('margin-left') ?? prop('left') ?? 0;
  const y = prop('margin-top') ?? prop('top') ?? 0;
  // §14.1.2.19 — the z-order lives in the shape's STYLE, not in an attribute of
  // its own; read as one it was always absent, so no VML shape was ever behind
  // the text.
  const zIndex = prop('z-index') ?? zIndexAttr;
  return {
    // §14.3.2 `w10:wrap` — a positioned VML shape wraps text only where it SAYS
    // so; with no `w10:wrap` it sits over or behind the text and the flow does
    // not see it. Wrapped square regardless, tdf108973_backgroundTextbox.docx's
    // box squeezed the sentence it is meant to sit behind into a column two
    // words wide, over two pages.
    wrap: vmlWrap(shape),
    // §14.1.2.2 `o:allowincell` — VML's spelling of `layoutInCell`: off, the
    // shape is placed against the page it names rather than the cell it sits
    // in (tdf129888vml.docx rules the page edge from inside a cell).
    ...(poAttrLocal(shape, 'allowincell') === 'f' ? { inCell: false } : {}),
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
      ...(vAlign ? { align: vAlign } : {}),
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
  // §14.1.2.22 `@fitshape` — WordArt is set at whatever size FILLS its box.
  // The parser has no font metrics, so it states a size in the right order of
  // magnitude and marks the body `fitToBox`; the layout, which has the metrics,
  // scales it to the box it measured. Left to the parser's own guess,
  // fdo78300.docx's title came out at half the size both references draw.
  const lines = string.split(/\r\n|[\r\n]/u);
  const fontSizePt = pt(Math.max(1, height / lines.length));
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
            // WordArt is glyphs in a box: no paragraph spacing above or below,
            // and single line spacing, or the fit measures the box against a
            // line taller than the letters it holds.
            properties: {
              alignment: 'center' as const,
              spacingBefore: pt(0),
              spacingAfter: pt(0),
              spacingLine: pt(12),
              spacingLineRule: 'auto' as const,
            },
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
        fitToBox: true,
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
  themeStyles?: ThemeStyles,
  parseDrawingText?: ParseDrawingText,
): ShapeData | null {
  const canvas = expandMcChildren(poChildren(graphicData)).find((c) =>
    poIsLocal(c, 'lockedCanvas'),
  );
  if (!canvas) return null;
  const raw = groupChildren(
    canvas,
    resolveColor,
    parseBody,
    resolveImage,
    themeLineWidths,
    themeStyles,
    parseDrawingText,
  );
  if (extentCx === undefined || extentCy === undefined) return null;
  // §20.1.7.5 — when the canvas states the space its members are written in
  // (`a:chOff`/`a:chExt`), that space maps onto the frame the document gives
  // it. fdo76249.docx writes its logo in a box 5.8× the frame, and drawn at
  // face value the statue alone filled half the page.
  const space = canvasChildSpace(canvas);
  let children: Array<ShapeGroupChild>;
  if (space) {
    const sx = emuToPt(extentCx) / space.w;
    const sy = emuToPt(extentCy) / space.h;
    children = raw.map((c) => ({
      shape: { ...c.shape, width: pt(c.shape.width * sx), height: pt(c.shape.height * sy) },
      xPt: pt((c.xPt - space.x) * sx),
      yPt: pt((c.yPt - space.y) * sy),
    }));
  } else {
    // The members keep the offsets of the document they were copied from, and
    // a canvas that states no child space is drawn from its OWN corner: what
    // Word and LibreOffice both do with fdo43641.docx, whose rectangle would
    // otherwise sit 22pt above the frame it belongs in.
    const origin = contentOrigin(raw);
    children = raw.map((c) => ({ ...c, xPt: pt(c.xPt - origin.x), yPt: pt(c.yPt - origin.y) }));
  }
  return {
    width: emuToPt(extentCx),
    height: emuToPt(extentCy),
    ...(children.length > 0 ? { children } : {}),
    geometry: { kind: 'custom', custom: EMPTY_GEOMETRY },
    fill: { kind: 'none' },
  };
}

// `a:grpSpPr/a:xfrm` on a locked canvas: the box its members are written in,
// in POINTS. Undefined when the canvas states none — then the members keep the
// size and the offsets they were copied with.
function canvasChildSpace(
  canvas: PoNode,
): { x: number; y: number; w: number; h: number } | undefined {
  const grpSpPr = poChildren(canvas).find((c) => poIsLocal(c, 'grpSpPr'));
  const xfrm = grpSpPr ? poChildren(grpSpPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  const chOff = xfrm ? poChildren(xfrm).find((c) => poIs(c, 'a:chOff')) : undefined;
  const chExt = xfrm ? poChildren(xfrm).find((c) => poIs(c, 'a:chExt')) : undefined;
  const cx = chExt ? poIntAttr(chExt, 'cx') : undefined;
  const cy = chExt ? poIntAttr(chExt, 'cy') : undefined;
  if (cx === undefined || cy === undefined || cx <= 0 || cy <= 0) return undefined;
  return {
    x: emuToPt((chOff ? poIntAttr(chOff, 'x') : undefined) ?? 0),
    y: emuToPt((chOff ? poIntAttr(chOff, 'y') : undefined) ?? 0),
    w: emuToPt(cx),
    h: emuToPt(cy),
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
  themeStyles?: ThemeStyles,
  parseDrawingText?: ParseDrawingText,
): ShapeData | null {
  const wgp = expandMcChildren(poChildren(graphicData)).find((c) => poIs(c, 'wpg:wgp'));
  if (!wgp) return null;
  const children = groupChildren(
    wgp,
    resolveColor,
    parseBody,
    resolveImage,
    themeLineWidths,
    themeStyles,
    parseDrawingText,
  );
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
//
// §20.1.7.5 — `unitScale` is EMU per unit of the space THIS group's `a:ext` is
// written in: 1 for a top-level group, whose `a:ext` is already EMU, and the
// parent's own scale for a nested one, whose `a:ext` is in the parent's CHILD
// units. Left at 1 all the way down, relorientation.docx's nested group scaled
// its two rectangles to a third of a point and the page came out blank.
function groupChildren(
  wgp: PoNode,
  resolveColor: ColorResolver,
  parseBody: ParseBody | undefined,
  resolveImage?: (relId: string) => ResourceId | undefined,
  themeLineWidths?: ReadonlyArray<number>,
  themeStyles?: ThemeStyles,
  parseDrawingText?: ParseDrawingText,
  unitScale: { readonly x: number; readonly y: number } = { x: 1, y: 1 },
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
  const sx = (ext && chExt && chExt.x > 0 ? ext.x / chExt.x : 1) * unitScale.x;
  const sy = (ext && chExt && chExt.y > 0 ? ext.y / chExt.y : 1) * unitScale.y;
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
      ? parseNestedGroup(
          child,
          cx * sx,
          cy * sy,
          resolveColor,
          parseBody,
          resolveImage,
          themeLineWidths,
          themeStyles,
          parseDrawingText,
          { x: sx, y: sy },
        )
      : picture
        ? parseGroupPicture(child, cx, cy, resolveImage)
        : parseWspNode(
            child,
            cx,
            cy,
            resolveColor,
            parseBody,
            resolveImage,
            themeLineWidths,
            themeStyles,
            parseDrawingText,
          );
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
  themeStyles?: ThemeStyles,
  parseDrawingText?: ParseDrawingText,
  unitScale?: { readonly x: number; readonly y: number },
): ShapeData {
  const children = groupChildren(
    grpSp,
    resolveColor,
    parseBody,
    resolveImage,
    themeLineWidths,
    themeStyles,
    parseDrawingText,
    unitScale,
  );
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
  themeStyles?: ThemeStyles,
  parseDrawingText?: ParseDrawingText,
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
    themeStyles,
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
  themeStyles?: ThemeStyles,
  parseDrawingText?: ParseDrawingText,
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

  // §21.1.2 — a member of a group or a locked canvas states its words in an
  // `a:txBody`, not a `wps:txbx`. fdo78658.docx labels every box of its diagram
  // that way, and we drew the boxes empty.
  // …which a shape may wrap in an `a:txSp` (§20.1.2.2.41), as Word does for
  // the WordArt it writes in DrawingML.
  const txSp = poChildren(wsp).find((c) => poIs(c, 'a:txSp'));
  const txBody =
    poChildren(wsp).find((c) => poIs(c, 'a:txBody')) ??
    (txSp ? poChildren(txSp).find((c) => poIs(c, 'a:txBody')) : undefined);
  const text =
    (parseBody ? parseTextBox(wsp, parseBody) : undefined) ??
    (txBody && parseDrawingText ? parseDrawingText(txBody, resolveColor) : undefined);
  // §20.1.4.2.19/20.1.4.2.10 — a shape drawn from a gallery style keeps its
  // fill and outline in `<wps:style>` and carries none in `spPr` at all; read
  // alone, spPr says the shape has neither. TextEffects_Groupshapes.docx's
  // rectangle asks for accent1 that way and we drew its caption on white.
  const style = poChildren(wsp).find((c) => poIsLocal(c, 'style'));
  // …but a fill the shape states ITSELF wins, and `a:noFill` is a statement:
  // §20.1.4.1.9's own properties override the reference. ShapeOverlappingWithSdt
  // spells `a:noFill` beside a `fillRef` and we painted its rectangle accent
  // blue, over the heading it is drawn around.
  if (style && fill.kind === 'none' && !statesFill(spPr)) {
    fill = styleRefFill(style, resolveColor, themeStyles);
  }
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
  // §20.1.8.40 — the shadow the shape casts: its own `a:effectLst`, or the one
  // the gallery style names through `a:effectRef` (§20.1.4.2.8). Word draws it
  // under every shape that asks; we read neither, so imgshadow.docx and
  // fdo78957.docx's boxes stood on the page with nothing under them.
  const shadow =
    (spPr ? parseShadow(spPr, resolveColor) : undefined) ??
    (style ? styleRefShadow(style, resolveColor, themeStyles?.effects) : undefined);

  if (widthEmu === undefined || heightEmu === undefined) return null;
  return {
    width: emuToPt(widthEmu),
    height: emuToPt(heightEmu),
    geometry,
    fill,
    ...(line ? { line } : {}),
    ...(shadow ? { shadow } : {}),
    ...(transform ? { transform } : {}),
    ...(styled ? { text: styled } : {}),
  };
}

/**
 * §20.1.4.2.8 `<a:effectRef>` — the shadow a gallery style names: a 1-based
 * index into the theme's `a:effectStyleLst`, plus the colour to put wherever
 * that style says `phClr`. The same mechanism as the fill and the outline
 * beside it.
 *
 * @param style             The shape's `<wps:style>` node.
 * @param resolveColor      The document's colour resolver.
 * @param themeEffectStyles The theme's effect styles, in list order.
 * @returns The shadow, or undefined when the style names none.
 */
function styleRefShadow(
  style: PoNode,
  resolveColor: ColorResolver,
  themeEffectStyles?: ReadonlyArray<PoNode>,
): ShapeShadow | undefined {
  const ref = poChildren(style).find((c) => poIs(c, 'a:effectRef'));
  const idx = ref ? Number(poAttr(ref, 'idx') ?? '') : NaN;
  if (!ref || !Number.isFinite(idx) || idx < 1) return undefined;
  const slot = themeEffectStyles?.[idx - 1];
  const list = slot ? poChildren(slot).find((c) => poIs(c, 'a:effectLst')) : undefined;
  const shdw = list ? poChildren(list).find((c) => poIs(c, 'a:outerShdw')) : undefined;
  if (!shdw) return undefined;
  const child = firstElementChild(ref);
  const phHex = child ? resolveColorNode(child, resolveColor) : undefined;
  return shadowFromOuterShdw(shdw, phHex ? placeholderColors(resolveColor, phHex) : resolveColor);
}

// wps:txbx/w:txbxContent (the text body) + wps:bodyPr (insets + vertical
// anchor). Returns undefined when the shape carries no text.
// Whether `spPr` states a fill of its own at all — including `a:noFill`, which
// is a shape saying it has none rather than saying nothing.
const FILL_TAGS: ReadonlySet<string> = new Set([
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
]);

// §20.1.8.33 — a `a:gradFill` with no `a:gsLst` states a DIRECTION and no
// colours: the run of stops comes from the gallery style's `a:fillRef`, and the
// shape only says which way to sweep it. Counted as a fill of its own, the
// shape ends up with none at all — 63200.pptx's ellipse drew as its own shadow,
// a grey disc where every reader has the theme's blue.
function emptyGradient(node: PoNode): boolean {
  return poIs(node, 'a:gradFill') && !poChildren(node).some((c) => poIs(c, 'a:gsLst'));
}

/**
 * The fill with the direction the SHAPE states, where it states one and no
 * colours: §20.1.8.33's `a:gradFill` holding only an `a:lin` says "sweep the
 * gallery style's run of colours this way". 63200.pptx's ellipse asks for the
 * diagonal where the theme's slot sweeps straight down.
 *
 * @param fill The fill resolved from the style (or anywhere else).
 * @param spPr The shape's own properties.
 * @returns The fill, turned; unchanged when the shape states no bare direction.
 */
export function withStatedDirection(fill: ShapeFill, spPr: PoNode | undefined): ShapeFill {
  const gradient = fill.kind === 'gradient' ? fill.gradient : undefined;
  if (gradient === undefined || gradient.kind !== 'linear' || spPr === undefined) return fill;
  const grad = poChildren(spPr).find((c) => emptyGradient(c));
  const lin = grad ? poChildren(grad).find((c) => poIs(c, 'a:lin')) : undefined;
  const ang = lin ? poIntAttr(lin, 'ang') : undefined;
  if (ang === undefined) return fill;
  return { ...fill, gradient: { ...gradient, angle: (ang / 60000) % 360 } };
}

/**
 * Whether an `spPr` states a fill AT ALL — including `a:noFill`, which is a
 * shape saying it has none rather than saying nothing. A placeholder that says
 * nothing inherits its prototype's.
 *
 * @param spPr The shape properties, or undefined.
 * @returns Whether a fill element is present.
 */
export function statesFill(spPr: PoNode | undefined): boolean {
  return (
    spPr !== undefined &&
    poChildren(spPr).some((c) => FILL_TAGS.has(poTag(c) ?? '') && !emptyGradient(c))
  );
}

// §20.1.4.2.13 `<a:fillRef>` — the fill a gallery style names. The theme's own
// `a:fillStyleLst` slot (which could make it a gradient) is out of reach here;
// the colour the reference names is what both references draw.
export function styleRefFill(
  style: PoNode,
  resolveColor: ColorResolver,
  themeStyles?: ThemeStyles,
): ShapeFill {
  const ref = poChildren(style).find((c) => poIs(c, 'a:fillRef'));
  if (!ref || poAttr(ref, 'idx') === '0') return { kind: 'none' };
  const child = firstElementChild(ref);
  const colorHex = child ? resolveColorNode(child, resolveColor) : undefined;
  // §20.1.4.1.14 — `@idx` names a slot of the theme's format scheme, and what
  // stands there is a whole FILL: the standard theme's slots are a solid, a
  // subtle gradient and a stronger one, and past 1000 the BACKGROUND fills.
  // Read as the reference's colour alone, fdo78957.docx's page-sized backdrop
  // — background slot 2, a white-to-grey sweep — came out flat white on white
  // paper, which is to say invisible.
  const idx = Number(poAttr(ref, 'idx') ?? '');
  const slot =
    Number.isFinite(idx) && idx >= 1001
      ? themeStyles?.bgFills?.[idx - 1001]
      : Number.isFinite(idx) && idx >= 1
        ? themeStyles?.fills?.[idx - 1]
        : undefined;
  if (slot) {
    const resolver = colorHex ? placeholderColors(resolveColor, colorHex) : resolveColor;
    const fill = fillFromNode(slot, resolver);
    if (fill) return fill;
  }
  return colorHex === undefined ? { kind: 'none' } : { kind: 'solid', colorHex };
}

// §20.1.4.2.19 `<a:lnRef>` — the outline a gallery style names. Its width lives
// in the theme's `a:lnStyleLst`, which is not reachable from here; the hairline
// below is what a shape with no stated width already draws.
export function styleRefLine(
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

/**
 * §20.1.4.2.14 `a:fontRef` — the colour a gallery style writes its text in.
 *
 * @param style        The shape's `p:style`/`wps:style` node.
 * @param resolveColor The document's colour resolver.
 * @returns The 6-hex colour, or `undefined` when the style names none.
 */
export function styleRefFontColor(style: PoNode, resolveColor: ColorResolver): string | undefined {
  const ref = poChildren(style).find((c) => poIs(c, 'a:fontRef'));
  const child = firstElementChild(ref);
  return child ? resolveColorNode(child, resolveColor) : undefined;
}

// §20.1.4.2.14 — give every run that names no colour of its own the one the
// gallery style's `a:fontRef` names. The theme's colour is the FLOOR of the
// cascade (§17.7.2), so a run that could inherit one — through its own
// character style or its paragraph's — keeps what the style sheet gives it:
// ColorOverwritten.docx writes its arrow's two lines in a "red" and a "green"
// paragraph style, and stamping the theme's white over them left the shape
// blank.
export function withStyleFontColor(
  text: ShapeTextBody,
  style: PoNode,
  resolveColor: ColorResolver,
): ShapeTextBody {
  const colorHex = styleRefFontColor(style, resolveColor);
  if (colorHex === undefined) return text;
  return {
    ...text,
    content: text.content.map((block) =>
      block.kind === 'paragraph'
        ? {
            ...block,
            paragraph: {
              ...block.paragraph,
              // Carried at the rank the theme's colour holds — under every
              // style the text names — instead of stamped on the runs that
              // happen to state nothing. Stamped, a run whose PARAGRAPH is
              // styled had to be skipped whether or not that style named a
              // colour, and tdf113258.docx's ellipse, whose Heading 1 names
              // none, drew black on blue where both references draw white.
              properties: {
                ...block.paragraph.properties,
                inheritedRun: { ...block.paragraph.properties.inheritedRun, colorHex },
              },
            },
          }
        : block,
    ),
  };
}

function parseTextBox(wsp: PoNode, parseBody: ParseBody): ShapeTextBody | undefined {
  const txbx = poChildren(wsp).find((c) => poIs(c, 'wps:txbx'));
  // A box that CONTINUES another carries `wps:linkedTxbx` in place of the text:
  // the words live in the chain's first box, and what will not fit there is
  // drawn here (LinkedTextBoxes.docx sets a newsletter in two such columns).
  const linked = poChildren(wsp).find((c) => poIs(c, 'wps:linkedTxbx'));
  const chainId = poAttr(txbx, 'id') ?? poAttr(linked, 'id');
  const chain =
    chainId === undefined
      ? undefined
      : { id: chainId, seq: linked ? (poIntAttr(linked, 'seq') ?? 1) : 0 };
  if (!txbx && !linked) return undefined;
  const txContent = txbx ? poChildren(txbx).find((c) => poIs(c, 'w:txbxContent')) : undefined;
  if (txbx && !txContent) return undefined;
  const content = txContent ? parseBody(poChildren(txContent)) : [];
  if (content.length === 0 && !linked) return undefined;

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
    ...(chain ? { chain } : {}),
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
    const own = fillFromNode(child, resolveColor, resolveImage);
    if (own) return own;
  }
  return { kind: 'none' };
}

/**
 * One fill ELEMENT — `a:noFill` / `a:solidFill` / `a:gradFill` / `a:blipFill` —
 * as a {@link ShapeFill}. The same reader serves a shape's own `a:spPr` and the
 * theme slot a `<a:fillRef>` names, which is a fill in its own right.
 *
 * @param child        The candidate node (anything else returns undefined).
 * @param resolveColor The colour resolver (with `phClr` bound, for a theme slot).
 * @param resolveImage Resolver for a picture fill's relationship.
 * @returns The fill, or undefined when the node is not one.
 */
function fillFromNode(
  child: PoNode,
  resolveColor: ColorResolver,
  resolveImage?: (relId: string) => ResourceId | undefined,
): ShapeFill | undefined {
  {
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
      const rect = fillRectBox(child);
      const duotone = blip ? duotoneOf(blip, resolveColor) : undefined;
      // §20.1.8.58 `a:tile` — the picture REPEATS at its own size instead of
      // being stretched over the box (§20.1.8.56 `a:stretch`).
      // NoFillAttrInImagedata.docx papers two text boxes with a texture that
      // way, and stretched it came out a brown blur.
      const tile = poChildren(child).find((c) => poIs(c, 'a:tile'));
      // §20.1.8.14 — the fill MODE is optional, and a `blipFill` that names
      // neither of them tiles rather than stretches. tdf128596 is a rounded
      // rectangle papered with a 32×32 tick and nothing else in its fill; read
      // as a stretch it came out as one tick blown up over the whole shape,
      // where both references paper it.
      const stretched = poChildren(child).some((c) => poIs(c, 'a:stretch'));
      const tiled = tile !== undefined || !stretched;
      // §20.1.8.58 `@sx` / `@sy` — the scale applied BEFORE the repeat, in
      // thousandths of a percent. Read as a plain repeat, a texture the file
      // halves tiles once where it should tile four times.
      const tileScale = tile ? tileScaleOf(tile) : undefined;
      // §20.1.8.4 `a:alphaModFix` — how opaque the PICTURE is drawn, stated on
      // the blip rather than on a colour. A slide backed by a photo at 70 %
      // showed it at full strength, which on tdf146223 is a saturated red
      // quadrant where the reference has a pale one.
      const fixed = blip ? poChildren(blip).find((c) => poIs(c, 'a:alphaModFix')) : undefined;
      const amt = fixed ? poIntAttr(fixed, 'amt') : undefined;
      const alpha = amt === undefined ? undefined : Math.min(1, Math.max(0, amt / 100000));
      return {
        kind: 'picture',
        imageResource: resource,
        ...(tiled ? { tiled: true } : {}),
        ...(tileScale ? { tileScale } : {}),
        ...(crop ? { imageCrop: crop } : {}),
        ...(rect ? { imageFillRect: rect } : {}),
        ...(duotone ? { duotone } : {}),
        ...(alpha !== undefined && alpha < 1 ? { alpha } : {}),
      };
    }
    if (poIs(child, 'a:solidFill')) {
      const hex = colorFromContainer(child, resolveColor);
      // §20.1.2.3.1 — the transparency is the fill's, not the colour's: read
      // off here it is drawn as transparency, and what is behind the shape
      // shows through. Left to the colour resolver it composites over white,
      // which is right only on white paper.
      const alpha = containerAlpha(child);
      return hex
        ? {
            kind: 'solid',
            // The resolver has already composited the colour over white for the
            // writers that cannot draw transparency; drawing THAT at the same
            // transparency would fade it twice. The compositing is linear, so
            // it undoes exactly.
            colorHex: alpha === undefined ? hex : unblendWhite(hex, alpha),
            ...(alpha !== undefined ? { alpha } : {}),
          }
        : { kind: 'none' };
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
      if (!gradient) return { kind: 'none' };
      const alpha = gradientAlpha(gradient.stops);
      return { kind: 'gradient', gradient, ...(alpha !== undefined ? { alpha } : {}) };
    }
  }
  return undefined;
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
 * §20.1.8.16 `a:clrChange` — the colour a picture declares away, and what it
 * becomes. `useA` (default true) says the destination's alpha counts, so a
 * destination at zero alpha knocks the colour OUT rather than replacing it.
 *
 * @param blip         The `a:blip` node.
 * @param resolveColor The colour resolver.
 * @returns The change, or `undefined` when the blip declares none.
 */
export function colorChangeOf(
  blip: PoNode,
  resolveColor: ColorResolver,
): { readonly fromHex: string; readonly toHex: string; readonly transparent: boolean } | undefined {
  const change = poChildren(blip).find((c) => poIs(c, 'a:clrChange'));
  if (!change) return undefined;
  const side = (tag: string): PoNode | undefined => poChildren(change).find((c) => poIs(c, tag));
  const from = side('a:clrFrom');
  const to = side('a:clrTo');
  const inner = (holder: PoNode | undefined): PoNode | undefined =>
    holder ? poChildren(holder).find((c) => poTag(c) !== undefined) : undefined;
  const fromNode = inner(from);
  const toNode = inner(to);
  const fromHex = fromNode ? resolveColorNode(fromNode, resolveColor) : undefined;
  if (fromHex === undefined) return undefined;
  const toHex = (toNode ? resolveColorNode(toNode, resolveColor) : undefined) ?? fromHex;
  const useA = poAttr(change, 'useA') !== '0';
  const alpha = toNode ? nodeAlpha(toNode) : undefined;
  return { fromHex, toHex, transparent: useA && alpha !== undefined && alpha <= 0.001 };
}

/** §20.1.2.3.1 — a colour node's own `a:alpha`, as a fraction. */
function nodeAlpha(color: PoNode): number | undefined {
  const alpha = poChildren(color).find((c) => poIs(c, 'a:alpha'));
  const val = alpha ? poIntAttr(alpha, 'val') : undefined;
  return val === undefined ? undefined : val / 100000;
}

/**
 * §20.1.8.23 `a:duotone` — the two colours a picture is recoloured between,
 * dark end first. Both are ordinary colour containers, so a theme's `phClr`
 * resolves through whatever resolver the caller bound.
 *
 * @param blip         The `a:blip` node.
 * @param resolveColor The colour resolver.
 * @returns The pair, or `undefined` when the blip states no duotone.
 */
function duotoneOf(
  blip: PoNode,
  resolveColor: ColorResolver,
): { readonly shadowHex: string; readonly highlightHex: string } | undefined {
  // §20.1.8.34 `a:grayscl` — the picture drawn in shades of grey. That is what
  // a duotone from black to white already is, so it is read as one rather than
  // grown a channel of its own: tdf112209's photographed chevron is a colour
  // photograph in the file and grey in every reader.
  if (poChildren(blip).some((c) => poIs(c, 'a:grayscl'))) {
    return { shadowHex: '000000', highlightHex: 'FFFFFF' };
  }
  const duotone = poChildren(blip).find((c) => poIs(c, 'a:duotone'));
  if (!duotone) return undefined;
  const colors = poChildren(duotone)
    .map((c) => resolveColorNode(c, resolveColor))
    .filter((hex): hex is string => hex !== undefined);
  const [shadowHex, highlightHex] = colors;
  return shadowHex !== undefined && highlightHex !== undefined
    ? { shadowHex, highlightHex }
    : undefined;
}

/**
 * §20.1.8.30 `a:stretch/a:fillRect` with POSITIVE insets — the part of the box
 * the picture is stretched INTO. The negative case is the zoom {@link
 * fillRectCrop} reads; this is the other one, and unread it drew a background
 * picture inset into the corner of a slide across the whole of it
 * (tdf153466.pptx: a triangle five times its size).
 *
 * @param blipFill The `a:blipFill` node.
 * @returns The rect as fractions of the box, or `undefined` when every inset
 *          is zero or negative.
 */
// §20.1.8.58 `a:tile @sx @sy` — the scale a tile is drawn at, in thousandths of
// a percent, before it is repeated. Absent or 100 % means the picture's own size.
function tileScaleOf(tile: PoNode): ShapeFill['tileScale'] {
  const pct = (name: string): number => {
    const v = poIntAttr(tile, name);
    return v === undefined || v <= 0 ? 1 : v / 100000;
  };
  const sx = pct('sx');
  const sy = pct('sy');
  return sx === 1 && sy === 1 ? undefined : { sx, sy };
}

function fillRectBox(blipFill: PoNode): ShapeFill['imageFillRect'] {
  const stretch = poChildren(blipFill).find((c) => poIs(c, 'a:stretch'));
  const rect = stretch ? poChildren(stretch).find((c) => poIs(c, 'a:fillRect')) : undefined;
  if (!rect) return undefined;
  const side = (name: string): number => Math.max(0, (poIntAttr(rect, name) ?? 0) / 100000);
  const [left, top, right, bottom] = [side('l'), side('t'), side('r'), side('b')];
  if (left + top + right + bottom === 0) return undefined;
  if (left + right >= 1 || top + bottom >= 1) return undefined;
  return { left, top, right, bottom };
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
  const alpha = alphaMod ? alphaMod.val : 1;
  const dist = emuToPt(poIntAttr(shdw, 'dist') ?? 0);
  // §20.1.10.13 ST_PositiveFixedAngle — 60 000ths of a degree.
  const dirDeg = (poIntAttr(shdw, 'dir') ?? 0) / 60000;
  const rad = (dirDeg * Math.PI) / 180;
  return {
    dxPt: dist * Math.cos(rad),
    dyPt: dist * Math.sin(rad),
    blurPt: emuToPt(poIntAttr(shdw, 'blurRad') ?? 0),
    // §20.1.2.3.1 — the resolver has already composited the transparency over
    // the paper, and the shadow carries it a second time when it paints: black
    // at 40 % came out as 40 % of a grey that was already 60 % white, a fringe
    // barely darker than the page. The colour the shadow wants is the one
    // BEFORE that wash (tdf104015's master casts exactly this).
    colorHex: alpha < 1 ? unblendWhite(colorHex, alpha) : colorHex,
    alpha,
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

/**
 * A picture's frame from the line its `pic:spPr` (or its VML shape) states —
 * colour and width only, since a picture border is a rule around the box and
 * carries none of a shape outline's ends or joins.
 *
 * @param line The parsed line, or `undefined`.
 * @returns The outline, or `undefined` when the picture is drawn bare.
 */
function pictureOutline(line: ShapeLine | undefined): PictureOutline | undefined {
  const stroke = buildStroke(line);
  if (!stroke || stroke.widthPt <= 0) return undefined;
  return { colorHex: stroke.colorHex, widthPt: pt(stroke.widthPt) };
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
// The colour that, composited over white at `alpha`, gives `hex` back.
function unblendWhite(hex: string, alpha: number): string {
  if (alpha <= 0.004) return hex;
  const n = parseInt(hex, 16);
  if (Number.isNaN(n)) return hex;
  const ch = (shift: number): string => {
    const v = ((n >> shift) & 255) / 255;
    const raw = (v - (1 - alpha)) / alpha;
    return Math.round(Math.max(0, Math.min(1, raw)) * 255)
      .toString(16)
      .padStart(2, '0');
  };
  return (ch(16) + ch(8) + ch(0)).toUpperCase();
}

// §20.1.2.3.1 — the `a:alpha` on the colour a fill container wraps, as 0..1.
function containerAlpha(parent: PoNode): number | undefined {
  for (const c of poChildren(parent)) {
    const mod = readColorMods(c).find((m) => m.kind === 'alpha');
    if (mod) return Math.max(0, Math.min(1, mod.val));
  }
  return undefined;
}

function colorFromContainer(parent: PoNode, resolveColor: ColorResolver): string | undefined {
  for (const c of poChildren(parent)) {
    const hex = resolveColorNode(c, resolveColor);
    if (hex) return hex;
  }
  return undefined;
}

/**
 * §20.1.2.3.1 — a gradient's stop transparencies, settled.
 *
 * The resolver composites a translucent colour over the PAPER, because a solid
 * fill has nowhere else to put the transparency. A gradient does: its stops
 * become a luminosity soft mask of the same sweep (§11.6.5.2,
 * `buildGradientAlphaMask`), so each stop keeps its own alpha and its own
 * colour, and the mask decides what shows through.
 *
 * That means the washing has to be undone here, on every stop — including the
 * opaque ones, where undoing it is a no-op. It used to be undone only when
 * EVERY stop was translucent, on the reasoning that a page paints a gradient at
 * one transparency; the mask makes that false, and the cost of the old reading
 * was a band that should have been half-transparent over a blue slide drawn as
 * an opaque pale stripe (45541's layout fades the middle of its left rule to
 * 50 %).
 */
function normalizeStopAlpha(stops: Array<GradientStop>): void {
  for (const [i, s] of stops.entries()) {
    stops[i] = { ...s, colorHex: unblendWhite(s.colorHex, s.alpha ?? 1) };
  }
}

/** The one transparency a gradient can be drawn at; see {@link normalizeStopAlpha}. */
function gradientAlpha(stops: ReadonlyArray<GradientStop>): number | undefined {
  const strongest = Math.max(...stops.map((s) => s.alpha ?? 1));
  return strongest < 1 ? strongest : undefined;
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
    const alpha = containerAlpha(gs);
    stops.push({ offset, colorHex: hex, ...(alpha !== undefined ? { alpha } : {}) });
  }
  if (stops.length === 0) return undefined;
  stops.sort((a, b) => a.offset - b.offset);
  normalizeStopAlpha(stops);
  // §20.1.8.46 `a:path` — a radial sweep, and it says both WHAT SHAPE its
  // contours are (`circle` / `rect` / `shape`) and WHERE it starts: the
  // `a:fillToRect` is the rectangle the FIRST stop fills, in the box's own
  // hundred-thousandths. The standard theme's background sweep starts a fifth
  // in from the left and halfway down, and centred it lit the wrong quarter of
  // fdo78957.docx's cover.
  const path = poChildren(grad).find((c) => poIs(c, 'a:path'));
  if (path) {
    const rect = poChildren(path).find((c) => poIs(c, 'a:fillToRect'));
    const side = (name: string): number | undefined => {
      const v = rect ? poIntAttr(rect, name) : undefined;
      return v === undefined ? undefined : v / 100000;
    };
    const l = side('l');
    const r = side('r');
    const t = side('t');
    const b = side('b');
    const center =
      l !== undefined || r !== undefined || t !== undefined || b !== undefined
        ? {
            x: clampUnit(((l ?? 0) + 1 - (r ?? 0)) / 2),
            y: clampUnit(((t ?? 0) + 1 - (b ?? 0)) / 2),
          }
        : undefined;
    return {
      kind: 'radial',
      stops,
      // §20.1.8.46 `@path` — `circle` sweeps in circles, `rect` in rectangles,
      // and `shape` follows the SHAPE's own outline, which for the rectangle a
      // background or a plain box is means rectangles again. Swept as a circle
      // instead, the contours run out to the corners at a different rate than
      // to the sides: tdf114848's centred glow came out as a band across the
      // middle of the slide.
      ...(poAttr(path, 'path') === 'rect' || poAttr(path, 'path') === 'shape'
        ? { sweep: 'rect' as const }
        : {}),
      ...(center ? { center } : {}),
    };
  }
  const lin = poChildren(grad).find((c) => poIs(c, 'a:lin'));
  const ang = lin ? poIntAttr(lin, 'ang') : undefined;
  const angle = ang !== undefined ? (ang / 60000) % 360 : 0;
  return { kind: 'linear', angle, stops };
}

function clampUnit(x: number): number {
  return x < 0 ? 0 : x > 1 ? 1 : x;
}
