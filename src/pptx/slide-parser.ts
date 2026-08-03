// E-PPTX PX1–PX5 — slide shapes → positioned FlowDoc elements.
//
// A PresentationML slide (p:cSld/p:spTree) is a canvas of shapes. Under Route A
// (epics.md) each shape becomes a floating element anchored at its EMU position,
// reusing the docx drawing model and DrawingML readers.
//   * PX1/PX2 — p:sp text boxes: own a:xfrm + direct a:rPr, else the placeholder
//     cascade (geometry + master p:txStyles) for a p:ph with no own transform.
//   * PX3 — p:pic floating images; p:sp visible geometry/fill/stroke (p:spPr).
//   * PX4 — p:graphicFrame: a c:chart floating ChartBlock or an a:tbl Table.
//   * PX5 — colours via the deck theme resolver; p:grpSp groups (a child→slide
//     transform). Backgrounds + theme wiring live in pptx-reader.
// Bullets/levels/alignment/anchor/autofit + hyperlinks come in PX6.

import type {
  Alignment,
  BodyElement,
  Border,
  CellBorders,
  CellMerge,
  ChartBlock,
  FloatAnchor,
  ImageBlock,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  ShapeBlock,
  ShapeFill,
  ShapeGeometry,
  ShapeTextBody,
  Table,
  TableCell,
  TableRow,
} from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { Loss, Pt, ResourceId } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import type { LevelBullet, PlaceholderCascade } from '@/pptx/placeholder-cascade';
import type { PlaceholderRef, ShapeBoxEmu } from '@/pptx/sp-helpers';

import { defaultColorResolver, placeholderColors, resolveColorNode } from '@/core/drawingml/colors';
import { FEATURES, emuToPt, pt } from '@/core/ir';
import {
  poAttr,
  poChildren,
  poFindDescendant,
  poIntAttr,
  poIs,
  poTag,
  poText,
} from '@/core/po-helpers';
import {
  colorChangeOf,
  parseCustGeom,
  parseFill,
  parseLine,
  parsePrstGeom,
  parseShadow,
  parseXfrm,
  statesFill,
  styleRefFill,
  styleRefFontColor,
  styleRefLine,
} from '@/word/drawing-parser';
import { parseBullet } from '@/pptx/placeholder-cascade';
import { boxFromXfrm, parsePh, parseXfrmBox, rPrToRunProps } from '@/pptx/sp-helpers';
import {
  cellStyle,
  lineBorder,
  tableStyleFlags,
  tableStyleId,
  withCellStyle,
  withoutFill,
} from '@/pptx/table-style';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const OLE_URI = 'http://schemas.openxmlformats.org/presentationml/2006/ole';
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

/**
 * Per-slide parsing context: the placeholder cascade (PX2), an image resolver
 * that turns a slide-scoped relationship id (`a:blip @r:embed`) into a
 * {@link ResourceId} (PX3a), and a chart resolver that parses a referenced chart
 * part (`c:chart @r:id`) and returns its document-unique key (PX4a). All optional
 * — a bare slide needs none.
 */
export interface SlideContext {
  readonly cascade?: PlaceholderCascade;
  /** A slide-scoped blip relationship id (`a:blip @r:embed`) → a stored resource (PX3a). */
  readonly resolveImage?: (relId: string) => ResourceId | undefined;
  /** The deck theme's fill style lists, for a `p:bgRef` on this slide. */
  readonly themeFills?: ThemeFillStyles;
  /** §20.1.4.1.21 `a:lnStyleLst` — the widths an `a:lnRef` indexes, in points. */
  readonly themeLineWidths?: ReadonlyArray<number>;
  /**
   * §19.3.1.43 — the slide's own background fill, which a shape marked
   * `useBgFill` is painted with.
   */
  readonly backgroundFill?: ShapeFill;
  /** The slide's size, which says WHERE that background lies under a shape. */
  readonly slideSize?: { readonly widthPt: Pt; readonly heightPt: Pt };
  /**
   * §19.3.2.4 — the picture an embedded object shows, by the `@spid` of its
   * shape in the slide's legacy VML drawing.
   */
  readonly resolveOlePreview?: (spid: string) => ResourceId | undefined;
  /**
   * §20.1.4.2.24 — the `a:tblStyle` a table names by GUID, from the deck's
   * `tableStyles.xml`. A table that names none wears none.
   */
  readonly resolveTableStyle?: (styleId: string | undefined) => PoNode | undefined;
  /** A chart relationship id (`c:chart @r:id`) → its document-unique key (PX4a). */
  readonly resolveChart?: (relId: string) => string | undefined;
  /**
   * The deck's colour resolver (master theme palette, PX5); defaults to the
   * Office palette when absent.
   */
  readonly colors?: ColorResolver;
  /** A run hyperlink (`a:hlinkClick @r:id`) → its external URL (PX6). */
  readonly resolveHyperlink?: (relId: string) => string | undefined;
  /**
   * A SmartArt data relationship (`dgm:relIds @r:dm`) → the diagram's pre-rendered
   * drawing override (its `dsp:spTree`), or `undefined` when the file ships no
   * override (E-SMARTART SA0).
   */
  readonly resolveDiagram?: (relId: string) => PoNode | undefined;
  /**
   * Sink for graceful-degradation notices (E-SMARTART SA3): a SmartArt that
   * declares a diagram but ships no drawing override records a dropped-feature
   * {@link Loss} here rather than vanishing without a trace.
   */
  readonly onLoss?: (loss: Loss) => void;
}

type LinkResolver = ((relId: string) => string | undefined) | undefined;

const ALGN_TO_ALIGNMENT: Readonly<Record<string, Alignment>> = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'both',
  dist: 'distribute',
};

const RECT_GEOMETRY: ShapeGeometry = { kind: 'preset', preset: 'rect', adjust: new Map() };

// A group transform maps a child-space EMU box to slide-space EMU (PX5c): a
// p:grpSp positions its children in its own coordinate frame (a:chOff/a:chExt),
// which scales + offsets onto the group's slide box (a:off/a:ext).
type GroupTransform = (box: ShapeBoxEmu) => ShapeBoxEmu;
const IDENTITY_TRANSFORM: GroupTransform = (box) => box;

/**
 * Walk a shape container (`p:spTree` or a `p:grpSp`), turning each `p:sp` into a
 * floating text/graphic shape, each `p:pic` into a floating image, each
 * `p:graphicFrame` into a chart/table, and recursing into nested `p:grpSp` groups
 * (composing their transforms).
 *
 * @param container The shape container node.
 * @param ctx       The per-slide parsing context.
 * @param transform Maps child-space boxes to the slide (identity at the top level).
 * @returns The container's shapes as positioned {@link BodyElement}s.
 */
export function parseSlideShapes(
  container: PoNode,
  ctx: SlideContext = {},
  transform: GroupTransform = IDENTITY_TRANSFORM,
  skipPlaceholders = false,
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  for (const child of poChildren(container)) {
    // On a master or a layout, a placeholder is a PROTOTYPE — the geometry and
    // text properties a slide's own placeholder inherits (PX2), and its text is
    // the prompt PowerPoint shows while editing. Drawn as a shape it would put
    // "Click to edit Master title style" on every slide.
    if (skipPlaceholders && isPlaceholder(child)) continue;
    if (poIs(child, 'p:sp')) {
      const shape = parseSp(child, ctx, transform);
      if (shape) out.push({ kind: 'shape', shape });
    } else if (poIs(child, 'p:pic')) {
      const image = parsePic(child, ctx, transform);
      if (image) out.push({ kind: 'image', image });
    } else if (poIs(child, 'p:graphicFrame')) {
      out.push(...parseGraphicFrame(child, ctx, transform));
    } else if (poIs(child, 'p:grpSp')) {
      out.push(
        ...parseSlideShapes(child, ctx, composeGroupTransform(child, transform), skipPlaceholders),
      );
    }
  }
  return out;
}

/**
 * Whether a shape is a placeholder — `p:ph` in its OWN non-visual properties.
 * Read as "anywhere below", a group holding one would take the whole group
 * with it.
 */
function isPlaceholder(shape: PoNode): boolean {
  const nv = poChildren(shape).find((c) => poTag(c)?.startsWith('p:nv') === true);
  const nvPr = nv ? poChildren(nv).find((c) => poIs(c, 'p:nvPr')) : undefined;
  return nvPr !== undefined && poChildren(nvPr).some((c) => poIs(c, 'p:ph'));
}

// p:grpSpPr/a:xfrm → a child→slide transform composed under the parent's. No (or
// degenerate) xfrm leaves the parent transform unchanged.
function composeGroupTransform(grpSp: PoNode, parent: GroupTransform): GroupTransform {
  const grpSpPr = poChildren(grpSp).find((c) => poIs(c, 'p:grpSpPr'));
  const xfrm = grpSpPr ? poChildren(grpSpPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  if (!xfrm) return parent;
  const off = poChildren(xfrm).find((c) => poIs(c, 'a:off'));
  const ext = poChildren(xfrm).find((c) => poIs(c, 'a:ext'));
  const chOff = poChildren(xfrm).find((c) => poIs(c, 'a:chOff'));
  const chExt = poChildren(xfrm).find((c) => poIs(c, 'a:chExt'));
  const extCx = ext ? poIntAttr(ext, 'cx') : undefined;
  const extCy = ext ? poIntAttr(ext, 'cy') : undefined;
  const chExtCx = chExt ? poIntAttr(chExt, 'cx') : undefined;
  const chExtCy = chExt ? poIntAttr(chExt, 'cy') : undefined;
  if (!extCx || !extCy || !chExtCx || !chExtCy) return parent;
  const offX = (off ? poIntAttr(off, 'x') : undefined) ?? 0;
  const offY = (off ? poIntAttr(off, 'y') : undefined) ?? 0;
  const chOffX = (chOff ? poIntAttr(chOff, 'x') : undefined) ?? 0;
  const chOffY = (chOff ? poIntAttr(chOff, 'y') : undefined) ?? 0;
  const sx = extCx / chExtCx;
  const sy = extCy / chExtCy;
  return (box) =>
    parent({
      x: offX + (box.x - chOffX) * sx,
      y: offY + (box.y - chOffY) * sy,
      cx: box.cx * sx,
      cy: box.cy * sy,
    });
}

// A page-absolute float anchor at the shape's EMU offset (the slide is the page).
function floatAt(box: ShapeBoxEmu): FloatAnchor {
  return {
    wrap: 'none',
    posH: { relativeFrom: 'page', offsetPt: emuToPt(box.x) },
    posV: { relativeFrom: 'page', offsetPt: emuToPt(box.y) },
  };
}

// p:sp → a floating shape: its geometry, fill and stroke (PX3), plus a text body
// (PX1/PX2). The box comes from the shape's own a:xfrm, else (for a placeholder)
// the cascade. Undefined when there is no geometry, or the shape is entirely
// invisible (no fill, no stroke, no text).
function parseSp(sp: PoNode, ctx: SlideContext, transform: GroupTransform): ShapeBlock | undefined {
  const ph = parsePh(sp);
  const colors = ctx.colors ?? defaultColorResolver;
  const spPr = poChildren(sp).find((c) => poIs(c, 'p:spPr'));
  let own: ShapeBoxEmu | undefined = parseXfrmBox(spPr);
  if (!own && ph && ctx.cascade) own = ctx.cascade.geometryFor(ph);
  if (!own) return undefined;
  const box = transform(own);

  // §20.1.4.2.10/§20.1.4.2.19/§20.1.4.2.14 — a shape drawn from a gallery keeps
  // its fill, its outline and the colour of its TEXT in `p:style` and carries
  // none of them in `spPr`: customGeo's title banner and the ellipse under it
  // are a theme gradient each, and read from `spPr` alone they were an empty
  // outline and nothing at all.
  const style = poChildren(sp).find((c) => poIs(c, 'p:style'));
  const txBody = poChildren(sp).find((c) => poIs(c, 'p:txBody'));
  // The style's colour is read BEFORE the text, because it belongs UNDER the
  // run's own and OVER the deck's default: themes.pptx's green box asks for
  // `lt1`, and stamped on afterwards it lost to the `tx1` every run inherits
  // from `p:defaultTextStyle` — white text drawn black on green.
  const styleColor = style ? styleRefFontColor(style, colors) : undefined;
  const text = txBody
    ? parseTxBody(txBody, ph, ctx.cascade, colors, ctx.resolveHyperlink, styleColor)
    : undefined;

  // Geometry/fill/stroke from p:spPr via the shared DrawingML readers, resolving
  // colours through the deck's theme palette (PX5).
  // A placeholder inherits what it does not state — its geometry, its fill and
  // its outline — from the prototype in the layout, else the master.
  // tdf95932's slide holds one placeholder with a word in it; the green
  // rounded panel it sits on, and the white the word is written in, are both
  // the layout's, and without the panel the word was white on white.
  const proto = ph && ctx.cascade ? ctx.cascade.shapePropsFor(ph) : undefined;
  const geomFrom = spPr && statesGeometry(spPr) ? spPr : (proto ?? spPr);
  const geometry = parseGeometry(geomFrom);
  // §19.3.1.43 `p:sp@useBgFill` — the shape is filled with the SLIDE's
  // background, which is how a deck cuts a hole in the decoration above it:
  // tdf93868's master lays a white rectangle over the whole slide and then a
  // rounded one on top that lets the slide's black gradient back through.
  const useBgFill = poAttr(sp, 'useBgFill') === '1';
  const fillFrom = statesFill(spPr) ? spPr : (proto ?? spPr);
  const fill: ShapeFill = useBgFill
    ? backgroundThrough(ctx.backgroundFill, box, ctx.slideSize)
    : fillFrom
      ? parseFill(fillFrom, colors, ctx.resolveImage)
      : { kind: 'none' };
  const lineFrom = spPr && poChildren(spPr).some((c) => poIs(c, 'a:ln')) ? spPr : (proto ?? spPr);
  let line = lineFrom ? parseLine(lineFrom, colors) : undefined;
  const themeStyles = ctx.themeFills
    ? { fills: ctx.themeFills.fills, bgFills: ctx.themeFills.backgrounds }
    : undefined;
  const styled: ShapeFill =
    style && fill.kind === 'none' && !statesFill(fillFrom)
      ? styleRefFill(style, colors, themeStyles)
      : fill;
  if (style) {
    const fromStyle = styleRefLine(style, colors, ctx.themeLineWidths);
    if (fromStyle) {
      line = line
        ? {
            ...fromStyle,
            ...Object.fromEntries(Object.entries(line).filter(([, v]) => v !== undefined)),
          }
        : fromStyle;
    }
  }
  const visibleLine = line !== undefined && line.fill !== 'none';
  if (!text && styled.kind === 'none' && !visibleLine) return undefined;

  // §20.1.7.6 — a shape may be turned or mirrored in its box, and a slide says
  // so on the same `a:xfrm` its position comes from. Unread, the blue triangle
  // ArtisticEffectSample's layout stands on its side (`rot="5400000"`) pointed
  // up instead of right.
  const xfrm = spPr ? poChildren(spPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  const spin = xfrm ? parseXfrm(xfrm) : undefined;

  return {
    float: floatAt(box),
    width: emuToPt(box.cx),
    height: emuToPt(box.cy),
    geometry,
    fill: styled,
    ...(line ? { line } : {}),
    ...(spin && Object.keys(spin).length > 0 ? { transform: spin } : {}),
    ...(text ? { text } : {}),
    paragraphProperties: {},
  };
}

// p:graphicFrame → floating chart (c:chart, PX4a), table (a:tbl, PX4b) or a
// SmartArt diagram's shapes (dgm, E-SMARTART SA0). Returns an array because a
// diagram expands to many shapes; chart/table yield one element. The frame's
// transform is p:xfrm (a:off + a:ext), not the a:xfrm of a shape.
function parseGraphicFrame(
  gf: PoNode,
  ctx: SlideContext,
  transform: GroupTransform,
): Array<BodyElement> {
  const own = boxFromXfrm(poChildren(gf).find((c) => poIs(c, 'p:xfrm')));
  if (!own) return [];
  const box = transform(own);
  const graphicData = poFindDescendant(gf, 'a:graphicData');
  const uri = graphicData ? poAttr(graphicData, 'uri') : undefined;

  if (uri === CHART_URI) {
    const cChart = poFindDescendant(gf, 'c:chart');
    const relId = cChart ? poAttr(cChart, 'id') : undefined; // r:id
    const key = relId !== undefined ? ctx.resolveChart?.(relId) : undefined;
    if (key === undefined) return [];
    const chart: ChartBlock = {
      float: floatAt(box),
      chartRelId: key,
      width: emuToPt(box.cx),
      height: emuToPt(box.cy),
      paragraphProperties: {},
    };
    return [{ kind: 'chart', chart }];
  }

  if (uri === TABLE_URI) {
    // §19.3.1.19 — a slide table stands where its FRAME says, like every other
    // graphic on the slide: the same float anchor a docx floating table uses.
    // Flowing it in-block put every table on the corpus at the top-left corner,
    // over the title (table_test2, and two conference decks besides).
    const tbl = poFindDescendant(gf, 'a:tbl');
    if (!tbl) return [];
    const table = parseTable(tbl, ctx.colors ?? defaultColorResolver, ctx.resolveTableStyle);
    return [
      {
        kind: 'table',
        table: { ...table, properties: { ...table.properties, float: floatAt(box) } },
      },
    ];
  }

  // §19.3.2.4 — an embedded OLE object shows a snapshot of itself. Modern decks
  // put a `p:pic` inside the `p:oleObj`; older ones point at a shape in the
  // slide's VML drawing with `@spid`, and the picture hangs off that.
  if (uri === OLE_URI) {
    const oleObj = poFindDescendant(gf, 'p:oleObj');
    if (!oleObj) return [];
    const pic = poChildren(oleObj).find((c) => poIs(c, 'p:pic'));
    if (pic) {
      const image = parsePic(pic, ctx, transform);
      return image ? [{ kind: 'image', image }] : [];
    }
    const spid = poAttr(oleObj, 'spid');
    const resource = spid !== undefined ? ctx.resolveOlePreview?.(spid) : undefined;
    if (resource === undefined) {
      ctx.onLoss?.({
        severity: 'dropped',
        feature: FEATURES.images,
        detail: 'embedded object with no preview picture — the object is not drawn',
      });
      return [];
    }
    return [
      {
        kind: 'image',
        image: {
          float: floatAt(box),
          resource,
          width: emuToPt(box.cx),
          height: emuToPt(box.cy),
          paragraphProperties: {},
        },
      },
    ];
  }

  // SmartArt: render the pre-rendered drawing override (dsp:spTree) as floating
  // shapes positioned within the frame box. No override ⇒ no shapes (SA0).
  if (uri === DIAGRAM_URI) {
    const relIds = poFindDescendant(gf, 'dgm:relIds');
    const dmRelId = relIds ? poAttr(relIds, 'dm') : undefined; // r:dm → data part
    const spTree = dmRelId !== undefined ? ctx.resolveDiagram?.(dmRelId) : undefined;
    if (!spTree) {
      // SmartArt is declared but ships no pre-rendered drawing override; record
      // a graceful loss instead of silently dropping the diagram (SA3).
      if (dmRelId !== undefined) ctx.onLoss?.(noDiagramOverrideLoss());
      return [];
    }
    const drawn = parseDiagramDrawing(
      spTree,
      diagramTransform(spTree, box),
      floatAt,
      ctx.colors ?? defaultColorResolver,
      ctx.resolveHyperlink,
    );
    // A drawing override that holds no shapes is a stub — the producer left the
    // layout to the reader, which is the same position as having none at all.
    if (drawn.length === 0) ctx.onLoss?.(noDiagramOverrideLoss());
    return drawn.map((shape) => ({ kind: 'shape', shape }));
  }
  return [];
}

/**
 * The diagram's child shapes live in the spTree's own coordinate space
 * (`dsp:grpSpPr/a:xfrm` `chOff`/`chExt`); map that onto a target box (the frame on
 * a slide, or the inline/anchored box in docx). Usually the child space equals
 * the box, so the scale is 1. Shared by pptx and docx (E-SMARTART).
 *
 * @param spTree The diagram drawing's `dsp:spTree`.
 * @param frame  The target box the diagram is placed into.
 * @returns A transform mapping diagram-space boxes onto `frame`.
 */
export function diagramTransform(spTree: PoNode, frame: ShapeBoxEmu): GroupTransform {
  const grpSpPr = poChildren(spTree).find((c) => poIs(c, 'dsp:grpSpPr'));
  const xfrm = grpSpPr ? poChildren(grpSpPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  const chOff = xfrm ? poChildren(xfrm).find((c) => poIs(c, 'a:chOff')) : undefined;
  const chExt = xfrm ? poChildren(xfrm).find((c) => poIs(c, 'a:chExt')) : undefined;
  const chExtCx = chExt ? poIntAttr(chExt, 'cx') : undefined;
  const chExtCy = chExt ? poIntAttr(chExt, 'cy') : undefined;
  const chOffX = (chOff ? poIntAttr(chOff, 'x') : undefined) ?? 0;
  const chOffY = (chOff ? poIntAttr(chOff, 'y') : undefined) ?? 0;
  const sx = chExtCx && chExtCx > 0 ? frame.cx / chExtCx : 1;
  const sy = chExtCy && chExtCy > 0 ? frame.cy / chExtCy : 1;
  return (b) => ({
    x: frame.x + (b.x - chOffX) * sx,
    y: frame.y + (b.y - chOffY) * sy,
    cx: b.cx * sx,
    cy: b.cy * sy,
  });
}

/**
 * A {@link Loss} for a SmartArt diagram that declares its data part but ships no
 * pre-rendered drawing override (older files, or a generator that omitted the
 * fallback). Ream renders the override rather than executing Office's layout
 * engine, so without it the diagram can't be drawn — this records the gap as a
 * dropped feature instead of letting it vanish. Shared by pptx and docx
 * (E-SMARTART SA3).
 *
 * @param where Optional location tag for the loss report (e.g. `slide 3`).
 */
export function noDiagramOverrideLoss(where?: string): Loss {
  return {
    severity: 'dropped',
    feature: FEATURES.smartArt,
    detail:
      'SmartArt diagram has no pre-rendered drawing override; its layout is not reconstructed',
    ...(where ? { where } : {}),
  };
}

/** One SmartArt node: the box it occupies in the target space, and its shape. */
export interface DiagramNode {
  readonly box: ShapeBoxEmu;
  readonly shape: ShapeBlock;
}

/**
 * Read a SmartArt drawing override (a `dsp:spTree`) into its nodes. The `dsp:`
 * wrapper holds an ordinary `a:` `spPr`/`txBody`, so the shared DrawingML
 * readers apply unchanged. Shared by pptx and docx (E-SMARTART); diagrams carry
 * no placeholder cascade.
 *
 * @param spTree       The diagram drawing's `dsp:spTree`.
 * @param transform    Maps each shape's diagram-space box to the target space.
 * @param colors       The colour resolver for the shapes' fills/strokes/text.
 * @param resolveLink  A run hyperlink resolver, or `undefined`.
 * @param resolveImage Resolves a picture fill's `r:embed` against the DRAWING
 *                     part's own relationships, or `undefined`.
 * @returns The diagram's visible nodes, in drawing order.
 */
export function parseDiagramNodes(
  spTree: PoNode,
  transform: GroupTransform,
  colors: ColorResolver,
  resolveLink: LinkResolver,
  resolveImage?: (relId: string) => ResourceId | undefined,
): Array<DiagramNode> {
  const out: Array<DiagramNode> = [];
  for (const sp of poChildren(spTree)) {
    if (!poIs(sp, 'dsp:sp')) continue;
    const spPr = poChildren(sp).find((c) => poIs(c, 'dsp:spPr'));
    const own = parseXfrmBox(spPr);
    if (!own) continue;
    const box = transform(own);

    const txBody = poChildren(sp).find((c) => poIs(c, 'dsp:txBody'));
    const parsed = txBody
      ? parseTxBody(txBody, undefined, undefined, colors, resolveLink)
      : undefined;
    // §20.1.4.2.14 — a diagram node states its text colour nowhere in the runs;
    // it is in the node's `dsp:style/a:fontRef`, and for the stock SmartArt
    // galleries that is `lt1`. smartart.docx puts white labels on three blue
    // boxes and we drew them black.
    const geometry = parseGeometry(spPr);
    const fill: ShapeFill = spPr ? parseFill(spPr, colors, resolveImage) : { kind: 'none' };
    const text = parsed ? withDiagramFontColor(parsed, sp, colors, fill) : undefined;
    const line = spPr ? parseLine(spPr, colors) : undefined;
    const shadow = spPr ? parseShadow(spPr, colors) : undefined;
    const visibleLine = line !== undefined && line.fill !== 'none';
    if (!text && fill.kind === 'none' && !visibleLine) continue;

    // §20.1.7.6 — a node may be turned in its box. fdo87488 stands one panel on
    // end with `rot="5400000"` and we drew it lying down, twice as wide as the
    // page. The LABEL does not turn with it: Word states the text's own frame
    // in `dsp:txXfrm` precisely so it stays upright, so a turned node hands its
    // text to a second, unturned box.
    const xfrm = spPr ? poChildren(spPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
    const spin = xfrm ? parseXfrm(xfrm) : undefined;
    const turned = (spin?.rotation60k ?? 0) % 21600000 !== 0;
    const txBox = turned
      ? boxFromXfrm(poChildren(sp).find((c) => poIs(c, 'dsp:txXfrm')))
      : undefined;

    out.push({
      box,
      shape: {
        width: emuToPt(box.cx),
        height: emuToPt(box.cy),
        geometry,
        fill,
        ...(line ? { line } : {}),
        ...(text && !txBox ? { text } : {}),
        ...(spin && Object.keys(spin).length > 0 ? { transform: spin } : {}),
        ...(shadow ? { shadow } : {}),
        paragraphProperties: {},
      },
    });
    if (text && txBox) {
      const placed = transform(txBox);
      out.push({
        box: placed,
        shape: {
          width: emuToPt(placed.cx),
          height: emuToPt(placed.cy),
          geometry: RECT_GEOMETRY,
          fill: { kind: 'none' },
          text,
          paragraphProperties: {},
        },
      });
    }
  }
  return out;
}

/**
 * The same nodes as free-standing floating shapes — what a slide wants, where
 * every shape is anchored to the page.
 *
 * @param spTree      The diagram drawing's `dsp:spTree`.
 * @param transform   Maps each shape's diagram-space box to the target space.
 * @param makeFloat   Anchors a node's box.
 * @param colors      The colour resolver for the shapes' fills/strokes/text.
 * @param resolveLink A run hyperlink resolver, or `undefined`.
 * @returns The diagram's visible shapes as positioned {@link ShapeBlock}s.
 */
export function parseDiagramDrawing(
  spTree: PoNode,
  transform: GroupTransform,
  makeFloat: (box: ShapeBoxEmu) => FloatAnchor,
  colors: ColorResolver,
  resolveLink: LinkResolver,
): Array<ShapeBlock> {
  return parseDiagramNodes(spTree, transform, colors, resolveLink).map(({ box, shape }) => ({
    ...shape,
    float: makeFloat(box),
  }));
}

// The text body with the node's `dsp:style/a:fontRef` colour filled in wherever
// a run declares none.
//
// The font reference belongs to the STYLE, and its colour is chosen to sit on
// the fill the style's `a:fillRef` names — white lettering for a filled node.
// A node that paints itself something else keeps the colour anyway, EXCEPT
// where the two are the same: tdf149551's gears state a white fill and a white
// font reference, and the labels inside them were white on white.
function withDiagramFontColor(
  text: ShapeTextBody,
  sp: PoNode,
  colors: ColorResolver,
  fill: ShapeFill,
): ShapeTextBody {
  const style = poChildren(sp).find((c) => poIs(c, 'dsp:style'));
  const ref = style ? poChildren(style).find((c) => poIs(c, 'a:fontRef')) : undefined;
  const child = ref ? poChildren(ref).find((c) => poTag(c) !== undefined) : undefined;
  const colorHex = child ? resolveColorNode(child, colors) : undefined;
  if (colorHex === undefined) return text;
  if (fill.kind === 'solid' && fill.colorHex === colorHex) return text;
  return {
    ...text,
    content: text.content.map((block) =>
      block.kind === 'paragraph'
        ? {
            ...block,
            paragraph: {
              ...block.paragraph,
              runs: block.paragraph.runs.map((run) =>
                run.properties.colorHex === undefined
                  ? { ...run, properties: { ...run.properties, colorHex } }
                  : run,
              ),
            },
          }
        : block,
    ),
  };
}

/**
 * The theme fill styles a `p:bgRef` indexes into (§19.3.1.2), as the raw nodes
 * the fill readers take.
 */
export interface ThemeFillStyles {
  /** `a:fillStyleLst` — what an index of 1…999 names. */
  readonly fills: ReadonlyArray<PoNode>;
  /** `a:bgFillStyleLst` — what an index past 1000 names, 1001 being the first. */
  readonly backgrounds: ReadonlyArray<PoNode>;
  /**
   * A picture named inside a style belongs to the THEME part: the Office
   * "Berlin" background is a photograph the theme carries and its `r:embed`
   * means nothing in the master's relationships (corpus: tdf123684, where it
   * resolved to a different image entirely).
   */
  readonly resolveImage?: (relId: string) => ResourceId | undefined;
}

/**
 * `p:bg` → the background fill, or `undefined` when it declares none.
 *
 * Two spellings. `p:bgPr` carries the fill itself — solid, gradient, or a
 * PICTURE, which needs the owning part's image resolver since the blip's
 * relationship is scoped to the part the background is written in. `p:bgRef`
 * (§19.3.1.2) carries no fill at all: an index into the theme's style lists and
 * the colour to put wherever that style says `phClr`. Read as its colour alone
 * — which is all this did — a deck whose theme opens with a black-to-grey
 * gradient came out flat, and one whose first background slot is a picture came
 * out a single colour.
 *
 * Used for the slide's own background and the inherited layout/master one (PX5b).
 *
 * @param bg           The `p:bg` node.
 * @param colors       The deck's colour resolver.
 * @param resolveImage Resolver for a picture background's relationship, scoped
 *                     to the part this `p:bg` is written in.
 * @param theme        The theme's fill style lists, for a `p:bgRef`.
 */
export function parseBackgroundFill(
  bg: PoNode,
  colors: ColorResolver,
  resolveImage?: (relId: string) => ResourceId | undefined,
  theme?: ThemeFillStyles,
): ShapeFill | undefined {
  const bgPr = poChildren(bg).find((c) => poIs(c, 'p:bgPr'));
  if (bgPr) {
    const fill = parseFill(bgPr, colors, resolveImage);
    return fill.kind !== 'none' ? fill : undefined;
  }
  const bgRef = poChildren(bg).find((c) => poIs(c, 'p:bgRef'));
  if (!bgRef) return undefined;
  const child = poChildren(bgRef).find((c) => poTag(c) !== undefined);
  const hex = child ? resolveColorNode(child, colors) : undefined;
  const idx = poIntAttr(bgRef, 'idx');
  const slot =
    idx === undefined
      ? undefined
      : idx > 1000
        ? theme?.backgrounds[idx - 1001]
        : theme?.fills[idx - 1];
  if (slot) {
    // The slot is a whole fill in the theme's own vocabulary; read it with the
    // shared reader, under a resolver where `phClr` is the colour named here.
    const themed = parseFill(
      { 'a:spPr': [slot] },
      hex === undefined ? colors : placeholderColors(colors, hex),
      theme?.resolveImage ?? resolveImage,
    );
    if (themed.kind !== 'none') return themed;
  }
  return hex === undefined ? undefined : { kind: 'solid', colorHex: hex };
}

/**
 * A full-slide backdrop element for a background fill: a rectangle covering the
 * page, anchored behind the content (PX5b).
 *
 * @param fill     The resolved background fill.
 * @param widthPt  The page width in points.
 * @param heightPt The page height in points.
 */
export function backdropElement(fill: ShapeFill, widthPt: Pt, heightPt: Pt): BodyElement {
  return {
    kind: 'shape',
    shape: {
      float: {
        wrap: 'none',
        behind: true,
        posH: { relativeFrom: 'page', offsetPt: emuToPt(0) },
        posV: { relativeFrom: 'page', offsetPt: emuToPt(0) },
      },
      width: widthPt,
      height: heightPt,
      geometry: RECT_GEOMETRY,
      fill,
      paragraphProperties: {},
    },
  };
}

// §19.3.1.43 — a shape that wears the slide's background wears the PIECE of it
// that lies under the shape, not a copy squeezed into its box. A picture
// background is stretched over the whole slide, so the destination rect is the
// SLIDE stated in fractions of the shape (§20.1.8.30 `a:fillRect`, where a
// negative inset reaches outside the box and the shape's own outline clips it
// back). Squeezed, tdf123684's text box drew a little diagonal of its own.
function backgroundThrough(
  bg: ShapeFill | undefined,
  box: ShapeBoxEmu,
  slide: { readonly widthPt: Pt; readonly heightPt: Pt } | undefined,
): ShapeFill {
  if (!bg) return { kind: 'none' };
  if (bg.kind !== 'picture' || !slide || box.cx <= 0 || box.cy <= 0) return bg;
  const [x, y] = [emuToPt(box.x), emuToPt(box.y)];
  const [w, h] = [emuToPt(box.cx), emuToPt(box.cy)];
  return {
    ...bg,
    imageFillRect: {
      left: -x / w,
      top: -y / h,
      right: -(slide.widthPt - x - w) / w,
      bottom: -(slide.heightPt - y - h) / h,
    },
  };
}

/**
 * The same element, painted BEHIND the slide's own content.
 *
 * §19.3.1 — a slide's shape tree is drawn over its layout's, and that over the
 * master's; nothing a deck states once can bury what a slide puts on it. Our
 * page paints by KIND (every image, then every shape), so without this a
 * layout's white card landed on top of the photograph the slide is made of —
 * themes.pptx's fifth slide is one picture and it vanished under the card.
 *
 * @param el The inherited element.
 * @returns The element with its float sunk to the backdrop layer.
 */
export function asBackdrop(el: BodyElement): BodyElement {
  const sink = <T extends { readonly float?: FloatAnchor }>(item: T): T =>
    item.float ? { ...item, float: { ...item.float, behind: true } } : item;
  switch (el.kind) {
    case 'shape':
      return { ...el, shape: sink(el.shape) };
    case 'image':
      return { ...el, image: sink(el.image) };
    case 'table':
      return { ...el, table: { ...el.table, properties: sink(el.table.properties) } };
    case 'chart':
      return { ...el, chart: sink(el.chart) };
    default:
      return el;
  }
}

/**
 * `p:spPr` geometry: `a:prstGeom` (preset) or `a:custGeom` (custom path), default
 * rect. Exported for sheet shapes (E-SHEET W2), whose `xdr:spPr` carries the same
 * `a:` children.
 */
/** Whether an `spPr` states a geometry of its own, rather than inheriting one. */
function statesGeometry(spPr: PoNode): boolean {
  return poChildren(spPr).some((c) => poIs(c, 'a:prstGeom') || poIs(c, 'a:custGeom'));
}

export function parseGeometry(spPr: PoNode | undefined): ShapeGeometry {
  if (!spPr) return RECT_GEOMETRY;
  const prst = poChildren(spPr).find((c) => poIs(c, 'a:prstGeom'));
  if (prst) return parsePrstGeom(prst);
  const cust = poChildren(spPr).find((c) => poIs(c, 'a:custGeom'));
  if (cust) return parseCustGeom(cust);
  return RECT_GEOMETRY;
}

// p:pic → a floating image. The bytes come from p:blipFill/a:blip @r:embed,
// resolved against the slide's relationships (PX3a); geometry from p:spPr/a:xfrm
// (picture placeholders that inherit it from the layout wait for a later slice).
function parsePic(
  pic: PoNode,
  ctx: SlideContext,
  transform: GroupTransform,
): ImageBlock | undefined {
  const spPr = poChildren(pic).find((c) => poIs(c, 'p:spPr'));
  const own = parseXfrmBox(spPr);
  if (!own) return undefined;
  const box = transform(own);

  const blipFill = poChildren(pic).find((c) => poIs(c, 'p:blipFill'));
  const blip = blipFill ? poChildren(blipFill).find((c) => poIs(c, 'a:blip')) : undefined;
  const relId = blip ? poAttr(blip, 'embed') : undefined;
  const resource = relId !== undefined ? ctx.resolveImage?.(relId) : undefined;
  // §20.1.8.16 — a picture may declare one of its colours away, which is how a
  // logo drawn on white sits on a dark slide.
  const colorChange = blip ? colorChangeOf(blip, ctx.colors ?? defaultColorResolver) : undefined;
  // §20.1.8.4 `a:alphaModFix` — how opaque the picture is DRAWN. A layout that
  // lays a photograph behind its title sets it low (ArtisticEffectSample's
  // cover shows one at 52%), and drawn full-strength the words are unreadable.
  const fixed = blip ? poChildren(blip).find((c) => poIs(c, 'a:alphaModFix')) : undefined;
  const amt = fixed ? poIntAttr(fixed, 'amt') : undefined;
  const alpha = amt === undefined ? undefined : Math.min(1, Math.max(0, amt / 100000));

  const altText = picAltText(pic);
  return {
    float: floatAt(box),
    ...(resource !== undefined ? { resource } : {}),
    ...(colorChange ? { colorChange } : {}),
    ...(alpha !== undefined && alpha < 1 ? { alpha } : {}),
    width: emuToPt(box.cx),
    height: emuToPt(box.cy),
    paragraphProperties: {},
    ...(altText ? { altText } : {}),
  };
}

// p:nvPicPr/p:cNvPr @descr (preferred) or @title → the picture's alternate text.
function picAltText(pic: PoNode): string | undefined {
  const nvPicPr = poChildren(pic).find((c) => poIs(c, 'p:nvPicPr'));
  const cNvPr = nvPicPr ? poChildren(nvPicPr).find((c) => poIs(c, 'p:cNvPr')) : undefined;
  const descr = cNvPr ? poAttr(cNvPr, 'descr') : undefined;
  const title = cNvPr ? poAttr(cNvPr, 'title') : undefined;
  return (descr ?? title)?.trim() || undefined;
}

// a:txBody → its paragraphs as BodyElements (shared by shape text bodies and
// table cells). Runs inherit the placeholder defaults when a cascade is given,
// and resolve colours through the deck's palette.
function txBodyParagraphs(
  txBody: PoNode,
  ph: PlaceholderRef | undefined,
  cascade: PlaceholderCascade | undefined,
  colors: ColorResolver,
  resolveLink: LinkResolver,
  styleColor?: string,
): Array<BodyElement> {
  const content: Array<BodyElement> = [];
  const counters: Array<number> = []; // per-level a:buAutoNum counters (PX6b)
  for (const child of poChildren(txBody)) {
    if (!poIs(child, 'a:p')) continue;
    content.push({
      kind: 'paragraph',
      paragraph: parseSlideParagraph(child, ph, cascade, colors, resolveLink, counters, styleColor),
    });
  }
  return content;
}

/**
 * `p:txBody` → {@link ShapeTextBody}. `a:bodyPr` carries the insets and the
 * vertical anchor (PX6: anchor `t`/`ctr`/`b`). Exported for sheet shapes
 * (E-SHEET W2) — like SmartArt, they pass no placeholder cascade, so runs use
 * their direct `a:rPr` formatting.
 *
 * @param txBody      The `p:txBody` node.
 * @param ph          The owning placeholder ref, when this is a placeholder shape.
 * @param cascade     The placeholder cascade supplying inherited run defaults.
 * @param colors      The colour resolver for run colours.
 * @param resolveLink A run hyperlink resolver, or `undefined`.
 * @returns The text body, or `undefined` when it holds no paragraphs.
 */
export function parseTxBody(
  txBody: PoNode,
  ph: PlaceholderRef | undefined,
  cascade: PlaceholderCascade | undefined,
  colors: ColorResolver,
  resolveLink: LinkResolver,
  styleColor?: string,
): ShapeTextBody | undefined {
  const content = txBodyParagraphs(txBody, ph, cascade, colors, resolveLink, styleColor);
  if (content.length === 0) return undefined;

  const bodyPr = poChildren(txBody).find((c) => poIs(c, 'a:bodyPr'));
  const lIns = bodyPr ? poIntAttr(bodyPr, 'lIns') : undefined;
  const tIns = bodyPr ? poIntAttr(bodyPr, 'tIns') : undefined;
  const rIns = bodyPr ? poIntAttr(bodyPr, 'rIns') : undefined;
  const bIns = bodyPr ? poIntAttr(bodyPr, 'bIns') : undefined;
  const a = bodyPr ? poAttr(bodyPr, 'anchor') : undefined;
  // A placeholder that states no anchor of its own sits where its prototype
  // says: a master title anchored `ctr` centres the slide's title in its box.
  const anchor: ShapeTextBody['anchor'] | undefined =
    a === 'ctr' || a === 'b' || a === 't' ? a : ph && cascade ? cascade.anchorFor(ph) : undefined;
  return {
    content,
    ...(lIns !== undefined ? { insetLeft: emuToPt(lIns) } : {}),
    ...(tIns !== undefined ? { insetTop: emuToPt(tIns) } : {}),
    ...(rIns !== undefined ? { insetRight: emuToPt(rIns) } : {}),
    ...(bIns !== undefined ? { insetBottom: emuToPt(bIns) } : {}),
    ...(anchor ? { anchor } : {}),
  };
}

// a:p → Paragraph. The outline level (a:pPr @lvl) selects the placeholder's
// default run formatting; @algn sets the alignment; @marL/@indent (else a
// per-level default) the indent; a:buChar/a:buAutoNum a materialized list marker
// run (PX6). Runs come from a:r and a:fld (a text field's cached a:t).
function parseSlideParagraph(
  aP: PoNode,
  ph: PlaceholderRef | undefined,
  cascade: PlaceholderCascade | undefined,
  colors: ColorResolver,
  resolveLink: LinkResolver,
  counters: Array<number>,
  styleColor?: string,
): Paragraph {
  const pPr = poChildren(aP).find((c) => poIs(c, 'a:pPr'));
  const level = (pPr ? poIntAttr(pPr, 'lvl') : undefined) ?? 0;
  const defaults: RunProperties = {
    ...(cascade ? cascade.defaultsFor(ph, level) : {}),
    ...(styleColor !== undefined ? { colorHex: styleColor } : {}),
  };
  // What the deck, the master and the prototypes say this paragraph looks like
  // — the paragraph's own pPr states only where it differs.
  const inherited: ParagraphProperties = cascade ? cascade.paragraphDefaultsFor(ph, level) : {};
  const algn = pPr ? poAttr(pPr, 'algn') : undefined;
  const alignment = algn !== undefined ? ALGN_TO_ALIGNMENT[algn] : undefined;

  const runs: Array<Run> = [];
  for (const child of poChildren(aP)) {
    if (poIs(child, 'a:r') || poIs(child, 'a:fld')) {
      const run = parseSlideRun(child, defaults, colors, resolveLink);
      if (run) runs.push(run);
    }
  }
  // §21.1.2.2.7 — marL is the text's left margin, @indent the first-line/hang.
  // Absent: indent nested levels by a default 0.5" per level (457200 EMU).
  const marL = pPr ? poIntAttr(pPr, 'marL') : undefined;
  const indent = pPr ? poIntAttr(pPr, 'indent') : undefined;
  const indentLeft =
    marL !== undefined ? emuToPt(marL) : level > 0 ? emuToPt(level * 457200) : undefined;
  const hanging = (indent !== undefined ? emuToPt(indent) : (inherited.indentFirstLine ?? 0)) < 0;
  const marker = bulletMarker(pPr, level, counters, cascade?.bulletFor(ph, level));
  if (marker !== undefined) {
    // §21.1.2.4.2/.3 — a bullet is drawn at the size of the TEXT IT LEADS,
    // scaled by `a:buSzPct` or overruled by `a:buSzPts`. Drawn at the level's
    // inherited size instead, ArtisticEffectSample's dots came out 48pt in
    // front of 18pt text and every line stood three times too tall.
    const base = runs[0]?.properties ?? defaults;
    runs.unshift({
      // §17.3.1.12 — a hanging indent is itself a tab stop, and the tab after
      // the marker is what carries the text out to the body indent. Written as
      // two spaces the words began wherever the dot happened to end.
      text: marker + (hanging ? '\t' : '  '),
      properties: bulletProps(base, pPr, colors),
      listMarker: true,
    });
  }

  // §21.1.2.2.3 `a:endParaRPr` — the properties of the paragraph MARK. On a
  // paragraph with no runs it is the only thing that says how tall the line is,
  // and a blank line that measures the layout's default instead of the size the
  // file names is the wrong height: shape-macro-ext-ref.xlsx opens its button's
  // text with an empty 14pt paragraph, and collapsing it drew the caption 12pt
  // above where both references put it.
  if (runs.length === 0) {
    const endRPr = poChildren(aP).find((c) => poIs(c, 'a:endParaRPr'));
    const sz = endRPr ? poIntAttr(endRPr, 'sz') : undefined;
    if (sz !== undefined && sz > 0) {
      runs.push({ text: '', properties: { ...defaults, fontSizePt: pt(sz / 100) } });
    }
  }

  return {
    properties: {
      ...inherited,
      ...(alignment ? { alignment } : {}),
      ...(indentLeft !== undefined ? { indentLeft } : {}),
      ...(indent !== undefined ? { indentFirstLine: emuToPt(indent) } : {}),
    },
    runs,
  };
}

// §21.1.2.4 — the marker's own formatting: the size of the text it leads (a
// percentage of it, or a size stated outright) and the colour the paragraph
// gives it.
function bulletProps(
  base: RunProperties,
  pPr: PoNode | undefined,
  colors: ColorResolver,
): RunProperties {
  const child = (tag: string): PoNode | undefined =>
    pPr ? poChildren(pPr).find((c) => poIs(c, tag)) : undefined;
  const pct = poIntAttr(child('a:buSzPct'), 'val');
  const pts = poIntAttr(child('a:buSzPts'), 'val');
  const clr = child('a:buClr');
  const colorHex = clr
    ? poChildren(clr)
        .map((c) => resolveColorNode(c, colors))
        .find((hex) => hex !== undefined)
    : undefined;
  const sizePt =
    pts !== undefined
      ? pts / 100
      : pct !== undefined && base.fontSizePt !== undefined
        ? (base.fontSizePt * pct) / 100000
        : undefined;
  return {
    ...base,
    ...(sizePt !== undefined ? { fontSizePt: pt(sizePt) } : {}),
    ...(colorHex !== undefined ? { colorHex } : {}),
  };
}

// a:pPr bullet → the marker text to prepend, or undefined for no bullet (the
// gap after it is the caller's, since only a hanging indent earns a tab).
// a:buNone suppresses; a:buChar is literal; a:buAutoNum advances the per-level
// counter and formats it (PX6b).
function bulletMarker(
  pPr: PoNode | undefined,
  level: number,
  counters: Array<number>,
  inherited: LevelBullet | undefined,
): string | undefined {
  // A paragraph states a bullet only where it differs from the one its level
  // already carries — from the prototype, the master's family style or the
  // deck's default. Read as "no bullet unless this paragraph draws one", every
  // list that leaves its dot to the master came out flat (themes.pptx writes
  // one bare line on its second slide and the dot in front of it is nine
  // levels up).
  const bullet = parseBullet(pPr) ?? inherited;
  if (!bullet || bullet.kind === 'none') return undefined;
  if (bullet.kind === 'char') return bullet.char;
  const prev = counters[level];
  const n = (prev === undefined ? bullet.startAt - 1 : prev) + 1;
  counters[level] = n;
  counters.length = level + 1; // deeper levels restart
  return `${n}${autoNumSuffix(bullet.type)}`;
}

// The trailing punctuation of an a:buAutoNum type (…Period → '.', …ParenR/Both →
// ')', …Plain → ''). v1 numbers arabic; alpha/roman folds onto arabic.
function autoNumSuffix(type: string): string {
  if (type.endsWith('ParenR') || type.endsWith('ParenBoth')) return ')';
  if (type.endsWith('Period')) return '.';
  return '';
}

// a:r / a:fld → Run. The placeholder defaults sit under the run's own a:rPr, so
// direct formatting always wins. a:rPr/a:hlinkClick @r:id resolves to a run href
// (PX6).
function parseSlideRun(
  node: PoNode,
  defaults: RunProperties,
  colors: ColorResolver,
  resolveLink: LinkResolver,
): Run | undefined {
  const t = poChildren(node).find((c) => poIs(c, 'a:t'));
  const text = t ? poText(t) : '';
  if (text.length === 0) return undefined;
  const rPr = poChildren(node).find((c) => poIs(c, 'a:rPr'));
  const hlink = rPr ? poChildren(rPr).find((c) => poIs(c, 'a:hlinkClick')) : undefined;
  const linkId = hlink ? poAttr(hlink, 'id') : undefined; // r:id
  const href = linkId !== undefined ? resolveLink?.(linkId) : undefined;
  return {
    text,
    properties: { ...defaults, ...rPrToRunProps(rPr, colors) },
    ...(href ? { href } : {}),
  };
}

// §21.1.3 a:tbl → a FlowDoc Table: grid column widths (a:tblGrid/a:gridCol @w),
// rows (a:tr) and cells (a:tc) with their text and merge state.
function parseTable(
  tbl: PoNode,
  colors: ColorResolver,
  resolveTableStyle?: (styleId: string | undefined) => PoNode | undefined,
): Table {
  const grid: Array<Pt> = [];
  const tblGrid = poChildren(tbl).find((c) => poIs(c, 'a:tblGrid'));
  if (tblGrid) {
    for (const col of poChildren(tblGrid)) {
      if (poIs(col, 'a:gridCol')) grid.push(emuToPt(poIntAttr(col, 'w') ?? 0));
    }
  }
  const rows: Array<TableRow> = [];
  // The cells kept beside the parsed rows: the style a cell wears depends on
  // what the cell's own `a:tcPr` says, which the FlowDoc cell no longer carries.
  const sources: Array<Array<PoNode>> = [];
  for (const tr of poChildren(tbl)) {
    if (!poIs(tr, 'a:tr')) continue;
    const cells = rowCells(tr);
    sources.push(cells);
    rows.push(parseTableRow(tr, cells, colors));
  }
  // §20.1.4.2.24 — almost everything a table LOOKS like is in the style it
  // names: the header's fill, the banding, the rules between the cells.
  const tblPr = poChildren(tbl).find((c) => poIs(c, 'a:tblPr'));
  const style = resolveTableStyle?.(tableStyleId(tblPr));
  if (!style) return { properties: {}, grid, rows };
  const flags = tableStyleFlags(tblPr);
  const styled = rows.map((row, r) => ({
    ...row,
    cells: row.cells.map((cell, c) => {
      const at = { row: r, rowCount: rows.length, col: c, colCount: grid.length };
      const part = cellStyle(style, flags, at, colors);
      const tc = sources[r]?.[c];
      return withCellStyle(cell, tc && cellSaysNoFill(tc) ? withoutFill(part) : part);
    }),
  }));
  return { properties: {}, grid, rows: styled };
}

// The cells a row HAS: a horizontal-merge continuation is covered by the
// gridSpan origin, so it is dropped and the origin's colSpan carries the width
// (the FlowDoc model omits placeholder cells for spanned columns).
function rowCells(tr: PoNode): Array<PoNode> {
  return poChildren(tr).filter((tc) => poIs(tc, 'a:tc') && poAttr(tc, 'hMerge') !== '1');
}

function parseTableRow(tr: PoNode, cells: Array<PoNode>, colors: ColorResolver): TableRow {
  const h = poIntAttr(tr, 'h');
  // §21.1.3.18 `a:tr@h` — the height the row asks for, which it keeps unless
  // its own text needs more. Stated without a rule it was ignored entirely, and
  // every slide table came out as tight as its text (table_test2's rows are
  // 36pt, 29pt, 15pt and we drew four equal thin ones).
  return {
    properties: { ...(h !== undefined ? { height: emuToPt(h), heightRule: 'atLeast' } : {}) },
    cells: cells.map((tc) => parseTableCell(tc, colors)),
  };
}

function parseTableCell(tc: PoNode, colors: ColorResolver): TableCell {
  const txBody = poChildren(tc).find((c) => poIs(c, 'a:txBody'));
  const content = txBody ? txBodyParagraphs(txBody, undefined, undefined, colors, undefined) : [];
  const gridSpan = poIntAttr(tc, 'gridSpan');
  const rowSpan = poIntAttr(tc, 'rowSpan');
  // Vertical merge: the origin carries @rowSpan (→ 'start'); a continuation cell
  // carries @vMerge="1" (→ 'middle') and is kept so the column slot stays filled.
  const merge: CellMerge | undefined =
    rowSpan !== undefined && rowSpan > 1
      ? 'start'
      : poAttr(tc, 'vMerge') === '1'
        ? 'middle'
        : undefined;
  const tcPr = poChildren(tc).find((c) => poIs(c, 'a:tcPr'));
  const shadingHex = tcPr ? cellFillHex(tcPr, colors) : undefined;
  const borders = tcPr ? cellOwnBorders(tcPr, colors) : undefined;
  return {
    properties: {
      ...(gridSpan !== undefined && gridSpan > 1 ? { colSpan: gridSpan } : {}),
      ...(merge ? { merge } : {}),
      ...(shadingHex ? { shading: { colorHex: shadingHex } } : {}),
      ...(borders ? { borders } : {}),
    },
    content,
  };
}

// §21.1.3.17 — a cell states its own four rules as `a:lnL`/`a:lnR`/`a:lnT`/
// `a:lnB`, each an ordinary `a:ln`, and they beat the table style's. tdf164936
// is one empty cell with three blue rules and a fourth at zero alpha.
const CELL_LINES: ReadonlyArray<readonly [string, keyof CellBorders]> = [
  ['a:lnL', 'left'],
  ['a:lnR', 'right'],
  ['a:lnT', 'top'],
  ['a:lnB', 'bottom'],
];

function cellOwnBorders(tcPr: PoNode, colors: ColorResolver): CellBorders | undefined {
  const out: { -readonly [K in keyof CellBorders]?: Border } = {};
  for (const [tag, side] of CELL_LINES) {
    const ln = poChildren(tcPr).find((c) => poIs(c, tag));
    if (ln) out[side] = lineBorder(ln, colors);
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

// §20.1.8.44 `a:noFill` on the cell — the cell is transparent BY ITS OWN WORD,
// which is not the same as saying nothing: the table style's fill applies to
// the silent one and not to this one.
function cellSaysNoFill(tc: PoNode): boolean {
  const tcPr = poChildren(tc).find((c) => poIs(c, 'a:tcPr'));
  return tcPr !== undefined && poChildren(tcPr).some((c) => poIs(c, 'a:noFill'));
}

// a:tcPr/a:solidFill → the cell background hex (srgb or theme scheme colour).
function cellFillHex(tcPr: PoNode, colors: ColorResolver): string | undefined {
  const solidFill = poChildren(tcPr).find((c) => poIs(c, 'a:solidFill'));
  if (!solidFill) return undefined;
  for (const c of poChildren(solidFill)) {
    const hex = resolveColorNode(c, colors);
    if (hex) return hex;
  }
  return undefined;
}
