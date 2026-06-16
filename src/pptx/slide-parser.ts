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
  CellMerge,
  ChartBlock,
  FloatAnchor,
  ImageBlock,
  Paragraph,
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
import type { Pt, ResourceId } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import type { PlaceholderCascade } from '@/pptx/placeholder-cascade';
import type { PlaceholderRef, ShapeBoxEmu } from '@/pptx/sp-helpers';

import { defaultColorResolver, resolveColorNode } from '@/core/drawingml/colors';
import { emuToPt } from '@/core/ir';
import { poAttr, poChildren, poFindDescendant, poIntAttr, poIs, poText } from '@/core/po-helpers';
import { parseCustGeom, parseFill, parseLine, parsePrstGeom } from '@/word/drawing-parser';
import { boxFromXfrm, parsePh, parseXfrmBox, rPrToRunProps } from '@/pptx/sp-helpers';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

// Per-slide parsing context: the placeholder cascade (PX2), an image resolver
// that turns a slide-scoped relationship id (a:blip @r:embed) into a
// ResourceStore id (PX3a), and a chart resolver that parses a referenced chart
// part (c:chart @r:id) and returns its document-unique key (PX4a). All optional
// — a bare slide needs none.
export interface SlideContext {
  readonly cascade?: PlaceholderCascade;
  readonly resolveImage?: (relId: string) => ResourceId | undefined;
  readonly resolveChart?: (relId: string) => string | undefined;
  // The deck's colour resolver (master theme palette, PX5); defaults to the
  // Office palette when absent.
  readonly colors?: ColorResolver;
  // A run hyperlink (a:hlinkClick @r:id) → its external URL (PX6).
  readonly resolveHyperlink?: (relId: string) => string | undefined;
  // A SmartArt data relationship (dgm:relIds @r:dm) → the diagram's pre-rendered
  // drawing override (its dsp:spTree), or undefined when the file ships no
  // override (E-SMARTART SA0).
  readonly resolveDiagram?: (relId: string) => PoNode | undefined;
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

// Walk a shape container (p:spTree or a p:grpSp), turning each p:sp into a
// floating text/graphic shape, each p:pic into a floating image, each
// p:graphicFrame into a chart/table, and recursing into nested p:grpSp groups
// (composing their transforms). `transform` maps child-space boxes to the slide.
export function parseSlideShapes(
  container: PoNode,
  ctx: SlideContext = {},
  transform: GroupTransform = IDENTITY_TRANSFORM,
): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  for (const child of poChildren(container)) {
    if (poIs(child, 'p:sp')) {
      const shape = parseSp(child, ctx, transform);
      if (shape) out.push({ kind: 'shape', shape });
    } else if (poIs(child, 'p:pic')) {
      const image = parsePic(child, ctx, transform);
      if (image) out.push({ kind: 'image', image });
    } else if (poIs(child, 'p:graphicFrame')) {
      out.push(...parseGraphicFrame(child, ctx, transform));
    } else if (poIs(child, 'p:grpSp')) {
      out.push(...parseSlideShapes(child, ctx, composeGroupTransform(child, transform)));
    }
  }
  return out;
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

  const txBody = poChildren(sp).find((c) => poIs(c, 'p:txBody'));
  const text = txBody
    ? parseTxBody(txBody, ph, ctx.cascade, colors, ctx.resolveHyperlink)
    : undefined;

  // Geometry/fill/stroke from p:spPr via the shared DrawingML readers, resolving
  // colours through the deck's theme palette (PX5).
  const geometry = parseGeometry(spPr);
  const fill: ShapeFill = spPr ? parseFill(spPr, colors) : { kind: 'none' };
  const line = spPr ? parseLine(spPr, colors) : undefined;
  const visibleLine = line !== undefined && line.fill !== 'none';
  if (!text && fill.kind === 'none' && !visibleLine) return undefined;

  return {
    float: floatAt(box),
    width: emuToPt(box.cx),
    height: emuToPt(box.cy),
    geometry,
    fill,
    ...(line ? { line } : {}),
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
    // A FlowDoc Table has no float, so a slide table flows in-block (the reader
    // zeroes the slide margins so it sits at the top-left). Its exact frame
    // position is a later refinement.
    const tbl = poFindDescendant(gf, 'a:tbl');
    return tbl
      ? [{ kind: 'table', table: parseTable(tbl, ctx.colors ?? defaultColorResolver) }]
      : [];
  }

  // SmartArt: render the pre-rendered drawing override (dsp:spTree) as floating
  // shapes positioned within the frame box. No override ⇒ no shapes (SA0).
  if (uri === DIAGRAM_URI) {
    const relIds = poFindDescendant(gf, 'dgm:relIds');
    const dmRelId = relIds ? poAttr(relIds, 'dm') : undefined; // r:dm → data part
    const spTree = dmRelId !== undefined ? ctx.resolveDiagram?.(dmRelId) : undefined;
    if (!spTree) return [];
    const childTransform = diagramTransform(spTree, box);
    const out: Array<BodyElement> = [];
    for (const child of poChildren(spTree)) {
      if (poIs(child, 'dsp:sp')) {
        const shape = parseDspSp(child, ctx, childTransform);
        if (shape) out.push({ kind: 'shape', shape });
      }
    }
    return out;
  }
  return [];
}

// The diagram's child shapes live in the spTree's own coordinate space
// (dsp:grpSpPr/a:xfrm chOff/chExt); map that onto the frame's slide-space box.
// Usually the child space equals the frame, so the scale is 1.
function diagramTransform(spTree: PoNode, frame: ShapeBoxEmu): GroupTransform {
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

// A single SmartArt drawing shape (dsp:sp). Mirrors parseSp, but the dsp:
// wrapper holds an ordinary a: spPr/txBody, so the shared DrawingML readers
// apply unchanged. Diagrams have no placeholder cascade.
function parseDspSp(
  sp: PoNode,
  ctx: SlideContext,
  transform: GroupTransform,
): ShapeBlock | undefined {
  const colors = ctx.colors ?? defaultColorResolver;
  const spPr = poChildren(sp).find((c) => poIs(c, 'dsp:spPr'));
  const own = parseXfrmBox(spPr);
  if (!own) return undefined;
  const box = transform(own);

  const txBody = poChildren(sp).find((c) => poIs(c, 'dsp:txBody'));
  const text = txBody
    ? parseTxBody(txBody, undefined, undefined, colors, ctx.resolveHyperlink)
    : undefined;

  const geometry = parseGeometry(spPr);
  const fill: ShapeFill = spPr ? parseFill(spPr, colors) : { kind: 'none' };
  const line = spPr ? parseLine(spPr, colors) : undefined;
  const visibleLine = line !== undefined && line.fill !== 'none';
  if (!text && fill.kind === 'none' && !visibleLine) return undefined;

  return {
    float: floatAt(box),
    width: emuToPt(box.cx),
    height: emuToPt(box.cy),
    geometry,
    fill,
    ...(line ? { line } : {}),
    ...(text ? { text } : {}),
    paragraphProperties: {},
  };
}

// p:bg → the background fill, or undefined when none/unsupported. p:bgPr carries
// the fill directly (a:solidFill/a:gradFill — the same children parseFill reads;
// a picture background is not yet supported); p:bgRef (a theme fill-style index)
// is approximated by its colour child as a solid fill. Used for the slide's own
// background and the inherited layout/master one (PX5b).
export function parseBackgroundFill(bg: PoNode, colors: ColorResolver): ShapeFill | undefined {
  const bgPr = poChildren(bg).find((c) => poIs(c, 'p:bgPr'));
  if (bgPr) {
    const fill = parseFill(bgPr, colors);
    return fill.kind !== 'none' ? fill : undefined;
  }
  const bgRef = poChildren(bg).find((c) => poIs(c, 'p:bgRef'));
  if (bgRef) {
    for (const c of poChildren(bgRef)) {
      const hex = resolveColorNode(c, colors);
      if (hex) return { kind: 'solid', colorHex: hex };
    }
  }
  return undefined;
}

// A full-slide backdrop element for a background fill: a rectangle covering the
// page, anchored behind the content (PX5b).
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

// p:spPr geometry: a:prstGeom (preset) or a:custGeom (custom path), default rect.
function parseGeometry(spPr: PoNode | undefined): ShapeGeometry {
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

  const altText = picAltText(pic);
  return {
    float: floatAt(box),
    ...(resource !== undefined ? { resource } : {}),
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
): Array<BodyElement> {
  const content: Array<BodyElement> = [];
  const counters: Array<number> = []; // per-level a:buAutoNum counters (PX6b)
  for (const child of poChildren(txBody)) {
    if (!poIs(child, 'a:p')) continue;
    content.push({
      kind: 'paragraph',
      paragraph: parseSlideParagraph(child, ph, cascade, colors, resolveLink, counters),
    });
  }
  return content;
}

// p:txBody → ShapeTextBody. a:bodyPr carries the insets and the vertical anchor
// (PX6: anchor t/ctr/b).
function parseTxBody(
  txBody: PoNode,
  ph: PlaceholderRef | undefined,
  cascade: PlaceholderCascade | undefined,
  colors: ColorResolver,
  resolveLink: LinkResolver,
): ShapeTextBody | undefined {
  const content = txBodyParagraphs(txBody, ph, cascade, colors, resolveLink);
  if (content.length === 0) return undefined;

  const bodyPr = poChildren(txBody).find((c) => poIs(c, 'a:bodyPr'));
  const lIns = bodyPr ? poIntAttr(bodyPr, 'lIns') : undefined;
  const tIns = bodyPr ? poIntAttr(bodyPr, 'tIns') : undefined;
  const rIns = bodyPr ? poIntAttr(bodyPr, 'rIns') : undefined;
  const bIns = bodyPr ? poIntAttr(bodyPr, 'bIns') : undefined;
  const a = bodyPr ? poAttr(bodyPr, 'anchor') : undefined;
  const anchor: ShapeTextBody['anchor'] | undefined =
    a === 'ctr' ? 'ctr' : a === 'b' ? 'b' : a === 't' ? 't' : undefined;
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
): Paragraph {
  const pPr = poChildren(aP).find((c) => poIs(c, 'a:pPr'));
  const level = (pPr ? poIntAttr(pPr, 'lvl') : undefined) ?? 0;
  const defaults: RunProperties = ph && cascade ? cascade.defaultsFor(ph, level) : {};
  const algn = pPr ? poAttr(pPr, 'algn') : undefined;
  const alignment = algn !== undefined ? ALGN_TO_ALIGNMENT[algn] : undefined;

  const runs: Array<Run> = [];
  const marker = bulletMarker(pPr, level, counters);
  if (marker !== undefined) runs.push({ text: marker, properties: defaults, listMarker: true });
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
  return {
    properties: {
      ...(alignment ? { alignment } : {}),
      ...(indentLeft !== undefined ? { indentLeft } : {}),
      ...(indent !== undefined ? { indentFirstLine: emuToPt(indent) } : {}),
    },
    runs,
  };
}

// a:pPr bullet → the marker text to prepend (with trailing spacing), or
// undefined for no bullet. a:buNone suppresses; a:buChar is literal; a:buAutoNum
// advances the per-level counter and formats it (PX6b).
function bulletMarker(
  pPr: PoNode | undefined,
  level: number,
  counters: Array<number>,
): string | undefined {
  if (!pPr) return undefined;
  if (poChildren(pPr).some((c) => poIs(c, 'a:buNone'))) return undefined;
  const buChar = poChildren(pPr).find((c) => poIs(c, 'a:buChar'));
  if (buChar) return `${poAttr(buChar, 'char') ?? '•'}  `;
  const buAuto = poChildren(pPr).find((c) => poIs(c, 'a:buAutoNum'));
  if (!buAuto) return undefined;
  const type = poAttr(buAuto, 'type') ?? 'arabicPeriod';
  const startAt = poIntAttr(buAuto, 'startAt') ?? 1;
  const prev = counters[level];
  const n = (prev === undefined ? startAt - 1 : prev) + 1;
  counters[level] = n;
  counters.length = level + 1; // deeper levels restart
  return `${n}${autoNumSuffix(type)}  `;
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
function parseTable(tbl: PoNode, colors: ColorResolver): Table {
  const grid: Array<Pt> = [];
  const tblGrid = poChildren(tbl).find((c) => poIs(c, 'a:tblGrid'));
  if (tblGrid) {
    for (const col of poChildren(tblGrid)) {
      if (poIs(col, 'a:gridCol')) grid.push(emuToPt(poIntAttr(col, 'w') ?? 0));
    }
  }
  const rows: Array<TableRow> = [];
  for (const tr of poChildren(tbl)) {
    if (poIs(tr, 'a:tr')) rows.push(parseTableRow(tr, colors));
  }
  return { properties: {}, grid, rows };
}

function parseTableRow(tr: PoNode, colors: ColorResolver): TableRow {
  const h = poIntAttr(tr, 'h');
  const cells: Array<TableCell> = [];
  for (const tc of poChildren(tr)) {
    if (!poIs(tc, 'a:tc')) continue;
    // Horizontal-merge continuation cells are covered by the gridSpan origin —
    // drop them so the origin's colSpan carries the width (the FlowDoc model
    // omits placeholder cells for spanned columns).
    if (poAttr(tc, 'hMerge') === '1') continue;
    cells.push(parseTableCell(tc, colors));
  }
  return { properties: { ...(h !== undefined ? { height: emuToPt(h) } : {}) }, cells };
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
  return {
    properties: {
      ...(gridSpan !== undefined && gridSpan > 1 ? { colSpan: gridSpan } : {}),
      ...(merge ? { merge } : {}),
      ...(shadingHex ? { shading: { colorHex: shadingHex } } : {}),
    },
    content,
  };
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
