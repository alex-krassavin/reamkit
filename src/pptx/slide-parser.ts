// E-PPTX PX1–PX4 — slide shapes → positioned FlowDoc elements.
//
// A PresentationML slide (p:cSld/p:spTree) is a canvas of shapes. Under Route A
// (epics.md) each shape becomes a floating element anchored at its EMU position,
// reusing the docx drawing model and DrawingML readers.
//   * PX1/PX2 — p:sp text boxes: own a:xfrm + direct a:rPr, else the placeholder
//     cascade (geometry + master p:txStyles) for a p:ph with no own transform.
//   * PX3 — p:pic floating images; p:sp visible geometry/fill/stroke (p:spPr).
//   * PX4 — p:graphicFrame: a c:chart floating ChartBlock or an a:tbl Table.
// Groups (p:grpSp) wait for PX5; bullets/levels/alignment/anchor/autofit for PX6.

import type {
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
}

const RECT_GEOMETRY: ShapeGeometry = { kind: 'preset', preset: 'rect', adjust: new Map() };

// Walk p:cSld/p:spTree, turning each text-bearing p:sp into a floating text box
// and each p:pic into a floating image, positioned on the slide. The context
// supplies the placeholder cascade and the image resolver.
export function parseSlideShapes(spTree: PoNode, ctx: SlideContext = {}): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  for (const child of poChildren(spTree)) {
    if (poIs(child, 'p:sp')) {
      const shape = parseSp(child, ctx);
      if (shape) out.push({ kind: 'shape', shape });
    } else if (poIs(child, 'p:pic')) {
      const image = parsePic(child, ctx);
      if (image) out.push({ kind: 'image', image });
    } else if (poIs(child, 'p:graphicFrame')) {
      const el = parseGraphicFrame(child, ctx);
      if (el) out.push(el);
    }
    // p:grpSp (groups) → later slice
  }
  return out;
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
function parseSp(sp: PoNode, ctx: SlideContext): ShapeBlock | undefined {
  const ph = parsePh(sp);
  const colors = ctx.colors ?? defaultColorResolver;
  const spPr = poChildren(sp).find((c) => poIs(c, 'p:spPr'));
  let box: ShapeBoxEmu | undefined = parseXfrmBox(spPr);
  if (!box && ph && ctx.cascade) box = ctx.cascade.geometryFor(ph);
  if (!box) return undefined;

  const txBody = poChildren(sp).find((c) => poIs(c, 'p:txBody'));
  const text = txBody ? parseTxBody(txBody, ph, ctx.cascade, colors) : undefined;

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

// p:graphicFrame → a floating chart (c:chart, PX4a) or table (a:tbl, PX4b). The
// frame's transform is p:xfrm (a:off + a:ext), not the a:xfrm of a shape.
function parseGraphicFrame(gf: PoNode, ctx: SlideContext): BodyElement | undefined {
  const box = boxFromXfrm(poChildren(gf).find((c) => poIs(c, 'p:xfrm')));
  if (!box) return undefined;
  const graphicData = poFindDescendant(gf, 'a:graphicData');
  const uri = graphicData ? poAttr(graphicData, 'uri') : undefined;

  if (uri === CHART_URI) {
    const cChart = poFindDescendant(gf, 'c:chart');
    const relId = cChart ? poAttr(cChart, 'id') : undefined; // r:id
    const key = relId !== undefined ? ctx.resolveChart?.(relId) : undefined;
    if (key === undefined) return undefined;
    const chart: ChartBlock = {
      float: floatAt(box),
      chartRelId: key,
      width: emuToPt(box.cx),
      height: emuToPt(box.cy),
      paragraphProperties: {},
    };
    return { kind: 'chart', chart };
  }

  if (uri === TABLE_URI) {
    // A FlowDoc Table has no float, so a slide table flows in-block (the reader
    // zeroes the slide margins so it sits at the top-left). Its exact frame
    // position is a later refinement.
    const tbl = poFindDescendant(gf, 'a:tbl');
    return tbl
      ? { kind: 'table', table: parseTable(tbl, ctx.colors ?? defaultColorResolver) }
      : undefined;
  }
  return undefined;
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
function parsePic(pic: PoNode, ctx: SlideContext): ImageBlock | undefined {
  const spPr = poChildren(pic).find((c) => poIs(c, 'p:spPr'));
  const box = parseXfrmBox(spPr);
  if (!box) return undefined;

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
): Array<BodyElement> {
  const content: Array<BodyElement> = [];
  for (const child of poChildren(txBody)) {
    if (!poIs(child, 'a:p')) continue;
    content.push({ kind: 'paragraph', paragraph: parseSlideParagraph(child, ph, cascade, colors) });
  }
  return content;
}

// p:txBody → ShapeTextBody. a:bodyPr insets are carried when present; the
// vertical anchor + autofit come in PX6.
function parseTxBody(
  txBody: PoNode,
  ph: PlaceholderRef | undefined,
  cascade: PlaceholderCascade | undefined,
  colors: ColorResolver,
): ShapeTextBody | undefined {
  const content = txBodyParagraphs(txBody, ph, cascade, colors);
  if (content.length === 0) return undefined;

  const bodyPr = poChildren(txBody).find((c) => poIs(c, 'a:bodyPr'));
  const lIns = bodyPr ? poIntAttr(bodyPr, 'lIns') : undefined;
  const tIns = bodyPr ? poIntAttr(bodyPr, 'tIns') : undefined;
  const rIns = bodyPr ? poIntAttr(bodyPr, 'rIns') : undefined;
  const bIns = bodyPr ? poIntAttr(bodyPr, 'bIns') : undefined;
  return {
    content,
    ...(lIns !== undefined ? { insetLeft: emuToPt(lIns) } : {}),
    ...(tIns !== undefined ? { insetTop: emuToPt(tIns) } : {}),
    ...(rIns !== undefined ? { insetRight: emuToPt(rIns) } : {}),
    ...(bIns !== undefined ? { insetBottom: emuToPt(bIns) } : {}),
  };
}

// a:p → Paragraph. The paragraph's outline level (a:pPr @lvl, default 0) selects
// the placeholder's default run formatting, applied under each run's own a:rPr.
// Runs come from a:r and a:fld (a text field whose cached a:t renders as text).
function parseSlideParagraph(
  aP: PoNode,
  ph: PlaceholderRef | undefined,
  cascade: PlaceholderCascade | undefined,
  colors: ColorResolver,
): Paragraph {
  const pPr = poChildren(aP).find((c) => poIs(c, 'a:pPr'));
  const level = (pPr ? poIntAttr(pPr, 'lvl') : undefined) ?? 0;
  const defaults: RunProperties = ph && cascade ? cascade.defaultsFor(ph, level) : {};

  const runs: Array<Run> = [];
  for (const child of poChildren(aP)) {
    if (poIs(child, 'a:r') || poIs(child, 'a:fld')) {
      const run = parseSlideRun(child, defaults, colors);
      if (run) runs.push(run);
    }
  }
  return { properties: {}, runs };
}

// a:r / a:fld → Run. The placeholder defaults sit under the run's own a:rPr, so
// direct formatting always wins.
function parseSlideRun(
  node: PoNode,
  defaults: RunProperties,
  colors: ColorResolver,
): Run | undefined {
  const t = poChildren(node).find((c) => poIs(c, 'a:t'));
  const text = t ? poText(t) : '';
  if (text.length === 0) return undefined;
  const rPr = poChildren(node).find((c) => poIs(c, 'a:rPr'));
  return { text, properties: { ...defaults, ...rPrToRunProps(rPr, colors) } };
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
  const content = txBody ? txBodyParagraphs(txBody, undefined, undefined, colors) : [];
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
