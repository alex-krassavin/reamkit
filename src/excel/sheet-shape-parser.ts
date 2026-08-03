// Sheet shapes (E-SHEET W2). A worksheet drawing's xdr:sp anchors render as
// floating shapes — geometry / fill / line / text — reusing the DrawingML readers
// the pptx + docx (SmartArt) shape paths use. Those readers operate on the
// preserveOrder PoNode tree (prefixed a:/xdr: tags), whereas sheet-drawing.ts
// reads charts + pictures via a removeNSPrefix fast-xml tree; the two configs are
// incompatible, so shapes parse the drawing a second time via parseXml. Like
// SmartArt, sheet shapes carry no placeholder cascade, so runs use their direct
// a:rPr formatting. A shape's box comes from its sheet ANCHOR (from/to tracks),
// not its a:xfrm (which is usually relative/zero on a sheet).

import type {
  ShapeBlock,
  ShapeFill,
  ShapeLine,
  ShapeShadow,
  ShapeTextBody,
} from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { ParsedWorksheet } from '@/core/spreadsheet-model';
import type { PoNode } from '@/core/po-helpers';

import { emuToPt, pt } from '@/core/ir';
import { placeholderColors, resolveColorNode } from '@/core/drawingml/colors';
import {
  poAttr,
  poChildren,
  poFirstChild,
  poIntAttr,
  poIs,
  poTag,
  poText,
} from '@/core/po-helpers';
import { parseXml } from '@/pptx/pptx-reader';
import { parseGeometry, parseTxBody } from '@/pptx/slide-parser';
import {
  parseFill,
  parseLine,
  parseShadow,
  parseXfrm,
  shadowFromOuterShdw,
} from '@/word/drawing-parser';
import { makeColWidthPt, makeRowHeightPt } from '@/excel/sheet-drawing';

interface SheetShape {
  readonly shape: ShapeBlock;
  readonly anchorRow: number;
}

/**
 * The namespace a shape's own children carry. A sheet drawing writes them under
 * `xdr:`; the fallback drawing of a SmartArt diagram writes the SAME children
 * under `dsp:` (§ MS-ODRAWXML 2.1) — one reader serves both.
 */
type ShapeNs = 'xdr' | 'dsp';

const ANCHOR_KINDS = ['xdr:twoCellAnchor', 'xdr:oneCellAnchor', 'xdr:absoluteAnchor'] as const;

/**
 * One `xdr:sp` / `dsp:sp` as a {@link ShapeBlock} in the box it occupies, or
 * undefined when it would draw nothing at all.
 *
 * @param sp     The shape element.
 * @param box    The box it occupies on the page.
 * @param ns     Which namespace its own children carry.
 * @param colors The theme colour resolver.
 * @param themeLineWidths   `a:lnStyleLst` widths, indexed by `<a:lnRef idx>`.
 * @param themeFillStyles   `a:fillStyleLst` nodes, indexed by `<a:fillRef idx>`.
 * @param themeEffectStyles `a:effectStyleLst` nodes, indexed by `<a:effectRef idx>`.
 */
function buildShape(
  sp: PoNode,
  box: AnchorBox,
  ns: ShapeNs,
  colors: ColorResolver,
  themeLineWidths: ReadonlyArray<number>,
  themeFillStyles: ReadonlyArray<PoNode>,
  themeEffectStyles: ReadonlyArray<PoNode>,
): ShapeBlock | undefined {
  // xdr:spPr wraps the same a: children (a:xfrm/a:prstGeom/a:solidFill/a:ln) the
  // shared readers expect; xdr:txBody holds a:bodyPr + a:p like a slide shape.
  const spPr = poChildren(sp).find((c) => poIs(c, `${ns}:spPr`));
  const txBody = poChildren(sp).find((c) => poIs(c, `${ns}:txBody`));
  const geometry = parseGeometry(spPr);
  const parsed = txBody ? parseTxBody(txBody, undefined, undefined, colors, undefined) : undefined;
  // An `<xdr:txBody>` is written whether or not it says anything, so the
  // test below is for CHARACTERS.
  const lettered = (parsed?.content ?? []).some(
    (b) => b.kind === 'paragraph' && b.paragraph.runs.some((r) => r.text.trim().length > 0),
  );
  // §20.1.7.6 — the anchor gives the box, `a:xfrm` gives how the shape sits
  // in it. Read for the box alone, a shape that is turned lies down flat:
  // tdf135828_Shape_Rect.xlsx points an `upArrow` up and to the right with
  // `rot="4616172"` (76.9°), and we drew it squashed across the page.
  const xfrm = spPr ? poChildren(spPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  const transform = xfrm ? parseXfrm(xfrm) : undefined;
  const turned = transform !== undefined && Object.keys(transform).length > 0;
  // A turned shape is the one case where the anchor is NOT its box. Excel
  // spans from/to across what the shape covers once rotated, and keeps the
  // shape's own size in `a:ext` — 23pt × 156pt for the tall thin `upArrow`
  // of tdf135828_Shape_Rect.xlsx, whose anchor is 164pt × 30pt because that
  // is the ground its 76.9° sweep covers. Drawn at the anchor's size the
  // arrow is a wide stub, and rotating THAT gives a sliver on its side.
  // Unturned the two agree, and the anchor stays authoritative.
  //
  // A turned shape with TEXT is left on its anchor. We lay a shape's text in
  // the page's frame rather than the shape's, so a label written 20pt wide
  // and turned 90° into a 156pt band would wrap every word: bnc762542.xlsx
  // has three of them and its "Description 1" broke after "Description".
  const own =
    turned && xfrm && !lettered ? poChildren(xfrm).find((c) => poIs(c, 'a:ext')) : undefined;
  const ownW = own ? emuToPt(poIntAttr(own, 'cx') ?? 0) : 0;
  const ownH = own ? emuToPt(poIntAttr(own, 'cy') ?? 0) : 0;
  // The rotation turns about the centre, so that is what the two boxes share.
  const useOwn = ownW > 0 && ownH > 0;
  const widthPt = useOwn ? ownW : box.widthPt;
  const heightPt = useOwn ? ownH : box.heightPt;
  const xPt = useOwn ? box.xPt + (box.widthPt - ownW) / 2 : box.xPt;
  const yPt = useOwn ? box.yPt + (box.heightPt - ownH) / 2 : box.yPt;
  const spFill: ShapeFill = spPr ? parseFill(spPr, colors) : { kind: 'none' };
  const fill = spFill.kind === 'none' ? styleFill(sp, ns, colors, themeFillStyles) : spFill;
  // §20.1.4.2.19 `<xdr:style><a:lnRef>` is where a shape drawn from a gallery
  // style keeps its outline — spPr then carries no `a:ln` at all, and read
  // alone it says the shape has no border. shape-macro-ext-ref.xlsx's macro
  // button lost the blue rule both references draw around it.
  // §20.1.2.2.24 — an `<a:ln/>` with nothing in it declares nothing, and it is
  // not a black hairline: 47504.xlsx writes one beside an `<a:lnRef>` that
  // carries the real outline, and taking the empty element at its word fenced
  // a gallery shape in black where both references draw the theme's own thin
  // accent rule.
  const directLine = spPr ? parseLine(spPr, colors) : undefined;
  const line =
    directLine && Object.keys(directLine).length > 0
      ? directLine
      : styleLine(sp, ns, colors, themeLineWidths);

  // §20.1.4.2.14 `<xdr:style><a:fontRef>` carries the colour the shape's text
  // takes when its runs name none — for anything drawn from a gallery style
  // that is where the colour IS. shape-macro-ext-ref.xlsx asks for `lt1` on a
  // green button and we drew black on green, because a run with no `a:rPr`
  // colour fell through to the layout's default.
  const text = parsed ? withStyleTextColor(parsed, sp, ns, colors) : undefined;
  const shadow =
    (spPr ? parseShadow(spPr, colors) : undefined) ??
    styleShadow(sp, ns, colors, themeEffectStyles);
  const visibleLine = line !== undefined && line.fill !== 'none';
  if (!text && fill.kind === 'none' && !visibleLine) return undefined;

  return {
    // §20.5.2.35: a `twoCellAnchor` is a TWO-dimensional placement. Emitted
    // as a plain block the drawing kept only its size and its order down
    // the page — everything landed against the left margin, which turned
    // bnc762542.xlsx's callout (three swatches, three leader lines, three
    // labels, side by side) into a single vertical stack. The layout floats
    // a shape at its anchor when one is given, so give it one.
    float: {
      wrap: 'none' as const,
      posH: { relativeFrom: 'margin' as const, offsetPt: pt(xPt) },
      posV: { relativeFrom: 'margin' as const, offsetPt: pt(yPt) },
    },
    width: pt(widthPt),
    height: pt(heightPt),
    geometry,
    fill,
    ...(transform && Object.keys(transform).length > 0 ? { transform } : {}),
    ...(line ? { line } : {}),
    ...(text ? { text } : {}),
    ...(shadow ? { shadow } : {}),
    paragraphProperties: {},
  };
}

/**
 * Parse a drawing's `xdr:sp` shape anchors (§20.5.2.30) into anchor-ordered
 * {@link ShapeBlock}s (E-SHEET W2), reusing the shared DrawingML geometry / fill /
 * line / text readers. A shape's box comes from its sheet anchor (from/to tracks),
 * not its `a:xfrm`. Returns `[]` when the drawing has no shapes (chart/picture-only),
 * so the reader's gate keeps non-shape sheets off this second parse.
 *
 * @param drawingXml The drawing part bytes (re-parsed via `parseXml` for the
 *                   preserveOrder tree the shape readers expect).
 * @param worksheet  The host worksheet, for the column/row track geometry.
 * @param colors     The theme colour resolver threaded into fill/line parsing.
 * @param themeLineWidths The theme's `a:lnStyleLst` widths in points, which an
 *                   `<a:lnRef idx>` indexes for a gallery-styled outline.
 * @param themeFillStyles The theme's `a:fillStyleLst` nodes, which an
 *                   `<a:fillRef idx>` indexes for a gallery-styled fill.
 * @param themeEffectStyles The theme's `a:effectStyleLst` nodes, which an
 *                   `<a:effectRef idx>` indexes for a gallery-styled shadow.
 */
export function parseSheetShapes(
  drawingXml: Uint8Array,
  worksheet: ParsedWorksheet,
  colors: ColorResolver,
  themeLineWidths: ReadonlyArray<number> = [],
  themeFillStyles: ReadonlyArray<PoNode> = [],
  themeEffectStyles: ReadonlyArray<PoNode> = [],
): Array<ShapeBlock> {
  const tree = parseXml(drawingXml);
  const wsDr = tree.find((n) => poIs(n, 'xdr:wsDr'));
  if (!wsDr) return [];
  const colWidthPt = makeColWidthPt(worksheet);
  const rowHeightPt = makeRowHeightPt(worksheet);

  const shapes: Array<SheetShape> = [];
  for (const anchor of poChildren(wsDr)) {
    if (!ANCHOR_KINDS.some((k) => poIs(anchor, k))) continue;
    const anchored = anchorBox(anchor, colWidthPt, rowHeightPt);
    if (!anchored) continue;
    // §20.5.2.17 — an anchor may frame a GROUP rather than a shape, and the
    // walk looked for a direct `xdr:sp` only: groupShape.xlsx nests two groups
    // over three rectangles and we drew none of them.
    const placed: Array<{ sp: PoNode; box: AnchorBox }> = [];
    const direct = poChildren(anchor).find((c) => poIs(c, 'xdr:sp'));
    if (direct) placed.push({ sp: direct, box: anchored });
    for (const group of poChildren(anchor).filter((c) => poIs(c, 'xdr:grpSp'))) {
      collectGrouped(group, anchored, placed);
    }
    for (const { sp, box } of placed) {
      if (isHiddenDrawing(sp, 'xdr:nvSpPr')) continue;

      const shape = buildShape(
        sp,
        box,
        'xdr',
        colors,
        themeLineWidths,
        themeFillStyles,
        themeEffectStyles,
      );
      if (!shape) continue;
      shapes.push({ shape, anchorRow: box.anchorRow });
    }
  }
  shapes.sort((a, b) => a.anchorRow - b.anchorRow);
  return shapes.map((s) => s.shape);
}

/**
 * The boxes the diagram frames of a drawing occupy, in document order.
 *
 * §20.5.2.16 `xdr:graphicFrame` is how a sheet hosts a diagram: the frame is
 * anchored like any other drawing and the diagram itself lives in its own
 * parts. The frame gives the corner its fallback shapes are laid out from —
 * see {@link parseDiagramShapes}.
 *
 * @param drawingXml The drawing part bytes.
 * @param worksheet  The host worksheet, for the column/row track geometry.
 * @returns One box per diagram frame, in the order the drawing writes them.
 */
export function parseDiagramFrames(
  drawingXml: Uint8Array,
  worksheet: ParsedWorksheet,
): Array<AnchorBox> {
  const tree = parseXml(drawingXml);
  const wsDr = tree.find((n) => poIs(n, 'xdr:wsDr'));
  if (!wsDr) return [];
  const colWidthPt = makeColWidthPt(worksheet);
  const rowHeightPt = makeRowHeightPt(worksheet);
  const out: Array<AnchorBox> = [];
  for (const anchor of poChildren(wsDr)) {
    if (!ANCHOR_KINDS.some((k) => poIs(anchor, k))) continue;
    const frame = poChildren(anchor).find((c) => poIs(c, 'xdr:graphicFrame'));
    if (!frame) continue;
    const data = poFirstChild(poFirstChild(frame, 'a:graphic'), 'a:graphicData');
    if (!poFirstChild(data, 'dgm:relIds')) continue;
    const box = anchorBox(anchor, colWidthPt, rowHeightPt);
    if (box) out.push(box);
  }
  return out;
}

/**
 * Parse the fallback drawing of a SmartArt diagram (§ MS-ODRAWXML 2.1
 * `dsp:drawing`) into the shapes it lays out.
 *
 * A diagram is FOUR parts of description — data, layout, quick style, colours —
 * that a renderer is meant to lay out itself, and nobody outside Office does.
 * The producer therefore also writes what it drew: a plain DrawingML picture of
 * the result, under `dsp:`, reachable from the drawing part by a
 * `diagramDrawing` relationship. Reading it is the difference between a diagram
 * and a blank space — tdf83671_SmartArt_import.xlsx draws three nested circles
 * where we drew nothing between its "start" and its "end".
 *
 * The shapes' `a:xfrm` offsets are relative to the frame the diagram sits in,
 * so the frame's own corner is all that has to be added.
 *
 * @param diagramXml The `dsp:drawing` part bytes.
 * @param frame      The box the graphic frame occupies on the page.
 * @param colors     The theme colour resolver.
 * @param themeLineWidths   `a:lnStyleLst` widths, indexed by `<a:lnRef idx>`.
 * @param themeFillStyles   `a:fillStyleLst` nodes, indexed by `<a:fillRef idx>`.
 * @param themeEffectStyles `a:effectStyleLst` nodes, indexed by `<a:effectRef idx>`.
 * @returns The shapes, in the order the diagram stacks them.
 */
export function parseDiagramShapes(
  diagramXml: Uint8Array,
  frame: AnchorBox,
  colors: ColorResolver,
  themeLineWidths: ReadonlyArray<number> = [],
  themeFillStyles: ReadonlyArray<PoNode> = [],
  themeEffectStyles: ReadonlyArray<PoNode> = [],
): Array<ShapeBlock> {
  const tree = parseXml(diagramXml);
  const drawing = tree.find((n) => poIs(n, 'dsp:drawing'));
  const spTree = drawing ? poFirstChild(drawing, 'dsp:spTree') : undefined;
  if (!spTree) return [];
  const out: Array<ShapeBlock> = [];
  for (const sp of poChildren(spTree)) {
    if (!poIs(sp, 'dsp:sp')) continue;
    const box = diagramBox(sp, frame);
    if (!box) continue;
    const shape = buildShape(
      sp,
      box,
      'dsp',
      colors,
      themeLineWidths,
      themeFillStyles,
      themeEffectStyles,
    );
    if (!shape) continue;
    // A diagram gives the LABEL its own rectangle: `dsp:txXfrm` is where the
    // words go, which is not the middle of the shape they belong to. Left on
    // the shape, all three of tdf83671_SmartArt_import.xlsx's labels stacked in
    // the centre of its circles instead of sitting one per band. The graphic
    // and its label become two blocks, since a block has one box for both.
    const text = shape.text;
    const label = text ? textBox(sp, frame) : undefined;
    if (!label || !text) {
      out.push(shape);
      continue;
    }
    const { text: _moved, ...graphic } = shape;
    out.push(graphic);
    out.push({
      float: {
        wrap: 'none' as const,
        posH: { relativeFrom: 'margin' as const, offsetPt: pt(label.xPt) },
        posV: { relativeFrom: 'margin' as const, offsetPt: pt(label.yPt) },
      },
      width: pt(label.widthPt),
      height: pt(label.heightPt),
      geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
      fill: { kind: 'none' },
      text,
      paragraphProperties: {},
    });
  }
  return out;
}

/** A diagram label's box: its `dsp:txXfrm`, moved to the frame's corner. */
function textBox(sp: PoNode, frame: AnchorBox): AnchorBox | undefined {
  const xfrm = poFirstChild(sp, 'dsp:txXfrm');
  const ext = xfrm ? poFirstChild(xfrm, 'a:ext') : undefined;
  if (!ext) return undefined;
  const widthPt = emuToPt(poIntAttr(ext, 'cx') ?? 0);
  const heightPt = emuToPt(poIntAttr(ext, 'cy') ?? 0);
  if (!(widthPt > 0 && heightPt > 0)) return undefined;
  const off = poFirstChild(xfrm, 'a:off');
  return {
    widthPt,
    heightPt,
    anchorRow: frame.anchorRow,
    xPt: frame.xPt + emuToPt(poIntAttr(off, 'x') ?? 0),
    yPt: frame.yPt + emuToPt(poIntAttr(off, 'y') ?? 0),
  };
}

/** A diagram shape's box: its own `a:xfrm`, moved to the frame's corner. */
function diagramBox(sp: PoNode, frame: AnchorBox): AnchorBox | undefined {
  const spPr = poFirstChild(sp, 'dsp:spPr');
  const xfrm = spPr ? poFirstChild(spPr, 'a:xfrm') : undefined;
  const off = xfrm ? poFirstChild(xfrm, 'a:off') : undefined;
  const ext = xfrm ? poFirstChild(xfrm, 'a:ext') : undefined;
  if (!ext) return undefined;
  const widthPt = emuToPt(poIntAttr(ext, 'cx') ?? 0);
  const heightPt = emuToPt(poIntAttr(ext, 'cy') ?? 0);
  if (!(widthPt > 0 && heightPt > 0)) return undefined;
  return {
    widthPt,
    heightPt,
    anchorRow: frame.anchorRow,
    xPt: frame.xPt + emuToPt(poIntAttr(off, 'x') ?? 0),
    yPt: frame.yPt + emuToPt(poIntAttr(off, 'y') ?? 0),
  };
}

/**
 * Place a group's children inside the box the group itself occupies.
 *
 * §20.5.2.17 `xdr:grpSp` — a group gives its children a coordinate space of its
 * own: `a:xfrm/a:chOff` + `a:chExt` name that space, and each child's own
 * `a:xfrm` is a rectangle inside it. Mapping the rectangle onto the group's box
 * places the child; a nested group recurses through the same map.
 *
 * @param group The `xdr:grpSp` element.
 * @param box   The box the group occupies on the page.
 * @param out   Collects each `xdr:sp` with the box it lands in.
 */
function collectGrouped(
  group: PoNode,
  box: AnchorBox,
  out: Array<{ sp: PoNode; box: AnchorBox }>,
): void {
  const grpPr = poChildren(group).find((c) => poIs(c, 'xdr:grpSpPr'));
  const xfrm = grpPr ? poChildren(grpPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  const chOff = xfrm ? poChildren(xfrm).find((c) => poIs(c, 'a:chOff')) : undefined;
  const chExt = xfrm ? poChildren(xfrm).find((c) => poIs(c, 'a:chExt')) : undefined;
  const ox = chOff ? (poIntAttr(chOff, 'x') ?? 0) : 0;
  const oy = chOff ? (poIntAttr(chOff, 'y') ?? 0) : 0;
  const cw = chExt ? (poIntAttr(chExt, 'cx') ?? 0) : 0;
  const ch = chExt ? (poIntAttr(chExt, 'cy') ?? 0) : 0;
  if (!(cw > 0) || !(ch > 0)) return;
  const place = (node: PoNode, prTag: string): AnchorBox | undefined => {
    const pr = poChildren(node).find((c) => poIs(c, prTag));
    const x = pr ? poChildren(pr).find((c) => poIs(c, 'a:xfrm')) : undefined;
    const off = x ? poChildren(x).find((c) => poIs(c, 'a:off')) : undefined;
    const ext = x ? poChildren(x).find((c) => poIs(c, 'a:ext')) : undefined;
    if (!off || !ext) return undefined;
    const w = ((poIntAttr(ext, 'cx') ?? 0) / cw) * box.widthPt;
    const h = ((poIntAttr(ext, 'cy') ?? 0) / ch) * box.heightPt;
    if (!(w > 0) || !(h > 0)) return undefined;
    return {
      widthPt: w,
      heightPt: h,
      anchorRow: box.anchorRow,
      xPt: box.xPt + (((poIntAttr(off, 'x') ?? 0) - ox) / cw) * box.widthPt,
      yPt: box.yPt + (((poIntAttr(off, 'y') ?? 0) - oy) / ch) * box.heightPt,
    };
  };
  for (const child of poChildren(group)) {
    if (poIs(child, 'xdr:sp')) {
      const b = place(child, 'xdr:spPr');
      if (b) out.push({ sp: child, box: b });
    } else if (poIs(child, 'xdr:grpSp')) {
      const b = place(child, 'xdr:grpSpPr');
      if (b) collectGrouped(child, b, out);
    }
  }
}

/**
 * Whether a drawing says it is not to be shown.
 *
 * §20.1.2.2.8 `cNvPr@hidden` — "Specifies whether this DrawingML object shall
 * be displayed", default false. POI writes one white rectangle per cell comment
 * under the name `_xssf_cell_comment` and marks it hidden; read without the
 * flag, 51850.xlsx grew a 494 × 677pt outlined box across both its pages.
 *
 * @param node      The `xdr:sp` / `xdr:pic` / `xdr:graphicFrame` element.
 * @param nvPropTag Its non-visual properties wrapper.
 * @returns True when the object declares itself hidden.
 */
function isHiddenDrawing(node: PoNode, nvPropTag: string): boolean {
  const nv = poChildren(node).find((c) => poIs(c, nvPropTag));
  const cNvPr = nv ? poChildren(nv).find((c) => poIs(c, 'xdr:cNvPr')) : undefined;
  const hidden = cNvPr ? poAttr(cNvPr, 'hidden') : undefined;
  return hidden === '1' || hidden === 'true';
}

/**
 * The outline `<xdr:style><a:lnRef>` names, if it names one.
 *
 * §20.1.4.2.19: the reference's `idx` is a 1-based index into the theme's
 * `a:lnStyleLst`, which is where the WIDTH lives — the reference itself carries
 * only the colour. Assuming a hairline drew 50299.xlsx's `idx="2"` rectangle in
 * 0.75pt where its theme asks for 2pt.
 */
function styleLine(
  sp: PoNode,
  ns: ShapeNs,
  colors: ColorResolver,
  themeLineWidths: ReadonlyArray<number>,
): ShapeLine | undefined {
  const style = poChildren(sp).find((c) => poIs(c, `${ns}:style`));
  const lnRef = style ? poChildren(style).find((c) => poIs(c, 'a:lnRef')) : undefined;
  const child = poFirstElement(lnRef);
  const colorHex = child ? resolveColorNode(child, colors) : undefined;
  if (colorHex === undefined) return undefined;
  const idx = Number(poAttr(lnRef, 'idx') ?? '');
  const width = Number.isFinite(idx) ? themeLineWidths[idx - 1] : undefined;
  return { colorHex, width: pt(width ?? 0.75) };
}

/**
 * The fill `<xdr:style><a:fillRef>` names, if it names one.
 *
 * §20.1.4.2.10, the same mechanism as the outline beside it: a shape drawn from
 * a gallery style carries no `a:solidFill` in its `spPr` at all, and read alone
 * it says the shape has no fill. 50299.xlsx's rectangle asks for `accent1` that
 * way, and we drew it empty on all six sheets.
 *
 * `idx="0"` is the explicit "no fill" slot. Above it the reference is a 1-based
 * index into the theme's `a:fillStyleLst` (§20.1.4.1.13), and THAT is the fill —
 * the reference only says which colour goes where the style says `phClr`. The
 * standard Office theme's third slot is a gradient, which is what 47504.xlsx
 * asks for and what both references draw; painting the referenced colour flat
 * instead lost the gradient on every gallery shape.
 */
/**
 * The first CHILD ELEMENT of a style reference — `<a:lnRef>`, `<a:fillRef>`,
 * `<a:effectRef>`, `<a:fontRef>` — each of which wraps exactly one colour.
 *
 * Not `poChildren(ref)[0]`: a part saved with indentation puts a whitespace
 * text node there, and the colour behind it went unread. Every gallery shape in
 * tdf139763ShapeAnchor.xlsx lost its fill AND its outline that way and drew
 * nothing at all on a page LibreOffice fills with two blue arrows.
 *
 * @param ref The style reference, or undefined.
 * @returns Its first element child, or undefined.
 */
function poFirstElement(ref: PoNode | undefined): PoNode | undefined {
  return ref ? poChildren(ref).find((c) => poTag(c) !== undefined) : undefined;
}

function styleFill(
  sp: PoNode,
  ns: ShapeNs,
  colors: ColorResolver,
  themeFillStyles: ReadonlyArray<PoNode>,
): ShapeFill {
  const style = poChildren(sp).find((c) => poIs(c, `${ns}:style`));
  const fillRef = style ? poChildren(style).find((c) => poIs(c, 'a:fillRef')) : undefined;
  if (!fillRef || poAttr(fillRef, 'idx') === '0') return { kind: 'none' };
  const child = poFirstElement(fillRef);
  const colorHex = child ? resolveColorNode(child, colors) : undefined;
  if (colorHex === undefined) return { kind: 'none' };
  const idx = Number(poAttr(fillRef, 'idx') ?? '');
  const slot = Number.isFinite(idx) ? themeFillStyles[idx - 1] : undefined;
  if (!slot) return { kind: 'solid', colorHex };
  // The slot is a whole fill written in the theme's own vocabulary, so read it
  // with the shared reader — under a resolver where `phClr` is the colour the
  // reference names (§20.1.4.2.10).
  const themed = parseFill({ 'a:spPr': [slot] }, placeholderColors(colors, colorHex));
  return themed.kind === 'none' ? { kind: 'solid', colorHex } : themed;
}

/**
 * The shadow `<xdr:style><a:effectRef>` names, if it names one.
 *
 * §20.1.4.2.8, the same mechanism as the fill and the outline beside it: a
 * shape drawn from a gallery style carries no `a:effectLst`, only an index into
 * the theme's `a:effectStyleLst`, and the colour to put where the style says
 * `phClr`. 47504.xlsx asks for slot 2 that way — the soft shadow both
 * references draw under it, and which we drew not at all.
 *
 * @param sp                The `xdr:sp` node.
 * @param colors            The workbook's colour resolver.
 * @param themeEffectStyles The theme's `a:effectStyleLst` nodes.
 * @returns The shadow, or undefined when the shape names no effect style.
 */
function styleShadow(
  sp: PoNode,
  ns: ShapeNs,
  colors: ColorResolver,
  themeEffectStyles: ReadonlyArray<PoNode>,
): ShapeShadow | undefined {
  const style = poChildren(sp).find((c) => poIs(c, `${ns}:style`));
  const ref = style ? poChildren(style).find((c) => poIs(c, 'a:effectRef')) : undefined;
  if (!ref) return undefined;
  const idx = Number(poAttr(ref, 'idx') ?? '');
  if (!Number.isFinite(idx) || idx < 1) return undefined;
  const slot = themeEffectStyles[idx - 1];
  const list = slot ? poFirstChild(slot, 'a:effectLst') : undefined;
  const shdw = list ? poFirstChild(list, 'a:outerShdw') : undefined;
  if (!shdw) return undefined;
  const child = poFirstElement(ref);
  const phHex = child ? resolveColorNode(child, colors) : undefined;
  return shadowFromOuterShdw(shdw, phHex ? placeholderColors(colors, phHex) : colors);
}

/** The colour `<xdr:style><a:fontRef>` names, if it names one. */
function styleFontColor(sp: PoNode, ns: ShapeNs, colors: ColorResolver): string | undefined {
  const style = poChildren(sp).find((c) => poIs(c, `${ns}:style`));
  const fontRef = style ? poChildren(style).find((c) => poIs(c, 'a:fontRef')) : undefined;
  // The colour is whichever colour child it carries — srgbClr, schemeClr, …
  const child = poFirstElement(fontRef);
  return child ? resolveColorNode(child, colors) : undefined;
}

/** The text body with that colour filled in wherever a run declares none. */
function withStyleTextColor(
  text: ShapeTextBody,
  sp: PoNode,
  ns: ShapeNs,
  colors: ColorResolver,
): ShapeTextBody {
  const colorHex = styleFontColor(sp, ns, colors);
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

export interface AnchorBox {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly anchorRow: number;
  /** Distance from the sheet's top-left corner to the shape's, in points. */
  readonly xPt: number;
  readonly yPt: number;
}

// The shape's size from its anchor: full tracks in [from..to) plus the offset
// difference (twoCellAnchor), or the explicit ext (one-cell / absolute anchor).
function anchorBox(
  anchor: PoNode,
  colWidthPt: (col: number) => number,
  rowHeightPt: (row: number) => number,
): AnchorBox | undefined {
  const from = marker(poChildren(anchor).find((c) => poIs(c, 'xdr:from')));
  if (poIs(anchor, 'xdr:twoCellAnchor')) {
    const to = marker(poChildren(anchor).find((c) => poIs(c, 'xdr:to')));
    if (!from || !to) return undefined;
    const widthPt = span(from.col, from.colOffPt, to.col, to.colOffPt, colWidthPt);
    const heightPt = span(from.row, from.rowOffPt, to.row, to.rowOffPt, rowHeightPt);
    if (!(widthPt > 0 && heightPt > 0)) return undefined;
    return {
      widthPt,
      heightPt,
      anchorRow: from.row,
      xPt: origin(from.col, from.colOffPt, colWidthPt),
      yPt: origin(from.row, from.rowOffPt, rowHeightPt),
    };
  }
  const ext = poChildren(anchor).find((c) => poIs(c, 'xdr:ext'));
  if (!ext) return undefined;
  const pos = absolutePos(anchor);
  const widthPt = emuToPt(poIntAttr(ext, 'cx') ?? 0);
  const heightPt = emuToPt(poIntAttr(ext, 'cy') ?? 0);
  if (!(widthPt > 0 && heightPt > 0)) return undefined;
  return {
    widthPt,
    heightPt,
    anchorRow: from?.row ?? 0,
    xPt: from ? origin(from.col, from.colOffPt, colWidthPt) : pos.xPt,
    yPt: from ? origin(from.row, from.rowOffPt, rowHeightPt) : pos.yPt,
  };
}

/**
 * §20.5.2.1 `absoluteAnchor` names no cell: it gives the distance from the
 * sheet's own corner in `<xdr:pos>`. Read nowhere, both of
 * tdf139763ShapeAnchor.xlsx's arrows piled into the top-left corner.
 *
 * @param anchor The anchor element.
 * @returns The offset in points; zero for an anchor that states none.
 */
function absolutePos(anchor: PoNode): { readonly xPt: number; readonly yPt: number } {
  const node = poChildren(anchor).find((c) => poIs(c, 'xdr:pos'));
  if (!node) return { xPt: 0, yPt: 0 };
  return { xPt: emuToPt(poIntAttr(node, 'x') ?? 0), yPt: emuToPt(poIntAttr(node, 'y') ?? 0) };
}

/**
 * The distance from the sheet's origin to a marker: every whole track before it
 * plus the marker's own offset into the one it lands in.
 */
function origin(index: number, offsetPt: number, trackPt: (i: number) => number): number {
  let total = 0;
  for (let i = 0; i < index; i++) total += trackPt(i);
  return total + offsetPt;
}

interface Marker {
  readonly col: number;
  readonly colOffPt: number;
  readonly row: number;
  readonly rowOffPt: number;
}

// §20.5.2.3 xdr:from/to — col/colOff(EMU)/row/rowOff(EMU) as element text.
function marker(node: PoNode | undefined): Marker | undefined {
  if (!node) return undefined;
  const col = childInt(node, 'xdr:col');
  const row = childInt(node, 'xdr:row');
  if (col === undefined || row === undefined) return undefined;
  return {
    col,
    row,
    colOffPt: emuToPt(childInt(node, 'xdr:colOff') ?? 0),
    rowOffPt: emuToPt(childInt(node, 'xdr:rowOff') ?? 0),
  };
}

function childInt(parent: PoNode, tag: string): number | undefined {
  const child = poFirstChild(parent, tag);
  if (!child) return undefined;
  const n = Number(poText(child));
  return Number.isFinite(n) ? n : undefined;
}

function span(
  from: number,
  fromOffPt: number,
  to: number,
  toOffPt: number,
  trackPt: (index: number) => number,
): number {
  let total = 0;
  for (let i = from; i < to; i++) total += trackPt(i);
  return total - fromOffPt + toOffPt;
}
