// Sheet shapes (E-SHEET W2). A worksheet drawing's xdr:sp anchors render as
// floating shapes — geometry / fill / line / text — reusing the DrawingML readers
// the pptx + docx (SmartArt) shape paths use. Those readers operate on the
// preserveOrder PoNode tree (prefixed a:/xdr: tags), whereas sheet-drawing.ts
// reads charts + pictures via a removeNSPrefix fast-xml tree; the two configs are
// incompatible, so shapes parse the drawing a second time via parseXml. Like
// SmartArt, sheet shapes carry no placeholder cascade, so runs use their direct
// a:rPr formatting. A shape's box comes from its sheet ANCHOR (from/to tracks),
// not its a:xfrm (which is usually relative/zero on a sheet).

import type { ShapeBlock, ShapeFill, ShapeLine, ShapeTextBody } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { ParsedWorksheet } from '@/core/spreadsheet-model';
import type { PoNode } from '@/core/po-helpers';

import { emuToPt, pt } from '@/core/ir';
import { applyColorMods, resolveColorNode } from '@/core/drawingml/colors';
import { poAttr, poChildren, poFirstChild, poIntAttr, poIs, poText } from '@/core/po-helpers';
import { parseXml } from '@/pptx/pptx-reader';
import { parseGeometry, parseTxBody } from '@/pptx/slide-parser';
import { parseFill, parseLine } from '@/word/drawing-parser';
import { makeColWidthPt, makeRowHeightPt } from '@/excel/sheet-drawing';

interface SheetShape {
  readonly shape: ShapeBlock;
  readonly anchorRow: number;
}

const ANCHOR_KINDS = ['xdr:twoCellAnchor', 'xdr:oneCellAnchor', 'xdr:absoluteAnchor'] as const;

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
 */
export function parseSheetShapes(
  drawingXml: Uint8Array,
  worksheet: ParsedWorksheet,
  colors: ColorResolver,
  themeLineWidths: ReadonlyArray<number> = [],
  themeFillStyles: ReadonlyArray<PoNode> = [],
): Array<ShapeBlock> {
  const tree = parseXml(drawingXml);
  const wsDr = tree.find((n) => poIs(n, 'xdr:wsDr'));
  if (!wsDr) return [];
  const colWidthPt = makeColWidthPt(worksheet);
  const rowHeightPt = makeRowHeightPt(worksheet);

  const shapes: Array<SheetShape> = [];
  for (const anchor of poChildren(wsDr)) {
    if (!ANCHOR_KINDS.some((k) => poIs(anchor, k))) continue;
    const sp = poChildren(anchor).find((c) => poIs(c, 'xdr:sp'));
    if (!sp) continue;
    if (isHiddenDrawing(sp, 'xdr:nvSpPr')) continue;
    const box = anchorBox(anchor, colWidthPt, rowHeightPt);
    if (!box) continue;

    // xdr:spPr wraps the same a: children (a:xfrm/a:prstGeom/a:solidFill/a:ln) the
    // shared readers expect; xdr:txBody holds a:bodyPr + a:p like a slide shape.
    const spPr = poChildren(sp).find((c) => poIs(c, 'xdr:spPr'));
    const geometry = parseGeometry(spPr);
    const spFill: ShapeFill = spPr ? parseFill(spPr, colors) : { kind: 'none' };
    const fill = spFill.kind === 'none' ? styleFill(sp, colors, themeFillStyles) : spFill;
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
        : styleLine(sp, colors, themeLineWidths);
    const txBody = poChildren(sp).find((c) => poIs(c, 'xdr:txBody'));
    const parsed = txBody
      ? parseTxBody(txBody, undefined, undefined, colors, undefined)
      : undefined;
    // §20.1.4.2.14 `<xdr:style><a:fontRef>` carries the colour the shape's text
    // takes when its runs name none — for anything drawn from a gallery style
    // that is where the colour IS. shape-macro-ext-ref.xlsx asks for `lt1` on a
    // green button and we drew black on green, because a run with no `a:rPr`
    // colour fell through to the layout's default.
    const text = parsed ? withStyleTextColor(parsed, sp, colors) : undefined;
    const visibleLine = line !== undefined && line.fill !== 'none';
    if (!text && fill.kind === 'none' && !visibleLine) continue;

    shapes.push({
      shape: {
        // §20.5.2.35: a `twoCellAnchor` is a TWO-dimensional placement. Emitted
        // as a plain block the drawing kept only its size and its order down
        // the page — everything landed against the left margin, which turned
        // bnc762542.xlsx's callout (three swatches, three leader lines, three
        // labels, side by side) into a single vertical stack. The layout floats
        // a shape at its anchor when one is given, so give it one.
        float: {
          wrap: 'none' as const,
          posH: { relativeFrom: 'margin' as const, offsetPt: pt(box.xPt) },
          posV: { relativeFrom: 'margin' as const, offsetPt: pt(box.yPt) },
        },
        width: pt(box.widthPt),
        height: pt(box.heightPt),
        geometry,
        fill,
        ...(line ? { line } : {}),
        ...(text ? { text } : {}),
        paragraphProperties: {},
      },
      anchorRow: box.anchorRow,
    });
  }
  shapes.sort((a, b) => a.anchorRow - b.anchorRow);
  return shapes.map((s) => s.shape);
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
  colors: ColorResolver,
  themeLineWidths: ReadonlyArray<number>,
): ShapeLine | undefined {
  const style = poChildren(sp).find((c) => poIs(c, 'xdr:style'));
  const lnRef = style ? poChildren(style).find((c) => poIs(c, 'a:lnRef')) : undefined;
  const child = lnRef ? poChildren(lnRef)[0] : undefined;
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
function styleFill(
  sp: PoNode,
  colors: ColorResolver,
  themeFillStyles: ReadonlyArray<PoNode>,
): ShapeFill {
  const style = poChildren(sp).find((c) => poIs(c, 'xdr:style'));
  const fillRef = style ? poChildren(style).find((c) => poIs(c, 'a:fillRef')) : undefined;
  if (!fillRef || poAttr(fillRef, 'idx') === '0') return { kind: 'none' };
  const child = poChildren(fillRef)[0];
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
 * §20.1.4.2.10 — inside a theme's style list, `phClr` stands for "the colour the
 * reference names". Every other colour resolves as it always did.
 *
 * @param colors The workbook's own resolver.
 * @param phHex  The colour the `a:fillRef` / `a:lnRef` carries.
 * @returns A resolver that answers `phClr` with that colour, transforms and all.
 */
function placeholderColors(colors: ColorResolver, phHex: string): ColorResolver {
  return (raw) =>
    'scheme' in raw && raw.scheme === 'phClr'
      ? applyColorMods(phHex, raw.mods ?? [])
      : colors(raw);
}

/** The colour `<xdr:style><a:fontRef>` names, if it names one. */
function styleFontColor(sp: PoNode, colors: ColorResolver): string | undefined {
  const style = poChildren(sp).find((c) => poIs(c, 'xdr:style'));
  const fontRef = style ? poChildren(style).find((c) => poIs(c, 'a:fontRef')) : undefined;
  // The colour is whichever colour child it carries — srgbClr, schemeClr, …
  const child = fontRef ? poChildren(fontRef)[0] : undefined;
  return child ? resolveColorNode(child, colors) : undefined;
}

/** The text body with that colour filled in wherever a run declares none. */
function withStyleTextColor(text: ShapeTextBody, sp: PoNode, colors: ColorResolver): ShapeTextBody {
  const colorHex = styleFontColor(sp, colors);
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

interface AnchorBox {
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
  const widthPt = emuToPt(poIntAttr(ext, 'cx') ?? 0);
  const heightPt = emuToPt(poIntAttr(ext, 'cy') ?? 0);
  if (!(widthPt > 0 && heightPt > 0)) return undefined;
  return {
    widthPt,
    heightPt,
    anchorRow: from?.row ?? 0,
    xPt: from ? origin(from.col, from.colOffPt, colWidthPt) : 0,
    yPt: from ? origin(from.row, from.rowOffPt, rowHeightPt) : 0,
  };
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
