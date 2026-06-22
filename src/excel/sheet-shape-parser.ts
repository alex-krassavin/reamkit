// Sheet shapes (E-SHEET W2). A worksheet drawing's xdr:sp anchors render as
// floating shapes — geometry / fill / line / text — reusing the DrawingML readers
// the pptx + docx (SmartArt) shape paths use. Those readers operate on the
// preserveOrder PoNode tree (prefixed a:/xdr: tags), whereas sheet-drawing.ts
// reads charts + pictures via a removeNSPrefix fast-xml tree; the two configs are
// incompatible, so shapes parse the drawing a second time via parseXml. Like
// SmartArt, sheet shapes carry no placeholder cascade, so runs use their direct
// a:rPr formatting. A shape's box comes from its sheet ANCHOR (from/to tracks),
// not its a:xfrm (which is usually relative/zero on a sheet).

import type { ShapeBlock, ShapeFill } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { ParsedWorksheet } from '@/core/spreadsheet-model';
import type { PoNode } from '@/core/po-helpers';

import { emuToPt, pt } from '@/core/ir';
import { poChildren, poFirstChild, poIntAttr, poIs, poText } from '@/core/po-helpers';
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
 */
export function parseSheetShapes(
  drawingXml: Uint8Array,
  worksheet: ParsedWorksheet,
  colors: ColorResolver,
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
    const box = anchorBox(anchor, colWidthPt, rowHeightPt);
    if (!box) continue;

    // xdr:spPr wraps the same a: children (a:xfrm/a:prstGeom/a:solidFill/a:ln) the
    // shared readers expect; xdr:txBody holds a:bodyPr + a:p like a slide shape.
    const spPr = poChildren(sp).find((c) => poIs(c, 'xdr:spPr'));
    const geometry = parseGeometry(spPr);
    const fill: ShapeFill = spPr ? parseFill(spPr, colors) : { kind: 'none' };
    const line = spPr ? parseLine(spPr, colors) : undefined;
    const txBody = poChildren(sp).find((c) => poIs(c, 'xdr:txBody'));
    const text = txBody ? parseTxBody(txBody, undefined, undefined, colors, undefined) : undefined;
    const visibleLine = line !== undefined && line.fill !== 'none';
    if (!text && fill.kind === 'none' && !visibleLine) continue;

    shapes.push({
      shape: {
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

interface AnchorBox {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly anchorRow: number;
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
    return widthPt > 0 && heightPt > 0 ? { widthPt, heightPt, anchorRow: from.row } : undefined;
  }
  const ext = poChildren(anchor).find((c) => poIs(c, 'xdr:ext'));
  if (!ext) return undefined;
  const widthPt = emuToPt(poIntAttr(ext, 'cx') ?? 0);
  const heightPt = emuToPt(poIntAttr(ext, 'cy') ?? 0);
  return widthPt > 0 && heightPt > 0 ? { widthPt, heightPt, anchorRow: from?.row ?? 0 } : undefined;
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
