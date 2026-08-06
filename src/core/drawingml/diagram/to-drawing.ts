// The laid-out diagram as the DRAWING part PowerPoint would have cached
// (§21.4.4 `dsp:drawing`). Emitting the same XML a file with a drawing part
// carries means the whole path behind it — geometry, fill, the node's text and
// its font colour, pictures — is the one already tested against files that DO
// carry one, instead of a second renderer that would drift from it.

import type { DiagramColors } from '@/core/drawingml/diagram/colors';
import type { LaidNode } from '@/core/drawingml/diagram/layout-engine';
import type { PoNode } from '@/core/po-helpers';
import { poChildren, poIs, poText } from '@/core/po-helpers';

const NS =
  'xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/**
 * A `dsp:drawing` for boxes the engine laid out.
 *
 * @param nodes  The laid-out boxes, in the diagram's own EMU frame.
 * @param frame  The frame the slide gives the diagram.
 * @param colors The colours part, when the file carries one.
 * @returns The XML text of the drawing part.
 */
export function diagramDrawingXml(
  nodes: ReadonlyArray<LaidNode>,
  frame: { readonly cx: number; readonly cy: number },
  colors?: DiagramColors,
): string {
  const shapes = nodes
    .map((n) => {
      const off = `<a:off x="${r(n.x)}" y="${r(n.y)}"/><a:ext cx="${r(n.cx)}" cy="${r(n.cy)}"/>`;
      // §21.4.5 — the colours part names a RUN of accents per style label, so
      // sibling boxes differ; only a file without one falls back to a single
      // accent for the lot.
      const fill =
        colors?.fill(n.styleLbl, n.index) ??
        '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
      const text = colors?.textFill(n.styleLbl, n.index);
      return (
        `<dsp:sp><dsp:spPr><a:xfrm>${off}</a:xfrm>` +
        `<a:prstGeom prst="${esc(n.shapeType)}"><a:avLst/></a:prstGeom>${fill}</dsp:spPr>` +
        `<dsp:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>` +
        `<dsp:txBody>${bodyXml(n, text)}</dsp:txBody></dsp:sp>`
      );
    })
    .join('');
  // §21.4.4 — the group's own transform states the CHILD extent the boxes are
  // laid out in, and the reader maps that onto the graphic frame the slide
  // gives the diagram. Without it the boxes land at their raw EMU sizes inside
  // a frame that is a different size entirely, which is how five 80pt boxes
  // came out at half that and floating in the corner.
  const grp =
    `<dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${r(frame.cx)}" cy="${r(frame.cy)}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${r(frame.cx)}" cy="${r(frame.cy)}"/></a:xfrm></dsp:grpSpPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing ${NS}><dsp:spTree>${grp}${shapes}</dsp:spTree></dsp:drawing>`;
}

// The node's own text, re-emitted as the drawing part spells it. Only the runs'
// text is carried: the size and colour a diagram's text takes come from the
// style, not from the data part's runs.
function bodyXml(node: LaidNode, textFill?: string): string {
  const paragraphs: Array<string> = [];
  for (const p of node.point.text ? poChildren(node.point.text) : []) {
    if (!poIs(p, 'a:p')) continue;
    let runs = '';
    for (const r of poChildren(p)) {
      if (!poIs(r, 'a:r')) continue;
      const text = poText(poChildren(r).find((c) => poIs(c, 'a:t')));
      if (text !== '') {
        const rPr =
          textFill === undefined
            ? '<a:rPr lang="en-US"/>'
            : `<a:rPr lang="en-US">${textFill}</a:rPr>`;
        runs += `<a:r>${rPr}<a:t>${esc(text)}</a:t></a:r>`;
      }
    }
    if (runs !== '') paragraphs.push(`<a:p><a:pPr algn="ctr"/>${runs}</a:p>`);
  }
  const body = '<a:bodyPr anchor="ctr"/><a:lstStyle/>';
  return paragraphs.length > 0 ? body + paragraphs.join('') : `${body}<a:p/>`;
}

/** Whether a parsed drawing part actually holds shapes to draw. */
export function hasShapes(tree: ReadonlyArray<PoNode>): boolean {
  return tree.length > 0;
}

const r = (v: number): number => Math.round(v);

function esc(s: string): string {
  return s
    .replace(/&/gu, '&amp;')
    .replace(/</gu, '&lt;')
    .replace(/>/gu, '&gt;')
    .replace(/"/gu, '&quot;');
}
