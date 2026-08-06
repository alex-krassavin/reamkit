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
      // §21.1.2.2.1 `a:br` — the break the author typed inside a paragraph.
      // Dropped, "Max size" and "(65 pt)" ran together and the box broke the
      // result where it fell instead: "Max size(65 / pt)".
      if (poIs(r, 'a:br')) {
        runs += '<a:br/>';
        continue;
      }
      if (!poIs(r, 'a:r')) continue;
      const text = poText(poChildren(r).find((c) => poIs(c, 'a:t')));
      if (text !== '') {
        // §21.4.3 — the layout states the point size (`primFontSz` for a node's
        // own label, `secFontSz` for its descendants'), and it is a large one:
        // a diagram's text fills its box rather than sitting at the deck's
        // default, which is why every box read small. It is a MAXIMUM — the
        // `normAutofit` below is what brings it down to what fits.
        const sz =
          node.fontSizePt === undefined ? '' : ` sz="${Math.round(node.fontSizePt * 100)}"`;
        const rPr =
          textFill === undefined
            ? `<a:rPr lang="en-US"${sz}/>`
            : `<a:rPr lang="en-US"${sz}>${textFill}</a:rPr>`;
        runs += `<a:r>${rPr}<a:t>${esc(text)}</a:t></a:r>`;
      }
    }
    // Bulleted text reads as the list it is: ranged left behind its bullets,
    // where a node's own label sits centred in its box.
    const pPr =
      node.bulleted === true
        ? '<a:pPr algn="l" marL="171450" indent="-171450"><a:buChar char="\u2022"/></a:pPr>'
        : '<a:pPr algn="ctr"/>';
    if (runs !== '') paragraphs.push(`<a:p>${pPr}${runs}</a:p>`);
  }
  // §20.1.10.42 — a `normAutofit` carrying no scale of its own asks the layout,
  // which has the glyphs, to measure this text and shrink it until it fits.
  // Nothing here can: the stated 65pt of "Automatically shrinked text" stood
  // three lines deep and half a box wide outside the box it names.
  const body =
    `<a:bodyPr anchor="ctr"${node.bulleted === true ? ' anchorCtr="0"' : ''}>` +
    `<a:normAutofit/></a:bodyPr><a:lstStyle/>`;
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
