// The laid-out diagram as the DRAWING part PowerPoint would have cached
// (§21.4.4 `dsp:drawing`). Emitting the same XML a file with a drawing part
// carries means the whole path behind it — geometry, fill, the node's text and
// its font colour, pictures — is the one already tested against files that DO
// carry one, instead of a second renderer that would drift from it.

import type { DiagramColors } from '@/core/drawingml/diagram/colors';
import type { LaidNode } from '@/core/drawingml/diagram/layout-engine';
import type { PoNode } from '@/core/po-helpers';
import { poAttr, poChildren, poIs, poText } from '@/core/po-helpers';

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
      // §21.4.3.9 — a box whose shape hides its geometry is there for the room
      // it takes, not to be seen: the picture lists space their nodes with one,
      // and painting it drew a bar of accent colour across the top of each.
      const paint =
        n.hideGeom === true
          ? '<a:noFill/><a:ln><a:noFill/></a:ln>'
          : (colors?.fill(n.styleLbl, n.index) ??
              (colors?.knows(n.styleLbl) === true
                ? '<a:noFill/>'
                : '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>')) +
            // §21.4.5 — `linClrLst` is the box's outline, and for the follower
            // boxes it is the whole of them: `conFgAcc1` is `lt1` at 90% inside
            // an `accent1` line, so unlined it is white on white.
            (colors?.line(n.styleLbl, n.index) ?? '');
      const text = colors?.textFill(n.styleLbl, n.index);
      return (
        `<dsp:sp><dsp:spPr><a:xfrm>${off}</a:xfrm>` +
        `<a:prstGeom prst="${esc(n.shapeType)}"><a:avLst/></a:prstGeom>${paint}</dsp:spPr>` +
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
        const src = poChildren(r).find((c) => poIs(c, 'a:rPr'));
        // §21.4.3 — the layout states a point size (`primFontSz` for a node's
        // own label, `secFontSz` for its descendants'), and it is a large one:
        // a diagram's text fills its box rather than sitting at the deck's
        // default. But it is only where the size starts when the RUN does not
        // state one of its own, and the run usually does. The drawings
        // PowerPoint itself cached prove it: customGeo.pptx writes `3600` in its
        // data and PowerPoint kept 3600 where it fitted and wrote 2500 and 3100
        // where it did not, while the files whose data names no size at all are
        // the ones it filled with 36pt to 65pt. Imposing 65 on every run set a
        // 16pt label three times over.
        const stated = poAttr(src, 'sz');
        const sz =
          stated !== undefined
            ? ` sz="${esc(stated)}"`
            : node.fontSizePt === undefined
              ? ''
              : ` sz="${Math.round(node.fontSizePt * 100)}"`;
        // Bold and italic are the author's own emphasis in the same way —
        // smartart-missing-bullet sets `b="1"` on its heading and nothing else
        // in the file says so.
        const emph = (['b', 'i', 'u'] as const)
          .map((a) => {
            const v = poAttr(src, a);
            return v === undefined ? '' : ` ${a}="${esc(v)}"`;
          })
          .join('');
        const rPr =
          textFill === undefined
            ? `<a:rPr lang="en-US"${sz}${emph}/>`
            : `<a:rPr lang="en-US"${sz}${emph}>${textFill}</a:rPr>`;
        runs += `<a:r>${rPr}<a:t>${esc(text)}</a:t></a:r>`;
      }
    }
    if (runs !== '') paragraphs.push(`<a:p>${paraPr(node, p)}${runs}</a:p>`);
  }
  // §20.1.10.42 — a `normAutofit` carrying no scale of its own asks the layout,
  // which has the glyphs, to measure this text and shrink it until it fits.
  // Nothing here can: the stated 65pt of "Automatically shrinked text" stood
  // three lines deep and half a box wide outside the box it names.
  const anchor = node.anchor === undefined ? 'ctr' : esc(node.anchor);
  const body =
    `<a:bodyPr anchor="${anchor}"${node.bulleted === true ? ' anchorCtr="0"' : ''}>` +
    `<a:normAutofit/></a:bodyPr><a:lstStyle/>`;
  return paragraphs.length > 0 ? body + paragraphs.join('') : `${body}<a:p/>`;
}

/**
 * One paragraph's properties. `bulletEnabled` on a layout node says only that
 * bullets are ALLOWED here; which bullet a paragraph actually takes — or that it
 * takes none — is the data part's to state, and smartart-missing-bullet is a
 * file of exactly two paragraphs, one `a:buNone` and one `a:buChar`, that came
 * out with the same marker on both. So a paragraph that states its own bullet or
 * its own alignment keeps them, and only one that states neither falls back to
 * the layout's answer: ranged left behind a bullet where the layout allows one,
 * centred where it does not.
 */
function paraPr(node: LaidNode, para: PoNode): string {
  const src = poChildren(para).find((c) => poIs(c, 'a:pPr'));
  const kids = poChildren(src);
  const bullet = kids.find(
    (c) => poIs(c, 'a:buNone') || poIs(c, 'a:buChar') || poIs(c, 'a:buAutoNum'),
  );
  const listed = bullet === undefined ? node.bulleted === true : !poIs(bullet, 'a:buNone');
  const algn = poAttr(src, 'algn') ?? node.align ?? (listed ? 'l' : 'ctr');
  if (!listed) return `<a:pPr algn="${esc(algn)}"/>`;
  const marL = poAttr(src, 'marL') ?? '171450';
  const indent = poAttr(src, 'indent') ?? '-171450';
  const char =
    bullet !== undefined && poIs(bullet, 'a:buChar') ? poAttr(bullet, 'char') : undefined;
  return (
    `<a:pPr algn="${esc(algn)}" marL="${esc(marL)}" indent="${esc(indent)}">` +
    `<a:buChar char="${esc(char ?? '•')}"/></a:pPr>`
  );
}

const r = (v: number): number => Math.round(v);

function esc(s: string): string {
  return s
    .replace(/&/gu, '&amp;')
    .replace(/</gu, '&lt;')
    .replace(/>/gu, '&gt;')
    .replace(/"/gu, '&quot;');
}
