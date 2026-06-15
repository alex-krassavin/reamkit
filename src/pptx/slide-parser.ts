// E-PPTX PX1 — slide shapes → positioned FlowDoc elements.
//
// A PresentationML slide (p:cSld/p:spTree) is a canvas of shapes. Under Route A
// (epics.md) each shape becomes a floating element anchored at its EMU position,
// reusing the docx drawing model: a text-bearing p:sp with an explicit a:xfrm
// maps to a floating ShapeBlock whose ShapeTextBody carries the paragraphs.
//
// PX1 handles p:sp with their OWN transform and text. Placeholder shapes that
// inherit geometry from the slide layout/master (no a:xfrm) wait for PX2;
// pictures (p:pic), graphic frames (p:graphicFrame) and groups (p:grpSp) for
// PX3–PX5; bullets/levels/alignment/anchor/autofit for PX6.

import type {
  BodyElement,
  FloatAnchor,
  Paragraph,
  Run,
  RunProperties,
  ShapeBlock,
  ShapeGeometry,
  ShapeTextBody,
} from '@/core/document-model';
import type { PoNode } from '@/core/po-helpers';

import { emuToPt, pt } from '@/core/ir';
import { poAttr, poChildren, poIntAttr, poIs, poText } from '@/core/po-helpers';

const RECT_GEOMETRY: ShapeGeometry = { kind: 'preset', preset: 'rect', adjust: new Map() };

const isTrue = (v: string | undefined): boolean => v === '1' || v === 'true' || v === 'on';

// Walk p:cSld/p:spTree, turning each text-bearing p:sp with an explicit
// transform into a floating shape BodyElement positioned on the slide.
export function parseSlideShapes(spTree: PoNode): Array<BodyElement> {
  const out: Array<BodyElement> = [];
  for (const child of poChildren(spTree)) {
    if (!poIs(child, 'p:sp')) continue; // p:pic / p:graphicFrame / p:grpSp → later slices
    const shape = parseSp(child);
    if (shape) out.push({ kind: 'shape', shape });
  }
  return out;
}

// p:sp → a floating text box, or undefined when the shape has no own geometry
// (a layout-inherited placeholder — PX2) or no text (a graphic-only shape — PX3).
function parseSp(sp: PoNode): ShapeBlock | undefined {
  const spPr = poChildren(sp).find((c) => poIs(c, 'p:spPr'));
  const xfrm = spPr ? poChildren(spPr).find((c) => poIs(c, 'a:xfrm')) : undefined;
  if (!xfrm) return undefined;

  const ext = poChildren(xfrm).find((c) => poIs(c, 'a:ext'));
  const cx = ext ? poIntAttr(ext, 'cx') : undefined;
  const cy = ext ? poIntAttr(ext, 'cy') : undefined;
  if (cx === undefined || cy === undefined || cx <= 0 || cy <= 0) return undefined;

  const txBody = poChildren(sp).find((c) => poIs(c, 'p:txBody'));
  const text = txBody ? parseTxBody(txBody) : undefined;
  if (!text) return undefined;

  const off = poChildren(xfrm).find((c) => poIs(c, 'a:off'));
  const x = (off ? poIntAttr(off, 'x') : undefined) ?? 0;
  const y = (off ? poIntAttr(off, 'y') : undefined) ?? 0;
  // §20.4.2.3 — placement is page-absolute (the slide IS the page): off.x/off.y
  // from the page's top-left corner, sized by ext.
  const float: FloatAnchor = {
    wrap: 'none',
    posH: { relativeFrom: 'page', offsetPt: emuToPt(x) },
    posV: { relativeFrom: 'page', offsetPt: emuToPt(y) },
  };

  return {
    float,
    width: emuToPt(cx),
    height: emuToPt(cy),
    geometry: RECT_GEOMETRY,
    fill: { kind: 'none' },
    text,
    paragraphProperties: {},
  };
}

// p:txBody → ShapeTextBody. Each a:p becomes a paragraph; a:bodyPr insets are
// carried when present. Vertical anchor + autofit come in PX6.
function parseTxBody(txBody: PoNode): ShapeTextBody | undefined {
  const content: Array<BodyElement> = [];
  for (const child of poChildren(txBody)) {
    if (!poIs(child, 'a:p')) continue;
    content.push({ kind: 'paragraph', paragraph: parseSlideParagraph(child) });
  }
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

// a:p → Paragraph. Runs from a:r (and a:fld, a text field whose cached a:t is
// rendered as static text). a:br (soft line break) and a:pPr depth wait for PX6.
function parseSlideParagraph(aP: PoNode): Paragraph {
  const runs: Array<Run> = [];
  for (const child of poChildren(aP)) {
    if (poIs(child, 'a:r') || poIs(child, 'a:fld')) {
      const run = parseSlideRun(child);
      if (run) runs.push(run);
    }
  }
  return { properties: {}, runs };
}

// a:r / a:fld → Run. a:t holds the text; a:rPr the direct run formatting.
function parseSlideRun(node: PoNode): Run | undefined {
  const t = poChildren(node).find((c) => poIs(c, 'a:t'));
  const text = t ? poText(t) : '';
  if (text.length === 0) return undefined;
  const rPr = poChildren(node).find((c) => poIs(c, 'a:rPr'));
  return { text, properties: rPrToRunProps(rPr) };
}

// a:rPr → RunProperties. PX1 reads the direct character formatting: size
// (sz, hundredths of a point), bold/italic/underline, solid colour and the
// latin typeface. Scheme colours and the theme font cascade come in PX5.
function rPrToRunProps(rPr: PoNode | undefined): RunProperties {
  if (!rPr) return {};
  const sz = poIntAttr(rPr, 'sz');
  const u = poAttr(rPr, 'u');
  const colorHex = solidFillHex(rPr);
  const latin = poChildren(rPr).find((c) => poIs(c, 'a:latin'));
  const typeface = latin ? poAttr(latin, 'typeface')?.trim() : undefined;
  return {
    ...(isTrue(poAttr(rPr, 'b')) ? { bold: true } : {}),
    ...(isTrue(poAttr(rPr, 'i')) ? { italic: true } : {}),
    ...(u !== undefined && u !== 'none' ? { underline: 'single' as const } : {}),
    ...(sz !== undefined ? { fontSizePt: pt(sz / 100) } : {}),
    ...(colorHex ? { colorHex } : {}),
    ...(typeface ? { fontFamily: { ascii: typeface } } : {}),
  };
}

// a:solidFill/a:srgbClr @val → '#rrggbb'. Scheme colours (a:schemeClr) need the
// theme resolver and are deferred to PX5.
function solidFillHex(rPr: PoNode): string | undefined {
  const solidFill = poChildren(rPr).find((c) => poIs(c, 'a:solidFill'));
  if (!solidFill) return undefined;
  const srgb = poChildren(solidFill).find((c) => poIs(c, 'a:srgbClr'));
  const val = srgb ? poAttr(srgb, 'val') : undefined;
  return val ? `#${val}` : undefined;
}
