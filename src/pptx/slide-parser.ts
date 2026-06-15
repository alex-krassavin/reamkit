// E-PPTX PX1/PX2 — slide shapes → positioned FlowDoc elements.
//
// A PresentationML slide (p:cSld/p:spTree) is a canvas of shapes. Under Route A
// (epics.md) each shape becomes a floating element anchored at its EMU position,
// reusing the docx drawing model: a text-bearing p:sp maps to a floating
// ShapeBlock whose ShapeTextBody carries the paragraphs.
//
// PX1 handles shapes with their OWN transform and direct run formatting. PX2
// adds the placeholder cascade: a p:sp that is a placeholder (p:ph) without its
// own a:xfrm inherits geometry from the slide layout/master, and its text
// inherits per-level size/colour defaults from the master's p:txStyles — so the
// title/body/number placeholders of a real deck land in place and at size.
// Pictures (p:pic), graphic frames (p:graphicFrame) and groups (p:grpSp) wait
// for PX3–PX5; bullets/levels/alignment/anchor/autofit for PX6.

import type {
  BodyElement,
  FloatAnchor,
  ImageBlock,
  Paragraph,
  Run,
  RunProperties,
  ShapeBlock,
  ShapeGeometry,
  ShapeTextBody,
} from '@/core/document-model';
import type { ResourceId } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import type { PlaceholderCascade } from '@/pptx/placeholder-cascade';
import type { PlaceholderRef, ShapeBoxEmu } from '@/pptx/sp-helpers';

import { emuToPt } from '@/core/ir';
import { poAttr, poChildren, poIntAttr, poIs, poText } from '@/core/po-helpers';
import { parsePh, parseXfrmBox, rPrToRunProps } from '@/pptx/sp-helpers';

// Per-slide parsing context: the placeholder cascade (PX2) and an image
// resolver that turns a slide-scoped relationship id (a:blip @r:embed) into a
// ResourceStore id (PX3). Both are optional — a bare slide needs neither.
export interface SlideContext {
  readonly cascade?: PlaceholderCascade;
  readonly resolveImage?: (relId: string) => ResourceId | undefined;
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
    }
    // p:graphicFrame (tables/charts) / p:grpSp (groups) → later slices
  }
  return out;
}

// p:sp → a floating text box. The box comes from the shape's own a:xfrm, else
// (for a placeholder) the cascade; undefined when neither supplies geometry or
// the shape has no text.
function parseSp(sp: PoNode, ctx: SlideContext): ShapeBlock | undefined {
  const ph = parsePh(sp);
  const spPr = poChildren(sp).find((c) => poIs(c, 'p:spPr'));
  let box: ShapeBoxEmu | undefined = parseXfrmBox(spPr);
  if (!box && ph && ctx.cascade) box = ctx.cascade.geometryFor(ph);
  if (!box) return undefined;

  const txBody = poChildren(sp).find((c) => poIs(c, 'p:txBody'));
  const text = txBody ? parseTxBody(txBody, ph, ctx.cascade) : undefined;
  if (!text) return undefined;

  // §20.4.2.3 — placement is page-absolute (the slide IS the page): off.x/off.y
  // from the page's top-left corner, sized by ext.
  const float: FloatAnchor = {
    wrap: 'none',
    posH: { relativeFrom: 'page', offsetPt: emuToPt(box.x) },
    posV: { relativeFrom: 'page', offsetPt: emuToPt(box.y) },
  };

  return {
    float,
    width: emuToPt(box.cx),
    height: emuToPt(box.cy),
    geometry: RECT_GEOMETRY,
    fill: { kind: 'none' },
    text,
    paragraphProperties: {},
  };
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
  const float: FloatAnchor = {
    wrap: 'none',
    posH: { relativeFrom: 'page', offsetPt: emuToPt(box.x) },
    posV: { relativeFrom: 'page', offsetPt: emuToPt(box.y) },
  };
  return {
    float,
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

// p:txBody → ShapeTextBody. Each a:p becomes a paragraph whose runs inherit the
// placeholder's per-level defaults (PX2) under their own a:rPr. a:bodyPr insets
// are carried when present; vertical anchor + autofit come in PX6.
function parseTxBody(
  txBody: PoNode,
  ph: PlaceholderRef | undefined,
  cascade?: PlaceholderCascade,
): ShapeTextBody | undefined {
  const content: Array<BodyElement> = [];
  for (const child of poChildren(txBody)) {
    if (!poIs(child, 'a:p')) continue;
    content.push({ kind: 'paragraph', paragraph: parseSlideParagraph(child, ph, cascade) });
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

// a:p → Paragraph. The paragraph's outline level (a:pPr @lvl, default 0) selects
// the placeholder's default run formatting, applied under each run's own a:rPr.
// Runs come from a:r and a:fld (a text field whose cached a:t renders as text).
function parseSlideParagraph(
  aP: PoNode,
  ph: PlaceholderRef | undefined,
  cascade?: PlaceholderCascade,
): Paragraph {
  const pPr = poChildren(aP).find((c) => poIs(c, 'a:pPr'));
  const level = (pPr ? poIntAttr(pPr, 'lvl') : undefined) ?? 0;
  const defaults: RunProperties = ph && cascade ? cascade.defaultsFor(ph, level) : {};

  const runs: Array<Run> = [];
  for (const child of poChildren(aP)) {
    if (poIs(child, 'a:r') || poIs(child, 'a:fld')) {
      const run = parseSlideRun(child, defaults);
      if (run) runs.push(run);
    }
  }
  return { properties: {}, runs };
}

// a:r / a:fld → Run. The placeholder defaults sit under the run's own a:rPr, so
// direct formatting always wins.
function parseSlideRun(node: PoNode, defaults: RunProperties): Run | undefined {
  const t = poChildren(node).find((c) => poIs(c, 'a:t'));
  const text = t ? poText(t) : '';
  if (text.length === 0) return undefined;
  const rPr = poChildren(node).find((c) => poIs(c, 'a:rPr'));
  return { text, properties: { ...defaults, ...rPrToRunProps(rPr) } };
}
