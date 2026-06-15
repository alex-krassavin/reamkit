// E-PPTX — shared shape (p:sp) parsing helpers, used by both the slide parser
// and the placeholder cascade (a leaf module, so neither imports the other).
//
// Slides, layouts and masters all describe shapes the same way: a p:nvSpPr (with
// an optional p:ph placeholder reference), a p:spPr (with the a:xfrm transform)
// and a p:txBody. These helpers read the pieces every layer shares.

import type { RunProperties } from '@/core/document-model';
import type { ColorResolver } from '@/core/drawingml/colors';
import type { PoNode } from '@/core/po-helpers';

import { resolveColorNode } from '@/core/drawingml/colors';
import { pt } from '@/core/ir';
import { poAttr, poChildren, poIntAttr, poIs } from '@/core/po-helpers';

// §19.3.1.36 p:ph — a shape's placeholder reference. type/idx tie a slide shape
// to the matching prototype in its layout and master.
export interface PlaceholderRef {
  readonly type?: string; // e.g. 'title', 'ctrTitle', 'subTitle', 'body', 'obj'
  readonly idx?: string;
}

// An a:xfrm box in EMU (off + ext).
export interface ShapeBoxEmu {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

const isTrue = (v: string | undefined): boolean => v === '1' || v === 'true' || v === 'on';

// p:nvSpPr/p:nvPr/p:ph → the placeholder reference, or undefined for an ordinary
// (non-placeholder) shape. A bare <p:ph/> returns an empty ref (still a
// placeholder: type defaults to 'obj').
export function parsePh(sp: PoNode): PlaceholderRef | undefined {
  const nvSpPr = poChildren(sp).find((c) => poIs(c, 'p:nvSpPr'));
  const nvPr = nvSpPr ? poChildren(nvSpPr).find((c) => poIs(c, 'p:nvPr')) : undefined;
  const ph = nvPr ? poChildren(nvPr).find((c) => poIs(c, 'p:ph')) : undefined;
  if (!ph) return undefined;
  const type = poAttr(ph, 'type');
  const idx = poAttr(ph, 'idx');
  return {
    ...(type !== undefined ? { type } : {}),
    ...(idx !== undefined ? { idx } : {}),
  };
}

// An xfrm node (a:off + a:ext children) → the EMU box. Works for both a:xfrm
// (in p:spPr) and p:xfrm (on a p:graphicFrame), whose children are identical.
export function boxFromXfrm(xfrm: PoNode | undefined): ShapeBoxEmu | undefined {
  if (!xfrm) return undefined;
  const off = poChildren(xfrm).find((c) => poIs(c, 'a:off'));
  const ext = poChildren(xfrm).find((c) => poIs(c, 'a:ext'));
  const cx = ext ? poIntAttr(ext, 'cx') : undefined;
  const cy = ext ? poIntAttr(ext, 'cy') : undefined;
  if (cx === undefined || cy === undefined || cx <= 0 || cy <= 0) return undefined;
  const x = (off ? poIntAttr(off, 'x') : undefined) ?? 0;
  const y = (off ? poIntAttr(off, 'y') : undefined) ?? 0;
  return { x, y, cx, cy };
}

// p:spPr/a:xfrm → the EMU box, or undefined when the shape carries no explicit
// transform (a placeholder that inherits it from the layout/master).
export function parseXfrmBox(spPr: PoNode | undefined): ShapeBoxEmu | undefined {
  return boxFromXfrm(spPr ? poChildren(spPr).find((c) => poIs(c, 'a:xfrm')) : undefined);
}

// a:rPr / a:defRPr → RunProperties. Both the direct run formatting on a slide
// (a:rPr) and the per-level defaults in a master's p:txStyles (a:defRPr) share
// this element, so the cascade reuses the same reader: size (sz, hundredths of a
// point), bold/italic/underline, solid colour and the latin typeface. Scheme
// colours and the theme font cascade come in PX5.
export function rPrToRunProps(rPr: PoNode | undefined, colors: ColorResolver): RunProperties {
  if (!rPr) return {};
  const sz = poIntAttr(rPr, 'sz');
  const u = poAttr(rPr, 'u');
  const colorHex = solidFillColor(rPr, colors);
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

// a:solidFill → resolved 6-hex (no '#'), via the deck's colour resolver so both
// a:srgbClr and a:schemeClr (PX5 theme) work.
function solidFillColor(rPr: PoNode, colors: ColorResolver): string | undefined {
  const solidFill = poChildren(rPr).find((c) => poIs(c, 'a:solidFill'));
  if (!solidFill) return undefined;
  for (const c of poChildren(solidFill)) {
    const hex = resolveColorNode(c, colors);
    if (hex) return hex;
  }
  return undefined;
}
