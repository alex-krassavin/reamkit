// ECMA-376 Part 1 §20 (DrawingML) + Part 3 (Markup Compatibility).
//
// Parses a <w:drawing> into either an embedded picture reference (the
// pre-existing behaviour) or a DrawingML shape (<wps:wsp>). Markup
// Compatibility <mc:AlternateContent> is resolved here, preferring the modern
// wps Choice over the legacy VML Fallback (which we cannot render).

import type {
  BodyElement,
  CustomPathCmd,
  ShapeDash,
  ShapeFill,
  ShapeGeometry,
  ShapeLine,
  ShapeTextBody,
  ShapeTransform,
} from '@/core/document-model';
import type { ColorMod, ColorResolver } from '@/core/drawingml/colors';
import type { PoNode } from '@/core/po-helpers';
import type { Pt } from '@/core/ir';
import { resolveColorNode } from '@/core/drawingml/colors';
import { emuToPt } from '@/core/ir';
import { poAttr, poChildren, poFindDescendant, poIntAttr, poIs, poTag } from '@/core/po-helpers';

const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

// Namespaces whose <mc:Choice> we can render. 'wps' = wordprocessingShape.
const UNDERSTOOD_NS = new Set(['wps']);

// Shape data without the owning paragraph's properties (attached by the caller,
// mirroring how the image branch returns size + id and the caller adds pPr).
export interface ShapeData {
  readonly width: Pt;
  readonly height: Pt;
  readonly geometry: ShapeGeometry;
  readonly fill: ShapeFill;
  readonly line?: ShapeLine;
  readonly transform?: ShapeTransform;
  readonly text?: ShapeTextBody;
}

// Parses the body elements of a w:txbxContent. Injected by the caller to avoid
// a cycle with document-parser (which imports this module).
export type ParseBody = (children: ReadonlyArray<PoNode>) => Array<BodyElement>;

export type DrawingContent =
  | {
      readonly kind: 'image';
      readonly imageId: string;
      readonly width: Pt;
      readonly height: Pt;
      // wp:docPr @descr/@title — alternate text for the tagged-PDF Figure.
      readonly altText?: string;
    }
  | { readonly kind: 'shape'; readonly data: ShapeData; readonly altText?: string }
  | {
      readonly kind: 'chart';
      readonly chartRelId: string;
      readonly width: Pt;
      readonly height: Pt;
      readonly altText?: string;
    };

// ECMA-376 Part 3 — resolve <mc:AlternateContent> to the children of the first
// <mc:Choice> whose Requires lists only namespaces we understand, else the
// <mc:Fallback> children, else nothing. (Requires holds space-separated
// namespace prefixes as declared in the document.)
export function resolveMc(altContent: PoNode): ReadonlyArray<PoNode> {
  for (const choice of poChildren(altContent)) {
    if (!poIs(choice, 'mc:Choice')) continue;
    const requires = (poAttr(choice, 'Requires') ?? '').split(/\s+/).filter(Boolean);
    if (requires.length > 0 && requires.every((r) => UNDERSTOOD_NS.has(r))) {
      return poChildren(choice);
    }
  }
  const fallback = poChildren(altContent).find((c) => poIs(c, 'mc:Fallback'));
  return fallback ? poChildren(fallback) : [];
}

// Flatten a children list, expanding any <mc:AlternateContent> to its chosen
// branch so downstream scanning sees plain elements (a <w:drawing>, or the VML
// we ignore). Used both at run level and inside a:graphicData.
export function expandMcChildren(children: ReadonlyArray<PoNode>): Array<PoNode> {
  const out: Array<PoNode> = [];
  for (const c of children) {
    if (poIs(c, 'mc:AlternateContent')) out.push(...resolveMc(c));
    else out.push(c);
  }
  return out;
}

export function parseDrawing(
  drawing: PoNode,
  resolveColor: ColorResolver,
  parseBody?: ParseBody,
): DrawingContent | null {
  const anchor =
    poChildren(drawing).find((c) => poIs(c, 'wp:inline')) ??
    poChildren(drawing).find((c) => poIs(c, 'wp:anchor'));
  if (!anchor) return null;

  const extent = poFindDescendant(anchor, 'wp:extent');
  const extentCx = extent ? poIntAttr(extent, 'cx') : undefined;
  const extentCy = extent ? poIntAttr(extent, 'cy') : undefined;

  // wp:docPr (a direct child of wp:inline/wp:anchor) carries the drawing's
  // alternate text: @descr preferred, then @title. Used for the Figure /Alt.
  const docPr = poChildren(anchor).find((c) => poIs(c, 'wp:docPr'));
  const descr = docPr ? poAttr(docPr, 'descr') : undefined;
  const title = docPr ? poAttr(docPr, 'title') : undefined;
  const altText = (descr ?? title)?.trim() || undefined;
  const alt = altText ? { altText } : {};

  const graphicData = poFindDescendant(anchor, 'a:graphicData');
  const uri = graphicData ? poAttr(graphicData, 'uri') : undefined;

  if (graphicData && uri === WPS_URI) {
    const data = parseWsp(graphicData, extentCx, extentCy, resolveColor, parseBody);
    return data ? { kind: 'shape', data, ...alt } : null;
  }

  if (graphicData && uri === CHART_URI) {
    const cChart = poFindDescendant(graphicData, 'c:chart');
    const chartRelId = cChart ? poAttr(cChart, 'id') : undefined; // r:id
    if (chartRelId && extentCx !== undefined && extentCy !== undefined) {
      return {
        kind: 'chart',
        chartRelId,
        width: emuToPt(extentCx),
        height: emuToPt(extentCy),
        ...alt,
      };
    }
    return null;
  }

  // Picture path: a:blip r:embed + extent.
  const blip = poFindDescendant(anchor, 'a:blip');
  const rId = blip ? poAttr(blip, 'embed') : undefined;
  if (extentCx !== undefined && extentCy !== undefined && rId) {
    return {
      kind: 'image',
      imageId: rId,
      width: emuToPt(extentCx),
      height: emuToPt(extentCy),
      ...alt,
    };
  }
  return null;
}

function parseWsp(
  graphicData: PoNode,
  extentCx: number | undefined,
  extentCy: number | undefined,
  resolveColor: ColorResolver,
  parseBody?: ParseBody,
): ShapeData | null {
  // wps:wsp is normally a direct child, but a nested mc:AlternateContent can
  // wrap it — expand first so either layout is found.
  const wsp = expandMcChildren(poChildren(graphicData)).find((c) => poIs(c, 'wps:wsp'));
  if (!wsp) return null;
  const spPr = poChildren(wsp).find((c) => poIs(c, 'wps:spPr'));

  let geometry: ShapeGeometry = { kind: 'preset', preset: 'rect', adjust: new Map() };
  let fill: ShapeFill = { kind: 'none' };
  let line: ShapeLine | undefined;
  let transform: ShapeTransform | undefined;
  let widthEmu = extentCx;
  let heightEmu = extentCy;

  if (spPr) {
    const xfrm = poChildren(spPr).find((c) => poIs(c, 'a:xfrm'));
    if (xfrm) {
      transform = parseXfrm(xfrm);
      if (widthEmu === undefined || heightEmu === undefined) {
        const ext = poChildren(xfrm).find((c) => poIs(c, 'a:ext'));
        if (ext) {
          widthEmu = widthEmu ?? poIntAttr(ext, 'cx');
          heightEmu = heightEmu ?? poIntAttr(ext, 'cy');
        }
      }
    }
    const prst = poChildren(spPr).find((c) => poIs(c, 'a:prstGeom'));
    const cust = poChildren(spPr).find((c) => poIs(c, 'a:custGeom'));
    if (prst) geometry = parsePrstGeom(prst);
    else if (cust) geometry = parseCustGeom(cust);
    fill = parseFill(spPr, resolveColor);
    line = parseLine(spPr, resolveColor);
  }

  const text = parseBody ? parseTextBox(wsp, parseBody) : undefined;

  if (widthEmu === undefined || heightEmu === undefined) return null;
  return {
    width: emuToPt(widthEmu),
    height: emuToPt(heightEmu),
    geometry,
    fill,
    ...(line ? { line } : {}),
    ...(transform ? { transform } : {}),
    ...(text ? { text } : {}),
  };
}

// wps:txbx/w:txbxContent (the text body) + wps:bodyPr (insets + vertical
// anchor). Returns undefined when the shape carries no text.
function parseTextBox(wsp: PoNode, parseBody: ParseBody): ShapeTextBody | undefined {
  const txbx = poChildren(wsp).find((c) => poIs(c, 'wps:txbx'));
  if (!txbx) return undefined;
  const txContent = poChildren(txbx).find((c) => poIs(c, 'w:txbxContent'));
  if (!txContent) return undefined;
  const content = parseBody(poChildren(txContent));
  if (content.length === 0) return undefined;

  const bodyPr = poChildren(wsp).find((c) => poIs(c, 'wps:bodyPr'));
  const lIns = bodyPr ? poIntAttr(bodyPr, 'lIns') : undefined;
  const tIns = bodyPr ? poIntAttr(bodyPr, 'tIns') : undefined;
  const rIns = bodyPr ? poIntAttr(bodyPr, 'rIns') : undefined;
  const bIns = bodyPr ? poIntAttr(bodyPr, 'bIns') : undefined;
  const a = bodyPr ? poAttr(bodyPr, 'anchor') : undefined;
  const anchor: ShapeTextBody['anchor'] | undefined =
    a === 'ctr' ? 'ctr' : a === 'b' ? 'b' : a === 't' ? 't' : undefined;

  return {
    content,
    ...(lIns !== undefined ? { insetLeft: emuToPt(lIns) } : {}),
    ...(tIns !== undefined ? { insetTop: emuToPt(tIns) } : {}),
    ...(rIns !== undefined ? { insetRight: emuToPt(rIns) } : {}),
    ...(bIns !== undefined ? { insetBottom: emuToPt(bIns) } : {}),
    ...(anchor ? { anchor } : {}),
  };
}

const isTrue = (v: string | undefined): boolean => v === '1' || v === 'true' || v === 'on';

function parseXfrm(xfrm: PoNode): ShapeTransform {
  const rot = poIntAttr(xfrm, 'rot');
  const flipH = isTrue(poAttr(xfrm, 'flipH'));
  const flipV = isTrue(poAttr(xfrm, 'flipV'));
  return {
    ...(rot !== undefined ? { rotation60k: rot } : {}),
    ...(flipH ? { flipH: true } : {}),
    ...(flipV ? { flipV: true } : {}),
  };
}

function parsePrstGeom(prst: PoNode): ShapeGeometry {
  const preset = poAttr(prst, 'prst') ?? 'rect';
  const adjust = new Map<string, number>();
  const avLst = poChildren(prst).find((c) => poIs(c, 'a:avLst'));
  if (avLst) {
    for (const gd of poChildren(avLst)) {
      if (!poIs(gd, 'a:gd')) continue;
      const nm = poAttr(gd, 'name');
      const fmla = poAttr(gd, 'fmla');
      if (!nm || !fmla) continue;
      const m = /^val\s+(-?\d+)/.exec(fmla);
      if (m) adjust.set(nm, Number(m[1]));
    }
  }
  return { kind: 'preset', preset, adjust };
}

// ECMA-376 §20.1.9.11 custGeom → §20.1.9.15 path. Parses the first <a:path>
// (multiple subpaths with differing w/h are a follow-up). Coordinates stay in
// path-space; the geometry layer scales + y-flips them. Falls back to a rect
// preset when the path is empty or has no usable size.
function parseCustGeom(cust: PoNode): ShapeGeometry {
  const pathLst = poChildren(cust).find((c) => poIs(c, 'a:pathLst'));
  const path = pathLst ? poChildren(pathLst).find((c) => poIs(c, 'a:path')) : undefined;
  if (!path) return { kind: 'preset', preset: 'rect', adjust: new Map() };

  const pathWidth = poIntAttr(path, 'w') ?? 0;
  const pathHeight = poIntAttr(path, 'h') ?? 0;
  const commands: Array<CustomPathCmd> = [];
  for (const node of poChildren(path)) {
    switch (poTag(node)) {
      case 'a:moveTo': {
        const p = firstPt(node);
        if (p) commands.push({ cmd: 'move', x: p.x, y: p.y });
        break;
      }
      case 'a:lnTo': {
        const p = firstPt(node);
        if (p) commands.push({ cmd: 'line', x: p.x, y: p.y });
        break;
      }
      case 'a:cubicBezTo': {
        const p = pts(node);
        if (p.length >= 3)
          commands.push({
            cmd: 'cubic',
            x1: p[0]!.x,
            y1: p[0]!.y,
            x2: p[1]!.x,
            y2: p[1]!.y,
            x: p[2]!.x,
            y: p[2]!.y,
          });
        break;
      }
      case 'a:quadBezTo': {
        const p = pts(node);
        if (p.length >= 2)
          commands.push({ cmd: 'quad', x1: p[0]!.x, y1: p[0]!.y, x: p[1]!.x, y: p[1]!.y });
        break;
      }
      case 'a:arcTo':
        commands.push({
          cmd: 'arc',
          wR: poIntAttr(node, 'wR') ?? 0,
          hR: poIntAttr(node, 'hR') ?? 0,
          stAng: poIntAttr(node, 'stAng') ?? 0,
          swAng: poIntAttr(node, 'swAng') ?? 0,
        });
        break;
      case 'a:close':
        commands.push({ cmd: 'close' });
        break;
    }
  }

  if (pathWidth <= 0 || pathHeight <= 0 || commands.length === 0) {
    return { kind: 'preset', preset: 'rect', adjust: new Map() };
  }
  return { kind: 'custom', custom: { pathWidth, pathHeight, commands } };
}

// <a:pt x= y=> children of a path command.
function pts(node: PoNode): Array<{ x: number; y: number }> {
  const out: Array<{ x: number; y: number }> = [];
  for (const c of poChildren(node)) {
    if (!poIs(c, 'a:pt')) continue;
    const x = poIntAttr(c, 'x');
    const y = poIntAttr(c, 'y');
    if (x !== undefined && y !== undefined) out.push({ x, y });
  }
  return out;
}

const firstPt = (node: PoNode): { x: number; y: number } | undefined => pts(node)[0];

function parseFill(spPr: PoNode, resolveColor: ColorResolver): ShapeFill {
  for (const child of poChildren(spPr)) {
    if (poIs(child, 'a:noFill')) return { kind: 'none' };
    if (poIs(child, 'a:solidFill')) {
      const hex = colorFromContainer(child, resolveColor);
      return hex ? { kind: 'solid', colorHex: hex } : { kind: 'none' };
    }
    if (poIs(child, 'a:gradFill')) {
      const hex = colorFromGradient(child, resolveColor);
      return hex ? { kind: 'solid', colorHex: hex } : { kind: 'none' };
    }
  }
  return { kind: 'none' };
}

function parseLine(spPr: PoNode, resolveColor: ColorResolver): ShapeLine | undefined {
  const ln = poChildren(spPr).find((c) => poIs(c, 'a:ln'));
  if (!ln) return undefined;
  const widthEmu = poIntAttr(ln, 'w');
  // a:ln @cap: flat | rnd | sq (§20.1.10.31).
  const capRaw = poAttr(ln, 'cap');
  const cap: ShapeLine['cap'] | undefined =
    capRaw === 'rnd'
      ? 'round'
      : capRaw === 'sq'
        ? 'square'
        : capRaw === 'flat'
          ? 'flat'
          : undefined;
  let noFill = false;
  let colorHex: string | undefined;
  let dash: ShapeDash | undefined;
  for (const c of poChildren(ln)) {
    if (poIs(c, 'a:noFill')) noFill = true;
    else if (poIs(c, 'a:solidFill')) colorHex = colorFromContainer(c, resolveColor);
    else if (poIs(c, 'a:prstDash')) dash = normalizeDash(poAttr(c, 'val'));
  }
  return {
    ...(widthEmu !== undefined ? { width: emuToPt(widthEmu) } : {}),
    ...(colorHex ? { colorHex } : {}),
    ...(dash ? { dash } : {}),
    ...(cap ? { cap } : {}),
    ...(noFill ? { fill: 'none' as const } : {}),
  };
}

const DASH_VALUES = new Set<ShapeDash>([
  'solid',
  'dot',
  'dash',
  'dashDot',
  'lgDash',
  'lgDashDot',
  'sysDash',
  'sysDot',
]);

// Map a:prstDash @val to a supported ShapeDash, folding rarer variants onto
// their nearest supported pattern.
function normalizeDash(v: string | undefined): ShapeDash | undefined {
  if (!v) return undefined;
  if (DASH_VALUES.has(v as ShapeDash)) return v as ShapeDash;
  if (v === 'lgDashDotDot') return 'lgDashDot';
  if (v === 'sysDashDot' || v === 'sysDashDotDot') return 'sysDash';
  return undefined;
}

// First a:srgbClr / a:schemeClr child → resolved hex (with colour transforms).
function colorFromContainer(parent: PoNode, resolveColor: ColorResolver): string | undefined {
  for (const c of poChildren(parent)) {
    const hex = resolveColorNode(c, resolveColor);
    if (hex) return hex;
  }
  return undefined;
}

// a:gradFill → the average of its gradient-stop colours (a solid approximation;
// PDF axial/radial shadings are not emitted). Reads a:gsLst/a:gs colour children.
function colorFromGradient(grad: PoNode, resolveColor: ColorResolver): string | undefined {
  const gsLst = poChildren(grad).find((c) => poIs(c, 'a:gsLst'));
  if (!gsLst) return undefined;
  let r = 0;
  let g = 0;
  let b = 0;
  let n = 0;
  for (const gs of poChildren(gsLst)) {
    if (!poIs(gs, 'a:gs')) continue;
    let hex: string | undefined;
    for (const c of poChildren(gs)) {
      hex = resolveColorNode(c, resolveColor);
      if (hex) break;
    }
    if (!hex) continue;
    const num = parseInt(hex, 16);
    r += (num >> 16) & 255;
    g += (num >> 8) & 255;
    b += num & 255;
    n++;
  }
  if (n === 0) return undefined;
  const toHex = (x: number): string =>
    Math.round(x / n)
      .toString(16)
      .padStart(2, '0');
  return (toHex(r) + toHex(g) + toHex(b)).toUpperCase();
}
