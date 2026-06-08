// From-scratch math typesetting engine (ECMA-376 §22 OfficeMathML → boxes).
//
// layoutMath recursively turns a MathNode into a MathBox: a width plus ascent
// (above the baseline) and descent (below), and a flat list of positioned draw
// items in a LOCAL frame (origin = box left, baseline at y=0, y up). Ordinary
// symbols become glyph items (rendered from the font); structural elements
// (fraction rules now; radicals, big operators, stretchy delimiters later)
// become rule/path items drawn as vector graphics — so no math font is needed.
//
// Pure: text widths arrive through an injected measure fn, mirroring
// chart-geometry. The renderer maps variants to fonts and emits the items.

import type {
  MathAccent,
  MathBar,
  MathDelimiter,
  MathEqArray,
  MathFunc,
  MathGroupChr,
  MathLimit,
  MathMatrix,
  MathNary,
  MathNode,
  MathRadical,
  MathRun,
  MathScript,
} from '@/document-model';
import type { PathSegment } from '@/pdf/vector-graphics';

export type MathVariant = 'regular' | 'italic' | 'bold' | 'boldItalic';

export interface MathGlyphItem {
  readonly kind: 'glyph';
  readonly x: number;
  readonly y: number; // baseline offset (up positive)
  readonly text: string;
  readonly variant: MathVariant;
  readonly sizePt: number;
}
export interface MathRuleItem {
  readonly kind: 'rule'; // filled rectangle (y = bottom edge)
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}
export interface MathPathItem {
  readonly kind: 'path'; // vector path (radicals, delimiters, big operators)
  readonly segments: ReadonlyArray<PathSegment>;
  readonly strokeWidthPt?: number; // present ⇒ stroked
  readonly fill?: boolean; // true ⇒ filled
}
export type MathDrawItem = MathGlyphItem | MathRuleItem | MathPathItem;

export interface MathBox {
  readonly width: number;
  readonly ascent: number;
  readonly descent: number;
  readonly items: ReadonlyArray<MathDrawItem>;
}

export type MeasureMath = (text: string, sizePt: number, variant: MathVariant) => number;

export interface MathCtx {
  readonly sizePt: number;
}

// Metrics as fractions of the em (font size). Loosely TeX-inspired.
const ASCENT = 0.72; // ordinary symbol ascent
const DESCENT = 0.2; // ordinary symbol descent
const AXIS = 0.26; // math axis height (where fraction bars sit)
const BAR_THICK = 0.045; // fraction rule thickness
const FRAC_GAP = 0.15; // gap between rule and num/den
const FRAC_PAD = 0.12; // horizontal padding around a fraction
const SCRIPT_SIZE = 0.7; // sub/superscript size relative to base

export function variantStyle(v: MathVariant): { bold: boolean; italic: boolean } {
  return {
    bold: v === 'bold' || v === 'boldItalic',
    italic: v === 'italic' || v === 'boldItalic',
  };
}

const variantOf = (italic: boolean, bold: boolean): MathVariant =>
  bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : 'regular';

// Auto-italic: Latin letters and Greek letters render italic in math by default.
function isAutoItalic(ch: string): boolean {
  return /[A-Za-z]/.test(ch) || (ch >= 'α' && ch <= 'ω') || (ch >= 'Α' && ch <= 'Ω');
}

function runCharVariant(ch: string, run: MathRun): MathVariant {
  const bold = run.bold ?? false;
  const italic = run.nor ? false : (run.italic ?? isAutoItalic(ch));
  return variantOf(italic, bold);
}

export function layoutMath(node: MathNode, ctx: MathCtx, measure: MeasureMath): MathBox {
  switch (node.type) {
    case 'run':
      return layoutRun(node, ctx, measure);
    case 'row':
      return layoutRow(node, ctx, measure);
    case 'fraction':
      return layoutFraction(node, ctx, measure);
    case 'script':
      return layoutScript(node, ctx, measure);
    case 'radical':
      return layoutRadical(node, ctx, measure);
    case 'nary':
      return layoutNary(node, ctx, measure);
    case 'func':
      return layoutFunc(node, ctx, measure);
    case 'limit':
      return layoutLimit(node, ctx, measure);
    case 'delimiter':
      return layoutDelimiter(node, ctx, measure);
    case 'matrix':
      return layoutMatrix(node, ctx, measure);
    case 'eqArr':
      return layoutEqArray(node, ctx, measure);
    case 'accent':
      return layoutAccent(node, ctx, measure);
    case 'bar':
      return layoutBar(node, ctx, measure);
    case 'groupChr':
      return layoutGroupChr(node, ctx, measure);
    default:
      // Constructs added in later milestones — render as empty until then.
      return emptyBox();
  }
}

const emptyBox = (): MathBox => ({ width: 0, ascent: 0, descent: 0, items: [] });

// Binary operators and relations take medium space on each side (math spacing).
const MATH_BINARY = new Set(['+', '−', '-', '±', '∓', '×', '÷', '⋅', '∗', '⊕', '⊗', '∘']);
const MATH_RELATION = new Set([
  '=',
  '≠',
  '<',
  '>',
  '≤',
  '≥',
  '≈',
  '≡',
  '≅',
  '∝',
  '→',
  '←',
  '↔',
  '∈',
  '∉',
  '⊂',
  '⊃',
  '⊆',
  '⊇',
  '≪',
  '≫',
]);

// Interior binary operators / relations are spaced; a leading sign (−x) is
// unary and gets no space.
function isSpacedOp(ch: string, i: number, n: number): boolean {
  return (MATH_BINARY.has(ch) || MATH_RELATION.has(ch)) && i > 0 && i < n - 1;
}

function layoutRun(run: MathRun, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const chars = [...run.text];
  if (chars.length === 0) return emptyBox();
  const med = 0.2 * size;
  const items: Array<MathDrawItem> = [];
  let x = 0;
  let i = 0;
  while (i < chars.length) {
    const v = runCharVariant(chars[i]!, run);
    if (isSpacedOp(chars[i]!, i, chars.length)) {
      x += med;
      items.push({ kind: 'glyph', x, y: 0, text: chars[i]!, variant: v, sizePt: size });
      x += measure(chars[i]!, size, v) + med;
      i++;
      continue;
    }
    // Group consecutive non-operator characters of the same variant.
    let seg = '';
    let j = i;
    while (
      j < chars.length &&
      runCharVariant(chars[j]!, run) === v &&
      !isSpacedOp(chars[j]!, j, chars.length)
    ) {
      seg += chars[j];
      j++;
    }
    items.push({ kind: 'glyph', x, y: 0, text: seg, variant: v, sizePt: size });
    x += measure(seg, size, v);
    i = j;
  }
  return { width: x, ascent: size * ASCENT, descent: size * DESCENT, items };
}

function layoutRow(
  node: { readonly children: ReadonlyArray<MathNode> },
  ctx: MathCtx,
  measure: MeasureMath,
): MathBox {
  let x = 0;
  let ascent = 0;
  let descent = 0;
  const items: Array<MathDrawItem> = [];
  for (const child of node.children) {
    const box = layoutMath(child, ctx, measure);
    for (const it of shiftItems(box.items, x, 0)) items.push(it);
    x += box.width;
    ascent = Math.max(ascent, box.ascent);
    descent = Math.max(descent, box.descent);
  }
  return { width: x, ascent, descent, items };
}

function layoutFraction(
  node: { readonly num: MathNode; readonly den: MathNode; readonly barless?: boolean },
  ctx: MathCtx,
  measure: MeasureMath,
): MathBox {
  const size = ctx.sizePt;
  const numBox = layoutMath(node.num, ctx, measure);
  const denBox = layoutMath(node.den, ctx, measure);
  const pad = FRAC_PAD * size;
  const width = Math.max(numBox.width, denBox.width) + 2 * pad;
  const barThick = Math.max(BAR_THICK * size, 0.5);
  const axis = AXIS * size;
  const gap = FRAC_GAP * size;

  const items: Array<MathDrawItem> = [];
  const numBaseY = axis + barThick / 2 + gap + numBox.descent;
  const numX = (width - numBox.width) / 2;
  for (const it of shiftItems(numBox.items, numX, numBaseY)) items.push(it);

  const denBaseY = axis - barThick / 2 - gap - denBox.ascent;
  const denX = (width - denBox.width) / 2;
  for (const it of shiftItems(denBox.items, denX, denBaseY)) items.push(it);

  if (!node.barless) {
    items.push({ kind: 'rule', x: pad * 0.5, y: axis - barThick / 2, w: width - pad, h: barThick });
  }
  const ascent = numBaseY + numBox.ascent;
  const descent = -(denBaseY - denBox.descent);
  return { width, ascent, descent, items };
}

// m:sSup / m:sSub / m:sSubSup / m:sPre — a base with sub/superscripts.
function layoutScript(node: MathScript, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const scriptSize = size * SCRIPT_SIZE;
  const baseBox = layoutMath(node.base, ctx, measure);
  const supBox = node.sup ? layoutMath(node.sup, { sizePt: scriptSize }, measure) : undefined;
  const subBox = node.sub ? layoutMath(node.sub, { sizePt: scriptSize }, measure) : undefined;
  const kern = 0.05 * size;
  const scriptW = Math.max(supBox?.width ?? 0, subBox?.width ?? 0);

  const baseX = node.pre ? scriptW + kern : 0;
  const scriptX = node.pre ? 0 : baseBox.width + kern;

  const items: Array<MathDrawItem> = [...shiftItems(baseBox.items, baseX, 0)];
  let ascent = baseBox.ascent;
  let descent = baseBox.descent;

  if (supBox) {
    const supShift = Math.max(baseBox.ascent - 0.25 * scriptSize, 0.45 * size);
    for (const it of shiftItems(supBox.items, scriptX, supShift)) items.push(it);
    ascent = Math.max(ascent, supShift + supBox.ascent);
  }
  if (subBox) {
    const subShift = Math.max(baseBox.descent + 0.2 * scriptSize, 0.18 * size);
    for (const it of shiftItems(subBox.items, scriptX, -subShift)) items.push(it);
    descent = Math.max(descent, subShift + subBox.descent);
  }
  return { width: baseBox.width + kern + scriptW, ascent, descent, items };
}

// m:rad — radical. The surd (√) is drawn as a stroked path sized to the
// radicand, with the vinculum (overbar) extending across it.
function layoutRadical(node: MathRadical, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const radBox = layoutMath(node.radicand, ctx, measure);
  const rule = Math.max(0.05 * size, 0.6);
  const gap = 0.12 * size;
  const radAsc = radBox.ascent;
  const radDesc = radBox.descent;
  const radH = radAsc + radDesc;
  const barY = radAsc + gap;
  const surdW = Math.max(0.5 * size, radH * 0.45);
  const bottom = -radDesc - rule * 0.5;
  const pad = 0.12 * size;

  const items: Array<MathDrawItem> = [];
  let degW = 0;
  if (node.degree) {
    const degBox = layoutMath(node.degree, { sizePt: size * 0.5 }, measure);
    degW = degBox.width;
    for (const it of shiftItems(degBox.items, 0, barY * 0.35)) items.push(it);
  }
  const surdX = degW;
  const startY = barY - radH * 0.5;
  const surd: Array<PathSegment> = [
    { op: 'move', x: surdX, y: startY },
    { op: 'line', x: surdX + surdW * 0.45, y: bottom },
    { op: 'line', x: surdX + surdW * 0.85, y: barY },
    { op: 'line', x: surdX + surdW + radBox.width + pad, y: barY },
  ];
  items.push({ kind: 'path', segments: surd, strokeWidthPt: Math.max(rule, 0.8) });
  for (const it of shiftItems(radBox.items, surdX + surdW, 0)) items.push(it);

  return {
    width: degW + surdW + radBox.width + pad,
    ascent: barY + rule,
    descent: radDesc + rule,
    items,
  };
}

// Big operators drawn as vector paths (so no math font is needed). Sum-like
// operators default to limits under/over; integrals to sub/super scripts.
const NARY_AXIS = 0.28; // visual centre above the baseline (× em)
const SUM_LIKE = new Set(['∑', '∏', '∐', '⋃', '⋂', '⋁', '⋀', '⨀', '⨁', '⨂']);

function layoutBigOp(op: string, size: number, measure: MeasureMath): MathBox {
  const axisC = NARY_AXIS * size;
  if (op === '∑') {
    const half = size * 0.75;
    const top = axisC + half;
    const bot = axisC - half;
    const w = size * 0.85;
    const segs: Array<PathSegment> = [
      { op: 'move', x: w, y: top },
      { op: 'line', x: 0, y: top },
      { op: 'line', x: w * 0.52, y: axisC },
      { op: 'line', x: 0, y: bot },
      { op: 'line', x: w, y: bot },
    ];
    return bigOpBox(w + size * 0.08, top, bot, segs, Math.max(size * 0.075, 1));
  }
  if (op === '∏') {
    const half = size * 0.75;
    const top = axisC + half;
    const bot = axisC - half;
    const w = size * 0.9;
    const segs: Array<PathSegment> = [
      { op: 'move', x: 0, y: top },
      { op: 'line', x: w, y: top },
      { op: 'move', x: w * 0.22, y: top },
      { op: 'line', x: w * 0.22, y: bot },
      { op: 'move', x: w * 0.78, y: top },
      { op: 'line', x: w * 0.78, y: bot },
    ];
    return bigOpBox(w + size * 0.08, top, bot, segs, Math.max(size * 0.08, 1));
  }
  if (op === '∫' || op === '∮') {
    const half = size * 0.95;
    const top = axisC + half;
    const bot = axisC - half;
    const w = size * 0.5;
    const segs: Array<PathSegment> = [
      { op: 'move', x: w * 0.72, y: top },
      { op: 'cubic', x1: w * 0.2, y1: top, x2: w * 0.78, y2: axisC, x: w * 0.42, y: axisC },
      { op: 'cubic', x1: w * 0.06, y1: axisC, x2: w * 0.64, y2: bot, x: w * 0.12, y: bot },
    ];
    return bigOpBox(w + size * 0.06, top, bot, segs, Math.max(size * 0.06, 0.9));
  }
  // Fallback: an enlarged glyph for less common operators.
  const opSize = size * 1.4;
  return {
    width: measure(op, opSize, 'regular'),
    ascent: opSize * ASCENT,
    descent: opSize * DESCENT,
    items: [{ kind: 'glyph', x: 0, y: 0, text: op, variant: 'regular', sizePt: opSize }],
  };
}

function bigOpBox(
  width: number,
  top: number,
  bot: number,
  segments: Array<PathSegment>,
  strokeWidthPt: number,
): MathBox {
  return { width, ascent: top, descent: -bot, items: [{ kind: 'path', segments, strokeWidthPt }] };
}

function layoutNary(node: MathNary, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const opBox = layoutBigOp(node.op, size, measure);
  const scriptSize = size * SCRIPT_SIZE;
  const subBox = node.sub ? layoutMath(node.sub, { sizePt: scriptSize }, measure) : undefined;
  const supBox = node.sup ? layoutMath(node.sup, { sizePt: scriptSize }, measure) : undefined;
  const bodyBox = layoutMath(node.body, ctx, measure);
  const undOvr = node.limLoc === 'undOvr' || (node.limLoc === undefined && SUM_LIKE.has(node.op));

  const items: Array<MathDrawItem> = [];
  let ascent = opBox.ascent;
  let descent = opBox.descent;
  let opGroupW: number;

  if (undOvr) {
    const stackW = Math.max(opBox.width, subBox?.width ?? 0, supBox?.width ?? 0);
    for (const it of shiftItems(opBox.items, (stackW - opBox.width) / 2, 0)) items.push(it);
    const gap = 0.06 * size;
    if (supBox) {
      const supY = opBox.ascent + gap + supBox.descent;
      for (const it of shiftItems(supBox.items, (stackW - supBox.width) / 2, supY)) items.push(it);
      ascent = Math.max(ascent, supY + supBox.ascent);
    }
    if (subBox) {
      const subY = -(opBox.descent + gap + subBox.ascent);
      for (const it of shiftItems(subBox.items, (stackW - subBox.width) / 2, subY)) items.push(it);
      descent = Math.max(descent, -(subY - subBox.descent));
    }
    opGroupW = stackW;
  } else {
    for (const it of opBox.items) items.push(it);
    const scriptX = opBox.width + 0.02 * size;
    let scriptW = 0;
    if (supBox) {
      const supY = opBox.ascent - 0.4 * scriptSize;
      for (const it of shiftItems(supBox.items, scriptX, supY)) items.push(it);
      ascent = Math.max(ascent, supY + supBox.ascent);
      scriptW = Math.max(scriptW, supBox.width);
    }
    if (subBox) {
      const subY = -opBox.descent + 0.2 * scriptSize - subBox.ascent;
      for (const it of shiftItems(subBox.items, scriptX, subY)) items.push(it);
      descent = Math.max(descent, -(subY - subBox.descent));
      scriptW = Math.max(scriptW, subBox.width);
    }
    opGroupW = opBox.width + 0.02 * size + scriptW;
  }

  const bodyX = opGroupW + 0.12 * size;
  for (const it of shiftItems(bodyBox.items, bodyX, 0)) items.push(it);
  return {
    width: bodyX + bodyBox.width,
    ascent: Math.max(ascent, bodyBox.ascent),
    descent: Math.max(descent, bodyBox.descent),
    items,
  };
}

// m:func — a function name then a thin space then its argument.
function layoutFunc(node: MathFunc, ctx: MathCtx, measure: MeasureMath): MathBox {
  const nameBox = layoutMath(node.name, ctx, measure);
  const bodyBox = layoutMath(node.body, ctx, measure);
  const sp = 0.16 * ctx.sizePt;
  const items: Array<MathDrawItem> = [
    ...nameBox.items,
    ...shiftItems(bodyBox.items, nameBox.width + sp, 0),
  ];
  return {
    width: nameBox.width + sp + bodyBox.width,
    ascent: Math.max(nameBox.ascent, bodyBox.ascent),
    descent: Math.max(nameBox.descent, bodyBox.descent),
    items,
  };
}

// m:limLow / m:limUpp — a limit centred below/above the base.
function layoutLimit(node: MathLimit, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const baseBox = layoutMath(node.base, ctx, measure);
  const limBox = layoutMath(node.lim, { sizePt: size * SCRIPT_SIZE }, measure);
  const w = Math.max(baseBox.width, limBox.width);
  const items: Array<MathDrawItem> = [...shiftItems(baseBox.items, (w - baseBox.width) / 2, 0)];
  const gap = 0.08 * size;
  let ascent = baseBox.ascent;
  let descent = baseBox.descent;
  if (node.pos === 'low') {
    const limY = -(baseBox.descent + gap + limBox.ascent);
    for (const it of shiftItems(limBox.items, (w - limBox.width) / 2, limY)) items.push(it);
    descent = -(limY - limBox.descent);
  } else {
    const limY = baseBox.ascent + gap + limBox.descent;
    for (const it of shiftItems(limBox.items, (w - limBox.width) / 2, limY)) items.push(it);
    ascent = limY + limBox.ascent;
  }
  return { width: w, ascent, descent, items };
}

// A stretchy delimiter drawn as a path, sized to the content's ascent/descent.
function bracketGlyph(
  chr: string,
  asc: number,
  desc: number,
  size: number,
): { width: number; items: Array<MathDrawItem> } {
  if (chr === '') return { width: 0, items: [] };
  const margin = 0.1 * size;
  const top = asc + margin;
  const bot = -desc - margin;
  const mid = (top + bot) / 2;
  const stroke = Math.max(0.06 * size, 0.8);
  const isOpen = '([{⟨〈⌊⌈'.includes(chr);
  const path = (
    segments: Array<PathSegment>,
    width: number,
  ): { width: number; items: Array<MathDrawItem> } => ({
    width,
    items: [{ kind: 'path', segments, strokeWidthPt: stroke }],
  });

  if (chr === '(' || chr === ')') {
    const w = 0.32 * size;
    const bulge = w * 0.85;
    const x0 = isOpen ? w : 0;
    const dir = isOpen ? -1 : 1;
    return path(
      [
        { op: 'move', x: x0, y: top },
        {
          op: 'cubic',
          x1: x0 + dir * bulge,
          y1: mid + 0.45 * (top - mid),
          x2: x0 + dir * bulge,
          y2: mid - 0.45 * (top - mid),
          x: x0,
          y: bot,
        },
      ],
      w,
    );
  }
  if (chr === '[' || chr === ']') {
    const w = 0.26 * size;
    const xV = isOpen ? 0 : w;
    const xA = isOpen ? w : 0;
    return path(
      [
        { op: 'move', x: xA, y: top },
        { op: 'line', x: xV, y: top },
        { op: 'line', x: xV, y: bot },
        { op: 'line', x: xA, y: bot },
      ],
      w,
    );
  }
  if (chr === '{' || chr === '}') {
    const w = 0.4 * size;
    const dir = isOpen ? 1 : -1;
    const xb = isOpen ? w : 0;
    const tip = xb - dir * w;
    return path(
      [
        { op: 'move', x: xb, y: top },
        {
          op: 'cubic',
          x1: xb - dir * w * 0.5,
          y1: top,
          x2: xb - dir * w * 0.5,
          y2: mid + 0.25 * (top - mid),
          x: tip,
          y: mid,
        },
        {
          op: 'cubic',
          x1: xb - dir * w * 0.5,
          y1: mid - 0.25 * (top - mid),
          x2: xb - dir * w * 0.5,
          y2: bot,
          x: xb,
          y: bot,
        },
      ],
      w,
    );
  }
  if (chr === '⟨' || chr === '⟩' || chr === '〈' || chr === '〉') {
    const w = 0.3 * size;
    const xPoint = isOpen ? 0 : w;
    const xOpen = isOpen ? w : 0;
    return path(
      [
        { op: 'move', x: xOpen, y: top },
        { op: 'line', x: xPoint, y: mid },
        { op: 'line', x: xOpen, y: bot },
      ],
      w,
    );
  }
  if (chr === '‖') {
    const w = 0.3 * size;
    return path(
      [
        { op: 'move', x: w * 0.3, y: top },
        { op: 'line', x: w * 0.3, y: bot },
        { op: 'move', x: w * 0.7, y: top },
        { op: 'line', x: w * 0.7, y: bot },
      ],
      w,
    );
  }
  // '|' and any other delimiter → a single vertical stroke.
  const w = 0.18 * size;
  return path(
    [
      { op: 'move', x: w / 2, y: top },
      { op: 'line', x: w / 2, y: bot },
    ],
    w,
  );
}

function layoutDelimiter(node: MathDelimiter, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const inner: Array<MathDrawItem> = [];
  let innerW = 0;
  let innerAsc = 0;
  let innerDesc = 0;
  node.children.forEach((child, i) => {
    if (i > 0 && node.sepChr) {
      const sepBox = layoutMath({ type: 'run', text: node.sepChr, nor: true }, ctx, measure);
      for (const it of shiftItems(sepBox.items, innerW, 0)) inner.push(it);
      innerW += sepBox.width;
      innerAsc = Math.max(innerAsc, sepBox.ascent);
      innerDesc = Math.max(innerDesc, sepBox.descent);
    }
    const box = layoutMath(child, ctx, measure);
    for (const it of shiftItems(box.items, innerW, 0)) inner.push(it);
    innerW += box.width;
    innerAsc = Math.max(innerAsc, box.ascent);
    innerDesc = Math.max(innerDesc, box.descent);
  });
  innerAsc = Math.max(innerAsc, size * 0.6);
  innerDesc = Math.max(innerDesc, size * 0.2);

  const open = bracketGlyph(node.begChr, innerAsc, innerDesc, size);
  const close = bracketGlyph(node.endChr, innerAsc, innerDesc, size);
  const items: Array<MathDrawItem> = [...open.items];
  for (const it of shiftItems(inner, open.width, 0)) items.push(it);
  for (const it of shiftItems(close.items, open.width + innerW, 0)) items.push(it);
  const margin = 0.1 * size;
  return {
    width: open.width + innerW + close.width,
    ascent: innerAsc + margin,
    descent: innerDesc + margin,
    items,
  };
}

function layoutMatrix(node: MathMatrix, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const cells = node.rows.map((row) => row.map((cell) => layoutMath(cell, ctx, measure)));
  const nCols = Math.max(0, ...cells.map((r) => r.length));
  const colW = new Array<number>(nCols).fill(0);
  const rowAsc: Array<number> = [];
  const rowDesc: Array<number> = [];
  cells.forEach((row, ri) => {
    let a = 0;
    let d = 0;
    row.forEach((cb, ci) => {
      colW[ci] = Math.max(colW[ci]!, cb.width);
      a = Math.max(a, cb.ascent);
      d = Math.max(d, cb.descent);
    });
    rowAsc[ri] = a;
    rowDesc[ri] = d;
  });
  const colGap = 0.7 * size;
  const rowGap = 0.35 * size;
  const totalW = colW.reduce((s, w) => s + w, 0) + colGap * Math.max(0, nCols - 1);
  const rowH = rowAsc.map((a, i) => a + rowDesc[i]!);
  const totalH = rowH.reduce((s, h) => s + h, 0) + rowGap * Math.max(0, cells.length - 1);
  const axis = AXIS * size;

  const items: Array<MathDrawItem> = [];
  let y = totalH / 2 + axis; // top of the matrix
  cells.forEach((row, ri) => {
    const baseline = y - rowAsc[ri]!;
    let x = 0;
    row.forEach((cb, ci) => {
      const cellX = x + (colW[ci]! - cb.width) / 2;
      for (const it of shiftItems(cb.items, cellX, baseline)) items.push(it);
      x += colW[ci]! + colGap;
    });
    y -= rowH[ri]! + rowGap;
  });
  return { width: totalW, ascent: totalH / 2 + axis, descent: totalH / 2 - axis, items };
}

// m:eqArr — equations stacked vertically, left-aligned, the block centred on
// the math axis (like a 1-column matrix but flush-left).
function layoutEqArray(node: MathEqArray, ctx: MathCtx, measure: MeasureMath): MathBox {
  const size = ctx.sizePt;
  const boxes = node.rows.map((r) => layoutMath(r, ctx, measure));
  const width = Math.max(0, ...boxes.map((b) => b.width));
  const rowGap = 0.5 * size;
  const rowH = boxes.map((b) => b.ascent + b.descent);
  const totalH = rowH.reduce((s, h) => s + h, 0) + rowGap * Math.max(0, boxes.length - 1);
  const axis = AXIS * size;

  const items: Array<MathDrawItem> = [];
  let y = totalH / 2 + axis; // top of the block
  boxes.forEach((b, i) => {
    const baseline = y - b.ascent;
    for (const it of shiftItems(b.items, 0, baseline)) items.push(it); // flush-left
    y -= rowH[i]! + rowGap;
  });
  return { width, ascent: totalH / 2 + axis, descent: totalH / 2 - axis, items };
}

function layoutAccent(node: MathAccent, ctx: MathCtx, measure: MeasureMath): MathBox {
  const baseBox = layoutMath(node.base, ctx, measure);
  const size = ctx.sizePt;
  const y = baseBox.ascent + 0.04 * size;
  const cx = baseBox.width / 2;
  const items: Array<MathDrawItem> = [...baseBox.items];
  const accH = drawAccent(node.char, cx, y, baseBox.width, size, items);
  return { width: baseBox.width, ascent: y + accH, descent: baseBox.descent, items };
}

function drawAccent(
  char: string,
  cx: number,
  y: number,
  baseW: number,
  size: number,
  items: Array<MathDrawItem>,
): number {
  const aw = Math.min(baseW * 0.85, 0.7 * size);
  const stroke = Math.max(0.05 * size, 0.7);
  const cp = char.codePointAt(0) ?? 0;
  if (char === '^' || char === 'ˆ' || cp === 0x0302) {
    const h = 0.18 * size;
    items.push({
      kind: 'path',
      segments: [
        { op: 'move', x: cx - aw / 2, y },
        { op: 'line', x: cx, y: y + h },
        { op: 'line', x: cx + aw / 2, y },
      ],
      strokeWidthPt: stroke,
    });
    return h;
  }
  if (char === '~' || char === '˜' || cp === 0x0303) {
    const h = 0.12 * size;
    items.push({
      kind: 'path',
      segments: [
        { op: 'move', x: cx - aw / 2, y },
        {
          op: 'cubic',
          x1: cx - aw * 0.25,
          y1: y + h,
          x2: cx - aw * 0.1,
          y2: y - h * 0.4,
          x: cx,
          y: y + h * 0.3,
        },
        {
          op: 'cubic',
          x1: cx + aw * 0.1,
          y1: y + h,
          x2: cx + aw * 0.25,
          y2: y - h * 0.4,
          x: cx + aw / 2,
          y,
        },
      ],
      strokeWidthPt: stroke,
    });
    return h;
  }
  if (char === '→' || cp === 0x20d7) {
    const sh = y + 0.06 * size;
    const head = 0.06 * size;
    items.push({
      kind: 'path',
      segments: [
        { op: 'move', x: cx - aw / 2, y: sh },
        { op: 'line', x: cx + aw / 2, y: sh },
        { op: 'move', x: cx + aw / 2 - head, y: sh + head },
        { op: 'line', x: cx + aw / 2, y: sh },
        { op: 'line', x: cx + aw / 2 - head, y: sh - head },
      ],
      strokeWidthPt: stroke,
    });
    return 0.14 * size;
  }
  if (char === '.' || cp === 0x0307) {
    const d = Math.max(0.07 * size, 1);
    items.push({ kind: 'rule', x: cx - d / 2, y, w: d, h: d });
    return 0.1 * size + d;
  }
  if (cp === 0x0308) {
    const d = Math.max(0.07 * size, 1);
    const off = 0.12 * size;
    items.push({ kind: 'rule', x: cx - off - d / 2, y, w: d, h: d });
    items.push({ kind: 'rule', x: cx + off - d / 2, y, w: d, h: d });
    return 0.1 * size + d;
  }
  // bar / macron and any other accent → a horizontal rule.
  items.push({ kind: 'rule', x: cx - aw / 2, y, w: aw, h: stroke });
  return 0.08 * size;
}

function layoutBar(node: MathBar, ctx: MathCtx, measure: MeasureMath): MathBox {
  const baseBox = layoutMath(node.base, ctx, measure);
  const size = ctx.sizePt;
  const rule = Math.max(0.05 * size, 0.6);
  const gap = 0.08 * size;
  const items: Array<MathDrawItem> = [...baseBox.items];
  if (node.pos === 'top') {
    const y = baseBox.ascent + gap;
    items.push({ kind: 'rule', x: 0, y, w: baseBox.width, h: rule });
    return { width: baseBox.width, ascent: y + rule, descent: baseBox.descent, items };
  }
  const y = -baseBox.descent - gap - rule;
  items.push({ kind: 'rule', x: 0, y, w: baseBox.width, h: rule });
  return { width: baseBox.width, ascent: baseBox.ascent, descent: -y, items };
}

function layoutGroupChr(node: MathGroupChr, ctx: MathCtx, measure: MeasureMath): MathBox {
  const baseBox = layoutMath(node.base, ctx, measure);
  const size = ctx.sizePt;
  const gap = 0.06 * size;
  const braceH = 0.18 * size;
  const stroke = Math.max(0.05 * size, 0.7);
  const w = baseBox.width;
  const items: Array<MathDrawItem> = [...baseBox.items];
  // A horizontal curly brace: two cubics meeting at a central tip.
  const hbrace = (yBase: number, dir: number): Array<PathSegment> => [
    { op: 'move', x: 0, y: yBase },
    {
      op: 'cubic',
      x1: w * 0.1,
      y1: yBase,
      x2: w * 0.4,
      y2: yBase,
      x: w / 2 - 0.02 * size,
      y: yBase + dir * braceH,
    },
    { op: 'move', x: w, y: yBase },
    {
      op: 'cubic',
      x1: w * 0.9,
      y1: yBase,
      x2: w * 0.6,
      y2: yBase,
      x: w / 2 + 0.02 * size,
      y: yBase + dir * braceH,
    },
  ];
  if (node.pos === 'top') {
    const yBase = baseBox.ascent + gap + braceH;
    items.push({ kind: 'path', segments: hbrace(yBase, -1), strokeWidthPt: stroke });
    return { width: w, ascent: yBase, descent: baseBox.descent, items };
  }
  const yBase = -baseBox.descent - gap - braceH;
  items.push({ kind: 'path', segments: hbrace(yBase, 1), strokeWidthPt: stroke });
  return { width: w, ascent: baseBox.ascent, descent: -yBase, items };
}

// Translate every item in a box by (dx, dy).
export function shiftItems(
  items: ReadonlyArray<MathDrawItem>,
  dx: number,
  dy: number,
): Array<MathDrawItem> {
  return items.map((it) => {
    if (it.kind === 'glyph') return { ...it, x: it.x + dx, y: it.y + dy };
    if (it.kind === 'rule') return { ...it, x: it.x + dx, y: it.y + dy };
    return { ...it, segments: it.segments.map((s) => shiftSegment(s, dx, dy)) };
  });
}

function shiftSegment(s: PathSegment, dx: number, dy: number): PathSegment {
  switch (s.op) {
    case 'move':
    case 'line':
      return { ...s, x: s.x + dx, y: s.y + dy };
    case 'cubic':
      return {
        ...s,
        x1: s.x1 + dx,
        y1: s.y1 + dy,
        x2: s.x2 + dx,
        y2: s.y2 + dy,
        x: s.x + dx,
        y: s.y + dy,
      };
    case 'close':
      return s;
  }
}

// Flatten a math tree into (variant, text) glyph segments — used to subset the
// right glyphs into the right font variant. Grows alongside layoutMath.
export function mathGlyphSegments(node: MathNode): Array<{ variant: MathVariant; text: string }> {
  const out: Array<{ variant: MathVariant; text: string }> = [];
  const visit = (n: MathNode): void => {
    switch (n.type) {
      case 'run': {
        const chars = [...n.text];
        let i = 0;
        while (i < chars.length) {
          const v = runCharVariant(chars[i]!, n);
          let seg = '';
          let j = i;
          while (j < chars.length && runCharVariant(chars[j]!, n) === v) {
            seg += chars[j];
            j++;
          }
          out.push({ variant: v, text: seg });
          i = j;
        }
        break;
      }
      case 'row':
        n.children.forEach(visit);
        break;
      case 'fraction':
        visit(n.num);
        visit(n.den);
        break;
      case 'script':
        visit(n.base);
        if (n.sub) visit(n.sub);
        if (n.sup) visit(n.sup);
        break;
      case 'radical':
        visit(n.radicand);
        if (n.degree) visit(n.degree);
        break;
      case 'nary':
        out.push({ variant: 'regular', text: n.op });
        visit(n.body);
        if (n.sub) visit(n.sub);
        if (n.sup) visit(n.sup);
        break;
      case 'func':
        visit(n.name);
        visit(n.body);
        break;
      case 'limit':
        visit(n.base);
        visit(n.lim);
        break;
      case 'delimiter':
        if (n.sepChr) out.push({ variant: 'regular', text: n.sepChr });
        n.children.forEach(visit);
        break;
      case 'matrix':
        n.rows.forEach((row) => row.forEach(visit));
        break;
      case 'eqArr':
        n.rows.forEach(visit);
        break;
      case 'accent':
      case 'bar':
      case 'groupChr':
        visit(n.base);
        break;
      default:
        break;
    }
  };
  visit(node);
  return out;
}
