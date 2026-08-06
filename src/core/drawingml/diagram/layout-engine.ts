// ECMA-376 §21.4.3 — a SmartArt LAYOUT part (`diagrams/layout1.xml`) is not a
// preset name, it is a small program: `layoutNode`s that pick an algorithm,
// `forEach`es that walk an axis of the data, `choose`/`if`/`else` on variables,
// and a `constrLst` that states each box's size in terms of another's.
// PowerPoint runs it and caches the result in a drawing part; a file written by
// a generator carries no drawing part, and then this has to run it.
//
// This is the first slice: the list layouts, whose whole program is "put the
// child nodes in a row or a column, with a gap between them, and set the text
// inside each". That is the `snake` / `lin` algorithms with `sp` for the gaps
// and `tx` for the text, and it covers the `default` layout — the Basic Block
// List every "insert SmartArt" starts from — plus the vertical and horizontal
// list families. The hierarchy and cycle algorithms are a separate slice.

import type { PoNode } from '@/core/po-helpers';
import type { DiagramData, DiagramPoint } from '@/core/drawingml/diagram/data-model';
import { poAttr, poChildren, poIs, poTag } from '@/core/po-helpers';

/** One box the engine laid out, in the diagram's own frame (EMU). */
export interface LaidNode {
  readonly point: DiagramPoint;
  readonly shapeType: string;
  /** The layout node's `styleLbl` — which run of colours the box takes. */
  readonly styleLbl?: string;
  /** The box's place among its siblings, which is where in that run it is. */
  readonly index: number;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/** A constraint from a `dgm:constrLst`, as the engine needs it. */
interface Constraint {
  readonly type: string;
  readonly for?: string;
  readonly forName?: string;
  readonly refType?: string;
  readonly refForName?: string;
  readonly fact: number;
  readonly val?: number;
}

/** The algorithms this slice runs. */
type AlgType = 'snake' | 'lin' | 'sp' | 'tx' | 'composite' | 'other';

interface LayoutNode {
  readonly name: string;
  readonly styleLbl?: string;
  /**
   * Whether the layout only lays this box out for a point that HAS children —
   * a `choose` on `func="cnt" axis="ch" ptType="node"`. The vertical block list
   * guards its descendants box that way, and a node with nothing under it gets
   * its label alone rather than an empty box beside it.
   */
  readonly needsChildren: boolean;
  readonly alg: AlgType;
  readonly algParams: ReadonlyMap<string, string>;
  readonly shapeType?: string;
  readonly constrs: ReadonlyArray<Constraint>;
  readonly children: ReadonlyArray<LayoutItem>;
}
interface ForEach {
  readonly kind: 'forEach';
  readonly axis: string;
  readonly ptType: string;
  readonly children: ReadonlyArray<LayoutItem>;
}
type LayoutItem = ({ kind: 'node' } & LayoutNode) | ForEach;

/**
 * Lay a diagram out from its data and layout parts.
 *
 * @param layout The parsed `layout1.xml` tree.
 * @param data   The parsed data part.
 * @param cx     The frame's width in EMU, as the drawing states it.
 * @param cy     The frame's height in EMU.
 * @returns The boxes to draw, or an empty list when the layout uses an
 *          algorithm this slice does not run — the caller then reports the
 *          diagram as a loss exactly as before.
 */
export function layoutDiagram(
  layout: ReadonlyArray<PoNode>,
  data: DiagramData,
  cx: number,
  cy: number,
): Array<LaidNode> {
  const root = findChild(layout, 'dgm:layoutDef');
  const top = root ? findChild(poChildren(root), 'dgm:layoutNode') : undefined;
  if (!top || !data.root) return [];
  const parsed = parseNode(top, { maxDepth: data.depth });
  // Only the flat list algorithms are run here; anything else is left to the
  // caller's graceful loss rather than laid out wrongly.
  if (parsed.alg !== 'snake' && parsed.alg !== 'lin') return [];

  const nodes = data.children(data.root.id, 'node');
  if (nodes.length === 0) return [];

  // The child layoutNode's own name is what the constraints refer to.
  const childName = firstNodeName(parsed.children) ?? 'node';
  const c = resolve(parsed.constrs, childName);

  // §21.4.3.2 — `snake` flows the children across and WRAPS; `lin` is the one
  // that does not. The number of columns is the one that leaves each cell
  // closest to the shape the constraints ask for (`h` stated against `w`),
  // which is what "fill the frame" means: five nodes in a 4:3 frame come out
  // two across and three down, as every reader draws them.
  const shapeType = childShapeType(parsed.children) ?? 'rect';
  const vertical = isVertical(parsed);
  const n = nodes.length;
  // A column list is one box wide however much room there is beside it, and a
  // row that does not snake is one box tall: only a snake across chooses.
  const cols = vertical ? 1 : parsed.alg === 'snake' ? bestColumns(n, cx, cy, c.aspect, c.gap) : n;
  const rows = Math.ceil(n / cols);

  // Each track is a cell plus the gap that follows it, the last gap trimmed.
  const cellW = cx / (cols + Math.max(0, cols - 1) * c.gap);
  const gapW = cellW * c.gap;
  const cellH = cy / (rows + Math.max(0, rows - 1) * c.gap);
  const gapH = cellH * c.gap;

  const out: Array<LaidNode> = [];
  nodes.forEach((point, i) => {
    const col = vertical ? Math.floor(i / rows) : i % cols;
    const row = vertical ? i % rows : Math.floor(i / cols);
    // A last row that is short is centred under the ones above it, which is
    // where PowerPoint leaves it.
    const inRow = vertical ? rows : Math.min(cols, n - row * cols);
    const rowWidth = inRow * cellW + Math.max(0, inRow - 1) * gapW;
    const left = vertical ? 0 : (cx - rowWidth) / 2;
    const cell = {
      x: left + col * (cellW + gapW),
      y: row * (cellH + gapH),
      cx: cellW,
      cy: cellH,
    };
    // §21.4.3 — the child layoutNode may not be a box at all: a list that
    // shows a node's label beside its children's text nests a `linNode` whose
    // own children are two `tx` boxes, sized by ITS constraints. Laid out as
    // one box, such a row is the whole cell where it should be two.
    const inner = splitCell(parsed, childName, point, data, cell, shapeType, i);
    out.push(...inner);
  });
  return out;
}

/**
 * One data node's cell, split the way the child layoutNode divides it.
 *
 * A leaf child (`tx` on the node itself) takes the cell whole. A child that
 * nests further hands its own children a share of the cell, and the two
 * families that do differ in which way they cut it:
 *
 * - the `linNode` of the vertical block lists states `w` fractions for its
 *   children (0.36 for the label, 0.64 for the descendants' text) — a row;
 * - the `composite` of the picture and colour lists gives both children the
 *   full width and states their `h` against a font size (0.8 against 1.22) —
 *   a column.
 *
 * Laid out as one box either way, such a cell holds one text where it should
 * hold the node's label and its children's text in boxes of their own.
 */
function splitCell(
  parent: { kind: 'node' } & LayoutNode,
  childName: string,
  point: DiagramPoint,
  data: DiagramData,
  cell: { x: number; y: number; cx: number; cy: number },
  shapeType: string,
  index: number,
): Array<LaidNode> {
  const child = findNamed(parent.children, childName);
  const inner = child ? namedBoxes(child) : [];
  const whole = (lbl: string | undefined): Array<LaidNode> => [
    { point, shapeType, ...(lbl !== undefined ? { styleLbl: lbl } : {}), index, ...cell },
  ];
  if (!child || inner.length !== 2) return whole(child?.styleLbl);
  const kids = data.children(point.id, 'node');

  const widths = inner.map((b) => sizeFact(child.constrs, b.name, 'w'));
  // A child given less than the whole width sits BESIDE its sibling; one given
  // all of it sits above, and then the heights are what divide the cell.
  const across = widths.some((v) => v !== undefined && v < 0.99);
  const shares = across
    ? fractions(widths)
    : fractions(
        inner.map((b) => sizeFact(child.constrs, b.name, 'h') ?? heightFact(parent, b.name)),
      );

  const out: Array<LaidNode> = [];
  let at = across ? cell.x : cell.y;
  inner.forEach((box, i) => {
    const span = (across ? cell.cx : cell.cy) * (shares[i] ?? 0.5);
    // The first box is the node's own label; the second holds its children's,
    // and is left out when the layout guards it and there are none.
    if (span > 0 && !(box.needsChildren && kids.length === 0)) {
      out.push({
        point: i === 0 ? point : gatheredText(point, kids),
        shapeType: i === 0 ? shapeType : (box.shapeType ?? 'rect'),
        ...(box.styleLbl !== undefined ? { styleLbl: box.styleLbl } : {}),
        index,
        x: across ? at : cell.x,
        y: across ? cell.y : at,
        cx: across ? span : cell.cx,
        cy: across ? cell.cy : span,
      });
    }
    at += span;
  });
  return out.length > 0 ? out : whole(child.styleLbl);
}

// Stated sizes as fractions of the whole: what is stated is kept, and what is
// not divides the remainder. Sizes stated against a font rather than the parent
// (`0.8` and `1.22`) are proportions of each other, so they are scaled to fit.
function fractions(facts: ReadonlyArray<number | undefined>): Array<number> {
  const total = facts.reduce<number>((a, b) => a + (b ?? 0), 0);
  if (total > 1.001) return facts.map((v) => (v ?? 0) / total);
  const unstated = facts.filter((v) => v === undefined).length;
  const rest = Math.max(0, 1 - total) / Math.max(1, unstated);
  return facts.map((v) => v ?? rest);
}

// The children's text as one point, since a list shows them together in the
// second box rather than one box apiece.
function gatheredText(parent: DiagramPoint, kids: ReadonlyArray<DiagramPoint>): DiagramPoint {
  const paragraphs = kids.flatMap((k) => (k.text ? poChildren(k.text) : []));
  // With no children the box is still laid out — the layout states it — but it
  // is empty. Falling back to the parent here printed every label twice.
  if (paragraphs.length === 0) return { id: `${parent.id}/des`, type: 'node' };
  return { id: `${parent.id}/des`, type: 'node', text: { 'dgm:t': paragraphs } };
}

// How many columns a snake takes: the count whose cell comes closest to the
// aspect the constraints state for a node.
function bestColumns(n: number, cx: number, cy: number, aspect: number, gap: number): number {
  if (n <= 1) return 1;
  let best = n;
  let bestErr = Infinity;
  for (let cols = 1; cols <= n; cols++) {
    const rows = Math.ceil(n / cols);
    const w = cx / (cols + Math.max(0, cols - 1) * gap);
    const h = cy / (rows + Math.max(0, rows - 1) * gap);
    const err = Math.abs(Math.log(h / w / (aspect || 1)));
    if (err < bestErr - 1e-9) {
      bestErr = err;
      best = cols;
    }
  }
  return best;
}

// The layoutNode with this name, anywhere below.
function findNamed(
  items: ReadonlyArray<LayoutItem>,
  name: string,
): ({ kind: 'node' } & LayoutNode) | undefined {
  for (const item of items) {
    if (item.kind === 'node') {
      if (item.name === name) return item;
      const inner = findNamed(item.children, name);
      if (inner) return inner;
    } else {
      const inner = findNamed(item.children, name);
      if (inner) return inner;
    }
  }
  return undefined;
}

// The boxes a node lays out directly inside itself.
function namedBoxes(node: { kind: 'node' } & LayoutNode): Array<{ kind: 'node' } & LayoutNode> {
  const out: Array<{ kind: 'node' } & LayoutNode> = [];
  for (const item of node.children) {
    if (item.kind === 'node') out.push(item);
    else for (const inner of item.children) if (inner.kind === 'node') out.push(inner);
  }
  return out;
}

// The size a node's constraints give one of its children, as a fraction of its
// own — a size stated against a SIBLING is a copy of that sibling, not a share
// of the parent, so it is left for the remainder to fill.
function sizeFact(
  constrs: ReadonlyArray<Constraint>,
  name: string,
  type: 'w' | 'h',
): number | undefined {
  for (const k of constrs) {
    if (k.type === type && k.forName === name && k.refType === type && k.refForName === undefined) {
      return k.fact;
    }
  }
  return undefined;
}

// A height the nested node does not state itself is stated for it further up,
// against a font size: `parTx` at 0.8 of it and `desTx` at 1.22 are a ratio
// between the two, which is all a split of the cell needs.
function heightFact(parent: { kind: 'node' } & LayoutNode, name: string): number | undefined {
  for (const k of parent.constrs) {
    if (k.type === 'h' && k.forName === name && k.refType === 'primFontSz') return k.fact;
  }
  return undefined;
}

// The child node's aspect and the gap between siblings, read off the top
// node's constraints: `h` of a child stated against its own `w` gives the
// aspect, and a `sibTrans` width stated the same way gives the gap. The gap is
// signed: a chevron process states `-0.1`, which is how the point of one
// chevron comes to sit in the notch of the next.
function resolve(
  constrs: ReadonlyArray<Constraint>,
  childName: string,
): { aspect: number; gap: number } {
  let aspect = 0.6;
  let gap = 0.1;
  for (const k of constrs) {
    if (k.forName === childName && k.type === 'h' && k.refType === 'w') aspect = k.fact;
    if (k.type === 'w' && k.refForName === childName && k.forName !== childName) gap = k.fact;
  }
  return { aspect, gap };
}

// `grDir`/`flowDir` say which way a snake runs; a column layout is the one
// whose flow is `col`, or whose growth starts down rather than across.
function isVertical(node: LayoutNode): boolean {
  const flow = node.algParams.get('flowDir');
  if (flow === 'col') return true;
  const dir = node.algParams.get('linDir');
  return dir === 'fromT' || dir === 'fromB';
}

function firstNodeName(items: ReadonlyArray<LayoutItem>): string | undefined {
  for (const item of items) {
    if (item.kind === 'node') return item.name;
    const inner = firstNodeName(item.children);
    if (inner !== undefined) return inner;
  }
  return undefined;
}

function childShapeType(items: ReadonlyArray<LayoutItem>): string | undefined {
  for (const item of items) {
    if (item.kind === 'node') return item.shapeType ?? childShapeType(item.children);
    const inner = childShapeType(item.children);
    if (inner !== undefined) return inner;
  }
  return undefined;
}

/** What a `choose` is evaluated against: the shape of the data it will lay out. */
interface LayoutContext {
  readonly maxDepth: number;
}

function parseNode(node: PoNode, ctx: LayoutContext): { kind: 'node' } & LayoutNode {
  const kids = poChildren(node);
  // §21.4.3.3 — a `choose` picks one branch by a function of the data. The
  // direction variables are left at their defaults, so those take the first
  // `if`; `maxDepth` is answered from the model, because the branch it guards
  // is the difference between one box in a cell and two.
  const flattened: Array<PoNode> = [];
  // Branches taken from a `choose` on how many children a point has: what they
  // lay out exists only for a point that has them.
  const guarded = new Set<PoNode>();
  for (const k of kids) {
    if (poIs(k, 'dgm:choose')) {
      const branches = poChildren(k);
      const taken =
        branches.find((b) => poIs(b, 'dgm:if') && holds(b, ctx)) ??
        branches.find((b) => poIs(b, 'dgm:else')) ??
        branches[0];
      const onCount = taken !== undefined && poIs(taken, 'dgm:if') && countsChildren(taken);
      for (const inner of poChildren(taken)) {
        flattened.push(inner);
        if (onCount) guarded.add(inner);
      }
    } else flattened.push(k);
  }
  const alg = flattened.find((k) => poIs(k, 'dgm:alg'));
  const shape = flattened.find((k) => poIs(k, 'dgm:shape'));
  const constrLst = flattened.find((k) => poIs(k, 'dgm:constrLst'));
  const params = new Map<string, string>();
  for (const p of alg ? poChildren(alg) : []) {
    const t = poAttr(p, 'type');
    const v = poAttr(p, 'val');
    if (t !== undefined && v !== undefined) params.set(t, v);
  }
  const shapeType = shape ? poAttr(shape, 'type') : undefined;
  const styleLbl = poAttr(node, 'styleLbl');
  return {
    kind: 'node',
    name: poAttr(node, 'name') ?? '',
    ...(styleLbl !== undefined && styleLbl !== '' ? { styleLbl } : {}),
    needsChildren: false,
    alg: algOf(alg),
    algParams: params,
    ...(shapeType !== undefined && shapeType !== '' ? { shapeType } : {}),
    constrs: parseConstraints(constrLst),
    children: flattened.flatMap((k) => parseItem(k, ctx, guarded.has(k))),
  };
}

// Whether a `dgm:if` holds. Only `maxDepth` is answered — every other test is
// on a variable this engine leaves at its default, where the `if` is the branch
// PowerPoint takes.
function holds(branch: PoNode, ctx: LayoutContext): boolean {
  if (poAttr(branch, 'func') !== 'maxDepth') return true;
  const val = Number(poAttr(branch, 'val') ?? '0');
  const op = poAttr(branch, 'op') ?? 'gte';
  if (op === 'gte') return ctx.maxDepth >= val;
  if (op === 'gt') return ctx.maxDepth > val;
  if (op === 'lte') return ctx.maxDepth <= val;
  if (op === 'lt') return ctx.maxDepth < val;
  if (op === 'neq') return ctx.maxDepth !== val;
  return ctx.maxDepth === val;
}

// Whether a branch tests a point's child count — `func="cnt" axis="ch"
// ptType="node" op="gte" val="1"`, the guard on a descendants box.
function countsChildren(branch: PoNode): boolean {
  return (
    poAttr(branch, 'func') === 'cnt' &&
    poAttr(branch, 'axis') === 'ch' &&
    poAttr(branch, 'ptType') === 'node' &&
    (poAttr(branch, 'op') ?? 'gte') === 'gte' &&
    Number(poAttr(branch, 'val') ?? '1') >= 1
  );
}

function parseItem(node: PoNode, ctx: LayoutContext, guarded = false): Array<LayoutItem> {
  if (poIs(node, 'dgm:layoutNode')) return [{ ...parseNode(node, ctx), needsChildren: guarded }];
  if (poIs(node, 'dgm:forEach')) {
    return [
      {
        kind: 'forEach',
        axis: poAttr(node, 'axis') ?? '',
        ptType: poAttr(node, 'ptType') ?? '',
        children: poChildren(node).flatMap((k) => parseItem(k, ctx, guarded)),
      },
    ];
  }
  return [];
}

function parseConstraints(list: PoNode | undefined): Array<Constraint> {
  const out: Array<Constraint> = [];
  for (const k of list ? poChildren(list) : []) {
    if (!poIs(k, 'dgm:constr')) continue;
    const type = poAttr(k, 'type');
    if (type === undefined) continue;
    const forName = poAttr(k, 'forName');
    const refType = poAttr(k, 'refType');
    const refForName = poAttr(k, 'refForName');
    const val = poAttr(k, 'val');
    out.push({
      type,
      ...(forName !== undefined ? { forName } : {}),
      ...(refType !== undefined ? { refType } : {}),
      ...(refForName !== undefined ? { refForName } : {}),
      fact: Number(poAttr(k, 'fact') ?? '1'),
      ...(val !== undefined ? { val: Number(val) } : {}),
    });
  }
  return out;
}

function algOf(alg: PoNode | undefined): AlgType {
  const t = alg ? poAttr(alg, 'type') : undefined;
  return t === 'snake' || t === 'lin' || t === 'sp' || t === 'tx' || t === 'composite'
    ? t
    : 'other';
}

function findChild(nodes: ReadonlyArray<PoNode>, tag: string): PoNode | undefined {
  return nodes.find((n) => poIs(n, tag) || poTag(n) === tag);
}
