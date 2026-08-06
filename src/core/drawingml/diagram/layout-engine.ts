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
  const parsed = parseNode(top);
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
    out.push({
      point,
      shapeType,
      x: left + col * (cellW + gapW),
      y: row * (cellH + gapH),
      cx: cellW,
      cy: cellH,
    });
  });
  return out;
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

// The child node's aspect and the gap between siblings, read off the top
// node's constraints: `h` of a child stated against its own `w` gives the
// aspect, and a `sibTrans` width stated the same way gives the gap.
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

function parseNode(node: PoNode): { kind: 'node' } & LayoutNode {
  const kids = poChildren(node);
  // §21.4.3.3 — a `choose` picks one branch by a variable; with no variables
  // set the first `if` is what PowerPoint takes, which is the `norm` direction.
  const flattened: Array<PoNode> = [];
  for (const k of kids) {
    if (poIs(k, 'dgm:choose')) {
      const first = poChildren(k).find((b) => poIs(b, 'dgm:if')) ?? poChildren(k)[0];
      if (first) flattened.push(...poChildren(first));
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
  return {
    kind: 'node',
    name: poAttr(node, 'name') ?? '',
    alg: algOf(alg),
    algParams: params,
    ...(shapeType !== undefined && shapeType !== '' ? { shapeType } : {}),
    constrs: parseConstraints(constrLst),
    children: flattened.flatMap(parseItem),
  };
}

function parseItem(node: PoNode): Array<LayoutItem> {
  if (poIs(node, 'dgm:layoutNode')) return [parseNode(node)];
  if (poIs(node, 'dgm:forEach')) {
    return [
      {
        kind: 'forEach',
        axis: poAttr(node, 'axis') ?? '',
        ptType: poAttr(node, 'ptType') ?? '',
        children: poChildren(node).flatMap(parseItem),
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
