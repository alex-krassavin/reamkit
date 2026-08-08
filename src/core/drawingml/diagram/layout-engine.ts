// ECMA-376 §21.4.3 — a SmartArt LAYOUT part (`diagrams/layout1.xml`) is not a
// preset name, it is a small program: `layoutNode`s that pick an algorithm,
// `forEach`es that walk an axis of the data, `choose`/`if`/`else` on variables,
// and a `constrLst` that states each box's size in terms of another's.
// PowerPoint runs it and caches the result in a drawing part; a file written by
// a generator carries no drawing part, and then this has to run it.
//
// It began with the list layouts, whose whole program is "put the child nodes
// in a row or a column, with a gap between them, and set the text inside each"
// — the `snake` / `lin` algorithms with `sp` for the gaps and `tx` for the
// text, covering the `default` layout that every "insert SmartArt" starts from
// plus the vertical and horizontal list families. `composite` places boxes
// rather than sharing a cell between them, `cycle` rings them, `conn` draws
// what joins two of them, and `hierRoot`/`hierChild` are the org chart. What is
// left — the pyramids and the matrices — divides a frame by AREA, and a layout
// naming one of those is still left to the caller's graceful loss.

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
  /** The point size the layout states for the box's text, when it states one. */
  readonly fontSizePt?: number;
  /** Whether the `tx` algorithm asks for the text to be bulleted. */
  readonly bulleted?: boolean;
  /**
   * §21.4.3.9 `dgm:shape@hideGeom` — the box reserves its room but its outline
   * is never drawn. The picture lists put an invisible spacer above each node.
   */
  readonly hideGeom?: boolean;
  /** §21.4.3.2 `txAnchorVert` — where in its box the `tx` algorithm sits text. */
  readonly anchor?: string;
  /** §21.4.3.2 `parTxLTRAlign` — how the layout ranges a node's own label. */
  readonly align?: string;
  /**
   * §21.4.3.5 `dgm:shape@rot` — how far the box's geometry is turned, in
   * 60 000ths of a degree clockwise. The cycle-matrix layout builds its disc
   * out of one quarter-wedge stated four times at 0°, 90°, 180° and 270°.
   */
  readonly rotation60k?: number;
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
  /** `@op` — `lte`/`gte` make the constraint a BOUND rather than a value. */
  readonly op?: string;
  readonly refType?: string;
  readonly refForName?: string;
  readonly fact: number;
  readonly val?: number;
}

/** The algorithms this slice runs. */
type AlgType = 'snake' | 'lin' | 'sp' | 'tx' | 'composite' | 'cycle' | 'hier' | 'other';

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
  readonly hideGeom?: boolean;
  /**
   * §21.4.3.7 `dgm:presOf@axis="desOrSelf"` — the box speaks for the point AND
   * everything under it, so a node with children shows its own label with theirs
   * beneath it rather than dropping them.
   */
  readonly presOfAxis?: string;
  /**
   * §21.4.3.7 — the same `presOf` when it names a PATH rather than one axis:
   * `axis="ch des" ptType="node node" st="1 1" cnt="1 0"` is "the first child,
   * then everything under it". Each step takes an axis of the points it has so
   * far and then the slice `st`/`cnt` names — both 1-based, and a count of 0
   * meaning all of them. This is how the cycle-matrix layout gives its four
   * quadrants one child each and its four callouts that child's children.
   */
  readonly presOfPath?: {
    readonly axis: ReadonlyArray<string>;
    readonly st: ReadonlyArray<number>;
    readonly cnt: ReadonlyArray<number>;
  };
  /** §21.4.3.5 `dgm:shape@rot` — degrees clockwise the box's geometry is turned. */
  readonly shapeRotation?: number;
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
  const parsed = parseNode(top, { maxDepth: data.depth, data, root: data.root });
  // Only the algorithms below are run; anything else is left to the caller's
  // graceful loss rather than laid out wrongly.
  if (
    parsed.alg !== 'snake' &&
    parsed.alg !== 'lin' &&
    parsed.alg !== 'composite' &&
    parsed.alg !== 'cycle' &&
    parsed.alg !== 'hier'
  ) {
    return [];
  }
  return layoutIn(parsed, data.root, { x: 0, y: 0, cx, cy }, data, 0);
}

/** A rectangle in the diagram's own frame (EMU). */
interface Box {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/**
 * One layout node run inside the box it was given, by whichever algorithm it
 * names.
 *
 * @param node  The layout node.
 * @param point The data point it is laying out.
 * @param box   The rectangle it has to fill.
 * @param data  The data part.
 * @param index The point's place among its siblings, for the colour run.
 */
function layoutIn(
  node: { kind: 'node' } & LayoutNode,
  point: DiagramPoint,
  box: Box,
  data: DiagramData,
  index: number,
): Array<LaidNode> {
  if (node.alg === 'composite') return compositeLayout(node, point, box, data, index);
  if (node.alg === 'hier') return hierarchyLayout(node, point, box, data);
  if (node.alg === 'cycle') return cycleLayout(node, point, box, data);
  if (node.alg === 'lin' || node.alg === 'snake') return listLayout(node, point, box, data);
  // §21.4.3.2 `sp` is the space algorithm — it reserves room. It reserves it
  // for a SHAPE when the node names one, which is how a process arrow gets
  // drawn behind the steps standing on it; naming none, it is just the gap.
  if (node.alg === 'sp') {
    // A space box shows no words, but it is still the point's box: whatever
    // formatting the author set on that point is this box's too. Dropped, the
    // orange outline a cycle-matrix's second callout was given by hand was
    // drawn on the text box behind it and painted over by the plain one.
    const own = spokenFor(node, point, data).spPr;
    return node.shapeType === undefined
      ? []
      : [
          {
            point: {
              id: `${point.id}/${node.name}`,
              type: 'node',
              ...(own !== undefined ? { spPr: own } : {}),
            },
            shapeType: node.shapeType,
            ...styleOf(node),
            index,
            ...box,
          },
        ];
  }
  // `tx` is the leaf that sets a point's own words. Anything else — a cycle, a
  // hierarchy, a pyramid — is an algorithm this engine does not run, and a box
  // filling its whole share is a worse answer than none: a centre-cycle came
  // out as one blue rectangle over the entire frame.
  if (node.alg !== 'tx') return [];
  const spoken = spokenFor(node, point, data);
  return [{ point: spoken, shapeType: node.shapeType ?? 'rect', ...styleOf(node), index, ...box }];
}

/** One node of the data tree, with the room its whole subtree needs. */
interface Branch {
  readonly point: DiagramPoint;
  readonly kids: ReadonlyArray<Branch>;
  /** The subtree's width, in box widths. */
  readonly units: number;
  /** What the children take between them, which is less when the box is wider. */
  readonly spread: number;
}

/**
 * §21.4.3.2 `hierRoot`/`hierChild` — the hierarchy family, which is the org
 * chart. `hierRoot` puts a node's box above its subtree and `hierChild` ranges
 * that subtree beneath it, each child being a `hierRoot` again; the two are one
 * tree walk written as two algorithms, and the engine runs them as one.
 *
 * The layout program that expresses this is deeply nested — a `choose` per
 * branch style at every level, with the same three boxes renamed each time —
 * but it says only one thing, and it says it in the constraints at the top: a
 * box is `w` wide and `0.5w` tall, levels are `0.21w` apart and so are
 * siblings. So the geometry is read from there and the shape of the tree from
 * the DATA, which is where it actually lives.
 *
 * @param node  The `hierChild`/`hierRoot` node at the top of the hierarchy.
 * @param point The data point it is rooted at.
 * @param box   The frame the tree has to fit.
 * @param data  The data part.
 */
function hierarchyLayout(
  node: { kind: 'node' } & LayoutNode,
  point: DiagramPoint,
  box: Box,
  data: DiagramData,
): Array<LaidNode> {
  const f = hierFactors(node.constrs);
  const face = firstTextBox(node);

  const build = (p: DiagramPoint): Branch => {
    const kids = data.children(p.id, 'node').map(build);
    const spread =
      kids.length === 0 ? 0 : kids.reduce((a, k) => a + k.units, 0) + f.sibSp * (kids.length - 1);
    return { point: p, kids, units: Math.max(1, spread), spread };
  };
  // The point handed in is the CONTAINER, as it is for every other algorithm
  // here: the roots of the tree are its children. Taking it for the tree's own
  // root drew the document itself as an empty box above the chart.
  const roots = data.children(point.id, 'node').map(build);
  if (roots.length === 0) return [];
  const wide = roots.reduce((a, r) => a + r.units, 0) + f.sibSp * (roots.length - 1);
  const rows = (t: Branch): number => 1 + t.kids.reduce((a, k) => Math.max(a, rows(k)), 0);
  const deep = roots.reduce((a, r) => Math.max(a, rows(r)), 0);
  // A row is a box plus the gap below it, the last gap trimmed.
  const tall = deep * f.aspect + (deep - 1) * f.sp;
  if (wide <= 0 || tall <= 0) return [];
  // One box's width: whichever of the two directions runs out first.
  const w = Math.min(box.cx / wide, box.cy / tall);

  const out: Array<LaidNode> = [];
  const top = box.y + (box.cy - tall * w) / 2;
  const place = (t: Branch, left: number, depth: number): void => {
    out.push({
      point: t.point,
      shapeType: face?.shapeType ?? 'rect',
      // §21.4.5 — a hierarchy's colours part names one label per level.
      styleLbl: `node${Math.min(depth, 4)}`,
      ...(face === undefined ? {} : textPlacing(face)),
      index: depth,
      // The box is centred over everything below it.
      x: left + (t.units * w - w) / 2,
      y: top + depth * (f.aspect + f.sp) * w,
      cx: w,
      cy: f.aspect * w,
    });
    let at = left + ((t.units - t.spread) * w) / 2;
    for (const kid of t.kids) {
      place(kid, at, depth + 1);
      at += (kid.units + f.sibSp) * w;
    }
  };
  let at = box.x + (box.cx - wide * w) / 2;
  for (const root of roots) {
    place(root, at, 0);
    at += (root.units + f.sibSp) * w;
  }
  return out;
}

// The three numbers a hierarchy is drawn from, each a multiple of one box's
// width: how tall a box is, how far apart two levels are, and how far apart two
// siblings are. The defaults are the org chart's own, for a layout that leaves
// one of them to the reader.
function hierFactors(constrs: ReadonlyArray<Constraint>): {
  aspect: number;
  sp: number;
  sibSp: number;
} {
  let aspect = 0.5;
  let sp = 0.21;
  let sibSp = 0.21;
  for (const k of constrs) {
    if (k.refType !== 'w') continue;
    // A height stated against another box's WIDTH is the box's own shape.
    if (k.type === 'h' && k.refForName !== undefined) aspect = k.fact;
    if (k.type === 'sp') sp = k.fact;
    if (k.type === 'sibSp') sibSp = k.fact;
  }
  return { aspect, sp, sibSp };
}

// The box a hierarchy draws for each of its nodes: the first `tx` in a program
// that renames the same three boxes at every level.
function firstTextBox(
  node: { kind: 'node' } & LayoutNode,
): ({ kind: 'node' } & LayoutNode) | undefined {
  const walk = (items: ReadonlyArray<LayoutItem>): ({ kind: 'node' } & LayoutNode) | undefined => {
    for (const item of items) {
      if (item.kind === 'node' && item.alg === 'tx') return item;
      const found = walk(item.children);
      if (found) return found;
    }
    return undefined;
  };
  return walk(node.children);
}

// §21.4.3.2 — where the `tx` algorithm puts the words inside the box it got.
function textPlacing(node: { kind: 'node' } & LayoutNode): { anchor?: string; align?: string } {
  const anchor = node.algParams.get('txAnchorVert');
  const align = node.algParams.get('parTxLTRAlign');
  return {
    ...(anchor !== undefined && anchor !== '' ? { anchor } : {}),
    ...(align !== undefined && align !== '' ? { align } : {}),
  };
}

// What a box takes from its own layout node, whatever laid it out.
function styleOf(node: { kind: 'node' } & LayoutNode): {
  styleLbl?: string;
  fontSizePt?: number;
  bulleted?: true;
  hideGeom?: true;
  anchor?: string;
  align?: string;
  rotation60k?: number;
} {
  const sz = fontSize(node.constrs, undefined, 'primFontSz');
  const anchor = node.algParams.get('txAnchorVert');
  const align = node.algParams.get('parTxLTRAlign');
  return {
    ...(node.styleLbl !== undefined ? { styleLbl: node.styleLbl } : {}),
    ...opt('fontSizePt', sz),
    ...(bulleted(node) ? { bulleted: true as const } : {}),
    ...(node.hideGeom === true ? { hideGeom: true as const } : {}),
    ...(anchor !== undefined && anchor !== '' ? { anchor } : {}),
    ...(align !== undefined && align !== '' ? { align } : {}),
    ...(node.shapeRotation !== undefined ? { rotation60k: node.shapeRotation * 60000 } : {}),
  };
}

/**
 * §21.4.3.7 `dgm:presOf` — which points a box speaks for. `desOrSelf` is the
 * point and everything under it; `des` is everything under it and NOT the point,
 * which is how a list's second box holds the children's words and not a second
 * copy of the label. Anything else is the point itself.
 */
function spokenFor(
  box: { kind: 'node' } & LayoutNode,
  point: DiagramPoint,
  data: DiagramData,
): DiagramPoint {
  if (box.presOfPath) return walkPath(box.presOfPath, point, data);
  if (box.presOfAxis === 'desOrSelf') return withDescendants(point, data);
  if (box.presOfAxis === 'des') return gatheredText(point, data.children(point.id, 'node'));
  return point;
}

/**
 * §21.4.3.7 — the `st`/`cnt` slice of a `presOf` path, read from the point the
 * box is being laid out for.
 *
 * The cycle-matrix layout is the reason this exists: each of its four quadrants
 * names ONE child (`axis="ch" st="3" cnt="1"`) and each of the four callouts
 * beside them names that child's own children (`axis="ch des" st="3 1"
 * cnt="1 0"`). Read as a plain axis, all eight boxes spoke for the root and the
 * whole diagram came out blank.
 *
 * @param path  The parsed axis / start / count triple.
 * @param point The point the box is laid out for.
 * @param data  The data part, for walking children.
 * @returns One point carrying the text of everything the path reached.
 */
function walkPath(
  path: NonNullable<LayoutNode['presOfPath']>,
  point: DiagramPoint,
  data: DiagramData,
): DiagramPoint {
  let at: Array<DiagramPoint> = [point];
  path.axis.forEach((step, i) => {
    const next =
      step === 'ch'
        ? at.flatMap((p) => data.children(p.id, 'node'))
        : step === 'des'
          ? at.flatMap((p) => descendants(p, data))
          : step === 'self'
            ? at
            : [];
    const from = Math.max(0, (path.st[i] ?? 1) - 1);
    const want = path.cnt[i] ?? 0;
    at = next.slice(from, want === 0 ? next.length : from + want);
  });
  // One point is that point, not a copy of its words: it carries the formatting
  // the author set on it, and a synthetic stand-in would drop it.
  if (at.length === 1) return at[0]!;
  const paragraphs = at.flatMap((p) => (p.text ? poChildren(p.text) : []));
  const id = `${point.id}/${path.axis.join('.')}`;
  if (paragraphs.length === 0) return { id, type: 'node' };
  return { id, type: 'node', text: { 'dgm:t': paragraphs } };
}

/** Every point under `point`, depth first — the `des` axis. */
function descendants(point: DiagramPoint, data: DiagramData): Array<DiagramPoint> {
  const out: Array<DiagramPoint> = [];
  const walk = (id: string): void => {
    for (const kid of data.children(id, 'node')) {
      out.push(kid);
      walk(kid.id);
    }
  };
  walk(point.id);
  return out;
}

/**
 * §21.4.3.7 — a `presOf`'s axis PATH, when it names more than one step or
 * slices the one it names. A single unsliced axis stays with the plain
 * `presOfAxis` reading, which the rest of the engine already answers.
 *
 * @param presOf The `dgm:presOf` element, when the box has one.
 * @returns The path, or `undefined` when there is nothing to walk.
 */
function axisPath(presOf: PoNode | undefined): LayoutNode['presOfPath'] {
  if (!presOf) return undefined;
  const axis = (poAttr(presOf, 'axis') ?? '').split(/\s+/u).filter((a) => a !== '');
  const nums = (name: string): Array<number> =>
    (poAttr(presOf, name) ?? '')
      .split(/\s+/u)
      .filter((v) => v !== '')
      .map((v) => Number(v));
  const st = nums('st');
  const cnt = nums('cnt');
  if (axis.length === 0) return undefined;
  // One step and no slice is the plain reading; anything else is a path.
  if (axis.length === 1 && st.length === 0 && cnt.length === 0) return undefined;
  if (!axis.every((a) => a === 'ch' || a === 'des' || a === 'self')) return undefined;
  return { axis, st, cnt };
}

// §21.4.3.7 `desOrSelf` — the point's own paragraphs, then every descendant's
// in depth-first order, as one text. Each descendant's paragraphs already carry
// the level they sit at, so they indent themselves.
function withDescendants(point: DiagramPoint, data: DiagramData): DiagramPoint {
  const paragraphs: Array<PoNode> = point.text ? [...poChildren(point.text)] : [];
  const walk = (id: string): void => {
    for (const kid of data.children(id, 'node')) {
      if (kid.text) paragraphs.push(...poChildren(kid.text));
      walk(kid.id);
    }
  };
  walk(point.id);
  if (paragraphs.length === 0) return point;
  return { ...point, text: { 'dgm:t': paragraphs } };
}

/**
 * §21.4.3.2 `composite` — the children are not a list, they are placed: each
 * states where its edges are against the parent's box or against a sibling's,
 * and then runs its own algorithm inside what it was given. A process arrow
 * across the frame with the steps standing along it is one node laid over
 * another, which no list algorithm can express.
 */
/**
 * §21.4.3.4 `<dgm:param type="ar"/>` — the width-to-height ratio the algorithm
 * wants its cell in. The cell shrinks to it on whichever axis is over, and
 * stays centred on what it had.
 *
 * The cycle-matrix layout is the reason this matters: it asks for 1.3, measures
 * its disc against the WIDTH it gets, and on a wide slide placeholder the disc
 * came out half again too big and its four quarters overlapped in a pinwheel.
 *
 * @param node The layout node, for its algorithm parameters.
 * @param box  The cell it was given.
 * @returns The cell at the stated ratio, or unchanged when it states none.
 */
function aspectFitted(node: { kind: 'node' } & LayoutNode, box: Box): Box {
  const ar = Number(node.algParams.get('ar') ?? '');
  if (!Number.isFinite(ar) || ar <= 0 || box.cx <= 0 || box.cy <= 0) return box;
  const cx = Math.min(box.cx, box.cy * ar);
  const cy = Math.min(box.cy, box.cx / ar);
  return { x: box.x + (box.cx - cx) / 2, y: box.y + (box.cy - cy) / 2, cx, cy };
}

function compositeLayout(
  node: { kind: 'node' } & LayoutNode,
  point: DiagramPoint,
  box: Box,
  data: DiagramData,
  index: number,
  outer: ReadonlyArray<Constraint> = [],
): Array<LaidNode> {
  const kids = namedBoxes(node);
  box = aspectFitted(node, box);
  // §21.4.3.1 — a composite places its children but need not size them: the
  // accent process states only `l`, `w` and `t` for its three, and their HEIGHTS
  // come from the node above as multiples of `primFontSz` (0.8 for the words,
  // 1.2 for the shape behind them, 1.6 for the descendants that hang below).
  // Without those every child took the whole cell and the descendants box, set
  // to start where the words end, was pushed clean out of the bottom.
  const run = (fontEmu: number): Array<Box> => {
    const placed = new Map<string, Box>();
    return kids.map((child) => {
      const rect = rectIn(node.constrs, child.name, box, placed, fontEmu, outer);
      placed.set(child.name, rect);
      return rect;
    });
  };
  // The size is what makes the children fill the cell: measure them in units of
  // it, and the deepest one states how many units tall the cell is.
  let rects = run(0);
  if (kids.every((c) => fontHeight(outer, c.name) !== undefined)) {
    const units = run(1).reduce((a, r) => Math.max(a, r.y - box.y + r.cy), 0);
    if (units > 0) rects = run(box.cy / units);
  }
  const out: Array<LaidNode> = [];
  kids.forEach((child, i) => {
    const rect = rects[i];
    if (rect) out.push(...layoutIn(child, point, rect, data, index));
  });
  return out;
}

// The height a node states for a named descendant as a multiple of the font
// size the layout runs at (`h refType="primFontSz"`).
function fontHeight(constrs: ReadonlyArray<Constraint>, name: string): number | undefined {
  for (const k of constrs) {
    if (k.type === 'h' && k.forName === name && k.refType === 'primFontSz' && k.op === undefined) {
      return k.fact;
    }
  }
  return undefined;
}

// Whether a node's constraints say where its children GO, rather than only in
// what order they stack.
//
// Every composite writes `l` and `t` zeros for its first child, and a stack
// writes the next one's `t` against the height of the one above it — that is
// the colour lists, whose two boxes are sized from further up and simply
// follow each other. A composite that CENTRES a child, or hangs it off the
// bottom or the right of its own box, is placing it, and then the placement is
// the whole point: a circle on a line with words above it and space below.
function places(constrs: ReadonlyArray<Constraint>): boolean {
  return constrs.some(
    (k) =>
      k.type === 'ctrX' ||
      k.type === 'ctrY' ||
      k.type === 'r' ||
      k.type === 'b' ||
      ((k.type === 'l' || k.type === 't') && k.refType !== undefined && k.refForName === undefined),
  );
}

// One child's rectangle, from the constraints its parent states for it. A
// constraint with no `refType` states nothing to measure against and reads as
// zero, which is how `<constr type="l" forName="arrow"/>` means the left edge.
function rectIn(
  constrs: ReadonlyArray<Constraint>,
  name: string,
  box: Box,
  placed: ReadonlyMap<string, Box>,
  fontEmu = 0,
  outer: ReadonlyArray<Constraint> = [],
): Box {
  let cx = box.cx;
  let cy = box.cy;
  let x: number | undefined;
  let y: number | undefined;
  const ref = (k: Constraint): number => {
    // A size stated against ITSELF (`w refType="h" refForName` its own name) is
    // the one being built, which is how a circle is made round.
    const from =
      k.refForName === undefined
        ? box
        : k.refForName === name
          ? { cx, cy }
          : (placed.get(k.refForName) ?? box);
    const base =
      k.refType === 'w'
        ? from.cx
        : k.refType === 'h'
          ? from.cy
          : k.refType === 'primFontSz'
            ? fontEmu
            : 0;
    return base * k.fact;
  };
  // The node ABOVE may state a descendant's height in font units while the
  // composite itself states only where the box goes; both are its constraints.
  const mine = [...outer, ...constrs].filter((k) => k.forName === name);
  for (const k of mine) {
    const v = ref(k);
    switch (k.type) {
      case 'w':
        cx = k.op === 'lte' ? Math.min(cx, v) : v;
        break;
      case 'h':
        cy = k.op === 'lte' ? Math.min(cy, v) : v;
        break;
      case 'l':
        x = box.x + v;
        break;
      case 'r':
        x = box.x + v - cx;
        break;
      case 't':
        y = box.y + v;
        break;
      case 'b':
        y = box.y + v - cy;
        break;
      case 'ctrX':
        x = box.x + v - cx / 2;
        break;
      case 'ctrY':
        y = box.y + v - cy / 2;
        break;
      default:
        break;
    }
  }
  return { x: x ?? box.x, y: y ?? box.y, cx, cy };
}

/**
 * §21.4.3.2 `cycle` — the point's children around a circle, starting at
 * `stAng` and spanning `spanAng` (both in degrees, clockwise from twelve
 * o'clock). `ctrShpMap="fNode"` makes the FIRST child the hub in the middle
 * and its own children the ring around it, which is what a radial cycle is.
 */
function cycleLayout(
  node: { kind: 'node' } & LayoutNode,
  point: DiagramPoint,
  box: Box,
  data: DiagramData,
): Array<LaidNode> {
  const boxes = namedBoxes(node);
  const hubNode = node.algParams.get('ctrShpMap') === 'fNode' ? boxes[0] : undefined;
  const ringNode = boxes[boxes.length - 1];
  const kids = data.children(point.id, 'node');
  const hub = hubNode ? kids[0] : undefined;
  const ring = hub ? data.children(hub.id, 'node') : kids;
  if (ring.length === 0 || !ringNode) return [];

  const n = ring.length;
  const stAng = Number(node.algParams.get('stAng') ?? '0');
  const spanAng = Number(node.algParams.get('spanAng') ?? '360');
  // A span of a full turn puts one node at each of `n` even steps; a partial
  // one puts the last node ON its far end, so the step is one shorter.
  const closed = Math.abs(spanAng) >= 359.5;
  const step = spanAng / (closed ? n : Math.max(1, n - 1));

  // §21.4.3.1 — a node's size is stated as a fraction of the diagram. Where it
  // is stated as the WHOLE of it, that is a ceiling and not a size, and the
  // size is the one that packs `n` of them round the ring: the chord between
  // two neighbours is one node plus the space between them.
  const stated = sizeFact(node.constrs, ringNode.name, 'w');
  const gap = siblingSpace(node.constrs, ringNode.name);
  const chord = Math.sin(Math.PI / Math.max(2, n));
  const least = Math.min(box.cx, box.cy);
  const size =
    stated !== undefined && stated < 0.99
      ? { cx: box.cx * stated, cy: box.cy * (sizeFact(node.constrs, ringNode.name, 'h') ?? stated) }
      : { cx: (least * chord) / (1 + gap + chord), cy: (least * chord) / (1 + gap + chord) };
  const radius = Math.min((box.cx - size.cx) / 2, (box.cy - size.cy) / 2);
  const midX = box.x + box.cx / 2;
  const midY = box.y + box.cy / 2;

  const out: Array<LaidNode> = [];
  if (hub && hubNode) {
    out.push(
      ...splitCell(
        node,
        hubNode.name,
        hub,
        data,
        { x: midX - size.cx / 2, y: midY - size.cy / 2, cx: size.cx, cy: size.cy },
        hubNode.shapeType ?? 'rect',
        0,
      ),
    );
  }
  ring.forEach((pt, i) => {
    const a = ((stAng + i * step) * Math.PI) / 180;
    const cellX = midX + radius * Math.sin(a) - size.cx / 2;
    const cellY = midY - radius * Math.cos(a) - size.cy / 2;
    out.push(
      ...splitCell(
        node,
        ringNode.name,
        pt,
        data,
        { x: cellX, y: cellY, cx: size.cx, cy: size.cy },
        ringNode.shapeType ?? 'rect',
        i,
      ),
    );
  });
  return out;
}

// The space a node's constraints leave between two siblings, as a fraction of
// the node's own width. A cycle may state a NEGATIVE one, and then the ring
// overlaps itself.
function siblingSpace(constrs: ReadonlyArray<Constraint>, childName: string): number {
  for (const k of constrs) {
    if (k.type === 'sibSp' && (k.refForName === childName || k.refForName === undefined)) {
      return k.fact;
    }
  }
  return 0.5;
}

/**
 * §21.4.3.2 `lin` / `snake` — the point's children in a row or a column, with
 * the gap the constraints state between them. `snake` wraps; `lin` does not.
 */
function listLayout(
  parsed: { kind: 'node' } & LayoutNode,
  parent: DiagramPoint,
  box: Box,
  data: DiagramData,
): Array<LaidNode> {
  const nodes = data.children(parent.id, 'node');
  if (nodes.length === 0) return [];

  // The child layoutNode's own name is what the constraints refer to.
  const childName = firstNodeName(parsed.children) ?? 'node';
  const c = resolve(parsed.constrs, childName);

  // The number of columns is the one that leaves each cell closest to the shape
  // the constraints ask for (`h` stated against `w`), which is what "fill the
  // frame" means: five nodes in a 4:3 frame come out two across and three down,
  // as every reader draws them.
  const shapeType = childShapeType(parsed.children) ?? 'rect';
  const vertical = isVertical(parsed);
  // The boxes the forEach lays out per point. One is the usual case; several
  // is a stack, and each of them is a box in its own right.
  const body = namedBoxes(parsed);
  // The shape the layout draws BETWEEN two cells, when it draws one.
  const trans = transitionBox(parsed);
  const n = nodes.length;
  // A column list is one box wide however much room there is beside it, and a
  // row that does not snake is one box tall: only a snake across chooses.
  const cols = vertical
    ? 1
    : parsed.alg === 'snake'
      ? bestColumns(n, box.cx, box.cy, c.aspect, c.gap)
      : n;
  const rows = Math.ceil(n / cols);

  // Each track is a cell plus the gap that follows it, the last gap trimmed.
  const cellW = box.cx / (cols + Math.max(0, cols - 1) * c.gap);
  const gapW = cellW * c.gap;
  const cellH = box.cy / (rows + Math.max(0, rows - 1) * c.gap);
  const gapH = cellH * c.gap;

  const out: Array<LaidNode> = [];
  nodes.forEach((point, i) => {
    const col = vertical ? Math.floor(i / rows) : i % cols;
    const row = vertical ? i % rows : Math.floor(i / cols);
    // A last row that is short is centred under the ones above it, which is
    // where PowerPoint leaves it.
    const inRow = vertical ? rows : Math.min(cols, n - row * cols);
    const rowWidth = inRow * cellW + Math.max(0, inRow - 1) * gapW;
    const left = vertical ? 0 : (box.cx - rowWidth) / 2;
    const cell = {
      x: box.x + left + col * (cellW + gapW),
      y: box.y + row * (cellH + gapH),
      cx: cellW,
      cy: cellH,
    };
    // §21.4.3 — the child layoutNode may not be a box at all: a list that
    // shows a node's label beside its children's text nests a `linNode` whose
    // own children are two `tx` boxes, sized by ITS constraints. Laid out as
    // one box, such a row is the whole cell where it should be two.
    out.push(
      ...(body.length > 1
        ? stackCell(parsed, body, point, data, cell, shapeType, i, vertical)
        : splitCell(parsed, childName, point, data, cell, shapeType, i)),
    );
    // §21.4.3.2 `conn` — the gap between two cells is where the layout puts the
    // thing that JOINS them, and a process draws an arrow there. The engine
    // already measured that gap and left it empty.
    const last = vertical ? row === rows - 1 : col === inRow - 1;
    if (trans && !last) {
      out.push(
        connector(
          trans,
          point,
          {
            x: vertical ? cell.x : cell.x + cellW,
            y: vertical ? cell.y + cellH : cell.y,
            cx: vertical ? cellW : gapW,
            cy: vertical ? gapH : cellH,
          },
          vertical,
        ),
      );
    }
  });
  return out;
}

/**
 * §21.4.3.2 — the box a list draws in the gap between two cells.
 *
 * `dgm:shape type="conn"` is not a preset geometry, it is "whatever joins these
 * two", and for a list laid out across that is an arrow pointing the way the
 * list runs. Its own constraint sizes it against the gap it was given — the
 * accent process asks for a height of 62% of its width — and it is centred in
 * what is left.
 */
function connector(
  node: { kind: 'node' } & LayoutNode,
  point: DiagramPoint,
  gap: Box,
  vertical: boolean,
): LaidNode {
  let { cx, cy } = gap;
  for (const k of node.constrs) {
    if (k.forName !== undefined || k.refType === undefined) continue;
    if (k.type === 'h' && k.refType === 'w') cy = Math.min(gap.cy, gap.cx * k.fact);
    if (k.type === 'w' && k.refType === 'h') cx = Math.min(gap.cx, gap.cy * k.fact);
  }
  return {
    point: { id: `${point.id}/${node.name}`, type: 'node' },
    // It points the way the list runs, not the way its gap happens to be shaped:
    // a row's gap is narrow and tall, and read off that the arrow pointed down.
    shapeType:
      node.shapeType === 'conn' || node.shapeType === undefined
        ? vertical
          ? 'downArrow'
          : 'rightArrow'
        : node.shapeType,
    ...styleOf(node),
    // §21.4.5 — a transition that draws a SHAPE and names no label of its own
    // takes the colours part's `sibTrans2D1`, the two-dimensional sibling
    // transition, which is the node accent at a 60% tint rather than the flat
    // accent everything unlabelled would otherwise fall back to.
    ...(node.styleLbl === undefined ? { styleLbl: 'sibTrans2D1' } : {}),
    index: 0,
    x: gap.x + (gap.cx - cx) / 2,
    y: gap.y + (gap.cy - cy) / 2,
    cx,
    cy,
  };
}

// The box a nested `forEach` over the sibling transitions lays out, when the
// layout gives that transition a shape rather than using it as bare spacing.
function transitionBox(
  node: { kind: 'node' } & LayoutNode,
): ({ kind: 'node' } & LayoutNode) | undefined {
  for (const item of node.children) {
    if (item.kind !== 'forEach') continue;
    for (const inner of item.children) {
      if (inner.kind !== 'forEach' || inner.ptType !== 'sibTrans') continue;
      for (const box of inner.children) {
        if (box.kind === 'node' && box.shapeType !== undefined && box.hideGeom !== true) return box;
      }
    }
  }
  return undefined;
}

/**
 * §21.4.3 — a `forEach` body of SEVERAL boxes is a stack, not a box.
 *
 * The vertical list family gives each point three: the row holding its label,
 * a negative space, and the box its descendants' words go in. Laying out only
 * the first of them drew the labels and dropped every descendant.
 *
 * What each box states for its height divides the cell between them. A space
 * stated as negative is an overlap — one box pulled up over the one before it —
 * which this engine does not run, so it takes no room and the rest share the
 * cell; LibreOffice draws these layouts without the overlap too.
 */
function stackCell(
  parsed: { kind: 'node' } & LayoutNode,
  body: ReadonlyArray<{ kind: 'node' } & LayoutNode>,
  point: DiagramPoint,
  data: DiagramData,
  cell: { x: number; y: number; cx: number; cy: number },
  shapeType: string,
  index: number,
  vertical: boolean,
): Array<LaidNode> {
  const stated = body.map((b) => stackFact(parsed, b) ?? 0);
  const total = stated.reduce((a, v) => a + Math.max(0, v), 0);
  const shares = stated.map((v) => (total > 0 ? Math.max(0, v) / total : 1 / body.length));

  const out: Array<LaidNode> = [];
  let at = vertical ? cell.y : cell.x;
  body.forEach((b, k) => {
    const span = (vertical ? cell.cy : cell.cx) * (shares[k] ?? 0);
    // A `sp` box naming no shape is space and nothing else.
    if (span > 0 && !(b.alg === 'sp' && b.shapeType === undefined)) {
      const slice = {
        x: vertical ? cell.x : at,
        y: vertical ? at : cell.y,
        cx: vertical ? cell.cx : span,
        cy: vertical ? span : cell.cy,
      };
      out.push(...splitCell(parsed, b.name, point, data, slice, b.shapeType ?? shapeType, index));
    }
    at += span;
  });
  return out;
}

/**
 * The boxes a wrapper INSETS, each at the width it states and where it falls.
 *
 * Widths that add up to the whole are a division and belong to the split; ones
 * that fall short are placements — a margin, then a label of seven tenths, then
 * nothing. Undefined when the boxes do not all state a width, or when they fill
 * the row between them.
 */
function insetBoxes(
  parent: { kind: 'node' } & LayoutNode,
  child: { kind: 'node' } & LayoutNode,
  width: number,
): Array<{ box: { kind: 'node' } & LayoutNode; x: number; cx: number }> | undefined {
  const boxes = namedBoxes(child);
  if (boxes.length < 2) return undefined;
  const facts = boxes.map(
    (b) => sizeFact(child.constrs, b.name, 'w') ?? sizeFact(parent.constrs, b.name, 'w'),
  );
  if (facts.some((v) => v === undefined)) return undefined;
  const total = facts.reduce<number>((a, v) => a + (v ?? 0), 0);
  if (total > 0.999) return undefined;
  const out: Array<{ box: { kind: 'node' } & LayoutNode; x: number; cx: number }> = [];
  let at = 0;
  boxes.forEach((box, i) => {
    const cx = width * (facts[i] ?? 0);
    out.push({ box, x: at, cx });
    at += cx;
  });
  return out;
}

// What one box of a stack states for its height, as a factor. A wrapper states
// nothing itself — the box it insets does.
function stackFact(
  parsed: { kind: 'node' } & LayoutNode,
  box: { kind: 'node' } & LayoutNode,
): number | undefined {
  const own = heightFact(parsed, box.name) ?? sizeFact(parsed.constrs, box.name, 'h');
  if (own !== undefined) return own;
  for (const inner of namedBoxes(box)) {
    const f = heightFact(parsed, inner.name) ?? sizeFact(parsed.constrs, inner.name, 'h');
    if (f !== undefined) return f;
  }
  return undefined;
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
  // §21.4.3.2 — a box that hides its geometry and runs the space algorithm is
  // spacing, not content, and so is never one half of a split. The vertical
  // list family insets its label with exactly such a margin box: split against
  // it, the node's own words went into the invisible half and vanished.
  const inner = (child ? namedBoxes(child) : []).filter(
    (b) => !(b.alg === 'sp' && b.hideGeom === true),
  );
  // A size can be stated on the box itself (no `forName`), by the node that
  // lays it out, or by the top node for every descendant of that name.
  const size = (
    box: ({ kind: 'node' } & LayoutNode) | undefined,
    type: 'primFontSz' | 'secFontSz',
  ): number | undefined =>
    box === undefined
      ? undefined
      : (fontSize(box.constrs, undefined, type) ??
        fontSize(child?.constrs ?? [], box.name, type) ??
        fontSize(parent.constrs, box.name, type));
  // A cell that is not divided wears the style of the one box that fills it —
  // the box itself when the child layoutNode IS the box, and its single content
  // box when the child is a wrapper that only insets one.
  const face = inner.length === 1 ? inner[0] : child;
  const whole = (): Array<LaidNode> => [
    {
      point: face === undefined ? point : spokenFor(face, point, data),
      shapeType: face?.shapeType ?? shapeType,
      ...(face?.styleLbl !== undefined ? { styleLbl: face.styleLbl } : {}),
      ...opt('fontSizePt', size(face, 'primFontSz')),
      ...(face !== undefined && bulleted(face) ? { bulleted: true } : {}),
      ...(face?.hideGeom === true ? { hideGeom: true as const } : {}),
      ...(face === undefined ? {} : textPlacing(face)),
      index,
      ...cell,
    },
  ];
  // A cell whose node PLACES its children rather than dividing them goes
  // through the composite algorithm: three boxes to a step — the words, the
  // circle on the line and the space opposite — is not a split of anything.
  // A composite that states only WIDTHS places nothing: the colour lists size
  // their two boxes from the top node, and stacking them is what that means.
  if (child?.alg === 'composite' && inner.length > 0 && places(child.constrs)) {
    return compositeLayout(child, point, cell, data, index, parent.constrs);
  }
  // A wrapper whose boxes state widths that do NOT fill it is INSETTING them,
  // not sharing it out: the vertical list gives its row a margin of a twentieth
  // and a label of seven tenths, and what is left over stays empty. Sharing the
  // row between them instead ran the label the whole way across.
  const inset = child ? insetBoxes(parent, child, cell.cx) : undefined;
  if (child && inset) {
    const out: Array<LaidNode> = [];
    for (const { box, x, cx } of inset) {
      if (box.alg === 'sp' && box.hideGeom === true) continue;
      out.push(
        ...splitCell(
          parent,
          box.name,
          point,
          data,
          { ...cell, x: cell.x + x, cx },
          box.shapeType ?? shapeType,
          index,
        ),
      );
    }
    if (out.length > 0) return out;
  }
  // Two boxes is a label and its descendants; more is a row of several — the
  // bracket list divides its row four ways (the label, the bracket, a gap and
  // the descendants), and taken as one box it drew the label alone.
  if (!child || inner.length < 2) return whole();
  const widths = inner.map((b) => sizeFact(child.constrs, b.name, 'w'));
  // A child given less than the whole width sits BESIDE its sibling; one given
  // all of it sits above, and then the heights are what divide the cell.
  const across = widths.some((v) => v !== undefined && v < 0.99);
  // Two boxes the engine shares out from what one of them states. More than two
  // it divides only when they are a ROW that adds up — every share stated and
  // the row filled — which is what tells the bracket list's four columns from a
  // node whose several full-width boxes the engine has no way to place.
  const share = widths.reduce<number>((a, v) => a + (v ?? 0), 0);
  if (
    inner.length > 2 &&
    !(across && widths.every((v) => v !== undefined) && Math.abs(share - 1) < 0.02)
  ) {
    return whole();
  }
  const kids = data.children(point.id, 'node');

  const shares = across
    ? fractions(widths)
    : fractions(
        inner.map((b) => sizeFact(child.constrs, b.name, 'h') ?? heightFact(parent, b.name)),
      );

  // §21.4.3.1 — a box may state a CEILING on its own height against its width
  // (`h refType="w" op="lte" fact="0.4"`). It is a bound, not a share: the
  // label of a colour list is at most four tenths of its column however tall
  // the column is, and what it does not take goes to the box below it.
  const cap =
    inner.length === 2
      ? heightCap(inner[0], across ? cell.cx * (shares[0] ?? 1) : cell.cx)
      : undefined;
  if (!across && cap !== undefined) {
    const first = Math.min(cell.cy * (shares[0] ?? 0.5), cap);
    shares[0] = first / cell.cy;
    shares[1] = 1 - shares[0];
  }

  const out: Array<LaidNode> = [];
  let at = across ? cell.x : cell.y;
  inner.forEach((box, i) => {
    const span = (across ? cell.cx : cell.cy) * (shares[i] ?? 0.5);
    // The first box is the node's own label; the second holds its children's,
    // and is left out when the layout guards it and there are none. A box that
    // names the points it speaks for is taken at its word instead.
    const spacer = box.alg === 'sp' && box.shapeType === undefined;
    if (span > 0 && !spacer && !(box.needsChildren && kids.length === 0)) {
      out.push({
        // A box running the space algorithm is a shape and no words, however
        // it was sized — the bracket between a label and its children drew a
        // stray copy of the children's text down its own spine.
        point:
          box.alg === 'sp'
            ? { id: `${point.id}/${box.name}`, type: 'node' }
            : box.presOfAxis === undefined
              ? i === 0
                ? point
                : gatheredText(point, kids)
              : spokenFor(box, point, data),
        shapeType: i === 0 ? shapeType : (box.shapeType ?? 'rect'),
        ...(box.styleLbl !== undefined ? { styleLbl: box.styleLbl } : {}),
        ...opt('fontSizePt', size(box, i === 0 ? 'primFontSz' : 'secFontSz')),
        ...(bulleted(box) ? { bulleted: true } : {}),
        ...(box.hideGeom === true ? { hideGeom: true } : {}),
        ...textPlacing(box),
        index,
        x: across ? at : cell.x,
        y: across ? cell.y : at,
        cx: across ? span : cell.cx,
        cy: across ? cell.cy : span,
      });
    }
    at += span;
  });
  return out.length > 0 ? out : whole();
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

// The point size a constraint states for a box's text. The primary size is the
// node's own label, the secondary its descendants' — and a secondary stated
// against the primary (`refType="primFontSz"`) is that size, whatever it is.
function fontSize(
  constrs: ReadonlyArray<Constraint>,
  name: string | undefined,
  type: 'primFontSz' | 'secFontSz',
): number | undefined {
  for (const k of constrs) {
    if (k.type !== type || k.forName !== name) continue;
    if (k.val !== undefined) return k.val;
    if (k.refType === 'primFontSz' && k.refForName !== undefined) {
      const from = fontSize(constrs, k.refForName, 'primFontSz');
      if (from !== undefined) return from * (k.fact || 1);
    }
  }
  return type === 'secFontSz' ? fontSize(constrs, name, 'primFontSz') : undefined;
}

// §21.4.3.2 `stBulletLvl` — the `tx` algorithm's parameter for text that starts
// bulleted, which is how a descendants box reads as a list and not a paragraph.
function bulleted(node: { kind: 'node' } & LayoutNode): boolean {
  return Number(node.algParams.get('stBulletLvl') ?? '0') >= 1;
}

function opt<T extends string>(name: T, v: number | undefined): { [P in T]?: number } {
  return (v === undefined ? {} : { [name]: v }) as { [P in T]?: number };
}

// The greatest height a box allows itself, as its own constraints state it
// against its width.
function heightCap(
  box: ({ kind: 'node' } & LayoutNode) | undefined,
  widthEmu: number,
): number | undefined {
  for (const k of box?.constrs ?? []) {
    if (k.type === 'h' && k.refType === 'w' && k.op === 'lte') return widthEmu * k.fact;
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
  readonly data: DiagramData;
  readonly root: DiagramPoint;
}

/**
 * §21.4.3.3 — a `choose` picks one branch by a function of the data, and what
 * the branch holds belongs to whatever the `choose` was standing in. The
 * direction variables are left at their defaults, so those take the first
 * `if`; `maxDepth` is answered from the model, because the branch it guards is
 * the difference between one box in a cell and two.
 *
 * @returns The children with every `choose` replaced by the branch taken, and
 *          the ones whose branch was guarded on a child count.
 */
function flattenChoose(
  kids: ReadonlyArray<PoNode>,
  ctx: LayoutContext,
): { flat: Array<PoNode>; guarded: Set<PoNode> } {
  const flat: Array<PoNode> = [];
  const guarded = new Set<PoNode>();
  for (const k of kids) {
    if (!poIs(k, 'dgm:choose')) {
      flat.push(k);
      continue;
    }
    const branches = poChildren(k);
    const taken =
      branches.find((b) => poIs(b, 'dgm:if') && holds(b, ctx)) ??
      branches.find((b) => poIs(b, 'dgm:else')) ??
      branches[0];
    const onCount = taken !== undefined && poIs(taken, 'dgm:if') && countsChildren(taken);
    // A branch may hold another `choose`, and a producer nests them freely.
    const inner = flattenChoose(poChildren(taken), ctx);
    for (const item of inner.flat) {
      flat.push(item);
      if (onCount || inner.guarded.has(item)) guarded.add(item);
    }
  }
  return { flat, guarded };
}

function parseNode(node: PoNode, ctx: LayoutContext): { kind: 'node' } & LayoutNode {
  const { flat: flattened, guarded } = flattenChoose(poChildren(node), ctx);
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
  const presOf = flattened.find((k) => poIs(k, 'dgm:presOf'));
  const presOfAxis = poAttr(presOf, 'axis');
  const presOfPath = axisPath(presOf);
  const shapeRot = shape ? Number(poAttr(shape, 'rot') ?? '') : NaN;
  return {
    kind: 'node',
    name: poAttr(node, 'name') ?? '',
    ...(styleLbl !== undefined && styleLbl !== '' ? { styleLbl } : {}),
    needsChildren: false,
    alg: algOf(alg),
    algParams: params,
    ...(shapeType !== undefined && shapeType !== '' ? { shapeType } : {}),
    ...(shape !== undefined && poAttr(shape, 'hideGeom') === '1' ? { hideGeom: true } : {}),
    ...(Number.isFinite(shapeRot) && shapeRot !== 0 ? { shapeRotation: shapeRot } : {}),
    ...(presOfAxis !== undefined && presOfAxis !== '' ? { presOfAxis } : {}),
    ...(presOfPath ? { presOfPath } : {}),
    constrs: parseConstraints(constrLst),
    children: flattened.flatMap((k) => parseItem(k, ctx, guarded.has(k))),
  };
}

// Whether a `dgm:if` holds. Only `maxDepth` is answered — every other test is
// on a variable this engine leaves at its default, where the `if` is the branch
// PowerPoint takes.
function holds(branch: PoNode, ctx: LayoutContext): boolean {
  const func = poAttr(branch, 'func');
  const axis = (poAttr(branch, 'axis') ?? '').split(/\s+/u).filter((a) => a !== '');
  if (func === 'maxDepth') return compare(ctx.maxDepth, branch, 'gte');
  // A `cnt` over a path of SEVERAL steps is a question about the model, asked
  // once: "how many children has the first child?" is what tells a radial
  // cycle where to start its ring. A one-step `cnt` is about the point being
  // laid out, which is not known here — that one is answered per point, by the
  // guard `countsChildren` puts on the box.
  if (func === 'cnt' && axis.length > 1) return compare(countAt(axis, branch, ctx), branch, 'gte');
  return true;
}

// The number of points at the end of a `dgm:if`'s path: each step takes an
// axis of the points it has so far, then the slice `st`/`cnt` names (both
// 1-based, and a count of 0 means all of them).
function countAt(axis: ReadonlyArray<string>, branch: PoNode, ctx: LayoutContext): number {
  const nums = (name: string): Array<number> =>
    (poAttr(branch, name) ?? '').split(/\s+/u).map((v) => Number(v));
  const st = nums('st');
  const cnt = nums('cnt');
  const types = (poAttr(branch, 'ptType') ?? '').split(/\s+/u);
  let at: Array<DiagramPoint> = [ctx.root];
  axis.forEach((step, i) => {
    if (step !== 'ch' && step !== 'des') {
      at = [];
      return;
    }
    const kind = types[i] === 'node' || types[i] === undefined ? 'node' : undefined;
    if (kind === undefined) {
      at = [];
      return;
    }
    const next = at.flatMap((p) => ctx.data.children(p.id, 'node'));
    const from = Math.max(0, (st[i] ?? 1) - 1);
    const want = cnt[i];
    const take = want === undefined || want === 0 ? next.length : want;
    at = next.slice(from, from + take);
  });
  return at.length;
}

// A `dgm:if`'s comparison, whatever it is comparing.
function compare(actual: number, branch: PoNode, fallback: string): boolean {
  const val = Number(poAttr(branch, 'val') ?? '0');
  const op = poAttr(branch, 'op') ?? fallback;
  if (op === 'gte') return actual >= val;
  if (op === 'gt') return actual > val;
  if (op === 'lte') return actual <= val;
  if (op === 'lt') return actual < val;
  if (op === 'neq') return actual !== val;
  return actual === val;
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
    // A `forEach` may hold a `choose` too — a process alternates the box above
    // the line with the one below it that way, and read as nothing the whole
    // repetition came out empty.
    const inner = flattenChoose(poChildren(node), ctx);
    return [
      {
        kind: 'forEach',
        axis: poAttr(node, 'axis') ?? '',
        ptType: poAttr(node, 'ptType') ?? '',
        children: inner.flat.flatMap((k) => parseItem(k, ctx, guarded || inner.guarded.has(k))),
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
    const op = poAttr(k, 'op');
    const refType = poAttr(k, 'refType');
    const refForName = poAttr(k, 'refForName');
    const val = poAttr(k, 'val');
    out.push({
      type,
      ...(forName !== undefined ? { forName } : {}),
      ...(op !== undefined ? { op } : {}),
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
  // §21.4.3.2 — `hierRoot` places a node above its subtree and `hierChild`
  // ranges that subtree beneath it. They are two halves of one tree walk, so
  // the engine runs them as one.
  if (t === 'hierRoot' || t === 'hierChild') return 'hier';
  return t === 'snake' ||
    t === 'lin' ||
    t === 'sp' ||
    t === 'tx' ||
    t === 'composite' ||
    t === 'cycle'
    ? t
    : 'other';
}

function findChild(nodes: ReadonlyArray<PoNode>, tag: string): PoNode | undefined {
  return nodes.find((n) => poIs(n, tag) || poTag(n) === tag);
}
