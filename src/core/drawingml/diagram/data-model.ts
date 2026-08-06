// ECMA-376 §21.4.2 — a SmartArt diagram's DATA part (`diagrams/data1.xml`):
// the points the user typed and the connections between them. It is a graph,
// not a tree, and the tree the layout walks is recovered from the connections:
// a `parOf` names a parent, a child and the child's position among its
// siblings, and hangs the two transition points (the arrow between siblings,
// the arrow down to a child) off the same connection.
//
// This part is always present. The DRAWING part beside it — the shapes
// PowerPoint laid out and cached — is not: a file written by a generator
// carries data, layout, colours and style and nothing else, and then the
// picture has to be computed. That is what this model feeds.

import type { PoNode } from '@/core/po-helpers';
import { poAttr, poChildren, poIs, poText } from '@/core/po-helpers';

/** §21.4.2.27 `dgm:pt@type` — what a point is in the model. */
export type PointType = 'doc' | 'node' | 'asst' | 'parTrans' | 'sibTrans' | 'pres';

/** One point of the data model. */
export interface DiagramPoint {
  readonly id: string;
  readonly type: PointType;
  /** `dgm:t` — the point's text, as the `a:txBody` every other reader takes. */
  readonly text?: PoNode;
  /** `dgm:spPr` — the formatting the user overrode on this node. */
  readonly spPr?: PoNode;
  /** `dgm:prSet` — presentation properties (placeholder text, custom sizes). */
  readonly prSet?: PoNode;
}

/** One `parOf` connection: a parent, a child, and where the child sits. */
interface ParentOf {
  readonly parent: string;
  readonly child: string;
  readonly ord: number;
  readonly sibTransId?: string;
  readonly parTransId?: string;
}

/**
 * The data part as the layout engine asks about it: the point tree recovered
 * from the connection list, plus the transition points that go with each edge.
 */
export class DiagramData {
  private readonly byId = new Map<string, DiagramPoint>();
  private readonly kids = new Map<string, Array<ParentOf>>();
  /** The document point every layout starts from. */
  readonly root: DiagramPoint | undefined;

  constructor(tree: ReadonlyArray<PoNode>) {
    const model = findDescendant(tree, 'dgm:dataModel');
    const ptLst = model ? poChildren(model).find((c) => poIs(c, 'dgm:ptLst')) : undefined;
    for (const pt of ptLst ? poChildren(ptLst) : []) {
      if (!poIs(pt, 'dgm:pt')) continue;
      const id = poAttr(pt, 'modelId');
      if (id === undefined) continue;
      const kids = poChildren(pt);
      this.byId.set(id, {
        id,
        type: (poAttr(pt, 'type') as PointType | undefined) ?? 'node',
        ...pick(kids, 'dgm:t', 'text'),
        ...pick(kids, 'dgm:spPr', 'spPr'),
        ...pick(kids, 'dgm:prSet', 'prSet'),
      });
    }
    const cxnLst = model ? poChildren(model).find((c) => poIs(c, 'dgm:cxnLst')) : undefined;
    for (const cxn of cxnLst ? poChildren(cxnLst) : []) {
      if (!poIs(cxn, 'dgm:cxn')) continue;
      // §21.4.2.9 — `parOf` is the only connection that builds the tree; the
      // rest (`presOf`, `presParOf`) bind presentation nodes, which a file
      // without a drawing part does not have.
      if ((poAttr(cxn, 'type') ?? 'parOf') !== 'parOf') continue;
      const parent = poAttr(cxn, 'srcId');
      const child = poAttr(cxn, 'destId');
      if (parent === undefined || child === undefined) continue;
      const list = this.kids.get(parent) ?? [];
      list.push({
        parent,
        child,
        ord: Number(poAttr(cxn, 'srcOrd') ?? '0'),
        ...opt('sibTransId', poAttr(cxn, 'sibTransId')),
        ...opt('parTransId', poAttr(cxn, 'parTransId')),
      });
      this.kids.set(parent, list);
    }
    for (const list of this.kids.values()) list.sort((a, b) => a.ord - b.ord);
    this.root = [...this.byId.values()].find((p) => p.type === 'doc');
  }

  /** The point with this id, if the model holds one. */
  point(id: string): DiagramPoint | undefined {
    return this.byId.get(id);
  }

  /**
   * The children of a point, in the order the connections give them.
   *
   * @param id       The parent's model id.
   * @param ptType   Which kind of child to return — `node` skips the
   *                 transitions, `sibTrans` returns them instead.
   */
  children(id: string, ptType: 'node' | 'sibTrans' | 'parTrans' = 'node'): Array<DiagramPoint> {
    const out: Array<DiagramPoint> = [];
    for (const edge of this.kids.get(id) ?? []) {
      const wanted =
        ptType === 'node' ? edge.child : ptType === 'sibTrans' ? edge.sibTransId : edge.parTransId;
      const pt = wanted === undefined ? undefined : this.byId.get(wanted);
      // A transition point exists for every edge but the layout draws one only
      // BETWEEN siblings, so the last child's is dropped by the caller.
      if (pt && (ptType !== 'node' || (pt.type !== 'parTrans' && pt.type !== 'sibTrans'))) {
        out.push(pt);
      }
    }
    return out;
  }

  /**
   * How many generations of `node` points the model holds below its root.
   *
   * §21.4.3.3 — a layout branches on this (`func="maxDepth" op="gte" val="2"`):
   * the same part lays a chevron out as one box when the nodes have no
   * children and as a labelled box over a body when they do.
   */
  get depth(): number {
    const walk = (id: string, seen: Set<string>): number => {
      if (seen.has(id)) return 0;
      seen.add(id);
      let best = 0;
      for (const kid of this.children(id, 'node')) best = Math.max(best, walk(kid.id, seen));
      return best + 1;
    };
    return this.root ? walk(this.root.id, new Set()) - 1 : 0;
  }

  /** The plain text of a point's `dgm:t`, paragraphs joined by newlines. */
  static textOf(pt: DiagramPoint | undefined): string {
    if (!pt?.text) return '';
    const lines: Array<string> = [];
    for (const p of poChildren(pt.text)) {
      if (!poIs(p, 'a:p')) continue;
      let line = '';
      for (const r of poChildren(p)) {
        // A break inside a paragraph is a line of its own, as `a:p` is.
        if (poIs(r, 'a:br')) {
          lines.push(line);
          line = '';
        } else if (poIs(r, 'a:r')) line += poText(poChildren(r).find((c) => poIs(c, 'a:t')));
      }
      lines.push(line);
    }
    return lines.join('\n');
  }
}

function pick<T extends string>(
  kids: ReadonlyArray<PoNode>,
  tag: string,
  as: T,
): { [P in T]?: PoNode } {
  const found = kids.find((c) => poIs(c, tag));
  return (found ? { [as]: found } : {}) as { [P in T]?: PoNode };
}

function opt(name: 'sibTransId' | 'parTransId', v: string | undefined): Record<string, string> {
  return v === undefined || v === '' ? {} : { [name]: v };
}

function findDescendant(tree: ReadonlyArray<PoNode>, tag: string): PoNode | undefined {
  for (const node of tree) {
    if (poIs(node, tag)) return node;
    const found = findDescendant(poChildren(node), tag);
    if (found) return found;
  }
  return undefined;
}
