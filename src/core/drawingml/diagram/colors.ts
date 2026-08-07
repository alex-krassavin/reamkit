// ECMA-376 §21.4.5 — a SmartArt COLOURS part (`diagrams/colors1.xml`): for each
// style label a layout node can carry, the run of colours its boxes take.
//
// The list is a run, not a colour: `alignNode1` names accent2 through accent6,
// and the nodes of a list walk it, so three columns come out in three different
// accents. The box that follows a node — its descendants' text — has its own
// label whose list is the same accents at a 40% tint with dark text on top.
// Without this every box is one accent and every diagram is monochrome.

import type { PoNode } from '@/core/po-helpers';
import { poAttr, poChildren, poIs, poTag } from '@/core/po-helpers';

/** One `dgm:*ClrLst`: the colours and how a run of boxes walks them. */
interface ColorRun {
  readonly meth: string;
  readonly colors: ReadonlyArray<PoNode>;
}

interface StyleColors {
  readonly fill?: ColorRun;
  readonly line?: ColorRun;
  readonly txFill?: ColorRun;
}

/** The colours part, as the drawing writer asks it. */
export class DiagramColors {
  private readonly byLabel = new Map<string, StyleColors>();

  constructor(tree: ReadonlyArray<PoNode>) {
    const def = findDescendant(tree, 'dgm:colorsDef');
    for (const lbl of def ? poChildren(def) : []) {
      if (!poIs(lbl, 'dgm:styleLbl')) continue;
      const name = poAttr(lbl, 'name');
      if (name === undefined) continue;
      const kids = poChildren(lbl);
      this.byLabel.set(name, {
        ...run(kids, 'dgm:fillClrLst', 'fill'),
        ...run(kids, 'dgm:linClrLst', 'line'),
        ...run(kids, 'dgm:txFillClrLst', 'txFill'),
      });
    }
  }

  /**
   * The `a:solidFill` for a box, or undefined when the part states none.
   *
   * @param label The layout node's `styleLbl`.
   * @param index The box's place in the run of siblings.
   */
  fill(label: string | undefined, index: number): string | undefined {
    return solid(this.pick(label, 'fill', index));
  }

  /**
   * Whether the part states this label at all.
   *
   * A label it names but gives no fill — `revTx`, the reversed text of a
   * process arrow — is a box with NO fill, which is not the same as a file
   * that ships no colours and falls back to one accent for everything.
   */
  knows(label: string | undefined): boolean {
    return label !== undefined && this.byLabel.has(label);
  }

  /**
   * The `a:ln` for a box's OUTLINE, where the part states one.
   *
   * A follower box is often a near-white wash with the node's own accent drawn
   * round it (`conFgAcc1` is `lt1` at 90% inside an `accent1` line): without the
   * line it is a white rectangle on a white slide and reads as nothing at all.
   *
   * @param label The layout node's `styleLbl`.
   * @param index The box's place in the run of siblings.
   */
  line(label: string | undefined, index: number): string | undefined {
    const fill = solid(this.pick(label, 'line', index));
    return fill === undefined ? undefined : `<a:ln>${fill}</a:ln>`;
  }

  /** The `a:solidFill` for a box's TEXT, where the part overrides the style. */
  textFill(label: string | undefined, index: number): string | undefined {
    return solid(this.pick(label, 'txFill', index));
  }

  private pick(
    label: string | undefined,
    which: keyof StyleColors,
    index: number,
  ): PoNode | undefined {
    const found = label === undefined ? undefined : this.byLabel.get(label)?.[which];
    const colors = found?.colors ?? [];
    if (colors.length === 0) return undefined;
    // §21.4.5.3 — `repeat` and `cycle` both walk the run and start over; `span`
    // spreads it across the siblings, which without their count is the same
    // first colour. Either way an index past the end wraps.
    return colors[((index % colors.length) + colors.length) % colors.length];
  }
}

function run(
  kids: ReadonlyArray<PoNode>,
  tag: string,
  as: 'fill' | 'line' | 'txFill',
): Partial<StyleColors> {
  const list = kids.find((c) => poIs(c, tag));
  const colors = list ? poChildren(list).filter((c) => poTag(c)?.endsWith('Clr') === true) : [];
  if (colors.length === 0) return {};
  return { [as]: { meth: poAttr(list, 'meth') ?? 'repeat', colors } };
}

function solid(color: PoNode | undefined): string | undefined {
  const xml = color === undefined ? '' : colorXml(color);
  return xml === '' ? undefined : `<a:solidFill>${xml}</a:solidFill>`;
}

// One colour re-emitted as DrawingML: the colour itself and the transforms on
// it (`a:tint`, `a:alpha`, `a:shade`, `a:lumMod`), which is what makes the
// follower box a light wash of the node's own accent rather than the accent.
function colorXml(color: PoNode): string {
  const tag = poTag(color);
  if (tag === undefined) return '';
  const val = poAttr(color, 'val');
  const mods = poChildren(color)
    .map((m) => {
      const t = poTag(m);
      const v = poAttr(m, 'val');
      return t === undefined || v === undefined ? '' : `<${t} val="${esc(v)}"/>`;
    })
    .join('');
  const open = val === undefined ? `<${tag}>` : `<${tag} val="${esc(val)}">`;
  return `${open}${mods}</${tag}>`;
}

function esc(s: string): string {
  return s.replace(/&/gu, '&amp;').replace(/</gu, '&lt;').replace(/"/gu, '&quot;');
}

function findDescendant(tree: ReadonlyArray<PoNode>, tag: string): PoNode | undefined {
  for (const node of tree) {
    if (poIs(node, tag)) return node;
    const found = findDescendant(poChildren(node), tag);
    if (found) return found;
  }
  return undefined;
}
