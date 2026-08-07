// Running a diagram's parts to the drawing PowerPoint would have cached.
//
// A SmartArt frame can turn up in a slide, a document or a sheet, and each of
// those resolves the parts its own way — Word hangs the layout off the document
// beside the data rel, PowerPoint off the slide. What happens next is the same
// either way: read the data, run the layout in the frame the drawing gives it,
// colour it from the colours part, and hand the result to the same `dsp` reader
// a file WITH a cached drawing goes through. This is that middle step, so it is
// written once. (The sheet reader still asks only for a cached drawing.)

import type { PoNode } from '@/core/po-helpers';
import { DiagramColors } from '@/core/drawingml/diagram/colors';
import { DiagramData } from '@/core/drawingml/diagram/data-model';
import { layoutDiagram } from '@/core/drawingml/diagram/layout-engine';
import { diagramDrawingXml } from '@/core/drawingml/diagram/to-drawing';
import { poChildren, poFindDescendant, poIs } from '@/core/po-helpers';

/** The parts a diagram is run from, as the reader resolved them. */
export interface DiagramParts {
  readonly data: Uint8Array;
  readonly layout?: Uint8Array | undefined;
  readonly colors?: Uint8Array | undefined;
}

/**
 * The picture a drawing part caches, when it caches one at all.
 *
 * A generator may write the drawing part and put nothing in it — the group
 * properties and no `dsp:sp` at all. That is a placeholder the generator never
 * got round to filling, not a diagram drawn as empty, and taking it at its word
 * paints a blank frame over a diagram whose data, layout and colours are all
 * present. So an empty tree reads as no tree, and the caller runs the layout
 * exactly as it does for a file that ships no drawing part in the first place.
 *
 * @param drawing The drawing part's bytes.
 * @param parse   The reader's XML parser.
 * @returns The cached `dsp:spTree`, or undefined when there is none to use.
 */
export function cachedDiagramTree(
  drawing: Uint8Array,
  parse: (bytes: Uint8Array) => Array<PoNode>,
): PoNode | undefined {
  for (const root of parse(drawing)) {
    const found = poFindDescendant(root, 'dsp:spTree');
    if (found && poChildren(found).some((c) => poIs(c, 'dsp:sp'))) return found;
  }
  return undefined;
}

/**
 * The `dsp:spTree` a diagram's own parts describe, for a file that ships no
 * cached drawing.
 *
 * @param parts The diagram's data, layout and colours parts.
 * @param frame The extent the drawing gives the diagram, in EMU.
 * @param parse The reader's XML parser (the same tree shape either way).
 * @returns The shape tree, or undefined when there is no layout part or the
 *          layout uses an algorithm this engine does not run — the caller then
 *          reports the diagram as a loss exactly as before.
 */
export function laidOutDiagramTree(
  parts: DiagramParts,
  frame: { readonly cx: number; readonly cy: number },
  parse: (bytes: Uint8Array) => Array<PoNode>,
): PoNode | undefined {
  if (!parts.layout) return undefined;
  const nodes = layoutDiagram(
    parse(parts.layout),
    new DiagramData(parse(parts.data)),
    frame.cx,
    frame.cy,
  );
  if (nodes.length === 0) return undefined;
  const colors = parts.colors ? new DiagramColors(parse(parts.colors)) : undefined;
  const xml = diagramDrawingXml(nodes, frame, colors);
  for (const root of parse(new TextEncoder().encode(xml))) {
    const found = poFindDescendant(root, 'dsp:spTree');
    if (found) return found;
  }
  return undefined;
}
