// ECMA-376 §21.4.3 — running a SmartArt LAYOUT part, for the files that carry
// no cached drawing. The layout is a program, and these are the shapes of it
// this engine answers: a flat list, a cell split in two beside itself, a cell
// split in two above itself, and the `maxDepth` branch that decides which of
// the last two a part means.

import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import type { PoNode } from '@/core/po-helpers';
import { DiagramData } from '@/core/drawingml/diagram/data-model';
import { layoutDiagram } from '@/core/drawingml/diagram/layout-engine';

const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
});
const parse = (xml: string): Array<PoNode> => parser.parse(xml) as Array<PoNode>;

const CX = 6_000_000;
const CY = 3_000_000;

/** A data part: `nodes` is a list of labels, each with its own child labels. */
function dataXml(nodes: ReadonlyArray<readonly [string, ReadonlyArray<string>]>): string {
  const pts: Array<string> = ['<dgm:pt modelId="doc" type="doc"/>'];
  const cxns: Array<string> = [];
  nodes.forEach(([label, kids], i) => {
    pts.push(
      `<dgm:pt modelId="${label}"><dgm:t><a:p><a:r><a:t>${label}</a:t></a:r></a:p></dgm:t></dgm:pt>`,
    );
    cxns.push(`<dgm:cxn modelId="c${label}" srcId="doc" destId="${label}" srcOrd="${i}"/>`);
    kids.forEach((kid, j) => {
      pts.push(
        `<dgm:pt modelId="${kid}"><dgm:t><a:p><a:r><a:t>${kid}</a:t></a:r></a:p></dgm:t></dgm:pt>`,
      );
      cxns.push(`<dgm:cxn modelId="c${kid}" srcId="${label}" destId="${kid}" srcOrd="${j}"/>`);
    });
  });
  return `<dgm:dataModel xmlns:dgm="d" xmlns:a="a"><dgm:ptLst>${pts.join('')}</dgm:ptLst><dgm:cxnLst>${cxns.join('')}</dgm:cxnLst></dgm:dataModel>`;
}

const layoutXml = (body: string): string =>
  `<dgm:layoutDef xmlns:dgm="d"><dgm:layoutNode name="top">${body}</dgm:layoutNode></dgm:layoutDef>`;

const run = (layout: string, data: string): ReturnType<typeof layoutDiagram> =>
  layoutDiagram(parse(layout), new DiagramData(parse(data)), CX, CY);

// The Basic Block List: a snake of boxes with a tenth of a box between them.
const BLOCK_LIST = layoutXml(
  `<dgm:alg type="snake"/>
   <dgm:constrLst>
     <dgm:constr type="w" for="ch" forName="node" refType="w"/>
     <dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/>
     <dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.1"/>
   </dgm:constrLst>
   <dgm:forEach name="each" axis="ch" ptType="node">
     <dgm:layoutNode name="node"><dgm:alg type="tx"/><dgm:shape type="rect"/></dgm:layoutNode>
   </dgm:forEach>`,
);

// The Vertical Block List: a column of rows, each row the node's label at 36%
// of the width beside its children's text at 64%.
const VERTICAL_BLOCK_LIST = layoutXml(
  `<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg>
   <dgm:constrLst>
     <dgm:constr type="h" for="ch" forName="linNode" refType="h"/>
     <dgm:constr type="w" for="ch" forName="linNode" refType="w"/>
     <dgm:constr type="h" for="ch" forName="sp" refType="h" refFor="ch" refForName="linNode" fact="0.05"/>
   </dgm:constrLst>
   <dgm:forEach name="each" axis="ch" ptType="node">
     <dgm:layoutNode name="linNode">
       <dgm:alg type="lin"/>
       <dgm:constrLst>
         <dgm:constr type="w" for="ch" forName="parentText" refType="w" fact="0.36"/>
         <dgm:constr type="w" for="ch" forName="descendantText" refType="w" fact="0.64"/>
       </dgm:constrLst>
       <dgm:layoutNode name="parentText"><dgm:alg type="tx"/><dgm:shape type="roundRect"/></dgm:layoutNode>
       <dgm:choose name="hasKids">
         <dgm:if name="some" axis="ch" ptType="node" func="cnt" op="gte" val="1">
           <dgm:layoutNode name="descendantText"><dgm:alg type="tx"/><dgm:shape type="rect"/></dgm:layoutNode>
         </dgm:if>
       </dgm:choose>
     </dgm:layoutNode>
   </dgm:forEach>`,
);

// The Vertical Picture List family: a row of columns, each column the label
// above its children's text, in the proportion their heights state.
const COLOUR_LIST = layoutXml(
  `<dgm:alg type="lin"/>
   <dgm:constrLst>
     <dgm:constr type="w" for="ch" forName="composite" refType="w"/>
     <dgm:constr type="h" for="des" forName="parTx" refType="primFontSz" refFor="des" refForName="parTx" fact="0.8"/>
     <dgm:constr type="h" for="des" forName="desTx" refType="primFontSz" refFor="des" refForName="parTx" fact="1.2"/>
     <dgm:constr type="w" for="ch" forName="space" refType="w" refFor="ch" refForName="composite" fact="0.14"/>
   </dgm:constrLst>
   <dgm:forEach name="each" axis="ch" ptType="node">
     <dgm:layoutNode name="composite">
       <dgm:alg type="composite"/>
       <dgm:constrLst>
         <dgm:constr type="w" for="ch" forName="parTx" refType="w"/>
         <dgm:constr type="w" for="ch" forName="desTx" refType="w" refFor="ch" refForName="parTx"/>
       </dgm:constrLst>
       <dgm:layoutNode name="parTx"><dgm:alg type="tx"/><dgm:shape type="rect"/></dgm:layoutNode>
       <dgm:layoutNode name="desTx"><dgm:alg type="tx"/><dgm:shape type="rect"/></dgm:layoutNode>
     </dgm:layoutNode>
   </dgm:forEach>`,
);

// The Chevron Process: two boxes to a cell when the nodes have children, one
// when they do not — and the chevrons overlap, which is the negative gap.
const CHEVRON = layoutXml(
  `<dgm:alg type="lin"/>
   <dgm:choose name="depth">
     <dgm:if name="deep" axis="des" func="maxDepth" op="gte" val="2">
       <dgm:constrLst>
         <dgm:constr type="w" for="ch" forName="composite" refType="w"/>
         <dgm:constr type="h" for="des" forName="parTx" refType="primFontSz" refFor="des" refForName="parTx" fact="1.5"/>
         <dgm:constr type="h" for="des" forName="desTx" refType="primFontSz" refFor="des" refForName="parTx" fact="0.5"/>
       </dgm:constrLst>
       <dgm:forEach name="eachDeep" axis="ch" ptType="node">
         <dgm:layoutNode name="composite">
           <dgm:alg type="composite"/>
           <dgm:constrLst>
             <dgm:constr type="w" for="ch" forName="parTx" refType="w"/>
             <dgm:constr type="w" for="ch" forName="desTx" refType="w" refFor="ch" refForName="parTx" fact="0.8"/>
           </dgm:constrLst>
           <dgm:layoutNode name="parTx"><dgm:alg type="tx"/><dgm:shape type="chevron"/></dgm:layoutNode>
           <dgm:layoutNode name="desTx"><dgm:alg type="tx"/><dgm:shape type="rect"/></dgm:layoutNode>
         </dgm:layoutNode>
       </dgm:forEach>
     </dgm:if>
     <dgm:else name="flat">
       <dgm:constrLst>
         <dgm:constr type="w" for="ch" forName="parTxOnly" refType="w"/>
         <dgm:constr type="w" for="ch" forName="parTxOnlySpace" refType="w" refFor="ch" refForName="parTxOnly" fact="-0.1"/>
       </dgm:constrLst>
       <dgm:forEach name="eachFlat" axis="ch" ptType="node">
         <dgm:layoutNode name="parTxOnly"><dgm:alg type="tx"/><dgm:shape type="chevron"/></dgm:layoutNode>
       </dgm:forEach>
     </dgm:else>
   </dgm:choose>`,
);

const label = (n: { readonly point: Parameters<typeof DiagramData.textOf>[0] }): string =>
  DiagramData.textOf(n.point);

describe('a SmartArt layout with no cached drawing', () => {
  it('puts one box per node and wraps them the way a snake does', () => {
    const boxes = run(
      BLOCK_LIST,
      dataXml([
        ['a', []],
        ['b', []],
        ['c', []],
      ]),
    );
    expect(boxes.map(label)).toEqual(['a', 'b', 'c']);
    expect(boxes.every((b) => b.shapeType === 'rect')).toBe(true);
    // Two across in a 2:1 frame is the count that leaves each cell closest to
    // the 0.6 the constraints ask for; the two cells and the tenth-of-a-cell
    // gap between them fill the width exactly.
    const w = CX / 2.1;
    expect(boxes[0]?.cx).toBeCloseTo(w, 3);
    expect((boxes[1]?.x ?? 0) - (boxes[0]?.x ?? 0)).toBeCloseTo(w * 1.1, 3);
    // The third wraps to a second row, and a short last row is centred.
    expect(boxes[2]?.y ?? 0).toBeGreaterThan(boxes[0]?.y ?? 0);
    expect(boxes[2]?.x).toBeCloseTo((CX - w) / 2, 3);
  });

  it('splits a row into the label and its children beside it', () => {
    const boxes = run(
      VERTICAL_BLOCK_LIST,
      dataXml([
        ['a', ['b', 'c']],
        ['x', ['y']],
      ]),
    );
    expect(boxes.map(label)).toEqual(['a', 'b\nc', 'x', 'y']);
    expect(boxes.map((b) => b.shapeType)).toEqual(['roundRect', 'rect', 'roundRect', 'rect']);
    // 36% beside 64%, side by side on one line.
    expect(boxes[0]?.cx).toBeCloseTo(CX * 0.36, 3);
    expect(boxes[1]?.cx).toBeCloseTo(CX * 0.64, 3);
    expect(boxes[1]?.x).toBeCloseTo(CX * 0.36, 3);
    expect(boxes[1]?.y).toBe(boxes[0]?.y);
    // ...and the second row below the first, not beside it.
    expect(boxes[2]?.y ?? 0).toBeGreaterThan(boxes[0]?.y ?? 0);
    expect(boxes[2]?.x).toBe(boxes[0]?.x);
  });

  it('leaves the second box out where the layout guards it on a child count', () => {
    const boxes = run(
      VERTICAL_BLOCK_LIST,
      dataXml([
        ['a', ['b']],
        ['alone', []],
      ]),
    );
    // A row for a node with nothing under it is its label and nothing beside
    // it — the `cnt` branch does not hold, so that box is never laid out.
    expect(boxes.map(label)).toEqual(['a', 'b', 'alone']);
    // The label keeps its 36%; the row does not stretch to fill the width.
    expect(boxes[2]?.cx).toBeCloseTo(CX * 0.36, 3);
  });

  it('splits a column into the label above its children, in the stated proportion', () => {
    const boxes = run(
      COLOUR_LIST,
      dataXml([
        ['a', ['b']],
        ['c', ['d']],
      ]),
    );
    expect(boxes.map(label)).toEqual(['a', 'b', 'c', 'd']);
    // 0.8 against 1.2 is a ratio between the two, not a share of the cell.
    expect(boxes[0]?.cy).toBeCloseTo(CY * 0.4, 3);
    expect(boxes[1]?.cy).toBeCloseTo(CY * 0.6, 3);
    expect(boxes[1]?.y).toBeCloseTo(CY * 0.4, 3);
    expect(boxes[1]?.x).toBe(boxes[0]?.x);
    expect(boxes[0]?.cx).toBe(boxes[1]?.cx);
  });

  it('leaves the descendants box empty rather than repeating the label', () => {
    const boxes = run(COLOUR_LIST, dataXml([['a', []]]));
    expect(boxes).toHaveLength(2);
    expect(boxes.map(label)).toEqual(['a', '']);
  });

  it('takes the branch the data depth chooses, and overlaps where it says to', () => {
    // Flat data: the `maxDepth gte 2` branch does not hold, so each node is one
    // chevron filling the cell — not a chevron cut to three quarters of it.
    const flat = run(
      CHEVRON,
      dataXml([
        ['a', []],
        ['b', []],
        ['c', []],
      ]),
    );
    expect(flat).toHaveLength(3);
    expect(flat.every((b) => b.cy === CY && b.shapeType === 'chevron')).toBe(true);
    // Three chevrons less two tenths of overlap fill the frame.
    const w = CX / 2.8;
    expect(flat[0]?.cx).toBeCloseTo(w, 3);
    expect((flat[1]?.x ?? 0) - (flat[0]?.x ?? 0)).toBeCloseTo(w * 0.9, 3);

    // Data one level deeper takes the other branch: a label over its children.
    const deep = run(
      CHEVRON,
      dataXml([
        ['a', ['b']],
        ['c', ['d']],
      ]),
    );
    expect(deep.map(label)).toEqual(['a', 'b', 'c', 'd']);
    expect(deep[0]?.cy).toBeCloseTo(CY * 0.75, 3);
    expect(deep[1]?.cy).toBeCloseTo(CY * 0.25, 3);
  });

  it('gives nothing for an algorithm this engine does not run', () => {
    const hierarchy = layoutXml('<dgm:alg type="hierChild"/>');
    expect(run(hierarchy, dataXml([['a', []]]))).toEqual([]);
  });
});
