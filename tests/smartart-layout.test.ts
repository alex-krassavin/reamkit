// ECMA-376 §21.4.3/§21.4.5 — running a SmartArt LAYOUT part and colouring what
// it lays out, for the files that carry no cached drawing. The layout is a
// program, and these are the shapes of it this engine answers: a flat list, a
// cell split in two beside itself, a cell split in two above itself, and the
// `maxDepth` branch that decides which of the last two a part means. The
// colours part then gives each box its place in a run of accents.

import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import type { PoNode } from '@/core/po-helpers';
import { DiagramColors } from '@/core/drawingml/diagram/colors';
import { DiagramData } from '@/core/drawingml/diagram/data-model';
import { layoutDiagram } from '@/core/drawingml/diagram/layout-engine';
import { cachedDiagramTree } from '@/core/drawingml/diagram/run';
import { diagramDrawingXml } from '@/core/drawingml/diagram/to-drawing';

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
     <dgm:constr type="primFontSz" for="des" forName="parentText" op="equ" val="65"/>
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
           <dgm:layoutNode name="descendantText">
             <dgm:alg type="tx"><dgm:param type="stBulletLvl" val="1"/></dgm:alg>
             <dgm:shape type="rect"/>
             <dgm:constrLst><dgm:constr type="secFontSz" val="40"/></dgm:constrLst>
           </dgm:layoutNode>
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
       <dgm:layoutNode name="parTx">
         <dgm:alg type="tx"/><dgm:shape type="rect"/>
         <dgm:constrLst><dgm:constr type="h" refType="w" op="lte" fact="0.4"/></dgm:constrLst>
       </dgm:layoutNode>
       <dgm:layoutNode name="desTx"><dgm:alg type="tx"/><dgm:shape type="rect"/></dgm:layoutNode>
     </dgm:layoutNode>
   </dgm:forEach>`,
);

// The same without the ceiling the label puts on its own height, to show what
// the height constraints alone would have given it.
const COLOUR_LIST_UNCAPPED = COLOUR_LIST.replace(
  '<dgm:constrLst><dgm:constr type="h" refType="w" op="lte" fact="0.4"/></dgm:constrLst>',
  '',
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

// A process arrow with the steps standing along it: the top node PLACES two
// children — a background arrow across the middle and the row of steps over it
// — and each step is itself placed, its words above the line and a circle on
// it. Nothing about this is a list of boxes.
const ARROW_PROCESS = layoutXml(
  `<dgm:alg type="composite"/>
   <dgm:constrLst>
     <dgm:constr type="w" for="ch" forName="arrow" refType="w"/>
     <dgm:constr type="h" for="ch" forName="arrow" refType="h" fact="0.4"/>
     <dgm:constr type="ctrY" for="ch" forName="arrow" refType="h" fact="0.5"/>
     <dgm:constr type="l" for="ch" forName="arrow"/>
     <dgm:constr type="w" for="ch" forName="points" refType="w" fact="0.9"/>
     <dgm:constr type="h" for="ch" forName="points" refType="h"/>
     <dgm:constr type="t" for="ch" forName="points"/>
     <dgm:constr type="l" for="ch" forName="points"/>
   </dgm:constrLst>
   <dgm:layoutNode name="arrow" styleLbl="bgShp">
     <dgm:alg type="sp"/><dgm:shape type="notchedRightArrow"/>
   </dgm:layoutNode>
   <dgm:layoutNode name="points">
     <dgm:alg type="lin"/>
     <dgm:constrLst>
       <dgm:constr type="w" for="ch" forName="step" refType="w"/>
     </dgm:constrLst>
     <dgm:forEach name="each" axis="ch" ptType="node">
       <dgm:choose name="side">
         <dgm:if name="odd" axis="self" ptType="node" func="posOdd" op="equ" val="1">
           <dgm:layoutNode name="step">
             <dgm:alg type="composite"/>
             <dgm:constrLst>
               <dgm:constr type="w" for="ch" forName="words" refType="w"/>
               <dgm:constr type="h" for="ch" forName="words" refType="h" fact="0.4"/>
               <dgm:constr type="t" for="ch" forName="words"/>
               <dgm:constr type="h" for="ch" forName="dot" refType="h" fact="0.1"/>
               <dgm:constr type="w" for="ch" forName="dot" refType="h" refFor="ch" refForName="dot" op="equ"/>
               <dgm:constr type="ctrY" for="ch" forName="dot" refType="h" fact="0.5"/>
               <dgm:constr type="ctrX" for="ch" forName="dot" refType="w" refFor="ch" refForName="words" fact="0.5"/>
             </dgm:constrLst>
             <dgm:layoutNode name="words"><dgm:alg type="tx"/><dgm:shape type="rect"/></dgm:layoutNode>
             <dgm:layoutNode name="dot"><dgm:alg type="sp"/><dgm:shape type="ellipse"/></dgm:layoutNode>
           </dgm:layoutNode>
         </dgm:if>
       </dgm:choose>
     </dgm:forEach>
   </dgm:layoutNode>`,
);

// A cycle: the children round a circle. The inner `choose` asks how many
// children the FIRST child has, which is what decides where the ring starts.
const CYCLE = layoutXml(
  `<dgm:choose name="where">
     <dgm:if name="hub" axis="ch ch" ptType="node node" st="1 1" cnt="1 0" func="cnt" op="lte" val="1">
       <dgm:alg type="cycle"><dgm:param type="stAng" val="90"/><dgm:param type="spanAng" val="360"/></dgm:alg>
     </dgm:if>
     <dgm:else name="round">
       <dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg>
     </dgm:else>
   </dgm:choose>
   <dgm:constrLst>
     <dgm:constr type="w" for="ch" forName="node" refType="w" fact="0.25"/>
     <dgm:constr type="h" for="ch" forName="node" refType="h" fact="0.25"/>
   </dgm:constrLst>
   <dgm:forEach name="each" axis="ch" ptType="node">
     <dgm:layoutNode name="node"><dgm:alg type="tx"/><dgm:shape type="ellipse"/></dgm:layoutNode>
   </dgm:forEach>`,
);

// The radial variant: the first child is the hub in the middle and its own
// children are the ring around it.
const RADIAL = layoutXml(
  `<dgm:alg type="cycle">
     <dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/>
     <dgm:param type="ctrShpMap" val="fNode"/>
   </dgm:alg>
   <dgm:constrLst>
     <dgm:constr type="w" for="ch" forName="node" refType="w" fact="0.25"/>
     <dgm:constr type="h" for="ch" forName="node" refType="h" fact="0.25"/>
   </dgm:constrLst>
   <dgm:forEach name="first" axis="ch" ptType="node" cnt="1">
     <dgm:layoutNode name="centerShape"><dgm:alg type="tx"/><dgm:shape type="ellipse"/></dgm:layoutNode>
   </dgm:forEach>
   <dgm:forEach name="each" axis="ch" ptType="node">
     <dgm:layoutNode name="node"><dgm:alg type="tx"/><dgm:shape type="ellipse"/></dgm:layoutNode>
   </dgm:forEach>`,
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

  it('takes the point size each box is given, and bullets what the tx algorithm marks', () => {
    const boxes = run(VERTICAL_BLOCK_LIST, dataXml([['a', ['b', 'c']]]));
    // The label's size is stated for every descendant of that name; the
    // descendants box states its own, on itself, with no `forName`.
    expect(boxes[0]?.fontSizePt).toBe(65);
    expect(boxes[1]?.fontSizePt).toBe(40);
    expect(boxes[0]?.bulleted).toBeUndefined();
    expect(boxes[1]?.bulleted).toBe(true);
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
      COLOUR_LIST_UNCAPPED,
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

  it('holds the label to the ceiling it puts on its own height', () => {
    const boxes = run(
      COLOUR_LIST,
      dataXml([
        ['a', ['b']],
        ['c', ['d']],
      ]),
    );
    // Two columns and the 0.14 gap between them: each is CX/2.14 wide, and the
    // label is at most four tenths of that however tall the column is — where
    // the height constraints alone would have given it two fifths of CY.
    const w = CX / 2.14;
    expect(boxes[0]?.cx).toBeCloseTo(w, 3);
    expect(boxes[0]?.cy).toBeCloseTo(w * 0.4, 3);
    expect(boxes[0]?.cy).toBeLessThan(CY * 0.4);
    // What the label does not take goes to the box below it; the column still
    // fills the cell.
    expect(boxes[1]?.y).toBeCloseTo(w * 0.4, 3);
    expect((boxes[1]?.cy ?? 0) + (boxes[0]?.cy ?? 0)).toBeCloseTo(CY, 3);
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

  it('places what a composite places, and runs each child inside it', () => {
    const boxes = run(
      ARROW_PROCESS,
      dataXml([
        ['a', []],
        ['b', []],
      ]),
    );
    // The background arrow: the whole width, two fifths of the height, centred.
    const arrow = boxes.find((b) => b.shapeType === 'notchedRightArrow');
    expect(arrow).toMatchObject({ x: 0, cx: CX, cy: CY * 0.4, styleLbl: 'bgShp' });
    expect(arrow?.y).toBeCloseTo(CY * 0.3, 3);
    // ...and the steps over it, inside the nine tenths of the width they were
    // given, each a `tx` box with the node's words above the line.
    expect(boxes.filter((b) => b.shapeType === 'rect').map(label)).toEqual(['a', 'b']);
    // Two cells and the default tenth-of-a-cell gap between them, inside the
    // nine tenths of the width the composite gave the row.
    const words = boxes.find((b) => b.shapeType === 'rect');
    expect(words?.cx).toBeCloseTo((CX * 0.9) / 2.1, 3);
    expect(words?.cy).toBeCloseTo(CY * 0.4, 3);
    expect(words?.y).toBe(0);
    // The dot is round — its width is stated against its own height — and sits
    // on the line, centred under the middle of the words above it.
    const dot = boxes.find((b) => b.shapeType === 'ellipse');
    expect(dot?.cx).toBeCloseTo(CY * 0.1, 3);
    expect(dot?.cy).toBeCloseTo(CY * 0.1, 3);
    expect((dot?.y ?? 0) + (dot?.cy ?? 0) / 2).toBeCloseTo(CY * 0.5, 3);
    expect((dot?.x ?? 0) + (dot?.cx ?? 0) / 2).toBeCloseTo((words?.cx ?? 0) / 2, 3);
    // A shape laid out by `sp` carries no words of its own.
    expect(label(dot as never)).toBe('');
  });

  it('puts a cycle round a circle, starting where the layout says', () => {
    // Four nodes, none with children of its own: the `cnt` over the first
    // child's children is zero, the branch holds, and the ring starts at three
    // o'clock as that branch asks.
    const boxes = run(
      CYCLE,
      dataXml([
        ['a', []],
        ['b', []],
        ['c', []],
        ['d', []],
      ]),
    );
    expect(boxes.map(label)).toEqual(['a', 'b', 'c', 'd']);
    const mid = (b: (typeof boxes)[number]): [number, number] => [b.x + b.cx / 2, b.y + b.cy / 2];
    // A quarter of the frame each, and a full turn in four steps about the
    // middle — the ring as wide as the shorter side allows.
    expect(boxes[0]?.cx).toBeCloseTo(CX * 0.25, 3);
    expect(boxes[0]?.cy).toBeCloseTo(CY * 0.25, 3);
    const r = (CY - CY * 0.25) / 2;
    expect(mid(boxes[0]!)[0]).toBeCloseTo(CX / 2 + r, 3);
    expect(mid(boxes[0]!)[1]).toBeCloseTo(CY / 2, 3);
    expect(mid(boxes[1]!)[1]).toBeCloseTo(CY / 2 + r, 3);
    expect(mid(boxes[2]!)[0]).toBeCloseTo(CX / 2 - r, 3);
    expect(mid(boxes[3]!)[1]).toBeCloseTo(CY / 2 - r, 3);
  });

  it('answers a `cnt` over a path, which is what moves the ring round', () => {
    // The same layout over data whose first child HAS children: the branch no
    // longer holds, and the ring starts at twelve o'clock instead.
    const boxes = run(
      CYCLE,
      dataXml([
        ['a', ['x', 'y']],
        ['b', []],
      ]),
    );
    const first = boxes[0]!;
    expect(first.x + first.cx / 2).toBeCloseTo(CX / 2, 3);
    expect(first.y + first.cy / 2).toBeLessThan(CY / 2);
  });

  it('puts the hub in the middle and its own children round it', () => {
    const boxes = run(RADIAL, dataXml([['hub', ['a', 'b']]]));
    expect(boxes.map(label)).toEqual(['hub', 'a', 'b']);
    // The hub sits at the centre of the frame; the ring is its children.
    expect(boxes[0]?.x).toBeCloseTo((CX - CX * 0.25) / 2, 3);
    expect(boxes[0]?.y).toBeCloseTo((CY - CY * 0.25) / 2, 3);
    expect(boxes[1]?.y).toBeLessThan(boxes[0]?.y ?? 0);
    expect(boxes[2]?.y).toBeGreaterThan(boxes[0]?.y ?? 0);
  });

  it('leaves a box out for an algorithm it cannot run inside a composite', () => {
    // A centre cycle is a composite of a `cycle` and its middle. Drawing the
    // cycle's whole share as one box put a rectangle over the entire frame.
    const centre = layoutXml(
      `<dgm:alg type="composite"/>
       <dgm:constrLst><dgm:constr type="w" for="ch" forName="ring" refType="w"/></dgm:constrLst>
       <dgm:layoutNode name="ring"><dgm:alg type="pyramid"/><dgm:shape type="ellipse"/></dgm:layoutNode>`,
    );
    expect(run(centre, dataXml([['a', []]]))).toEqual([]);
  });
});

describe('the drawing part the engine writes', () => {
  const sizesIn = (xml: string): Array<number> =>
    [...xml.matchAll(/sz="(\d+)"/gu)].map((m) => Number(m[1]) / 100);

  it('writes the size the layout states, and asks the layout to fit it', () => {
    const xml = diagramDrawingXml(run(VERTICAL_BLOCK_LIST, dataXml([['a', ['b', 'c']]])), {
      cx: CX,
      cy: CY,
    });
    expect(sizesIn(xml)).toEqual([65, 40, 40]);
    // The stated size is a maximum; the shrink needs glyphs, so every box asks
    // for it rather than guessing here.
    expect(xml.match(/<a:normAutofit\/>/gu)).toHaveLength(2);
  });

  it('carries a break inside a paragraph rather than running the halves together', () => {
    const withBreak = dataXml([['a', []]]).replace(
      '<a:r><a:t>a</a:t></a:r>',
      '<a:r><a:t>Max size</a:t></a:r><a:br/><a:r><a:t>(65 pt)</a:t></a:r>',
    );
    const xml = diagramDrawingXml(run(VERTICAL_BLOCK_LIST, withBreak), { cx: CX, cy: CY });
    expect(xml).toContain('<a:t>Max size</a:t></a:r><a:br/><a:r>');
  });

  it('bullets and ranges left only the box the layout marks', () => {
    const xml = diagramDrawingXml(run(VERTICAL_BLOCK_LIST, dataXml([['a', ['b']]])), {
      cx: CX,
      cy: CY,
    });
    expect(xml).toContain('<a:buChar char="\u2022"/>');
    expect(xml).toContain('<a:pPr algn="ctr"/>');
  });
});

const COLORS = `<dgm:colorsDef xmlns:dgm="d" xmlns:a="a" uniqueId="urn:x">
  <dgm:styleLbl name="alignNode1">
    <dgm:fillClrLst meth="repeat">
      <a:schemeClr val="accent2"/><a:schemeClr val="accent3"/>
    </dgm:fillClrLst>
    <dgm:txFillClrLst/>
  </dgm:styleLbl>
  <dgm:styleLbl name="alignAccFollowNode1">
    <dgm:fillClrLst meth="repeat">
      <a:schemeClr val="accent2"><a:tint val="40000"/><a:alpha val="90000"/></a:schemeClr>
      <a:schemeClr val="accent3"><a:tint val="40000"/><a:alpha val="90000"/></a:schemeClr>
    </dgm:fillClrLst>
    <dgm:txFillClrLst meth="repeat"><a:schemeClr val="dk1"/></dgm:txFillClrLst>
  </dgm:styleLbl>
</dgm:colorsDef>`;

describe('the colours a SmartArt part gives its boxes', () => {
  const colors = new DiagramColors(parse(COLORS));

  it('walks the run, so siblings differ and the run starts over', () => {
    expect(colors.fill('alignNode1', 0)).toBe(
      '<a:solidFill><a:schemeClr val="accent2"></a:schemeClr></a:solidFill>',
    );
    expect(colors.fill('alignNode1', 1)).toBe(
      '<a:solidFill><a:schemeClr val="accent3"></a:schemeClr></a:solidFill>',
    );
    expect(colors.fill('alignNode1', 2)).toBe(colors.fill('alignNode1', 0));
  });

  it('keeps the transforms that make a follower box a wash of the accent', () => {
    expect(colors.fill('alignAccFollowNode1', 0)).toBe(
      '<a:solidFill><a:schemeClr val="accent2"><a:tint val="40000"/><a:alpha val="90000"/></a:schemeClr></a:solidFill>',
    );
  });

  it('gives a label it knows but leaves unfilled no fill at all', () => {
    // `revTx` — the reversed text of a process arrow — is named with an empty
    // fill list, which is a transparent box, not the accent a file with no
    // colours part at all falls back to.
    const withRevTx = new DiagramColors(
      parse(
        COLORS.replace(
          '</dgm:colorsDef>',
          '<dgm:styleLbl name="revTx"><dgm:fillClrLst/></dgm:styleLbl></dgm:colorsDef>',
        ),
      ),
    );
    expect(withRevTx.knows('revTx')).toBe(true);
    expect(withRevTx.fill('revTx', 0)).toBeUndefined();
    expect(withRevTx.knows('neverHeardOfIt')).toBe(false);
  });

  it('overrides the text colour only where the part states one', () => {
    expect(colors.textFill('alignAccFollowNode1', 0)).toBe(
      '<a:solidFill><a:schemeClr val="dk1"></a:schemeClr></a:solidFill>',
    );
    // An empty list is the style's own font colour, not a colour of its own.
    expect(colors.textFill('alignNode1', 0)).toBeUndefined();
    expect(colors.fill('noSuchLabel', 0)).toBeUndefined();
    expect(colors.fill(undefined, 0)).toBeUndefined();
  });
});

// §21.4.3.7/§21.4.3.9 and §21.4.4 — what a box takes from the file rather than
// from the layout's defaults: the descendants it speaks for, the geometry it
// hides, the bullet its own data states, and the picture a drawing part is
// allowed to have cached as nothing at all.
describe('SmartArt: what the file states for itself', () => {
  // A picture list's node: it speaks for its point and everything under it, and
  // it is spaced by a box that reserves room without being drawn.
  const PICTURE_LIST = layoutXml(
    `<dgm:alg type="lin"/>
     <dgm:constrLst>
       <dgm:constr type="w" for="ch" forName="cell" refType="w"/>
     </dgm:constrLst>
     <dgm:forEach name="each" axis="ch" ptType="node">
       <dgm:layoutNode name="cell">
         <dgm:alg type="composite"/>
         <dgm:constrLst>
           <dgm:constr type="h" for="ch" forName="spacer" refType="h" fact="0.1"/>
           <dgm:constr type="t" for="ch" forName="spacer"/>
           <dgm:constr type="h" for="ch" forName="label" refType="h" fact="0.9"/>
           <dgm:constr type="b" for="ch" forName="label" refType="h"/>
         </dgm:constrLst>
         <dgm:layoutNode name="label">
           <dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst>
           <dgm:alg type="tx"><dgm:param type="txAnchorVert" val="t"/></dgm:alg>
           <dgm:shape type="rect"/>
           <dgm:presOf axis="desOrSelf" ptType="node"/>
         </dgm:layoutNode>
         <dgm:layoutNode name="spacer">
           <dgm:alg type="sp"/>
           <dgm:shape type="roundRect" hideGeom="1"/>
         </dgm:layoutNode>
       </dgm:layoutNode>
     </dgm:forEach>`,
  );

  it('a box with presOf desOrSelf speaks for its descendants too', () => {
    const nodes = run(PICTURE_LIST, dataXml([['Parent', ['Kid1', 'Kid2']]]));
    const spoken = nodes.find((n) => n.point.text !== undefined);
    const xml = diagramDrawingXml(nodes, { cx: CX, cy: CY });
    // Its own label first, then each descendant's, in depth-first order.
    expect(xml.indexOf('Parent')).toBeGreaterThan(-1);
    expect(xml.indexOf('Kid1')).toBeGreaterThan(xml.indexOf('Parent'));
    expect(xml.indexOf('Kid2')).toBeGreaterThan(xml.indexOf('Kid1'));
    expect(spoken?.anchor).toBe('t');
  });

  it('a hidden geometry reserves its room and paints nothing', () => {
    const nodes = run(PICTURE_LIST, dataXml([['Parent', []]]));
    const spacer = nodes.find((n) => n.hideGeom === true);
    expect(spacer).toBeDefined();
    // It is a real box — the label below it starts where the spacer ends.
    expect(spacer?.cy).toBeGreaterThan(0);
    const xml = diagramDrawingXml(nodes, { cx: CX, cy: CY });
    expect(xml).toContain('<a:noFill/><a:ln><a:noFill/></a:ln>');
  });

  it("a paragraph's own bullet wins over the layout's permission to have one", () => {
    const data = dataXml([['Parent', ['Kid']]])
      .replace('<a:p><a:r><a:t>Parent', '<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:t>Parent')
      .replace(
        '<a:p><a:r><a:t>Kid',
        '<a:p><a:pPr algn="l"><a:buChar char="-"/></a:pPr><a:r><a:t>Kid',
      );
    const xml = diagramDrawingXml(run(PICTURE_LIST, data), { cx: CX, cy: CY });
    // The heading says it takes no bullet, so it takes none; the child names
    // its own character rather than the engine's default.
    expect(xml).toContain('<a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"/><a:t>Parent</a:t>');
    expect(xml).toContain('<a:buChar char="-"/>');
    expect(xml).not.toContain('<a:buChar char="•"/>');
  });

  it("carries the author's own bold through, but not their size", () => {
    const data = dataXml([['Parent', []]]).replace(
      '<a:r><a:t>Parent',
      '<a:r><a:rPr b="1" sz="900"/><a:t>Parent',
    );
    const xml = diagramDrawingXml(run(PICTURE_LIST, data), { cx: CX, cy: CY });
    expect(xml).toContain(' b="1"');
    expect(xml).not.toContain('sz="900"');
  });
});

// §21.4.4 — the drawing part a file caches, and what it means when it caches
// nothing. smartart-missing-bullet.pptx ships a `dsp:drawing` holding only the
// group properties: the generator wrote the part and never filled it.
describe('SmartArt: an empty cached drawing is no drawing', () => {
  const EMPTY = `<dsp:drawing xmlns:dsp="s" xmlns:a="a"><dsp:spTree>
      <dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>
      <dsp:grpSpPr/></dsp:spTree></dsp:drawing>`;
  const DRAWN = EMPTY.replace(
    '</dsp:spTree>',
    '<dsp:sp><dsp:spPr/><dsp:txBody><a:p/></dsp:txBody></dsp:sp></dsp:spTree>',
  );
  const bytes = (xml: string): Uint8Array => new TextEncoder().encode(xml);
  // `cachedDiagramTree` takes the reader's own parser, which reads bytes.
  const parseBytes = (b: Uint8Array): Array<PoNode> => parse(new TextDecoder().decode(b));

  it('reads a cached tree that holds shapes', () => {
    expect(cachedDiagramTree(bytes(DRAWN), parseBytes)).toBeDefined();
  });

  it('reads a cached tree that holds none as absent, so the layout runs', () => {
    expect(cachedDiagramTree(bytes(EMPTY), parseBytes)).toBeUndefined();
    expect(cachedDiagramTree(bytes('<a:notADrawing xmlns:a="a"/>'), parseBytes)).toBeUndefined();
  });
});

// §21.4.5 `dgm:linClrLst` — the outline the colours part states for a label.
describe('SmartArt: the colours part draws outlines too', () => {
  const OUTLINED = `<dgm:colorsDef xmlns:dgm="d" xmlns:a="a">
      <dgm:styleLbl name="conFgAcc1">
        <dgm:fillClrLst meth="repeat"><a:schemeClr val="lt1"><a:alpha val="90000"/></a:schemeClr></dgm:fillClrLst>
        <dgm:linClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:linClrLst>
      </dgm:styleLbl>
      <dgm:styleLbl name="plain"><dgm:fillClrLst><a:schemeClr val="accent2"/></dgm:fillClrLst></dgm:styleLbl>
    </dgm:colorsDef>`;
  const outlined = new DiagramColors(parse(OUTLINED));

  it('gives a follower box its line, which is the whole of it', () => {
    // Near-white inside an accent line: unlined it is white on white.
    expect(outlined.fill('conFgAcc1', 0)).toContain('val="lt1"');
    expect(outlined.line('conFgAcc1', 0)).toBe(
      '<a:ln><a:solidFill><a:schemeClr val="accent1"></a:schemeClr></a:solidFill></a:ln>',
    );
  });

  it('leaves a label that states no line without one', () => {
    expect(outlined.line('plain', 0)).toBeUndefined();
    expect(outlined.line('noSuchLabel', 0)).toBeUndefined();
    expect(outlined.line(undefined, 0)).toBeUndefined();
  });
});

// §21.4.3.2 — the vertical list family insets its label with a hidden margin
// box, which is spacing and not one half of a two-way split.
describe('SmartArt: a hidden margin is not half of a split', () => {
  const INSET_LABEL = layoutXml(
    `<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg>
     <dgm:constrLst>
       <dgm:constr type="w" for="ch" forName="row" refType="w"/>
       <dgm:constr type="w" for="des" forName="leftMargin" refType="w" fact="0.05"/>
       <dgm:constr type="w" for="des" forName="label" refType="w" fact="0.7"/>
     </dgm:constrLst>
     <dgm:forEach name="each" axis="ch" ptType="node">
       <dgm:layoutNode name="row">
         <dgm:alg type="lin"><dgm:param type="linDir" val="fromL"/></dgm:alg>
         <dgm:constrLst/>
         <dgm:layoutNode name="leftMargin">
           <dgm:alg type="sp"/><dgm:shape type="rect" hideGeom="1"/>
         </dgm:layoutNode>
         <dgm:layoutNode name="label" styleLbl="node1">
           <dgm:alg type="tx"/><dgm:shape type="roundRect"/>
         </dgm:layoutNode>
       </dgm:layoutNode>
     </dgm:forEach>`,
  );

  it('gives the cell to the one box that holds the words', () => {
    const nodes = run(INSET_LABEL, dataXml([['P1', []]]));
    // One box, not two: split against the margin, the label's own words went
    // into the invisible half and the diagram came out blank.
    expect(nodes).toHaveLength(1);
    expect(nodes[0]?.styleLbl).toBe('node1');
    expect(nodes[0]?.shapeType).toBe('roundRect');
    expect(nodes[0]?.hideGeom).toBeUndefined();
    expect(diagramDrawingXml(nodes, { cx: CX, cy: CY })).toContain('<a:t>P1</a:t>');
  });
});

// §21.4.3 — a forEach body of several boxes is a stack, one box per point per
// entry, and §21.4.3.7 `presOf axis="des"` puts the descendants in the second.
describe('SmartArt: a forEach body of several boxes', () => {
  const STACK = layoutXml(
    `<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg>
     <dgm:constrLst>
       <dgm:constr type="w" for="ch" forName="row" refType="w"/>
       <dgm:constr type="h" for="des" forName="label" refType="primFontSz" refFor="des" refForName="label" fact="0.8"/>
       <dgm:constr type="h" for="ch" forName="gap" refType="primFontSz" refFor="des" refForName="label" fact="-0.4"/>
       <dgm:constr type="h" for="ch" forName="kids" refType="primFontSz" refFor="des" refForName="label" fact="0.4"/>
     </dgm:constrLst>
     <dgm:forEach name="each" axis="ch" ptType="node">
       <dgm:layoutNode name="row">
         <dgm:alg type="lin"/><dgm:constrLst/>
         <dgm:layoutNode name="label" styleLbl="node1">
           <dgm:alg type="tx"/><dgm:shape type="roundRect"/>
           <dgm:presOf axis="self" ptType="node"/>
         </dgm:layoutNode>
       </dgm:layoutNode>
       <dgm:layoutNode name="gap"><dgm:alg type="sp"/><dgm:shape/></dgm:layoutNode>
       <dgm:layoutNode name="kids" styleLbl="conFgAcc1">
         <dgm:alg type="tx"/><dgm:shape type="rect"/>
         <dgm:presOf axis="des" ptType="node"/>
       </dgm:layoutNode>
     </dgm:forEach>`,
  );

  it('lays out every box of the body, not only the first', () => {
    const nodes = run(STACK, dataXml([['P1', ['K1']]]));
    // The label and the descendants box; the negative space draws nothing.
    expect(nodes.map((n) => n.styleLbl)).toEqual(['node1', 'conFgAcc1']);
    const [chip, kids] = nodes;
    // Stacked, in the proportion their heights state (0.8 and 0.4 of the
    // stated font size); the negative space is an overlap, so it takes none.
    expect(chip?.y).toBe(0);
    expect(kids?.y).toBeCloseTo(CY * (0.8 / 1.2), 0);
    expect((chip?.cy ?? 0) / (kids?.cy ?? 1)).toBeCloseTo(2, 5);
  });

  it('puts the point in the first box and its descendants in the second', () => {
    const xml = diagramDrawingXml(run(STACK, dataXml([['P1', ['K1']]])), { cx: CX, cy: CY });
    expect(xml.indexOf('<a:t>P1</a:t>')).toBeGreaterThan(-1);
    // Not a second copy of the label: the descendants box holds only the kids.
    expect(xml.indexOf('<a:t>K1</a:t>')).toBeGreaterThan(xml.indexOf('<a:t>P1</a:t>'));
    expect(xml.split('<a:t>P1</a:t>')).toHaveLength(2);
  });
});

// §21.4.3.1 — widths that do not add up to the row are placements, not shares.
describe('SmartArt: a wrapper insets its boxes', () => {
  const INSET = layoutXml(
    `<dgm:alg type="lin"><dgm:param type="linDir" val="fromT"/></dgm:alg>
     <dgm:constrLst>
       <dgm:constr type="w" for="ch" forName="row" refType="w"/>
       <dgm:constr type="w" for="des" forName="margin" refType="w" fact="0.05"/>
       <dgm:constr type="w" for="des" forName="label" refType="w" fact="0.7"/>
     </dgm:constrLst>
     <dgm:forEach name="each" axis="ch" ptType="node">
       <dgm:layoutNode name="row">
         <dgm:alg type="lin"><dgm:param type="linDir" val="fromL"/></dgm:alg>
         <dgm:constrLst/>
         <dgm:layoutNode name="margin"><dgm:alg type="sp"/><dgm:shape type="rect" hideGeom="1"/></dgm:layoutNode>
         <dgm:layoutNode name="label" styleLbl="node1">
           <dgm:alg type="tx"><dgm:param type="parTxLTRAlign" val="l"/></dgm:alg>
           <dgm:shape type="roundRect"/>
         </dgm:layoutNode>
       </dgm:layoutNode>
     </dgm:forEach>`,
  );

  it('places each box where its stated width puts it', () => {
    const nodes = run(INSET, dataXml([['P1', []]]));
    expect(nodes).toHaveLength(1);
    // A twentieth in from the left and seven tenths across; the last quarter of
    // the row is stated by nothing and stays empty.
    expect(nodes[0]?.x).toBeCloseTo(CX * 0.05, 5);
    expect(nodes[0]?.cx).toBeCloseTo(CX * 0.7, 5);
  });

  it("ranges the label the way the layout's own parameter asks", () => {
    const xml = diagramDrawingXml(run(INSET, dataXml([['P1', []]])), { cx: CX, cy: CY });
    expect(xml).toContain('<a:pPr algn="l"/>');
  });
});
