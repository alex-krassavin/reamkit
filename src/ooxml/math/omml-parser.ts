// ECMA-376 Part 1 §22 — OfficeMathML (OMML) parser.
//
// Turns an <m:oMath> subtree into a recursive MathNode. The element coverage
// grows by milestone; this file currently handles runs (m:r/m:t) and fractions
// (m:f). Unrecognised constructs degrade by recursing into their content so any
// literal symbols still appear.

import type { MathFraction, MathNode, MathRun } from '@/document-model';
import type { PoNode } from '@/ooxml/wordproc/po-helpers';
import {
  poChildren,
  poChildrenWith,
  poIs,
  poTag,
  poText,
  poToggle,
  poVal,
} from '@/ooxml/wordproc/po-helpers';

export function parseOMath(oMath: PoNode): MathNode {
  return { type: 'row', children: parseMathSeq(poChildren(oMath)) };
}

// Parse a sequence of sibling elements into content nodes (skipping property
// elements and nulls).
function parseMathSeq(children: ReadonlyArray<PoNode>): Array<MathNode> {
  const out: Array<MathNode> = [];
  for (const child of children) {
    const node = parseMathNode(child);
    if (node) out.push(node);
  }
  return out;
}

function parseMathNode(node: PoNode): MathNode | null {
  switch (poTag(node)) {
    case 'm:r':
      return parseMathRun(node);
    case 'm:f':
      return parseFraction(node);
    case 'm:sSup':
      return parseScript(node, false, true);
    case 'm:sSub':
      return parseScript(node, true, false);
    case 'm:sSubSup':
      return parseScript(node, true, true);
    case 'm:sPre':
      return {
        type: 'script',
        base: childRow(node, 'm:e'),
        sub: childRow(node, 'm:sub'),
        sup: childRow(node, 'm:sup'),
        pre: true,
      };
    case 'm:rad':
      return parseRadical(node);
    case 'm:nary':
      return parseNary(node);
    case 'm:func':
      return { type: 'func', name: childRow(node, 'm:fName'), body: childRow(node, 'm:e') };
    case 'm:limLow':
      return {
        type: 'limit',
        base: childRow(node, 'm:e'),
        lim: childRow(node, 'm:lim'),
        pos: 'low',
      };
    case 'm:limUpp':
      return {
        type: 'limit',
        base: childRow(node, 'm:e'),
        lim: childRow(node, 'm:lim'),
        pos: 'upp',
      };
    case 'm:d':
      return parseDelimiter(node);
    case 'm:m':
      return parseMatrix(node);
    case 'm:eqArr':
      return parseEqArray(node);
    case 'm:acc':
      return parseAccent(node);
    case 'm:bar':
      return parseBar(node);
    case 'm:groupChr':
      return parseGroupChr(node);
    // Property elements carry no content.
    case 'm:rPr':
    case 'm:fPr':
    case 'm:ctrlPr':
      return null;
    // Generic content wrappers / unknown elements: surface their content.
    default: {
      const kids = parseMathSeq(poChildren(node));
      return kids.length > 0 ? { type: 'row', children: kids } : null;
    }
  }
}

// Collect a named child element's content as a row (m:num, m:den, m:e, …).
export function childRow(parent: PoNode, tag: string): MathNode {
  const el = poChildren(parent).find((c) => poIs(c, tag));
  return { type: 'row', children: el ? parseMathSeq(poChildren(el)) : [] };
}

function parseMathRun(r: PoNode): MathRun {
  let text = '';
  for (const c of poChildren(r)) {
    if (poIs(c, 'm:t')) text += poText(c);
  }
  const rPr = poChildren(r).find((c) => poIs(c, 'm:rPr'));
  let italic = false;
  let bold = false;
  let nor = false;
  if (rPr) {
    // m:sty: p (plain/upright) | b (bold) | i (italic) | bi (bold-italic).
    const sty = poVal(poChildren(rPr).find((c) => poIs(c, 'm:sty')));
    if (sty === 'i') italic = true;
    else if (sty === 'b') bold = true;
    else if (sty === 'bi') {
      italic = true;
      bold = true;
    } else if (sty === 'p') nor = true;
    if (poChildren(rPr).some((c) => poIs(c, 'm:nor'))) nor = true;
  }
  return {
    type: 'run',
    text,
    ...(italic ? { italic: true } : {}),
    ...(bold ? { bold: true } : {}),
    ...(nor ? { nor: true } : {}),
  };
}

function parseFraction(f: PoNode): MathFraction {
  const num = childRow(f, 'm:num');
  const den = childRow(f, 'm:den');
  const fPr = poChildren(f).find((c) => poIs(c, 'm:fPr'));
  const fType = fPr ? poVal(poChildren(fPr).find((c) => poIs(c, 'm:type'))) : undefined;
  return { type: 'fraction', num, den, ...(fType === 'noBar' ? { barless: true } : {}) };
}

// m:sSup / m:sSub / m:sSubSup — base (m:e) with sub/sup scripts.
function parseScript(node: PoNode, hasSub: boolean, hasSup: boolean): MathNode {
  return {
    type: 'script',
    base: childRow(node, 'm:e'),
    ...(hasSub ? { sub: childRow(node, 'm:sub') } : {}),
    ...(hasSup ? { sup: childRow(node, 'm:sup') } : {}),
  };
}

// m:nary — n-ary operator. m:naryPr/m:chr is the operator (default ∫);
// m:limLoc places the limits; m:subHide/m:supHide drop a limit.
function parseNary(node: PoNode): MathNode {
  const naryPr = poChildren(node).find((c) => poIs(c, 'm:naryPr'));
  let op = '∫';
  let limLoc: 'undOvr' | 'subSup' | undefined;
  const hidden = (tag: string): boolean => {
    if (!naryPr) return false;
    const el = poChildren(naryPr).find((c) => poIs(c, tag));
    return el !== undefined && poToggle(el) !== false;
  };
  if (naryPr) {
    const chr = poVal(poChildren(naryPr).find((c) => poIs(c, 'm:chr')));
    if (chr) op = chr;
    const ll = poVal(poChildren(naryPr).find((c) => poIs(c, 'm:limLoc')));
    if (ll === 'undOvr' || ll === 'subSup') limLoc = ll;
  }
  return {
    type: 'nary',
    op,
    body: childRow(node, 'm:e'),
    ...(hidden('m:subHide') ? {} : { sub: childRow(node, 'm:sub') }),
    ...(hidden('m:supHide') ? {} : { sup: childRow(node, 'm:sup') }),
    ...(limLoc ? { limLoc } : {}),
  };
}

// m:d — delimiters (m:dPr/m:begChr,m:endChr,m:sepChr) around one or more m:e.
function parseDelimiter(node: PoNode): MathNode {
  const dPr = poChildren(node).find((c) => poIs(c, 'm:dPr'));
  const chrVal = (tag: string): string | undefined =>
    dPr ? poVal(poChildren(dPr).find((c) => poIs(c, tag))) : undefined;
  const beg = chrVal('m:begChr');
  const end = chrVal('m:endChr');
  const sep = chrVal('m:sepChr');
  const children = poChildrenWith(node, 'm:e').map(
    (e): MathNode => ({ type: 'row', children: parseMathSeq(poChildren(e)) }),
  );
  return {
    type: 'delimiter',
    begChr: beg ?? '(',
    endChr: end ?? ')',
    children,
    ...(sep !== undefined ? { sepChr: sep } : {}),
  };
}

// m:m — a matrix of m:mr rows, each holding m:e cells.
function parseMatrix(node: PoNode): MathNode {
  const rows = poChildrenWith(node, 'm:mr').map((mr) =>
    poChildrenWith(mr, 'm:e').map(
      (e): MathNode => ({ type: 'row', children: parseMathSeq(poChildren(e)) }),
    ),
  );
  return { type: 'matrix', rows };
}

// m:eqArr — an equation array: each m:e child is one stacked equation row.
function parseEqArray(node: PoNode): MathNode {
  const rows = poChildrenWith(node, 'm:e').map(
    (e): MathNode => ({ type: 'row', children: parseMathSeq(poChildren(e)) }),
  );
  return { type: 'eqArr', rows };
}

// m:acc — accent (m:accPr/m:chr, default combining circumflex) over m:e.
function parseAccent(node: PoNode): MathNode {
  const accPr = poChildren(node).find((c) => poIs(c, 'm:accPr'));
  const chr = accPr ? poVal(poChildren(accPr).find((c) => poIs(c, 'm:chr'))) : undefined;
  return { type: 'accent', char: chr ?? '̂', base: childRow(node, 'm:e') };
}

// m:bar — a bar above (default) or below the base.
function parseBar(node: PoNode): MathNode {
  const barPr = poChildren(node).find((c) => poIs(c, 'm:barPr'));
  const pos = barPr ? poVal(poChildren(barPr).find((c) => poIs(c, 'm:pos'))) : undefined;
  return { type: 'bar', base: childRow(node, 'm:e'), pos: pos === 'bot' ? 'bot' : 'top' };
}

// m:groupChr — a grouping character (default under-brace) above or below.
function parseGroupChr(node: PoNode): MathNode {
  const gPr = poChildren(node).find((c) => poIs(c, 'm:groupChrPr'));
  const chr = gPr ? poVal(poChildren(gPr).find((c) => poIs(c, 'm:chr'))) : undefined;
  const pos = gPr ? poVal(poChildren(gPr).find((c) => poIs(c, 'm:pos'))) : undefined;
  return {
    type: 'groupChr',
    char: chr ?? '⏟', // ⏟ bottom brace
    base: childRow(node, 'm:e'),
    pos: pos === 'top' ? 'top' : 'bot',
  };
}

// m:rad — radical. The degree (m:deg) is shown only when present and non-empty.
function parseRadical(node: PoNode): MathNode {
  const degEl = poChildren(node).find((c) => poIs(c, 'm:deg'));
  const hasDeg = degEl !== undefined && parseMathSeq(poChildren(degEl)).length > 0;
  return {
    type: 'radical',
    radicand: childRow(node, 'm:e'),
    ...(hasDeg ? { degree: childRow(node, 'm:deg') } : {}),
  };
}
