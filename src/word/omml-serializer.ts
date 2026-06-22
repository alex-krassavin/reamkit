// ECMA-376 Part 1 §22 — OfficeMathML (OMML) serializer. The inverse of
// omml-parser.ts: a MathNode tree → the inner content of an <m:oMath> element
// (the m: namespace is declared by the caller). Covers every node the parser
// produces, so a parse → serialize → parse round-trip preserves the math (WT3).

import type { MathNode } from '@/core/document-model';

/**
 * Serialize a {@link MathNode} tree to the inner content of an `<m:oMath>`
 * element (ECMA-376 Part 1 §22). The inverse of `parseOMath`; the `m:` namespace
 * is declared by the caller, and the output round-trips back through the parser
 * (WT3).
 *
 * @param node The math tree (typically a `row` node).
 * @returns The serialized OMML markup.
 */
export function omathXml(node: MathNode): string {
  return ser(node);
}

function ser(node: MathNode): string {
  switch (node.type) {
    case 'row':
      return node.children.map(ser).join('');
    case 'run':
      return `<m:r>${runPr(node)}<m:t xml:space="preserve">${esc(node.text)}</m:t></m:r>`;
    case 'fraction':
      return (
        '<m:f>' +
        (node.barless ? '<m:fPr><m:type m:val="noBar"/></m:fPr>' : '') +
        `<m:num>${ser(node.num)}</m:num><m:den>${ser(node.den)}</m:den></m:f>`
      );
    case 'script':
      return scriptXml(node);
    case 'radical':
      return (
        '<m:rad>' +
        (node.degree ? '' : '<m:radPr><m:degHide m:val="1"/></m:radPr>') +
        (node.degree ? `<m:deg>${ser(node.degree)}</m:deg>` : '<m:deg/>') +
        `<m:e>${ser(node.radicand)}</m:e></m:rad>`
      );
    case 'nary':
      return (
        '<m:nary><m:naryPr>' +
        `<m:chr m:val="${attr(node.op)}"/>` +
        (node.limLoc ? `<m:limLoc m:val="${node.limLoc}"/>` : '') +
        (node.sub ? '' : '<m:subHide m:val="1"/>') +
        (node.sup ? '' : '<m:supHide m:val="1"/>') +
        '</m:naryPr>' +
        `<m:sub>${node.sub ? ser(node.sub) : ''}</m:sub>` +
        `<m:sup>${node.sup ? ser(node.sup) : ''}</m:sup>` +
        `<m:e>${ser(node.body)}</m:e></m:nary>`
      );
    case 'func':
      return `<m:func><m:fName>${ser(node.name)}</m:fName><m:e>${ser(node.body)}</m:e></m:func>`;
    case 'limit':
      return (
        `<m:lim${node.pos === 'upp' ? 'Upp' : 'Low'}>` +
        `<m:e>${ser(node.base)}</m:e><m:lim>${ser(node.lim)}</m:lim>` +
        `</m:lim${node.pos === 'upp' ? 'Upp' : 'Low'}>`
      );
    case 'delimiter':
      return (
        '<m:d><m:dPr>' +
        `<m:begChr m:val="${attr(node.begChr)}"/>` +
        (node.sepChr !== undefined ? `<m:sepChr m:val="${attr(node.sepChr)}"/>` : '') +
        `<m:endChr m:val="${attr(node.endChr)}"/></m:dPr>` +
        node.children.map((c) => `<m:e>${ser(c)}</m:e>`).join('') +
        '</m:d>'
      );
    case 'matrix':
      return (
        '<m:m>' +
        node.rows
          .map((row) => `<m:mr>${row.map((c) => `<m:e>${ser(c)}</m:e>`).join('')}</m:mr>`)
          .join('') +
        '</m:m>'
      );
    case 'accent':
      return `<m:acc><m:accPr><m:chr m:val="${attr(node.char)}"/></m:accPr><m:e>${ser(node.base)}</m:e></m:acc>`;
    case 'bar':
      return `<m:bar><m:barPr><m:pos m:val="${node.pos}"/></m:barPr><m:e>${ser(node.base)}</m:e></m:bar>`;
    case 'groupChr':
      return (
        '<m:groupChr><m:groupChrPr>' +
        `<m:chr m:val="${attr(node.char)}"/><m:pos m:val="${node.pos}"/></m:groupChrPr>` +
        `<m:e>${ser(node.base)}</m:e></m:groupChr>`
      );
    case 'eqArr':
      return `<m:eqArr>${node.rows.map((r) => `<m:e>${ser(r)}</m:e>`).join('')}</m:eqArr>`;
  }
}

// m:sSup / m:sSub / m:sSubSup / m:sPre — the run-script element matching the
// present sub/sup (m:sPre always carries both, empty when absent).
function scriptXml(node: Extract<MathNode, { type: 'script' }>): string {
  const base = `<m:e>${ser(node.base)}</m:e>`;
  const sub = `<m:sub>${node.sub ? ser(node.sub) : ''}</m:sub>`;
  const sup = `<m:sup>${node.sup ? ser(node.sup) : ''}</m:sup>`;
  if (node.pre) return `<m:sPre>${base}${sub}${sup}</m:sPre>`;
  if (node.sub && node.sup) return `<m:sSubSup>${base}${sub}${sup}</m:sSubSup>`;
  if (node.sub) return `<m:sSub>${base}${sub}</m:sSub>`;
  if (node.sup) return `<m:sSup>${base}${sup}</m:sSup>`;
  return ser(node.base); // a script with neither degrades to the base
}

// m:rPr/m:sty — p (upright/nor) | b | i | bi; absent lets letters auto-italicise.
function runPr(run: Extract<MathNode, { type: 'run' }>): string {
  let sty: string | undefined;
  if (run.bold && run.italic) sty = 'bi';
  else if (run.bold) sty = 'b';
  else if (run.italic) sty = 'i';
  else if (run.nor) sty = 'p';
  return sty ? `<m:rPr><m:sty m:val="${sty}"/></m:rPr>` : '';
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function attr(s: string): string {
  return esc(s).replace(/"/g, '&quot;');
}
