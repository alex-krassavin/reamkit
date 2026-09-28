// §14.7–§14.8 — the structure tree as the tagged reading needs it: a
// document's own types read through its /RoleMap, inline elements read as the
// stretch of their parent's line they are, and an element's own text kept
// where it stands among its children.

import { describe, expect, it } from 'vitest';

import type { StructNode } from '@/pdf-reader/struct-tree';
import { PdfFile } from '@/pdf-reader/document';
import { readStructTree } from '@/pdf-reader/struct-tree';

/**
 * One page whose content marks six stretches, MCIDs 0 to 5, under the tree the
 * objects from 6 on describe. Object 3 is the page, 5 the tree root.
 */
function tagged(tree: ReadonlyArray<string>, roleMap = ''): Uint8Array {
  const content = [0, 1, 2, 3, 4, 5]
    .map(
      (id) =>
        `/P <</MCID ${String(id)}>> BDC BT /F1 12 Tf ${String(20 + id * 20)} 100 Td (w${String(id)}) Tj ET EMC`,
    )
    .join('\n');
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 5 0 R /MarkInfo << /Marked true >> >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R ' +
      '/Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    `<< /Type /StructTreeRoot /K 6 0 R ${roleMap} >>`,
    ...tree,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

/** A node as `type[mcids]{children}`, for comparing shapes at a glance. */
const shape = (n: StructNode): string =>
  `${n.type}[${n.mcids.map((m) => m.mcid).join(',')}]` +
  (n.children.length > 0 ? `{${n.children.map(shape).join(' ')}}` : '');

describe('a structure tree read the way its document means it (§14.8)', () => {
  // bug1937438_mml_from_latex.pdf: a LaTeX document tags its heading as a
  // `section` carrying its number as a `section-number`, and a sentence as a
  // `text` with a `Formula` of MathML set between two of its stretches.
  const latex = tagged(
    [
      '<< /Type /StructElem /S /Document /K [7 0 R] >>',
      '<< /Type /StructElem /S /sect /Pg 3 0 R /K [8 0 R 10 0 R 5] >>',
      '<< /Type /StructElem /S /section /Pg 3 0 R /K [9 0 R 1] >>',
      '<< /Type /StructElem /S /section-number /Pg 3 0 R /K 0 >>',
      '<< /Type /StructElem /S /text /Pg 3 0 R /K [2 11 0 R 4] >>',
      '<< /Type /StructElem /S /Formula /Pg 3 0 R /K [12 0 R] >>',
      '<< /Type /StructElem /S /mi /NS 13 0 R /Pg 3 0 R /K 3 >>',
      '<< /Type /Namespace /NS (http://www.w3.org/1998/Math/MathML) >>',
    ],
    '/RoleMap << /sect /Sect /section /H1 /section-number /Span /text /P >>',
  );

  it('reads a document’s own types through its /RoleMap', () => {
    const root = readStructTree(PdfFile.parse(latex));
    expect(root?.children[0]?.type).toBe('Sect');
    expect(root?.children[0]?.children.map((c) => c.type)).toContain('H1');
  });

  it('reads an inline element, MathML among them, as a stretch of its parent’s line', () => {
    // Read as blocks of their own, the heading came back "1" without its
    // words, and each glyph of the formula as a paragraph apiece.
    const root = readStructTree(PdfFile.parse(latex));
    const sect = root?.children[0];
    expect(sect ? shape(sect) : '').toBe('Sect[]{H1[0,1] P[2,3,4] P[5]}');
  });

  it('keeps an element’s own text where it stands among its children', () => {
    // MCID 5 hangs on the section itself, after its paragraph: it is a
    // paragraph of its own, after that one — not text dropped because the
    // element also had children.
    const root = readStructTree(PdfFile.parse(latex));
    const last = root?.children[0]?.children.at(-1);
    expect(last?.mcids.map((m) => m.mcid)).toEqual([5]);
  });

  it('leaves a figure’s own marked content its own', () => {
    // A figure's marked content is its picture, whatever else it holds.
    const figure = tagged([
      '<< /Type /StructElem /S /Document /K [7 0 R] >>',
      '<< /Type /StructElem /S /Figure /Pg 3 0 R /K [0 8 0 R] >>',
      '<< /Type /StructElem /S /Caption /Pg 3 0 R /K 1 >>',
    ]);
    const root = readStructTree(PdfFile.parse(figure));
    expect(root?.children[0] ? shape(root.children[0]) : '').toBe('Figure[0]{Caption[1]}');
  });
});
