// A page set in columns under a line across it opens a section of its own, and
// that section starts the page itself (§17.18.77 nextPage): Word does not break
// before a paragraph that carries the section before it, and comments.pdf's
// twelfth page came back under its eleventh on one sheet. And a section's head
// and foot are its sheet's (§17.6.11): measured on the one page it opens, its
// top margin stood wherever that page's first line happened to.

import { describe, expect, it } from 'vitest';

import type { PdfValue } from '@/pdf/objects';
import { OpcPackage } from '@/core/opc';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';
import { dict, name, stream } from '@/pdf/objects';
import { PdfDocument } from '@/pdf/writer';
import { writeDocx } from '@/word/docx-writer';

/** Letter pages of `ops` in Helvetica, one page each. */
function pages(sheets: ReadonlyArray<ReadonlyArray<string>>): Uint8Array {
  const doc = new PdfDocument();
  const font = doc.add(
    dict({ Type: name('Font'), Subtype: name('Type1'), BaseFont: name('Helvetica') }),
  );
  const pagesMap = dict({ Type: name('Pages'), Kids: [], Count: sheets.length });
  const pagesRef = doc.add(pagesMap);
  for (const ops of sheets) {
    const content = doc.add(stream({}, new TextEncoder().encode(ops.join('\n'))));
    const leaf = doc.add(
      dict({
        Type: name('Page'),
        Parent: pagesRef,
        MediaBox: [0, 0, 612, 792],
        Resources: dict({ Font: dict({ F1: font }) }),
        Contents: content,
      }),
    );
    (pagesMap.get('Kids') as Array<PdfValue>).push(leaf);
  }
  const catalog = doc.add(dict({ Type: name('Catalog'), Pages: pagesRef }));
  return doc.build(catalog);
}

const at = (x: number, y: number, text: string): string =>
  `1 0 0 1 ${String(x)} ${String(y)} Tm (${text}) Tj`;

/** Two columns of prose from `top`, `rows` lines deep, each line tagged `tag`. */
function columns(top: number, rows: number, tag: string): Array<string> {
  const ops: Array<string> = [];
  for (let k = 0; k < rows; k++) {
    const n = String(k + 1).padStart(2, '0');
    const y = top - k * 12;
    ops.push(at(54, y, `L${tag}${n} a line of prose that runs across its column`));
    ops.push(at(317, y, `R${tag}${n} a line of prose that runs across its column`));
  }
  return ops;
}

// Three sheets of a paper in two columns. The second opens on a line across
// the page, set low: what stands over it on the page is a drawing, which says
// nothing to a margin.
const doc = reconstructByLayout(
  PdfFile.parse(
    pages([
      ['BT /F1 9 Tf', ...columns(740, 50, 'a'), 'ET'],
      [
        'BT /F1 9 Tf',
        at(54, 400, 'A line set across the head of the columns under it, wider than one'),
        ...columns(380, 25, 'b'),
        'ET',
      ],
      ['BT /F1 9 Tf', ...columns(740, 50, 'c'), 'ET'],
    ]),
  ),
).doc;

describe('a page that opens a section of its own', () => {
  const across = doc.body.findIndex(
    (el) =>
      el.kind === 'paragraph' && el.paragraph.runs.some((r) => r.text.startsWith('A line set')),
  );
  const opened = doc.sections.find((s) => s.endIndex > across)!;
  const before = doc.sections[doc.sections.indexOf(opened) - 1]!;

  it('starts the page itself, and the section before it ends on its own page', () => {
    expect(across).toBeGreaterThan(0);
    expect(opened.properties.sectionStart).toBeUndefined();
    // …on that page's own last paragraph: a carrier of its own, on a page the
    // words fill to the foot, goes over onto a sheet of its own.
    const carrier = doc.body[before.endIndex - 1];
    if (carrier?.kind !== 'paragraph') throw new Error('a paragraph');
    expect(carrier.paragraph.properties.pageBreakBefore).not.toBe(true);
    expect(carrier.paragraph.runs.map((r) => r.text).join('')).toMatch(/^Ra01/u);
  });

  it('never breaks on a paragraph that carries a section: Word does not', () => {
    const xml = new TextDecoder().decode(
      OpcPackage.open(writeDocx(doc).bytes).getMainDocument().data,
    );
    const carriers = xml.match(/<w:pPr>(?:(?!<\/w:pPr>).)*<w:sectPr>/gsu) ?? [];
    expect(carriers.length).toBeGreaterThan(0);
    expect(carriers.filter((pPr) => pPr.includes('<w:pageBreakBefore/>'))).toEqual([]);
  });
});

describe('the head and foot of a section', () => {
  it('are its sheet’s, whatever page it opens', () => {
    const tops = new Set(doc.sections.map((s) => s.properties.margins?.top));
    const bottoms = new Set(doc.sections.map((s) => s.properties.margins?.bottom));
    expect(tops.size).toBe(1);
    expect(bottoms.size).toBe(1);
    // …the sheet's head, where its first page's text begins, not 400 points down.
    expect([...tops][0]).toBeLessThan(792 - 700);
  });
});
