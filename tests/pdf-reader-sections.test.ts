// A page set in columns under a line across it opens a section of its own, and
// that section starts the page itself (§17.18.77 nextPage): Word does not break
// before a paragraph that carries the section before it, and comments.pdf's
// twelfth page came back under its eleventh on one sheet. And a section's head
// and foot are its sheet's (§17.6.11): measured on the one page it opens, its
// top margin stood wherever that page's first line happened to.

import { describe, expect, it } from 'vitest';

import type { BodyElement } from '@/core/document-model';
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
    // (The right column opens on a break to it: the foot of the sheet is the
    // second page's, sixty points under the first page's columns.)
    const words = carrier.paragraph.runs.filter((r) => r.columnBreak !== true);
    expect(words.map((r) => r.text).join('')).toMatch(/^Ra01/u);
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

describe('a last page that sets its columns balanced', () => {
  /** A full page of two columns, and a last page of `last`. */
  const read = (last: ReadonlyArray<string>) =>
    reconstructByLayout(
      PdfFile.parse(
        pages([
          ['BT /F1 9 Tf', ...columns(740, 50, 'a'), 'ET'],
          ['BT /F1 9 Tf', ...last, 'ET'],
        ]),
      ),
    ).doc;

  it('ends their section before the document does, for a word processor to balance them', () => {
    // comments.pdf's references stand nine to a column on its last page, and
    // came back all nineteen down the left one, the right one empty.
    const [columned, closing] = read(columns(740, 20, 'z')).sections.slice(-2);
    expect(columned?.properties.columns?.count).toBe(2);
    expect(closing?.properties.columns).toBeUndefined();
    expect(closing?.properties.sectionStart).toBe('continuous');
  });

  it('leaves a last page whose first column runs to the foot as it was', () => {
    const ops: Array<string> = [];
    for (let k = 0; k < 50; k++) {
      ops.push(at(54, 740 - k * 12, `Lz${String(k)} a line of prose that runs across its column`));
    }
    for (let k = 0; k < 10; k++) {
      ops.push(at(317, 740 - k * 12, `Rz${String(k)} a line of prose that runs across its column`));
    }
    expect(read(ops).sections.at(-1)?.properties.columns?.count).toBe(2);
  });
});

describe('a column the page ends short of the foot of its text', () => {
  /** The one break to a column in `flow`, what follows it, and the page's first paragraph. */
  const turn = (flow: { body: ReadonlyArray<BodyElement> }) => {
    const index = flow.body.findIndex(
      (el) => el.kind === 'paragraph' && el.paragraph.runs.some((r) => r.columnBreak === true),
    );
    const breaks = flow.body.filter(
      (el) => el.kind === 'paragraph' && el.paragraph.runs.some((r) => r.columnBreak === true),
    );
    const carrier = flow.body[index];
    const opens = flow.body
      .slice(index + 1)
      .find((el) => el.kind === 'paragraph' && el.paragraph.runs.length > 0);
    const first = flow.body.find((el) => el.kind === 'paragraph');
    if (carrier?.kind !== 'paragraph' || opens?.kind !== 'paragraph' || first?.kind !== 'paragraph')
      throw new Error('paragraphs');
    return {
      breaks: breaks.length,
      carrier: carrier.paragraph,
      opens: opens.paragraph,
      first: first.paragraph,
    };
  };

  it('breaks to the next column there, where a word processor would run it on', () => {
    // canvas.pdf ends its first sheet's left column fifty points above the
    // foot its second sheet sets, and the head of the right column came back
    // under the left one.
    const flow = reconstructByLayout(
      PdfFile.parse(
        pages([
          ['BT /F1 9 Tf', ...columns(740, 30, 'a'), 'ET'],
          ['BT /F1 9 Tf', ...columns(740, 50, 'b'), 'ET'],
        ]),
      ),
    ).doc;
    const { breaks, carrier, opens, first } = turn(flow);
    // Before the first line of the right column of the first sheet, and nowhere
    // else: the second sheet's columns run to the foot.
    expect(breaks).toBe(1);
    // …as a paragraph of its own: inside one, a word processor leaves what is
    // before the break in the column it turns from.
    expect(carrier.runs.map((r) => r.text).join('')).toBe('\n');
    expect(opens.runs.map((r) => r.text).join('')).toMatch(/^Ra01/u);
    // The right column's first line stands off the head of the column as the
    // left one's stands off the head of the page: both are on one baseline.
    expect(opens.properties.spacingBefore).toBeCloseTo(first.properties.spacingBefore ?? 0, 1);
  });

  it("breaks before the next column's first line past a drawing anchored over it", () => {
    // canvas.pdf's right column opens on a drawing anchored to the page, which
    // takes no room in the column: the right column's first line is past it.
    const flow = reconstructByLayout(
      PdfFile.parse(
        pages([
          ['0.5 g 317 752 40 6 re f', 'BT /F1 9 Tf 0 g', ...columns(740, 30, 'a'), 'ET'],
          ['BT /F1 9 Tf', ...columns(740, 50, 'b'), 'ET'],
        ]),
      ),
    ).doc;
    const { breaks, opens, first } = turn(flow);
    expect(breaks).toBe(1);
    expect(opens.runs.map((r) => r.text).join('')).toMatch(/^Ra01/u);
    expect(opens.properties.spacingBefore).toBeCloseTo(first.properties.spacingBefore ?? 0, 1);
  });
});
