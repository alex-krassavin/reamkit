import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

function layoutOf(docx: Uint8Array) {
  const flow = Ream.parse(docx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
}

/** The text of every line drawn, page by page. */
function textByPage(laid: ReturnType<typeof layoutOf>): Array<Array<string>> {
  return laid.pages.map((p) =>
    p.commands
      .filter((c) => c.type === 'line')
      .map((c) =>
        (c as unknown as { line: { tokens: ReadonlyArray<{ text?: string }> } }).line.tokens
          .map((t) => t.text ?? '')
          .join(''),
      )
      .filter((t) => t !== ''),
  );
}

// Half a page of body text, then a one-cell row that needs most of a page.
const FILLER = Array.from(
  { length: 24 },
  (_, i) => `<w:p><w:r><w:t>Filler${String(i)}</w:t></w:r></w:p>`,
).join('');
const CELL = Array.from(
  { length: 30 },
  (_, i) => `<w:p><w:r><w:t>Row${String(i)}</w:t></w:r></w:p>`,
).join('');

const doc = (trPr = '') =>
  buildDocxFromBody(
    `${FILLER}<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>` +
      `<w:tr>${trPr}<w:tc>${CELL}</w:tc></w:tr></w:tbl>`,
  );

describe('a table row at the page edge (§17.4.6)', () => {
  it('breaks the row where the page ends', () => {
    // table-row-data-displayed-twice.docx is one 713pt row: drawn whole it ran
    // off the bottom of the page and over the section that follows it.
    const pages = textByPage(layoutOf(doc()));
    expect(pages.length).toBe(2);
    expect(pages[0]!).toContain('Row0'); // the row starts beside the filler…
    expect(pages[1]!).toContain('Row29'); // …and ends on the next page
    // Nothing is lost or repeated: every line of the cell is drawn once.
    const all = pages.flat().filter((t) => t.startsWith('Row'));
    expect(all.length).toBe(30);
    expect(new Set(all).size).toBe(30);
  });

  it('moves a `w:cantSplit` row whole instead', () => {
    const pages = textByPage(layoutOf(doc('<w:trPr><w:cantSplit/></w:trPr>')));
    expect(pages[0]!.some((t) => t.startsWith('Row'))).toBe(false);
    expect(pages[1]!).toContain('Row0');
  });

  it('counts the space between a cell’s paragraphs when it splits one', () => {
    // The chunk sizer measured lines only, so a cell of spaced paragraphs
    // reported less height than it drew — and the piece it cut off overflowed
    // the page it was cut to fit.
    const spaced = Array.from(
      { length: 30 },
      (_, i) =>
        `<w:p><w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr>` +
        `<w:r><w:t>Row${String(i)}</w:t></w:r></w:p>`,
    ).join('');
    const laid = layoutOf(
      buildDocxFromBody(
        `${FILLER}<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>` +
          `<w:tr><w:tc>${spaced}</w:tc></w:tr></w:tbl>`,
      ),
    );
    // Letter minus 1" margins: nothing may be drawn below the bottom margin.
    for (const page of laid.pages) {
      for (const c of page.commands) {
        if (c.type !== 'line') continue;
        expect((c as unknown as { baselineY: number }).baselineY).toBeLessThanOrEqual(721);
      }
    }
  });
});
