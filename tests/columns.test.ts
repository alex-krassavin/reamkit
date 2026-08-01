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

function lineXs(commands: ReadonlyArray<{ type: string }>): Array<number> {
  return commands
    .filter((c) => c.type === 'line')
    .map((c) => (c as unknown as { originX: number }).originX);
}

function layoutOf(docx: Uint8Array) {
  const flow = Ream.parse(docx).flow;
  return layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
}

describe('multi-column sections (§17.6.4)', () => {
  // A5 portrait-ish page, 2 columns, enough lines to overflow column 1.
  const docx = (cols: string, lines = 30) =>
    buildDocxFromBody(
      Array.from({ length: lines }, (_, i) => `<w:p><w:r><w:t>line ${i}</w:t></w:r></w:p>`).join(
        '',
      ) +
        `<w:sectPr><w:pgSz w:w="8400" w:h="11900"/>` +
        `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>${cols}</w:sectPr>`,
    );

  it('fills column one, then column two, then the next page', () => {
    const laid = layoutOf(docx('<w:cols w:num="2" w:space="720"/>', 90));
    const xs = lineXs(laid.pages[0]!.commands);
    // content width = 420−72 = 348pt; column = (348−36)/2 = 156; col2 x = 36+156+36 = 228.
    const col1 = xs.filter((x) => Math.abs(x - 36) < 1);
    const col2 = xs.filter((x) => Math.abs(x - 228) < 1);
    expect(col1.length).toBeGreaterThan(0);
    expect(col2.length).toBeGreaterThan(0);
    // Reading order on the page: all column-1 lines come before column-2 ones.
    const firstCol2Idx = xs.findIndex((x) => Math.abs(x - 228) < 1);
    expect(xs.slice(firstCol2Idx).every((x) => Math.abs(x - 228) < 1)).toBe(true);
    // 90 lines at ~37 per column overflow both columns → a second page exists
    // and starts back at column one.
    expect(laid.pages.length).toBeGreaterThan(1);
    const xs2 = lineXs(laid.pages[1]!.commands);
    expect(Math.abs(xs2[0]! - 36)).toBeLessThan(1);
  });

  it('explicit unequal w:col widths position the second column correctly', () => {
    const laid = layoutOf(
      docx('<w:cols><w:col w:w="3000" w:space="720"/><w:col w:w="3240"/></w:cols>', 60),
    );
    const xs = lineXs(laid.pages[0]!.commands);
    // col1 at 36pt; col2 at 36 + 150 + 36 = 222pt.
    expect(xs.some((x) => Math.abs(x - 222) < 1)).toBe(true);
  });

  it('single-column documents are untouched by the column machinery', () => {
    const laid = layoutOf(docx('', 5));
    const xs = lineXs(laid.pages[0]!.commands);
    expect(xs.every((x) => Math.abs(x - 36) < 1)).toBe(true);
  });
});

describe('a column break (§17.3.3.1 w:br w:type="column")', () => {
  const twoCols = (bodyXml: string) =>
    buildDocxFromBody(
      bodyXml +
        `<w:sectPr><w:pgSz w:w="8400" w:h="11900"/>` +
        `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>` +
        `<w:cols w:num="2" w:space="720"/></w:sectPr>`,
    );

  it('sends what follows it to the next column, mid-paragraph', () => {
    // columnbreak.docx breaks inside one paragraph, and read as a soft line
    // break both halves stayed in the first column.
    const laid = layoutOf(
      twoCols(
        '<w:p><w:r><w:t>before</w:t></w:r>' +
          '<w:r><w:br w:type="column"/></w:r>' +
          '<w:r><w:t>after</w:t></w:r></w:p>',
      ),
    );
    const xs = lineXs(laid.pages[0]!.commands);
    expect(xs).toHaveLength(2);
    expect(xs[0]).toBeCloseTo(36, 0); // column one
    expect(xs[1]).toBeCloseTo(228, 0); // column two
    expect(laid.pages).toHaveLength(1);
  });

  it('starts a fresh page when the break falls in the last column', () => {
    const laid = layoutOf(
      twoCols(
        '<w:p><w:r><w:t>a</w:t></w:r><w:r><w:br w:type="column"/></w:r>' +
          '<w:r><w:t>b</w:t></w:r><w:r><w:br w:type="column"/></w:r>' +
          '<w:r><w:t>c</w:t></w:r></w:p>',
      ),
    );
    expect(laid.pages).toHaveLength(2);
    expect(lineXs(laid.pages[1]!.commands)[0]).toBeCloseTo(36, 0);
  });
});

describe('a text frame (§17.3.1.11 w:framePr)', () => {
  const framed = (framePr: string, text = 'Name', after = '') =>
    buildDocxFromBody(
      '<w:p><w:r><w:t>first</w:t></w:r></w:p>' +
        `<w:p><w:pPr>${framePr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>` +
        after +
        '<w:p><w:r><w:t>last</w:t></w:r></w:p>' +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1418" w:right="1134" w:bottom="1134" w:left="1366"/></w:sectPr>',
    );

  it('leaves the flow and lands where the frame says', () => {
    // content-control-grab-bag.docx hangs its frame off the page at
    // x=6561/y=2705 twips; read inline it sat between "first" and "last".
    const laid = layoutOf(
      framed(
        '<w:framePr w:w="3969" w:h="2382" w:hRule="exact" w:wrap="around"' +
          ' w:vAnchor="page" w:hAnchor="page" w:x="6561" w:y="2705"/>',
      ),
    );
    const lines = laid.pages[0]!.commands.filter((c) => c.type === 'line');
    const texts = lines.map((c) =>
      (c as unknown as { line: { tokens: ReadonlyArray<{ text?: string }> } }).line.tokens
        .map((t) => t.text ?? '')
        .join(''),
    );
    const xs = lineXs(laid.pages[0]!.commands);
    const idx = texts.indexOf('Name');
    expect(idx).toBeGreaterThanOrEqual(0);
    // 6561 twips = 328.05pt from the page's left edge.
    expect(xs[idx]).toBeCloseTo(328.05, 1);
    // …and the two body paragraphs sit together at the text margin.
    expect(xs[texts.indexOf('first')]).toBeCloseTo(68.3, 1);
    expect(xs[texts.indexOf('last')]).toBeCloseTo(68.3, 1);
  });

  it('is as wide as its text when it states no width', () => {
    // CT-with-frame.docx's marginal note states none: taken as the full column
    // it excluded the whole of it and pushed the paragraph beside it away.
    const laid = layoutOf(
      framed('<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:y="1"/>', '0123'),
    );
    const body = laid.pages[0]!.commands.filter((c) => c.type === 'line')
      .map(
        (c) =>
          c as unknown as { originX: number; line: { tokens: ReadonlyArray<{ text?: string }> } },
      )
      .filter((c) => c.line.tokens.map((t) => t.text ?? '').join('') === 'last');
    expect(body[0]!.originX).toBeCloseTo(68.3, 1);
  });

  it('holds consecutive paragraphs of the same frame together', () => {
    const laid = layoutOf(
      framed(
        '<w:framePr w:w="2000" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="6000" w:y="3000"/>',
        'one',
        '<w:p><w:pPr><w:framePr w:w="2000" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="6000" w:y="3000"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>',
      ),
    );
    const lines = laid.pages[0]!.commands.filter((c) => c.type === 'line').map(
      (c) =>
        c as unknown as {
          originX: number;
          baselineY: number;
          line: { tokens: ReadonlyArray<{ text?: string }> };
        },
    );
    const one = lines.find((l) => l.line.tokens.map((t) => t.text ?? '').join('') === 'one')!;
    const two = lines.find((l) => l.line.tokens.map((t) => t.text ?? '').join('') === 'two')!;
    // Same box, one under the other — not two frames at the same spot.
    expect(two.originX).toBeCloseTo(one.originX, 1);
    expect(two.baselineY).toBeGreaterThan(one.baselineY);
  });
});
