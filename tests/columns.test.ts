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
