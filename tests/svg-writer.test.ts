// The SVG target: text drawn from the faces' own outlines, with the words laid
// unseen over them, and a workbook drawn as images of its sheets — each whole on
// a page of its own, cut to what is on it, with the window's gridlines.

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';

const fontsDir = resolve(__dirname, 'fixtures/fonts');
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(fontsDir, 'Roboto-Regular.ttf'))),
};

const svgOf = async (bytes: Uint8Array): Promise<string> =>
  new TextDecoder().decode(await Ream.parse(bytes).convert('svg', { fonts: FONTS }));

describe('svg text (glyph outlines)', () => {
  it('draws each glyph once from its outline and places it, the words laid over unseen', async () => {
    const svg = await svgOf(buildDocxFromBody('<w:p><w:r><w:t>Hello hello</w:t></w:r></w:p>'));
    // One definition per glyph ("l" is used four times and defined once)…
    const defs = [...svg.matchAll(/<path id="(g\d+-\d+)" d="/gu)].map((m) => m[1]);
    expect(new Set(defs).size).toBe(defs.length);
    expect(defs.length).toBeGreaterThan(3);
    // …placed by reference, in the text's colour…
    expect(svg).toMatch(/<g transform="matrix\([^)]*\)" fill="#000000"><use href="#g\d+-\d+"/u);
    // …with no sans standing in for the face, and the words still searchable.
    expect(svg).not.toContain('font-family="sans-serif" font-size="11"');
    expect(svg).toContain('fill-opacity="0" xml:space="preserve">Hello hello</text>');
  });

  it('strokes a faked bold and rules an underline, as the PDF draws them', async () => {
    // One regular face: the bold is made by stroking the glyphs a thirtieth of
    // the em wide.
    const svg = await svgOf(
      buildDocxFromBody(
        '<w:p><w:r><w:rPr><w:b/><w:u w:val="single"/><w:color w:val="C00000"/></w:rPr><w:t>Bold</w:t></w:r></w:p>',
      ),
    );
    expect(svg).toMatch(/fill="#C00000" stroke="#C00000" stroke-width="0\.03"/u);
    expect(svg).toMatch(
      /<rect x="[\d.]+" y="[\d.]+" width="[\d.]+" height="[\d.]+" fill="#C00000"\/>/u,
    );
  });
});

describe('a workbook as images of its sheets', () => {
  it('puts each visible sheet whole on a page as large as its content', async () => {
    const wide = Array.from({ length: 20 }, (_, i) => `column ${String(i + 1)}`);
    const xlsx = buildXlsx({
      sheets: [
        { name: 'Wide', rows: [wide], columns: [{ min: 1, max: 20, widthChars: 20 }] },
        { name: 'Small', rows: [['a', 'b']] },
      ],
    });
    const svg = await svgOf(xlsx);
    const pages = [
      ...svg.matchAll(/data-page="\d+">\n<rect x="0" y="0" width="([\d.]+)" height="([\d.]+)"/gu),
    ].map((m) => [Number(m[1]), Number(m[2])]);
    // Two sheets, two pages — the wide one in one piece, wider than any paper,
    // and neither as tall as a page of it.
    expect(pages).toHaveLength(2);
    expect(pages[0]![0]).toBeGreaterThan(20 * 100);
    expect(pages[0]![1]).toBeLessThan(40);
    expect(pages[1]![0]).toBeLessThan(120);
    expect(svg).toContain('column 20');
  });

  it("rules the window's gridlines, and none where the sheet hides them", async () => {
    const shown = await svgOf(buildXlsx({ rows: [['x', 'y']] }));
    const hidden = await svgOf(
      buildXlsx({ sheets: [{ name: 'S', rows: [['x', 'y']], hideGridLines: true }] }),
    );
    expect(shown).toContain('stroke="#D4D4D4"');
    expect(hidden).not.toContain('stroke="#D4D4D4"');
  });
});
