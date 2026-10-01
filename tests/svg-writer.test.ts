// The SVG target draws a workbook as images of its sheets — each whole on a page
// of its own, cut to what is on it, with the window's gridlines.

import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { Ream } from '@/core/converter/ream';

const fontsDir = resolve(__dirname, 'fixtures/fonts');
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(fontsDir, 'Roboto-Regular.ttf'))),
};

const svgOf = async (bytes: Uint8Array): Promise<string> =>
  new TextDecoder().decode(await Ream.parse(bytes).convert('svg', { fonts: FONTS }));

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
