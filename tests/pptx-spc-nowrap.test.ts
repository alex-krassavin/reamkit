// §21.1.2.3.9 `a:rPr@spc` — character tracking, stated in hundredths of a
// point like `sz`. And ST_TextWrappingType §20.1.10.85 — `a:bodyPr@wrap="none"`,
// the box whose lines run as far as their words do however narrow it is.
//
// Both attributes were dropped by the PPTX reader while every consumer of them
// already existed: `letterSpacingPt` measured and drawn by the layout and every
// writer, `noWrap` read as an unbounded inner width. The tests below pin the
// reader's side of both contracts, plus one end-to-end check each that the
// dropped attribute actually changes the page.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildPptx } from './fixtures/build-pptx';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

// One text shape: inner <p:spTree> XML for a p:sp at (x,y) sized (cx,cy) EMU,
// carrying a single run. Defaults: a 1in,1in box, 4in×2in.
function textShape(opts: {
  readonly text: string;
  readonly cx?: number;
  readonly cy?: number;
  readonly sz?: number;
  readonly spc?: number;
  readonly wrapNone?: boolean;
}): string {
  const { text, cx = 3657600, cy = 1828800, sz, spc, wrapNone } = opts;
  const rPr = `${sz !== undefined ? ` sz="${sz}"` : ''}${spc !== undefined ? ` spc="${spc}"` : ''}`;
  const bodyPr = wrapNone ? '<a:bodyPr wrap="none"/>' : '<a:bodyPr/>';
  return (
    `<p:sp><p:spPr>` +
    `<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</p:spPr>` +
    `<p:txBody>${bodyPr}<a:p><a:r><a:rPr${rPr}/><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
    `</p:sp>`
  );
}

// The first run of the first text-bearing shape in the parsed flow.
function firstShapeRun(doc: ReturnType<typeof Ream.parse>) {
  const text = firstShapeText(doc);
  if (!text) return undefined;
  for (const child of text.content) {
    if (child.kind === 'paragraph' && child.paragraph.runs.length > 0) {
      return child.paragraph.runs[0];
    }
  }
  return undefined;
}

// The first text-bearing shape's body in the parsed flow.
function firstShapeText(doc: ReturnType<typeof Ream.parse>) {
  for (const el of doc.flow.body) {
    if (el.kind === 'shape' && el.shape.text) return el.shape.text;
  }
  return undefined;
}

// The painted line's width in the PDF: the leftmost glyph origin to the pen
// position after the last glyph, across every piece of the single line.
async function paintedWidth(text: string, spc?: number): Promise<number> {
  const pptx = buildPptx([textShape({ text, sz: 2000, cx: 8000000, spc })]);
  const file = PdfFile.parse(await Ream.parse(pptx).convert('pdf', { fonts: FONTS }));
  const runs = extractPageText(file, file.pages()[0]!).filter(
    (r) => r.text.replace(/[\s\u200b]/g, '').length > 0,
  );
  expect(runs.length).toBeGreaterThan(0);
  const left = Math.min(...runs.map((r) => r.x));
  const right = Math.max(...runs.map((r) => r.endX));
  return right - left;
}

describe('a:rPr@spc — run tracking (§21.1.2.3.9)', () => {
  it('reads spc into letterSpacingPt', () => {
    const tracked = Ream.parse(buildPptx([textShape({ text: 'Tracked', sz: 2000, spc: 300 })]));
    expect(firstShapeRun(tracked)?.properties.letterSpacingPt).toBe(3);

    const plain = Ream.parse(buildPptx([textShape({ text: 'Tracked', sz: 2000 })]));
    expect(firstShapeRun(plain)?.properties.letterSpacingPt).toBeUndefined();
  });

  it('widens the painted line by the spacing between its glyphs', async () => {
    // Ten glyphs, spc=300: the PDF writer tracks with Tc, which widens every
    // glyph's advance by 3 pt — 30 pt over the untracked line.
    const plain = await paintedWidth('HHHHHHHHHH');
    const tracked = await paintedWidth('HHHHHHHHHH', 300);
    expect(tracked - plain).toBeCloseTo(30, 1);
  });
});

describe('a:bodyPr@wrap="none" — the box that does not wrap (§20.1.10.85)', () => {
  it('reads wrap="none" into noWrap', () => {
    const noWrap = Ream.parse(buildPptx([textShape({ text: 'Long line', wrapNone: true })]));
    expect(firstShapeText(noWrap)?.noWrap).toBe(true);

    const wrapped = Ream.parse(buildPptx([textShape({ text: 'Long line' })]));
    expect(firstShapeText(wrapped)?.noWrap).toBeUndefined();
  });

  it('keeps the paragraph on one line where the same box wraps', async () => {
    // 70 characters at 20 pt far exceed the 4 in box: wrapped, several
    // baselines; wrap="none", every glyph stays on the box's first line.
    const text = 'Thequickbrownfoxjumpsoverthelazydog'.repeat(2);
    const lines = async (wrapNone: boolean): Promise<number> => {
      const pptx = buildPptx([textShape({ text, sz: 2000, wrapNone })]);
      const file = PdfFile.parse(await Ream.parse(pptx).convert('pdf', { fonts: FONTS }));
      const ys = extractPageText(file, file.pages()[0]!)
        .filter((r) => r.text.replace(/[\s\u200b]/g, '').length > 0)
        .map((r) => Math.round(r.y));
      return new Set(ys).size;
    };
    expect(await lines(false)).toBeGreaterThan(2);
    expect(await lines(true)).toBe(1);
  });
});
