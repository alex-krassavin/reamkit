// E-PPTX PX1 — slide content. A text-bearing p:sp with an explicit a:xfrm is
// read into a floating text box positioned at its EMU offset on the slide, so
// the slide's text flows through to every target (here HTML for content, PDF
// for geometry). Placeholder text (no own a:xfrm) waits for PX2.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildPptx } from './fixtures/build-pptx';
import { buildTinyPng } from './fixtures/build-png';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

const latin1 = new TextDecoder('latin1');
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};
const decoder = new TextDecoder();

// One text shape: inner <p:spTree> XML for a p:sp at (x,y) sized (cx,cy) EMU,
// carrying a single run. Defaults: a 1in,1in box, 4in×2in.
function textShape(opts: {
  readonly text: string;
  readonly x?: number;
  readonly y?: number;
  readonly cx?: number;
  readonly cy?: number;
  readonly sz?: number;
  readonly bold?: boolean;
}): string {
  const { text, x = 914400, y = 914400, cx = 3657600, cy = 1828800, sz, bold } = opts;
  const rPr = `${sz !== undefined ? ` sz="${sz}"` : ''}${bold ? ' b="1"' : ''}`;
  return (
    `<p:sp><p:spPr>` +
    `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</p:spPr>` +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr${rPr}/><a:t>${text}</a:t></a:r></a:p></p:txBody>` +
    `</p:sp>`
  );
}

describe('pptx slide text (E-PPTX PX1)', () => {
  it('flows a slide text shape through to the HTML', async () => {
    const pptx = buildPptx([textShape({ text: 'Hello PPTX' })]);
    const html = decoder.decode(await Ream.parse(pptx).convert('html'));
    expect(html).toContain('Hello PPTX');
  });

  it("renders each slide's text on its own page", async () => {
    const pptx = buildPptx([textShape({ text: 'SlideOne' }), textShape({ text: 'SlideTwo' })]);
    const html = decoder.decode(await Ream.parse(pptx).convert('html'));
    expect(html).toContain('SlideOne');
    expect(html).toContain('SlideTwo');
    const pdf = await Ream.parse(pptx).convert('pdf', { fonts: FONTS });
    expect(PdfFile.parse(pdf).pages().length).toBe(2);
  });

  it('positions a text box at its EMU offset on the slide', async () => {
    // Box at 2in,1.5in on the default 16:9 deck (960×540 pt): 144 pt from the
    // left, 108 pt from the top.
    const pptx = buildPptx([textShape({ text: 'Positioned', x: 1828800, y: 1371600, sz: 2400 })]);
    const pdf = await Ream.parse(pptx).convert('pdf', { fonts: FONTS });
    const file = PdfFile.parse(pdf);
    const runs = extractPageText(file, file.pages()[0]!);
    const run = runs.find((r) => r.text.replace(/\s/g, '').includes('Positioned'));
    expect(run).toBeDefined();
    // The glyph origin sits just inside the box's left edge (144 pt) — not at
    // x≈0 (which would mean the float position was ignored).
    expect(run!.x).toBeGreaterThan(140);
    expect(run!.x).toBeLessThan(200);
    // PDF y grows upward: the box top is at 540 − 108 = 432 pt; the first
    // baseline sits a little below it, and well above the page middle.
    expect(run!.y).toBeGreaterThan(390);
    expect(run!.y).toBeLessThan(432);
  });

  it('carries bold run formatting from a:rPr', async () => {
    const pptx = buildPptx([textShape({ text: 'BoldText', bold: true, sz: 2000 })]);
    const html = decoder.decode(await Ream.parse(pptx).convert('html'));
    // The bold run renders inside a weighted span.
    expect(html).toMatch(/font-weight:\s*(bold|[67]00)/);
    expect(html).toContain('BoldText');
  });
});

// A layout title placeholder carrying geometry, and master text styles. The
// title's level-1 default is 44 pt bold.
const LAYOUT_TITLE =
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/>` +
  `<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="838200" y="457200"/><a:ext cx="7772400" cy="1143000"/></a:xfrm></p:spPr></p:sp>`;
const TX_STYLES =
  `<p:txStyles>` +
  `<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400" b="1"/></a:lvl1pPr></p:titleStyle>` +
  `<p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle>` +
  `<p:otherStyle/></p:txStyles>`;

// A slide title placeholder with NO own geometry, carrying `text` (and an
// optional own run size that should override the master default).
function titlePlaceholder(text: string, ownSz?: number): string {
  const rPr = ownSz !== undefined ? `<a:rPr sz="${ownSz}"/>` : '';
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/>` +
    `<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr/>` +
    `<p:txBody><a:bodyPr/><a:p><a:r>${rPr}<a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

// The first run of the first text-bearing shape in the parsed flow.
function firstShapeRun(doc: ReturnType<typeof Ream.parse>) {
  for (const el of doc.flow.body) {
    if (el.kind === 'shape' && el.shape.text) {
      for (const child of el.shape.text.content) {
        if (child.kind === 'paragraph' && child.paragraph.runs.length > 0) {
          return child.paragraph.runs[0];
        }
      }
    }
  }
  return undefined;
}

describe('pptx placeholder cascade (E-PPTX PX2)', () => {
  it('renders a placeholder that inherits its geometry from the layout', async () => {
    const pptx = buildPptx([titlePlaceholder('Inherited Title')], {
      layoutMaster: { layoutSpTree: LAYOUT_TITLE, txStyles: TX_STYLES },
    });
    // Built at all (a placeholder with no own xfrm) → the cascade supplied geometry.
    const html = decoder.decode(await Ream.parse(pptx).convert('html'));
    expect(html).toContain('Inherited Title');
    // Positioned at the layout's xfrm: x = 838200 EMU = 66 pt from the left.
    // (Multi-word text breaks into per-word runs, so match the first word.)
    const file = PdfFile.parse(await Ream.parse(pptx).convert('pdf', { fonts: FONTS }));
    const run = extractPageText(file, file.pages()[0]!).find((r) =>
      r.text.replace(/\s/g, '').includes('Inherited'),
    );
    expect(run).toBeDefined();
    expect(run!.x).toBeGreaterThan(60);
    expect(run!.x).toBeLessThan(110);
  });

  it('applies the master title text size to a placeholder run', () => {
    const pptx = buildPptx([titlePlaceholder('Sized Title')], {
      layoutMaster: { layoutSpTree: LAYOUT_TITLE, txStyles: TX_STYLES },
    });
    const run = firstShapeRun(Ream.parse(pptx));
    expect(run?.properties.fontSizePt).toBe(44); // titleStyle/lvl1pPr/defRPr sz=4400
    expect(run?.properties.bold).toBe(true);
  });

  it("lets a run's own a:rPr override the master default size", () => {
    const pptx = buildPptx([titlePlaceholder('Big', 6000)], {
      layoutMaster: { layoutSpTree: LAYOUT_TITLE, txStyles: TX_STYLES },
    });
    expect(firstShapeRun(Ream.parse(pptx))?.properties.fontSizePt).toBe(60); // own sz wins
  });

  it('drops a placeholder with no geometry anywhere (no layout/master)', async () => {
    const pptx = buildPptx([titlePlaceholder('Orphan')]); // no cascade to inherit from
    const html = decoder.decode(await Ream.parse(pptx).convert('html'));
    expect(html).not.toContain('Orphan');
  });
});

// A slide with a single p:pic at 2in,1in sized 3in×3in, its a:blip resolving
// through the slide rel to a 2×2 red PNG in ppt/media.
function picDeck(): Uint8Array {
  const pic =
    `<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture 4" descr="a red square"/>` +
    `<p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="2743200" cy="2743200"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  return buildPptx([pic], {
    media: { 'ppt/media/image1.png': buildTinyPng(2, 2, [255, 0, 0, 255]) },
    slideRels: [`<Relationship Id="rId7" Type="${IMAGE_REL}" Target="../media/image1.png"/>`],
  });
}

describe('pptx slide images (E-PPTX PX3)', () => {
  it('reads a p:pic into a positioned image with its bytes in the store', () => {
    const doc = Ream.parse(picDeck());
    const el = doc.flow.body.find((e) => e.kind === 'image');
    expect(el?.kind).toBe('image');
    if (el?.kind !== 'image') return;
    const img = el.image;
    // The blip resolved to a stored resource (its bytes are the PNG).
    expect(img.resource).toBeDefined();
    expect(doc.flow.resources.get(img.resource!)).toBeDefined();
    // ext 2743200 EMU = 216 pt; off 1828800,914400 EMU = 144,72 pt from the page.
    expect(Math.round(img.width)).toBe(216);
    expect(Math.round(img.height)).toBe(216);
    expect(img.float?.posH?.relativeFrom).toBe('page');
    expect(Math.round(img.float?.posH?.offsetPt ?? -1)).toBe(144);
    expect(Math.round(img.float?.posV?.offsetPt ?? -1)).toBe(72);
    expect(img.altText).toBe('a red square'); // p:cNvPr @descr
  });

  it('embeds the slide image into the rendered PDF', async () => {
    const pdf = await Ream.parse(picDeck()).convert('pdf', { fonts: FONTS });
    expect(PdfFile.parse(pdf).pages().length).toBe(1);
    // An image XObject made it into the PDF.
    expect(latin1.decode(pdf)).toContain('/Image');
  });
});
