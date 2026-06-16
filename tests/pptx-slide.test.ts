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

// One p:sp at a fixed box with the given inner p:spPr children (geometry/fill/…).
function shapeDeck(spPrInner: string): Uint8Array {
  return buildPptx([
    `<p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>` +
      `${spPrInner}</p:spPr></p:sp>`,
  ]);
}
function firstShape(doc: ReturnType<typeof Ream.parse>) {
  const el = doc.flow.body.find((e) => e.kind === 'shape');
  return el?.kind === 'shape' ? el.shape : undefined;
}

describe('pptx slide shapes (E-PPTX PX3)', () => {
  it('reads solid fill, stroke and preset geometry on a textless shape', () => {
    const shp = firstShape(
      Ream.parse(
        shapeDeck(
          `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
            `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>` +
            `<a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>`,
        ),
      ),
    );
    expect(shp).toBeDefined(); // no text, but a visible fill keeps it
    expect(shp?.fill.kind).toBe('solid');
    expect(shp?.fill.colorHex).toBe('FF0000');
    expect(shp?.geometry.preset).toBe('roundRect');
    expect(shp?.line?.colorHex).toBe('0000FF');
    expect(shp?.text).toBeUndefined();
  });

  it('reads a gradient fill', () => {
    const shp = firstShape(
      Ream.parse(
        shapeDeck(
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
            `<a:gradFill><a:gsLst>` +
            `<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>` +
            `<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>` +
            `</a:gsLst><a:lin ang="0"/></a:gradFill>`,
        ),
      ),
    );
    expect(shp?.fill.kind).toBe('gradient');
    expect(shp?.fill.gradient?.stops.length).toBe(2);
  });

  it('drops an invisible shape (no fill, no stroke, no text)', () => {
    const shp = firstShape(
      Ream.parse(shapeDeck(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>`)),
    );
    expect(shp).toBeUndefined();
  });

  it('renders a filled shape into the PDF', async () => {
    const pdf = await Ream.parse(
      shapeDeck(
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
          `<a:solidFill><a:srgbClr val="00AA00"/></a:solidFill>`,
      ),
    ).convert('pdf', { fonts: FONTS });
    expect(PdfFile.parse(pdf).pages().length).toBe(1); // a textless filled box still makes a page
  });
});

const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const BAR_CHART =
  `<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/>` +
  `<c:cat><c:strRef><c:strCache><c:ptCount val="2"/>` +
  `<c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:numCache><c:ptCount val="2"/>` +
  `<c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;

// A slide with a c:chart graphicFrame at 2in,1in sized 6in×3.5in, its r:id
// resolving through the slide rel to a bar chart part.
function chartDeck(): Uint8Array {
  const gf =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Chart 5"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="1828800" y="914400"/><a:ext cx="5486400" cy="3200400"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${CHART_NS}">` +
    `<c:chart xmlns:c="${CHART_NS}" r:id="rId8"/>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`;
  return buildPptx([gf], {
    media: { 'ppt/charts/chart1.xml': new TextEncoder().encode(BAR_CHART) },
    slideRels: [`<Relationship Id="rId8" Type="${CHART_REL}" Target="../charts/chart1.xml"/>`],
  });
}

describe('pptx slide charts (E-PPTX PX4)', () => {
  it('reads a c:chart graphicFrame into a positioned ChartBlock', () => {
    const doc = Ream.parse(chartDeck());
    const el = doc.flow.body.find((e) => e.kind === 'chart');
    expect(el?.kind).toBe('chart');
    if (el?.kind !== 'chart') return;
    const ch = el.chart;
    // The block references a parsed chart in the document's charts map.
    expect(doc.flow.charts?.get(ch.chartRelId)?.type).toBe('bar');
    // ext 5486400×3200400 EMU = 432×252 pt; off 1828800,914400 = 144,72 pt.
    expect(Math.round(ch.width)).toBe(432);
    expect(Math.round(ch.height)).toBe(252);
    expect(Math.round(ch.float?.posH?.offsetPt ?? -1)).toBe(144);
    expect(Math.round(ch.float?.posV?.offsetPt ?? -1)).toBe(72);
  });

  it('renders the slide chart into the PDF', async () => {
    const pdf = await Ream.parse(chartDeck()).convert('pdf', { fonts: FONTS });
    expect(PdfFile.parse(pdf).pages().length).toBe(1);
  });
});

const TABLE_NS = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const GRID2 = `<a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid>`;

// An a:tc with text + optional attributes (gridSpan/rowSpan/vMerge/hMerge) + tcPr.
function cell(text: string, attrs = '', tcPr = ''): string {
  return (
    `<a:tc${attrs ? ` ${attrs}` : ''}><a:txBody><a:bodyPr/>` +
    `<a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody>${tcPr}</a:tc>`
  );
}

// A slide carrying a table graphicFrame whose a:tbl is `tblInner` (grid + rows).
function tableDeck(tblInner: string): Uint8Array {
  const gf =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table 6"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1828800"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${TABLE_NS}"><a:tbl>${tblInner}</a:tbl>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`;
  return buildPptx([gf]);
}
function firstTable(doc: ReturnType<typeof Ream.parse>) {
  const el = doc.flow.body.find((e) => e.kind === 'table');
  return el?.kind === 'table' ? el.table : undefined;
}

describe('pptx slide tables (E-PPTX PX4)', () => {
  it('reads an a:tbl graphicFrame into a FlowDoc table', () => {
    const tbl = firstTable(
      Ream.parse(
        tableDeck(
          GRID2 +
            `<a:tr h="370840">` +
            cell(
              'H1',
              '',
              `<a:tcPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:tcPr>`,
            ) +
            cell('H2') +
            `</a:tr><a:tr h="370840">${cell('A1')}${cell('B1')}</a:tr>`,
        ),
      ),
    );
    expect(tbl).toBeDefined();
    expect(tbl?.grid.length).toBe(2);
    expect(Math.round(tbl!.grid[0]!)).toBe(216); // 2743200 EMU = 216 pt
    expect(tbl?.rows.length).toBe(2);
    expect(tbl?.rows[0]?.cells.length).toBe(2);
    expect(tbl?.rows[0]?.cells[0]?.properties.shading?.colorHex).toBe('4472C4');
  });

  it('flows table cell text through to the HTML', async () => {
    const html = decoder.decode(
      await Ream.parse(
        tableDeck(GRID2 + `<a:tr>${cell('CellText')}${cell('Other')}</a:tr>`),
      ).convert('html'),
    );
    expect(html).toContain('CellText');
    expect(html).toContain('Other');
  });

  it('honours gridSpan (colSpan) and drops the hMerge continuation', () => {
    const tbl = firstTable(
      Ream.parse(
        tableDeck(GRID2 + `<a:tr>${cell('Wide', 'gridSpan="2"')}${cell('', 'hMerge="1"')}</a:tr>`),
      ),
    );
    expect(tbl?.rows[0]?.cells.length).toBe(1); // continuation dropped
    expect(tbl?.rows[0]?.cells[0]?.properties.colSpan).toBe(2);
  });

  it('renders a slide table into the PDF', async () => {
    const pdf = await Ream.parse(
      tableDeck(GRID2 + `<a:tr>${cell('X')}${cell('Y')}</a:tr>`),
    ).convert('pdf', { fonts: FONTS });
    expect(PdfFile.parse(pdf).pages().length).toBe(1);
  });
});

// A positioned shape filled with the accent1 scheme colour.
const SCHEME_SHAPE =
  `<p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill></p:spPr></p:sp>`;

describe('pptx deck theme (E-PPTX PX5)', () => {
  it('resolves a scheme-colour fill through the deck theme', () => {
    const pptx = buildPptx([SCHEME_SHAPE], {
      layoutMaster: { theme: `<a:accent1><a:srgbClr val="FF8800"/></a:accent1>` },
    });
    const shp = firstShape(Ream.parse(pptx));
    expect(shp?.fill.kind).toBe('solid');
    expect(shp?.fill.colorHex).toBe('FF8800'); // the deck's accent1, not the default
  });

  it('falls back to the Office palette when the deck has no theme', () => {
    const shp = firstShape(Ream.parse(buildPptx([SCHEME_SHAPE])));
    expect(shp?.fill.colorHex).toBe('4472C4'); // default Office accent1
  });
});

const bgFill = (hex: string): string =>
  `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></p:bgPr></p:bg>`;

describe('pptx slide backgrounds (E-PPTX PX5b)', () => {
  it('renders the slide background as a behind-everything full-slide backdrop', () => {
    const doc = Ream.parse(buildPptx([''], { slideBg: [bgFill('112233')] }));
    const shp = firstShape(doc);
    expect(shp).toBeDefined();
    expect(shp?.fill.colorHex).toBe('112233');
    expect(shp?.float?.behind).toBe(true);
    expect(Math.round(shp!.width)).toBe(960); // the full 16:9 deck
    expect(Math.round(shp!.height)).toBe(540);
  });

  it('inherits the master background when the slide has none', () => {
    const doc = Ream.parse(buildPptx([''], { layoutMaster: { masterBg: bgFill('445566') } }));
    expect(firstShape(doc)?.fill.colorHex).toBe('445566');
  });

  it("prefers the slide's own background over the master's", () => {
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [bgFill('112233')],
        layoutMaster: { masterBg: bgFill('445566') },
      }),
    );
    expect(firstShape(doc)?.fill.colorHex).toBe('112233');
  });
});

// A p:grpSp whose child shape (filled `hex`) sits at child-box (cx,cy,ex,ey),
// inside the group transform off/ext (chOff 0, chExt = `chExt`).
function groupDeck(opts: {
  readonly off: [number, number];
  readonly ext: [number, number];
  readonly chExt: [number, number];
  readonly child: [number, number, number, number]; // x, y, cx, cy
  readonly hex: string;
}): Uint8Array {
  const [ox, oy] = opts.off;
  const [ex, ey] = opts.ext;
  const [chx, chy] = opts.chExt;
  const [cx, cy, ccx, ccy] = opts.child;
  const grp =
    `<p:grpSp><p:grpSpPr><a:xfrm><a:off x="${ox}" y="${oy}"/><a:ext cx="${ex}" cy="${ey}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${chx}" cy="${chy}"/></a:xfrm></p:grpSpPr>` +
    `<p:sp><p:spPr><a:xfrm><a:off x="${cx}" y="${cy}"/><a:ext cx="${ccx}" cy="${ccy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${opts.hex}"/></a:solidFill></p:spPr></p:sp></p:grpSp>`;
  return buildPptx([grp]);
}

describe('pptx slide groups (E-PPTX PX5c)', () => {
  it('offsets a grouped shape into slide coordinates', () => {
    // chExt == ext → scale 1; the child is just shifted by the group's off.
    const shp = firstShape(
      Ream.parse(
        groupDeck({
          off: [1828800, 914400], // 144, 72 pt
          ext: [3657600, 3657600],
          chExt: [3657600, 3657600],
          child: [914400, 0, 914400, 914400],
          hex: 'FF0000',
        }),
      ),
    );
    expect(shp?.fill.colorHex).toBe('FF0000');
    // 1828800 + 914400 = 2743200 EMU = 216 pt; y = 914400 = 72 pt; size 72 pt.
    expect(Math.round(shp?.float?.posH?.offsetPt ?? -1)).toBe(216);
    expect(Math.round(shp?.float?.posV?.offsetPt ?? -1)).toBe(72);
    expect(Math.round(shp!.width)).toBe(72);
  });

  it('scales a grouped shape by the group ext / chExt ratio', () => {
    // ext is half of chExt → scale 0.5.
    const shp = firstShape(
      Ream.parse(
        groupDeck({
          off: [0, 0],
          ext: [1828800, 1828800],
          chExt: [3657600, 3657600],
          child: [0, 0, 1828800, 1828800],
          hex: '00FF00',
        }),
      ),
    );
    expect(Math.round(shp!.width)).toBe(72); // 1828800 * 0.5 = 914400 EMU = 72 pt
  });
});

const HLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

// A text shape at a fixed box with the given p:txBody inner XML.
function textBodyShape(txBodyInner: string): string {
  return (
    `<p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${txBodyInner}</p:txBody></p:sp>`
  );
}
function firstShapeParagraph(doc: ReturnType<typeof Ream.parse>) {
  const el = doc.flow.body.find((e) => e.kind === 'shape');
  if (el?.kind !== 'shape') return undefined;
  const p = el.shape.text?.content[0];
  return p?.kind === 'paragraph' ? p.paragraph : undefined;
}

describe('pptx text depth + links (E-PPTX PX6)', () => {
  it('reads paragraph alignment and the body vertical anchor', () => {
    const doc = Ream.parse(
      buildPptx([
        textBodyShape(
          `<a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:t>Centered</a:t></a:r></a:p>`,
        ),
      ]),
    );
    const el = doc.flow.body.find((e) => e.kind === 'shape');
    expect(el?.kind === 'shape' && el.shape.text?.anchor).toBe('ctr');
    expect(firstShapeParagraph(doc)?.properties.alignment).toBe('center');
  });

  it('resolves a run hyperlink to its external URL and renders an anchor', async () => {
    const pptx = buildPptx(
      [
        textBodyShape(
          `<a:bodyPr/><a:p><a:r><a:rPr><a:hlinkClick r:id="rId5"/></a:rPr><a:t>ClickMe</a:t></a:r></a:p>`,
        ),
      ],
      {
        slideRels: [
          `<Relationship Id="rId5" Type="${HLINK_REL}" Target="https://example.com/" TargetMode="External"/>`,
        ],
      },
    );
    expect(firstShapeParagraph(Ream.parse(pptx))?.runs[0]?.href).toBe('https://example.com/');
    const html = decoder.decode(await Ream.parse(pptx).convert('html'));
    expect(html).toContain('href="https://example.com/"');
  });
});

// An a:p with the given a:pPr inner XML and a single run.
function para(pPrInner: string, text: string): string {
  return `<a:p><a:pPr>${pPrInner}</a:pPr><a:r><a:t>${text}</a:t></a:r></a:p>`;
}
function shapeParagraphs(doc: ReturnType<typeof Ream.parse>) {
  const el = doc.flow.body.find((e) => e.kind === 'shape');
  return el?.kind === 'shape' ? el.shape.text?.content : undefined;
}

describe('pptx bullets + indent (E-PPTX PX6b)', () => {
  it('materializes a buChar bullet as a leading list-marker run', () => {
    const doc = Ream.parse(
      buildPptx([textBodyShape(`<a:bodyPr/>${para('<a:buChar char="•"/>', 'Item')}`)]),
    );
    const p = firstShapeParagraph(doc);
    expect(p?.runs[0]?.listMarker).toBe(true);
    expect(p?.runs[0]?.text).toContain('•');
    expect(p?.runs[1]?.text).toBe('Item');
  });

  it('numbers buAutoNum paragraphs per level', () => {
    const body =
      `<a:bodyPr/>` +
      para('<a:buAutoNum type="arabicPeriod"/>', 'One') +
      para('<a:buAutoNum type="arabicPeriod"/>', 'Two');
    const paras = shapeParagraphs(Ream.parse(buildPptx([textBodyShape(body)])));
    const marker = (i: number) =>
      paras?.[i]?.kind === 'paragraph' ? paras[i].paragraph.runs[0]?.text.trim() : undefined;
    expect(marker(0)).toBe('1.');
    expect(marker(1)).toBe('2.');
  });

  it('suppresses a buNone bullet and indents by outline level', () => {
    const doc = Ream.parse(
      buildPptx([
        textBodyShape(
          `<a:bodyPr/><a:p><a:pPr lvl="1"><a:buNone/></a:pPr><a:r><a:t>Plain</a:t></a:r></a:p>`,
        ),
      ]),
    );
    const p = firstShapeParagraph(doc);
    expect(p?.runs[0]?.listMarker).toBeUndefined(); // no marker
    expect(p?.runs[0]?.text).toBe('Plain');
    expect(Math.round(p?.properties.indentLeft ?? -1)).toBe(36); // level 1 × 0.5"
  });
});

// E-SMARTART SA0 — a graphicFrame referencing a SmartArt data part (dgm:relIds
// @r:dm) whose pre-rendered drawing override (diagrams/drawing1.xml) holds two
// dsp:sp nodes. The reader follows slide → data1.xml → drawing1.xml and renders
// the nodes as floating shapes positioned within the frame box.
const DIAGRAM_DATA_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData';
const DIAGRAM_DRAWING_REL = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
const A_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const srgbFill = (hex: string): string => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
// accent1 scheme fill markup — for the SA3 theme-resolution test.
const SCHEME_ACCENT1_FILL = `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`;

function smartArtDeck(
  opts: { readonly fillA?: string; readonly build?: Parameters<typeof buildPptx>[1] } = {},
): Uint8Array {
  const frame =
    `<p:graphicFrame>` +
    `<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="2743200"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
    `<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ` +
    `r:dm="rId100" r:lo="rId101" r:qs="rId102" r:cs="rId103"/>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`;

  const node = (text: string, x: number, fill: string): string =>
    `<dsp:sp><dsp:spPr>` +
    `<a:xfrm><a:off x="${x}" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `${fill}` +
    `</dsp:spPr>` +
    `<dsp:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dsp:txBody>` +
    `</dsp:sp>`;

  const drawing =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="${A_MAIN}">` +
    `<dsp:spTree>` +
    node('NodeA', 0, opts.fillA ?? srgbFill('4472C4')) +
    node('NodeB', 2743200, srgbFill('ED7D31')) +
    `</dsp:spTree></dsp:drawing>`;

  const enc = new TextEncoder();
  const rels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="${DIAGRAM_DRAWING_REL}" Target="drawing1.xml"/>` +
    `</Relationships>`;
  return buildPptx([frame], {
    slideRels: [
      `<Relationship Id="rId100" Type="${DIAGRAM_DATA_REL}" Target="../diagrams/data1.xml"/>`,
    ],
    media: {
      'ppt/diagrams/data1.xml': enc.encode(
        `<?xml version="1.0"?>\n<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>`,
      ),
      'ppt/diagrams/_rels/data1.xml.rels': enc.encode(rels),
      'ppt/diagrams/drawing1.xml': enc.encode(drawing),
    },
    ...opts.build,
  });
}

function shapeTexts(doc: ReturnType<typeof Ream.parse>): Array<string> {
  const out: Array<string> = [];
  for (const el of doc.flow.body) {
    if (el.kind === 'shape' && el.shape.text) {
      const words = el.shape.text.content
        .flatMap((c) => (c.kind === 'paragraph' ? c.paragraph.runs.map((r) => r.text) : []))
        .join('');
      if (words) out.push(words);
    }
  }
  return out;
}

describe('SmartArt diagrams (E-SMARTART SA0)', () => {
  it('renders the drawing-override nodes as floating shapes', () => {
    const texts = shapeTexts(Ream.parse(smartArtDeck()));
    expect(texts).toContain('NodeA');
    expect(texts).toContain('NodeB');
  });

  it('positions each node within the frame box', () => {
    const xs = Ream.parse(smartArtDeck())
      .flow.body.filter((e) => e.kind === 'shape')
      .map((e) => Math.round(e.shape.float?.posH?.offsetPt ?? -1))
      .sort((a, b) => a - b);
    // NodeA at the frame offset (914400 EMU = 72pt); NodeB at +2743200 EMU (+216pt).
    expect(xs).toEqual([72, 288]);
  });

  it('flows the diagram text through to PDF', async () => {
    const file = PdfFile.parse(await Ream.parse(smartArtDeck()).convert('pdf', { fonts: FONTS }));
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('')
      .replace(/\s/g, '');
    expect(text).toContain('NodeA');
    expect(text).toContain('NodeB');
  });

  it('resolves a node scheme-colour fill through the deck theme (SA3)', () => {
    // NodeA fills with accent1; the deck theme maps accent1 → FF8800. The shared
    // ColorResolver that styles ordinary slide shapes styles diagram shapes too.
    const deck = smartArtDeck({
      fillA: SCHEME_ACCENT1_FILL,
      build: { layoutMaster: { theme: `<a:accent1><a:srgbClr val="FF8800"/></a:accent1>` } },
    });
    // NodeA (the first diagram shape) fills with accent1, mapped to FF8800 by the
    // deck theme — the same ColorResolver path that styles ordinary slide shapes.
    expect(firstShape(Ream.parse(deck))?.fill.colorHex).toBe('FF8800');
  });

  it('degrades gracefully — and records a loss — when no drawing override ships', () => {
    // Same frame, but no diagrams/* parts → resolveDiagram yields nothing.
    const frame =
      `<p:graphicFrame>` +
      `<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="2743200"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId100"/>` +
      `</a:graphicData></a:graphic></p:graphicFrame>`;
    const doc = Ream.parse(buildPptx([frame]));
    expect(doc.flow.body.filter((e) => e.kind === 'shape')).toHaveLength(0);
    // SA3: the diagram is dropped explicitly, located to the slide it sat on.
    const loss = doc.losses.find((l) => l.feature === 'shapes.smartArt');
    expect(loss?.severity).toBe('dropped');
    expect(loss?.where).toBe('slide 1');
  });
});
