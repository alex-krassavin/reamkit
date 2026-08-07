// E-PPTX PX1 — slide content. A text-bearing p:sp with an explicit a:xfrm is
// read into a floating text box positioned at its EMU offset on the slide, so
// the slide's text flows through to every target (here HTML for content, PDF
// for geometry). Placeholder text (no own a:xfrm) waits for PX2.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildPptx } from './fixtures/build-pptx';
import { buildTinyPng } from './fixtures/build-png';
import type { ShapeFill } from '@/core/document-model';
import { FontRegistry } from '@/core/font';
import { paintPlan } from '@/layout/page-doc';
import { layoutStyledDocument } from '@/layout/styled-layout';
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

  // §21.1.2.2.5/.9/.10 — both spacings come in two spellings: `a:spcPts` is a
  // distance, `a:spcPct` a FRACTION of a line. Only the distance was read, so a
  // deck that spaces its bullets the usual way — by fraction — set them solid.
  it('reads the line height and the space around a paragraph stated as fractions', () => {
    const styles =
      `<p:txStyles><p:titleStyle/><p:bodyStyle><a:lvl1pPr>` +
      `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>` +
      `<a:spcBef><a:spcPct val="20000"/></a:spcBef>` +
      `<a:defRPr sz="3000"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles>`;
    const slide = (pPr: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="1828800"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p>${pPr}<a:r><a:rPr lang="en"/><a:t>Spaced</a:t></a:r></a:p>` +
      `<a:p><a:endParaRPr lang="en"/></a:p></p:txBody></p:sp>`;
    const paras = (pPr: string) => {
      const doc = Ream.parse(buildPptx([slide(pPr)], { layoutMaster: { txStyles: styles } }));
      const el = doc.flow.body.find((e) => e.kind === 'shape');
      return el?.kind === 'shape'
        ? (el.shape.text?.content.flatMap((c) => (c.kind === 'paragraph' ? [c.paragraph] : [])) ??
            [])
        : [];
    };
    // From the level: 90% of a line (12pt states "single") and 20% of one
    // before the paragraph, resolved against the 30pt the same level gives.
    const [first, blank] = paras('');
    expect(first?.properties.spacingLine).toBeCloseTo(10.8, 5);
    expect(first?.properties.spacingLineRule).toBe('auto');
    expect(first?.properties.spacingBefore).toBeCloseTo(0.2 * 30 * 1.2, 5);
    // §21.1.2.2.3 — a paragraph with no runs is as tall as its MARK, which
    // takes the level's size when it states none of its own.
    expect(blank?.runs).toHaveLength(0);
    expect(blank?.properties.runProperties?.fontSizePt).toBe(30);
    // The paragraph's own spelling wins, in either form.
    const own = paras(
      '<a:pPr><a:lnSpc><a:spcPts val="1800"/></a:lnSpc>' +
        '<a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>',
    )[0];
    expect(own?.properties.spacingLine).toBe(18);
    expect(own?.properties.spacingLineRule).toBe('exact');
    expect(own?.properties.spacingBefore).toBe(6);
  });

  // ISO/IEC 29500-3 §10.2 — a deck writes the same shape twice: an `mc:Choice`
  // for the application whose namespace it names, and an `mc:Fallback` in plain
  // OOXML for everyone else. Reading the choice means reading markup written
  // for someone else — tdf143222's whole slide is an embedded worksheet whose
  // preview picture lives in the fallback alone, and it came out blank.
  it('takes the mc:Fallback of an AlternateContent, not the choice', () => {
    const alt =
      `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
      `<mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">` +
      `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>` +
      `</mc:Choice><mc:Fallback>` +
      `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:spPr></p:sp>` +
      `</mc:Fallback></mc:AlternateContent>`;
    const fills = Ream.parse(buildPptx([alt])).flow.body.flatMap((el) =>
      el.kind === 'shape' ? [el.shape.fill.colorHex] : [],
    );
    expect(fills).toEqual(['00FF00']); // the fallback's green, and only once
  });

  // §21.1.2.1.2 `a:normAutofit` — the shrink PowerPoint already worked out and
  // wrote down. Unread, the text it squeezed to 62% of its style overflowed the
  // placeholder it was squeezed to fit.
  it('sets the text at the scale the autofit already settled on', () => {
    const body = (bodyPr: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm></p:spPr>` +
      `<p:txBody>${bodyPr}<a:p><a:r><a:rPr lang="en" sz="3200"/><a:t>Squeezed</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`;
    const fittedPara = (bodyPr: string) => {
      const doc = Ream.parse(buildPptx([body(bodyPr)]));
      const el = doc.flow.body.find((e) => e.kind === 'shape');
      const block =
        el?.kind === 'shape'
          ? el.shape.text?.content.find((c) => c.kind === 'paragraph')
          : undefined;
      return block?.kind === 'paragraph' ? block.paragraph : undefined;
    };
    // 62.5% of 32pt, and a fifth off the line spacing (12pt states "single").
    const fitted = fittedPara(
      '<a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>',
    );
    expect(fitted?.runs[0]?.properties.fontSizePt).toBeCloseTo(20, 5);
    expect(fitted?.properties.spacingLine).toBeCloseTo(9.6, 5);
    expect(fitted?.properties.spacingLineRule).toBe('auto');
    // A bare normAutofit states the RULE and no result — that shrink is ours to
    // compute, and until we do there is nothing to apply.
    expect(
      fittedPara('<a:bodyPr><a:normAutofit/></a:bodyPr>')?.runs[0]?.properties.fontSizePt,
    ).toBe(32);
    expect(fittedPara('<a:bodyPr/>')?.runs[0]?.properties.fontSizePt).toBe(32);
  });

  // §20.1.4.1.14 — a slide names its typeface by TOKEN more often than by name:
  // `+mn-lt` is "whatever the theme calls its body font". Left unresolved the
  // token travels into the model as if it WERE a typeface, no substitution
  // table knows it, and 45541_Header's Times deck came out in a grotesque.
  it('resolves a +mn-lt/+mj-lt typeface through the theme font scheme', () => {
    const fontScheme =
      `<a:fontScheme name="t"><a:majorFont><a:latin typeface="Georgia"/>` +
      `<a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
      `<a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface="MS Mincho"/>` +
      `<a:cs typeface=""/></a:minorFont></a:fontScheme>`;
    const styles =
      `<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr>` +
      `<a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
      `<p:bodyStyle/><p:otherStyle/></p:txStyles>`;
    const run = (rPr: string) =>
      firstShapeRun(
        Ream.parse(
          buildPptx(
            [
              `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/>` +
                `<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>` +
                `<p:txBody><a:bodyPr/><a:p><a:r>${rPr}<a:t>Themed</a:t></a:r></a:p>` +
                `</p:txBody></p:sp>`,
            ],
            {
              layoutMaster: {
                layoutSpTree: LAYOUT_TITLE,
                txStyles: styles,
                themeFonts: fontScheme,
              },
            },
          ),
        ),
      );
    // The master's title style names the MAJOR font by token…
    expect(run('<a:rPr lang="en"/>')?.properties.fontFamily?.ascii).toBe('Georgia');
    // …and the run's own token wins over it, resolved the same way.
    expect(
      run('<a:rPr lang="en"><a:latin typeface="+mn-lt"/></a:rPr>')?.properties.fontFamily?.ascii,
    ).toBe('Times New Roman');
    // A typeface stated by NAME is untouched.
    expect(
      run('<a:rPr lang="en"><a:latin typeface="Verdana"/></a:rPr>')?.properties.fontFamily?.ascii,
    ).toBe('Verdana');
  });

  // A run states bold/italic/underline to turn them OFF as much as on:
  // 45541_Header's master body style is bold and every slide's own runs say
  // `b="0"`, so read as "bold when true, silent otherwise" the deck came out
  // bold from end to end.
  it('lets a run turn OFF the bold its master states', () => {
    const styles =
      `<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr b="1" i="1" u="sng"/></a:lvl1pPr>` +
      `</p:titleStyle><p:bodyStyle/><p:otherStyle/></p:txStyles>`;
    const deck = (rPr: string) =>
      firstShapeRun(
        Ream.parse(
          buildPptx(
            [
              `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/>` +
                `<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>` +
                `<p:txBody><a:bodyPr/><a:p><a:r>${rPr}<a:t>Plain</a:t></a:r></a:p>` +
                `</p:txBody></p:sp>`,
            ],
            { layoutMaster: { layoutSpTree: LAYOUT_TITLE, txStyles: styles } },
          ),
        ),
      );
    expect(deck('<a:rPr lang="en"/>')?.properties.bold).toBe(true);
    expect(deck('<a:rPr lang="en" b="0"/>')?.properties.bold).toBe(false);
    expect(deck('<a:rPr lang="en" i="0"/>')?.properties.italic).toBe(false);
    expect(deck('<a:rPr lang="en" u="none"/>')?.properties.underline).toBe('none');
  });

  // §21.1.2.2.7 `@cap` — the text is DISPLAYED in capitals whatever it stores.
  // Stated in a master's title style it reaches every slide: themes.pptx writes
  // "Trade show" and shows TRADE SHOW.
  it('takes the capitals its master states, and gives them back on cap="none"', () => {
    const caps =
      `<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr cap="all"/></a:lvl1pPr></p:titleStyle>` +
      `<p:bodyStyle/><p:otherStyle/></p:txStyles>`;
    const deck = (rPr: string): ReturnType<typeof Ream.parse> =>
      Ream.parse(
        buildPptx(
          [
            `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/>` +
              `<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>` +
              `<p:txBody><a:bodyPr/><a:p><a:r>${rPr}<a:t>Trade show</a:t></a:r></a:p>` +
              `</p:txBody></p:sp>`,
          ],
          { layoutMaster: { layoutSpTree: LAYOUT_TITLE, txStyles: caps } },
        ),
      );
    expect(firstShapeRun(deck('<a:rPr lang="en"/>'))?.properties.caps).toBe(true);
    expect(firstShapeRun(deck('<a:rPr lang="en" cap="none"/>'))?.properties.caps).toBe(false);
    expect(firstShapeRun(deck('<a:rPr lang="en" cap="small"/>'))?.properties.smallCaps).toBe(true);
  });

  // §21.1.2.4.x — a paragraph states a bullet only where it differs from the
  // one its level already carries.
  it("draws the bullet the master's body style names, without one on the slide", () => {
    const body =
      `<p:txStyles><p:titleStyle/>` +
      `<p:bodyStyle><a:lvl1pPr marL="285750" indent="-285750">` +
      `<a:buChar char="–"/><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle>` +
      `<p:otherStyle/></p:txStyles>`;
    const slide =
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>With a solid fill</a:t></a:r>` +
      `</a:p></p:txBody></p:sp>`;
    const doc = Ream.parse(buildPptx([slide], { layoutMaster: { txStyles: body } }));
    const runs = doc.flow.body.flatMap((el) =>
      el.kind === 'shape' && el.shape.text
        ? el.shape.text.content.flatMap((c) => (c.kind === 'paragraph' ? c.paragraph.runs : []))
        : [],
    );
    expect(runs[0]?.listMarker).toBe(true);
    expect(runs[0]?.text.trim()).toBe('–');
    expect(runs[1]?.text).toBe('With a solid fill');
    // §17.3.1.12 — a hanging indent is itself a tab stop, and the tab after the
    // marker is what carries the text out to the body indent (marL 285750).
    expect(runs[0]?.text).toBe('–\t');
  });

  // §21.1.2.4.5 `a:buFont` — the character is stated IN THAT FACE, and the
  // symbol faces are not alphabets: Wingdings `l` is a filled circle. Read as a
  // letter it printed a column of `l`s down all eleven slides of 45541_Header.
  // Its size and colour come from the level too, not only from the paragraph.
  it('reads a symbol-font bullet, and takes its size and colour from the level', async () => {
    const body =
      `<p:txStyles><p:titleStyle/>` +
      `<p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900">` +
      `<a:buClr><a:srgbClr val="00B0F0"/></a:buClr><a:buSzPct val="80000"/>` +
      `<a:buFont typeface="Wingdings"/><a:buChar char="l"/>` +
      `<a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles>`;
    const slide =
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>Donors</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`;
    const pptx = buildPptx([slide], { layoutMaster: { txStyles: body } });
    const marker = Ream.parse(pptx).flow.body.flatMap((el) =>
      el.kind === 'shape' && el.shape.text
        ? el.shape.text.content.flatMap((c) =>
            c.kind === 'paragraph' ? c.paragraph.runs.filter((r) => r.listMarker) : [],
          )
        : [],
    )[0];
    expect(marker?.text.trim()).toBe('●'); // Wingdings `l` is a black circle
    expect(marker?.properties.colorHex).toBe('00B0F0');
    expect(marker?.properties.fontSizePt).toBeCloseTo(25.6, 5); // 80% of 32pt
    // …and the font a document is RENDERED with may have no such glyph, which
    // prints a box. The dot every text font carries says the same thing.
    const file = PdfFile.parse(await Ream.parse(pptx).convert('pdf', { fonts: FONTS }));
    const drawn = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('');
    expect(drawn).toContain('•');
    expect(drawn).not.toContain('●');
  });

  // The page, the band and a table cell all apply a paragraph's indent; a TEXT
  // BOX did not, so a slide's bullets stood in the margin and their words began
  // wherever the dot ended.
  it('stands the bulleted text at the indent its paragraph asks for', async () => {
    const body =
      `<p:txStyles><p:titleStyle/>` +
      `<p:bodyStyle><a:lvl1pPr marL="914400" indent="-914400">` +
      `<a:buChar char="•"/><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle>` +
      `<p:otherStyle/></p:txStyles>`;
    const slide =
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr lIns="0"/><a:p><a:r><a:rPr lang="en"/><a:t>Indented</a:t></a:r>` +
      `</a:p></p:txBody></p:sp>`;
    const pptx = buildPptx([slide], { layoutMaster: { txStyles: body } });
    const file = PdfFile.parse(await Ream.parse(pptx).convert('pdf', { fonts: FONTS }));
    const runs = extractPageText(file, file.pages()[0]!);
    const word = runs.find((r) => r.text.replace(/\s/gu, '').includes('Indented'));
    // marL 914400 EMU = 72 pt from the box's left edge (which is the page's).
    expect(word?.x).toBeGreaterThan(66);
    expect(word?.x).toBeLessThan(78);
    // …and the dot itself stands out in front of it, at the box's edge.
    const dot = runs.find((r) => r.text.includes('•'));
    expect(dot?.x).toBeLessThan(6);
  });

  // §21.1.2.4.2/.3 — a bullet is drawn at the size of the TEXT IT LEADS,
  // scaled by `a:buSzPct`. Drawn at the level's inherited size instead,
  // ArtisticEffectSample's dots came out 48pt in front of 18pt text and every
  // line stood three times too tall.
  it('sizes the bullet by the run it leads, not by the level default', () => {
    const body =
      `<p:txStyles><p:titleStyle/>` +
      `<p:bodyStyle><a:lvl1pPr><a:buChar char="•"/><a:defRPr sz="4800"/></a:lvl1pPr>` +
      `</p:bodyStyle><p:otherStyle/></p:txStyles>`;
    const slide = (pPr: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p>${pPr}<a:r><a:rPr lang="en" sz="1800"/><a:t>Small</a:t></a:r>` +
      `</a:p></p:txBody></p:sp>`;
    const marker = (pPr: string) => {
      const doc = Ream.parse(buildPptx([slide(pPr)], { layoutMaster: { txStyles: body } }));
      return doc.flow.body.flatMap((el) =>
        el.kind === 'shape' && el.shape.text
          ? el.shape.text.content.flatMap((c) =>
              c.kind === 'paragraph' ? c.paragraph.runs.filter((r) => r.listMarker) : [],
            )
          : [],
      )[0];
    };
    expect(marker('')?.properties.fontSizePt).toBe(18); // the run's, not 48
    expect(marker('<a:pPr><a:buSzPct val="50000"/></a:pPr>')?.properties.fontSizePt).toBe(9);
    expect(marker('<a:pPr><a:buSzPts val="1200"/></a:pPr>')?.properties.fontSizePt).toBe(12);
    expect(
      marker('<a:pPr><a:buClr><a:srgbClr val="FF0000"/></a:buClr></a:pPr>')?.properties.colorHex,
    ).toBe('FF0000');
  });

  // §20.1.4.2.14 — a gallery style's `a:fontRef` colour belongs UNDER the run's
  // own and OVER the deck's default. Stamped on afterwards it lost to the `tx1`
  // every run inherits from `p:defaultTextStyle`: themes.pptx's green box asks
  // for `lt1` and drew its white caption black.
  it("writes a styled shape's text in the colour its style names", () => {
    const shape =
      `<p:sp><p:nvSpPr><p:cNvPr id="5" name="Rectangle 4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2971800" cy="1752600"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
      `<a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>white</a:t></a:r></a:p></p:txBody>` +
      `</p:sp>`;
    const doc = Ream.parse(
      buildPptx([shape], {
        defaultTextStyle:
          `<a:lvl1pPr><a:defRPr sz="1800"><a:solidFill>` +
          `<a:srgbClr val="000000"/></a:solidFill></a:defRPr></a:lvl1pPr>`,
      }),
    );
    expect(firstShapeRun(doc)?.properties.colorHex).toBe('FFFFFF');
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

  // §20.1.8.4 `a:alphaModFix` — how opaque the picture is DRAWN. A layout that
  // lays a photograph behind its title sets it low (ArtisticEffectSample's
  // cover shows one at 52%), and drawn full-strength the words on it are lost.
  it('draws a picture at the transparency its blip fixes', () => {
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="5" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId7"><a:alphaModFix amt="52000"/></a:blip>` +
      `<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const doc = Ream.parse(
      buildPptx([pic], {
        media: { 'ppt/media/image1.png': buildTinyPng(2, 2, [255, 0, 0, 255]) },
        slideRels: [`<Relationship Id="rId7" Type="${IMAGE_REL}" Target="../media/image1.png"/>`],
      }),
    );
    const el = doc.flow.body.find((e) => e.kind === 'image');
    expect(el?.kind === 'image' && el.image.alpha).toBeCloseTo(0.52, 5);
  });

  it('embeds the slide image into the rendered PDF', async () => {
    const pdf = await Ream.parse(picDeck()).convert('pdf', { fonts: FONTS });
    expect(PdfFile.parse(pdf).pages().length).toBe(1);
    // An image XObject made it into the PDF.
    expect(latin1.decode(pdf)).toContain('/Image');
  });
});

describe('what a slide shape states about its own box', () => {
  // §20.1.7.6 — a shape may be turned in its box, and a slide says so on the
  // same `a:xfrm` its position comes from. Unread, the blue triangle
  // ArtisticEffectSample's layout stands on its side pointed up, not right.
  it('turns and mirrors a shape the way its a:xfrm says', () => {
    const sp = (attrs: string): string =>
      `<p:sp><p:spPr><a:xfrm ${attrs}><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="1F7ABF"/></a:solidFill></p:spPr></p:sp>`;
    const shape = (attrs: string) => {
      const el = Ream.parse(buildPptx([sp(attrs)])).flow.body.find((e) => e.kind === 'shape');
      return el?.kind === 'shape' ? el.shape : undefined;
    };
    expect(shape('rot="5400000"')?.transform?.rotation60k).toBe(5400000);
    expect(shape('flipH="1"')?.transform?.flipH).toBe(true);
    expect(shape('')?.transform).toBeUndefined();
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

  // §20.1.8.14 — a chart may be papered with a PICTURE rather than filled with
  // a colour, and the blip is named through the CHART part's own relationships.
  // chart-texture-bg.pptx tiles a woven cloth over its whole frame, and a chart
  // on white is a different chart.
  it('papers a chart with the picture its frame is filled with', async () => {
    const textured = BAR_CHART.replace(
      '</c:chart>',
      '</c:chart><c:spPr><a:blipFill><a:blip r:embed="rIdTex"/>' +
        '<a:tile tx="0" ty="0" sx="100000" sy="100000"/></a:blipFill></c:spPr>',
    ).replace(
      '<c:chartSpace ',
      '<c:chartSpace xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
    );
    const gf =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Chart 5"/>` +
      `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="3200400"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="${CHART_NS}">` +
      `<c:chart xmlns:c="${CHART_NS}" r:id="rId8"/></a:graphicData></a:graphic></p:graphicFrame>`;
    const pptx = buildPptx([gf], {
      media: {
        'ppt/charts/chart1.xml': new TextEncoder().encode(textured),
        'ppt/charts/_rels/chart1.xml.rels': new TextEncoder()
          .encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdTex" Type="${IMAGE_REL}" Target="../media/tex.png"/></Relationships>`),
        'ppt/media/tex.png': buildTinyPng(2, 2, [200, 170, 100, 255]),
      },
      slideRels: [`<Relationship Id="rId8" Type="${CHART_REL}" Target="../charts/chart1.xml"/>`],
    });
    const doc = Ream.parse(pptx);
    const chart = [...(doc.flow.charts?.values() ?? [])][0];
    expect(chart?.frameFillImage?.tiled).toBe(true);
    expect(chart?.frameFillImage?.resource).toBeDefined();
    // …and the picture reaches the page: an XObject is drawn for it.
    const pdf = await Ream.parse(pptx).convert('pdf', { fonts: FONTS });
    expect(latin1.decode(pdf)).toContain('/Image');
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

describe('a slide table wears its style and stands in its frame', () => {
  const TABLE_STYLES_REL =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles';
  const GUID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
  const styles =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<a:tblStyleLst xmlns:a="${A_MAIN}" def="${GUID}"><a:tblStyle styleId="${GUID}" styleName="s">` +
    `<a:wholeTbl><a:tcStyle><a:tcBdr><a:insideH><a:ln w="12700">` +
    `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></a:insideH></a:tcBdr>` +
    `<a:fill><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>` +
    `<a:band1H><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="BBCCEE"/></a:solidFill></a:fill>` +
    `</a:tcStyle></a:band1H>` +
    `<a:firstRow><a:tcTxStyle b="on"><a:srgbClr val="FFFFFF"/></a:tcTxStyle>` +
    `<a:tcStyle><a:fill><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:fill>` +
    `</a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>`;
  const tc = (text: string, own = ''): string =>
    `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>${text}</a:t></a:r></a:p>` +
    `</a:txBody><a:tcPr>${own}</a:tcPr></a:tc>`;
  const deck = (tblPr: string): ReturnType<typeof Ream.parse> =>
    Ream.parse(
      buildPptx(
        [
          `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="t"/><p:cNvGraphicFramePr/>` +
            `<p:nvPr/></p:nvGraphicFramePr>` +
            `<p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="5486400" cy="1828800"/></p:xfrm>` +
            `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
            `<a:tbl>${tblPr}<a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid>` +
            `<a:tr h="457200">${tc('H1')}${tc('H2')}</a:tr>` +
            `<a:tr h="457200">${tc('a')}${tc('b')}</a:tr>` +
            `<a:tr h="457200">${tc('c')}` +
            tc('own', '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>') +
            `</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
        ],
        {
          presentationRels: `<Relationship Id="rIdTs" Type="${TABLE_STYLES_REL}" Target="tableStyles.xml"/>`,
          media: { 'ppt/tableStyles.xml': new TextEncoder().encode(styles) },
        },
      ),
    );
  const table = (doc: ReturnType<typeof Ream.parse>) =>
    doc.flow.body.flatMap((el) => (el.kind === 'table' ? [el.table] : []))[0];

  it('stands at the frame its slide puts it in', () => {
    const t = table(deck('<a:tblPr/>'));
    expect(t?.properties.float?.posH?.offsetPt).toBe(72); // 914400 EMU
    expect(t?.properties.float?.posV?.offsetPt).toBe(144);
  });

  // §20.1.4.2.24 — a table wears the style it NAMES; the flags say which of its
  // conditional parts reach the cells.
  const named = (flags: string): string =>
    `<a:tblPr ${flags}><a:tableStyleId>${GUID}</a:tableStyleId></a:tblPr>`;

  it('takes the fills, the rules and the bold its style names', () => {
    const t = table(deck(named('firstRow="1" bandRow="1"')));
    const shading = (r: number, c: number): string | undefined =>
      t?.rows[r]?.cells[c]?.properties.shading?.colorHex;
    expect(shading(0, 0)).toBe('4472C4'); // the header row
    expect(shading(1, 0)).toBe('BBCCEE'); // …the first band under it
    expect(shading(2, 0)).toBe('DDDDDD'); // …and the whole-table fill on the next
    expect(shading(2, 1)).toBe('FF0000'); // a cell's OWN fill beats the style
    expect(t?.rows[0]?.cells[0]?.properties.borders?.insideH?.colorHex).toBe('FFFFFF');
    const run = t?.rows[0]?.cells[0]?.content.flatMap((b) =>
      b.kind === 'paragraph' ? b.paragraph.runs : [],
    )[0];
    expect(run?.properties.bold).toBe(true);
  });

  // §20.1.4.2.24 — a table may name one of the GALLERY's own styles, and then
  // the deck ships no definition at all: `tableStyles.xml` lists what a deck
  // CUSTOMISES. Eight decks of the corpus do that, and predefined-table-style
  // came out as bare text with no fill and no rule anywhere.
  it('knows the styles PowerPoint has built in, which no deck ships', () => {
    const MEDIUM2_ACCENT1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
    const doc = Ream.parse(
      buildPptx([
        `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="t"/><p:cNvGraphicFramePr/>` +
          `<p:nvPr/></p:nvGraphicFramePr>` +
          `<p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="5486400" cy="1828800"/></p:xfrm>` +
          `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
          `<a:tbl><a:tblPr firstRow="1" bandRow="1">` +
          `<a:tableStyleId>${MEDIUM2_ACCENT1}</a:tableStyleId></a:tblPr>` +
          `<a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid>` +
          `<a:tr h="457200">${tc('H1')}${tc('H2')}</a:tr>` +
          `<a:tr h="457200">${tc('a')}${tc('b')}</a:tr>` +
          `<a:tr h="457200">${tc('c')}${tc('d')}</a:tr>` +
          `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
      ]),
    );
    const t = table(doc);
    // The definition is theme-RELATIVE, so the fills are the deck's accent and
    // its light shade, not colours baked into the style.
    expect(t?.rows[0]?.cells[0]?.properties.shading?.colorHex).toBeDefined();
    expect(t?.rows[1]?.cells[0]?.properties.shading?.colorHex).toBeDefined();
    expect(t?.rows[0]?.cells[0]?.properties.shading?.colorHex).not.toBe(
      t?.rows[1]?.cells[0]?.properties.shading?.colorHex,
    );
    // …and a GUID that is NOT one of them still wears nothing.
    const none = Ream.parse(
      buildPptx([
        `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="t"/><p:cNvGraphicFramePr/>` +
          `<p:nvPr/></p:nvGraphicFramePr>` +
          `<p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="5486400" cy="1828800"/></p:xfrm>` +
          `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
          `<a:tbl><a:tblPr firstRow="1"><a:tableStyleId>{00000000-0000-0000-0000-000000000000}` +
          `</a:tableStyleId></a:tblPr>` +
          `<a:tblGrid><a:gridCol w="2743200"/></a:tblGrid>` +
          `<a:tr h="457200">${tc('H1')}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
      ]),
    );
    expect(table(none)?.rows[0]?.cells[0]?.properties.shading).toBeUndefined();
  });

  // §20.1.4.2.25 / §20.1.4.2.19 — a style may POINT at the theme instead of
  // spelling a fill or a rule out: `a:tblBg/a:fillRef` is the table's own
  // background and `a:lnRef` is a rule whose width lives in the theme's line
  // style list. bnc480256's every rule and every second row are written that
  // way, and read as nothing the table came out with three pale bars and no
  // borders at all.
  it('resolves a background and a rule its style points at in the theme', () => {
    const themed =
      `<?xml version="1.0"?>\n<a:tblStyleLst xmlns:a="${A_MAIN}" def="${GUID}">` +
      `<a:tblStyle styleId="${GUID}" styleName="s">` +
      `<a:tblBg><a:fillRef idx="2"><a:srgbClr val="336699"/></a:fillRef></a:tblBg>` +
      `<a:wholeTbl><a:tcStyle><a:tcBdr><a:insideH>` +
      `<a:lnRef idx="2"><a:srgbClr val="FFCC00"/></a:lnRef>` +
      `</a:insideH></a:tcBdr></a:tcStyle></a:wholeTbl>` +
      `<a:band1H><a:tcStyle><a:fill><a:solidFill><a:srgbClr val="BBCCEE"/></a:solidFill>` +
      `</a:fill></a:tcStyle></a:band1H></a:tblStyle></a:tblStyleLst>`;
    const doc = Ream.parse(
      buildPptx(
        [
          `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="t"/><p:cNvGraphicFramePr/>` +
            `<p:nvPr/></p:nvGraphicFramePr>` +
            `<p:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="1828800"/></p:xfrm>` +
            `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
            `<a:tbl>${named('bandRow="1"')}<a:tblGrid><a:gridCol w="2743200"/></a:tblGrid>` +
            `<a:tr h="457200">${tc('a')}</a:tr><a:tr h="457200">${tc('b')}</a:tr>` +
            `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
        ],
        {
          presentationRels: `<Relationship Id="rIdTs" Type="${TABLE_STYLES_REL}" Target="tableStyles.xml"/>`,
          media: { 'ppt/tableStyles.xml': new TextEncoder().encode(themed) },
        },
      ),
    );
    const t = table(doc);
    // The band paints over the table's background; the row without one shows it.
    expect(t?.rows[0]?.cells[0]?.properties.shading?.colorHex).toBe('BBCCEE');
    expect(t?.rows[1]?.cells[0]?.properties.shading?.colorHex).toBe('336699');
    // The rule takes its colour from the reference and its width from the theme.
    expect(t?.rows[0]?.cells[0]?.properties.borders?.insideH?.colorHex).toBe('FFCC00');
  });

  it('leaves the style alone when the table asks for no part of it', () => {
    const t = table(deck(named('firstRow="0" bandRow="0"')));
    expect(t?.rows[0]?.cells[0]?.properties.shading?.colorHex).toBe('DDDDDD');
  });

  // The list's `@def` is the style PowerPoint applies when a table is INSERTED,
  // and the table then records that GUID itself — it is not a fallback for one
  // that names nothing (table-with-no-theme is two bare rows, not blue banding).
  it('wears nothing when it names no style, whatever the list defaults to', () => {
    const t = table(deck('<a:tblPr firstRow="1" bandRow="1"/>'));
    expect(t?.rows[0]?.cells[0]?.properties.shading).toBeUndefined();
    expect(t?.rows[2]?.cells[1]?.properties.shading?.colorHex).toBe('FF0000');
  });

  // §21.1.3.17 — the cell's own word about its fill and its four rules.
  it('keeps a cell transparent when the cell itself says a:noFill', () => {
    const t = table(
      Ream.parse(
        buildPptx(
          [
            `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="t"/>` +
              `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
              `<p:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></p:xfrm>` +
              `<a:graphic><a:graphicData uri="${TABLE_NS}"><a:tbl>${named('firstRow="1"')}` +
              `<a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid><a:tr h="457200">` +
              tc(
                'bare',
                '<a:lnL w="38100"><a:solidFill><a:srgbClr val="2670C9"><a:alpha val="0"/>' +
                  '</a:srgbClr></a:solidFill></a:lnL><a:lnB w="38100"><a:solidFill>' +
                  '<a:srgbClr val="2670C9"/></a:solidFill></a:lnB><a:noFill/>',
              ) +
              tc('styled') +
              `</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`,
          ],
          {
            presentationRels: `<Relationship Id="rIdTs" Type="${TABLE_STYLES_REL}" Target="tableStyles.xml"/>`,
            media: { 'ppt/tableStyles.xml': new TextEncoder().encode(styles) },
          },
        ),
      ),
    );
    const bare = t?.rows[0]?.cells[0];
    expect(bare?.properties.shading).toBeUndefined(); // the style's header fill is off
    expect(bare?.properties.borders?.left?.style).toBe('none'); // alpha 0 draws nothing
    expect(bare?.properties.borders?.bottom?.colorHex).toBe('2670C9');
    expect(bare?.properties.borders?.bottom?.width).toBe(3); // 38100 EMU
    expect(t?.rows[0]?.cells[1]?.properties.shading?.colorHex).toBe('4472C4');
  });

  it('keeps each row as tall as it asks (a:tr@h is a minimum)', () => {
    const t = table(deck('<a:tblPr/>'));
    expect(t?.rows[0]?.properties.height).toBe(36); // 457200 EMU
    expect(t?.rows[0]?.properties.heightRule).toBe('atLeast');
  });
});

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

const THEME = // dk2 is the blue this deck calls its background
  `<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFF00"/></a:lt1>` +
  `<a:dk2><a:srgbClr val="0066CC"/></a:dk2><a:lt2><a:srgbClr val="273943"/></a:lt2>`;
const SCHEME_BG =
  // a background painted with the bg1 ALIAS, not a slot
  `<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg>`;

describe("pptx colour map — what a deck's bg1 and tx1 mean (§19.3.1.6)", () => {
  it("reads bg1 through the master's map, not the fixed DrawingML alias", () => {
    // Without the map, `bg1` means lt1 — here yellow, and a deck that is blue
    // in PowerPoint came out yellow across every slide.
    const mapped = buildPptx([''], {
      layoutMaster: { theme: THEME, masterBg: SCHEME_BG, clrMap: 'bg1="dk2" tx1="lt1"' },
    });
    expect(firstShape(Ream.parse(mapped))?.fill.colorHex).toBe('0066CC');
    const unmapped = buildPptx([''], { layoutMaster: { theme: THEME, masterBg: SCHEME_BG } });
    expect(firstShape(Ream.parse(unmapped))?.fill.colorHex).toBe('FFFF00');
  });

  it('lets a layout override the map for the slides on it (§19.3.1.7)', () => {
    const doc = Ream.parse(
      buildPptx([''], {
        layoutMaster: {
          theme: THEME,
          masterBg: SCHEME_BG,
          clrMap: 'bg1="dk2"',
          layoutClrMapOvr: 'bg1="lt2"',
        },
      }),
    );
    expect(firstShape(doc)?.fill.colorHex).toBe('273943'); // the layout's word, not the master's
  });

  it("lets a slide override the map for its OWN content, not for the master's", () => {
    // §19.3.1.7 — the override governs the slide. What the master draws keeps
    // reading under the master's map: chart_pt_color_bg1 flips bg1 to dk1 for
    // its chart and its white deck came out black when the flip reached the
    // master's background too.
    const shape =
      `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:spPr></p:sp>`;
    const doc = Ream.parse(
      buildPptx([shape, shape], {
        layoutMaster: { theme: THEME, masterBg: SCHEME_BG, clrMap: 'bg1="dk2"' },
        slideClrMapOvr: [undefined, 'bg1="lt1"'],
      }),
    );
    const fills = doc.flow.body.flatMap((el) =>
      el.kind === 'shape' ? [el.shape.fill.colorHex] : [],
    );
    // Per slide: the master's backdrop (blue either way) then the slide's own
    // shape — blue on the first, yellow on the one that overrides.
    expect(fills).toEqual(['0066CC', '0066CC', '0066CC', 'FFFF00']);
  });

  it('paints a real deck the colour its map names', () => {
    // corpus: the master's bg is `schemeClr bg1` and its map says bg1 = dk2.
    const deck = new Uint8Array(readFileSync('tests/fixtures/real/master-bg-color.pptx'));
    expect(firstShape(Ream.parse(deck))?.fill.colorHex).toBe('009DF0');
  });
});

describe('pptx background — a reference and a picture (E-PPTX PX5b)', () => {
  const BG_STYLES =
    `<a:fillStyleLst><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:fillStyleLst>` +
    `<a:bgFillStyleLst>` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>` +
    `</a:bgFillStyleLst>`;
  const bgRef = (idx: number, inner: string): string =>
    `<p:bg><p:bgRef idx="${String(idx)}">${inner}</p:bgRef></p:bg>`;

  it('reads a p:bgRef as the theme fill it indexes, with phClr bound', () => {
    // §19.3.1.2 — 1001 is the first background style, and the colour on the
    // reference is what the style means by `phClr`.
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [bgRef(1001, '<a:srgbClr val="AABBCC"/>')],
        layoutMaster: { themeFmt: BG_STYLES },
      }),
    );
    expect(firstShape(doc)?.fill.colorHex).toBe('AABBCC');
  });

  it('takes the slot itself when the slot names its own colour', () => {
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [bgRef(1002, '<a:srgbClr val="AABBCC"/>')],
        layoutMaster: { themeFmt: BG_STYLES },
      }),
    );
    expect(firstShape(doc)?.fill.colorHex).toBe('00FF00'); // the style, not the reference
  });

  it('indexes the ordinary fill styles below 1000', () => {
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [bgRef(1, '<a:srgbClr val="AABBCC"/>')],
        layoutMaster: { themeFmt: BG_STYLES },
      }),
    );
    expect(firstShape(doc)?.fill.colorHex).toBe('112233');
  });

  it('falls back to the reference colour when the theme has no such slot', () => {
    const doc = Ream.parse(
      buildPptx([''], { slideBg: [bgRef(1001, '<a:srgbClr val="AABBCC"/>')] }),
    );
    expect(firstShape(doc)?.fill.colorHex).toBe('AABBCC');
  });

  it('paints a picture background, at the opacity the blip fixes', () => {
    // §20.1.8.4 `a:alphaModFix` — tdf146223 backs a slide with a photo at 70 %.
    const png = buildTinyPng(2, 2, [255, 0, 0, 255]);
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [
          `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg">` +
            `<a:alphaModFix amt="70000"/></a:blip><a:tile/></a:blipFill></p:bgPr></p:bg>`,
        ],
        slideRels: [`<Relationship Id="rIdBg" Type="${IMAGE_REL}" Target="../media/bg.png"/>`],
        media: { 'ppt/media/bg.png': png },
      }),
    );
    const fill = firstShape(doc)?.fill;
    expect(fill?.kind).toBe('picture');
    expect(fill?.imageResource).toBeDefined();
    expect(fill?.tiled).toBe(true);
    expect(fill?.alpha).toBeCloseTo(0.7, 5);
  });

  it('recolours a picture between the two tones a duotone names', () => {
    // §20.1.8.23 — an Office theme ships a grey photograph and tints it; the
    // deck whose background is a brown ridged texture stores a grey one
    // (corpus: themes.pptx).
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [
          `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg">` +
            `<a:duotone><a:srgbClr val="1A0F00"/><a:srgbClr val="E8C9A0"/></a:duotone>` +
            `</a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>`,
        ],
        slideRels: [`<Relationship Id="rIdBg" Type="${IMAGE_REL}" Target="../media/bg.png"/>`],
        media: { 'ppt/media/bg.png': buildTinyPng(2, 2, [128, 128, 128, 255]) },
      }),
    );
    expect(firstShape(doc)?.fill.duotone).toEqual({
      shadowHex: '1A0F00',
      highlightHex: 'E8C9A0',
    });
  });

  it('takes a run colour from the gradient it is filled with', () => {
    // §20.1.8.33 — WordArt is always filled with a gradient, never a solid.
    // A run here wears one colour, so it takes the stop nearest the middle —
    // the one most of the glyph is painted in. Unread, the run had no colour
    // at all and every word came out in the default black (tdf114848 page 2).
    const run =
      `<a:r><a:rPr lang="en-US" sz="5400" b="1"><a:gradFill><a:gsLst>` +
      `<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>` +
      `<a:gs pos="50000"><a:srgbClr val="1F4E79"/></a:gs>` +
      `<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>` +
      `</a:gsLst><a:lin ang="5400000"/></a:gradFill></a:rPr>` +
      `<a:t>Wave</a:t></a:r>`;
    const sp =
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="W"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
      `<p:txBody><a:bodyPr><a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp></a:bodyPr>` +
      `<a:lstStyle/><a:p>${run}</a:p></p:txBody></p:sp>`;
    const doc = Ream.parse(buildPptx([sp]));
    const shape = doc.flow.body.find((e) => e.kind === 'shape');
    if (shape?.kind !== 'shape') throw new Error('no shape');
    const block = shape.shape.text?.content[0];
    const only = block?.kind === 'paragraph' ? block.paragraph.runs[0] : undefined;
    expect(only?.text).toBe('Wave');
    // The middle stop, not the default black a run with no fill falls back to.
    expect(only?.properties.colorHex).toBe('1F4E79');
  });

  it('clips a picture to the geometry it names for itself', () => {
    // §19.3.1.37 — a `p:pic` may carry a geometry of its own, and then the
    // picture is CLIPPED to it: crop-to-shape.pptx is one photograph in an
    // ellipse, and drawn as a plain rectangle the ellipse came out square.
    const pic = (prst: string): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture 3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdC"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/></a:xfrm>` +
      `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const build = (prst: string) =>
      Ream.parse(
        buildPptx([pic(prst)], {
          slideRels: [`<Relationship Id="rIdC" Type="${IMAGE_REL}" Target="../media/c.png"/>`],
          media: { 'ppt/media/c.png': buildTinyPng(2, 2, [0, 200, 0, 255]) },
        }),
      );
    // An ellipse is a SHAPE wearing the picture as its fill — that is what
    // clips it, and the shape path already draws one.
    const round = build('ellipse').flow.body.find((e) => e.kind === 'shape');
    expect(round?.kind).toBe('shape');
    if (round?.kind !== 'shape') throw new Error('not a shape');
    expect(round.shape.geometry).toMatchObject({ kind: 'preset', preset: 'ellipse' });
    expect(round.shape.fill.kind).toBe('picture');
    // A plain rectangle needs no clipping and stays the picture it was: the
    // image path measures and embeds one far more directly.
    expect(build('rect').flow.body.some((e) => e.kind === 'image')).toBe(true);
    expect(build('rect').flow.body.some((e) => e.kind === 'shape')).toBe(false);
  });

  it('places a PICTURE placeholder from the layout when it carries no transform', () => {
    // §19.3.1.35 — a content placeholder filled with a photograph is a `p:pic`
    // whose `p:spPr` is EMPTY: its whole geometry is the layout's. Read as a
    // picture with no box of its own it drew nothing at all
    // (customshape-bitmapfill-srcrect.pptx is one such picture and no more).
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="6" name="Content Placeholder"/><p:cNvPicPr/>` +
      `<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdP"/><a:stretch/></p:blipFill>` +
      `<p:spPr/></p:pic>`;
    const layout =
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="6192000" y="1332000"/><a:ext cx="5493600" cy="4012789"/>` +
      `</a:xfrm></p:spPr></p:sp>`;
    const doc = Ream.parse(
      buildPptx([pic], {
        layoutMaster: { layoutSpTree: layout },
        slideRels: [`<Relationship Id="rIdP" Type="${IMAGE_REL}" Target="../media/p.png"/>`],
        media: { 'ppt/media/p.png': buildTinyPng(2, 2, [255, 0, 0, 255]) },
      }),
    );
    const pics = doc.flow.body.filter((e) => e.kind === 'image');
    expect(pics).toHaveLength(1);
    const only = pics[0];
    if (only?.kind !== 'image') throw new Error('no picture');
    // The layout's own box, to the point.
    expect(only.image.width).toBeCloseTo(5493600 / 12700, 4);
    expect(only.image.float?.posH?.offsetPt).toBeCloseTo(6192000 / 12700, 4);
  });

  it('reads a grayscale blip as the duotone from black to white that it is', () => {
    // §20.1.8.34 `a:grayscl` — the picture drawn in shades of grey, which is
    // what a duotone between black and white already is. tdf112209's chevron
    // is a colour photograph in the file and grey in every reader.
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [
          `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg"><a:grayscl/></a:blip>` +
            `<a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>`,
        ],
        slideRels: [`<Relationship Id="rIdBg" Type="${IMAGE_REL}" Target="../media/bg.png"/>`],
        media: { 'ppt/media/bg.png': buildTinyPng(2, 2, [20, 90, 200, 255]) },
      }),
    );
    expect(firstShape(doc)?.fill.duotone).toEqual({
      shadowHex: '000000',
      highlightHex: 'FFFFFF',
    });
  });

  it('paints a duotone through the picture, as a luminosity mask', async () => {
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [
          `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg">` +
            `<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FF0000"/></a:duotone>` +
            `</a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>`,
        ],
        slideRels: [`<Relationship Id="rIdBg" Type="${IMAGE_REL}" Target="../media/bg.png"/>`],
        media: { 'ppt/media/bg.png': buildTinyPng(2, 2, [200, 200, 200, 255]) },
      }),
    );
    const pdf = await doc.convert('pdf', { fonts: FONTS });
    const bytes = latin1.decode(pdf);
    // The picture is not painted at all: it masks the light colour over the
    // dark one (ISO 32000-1 §11.6.5.2), which is the two-tone map itself.
    expect(bytes).toContain('/Luminosity');
    expect(bytes).toContain('/SMask');
  });

  it('stretches a background picture into the fill rect, not across the slide', () => {
    // §20.1.8.30 — POSITIVE insets say where the picture goes IN the box.
    // tdf153466 insets one 55 % from the left and 56 % from the top; drawn
    // over the whole slide it is a triangle five times its size.
    const doc = Ream.parse(
      buildPptx([''], {
        slideBg: [
          `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg"/>` +
            `<a:stretch><a:fillRect l="55000" t="56000"/></a:stretch></a:blipFill></p:bgPr></p:bg>`,
        ],
        slideRels: [`<Relationship Id="rIdBg" Type="${IMAGE_REL}" Target="../media/bg.png"/>`],
        media: { 'ppt/media/bg.png': buildTinyPng(2, 2, [255, 0, 0, 255]) },
      }),
    );
    expect(firstShape(doc)?.fill.imageFillRect).toEqual({
      left: 0.55,
      top: 0.56,
      right: 0,
      bottom: 0,
    });
  });

  it("resolves the MASTER's background picture through the master's own rels", () => {
    // The blip's relationship id is scoped to the part the background is
    // written in, so a master background cannot be resolved against the slide.
    const doc = Ream.parse(
      buildPptx([''], {
        layoutMaster: {
          masterBg: `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdM"/></a:blipFill></p:bgPr></p:bg>`,
          masterRels: `<Relationship Id="rIdM" Type="${IMAGE_REL}" Target="../media/m.png"/>`,
        },
        media: { 'ppt/media/m.png': buildTinyPng(2, 2, [0, 0, 255, 255]) },
      }),
    );
    expect(firstShape(doc)?.fill.kind).toBe('picture');
  });
});

describe('a picture that declares a colour away (§20.1.8.16)', () => {
  it('reads a:clrChange off the blip, and whether it knocks the colour out', () => {
    const pic = (change: string): string =>
      `<p:pic><p:nvPicPr><p:cNvPr id="5" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdImg">${change}</a:blip><a:stretch/></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const deck = (change: string): ReturnType<typeof Ream.parse> =>
      Ream.parse(
        buildPptx([pic(change)], {
          slideRels: [`<Relationship Id="rIdImg" Type="${IMAGE_REL}" Target="../media/i.png"/>`],
          media: { 'ppt/media/i.png': buildTinyPng(2, 2, [255, 255, 255, 255]) },
        }),
      );
    const image = (doc: ReturnType<typeof Ream.parse>) =>
      doc.flow.body.flatMap((e) => (e.kind === 'image' ? [e.image] : []))[0];
    // Zero alpha on the destination: the colour goes.
    const gone = deck(
      `<a:clrChange><a:clrFrom><a:srgbClr val="FFFFFF"/></a:clrFrom>` +
        `<a:clrTo><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>`,
    );
    expect(image(gone)?.colorChange).toEqual({
      fromHex: 'FFFFFF',
      toHex: 'FFFFFF',
      transparent: true,
    });
    // A destination with no alpha of its own repaints instead.
    const swapped = deck(
      `<a:clrChange><a:clrFrom><a:srgbClr val="FFFFFF"/></a:clrFrom>` +
        `<a:clrTo><a:srgbClr val="112233"/></a:clrTo></a:clrChange>`,
    );
    expect(image(swapped)?.colorChange).toEqual({
      fromHex: 'FFFFFF',
      toHex: '112233',
      transparent: false,
    });
    expect(image(deck(''))?.colorChange).toBeUndefined();
  });
});

describe('what a slide puts behind its content', () => {
  it("paints the backdrop before the slide's own picture, not over it", () => {
    // Every shape paints after every image in the ordinary passes, so a white
    // backdrop landed on top of the photograph the slide is made of
    // (corpus: tdf156808, tdf157635, tdf156856 — three dark slides drawn
    // blank). What the page puts behind its content paints first instead.
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="4" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const doc = Ream.parse(
      buildPptx([pic], {
        layoutMaster: { masterBg: bgFill('FFFFFF') },
        slideRels: [`<Relationship Id="rIdImg" Type="${IMAGE_REL}" Target="../media/i.png"/>`],
        media: { 'ppt/media/i.png': buildTinyPng(2, 2, [0, 0, 0, 255]) },
      }),
    );
    const laid = layoutStyledDocument(doc.flow.body, {
      registry: FontRegistry.fromBytes({ regular: FONTS.regular }),
      resources: doc.flow.resources,
      ...(doc.flow.section ? { section: doc.flow.section } : {}),
      styles: doc.flow.styles,
    });
    const plan = paintPlan(laid.pages[0]!.commands);
    expect(plan.behind.map((c) => c.type)).toEqual(['shape']); // the backdrop
    expect(plan.images).toHaveLength(1); // …and the picture over it
    expect(plan.shapes).toHaveLength(0);
  });
});

describe("pptx inherited shapes — the deck's own decoration (E-PPTX PX5d)", () => {
  const rect = (hex: string): string =>
    `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></p:spPr></p:sp>`;
  // A placeholder on a master is a prototype, not a drawn shape.
  const phRect = (hex: string): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="9" name="ph"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>` +
    `</p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></p:spPr></p:sp>`;
  const fills = (doc: ReturnType<typeof Ream.parse>): Array<string | undefined> =>
    doc.flow.body.flatMap((el) => (el.kind === 'shape' ? [el.shape.fill.colorHex] : []));

  it("draws the master's shapes, then the layout's, under the slide's own", () => {
    const doc = Ream.parse(
      buildPptx([rect('333333')], {
        layoutMaster: { masterSpTree: rect('111111'), layoutSpTree: rect('222222') },
      }),
    );
    expect(fills(doc)).toEqual(['111111', '222222', '333333']);
  });

  it('leaves the placeholders behind — they are prototypes, not decoration', () => {
    const doc = Ream.parse(
      buildPptx([''], {
        layoutMaster: {
          masterSpTree: phRect('111111') + rect('AAAAAA'),
          layoutSpTree: phRect('222222'),
        },
      }),
    );
    expect(fills(doc)).toEqual(['AAAAAA']);
  });

  it('shows none of them on a slide that says showMasterSp="0" (§19.3.1.38)', () => {
    const doc = Ream.parse(
      buildPptx([rect('333333')], {
        layoutMaster: { masterSpTree: rect('111111'), layoutSpTree: rect('222222') },
        hideMasterShapes: [true],
      }),
    );
    expect(fills(doc)).toEqual(['333333']);
  });

  it("…and only the master's on a LAYOUT that says so (§19.3.1.39)", () => {
    const doc = Ream.parse(
      buildPptx([rect('333333')], {
        layoutMaster: {
          masterSpTree: rect('111111'),
          layoutSpTree: rect('222222'),
          hideMasterShapes: true,
        },
      }),
    );
    expect(fills(doc)).toEqual(['222222', '333333']);
  });

  it('paints a useBgFill shape with the slide background it stands on', () => {
    // §19.3.1.43 — tdf93868's master lays a white rectangle over the whole
    // slide, then a rounded one marked `useBgFill` that lets the background
    // back through. Read without it the deck is a blank white page.
    const bgFilled =
      `<p:sp useBgFill="1"><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr></p:sp>`;
    const doc = Ream.parse(
      buildPptx([''], {
        layoutMaster: { masterBg: bgFill('102030'), masterSpTree: rect('FFFFFF') + bgFilled },
      }),
    );
    // The backdrop, the white rectangle over it, and the shape that is the
    // background again — not the white its own fill states.
    expect(fills(doc)).toEqual(['102030', 'FFFFFF', '102030']);
  });

  // §19.3.1.43 + §20.1.8.30 — a PICTURE background is stretched over the whole
  // slide, so a shape that wears it wears the piece under itself. Squeezed into
  // the box instead, tdf123684's text box drew a little copy of the slide's own
  // diagonal inside itself.
  it('gives a useBgFill shape the piece of a picture background under it', () => {
    const bgFilled =
      `<p:sp useBgFill="1"><p:spPr>` +
      `<a:xfrm><a:off x="3048000" y="1714500"/><a:ext cx="3048000" cy="1714500"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
    const doc = Ream.parse(
      buildPptx([bgFilled], {
        cx: 12192000,
        cy: 6858000,
        slideBg: [
          `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg"/>` +
            `<a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>`,
        ],
        slideRels: [`<Relationship Id="rIdBg" Type="${IMAGE_REL}" Target="../media/bg.png"/>`],
        media: { 'ppt/media/bg.png': buildTinyPng(2, 2, [10, 20, 30, 255]) },
      }),
    );
    // The backdrop is the first shape; the second is the one wearing it.
    const fill = doc.flow.body.flatMap((el) => (el.kind === 'shape' ? [el.shape.fill] : []))[1];
    expect(fill?.kind).toBe('picture');
    // The box is a quarter of the slide, a quarter in: the picture's rect
    // reaches one box-width left and up, and one right and down.
    expect(fill?.imageFillRect?.left).toBeCloseTo(-1, 5);
    expect(fill?.imageFillRect?.top).toBeCloseTo(-1, 5);
    expect(fill?.imageFillRect?.right).toBeCloseTo(-2, 5);
    expect(fill?.imageFillRect?.bottom).toBeCloseTo(-2, 5);
  });

  it('draws them on every slide of the layout, and behind the background', () => {
    const doc = Ream.parse(
      buildPptx(['', ''], {
        layoutMaster: { masterSpTree: rect('111111'), masterBg: bgFill('445566') },
      }),
    );
    // Per slide: the backdrop first, then the inherited shape over it.
    expect(fills(doc)).toEqual(['445566', '111111', '445566', '111111']);
  });

  // The backdrop layer paints in its own order, which the emitter once wrote as
  // "the first token at a point" — enough for a metafile's label and nothing
  // like a sentence. WithMaster.pptx printed "This" for its master's caption.
  it('sets every word of the text one of them carries', async () => {
    const caption =
      `<p:sp><p:nvSpPr><p:cNvPr id="8" name="cap"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="457200"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en" sz="1800"/>` +
      `<a:t>This text comes from the Master Slide</a:t></a:r></a:p></p:txBody></p:sp>`;
    const pptx = buildPptx([''], { layoutMaster: { masterSpTree: caption } });
    const file = PdfFile.parse(await Ream.parse(pptx).convert('pdf', { fonts: FONTS }));
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join(' ')
      .replace(/\s+/gu, ' ');
    expect(text).toContain('This');
    expect(text).toContain('Master');
    expect(text).toContain('Slide');
  });

  // §19.3.1 — the page paints by KIND (every image, then every shape), so a
  // layout's card would land on top of the photograph a slide puts on it.
  it("sinks them BEHIND the slide's own content, whatever kind that is", () => {
    const doc = Ream.parse(
      buildPptx(
        [
          `<p:pic><p:nvPicPr><p:cNvPr id="5" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
            `<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch/></p:blipFill>` +
            `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
            `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
        ],
        {
          layoutMaster: { masterSpTree: rect('111111') },
          media: { 'ppt/media/i.png': buildTinyPng(2, 2, [255, 0, 0, 255]) },
          slideRels: [`<Relationship Id="rIdImg" Type="${IMAGE_REL}" Target="../media/i.png"/>`],
        },
      ),
    );
    const inheritedShape = doc.flow.body.find((el) => el.kind === 'shape');
    const slidePicture = doc.flow.body.find((el) => el.kind === 'image');
    expect(inheritedShape?.kind === 'shape' && inheritedShape.shape.float?.behind).toBe(true);
    expect(slidePicture?.kind === 'image' && slidePicture.image.float?.behind).toBeUndefined();
  });
});

describe('what a placeholder inherits from its prototype', () => {
  const protoRect =
    `<p:sp><p:nvSpPr><p:cNvPr id="9" name="body"/><p:cNvSpPr/><p:nvPr><p:ph idx="13"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="228600" y="800100"/><a:ext cx="4572000" cy="2286000"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="76BF3D"/></a:solidFill>` +
    `<a:ln w="19050"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`;
  const slidePh =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="b"/><p:cNvSpPr/><p:nvPr><p:ph idx="13"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>Test</a:t></a:r></a:p></p:txBody></p:sp>`;

  it('takes the fill, the outline and the geometry it does not state', () => {
    // tdf95932 — "Test inheritance of shape properties from slide master": the
    // green panel is the layout's, and the word on it is white, so without the
    // panel the slide read as blank paper.
    const doc = Ream.parse(buildPptx([slidePh], { layoutMaster: { layoutSpTree: protoRect } }));
    const shape = doc.flow.body.flatMap((e) =>
      e.kind === 'shape' && e.shape.text ? [e.shape] : [],
    )[0];
    expect(shape?.fill.colorHex).toBe('76BF3D');
    expect(shape?.geometry.kind === 'preset' ? shape.geometry.preset : '').toBe('roundRect');
    expect(shape?.line?.colorHex).toBe('112233');
  });

  // §19.3.1.36 — the inheritance is per PROPERTY. A layout prototype that
  // states a box and nothing else must not stop the search: tdf104015's title
  // is red in every reader, and the red is the MASTER's, three levels up from
  // a slide that states only a blue outline.
  it('keeps looking past a prototype that states only some of them', () => {
    const layoutBoxOnly =
      `<p:sp><p:nvSpPr><p:cNvPr id="9" name="body"/><p:cNvSpPr/><p:nvPr><p:ph idx="13"/></p:nvPr>` +
      `</p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="2286000"/></a:xfrm>` +
      `</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`;
    const doc = Ream.parse(
      buildPptx([slidePh], {
        layoutMaster: { layoutSpTree: layoutBoxOnly, masterSpTree: protoRect },
      }),
    );
    const shape = doc.flow.body.flatMap((e) =>
      e.kind === 'shape' && e.shape.text ? [e.shape] : [],
    )[0];
    // The box is the layout's…
    expect(shape?.width).toBe(360); // 4572000 EMU
    // …and the fill it never states comes from the master behind it.
    expect(shape?.fill.colorHex).toBe('76BF3D');
  });

  it('…but what the slide states itself still wins', () => {
    const own = slidePh.replace(
      '<p:spPr/>',
      `<p:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>` +
        `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`,
    );
    const doc = Ream.parse(buildPptx([own], { layoutMaster: { layoutSpTree: protoRect } }));
    const shape = doc.flow.body.flatMap((e) =>
      e.kind === 'shape' && e.shape.text ? [e.shape] : [],
    )[0];
    expect(shape?.fill.colorHex).toBe('FF0000');
    expect(shape?.geometry.kind === 'preset' ? shape.geometry.preset : '').toBe('ellipse');
  });
});

describe('a shape drawn from a gallery style (§20.1.4.2)', () => {
  it('takes its fill, outline and text colour from p:style', () => {
    // customGeo's title banner and the ellipse under it carry no fill in their
    // spPr at all — both are a theme slot named by `a:fillRef`.
    const themeFmt =
      `<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
      `<a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>`;
    const sp =
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="g"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:style><a:lnRef idx="1"><a:srgbClr val="223344"/></a:lnRef>` +
      `<a:fillRef idx="1"><a:srgbClr val="4488CC"/></a:fillRef>` +
      `<a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef>` +
      `<a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>` +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>`;
    const doc = Ream.parse(buildPptx([sp], { layoutMaster: { themeFmt } }));
    const shape = doc.flow.body.flatMap((e) => (e.kind === 'shape' ? [e.shape] : []))[0];
    expect(shape?.fill.colorHex).toBe('4488CC');
    expect(shape?.line?.colorHex).toBe('223344');
    const run = shape?.text?.content.flatMap((c) =>
      c.kind === 'paragraph' ? c.paragraph.runs : [],
    )[0];
    expect(run?.properties.colorHex).toBe('FFFFFF');
  });

  // §20.1.8.33 — a `a:gradFill` holding only a direction states no colours at
  // all: the run of them comes from the style's `a:fillRef` slot and the shape
  // says only which way to sweep it. Counted as a fill of its own, the shape
  // ended up with none — 63200.pptx's ellipse drew as its own SHADOW, a grey
  // disc where every reader has the theme's blue.
  it('takes the run of colours from the style when the shape states only a direction', () => {
    const themeFmt =
      `<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
      `<a:gradFill><a:gsLst>` +
      `<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:gs>` +
      `<a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs>` +
      `</a:gsLst><a:lin ang="5400000"/></a:gradFill></a:fillStyleLst>` +
      `<a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>`;
    const sp = (fill: string): string =>
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="g"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>${fill}</p:spPr>` +
      `<p:style><a:lnRef idx="1"><a:srgbClr val="223344"/></a:lnRef>` +
      `<a:fillRef idx="2"><a:srgbClr val="4488CC"/></a:fillRef>` +
      `<a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style></p:sp>`;
    const fillOf = (inner: string): ShapeFill | undefined =>
      Ream.parse(buildPptx([sp(inner)], { layoutMaster: { themeFmt } })).flow.body.flatMap((e) =>
        e.kind === 'shape' ? [e.shape.fill] : [],
      )[0];

    // The theme's second slot is a gradient of the reference's colour…
    const bare = fillOf('<a:gradFill><a:lin ang="2700000"/><a:tileRect/></a:gradFill>');
    expect(bare?.kind).toBe('gradient');
    expect(bare?.gradient?.stops.map((stop) => stop.colorHex)).toEqual(['B0C3E2', '4488CC']);
    // …swept the way the SHAPE asks, not the way the slot does (90°).
    expect(bare?.gradient?.kind === 'linear' ? bare.gradient.angle : undefined).toBe(45);

    // A shape that states nothing takes the slot's own direction.
    expect(
      (() => {
        const f = fillOf('');
        return f?.gradient?.kind === 'linear' ? f.gradient.angle : undefined;
      })(),
    ).toBe(90);
  });
});

describe('pptx inherited text — what a slide is written in (E-PPTX PX2)', () => {
  const textBox = (text: string): string =>
    `<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const titlePh = (text: string): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const shapeOf = (doc: ReturnType<typeof Ream.parse>) =>
    doc.flow.body.flatMap((el) => (el.kind === 'shape' && el.shape.text ? [el.shape] : []))[0];
  const firstParagraph = (doc: ReturnType<typeof Ream.parse>) => {
    const c = shapeOf(doc)?.text?.content[0];
    return c?.kind === 'paragraph' ? c.paragraph : undefined;
  };

  it("writes a plain text box in the deck's default text style (§19.2.1.8)", () => {
    // tdf93868's only shape is a text box with no colour of its own; the deck
    // says tx1, its map says tx1 is lt1, and its lt1 is white — so the slide
    // reads white on black, not black on black.
    const doc = Ream.parse(
      buildPptx([textBox('plain')], {
        defaultTextStyle:
          `<a:lvl1pPr algn="ctr"><a:defRPr sz="2800">` +
          `<a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr>`,
        layoutMaster: { theme: THEME, clrMap: 'tx1="lt1"' },
      }),
    );
    const p = firstParagraph(doc);
    expect(p?.runs[0]?.properties.colorHex).toBe('FFFF00'); // lt1 through the map
    expect(p?.runs[0]?.properties.fontSizePt).toBe(28);
    expect(p?.properties.alignment).toBe('center');
  });

  it("takes a placeholder's alignment from the master's text styles", () => {
    const doc = Ream.parse(
      buildPptx([titlePh('Title')], {
        layoutMaster: {
          txStyles: `<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle></p:txStyles>`,
        },
      }),
    );
    expect(firstParagraph(doc)?.properties.alignment).toBe('center');
    expect(firstParagraph(doc)?.runs[0]?.properties.fontSizePt).toBe(44);
  });

  it("lets the layout's own prototype override the master's family style", () => {
    const doc = Ream.parse(
      buildPptx([titlePh('Title')], {
        layoutMaster: {
          txStyles: `<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle></p:txStyles>`,
          layoutSpTree:
            `<p:sp><p:nvSpPr><p:cNvPr id="8" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
            `<p:spPr/><p:txBody><a:bodyPr anchor="b"/>` +
            `<a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="2000"/></a:lvl1pPr></a:lstStyle>` +
            `<a:p/></p:txBody></p:sp>`,
        },
      }),
    );
    expect(firstParagraph(doc)?.properties.alignment).toBe('right'); // the layout's word
    expect(firstParagraph(doc)?.runs[0]?.properties.fontSizePt).toBe(20);
    // …and the prototype's vertical anchor comes with it.
    expect(shapeOf(doc)?.text?.anchor).toBe('b');
  });

  it("keeps the paragraph's own properties over everything inherited", () => {
    const doc = Ream.parse(
      buildPptx([titlePh('Title').replace('<a:p>', '<a:p><a:pPr algn="l"/>')], {
        layoutMaster: {
          txStyles: `<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle></p:txStyles>`,
        },
      }),
    );
    expect(firstParagraph(doc)?.properties.alignment).toBe('left');
  });
});

describe('pptx embedded objects — the picture they show (E-PPTX PX4c)', () => {
  const OLE_URI = 'http://schemas.openxmlformats.org/presentationml/2006/ole';
  const VML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
  const frame = (inner: string): string =>
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Object 1"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${OLE_URI}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`;
  const enc = (t: string): Uint8Array => new TextEncoder().encode(t);
  const firstImage = (doc: ReturnType<typeof Ream.parse>) =>
    doc.flow.body.flatMap((el) => (el.kind === 'image' ? [el.image] : []))[0];

  it("draws the preview a legacy object keeps in the slide's VML drawing", () => {
    // 45541_Footer's eighth slide is one embedded deck and nothing else: the
    // `@spid` names a VML shape whose `v:imagedata` is the snapshot.
    const doc = Ream.parse(
      buildPptx(
        [frame('<p:oleObj spid="_x0000_s1026" name="Slide" r:id="rIdOle"><p:embed/></p:oleObj>')],
        {
          slideRels: [
            `<Relationship Id="rIdVml" Type="${VML_REL}" Target="../drawings/vmlDrawing1.vml"/>`,
          ],
          media: {
            'ppt/drawings/vmlDrawing1.vml': enc(
              `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">` +
                `<v:shape id="_x0000_s1026" type="#_x0000_t75" style="width:10in;height:540pt">` +
                `<v:imagedata o:relid="rId1" o:title=""/></v:shape></xml>`,
            ),
            'ppt/drawings/_rels/vmlDrawing1.vml.rels': enc(
              `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                `<Relationship Id="rId1" Type="${IMAGE_REL}" Target="../media/prev.png"/></Relationships>`,
            ),
            'ppt/media/prev.png': buildTinyPng(2, 2, [0, 128, 255, 255]),
          },
        },
      ),
    );
    const img = firstImage(doc);
    expect(img?.resource).toBeDefined();
    expect(Math.round(img?.width ?? 0)).toBe(720); // the frame, 10in wide
    expect(Math.round(img?.height ?? 0)).toBe(540);
  });

  it('draws the p:pic a modern object carries inside itself', () => {
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="8" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const doc = Ream.parse(
      buildPptx([frame(`<p:oleObj spid="_x0000_s1027" r:id="rIdOle">${pic}</p:oleObj>`)], {
        slideRels: [`<Relationship Id="rIdImg" Type="${IMAGE_REL}" Target="../media/p.png"/>`],
        media: { 'ppt/media/p.png': buildTinyPng(2, 2, [255, 0, 0, 255]) },
      }),
    );
    const img = firstImage(doc);
    expect(img?.resource).toBeDefined();
    expect(Math.round(img?.width ?? 0)).toBe(144); // the pic's own box, 2in
  });

  it('says so when an embedded object shows no picture at all', () => {
    const doc = Ream.parse(
      buildPptx([frame('<p:oleObj spid="_x0000_s1028" r:id="rIdOle"><p:embed/></p:oleObj>')]),
    );
    expect(firstImage(doc)).toBeUndefined();
    expect(doc.losses.some((l) => /embedded object/u.test(l.detail))).toBe(true);
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
  opts: {
    readonly fillA?: string;
    readonly build?: Parameters<typeof buildPptx>[1];
    /** Omit the data part's own .rels, as PowerPoint does. */
    readonly dropDataRels?: boolean;
  } = {},
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
  const parts: Record<string, Uint8Array> = {
    'ppt/diagrams/data1.xml': enc.encode(
      `<?xml version="1.0"?>\n<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>`,
    ),
    'ppt/diagrams/drawing1.xml': enc.encode(drawing),
    ...(opts.dropDataRels === true
      ? {}
      : { 'ppt/diagrams/_rels/data1.xml.rels': enc.encode(rels) }),
  };
  const { media, slideRels, ...rest } = opts.build ?? {};
  return buildPptx([frame], {
    slideRels: slideRels ?? [
      `<Relationship Id="rId100" Type="${DIAGRAM_DATA_REL}" Target="../diagrams/data1.xml"/>`,
    ],
    media: { ...parts, ...media },
    ...rest,
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

  it('finds the drawing PowerPoint names from inside the data (dsp:dataModelExt)', () => {
    // The drawing's relationship is the SLIDE's, and the data part points at it
    // by id from its own extension list. Looked for on the data part alone, two
    // corpus decks with a drawing sitting right there rendered nothing
    // (smartart-missing-bullet, tdf145528_SmartArt_Matrix).
    const deck = smartArtDeck({
      build: {
        slideRels: [
          `<Relationship Id="rId100" Type="${DIAGRAM_DATA_REL}" Target="../diagrams/data1.xml"/>` +
            `<Relationship Id="rId7" Type="${DIAGRAM_DRAWING_REL}" Target="../diagrams/drawing1.xml"/>`,
        ],
        media: {
          // The data names the drawing by a relationship of the SLIDE…
          'ppt/diagrams/data1.xml': new TextEncoder().encode(
            `<?xml version="1.0"?>\n<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
              `<dgm:extLst><a:ext xmlns:a="${A_MAIN}" uri="{x}">` +
              `<dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rId7"/>` +
              `</a:ext></dgm:extLst></dgm:dataModel>`,
          ),
        },
      },
      // …and the data part carries no relationships of its own.
      dropDataRels: true,
    });
    expect(shapeTexts(Ream.parse(deck))).toContain('NodeA');
  });

  it('gives each diagram on the slide ITS drawing, not the first one', () => {
    // tdf125551 carries four, each naming its own through `dsp:dataModelExt`.
    // Resolved by a slide-wide fallback they were one diagram drawn four times.
    const frame = (dm: string): string =>
      `<p:graphicFrame>` +
      `<p:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
      `<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ` +
      `r:dm="${dm}" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/>` +
      `</a:graphicData></a:graphic></p:graphicFrame>`;
    const enc = new TextEncoder();
    const drawing = (text: string): Uint8Array =>
      enc.encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="${A_MAIN}">` +
          `<dsp:spTree><dsp:sp><dsp:spPr>` +
          `<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${srgbFill('4472C4')}</dsp:spPr>` +
          `<dsp:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dsp:txBody>` +
          `</dsp:sp></dsp:spTree></dsp:drawing>`,
      );
    const data = (relId: string): Uint8Array =>
      enc.encode(
        `<?xml version="1.0"?>\n<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
          `<dgm:extLst><a:ext xmlns:a="${A_MAIN}" uri="{x}">` +
          `<dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="${relId}"/>` +
          `</a:ext></dgm:extLst></dgm:dataModel>`,
      );
    const doc = Ream.parse(
      buildPptx([frame('rIdD1') + frame('rIdD2')], {
        slideRels: [
          `<Relationship Id="rIdD1" Type="${DIAGRAM_DATA_REL}" Target="../diagrams/data1.xml"/>` +
            `<Relationship Id="rIdD2" Type="${DIAGRAM_DATA_REL}" Target="../diagrams/data2.xml"/>` +
            `<Relationship Id="rIdW1" Type="${DIAGRAM_DRAWING_REL}" Target="../diagrams/drawing1.xml"/>` +
            `<Relationship Id="rIdW2" Type="${DIAGRAM_DRAWING_REL}" Target="../diagrams/drawing2.xml"/>`,
        ],
        media: {
          'ppt/diagrams/data1.xml': data('rIdW1'),
          'ppt/diagrams/data2.xml': data('rIdW2'),
          'ppt/diagrams/drawing1.xml': drawing('First'),
          'ppt/diagrams/drawing2.xml': drawing('Second'),
        },
      }),
    );
    expect(shapeTexts(doc).sort()).toEqual(['First', 'Second']);
  });

  it('says so when the drawing override holds no shapes at all', () => {
    const stub =
      `<?xml version="1.0"?>\n<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ` +
      `xmlns:a="${A_MAIN}"><dsp:spTree/></dsp:drawing>`;
    const deck = smartArtDeck({
      build: { media: { 'ppt/diagrams/drawing1.xml': new TextEncoder().encode(stub) } },
    });
    const doc = Ream.parse(deck);
    expect(doc.flow.body.some((e) => e.kind === 'shape')).toBe(false);
    expect(doc.losses.some((l) => /drawing override/u.test(l.detail))).toBe(true);
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

// §20.1.10.42 `a:normAutofit` — the text shrinks until it fits the box it is
// in. PowerPoint usually writes the scale it settled on into `@fontScale` and
// that is applied at parse; a box that arrives WITHOUT one has never been
// fitted, and every box of a diagram laid out from its layout part is such a
// box. Unfitted, "Automatically shrinked text" at the layout's stated 65pt
// stood three lines deep and half a box wide outside the box that names it.
describe('a shape whose text shrinks to fit it', () => {
  // A 2in × 0.6in box holding more 65pt words than it can possibly take.
  const boxed = (bodyPr: string): string =>
    `<p:sp><p:spPr>` +
    `<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="548640"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody>${bodyPr}<a:p><a:r><a:rPr sz="6500"/>` +
    `<a:t>Automatically shrinked text</a:t></a:r></a:p></p:txBody></p:sp>`;

  // The largest size the page sets: the deck's own empty placeholder text sets
  // the default 11pt beside it, which is not what this box is about.
  const sizeOf = async (bodyPr: string): Promise<number> => {
    const pdf = await Ream.parse(buildPptx([boxed(bodyPr)])).convert('pdf', { fonts: FONTS });
    const sizes = [...latin1.decode(pdf).matchAll(/\/[A-Za-z0-9]+ ([\d.]+) Tf/gu)].map((m) =>
      Number(m[1]),
    );
    return Math.max(...sizes);
  };

  it('leaves the stated size alone without one', async () => {
    expect(await sizeOf('<a:bodyPr/>')).toBe(65);
  });

  it('brings it down to what the box holds with one', async () => {
    const size = await sizeOf('<a:bodyPr><a:normAutofit/></a:bodyPr>');
    expect(size).toBeLessThan(65);
    expect(size).toBeGreaterThan(0);
  });

  it('trusts the scale the producer already settled on', async () => {
    // A stated `fontScale` is PowerPoint's own answer and is applied as given;
    // measuring it again would fight the file over a box it has already fitted.
    expect(await sizeOf('<a:bodyPr><a:normAutofit fontScale="40000"/></a:bodyPr>')).toBe(26);
  });
});

// §19.3.1.15 `p:controls` — an ActiveX control sits BESIDE the shape tree, and
// its `mc:Fallback` is the picture a reader that cannot run it should draw.
describe('pptx slide: ActiveX controls', () => {
  const CONTROLS =
    `<p:controls>` +
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v">` +
    `<p:control spid="1" name="CommandButton1" r:id="rId9" imgW="1828800" imgH="1085760"/>` +
    `</mc:Choice>` +
    `<mc:Fallback>` +
    `<p:control name="CommandButton1" r:id="rId9" imgW="1828800" imgH="1085760">` +
    `<p:pic><p:nvPicPr><p:cNvPr id="8" name="CommandButton1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="rIdCtl"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="1000000" y="500000"/><a:ext cx="1828800" cy="1085760"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>` +
    `</p:control>` +
    `</mc:Fallback>` +
    `</mc:AlternateContent>` +
    `</p:controls>`;

  const deck = (): Uint8Array =>
    buildPptx([''], {
      slideControls: [CONTROLS],
      media: { 'ppt/media/ctl.png': buildTinyPng(2, 2, [0, 0, 255, 255]) },
      slideRels: [
        `<Relationship Id="rIdCtl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/ctl.png"/>`,
      ],
    });

  it("draws a control's cached picture, though the shape tree is empty", () => {
    // The control is not in `p:spTree` at all, so a reader that walks only that
    // tree drew a blank page — activex_picture.pptx is seventeen of them over
    // an empty tree.
    const doc = Ream.parse(deck());
    const pics = doc.flow.body.filter((e) => e.kind === 'image');
    expect(pics).toHaveLength(1);
    const only = pics[0];
    // Placed where the fallback picture's own transform puts it, floating, as
    // any other picture on the slide would be.
    if (only?.kind !== 'image') throw new Error('no picture');
    const float = only.image.float;
    expect(float?.posH?.offsetPt).toBeCloseTo(1000000 / 12700, 5);
    expect(float?.posV?.offsetPt).toBeCloseTo(500000 / 12700, 5);
  });
});
