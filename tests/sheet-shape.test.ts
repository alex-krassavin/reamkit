// E-SHEET W2 — floating shapes on a worksheet. An xdr:sp anchor renders as a
// ShapeBlock (geometry / fill / line / text), reusing the DrawingML readers the
// pptx + SmartArt paths use. The shape's box comes from its sheet anchor; runs
// use their direct a:rPr formatting (no placeholder cascade). It projects to a
// shape block after the grid — PDF/HTML draw it through the existing shape path.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import { readXlsxToSheetDoc } from '@/excel/xlsx-reader';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

const shapeXlsx = () =>
  buildXlsx({ rows: [['cell']], sheetShape: { text: 'Shape text', fillHex: '4472C4' } });

describe('sheet shapes — resolve (E-SHEET W2)', () => {
  it('resolves an xdr:sp anchor into a ShapeBlock with geometry, fill, line and text', () => {
    const sheet = readXlsxToSheetDoc(shapeXlsx()).sheets[0]!;
    expect(sheet.shapes).toHaveLength(1);
    const shape = sheet.shapes![0]!;
    expect(shape.width).toBeGreaterThan(0);
    expect(shape.height).toBeGreaterThan(0);
    expect(shape.geometry).toMatchObject({ kind: 'preset', preset: 'roundRect' });
    expect(shape.fill).toMatchObject({ kind: 'solid', colorHex: '4472C4' });
    expect(shape.line).toBeDefined();
    const para = shape.text?.content[0];
    expect(para?.kind).toBe('paragraph');
    if (para?.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(para.paragraph.runs[0]?.text).toBe('Shape text');
  });

  it('takes its outline and its text colour from <xdr:style> when spPr has none', () => {
    // §20.1.4.2.19/§20.1.4.2.14 — a shape drawn from a gallery style keeps its
    // outline in `a:lnRef` and its text colour in `a:fontRef`, and its spPr then
    // carries no `a:ln` at all. Read alone, spPr says the shape has no border
    // and its runs no colour: shape-macro-ext-ref.xlsx drew black text on a
    // green button with no rule around it, where both references draw white
    // text inside a blue one.
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [['cell']], sheetShape: { text: 'Go', styleOnly: true } }),
    ).sheets[0]!;
    const shape = sheet.shapes![0]!;
    expect(shape.line).toMatchObject({ colorHex: '123456' });
    const para = shape.text?.content[0];
    if (para?.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(para.paragraph.runs[0]?.properties.colorHex).toBe('FFFFFF');
  });

  it('lets a run\u2019s own <a:sysClr> beat the shape style\u2019s font colour', () => {
    // §20.1.2.3.32 — a system colour is STATED, not referenced: it carries its
    // own value in `lastClr` and there is nothing to look up. Unread, the run
    // fell back to `<a:fontRef>`, which on ConditionalFormattingSamples.xlsx
    // asks for `lt1`: fifteen navigation buttons came out white on pale blue.
    const sheet = readXlsxToSheetDoc(
      buildXlsx({
        rows: [['cell']],
        sheetShape: {
          text: 'Go',
          styleOnly: true,
          runColorXml: '<a:sysClr val="windowText" lastClr="000000"/>',
        },
      }),
    ).sheets[0]!;
    const para = sheet.shapes![0]!.text?.content[0];
    if (para?.kind !== 'paragraph') throw new Error('expected a paragraph');
    expect(para.paragraph.runs[0]?.properties.colorHex).toBe('000000');
  });

  it('takes its FILL from <a:fillRef> and its width from the theme (50299)', () => {
    // §20.1.4.2.10 is the outline's twin, and it was the half left unread: this
    // rectangle's spPr holds an `a:xfrm` and an `a:prstGeom` and nothing else,
    // so we drew it hollow on all six sheets that repeat it. Its `a:lnRef
    // idx="2"` indexes the theme's line styles, where the width is — assuming a
    // hairline drew a 2pt rule in 0.75pt.
    const sheet = readXlsxToSheetDoc(new Uint8Array(readFileSync('tests/fixtures/real/50299.xlsx')))
      .sheets[0]!;
    const shape = sheet.shapes![0]!;
    expect(shape.fill).toMatchObject({ kind: 'solid', colorHex: '4F81BD' }); // accent1
    expect(shape.line?.width).toBeCloseTo(2); // lnStyleLst[1] = 25400 EMU
  });

  it('walks into a shape GROUP and places each child in it', () => {
    // §20.5.2.17 — an anchor may frame a group rather than a shape, and the
    // walk looked for a direct `xdr:sp` only. groupShape.xlsx nests two groups
    // over three rectangles and we drew none of them.
    const drawing =
      `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="9" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
      `<xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></xdr:grpSpPr>` +
      `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="A"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500" cy="1000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></xdr:spPr></xdr:sp>` +
      `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="3" name="B"/><xdr:cNvSpPr/></xdr:nvSpPr>` +
      `<xdr:spPr><a:xfrm><a:off x="500" y="0"/><a:ext cx="500" cy="1000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
      `<a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></xdr:spPr></xdr:sp></xdr:grpSp>`;
    const sheet = readXlsxToSheetDoc(
      buildXlsx({ rows: [['cell']], sheetShape: { rawShapeXml: drawing } }),
    ).sheets[0]!;
    expect(sheet.shapes).toHaveLength(2);
    const [a, b] = sheet.shapes!;
    // Each child takes half the group's box, side by side.
    expect(a!.fill).toMatchObject({ kind: 'solid', colorHex: '4472C4' });
    expect(b!.fill).toMatchObject({ kind: 'solid', colorHex: 'ED7D31' });
    expect(a!.width).toBeCloseTo(b!.width);
    expect(b!.float?.posH?.offsetPt ?? 0).toBeGreaterThan(a!.float?.posH?.offsetPt ?? 0);
  });

  it('leaves a sheet with no drawing without a shapes field', () => {
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1]] })).sheets[0]!;
    expect(sheet.shapes).toBeUndefined();
  });

  it('skips a chart-only drawing (no shapes field)', () => {
    const chartXml = `<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
    const sheet = readXlsxToSheetDoc(buildXlsx({ rows: [[1]], sheetChart: { chartXml } }))
      .sheets[0]!;
    expect(sheet.shapes).toBeUndefined();
  });
});

describe('sheet shapes — projection (E-SHEET W2)', () => {
  it('projects the shape as a shape block after the grid', () => {
    const body = Ream.parse(shapeXlsx()).flow.body;
    const shape = body.find((el) => el.kind === 'shape');
    if (shape?.kind !== 'shape') throw new Error('expected a shape block');
    expect(shape.shape.fill).toMatchObject({ kind: 'solid', colorHex: '4472C4' });
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

describe('sheet shapes — render (E-SHEET W2)', () => {
  it('draws the shape fill and text into HTML', async () => {
    const html = new TextDecoder().decode(await Ream.parse(shapeXlsx()).convert('html'));
    expect(html).toContain('#4472C4'); // the shape fill
    expect(html).toContain('Shape text'); // the shape's text body
  });

  it('renders a sheet with a shape to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(shapeXlsx(), { fonts: FONTS });
    expect(pdf.length).toBeGreaterThan(1000);
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
  });
});

describe('a drawing that says it is hidden (§20.1.2.2.8)', () => {
  it('is not drawn at all', () => {
    // `cNvPr@hidden` — "Specifies whether this DrawingML object shall be
    // displayed". POI writes one white, black-outlined rectangle per cell
    // comment under the name `_xssf_cell_comment` and marks it hidden; read
    // without the flag, 51850.xlsx grew a 494 × 677pt box across both pages.
    const shown = readXlsxToSheetDoc(
      buildXlsx({ rows: [['cell']], sheetShape: { text: 'Ghost', fillHex: 'FFFFFF' } }),
    ).sheets[0]!;
    expect(shown.shapes).toHaveLength(1);
    const hidden = readXlsxToSheetDoc(
      buildXlsx({
        rows: [['cell']],
        sheetShape: { text: 'Ghost', fillHex: 'FFFFFF', hidden: true },
      }),
    ).sheets[0]!;
    expect(hidden.shapes).toBeUndefined();
  });
});
