// The DrawingML and chart percentages written with their sign — ST_Percentage
// (§20.1.10.40) and its kin are a union of a bare number and `"75%"` in the
// transitional schemas, and Office reads both. Each test spells one document
// both ways and expects the same model.

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { buildPptx } from './fixtures/build-pptx';
import type { PoNode } from '@/core/po-helpers';
import { Ream } from '@/core/converter/ream';
import { parseChart } from '@/core/drawingml/chart-parser';
import { defaultColorResolver, readColorMods } from '@/core/drawingml/colors';
import { readDocx } from '@/word/docx-reader';
import { parseSrcRect } from '@/word/drawing-parser';

/** One `a:x` element with `attrs`, and `kids` under it. */
const el = (tag: string, attrs: Record<string, string>, kids: Array<PoNode> = []): PoNode => ({
  [tag]: kids,
  ':@': Object.fromEntries(Object.entries(attrs).map(([k, v]) => [`@_${k}`, v])),
});

/** Every `"N%"` in `xml` spelled as the bare number `N × per` instead. */
const bare = (xml: string, per: number): string =>
  xml.replace(/"(-?[0-9]+(?:\.[0-9]+)?)%"/gu, (_, n: string) => `"${Number(n) * per}"`);

describe('DrawingML percentages written with their sign', () => {
  it('reads colour transforms', () => {
    const mods = (lumMod: string, alpha: string) =>
      readColorMods(
        el('a:schemeClr', { val: 'accent1' }, [
          el('a:lumMod', { val: lumMod }),
          el('a:lumOff', { val: '25%' }),
          el('a:alpha', { val: alpha }),
        ]),
      );
    expect(mods('75%', '50%')).toEqual([
      { kind: 'lumMod', val: 0.75 },
      { kind: 'lumOff', val: 0.25 },
      { kind: 'alpha', val: 0.5 },
    ]);
    expect(mods('75000', '50000')[0]).toEqual({ kind: 'lumMod', val: 0.75 });
  });

  it('reads a picture crop', () => {
    expect(parseSrcRect(el('a:srcRect', { l: '25%', t: '10.5%' }))).toEqual({
      left: 0.25,
      top: 0.105,
      right: 0,
      bottom: 0,
    });
  });

  it('reads a slide the same in either spelling', () => {
    const shape =
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:gradFill><a:gsLst><a:gs pos="0%"><a:srgbClr val="FF0000"><a:alpha val="40%"/></a:srgbClr></a:gs>' +
      '<a:gs pos="35%"><a:schemeClr val="accent1"><a:lumMod val="75%"/><a:lumOff val="25%"/></a:schemeClr></a:gs>' +
      '<a:gs pos="100%"><a:srgbClr val="0000FF"/></a:gs></a:gsLst>' +
      '<a:path path="circle"><a:fillToRect l="20%" t="50%" r="80%" b="50%"/></a:path></a:gradFill></p:spPr>' +
      '<p:txBody><a:bodyPr><a:normAutofit fontScale="62.5%" lnSpcReduction="20%"/></a:bodyPr>' +
      '<a:p><a:pPr><a:lnSpc><a:spcPct val="90%"/></a:lnSpc><a:spcBef><a:spcPct val="20%"/></a:spcBef>' +
      '<a:buSzPct val="80%"/><a:buChar char="•"/></a:pPr>' +
      '<a:r><a:rPr lang="en" sz="3200"/><a:t>Squeezed</a:t></a:r></a:p></p:txBody></p:sp>';
    const pic =
      '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture 4"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId7"><a:alphaModFix amt="52%"/></a:blip>' +
      '<a:srcRect l="25%" r="12.5%"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
    const body = (slide: string) =>
      Ream.parse(
        buildPptx([slide], {
          media: { 'ppt/media/image1.png': buildTinyPng(2, 2, [255, 0, 0, 255]) },
          slideRels: [
            '<Relationship Id="rId7" Target="../media/image1.png" ' +
              'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>',
          ],
        }),
      ).flow.body;
    const signed = body(shape + pic);
    expect(signed).toEqual(body(bare(shape + pic, 1000)));

    // ...and the bare spelling is not the default either: the numbers arrived.
    const sp = signed.find((e) => e.kind === 'shape');
    if (sp?.kind !== 'shape') throw new Error('expected a shape');
    const para = sp.shape.text?.content.find((c) => c.kind === 'paragraph');
    if (para?.kind !== 'paragraph') throw new Error('expected a paragraph');
    const runs = para.paragraph.runs;
    // 62.5% of 32pt, and its bullet 80% of that.
    expect(runs.find((r) => r.text === 'Squeezed')?.properties.fontSizePt).toBeCloseTo(20, 5);
    expect(runs.find((r) => r.text !== 'Squeezed')?.properties.fontSizePt).toBeCloseTo(16, 5);
    const img = signed.find((e) => e.kind === 'image');
    if (img?.kind !== 'image') throw new Error('expected a picture');
    expect(img.image.crop).toEqual({ left: 0.25, top: 0, right: 0.125, bottom: 0 });
  });

  it('reads a custom dash in a .docx the same in either spelling', () => {
    const rule = (ds: string): string =>
      '<w:p><w:r><w:drawing>' +
      '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
      '<wp:extent cx="914400" cy="12700"/><wp:docPr id="1" name="Rule"/>' +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="12700"/></a:xfrm>' +
      '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>' +
      `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:custDash>${ds}</a:custDash></a:ln>` +
      '</wps:spPr><wps:bodyPr/></wps:wsp>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const read = (ds: string) => readDocx(buildDocxFromBody(rule(ds))).doc.body;
    const signed = read('<a:ds d="300%" sp="100%"/><a:ds d="100%" sp="100%"/>');
    expect(signed).toEqual(read('<a:ds d="300000" sp="100000"/><a:ds d="100000" sp="100000"/>'));
    expect(JSON.stringify(signed)).toContain('"customDash":[3,1,1,1]');
  });
});

describe('chart percentages written with their sign', () => {
  const NS =
    'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const SERIES =
    '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:numCache><c:ptCount val="2"/>' +
    '<c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>';
  const AXES =
    '<c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
    '<c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>';
  const chart = (xml: string) => parseChart(new TextEncoder().encode(xml), defaultColorResolver)!;

  it('reads the gap and overlap of a bar chart as whole percents', () => {
    const bars = (gap: string, overlap: string): string =>
      `<c:chartSpace ${NS}><c:chart><c:plotArea><c:barChart><c:barDir val="col"/>` +
      `<c:grouping val="clustered"/>${SERIES}<c:gapWidth val="${gap}"/><c:overlap val="${overlap}"/>` +
      `<c:axId val="1"/><c:axId val="2"/></c:barChart>${AXES}</c:plotArea></c:chart></c:chartSpace>`;
    const signed = chart(bars('80%', '-27%'));
    expect(signed.gapPercent).toBe(80);
    expect(signed).toEqual(chart(bars('80', '-27')));
  });

  it("reads a 3-D bar chart's depth and gap as whole percents", () => {
    const bars = (depth: string, gap: string): string =>
      `<c:chartSpace ${NS}><c:chart><c:view3D><c:rotX val="15"/><c:rotY val="20"/>` +
      `<c:depthPercent val="${depth}"/><c:rAngAx val="1"/></c:view3D><c:plotArea><c:bar3DChart>` +
      `<c:barDir val="col"/><c:grouping val="clustered"/>${SERIES}<c:gapDepth val="${gap}"/>` +
      `<c:axId val="1"/><c:axId val="2"/><c:axId val="0"/></c:bar3DChart>${AXES}</c:plotArea>` +
      `</c:chart></c:chartSpace>`;
    expect(chart(bars('300%', '80%')).bar3D).toEqual({
      rotX: 15,
      rotY: 20,
      depthPercent: 300,
      gapDepth: 80,
    });
  });
});
