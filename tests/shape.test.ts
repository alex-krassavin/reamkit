import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { countShown, showPattern } from './fixtures/pdf-show';
import { eighthPtToPt, emuToPt, halfPtToPt, twipsToPt } from '@/core/ir';

import { convertDocxToPdfSync } from '@/core/converter';
import { parseTtf } from '@/core/font';
import { OpcPackage } from '@/core/opc';
import { applyColorMods } from '@/core/drawingml/colors';
import { parseDocument } from '@/word';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};
const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

// A <w:drawing> carrying a wps:wsp shape with the given spPr inner XML.
function drawingEl(spPrInner: string, cx = 1828800, cy = 914400): string {
  return `<w:drawing>
    <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
      <wp:extent cx="${cx}" cy="${cy}"/>
      <wp:docPr id="1" name="Shape 1"/>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
              ${spPrInner}
            </wps:spPr>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;
}

const shapeRun = (spPrInner: string, cx?: number, cy?: number): string =>
  `<w:r>${drawingEl(spPrInner, cx, cy)}</w:r>`;

const RECT_FILL_STROKE = `
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
  <a:ln w="12700"><a:solidFill><a:srgbClr val="2F528F"/></a:solidFill></a:ln>`;

describe('DrawingML shape parsing', () => {
  it('parses a wps:wsp rect into a shape BodyElement', () => {
    const docx = buildDocxFromBody(`<w:p>${shapeRun(RECT_FILL_STROKE)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    expect(parsed).toHaveLength(1);
    expect(parsed[0]!.kind).toBe('shape');
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    const shape = parsed[0]!.shape;
    expect(shape.width).toBe(emuToPt(1828800));
    expect(shape.height).toBe(emuToPt(914400));
    expect(shape.geometry.kind).toBe('preset');
    expect(shape.geometry.preset).toBe('rect');
    expect(shape.fill).toEqual({ kind: 'solid', colorHex: '4472C4' });
    expect(shape.line?.colorHex).toBe('2F528F');
    expect(shape.line?.width).toBe(emuToPt(12700));
  });

  it('parses prstGeom adjust values', () => {
    const inner = `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
      <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>`;
    const docx = buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    expect(parsed[0]!.shape.geometry.adjust?.get('adj')).toBe(16667);
  });

  it('resolves an unknown schemeClr through the default Office palette', () => {
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`;
    const docx = buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    expect(parsed[0]!.shape.fill).toEqual({ kind: 'solid', colorHex: '4472C4' });
  });

  it('applies a lumMod/lumOff transform to a scheme fill (Accent 1, Lighter 40%)', () => {
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill>`;
    const docx = buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    const fill = parsed[0]!.shape.fill;
    if (fill.kind !== 'solid') throw new Error('expected solid fill');
    expect(fill.colorHex).not.toBe('4472C4'); // lightened, not the raw accent
    const sum = [0, 2, 4].reduce((a, i) => a + parseInt(fill.colorHex!.slice(i, i + 2), 16), 0);
    expect(sum).toBeGreaterThan(0x44 + 0x72 + 0xc4); // brighter than the base
  });

  it('parses a gradient fill into stops + linear direction (E-PDF EP16)', () => {
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:gradFill><a:gsLst>
        <a:gs pos="0"><a:srgbClr val="000000"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
      </a:gsLst><a:lin ang="2700000"/></a:gradFill>`;
    const docx = buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    expect(parsed[0]!.shape.fill).toEqual({
      kind: 'gradient',
      gradient: {
        kind: 'linear',
        angle: 45, // 2700000 / 60000
        stops: [
          { offset: 0, colorHex: '000000' },
          { offset: 1, colorHex: 'FFFFFF' },
        ],
      },
    });
  });
});

describe('colour transforms (§20.1.2.3)', () => {
  it('shade darkens toward black, on LINEAR light', () => {
    // §20.1.2.3.31 — the fraction is of the light the colour stands for, not of
    // the byte that encodes it. Taken on the byte, half of the Office accent
    // came out 223962, a shade both references draw as 2F528F.
    expect(applyColorMods('4472C4', [{ kind: 'shade', val: 0.5 }])).toBe('2F528F');
  });

  it('tint lightens toward white, on the same light', () => {
    expect(applyColorMods('000000', [{ kind: 'tint', val: 0.5 }])).toBe('BCBCBC');
  });

  it('alpha composites the colour over the white page', () => {
    // §20.1.2.3.1 — dml-groupshape-childposition.docx draws eleven of its
    // strokes at 20% opacity, which every reader shows as a pale tint; ignored,
    // they came out in full dark navy.
    expect(applyColorMods('000000', [{ kind: 'alpha', val: 0.2 }])).toBe('CCCCCC');
    expect(applyColorMods('4472C4', [{ kind: 'alpha', val: 1 }])).toBe('4472C4');
  });

  it('no transforms is an identity', () => {
    expect(applyColorMods('4472C4', [])).toBe('4472C4');
  });
});

describe('shape edge cases', () => {
  it('renders an anchored (floating) shape as a block', () => {
    const cx = 1828800;
    const cy = 914400;
    const body = `<w:p><w:r><w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" simplePos="0" behindDoc="0">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="${cx}" cy="${cy}"/>
        <wp:docPr id="1" name="Shape"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                ${RECT_FILL_STROKE}
              </wps:spPr>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing></w:r></w:p>`;
    const docx = buildDocxFromBody(body);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    expect(parsed[0]!.kind).toBe('shape');
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    expect(parsed[0]!.shape.width).toBe(emuToPt(cx));
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toMatch(/\nh\nB\n/);
  });

  it('clamps an oversized shape so it stays on the page', () => {
    // 50-inch-tall shape on A4 → must be scaled to fit, not pushed off-page.
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`;
    const docx = buildDocxFromBody(`<w:p>${shapeRun(inner, 914400, 45720000)}</w:p>`);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    const m = /1 0 0 1 (-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) cm/.exec(text);
    expect(m).not.toBeNull();
    // The y-translate (shape bottom) must be on the page, not far negative.
    expect(Number(m![2])).toBeGreaterThanOrEqual(-1);
  });

  it('draws a shape in a mixed text+shape run ahead of the paragraph', () => {
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`;
    const body = `<w:p><w:r><w:t>Hello</w:t></w:r>${shapeRun(inner)}</w:p>`;
    const docx = buildDocxFromBody(body);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    // The shape leaves the run for a block of its own; the paragraph keeps
    // its text (the line model carries pictures, not shapes).
    expect(parsed.map((el) => el.kind)).toEqual(['shape', 'paragraph']);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toContain('0.266667 0.447059 0.768627 rg');
  });
});

describe('shape theme colours (end-to-end)', () => {
  const customTheme = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:themeElements><a:clrScheme name="Custom">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F3864"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
    </a:clrScheme></a:themeElements>
  </a:theme>`;
  const schemeFill = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`;

  it('uses the document theme accent over the default palette', () => {
    const docx = buildDocxFromBody(`<w:p>${shapeRun(schemeFill)}</w:p>`, { themeXml: customTheme });
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toContain('1 0 0 rg'); // accent1 = FF0000 from the theme
  });

  it('falls back to the default Office accent when no theme part exists', () => {
    const docx = buildDocxFromBody(`<w:p>${shapeRun(schemeFill)}</w:p>`);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // default accent1 4472C4 → 0.266667 0.447059 0.768627 rg
    expect(text).toContain('0.266667 0.447059 0.768627 rg');
  });
});

describe('DrawingML shape rendering end-to-end', () => {
  it('emits a filled + stroked vector path for a rect shape', () => {
    const docx = buildDocxFromBody(`<w:p>${shapeRun(RECT_FILL_STROKE)}</w:p>`);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    expect(text).toMatch(/ cm\n/); // placement transform
    expect(text).toContain('0 0 m'); // path start (bottom-left)
    expect(text).toMatch(/\nh\nB\n/); // close then fill+stroke
    expect(text).toMatch(/ rg\n/); // non-stroking (fill) colour
    expect(text).toMatch(/ RG\n/); // stroking (line) colour
  });

  it('fill-only shape paints with f (no stroke)', () => {
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>`;
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`), { fonts: FONTS }),
    );
    expect(text).toMatch(/\nh\nf\n/);
    expect(text).not.toMatch(/ RG\n/);
  });

  it('no-fill + line shape paints with S (stroke only)', () => {
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/>
      <a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>`;
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`), { fonts: FONTS }),
    );
    expect(text).toMatch(/\nh\nS\n/);
    expect(text).not.toMatch(/ rg\n/);
  });

  it('roundRect and ellipse emit Bézier (c) operators', () => {
    const round = `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`;
    const ellipse = `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`;
    for (const inner of [round, ellipse]) {
      const text = asLatin1(
        convertDocxToPdfSync(buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`), { fonts: FONTS }),
      );
      expect(text).toMatch(/ c\n/);
    }
  });

  it('an unknown preset falls back to a filled bounding rectangle', () => {
    const inner = `<a:prstGeom prst="cloudCallout"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="A5A5A5"/></a:solidFill>`;
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`), { fonts: FONTS }),
    );
    // Rect fallback: a closed path filled with f (no Bézier).
    expect(text).toMatch(/\nh\nf\n/);
    expect(text).not.toMatch(/ c\n/);
  });
});

describe('Markup Compatibility (mc:AlternateContent)', () => {
  it('prefers the wps Choice over the VML Fallback', () => {
    const choiceDrawing = drawingEl(RECT_FILL_STROKE);
    const body = `<w:p><w:r>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice Requires="wps">${choiceDrawing}</mc:Choice>
        <mc:Fallback><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml"/></w:pict></mc:Fallback>
      </mc:AlternateContent>
    </w:r></w:p>`;
    const parsed = parseDocument(OpcPackage.open(buildDocxFromBody(body)).getMainDocument().data);
    expect(parsed).toHaveLength(1);
    expect(parsed[0]!.kind).toBe('shape');
  });

  it('ignores a Choice whose Requires we do not understand (falls to Fallback)', () => {
    const body = `<w:p><w:r>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <mc:Choice Requires="aink">${drawingEl(RECT_FILL_STROKE)}</mc:Choice>
        <mc:Fallback><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml"/></w:pict></mc:Fallback>
      </mc:AlternateContent>
    </w:r></w:p>`;
    const parsed = parseDocument(OpcPackage.open(buildDocxFromBody(body)).getMainDocument().data);
    // The wps Choice required an unknown namespace, so we take the VML Fallback
    // (which we can't render) → no shape, just an (empty) paragraph.
    expect(parsed[0]!.kind).toBe('paragraph');
  });
});

describe('DrawingML shape line styling', () => {
  const dashedInner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln w="25400" cap="rnd">
      <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      <a:prstDash val="dash"/>
    </a:ln>`;

  it('parses prstDash and cap', () => {
    const docx = buildDocxFromBody(`<w:p>${shapeRun(dashedInner)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    expect(parsed[0]!.shape.line?.dash).toBe('dash');
    expect(parsed[0]!.shape.line?.cap).toBe('round');
    expect(parsed[0]!.shape.line?.width).toBe(emuToPt(25400));
  });

  it('emits a dash array, line width and round cap', () => {
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(`<w:p>${shapeRun(dashedInner)}</w:p>`), {
        fonts: FONTS,
      }),
    );
    expect(text).toContain('2 w'); // 25400 EMU = 2pt
    expect(text).toContain('[8 6] 0 d'); // dash = [4w, 3w] at w=2
    expect(text).toContain('1 J'); // round cap
  });
});

describe('custom geometry (custGeom)', () => {
  const triangleInner = `<a:custGeom><a:pathLst><a:path w="100" h="100">
      <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
      <a:lnTo><a:pt x="100" y="0"/></a:lnTo>
      <a:lnTo><a:pt x="50" y="100"/></a:lnTo>
      <a:close/>
    </a:path></a:pathLst></a:custGeom>
    <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`;

  it('parses custGeom into a custom-kind geometry', () => {
    const parsed = parseDocument(
      OpcPackage.open(buildDocxFromBody(`<w:p>${shapeRun(triangleInner)}</w:p>`)).getMainDocument()
        .data,
    );
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    const geom = parsed[0]!.shape.geometry;
    expect(geom.kind).toBe('custom');
    expect(geom.custom?.pathWidth).toBe(100);
    expect(geom.custom?.commands.map((c) => c.cmd)).toEqual(['move', 'line', 'line', 'close']);
  });

  it('renders a cubicBezTo as a c operator', () => {
    const inner = `<a:custGeom><a:pathLst><a:path w="100" h="100">
        <a:moveTo><a:pt x="0" y="0"/></a:moveTo>
        <a:cubicBezTo><a:pt x="0" y="100"/><a:pt x="100" y="100"/><a:pt x="100" y="0"/></a:cubicBezTo>
        <a:close/>
      </a:path></a:pathLst></a:custGeom>
      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`;
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(`<w:p>${shapeRun(inner)}</w:p>`), { fonts: FONTS }),
    );
    expect(text).toMatch(/ c\n/);
  });
});

describe('text in shape (wps:txbx)', () => {
  const textBox = (inner: string, bodyPr = '<wps:bodyPr/>'): string => {
    const cx = 2743200;
    const cy = 1097280;
    return `<w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="${cx}" cy="${cy}"/>
        <wp:docPr id="1" name="TextBox"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:srgbClr val="DEEBF7"/></a:solidFill>
              </wps:spPr>
              <wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx>
              ${bodyPr}
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing></w:r>`;
  };

  it('parses the text box content, insets and anchor', () => {
    const inner = `<w:p><w:r><w:t>Label</w:t></w:r></w:p>`;
    const bodyPr = `<wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"/>`;
    const docx = buildDocxFromBody(`<w:p>${textBox(inner, bodyPr)}</w:p>`);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    if (parsed[0]!.kind !== 'shape') throw new Error('unreachable');
    const t = parsed[0]!.shape.text;
    expect(t).toBeDefined();
    expect(t!.content).toHaveLength(1);
    expect(t!.anchor).toBe('ctr');
    expect(t!.insetLeft).toBe(emuToPt(91440));
  });

  it('renders the text-box glyphs on top of the shape fill', () => {
    const inner = `<w:p><w:r><w:t>Label</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(`<w:p>${textBox(inner)}</w:p>`);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));

    const parsed = parseTtf(FONTS.regular);

    const label = text.search(showPattern(parsed, 'Label'));
    expect(label).toBeGreaterThan(-1); // text-box content rendered
    // Shape fill paints (f) before the text pass (BT) → text on top.
    expect(text.indexOf('\nf\n')).toBeLessThan(text.indexOf('BT'));
    expect(text.indexOf('BT')).toBeLessThan(label);
  });
});

// §20.5.2.17 `wpg:wgp` — a drawing group: a box holding shapes of its own, in
// its own coordinate space. Left unread, Tdf147485.docx — whose whole picture
// is one — printed an empty page, because the mc:Fallback beside it is VML.
describe('a drawing group', () => {
  const member = (x: number, y: number, cx: number, cy: number, hex: string): string =>
    `<wps:wsp><wps:spPr>
       <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
       <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
       <a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>
     </wps:spPr><wps:bodyPr/></wps:wsp>`;

  const groupDocx = (grpXfrm: string, members: string): Uint8Array =>
    buildDocxFromBody(`<w:p><w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="1828800" cy="914400"/>
        <wp:docPr id="1" name="Group 1"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                     xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wpg:grpSpPr><a:xfrm>${grpXfrm}</a:xfrm></wpg:grpSpPr>
              ${members}
            </wpg:wgp>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing></w:r></w:p>`);

  const IDENTITY =
    '<a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/>';

  it('draws every member of the group', () => {
    const text = asLatin1(
      convertDocxToPdfSync(
        groupDocx(
          IDENTITY,
          member(0, 0, 914400, 457200, 'FF0000') + member(914400, 457200, 914400, 457200, '00FF00'),
        ),
        { fonts: FONTS },
      ),
    );
    // One fill each, in the colours the members name (1.0 0 0 and 0 1.0 0).
    expect(text).toContain('1 0 0 rg');
    expect(text).toContain('0 1 0 rg');
  });

  it('maps a member out of the group’s child coordinate space', () => {
    // The child space is half the size of the box, so a member at (0, 0) of
    // 457200×228600 fills the same quarter as one at (0, 0) of 914400×457200
    // would in an identity space: 72pt × 36pt.
    const halved =
      '<a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/>';
    const text = asLatin1(
      convertDocxToPdfSync(groupDocx(halved, member(0, 0, 457200, 228600, 'FF0000')), {
        fonts: FONTS,
      }),
    );
    expect(text).toMatch(/0 0 m\n72 0 l\n72 36 l\n0 36 l/u);
  });

  it('offsets a member by the child origin, not the page’s', () => {
    const shifted =
      '<a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>' +
      '<a:chOff x="914400" y="457200"/><a:chExt cx="1828800" cy="914400"/>';
    // A member sitting AT the child origin is at the group's own corner.
    const text = asLatin1(
      convertDocxToPdfSync(groupDocx(shifted, member(914400, 457200, 914400, 457200, 'FF0000')), {
        fonts: FONTS,
      }),
    );
    // Its transform translates by the group's corner, no further.
    expect(text).toMatch(/1 0 0 1 72 \d/u);
  });

  it('skips a member that states no size', () => {
    const text = asLatin1(
      convertDocxToPdfSync(
        groupDocx(
          IDENTITY,
          '<wps:wsp><wps:spPr/><wps:bodyPr/></wps:wsp>' + member(0, 0, 914400, 457200, 'FF0000'),
        ),
        { fonts: FONTS },
      ),
    );
    expect(text).toContain('1 0 0 rg');
  });
});

// §20.1.4.2.13/19 — a shape drawn from a gallery style keeps its fill and its
// outline in `<wps:style>` and carries neither in `spPr`. Read alone, spPr says
// the shape has none: TextEffects_Groupshapes.docx drew its caption on white
// where LibreOffice fills an accent-blue rectangle behind it.
describe('a shape filled with a picture (§20.1.8.14)', () => {
  const png = buildTinyPng(2, 2, [255, 0, 0, 255]);
  const filled = (blipFill: string): string =>
    asLatin1(
      convertDocxToPdfSync(
        buildDocxFromBody(
          `<w:p><w:r><w:drawing>
          <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
            <wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="S"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <wps:spPr>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    ${blipFill}
                  </wps:spPr>
                  <wps:style><a:fillRef idx="1"><a:srgbClr val="ED7D31"/></a:fillRef></wps:style>
                  <wps:bodyPr/>
                </wps:wsp>
              </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`,
          { images: { rId9: { bytes: png, extension: 'png', contentType: 'image/png' } } },
        ),
        { fonts: FONTS },
      ),
    );

  it('draws the picture, not the gallery colour beneath it', () => {
    // crop-roundtrip.docx fills a rectangle with a photo; read as no fill at
    // all, the shape fell through to its style and drew a plain orange box.
    const text = filled(
      '<a:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
    );
    expect(text).toContain(' Do'); // the image XObject is painted
    expect(text).not.toMatch(/0\.929412 0\.490196 0\.192157 rg/u); // …and ED7D31 is not
  });

  it('shows the part a negative fillRect frames', () => {
    // §20.1.8.30 — a negative inset pushes the picture's edge outside the box,
    // so what remains inside is a zoomed-in part of it: the emitter clips.
    const text = filled(
      '<a:blipFill><a:blip r:embed="rId9"/>' +
        '<a:stretch><a:fillRect t="-100000" b="0"/></a:stretch></a:blipFill>',
    );
    expect(text).toMatch(/re\nW\nn/u); // a clip path around the box
  });
});

describe('a gallery-styled shape', () => {
  const styled = (styleXml: string, spPrInner: string, extra = ''): string =>
    asLatin1(
      convertDocxToPdfSync(
        buildDocxFromBody(`<w:p><w:r><w:drawing>
          <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
            <wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="S"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                    ${spPrInner}
                  </wps:spPr>
                  ${styleXml}
                  ${extra}
                  <wps:bodyPr/>
                </wps:wsp>
              </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`),
        { fonts: FONTS },
      ),
    );
  const RECT = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';

  it('takes the fill and outline the style names', () => {
    const text = styled(
      '<wps:style><a:lnRef idx="2"><a:srgbClr val="2E75B6"/></a:lnRef>' +
        '<a:fillRef idx="1"><a:srgbClr val="5B9BD5"/></a:fillRef></wps:style>',
      RECT,
    );
    // 0x5B/255, 0x9B/255, 0xD5/255 — filled …
    expect(text).toMatch(/0\.356863 0\.607843 0\.835294 rg/u);
    // … and stroked in 0x2E/0x75/0xB6.
    expect(text).toMatch(/0\.180392 0\.458824 0\.713725 RG/u);
  });

  it('draws a pattern fill as the ink it lays down', () => {
    // §20.1.8.37 — dml-shape-fillpattern.docx rules twelve rectangles with
    // `a:pattFill` and we drew twelve empty boxes. A tile is beyond a vector
    // fill; the two colours blended by the pattern's coverage is much nearer
    // than nothing at all.
    const text = styled(
      '',
      `${RECT}<a:pattFill prst="ltHorz"><a:fgClr><a:srgbClr val="00FF00"/></a:fgClr>` +
        '<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
    );
    // 15% of pure green over white.
    expect(text).toContain('0.85098 1 0.85098 rg');
  });

  it('keeps a fill the shape states for itself', () => {
    const text = styled(
      '<wps:style><a:fillRef idx="1"><a:srgbClr val="5B9BD5"/></a:fillRef></wps:style>',
      `${RECT}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`,
    );
    expect(text).toContain('1 0 0 rg');
    expect(text).not.toMatch(/0\.356863 0\.607843 0\.835294 rg/u);
  });

  it('gives the text the colour the style names', () => {
    // §20.1.4.2.14 — a run with no colour of its own takes the fontRef's.
    // LineStyle_DashType.docx asks for white on seven blue rectangles.
    const text = styled(
      '<wps:style><a:fillRef idx="1"><a:srgbClr val="4472C4"/></a:fillRef>' +
        '<a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></wps:style>',
      `${RECT}`,
      '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Caption</w:t></w:r></w:p></w:txbxContent></wps:txbx>',
    );
    expect(text).toContain('1 1 1 rg');
  });

  it("leaves a run that inherits a colour with the style's", () => {
    // §17.7.2 — the theme's font colour is the FLOOR of the cascade, not the
    // top of it. ColorOverwritten.docx writes its arrow's two lines in a "red"
    // and a "green" paragraph style, and stamping the theme's white over them
    // left the shape looking blank.
    const body = `<w:p><w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="S"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
              <wps:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></wps:style>
              <wps:txbx><w:txbxContent>
                <w:p><w:pPr><w:pStyle w:val="red"/></w:pPr><w:r><w:t>Ausgang</w:t></w:r></w:p>
                <w:p><w:r><w:t>plain</w:t></w:r></w:p>
              </w:txbxContent></wps:txbx>
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const text = asLatin1(
      convertDocxToPdfSync(
        buildDocxFromBody(body, {
          stylesXml:
            '<w:style w:type="paragraph" w:styleId="red"><w:name w:val="red"/>' +
            '<w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>',
        }),
        { fonts: FONTS },
      ),
    );
    expect(text).toContain('1 0 0 rg'); // the style's red survives
    expect(text).toContain('1 1 1 rg'); // …and the plain paragraph takes white
  });

  it('takes the style colour under a rule that states only a width', () => {
    // §20.1.4.2.19 — the shape's own `a:ln` says how THICK and how dashed, the
    // gallery style says what COLOUR. dashed_line_custdash_percentage.docx
    // rules a 4.5pt accent-blue line that way and we drew a black hairline.
    const text = styled(
      '<wps:style><a:lnRef idx="1"><a:srgbClr val="4472C4"/></a:lnRef></wps:style>',
      `${RECT}<a:ln w="57150"><a:custDash><a:ds d="800000" sp="300000"/></a:custDash></a:ln>`,
    );
    expect(text).toContain('0.266667 0.447059 0.768627 RG'); // 4472C4
    expect(text).toContain('4.5 w');
    // §20.1.8.21 — dash 800% and space 300% of a 4.5pt rule.
    expect(text).toContain('[36 13.5] 0 d');
  });

  it('finds the colour inside a pretty-printed reference', () => {
    // A text node between the elements is not the element: the "not #text"
    // test took the indentation and read no colour at all.
    const text = styled(
      '<wps:style>\n  <a:fillRef idx="1">\n    <a:srgbClr val="5B9BD5"/>\n  </a:fillRef>\n</wps:style>',
      RECT,
    );
    expect(text).toContain('0.356863 0.607843 0.835294 rg');
  });

  it('reads idx="0" as naming nothing', () => {
    const text = styled(
      '<wps:style><a:fillRef idx="0"><a:srgbClr val="5B9BD5"/></a:fillRef></wps:style>',
      RECT,
    );
    expect(text).not.toMatch(/0\.356863 0\.607843 0\.835294 rg/u);
  });
});

// §14.1.2.22 `v:textpath` — legacy WordArt, whose words live in an attribute
// rather than in the document body. Read nowhere, WordArt.docx printed an empty
// page. The preset path is a shapetype of formulas we do not evaluate, so the
// words are set flat in the shape's box — which is the whole of what it says.
describe('a legacy VML shape (§14.1.2)', () => {
  const pictOf = (inner: string): string =>
    asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(`<w:p><w:r><w:pict>${inner}</w:pict></w:r></w:p>`), {
        fonts: FONTS,
      }),
    );

  it('draws a rectangle with its fill, its outline and its words', () => {
    // drawinglayer-pic-pos.docx frames its title in one of these, and read
    // only for a `v:imagedata` the page came out without it.
    const text = pictOf(
      '<v:rect style="position:absolute;margin-left:10pt;margin-top:20pt;width:200pt;height:100pt"' +
        ' fillcolor="#4472C4" strokecolor="red" strokeweight="2pt">' +
        '<v:textbox><w:txbxContent><w:p><w:r><w:t>Framed</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
        '</v:rect>',
    );
    expect(text).toContain('0.266667 0.447059 0.768627 rg'); // 4472C4 fill
    expect(text).toContain('1 0 0 RG'); // red outline
    expect(text).toContain('2 w');
    expect(text).toMatch(/<[0-9A-F]+>[^\n]*T[Jj]/u); // the words in it
  });

  it('reads a named colour and an unfilled shape', () => {
    const text = pictOf('<v:oval style="width:50pt;height:50pt" filled="f" strokecolor="navy"/>');
    expect(text).toContain('0 0 0.501961 RG'); // navy
    expect(text).not.toMatch(/ rg\n/u); // nothing filled
  });

  it("places a group's members in the group's own coordinate space", () => {
    // dml-textshape.docx draws its whole diagram inside a v:group, and read as
    // one shape it drew nothing at all.
    const text = pictOf(
      '<v:group style="width:200pt;height:100pt" coordsize="2000,1000">' +
        '<v:rect style="position:absolute;left:1000;top:500;width:1000;height:500" fillcolor="#00FF00"/>' +
        '</v:group>',
    );
    // Half the group's box, at its centre: a 100×50pt rectangle.
    expect(text).toMatch(/100 0 l/u);
    expect(text).toContain('0 1 0 rg');
  });
});

describe('VML WordArt', () => {
  const pict = (shapeXml: string): string =>
    asLatin1(
      convertDocxToPdfSync(
        buildDocxFromBody(`<w:p><w:r><w:pict>${shapeXml}</w:pict></w:r></w:p>`),
        {
          fonts: FONTS,
        },
      ),
    );
  // The shapetype template beside the shape carries a textpath of its own, with
  // no string on it — the words are on the SHAPE's.
  const SHAPETYPE =
    '<v:shapetype id="_x0000_t144" o:spt="144"><v:textpath on="t" fitpath="t"/></v:shapetype>';

  it('sets the string the textpath carries', () => {
    const text = pict(
      `${SHAPETYPE}<v:shape id="s" type="#_x0000_t144" style="width:286.45pt;height:134.8pt" fillcolor="black">` +
        '<v:textpath style="font-family:&quot;Arial Black&quot;" string="WORD-ART"/></v:shape>',
    );
    const parsed = parseTtf(FONTS.regular);
    expect(text).toMatch(showPattern(parsed, 'WORD-ART'));
  });

  it('sizes it to the box it is given, width included', () => {
    // Half the height per line is close for a line of capitals, but a size
    // that overflows the width wraps: "WORD-ART" came out as "WORD-A / RT".
    const text = pict(
      `${SHAPETYPE}<v:shape id="s" type="#_x0000_t144" style="width:286.45pt;height:134.8pt">` +
        '<v:textpath string="WORD-ART"/></v:shape>',
    );
    const size = /\/F\d+ ([\d.]+) Tf/u.exec(text);
    expect(size).not.toBeNull();
    // 286.45 / (8 × 0.62) = 57.8pt, under the 67.4pt the height alone allows.
    expect(Number(size![1])).toBeCloseTo(57.75, 0);
  });

  it('ignores a shape that names no string or no size', () => {
    expect(pict(`${SHAPETYPE}<v:shape id="s" style="width:100pt;height:50pt"/>`)).not.toContain(
      ' Tj',
    );
    expect(pict(`${SHAPETYPE}<v:shape id="s"><v:textpath string="X"/></v:shape>`)).not.toContain(
      ' Tj',
    );
  });
});

// §20.1.10.83 `a:bodyPr @vert` — text set along the box's long axis rather than
// across it. btlr-textbox.docx reads bottom-to-top and we set it flat, so it
// ran out of the box the wrong way.
describe('vertical text in a shape', () => {
  const vertical = (vert: string): string => {
    const inner = '<w:p><w:r><w:t>Sideways</w:t></w:r></w:p>';
    const body = `<w:p><w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="2743200" cy="1828800"/><wp:docPr id="1" name="S"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
              <wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx>
              <wps:bodyPr vert="${vert}"/>
            </wps:wsp>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    return asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
  };

  it('turns the line a quarter, one way for each mode', () => {
    // A quarter turn is a text matrix of (0 1 -1 0) or its opposite, not the
    // (1 0 0 1) a flat line uses.
    expect(vertical('vert270')).toMatch(/\n0 1 -1 0 [\d.]+ [\d.]+ Tm\n/u);
    expect(vertical('vert')).toMatch(/\n0 -1 1 0 [\d.]+ [\d.]+ Tm\n/u);
  });

  it('leaves horizontal text flat', () => {
    expect(vertical('horz')).toMatch(/\n1 0 0 1 [\d.]+ [\d.]+ Tm\n/u);
  });
});

// §20.1.10.28 `a:spAutoFit` — the shape follows its text: the box it states is
// a starting size and the height is whatever the text needs. Ignored,
// autofit.docx drew its one-line box as tall as the four-line box beside it.
describe('a shape that fits itself to its text', () => {
  const boxed = (bodyPr: string): string => {
    const body = `<w:p><w:r><w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="2743200" cy="1828800"/><wp:docPr id="1" name="S"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>
              <wps:txbx><w:txbxContent><w:p><w:r><w:t>One line.</w:t></w:r></w:p></w:txbxContent></wps:txbx>
              ${bodyPr}
            </wps:wsp>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    return asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
  };
  // The outline's own path says how tall the shape came out.
  const heightOf = (text: string): number =>
    Number(/0 0 m\n[\d.]+ 0 l\n[\d.]+ ([\d.]+) l/u.exec(text)![1]);

  it('shrinks the box to the text it holds', () => {
    // 144pt as stated, against one 11pt line plus the default 3.6pt insets.
    expect(heightOf(boxed('<wps:bodyPr/>'))).toBeCloseTo(144, 0);
    expect(heightOf(boxed('<wps:bodyPr><a:spAutoFit/></wps:bodyPr>'))).toBeLessThan(30);
  });
});

describe('a floating drawing and the text column', () => {
  const anchored = (cxEmu: number, inner: string): string =>
    `<w:p><w:r><w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                 distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1"
                 behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="${cxEmu}" cy="914400"/><wp:wrapNone/><wp:docPr id="1" name="S"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
            <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxEmu}" cy="914400"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>
              ${inner}
              <wps:bodyPr/>
            </wps:wsp>
          </a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;

  it('draws the floats in the z-order they state', () => {
    // §20.4.2.3 `relativeHeight` — dml-rectangle-relsize.docx writes its blue
    // bar FIRST and gives it the higher z, so it belongs over the white
    // rectangle that follows; drawn in document order the rectangle hid it.
    const float = (z: number, hex: string, cy: number): string =>
      `<w:p><w:r><w:drawing>
        <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                   distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${z}"
                   behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
          <wp:extent cx="914400" cy="${cy}"/><wp:wrapNone/><wp:docPr id="1" name="S"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="${cy}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  <a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></wps:spPr>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
    const text = asLatin1(
      convertDocxToPdfSync(
        // The taller white one is written second but sits BELOW.
        buildDocxFromBody(float(20, 'FF0000', 114300) + float(10, 'FFFFFF', 914400)),
        { fonts: FONTS },
      ),
    );
    expect(text.indexOf('1 1 1 rg')).toBeLessThan(text.indexOf('1 0 0 rg'));
  });

  it('takes a size stated as a share of the margins', () => {
    // `wp14:sizeRelH/V` — dml-shape-relsize.docx asks for 40% of the margin
    // width and 20% of its height; read as nothing, the shape came out at the
    // fallback extent, less than half as wide.
    const relative =
      `<w:p><w:r><w:drawing>
        <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                   xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
                   distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1"
                   behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
          <wp:simplePos x="0" y="0"/>
          <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
          <wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="1" name="S"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData></a:graphic>
          <wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>50000</wp14:pctWidth></wp14:sizeRelH>
        </wp:anchor></w:drawing></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(relative), { fonts: FONTS }));
    // Half of the 468pt column, and the height follows to keep the shape square.
    expect(text).toMatch(/234 0 l/u);
    expect(text).toMatch(/234 117 l/u);
  });

  it('keeps the width it states, past the column', () => {
    // §20.4.2.3 — a float is not in the text column and may hang into the
    // margins: dml-groupshape-capitalization.docx anchors a 547pt group on a
    // 454pt column, and shrinking it to fit set every word inside at 83%.
    const wide = 6858000; // 540pt, wider than the 468pt column of a letter page
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(anchored(wide, '')), { fonts: FONTS }),
    );
    // The fill path spans the full 540pt.
    expect(text).toMatch(/540 0 l/u);
  });

  it('leaves the paragraph spacing inside a text box on the page', () => {
    // The box's height counted the space between its paragraphs but the
    // emitter never left it, so dml-groupshape-capitalization.docx's caption
    // ran its four paragraphs together.
    const two =
      '<wps:txbx><w:txbxContent>' +
      '<w:p><w:pPr><w:spacing w:after="400"/></w:pPr><w:r><w:t>one</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>two</w:t></w:r></w:p>' +
      '</w:txbxContent></wps:txbx>';
    const text = asLatin1(
      convertDocxToPdfSync(buildDocxFromBody(anchored(2743200, two)), { fonts: FONTS }),
    );
    const ys = [...text.matchAll(/1 0 0 1 [\d.]+ ([\d.]+) Tm/gu)].map((m) => Number(m[1]));
    // Three positions, two of them the box's: the first is the empty line the
    // paragraph the drawing is ANCHORED to stands on (§20.4.2.3 — the drawing
    // leaves the flow, the paragraph mark does not). Nothing is drawn there.
    expect(ys).toHaveLength(3);
    // 20pt of spacing plus the line's own height — well past a bare line.
    expect(ys[1]! - ys[2]!).toBeGreaterThan(25);
  });
});
