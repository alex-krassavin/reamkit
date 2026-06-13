import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
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
  it('shade darkens toward black', () => {
    expect(applyColorMods('4472C4', [{ kind: 'shade', val: 0.5 }])).toBe('223962');
  });

  it('tint lightens toward white', () => {
    expect(applyColorMods('000000', [{ kind: 'tint', val: 0.5 }])).toBe('808080');
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

  it('drops a shape in a mixed text+shape run but keeps the text', () => {
    const inner = `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>`;
    const body = `<w:p><w:r><w:t>Hello</w:t></w:r>${shapeRun(inner)}</w:p>`;
    const docx = buildDocxFromBody(body);
    const parsed = parseDocument(OpcPackage.open(docx).getMainDocument().data);
    expect(parsed[0]!.kind).toBe('paragraph'); // not collapsed to a shape
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // The shape fill (4472C4) is dropped; only the text renders.
    expect(text).not.toContain('0.266667 0.447059 0.768627 rg');
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
    const hexOf = (s: string): string =>
      [...s]
        .map((c) => parsed.glyphForCodepoint(c.codePointAt(0)!))
        .map((g) => g.toString(16).padStart(4, '0').toUpperCase())
        .join('');

    const labelTj = `<${hexOf('Label')}> Tj`;
    expect(text).toContain(labelTj); // text-box content rendered
    // Shape fill paints (f) before the text pass (BT) → text on top.
    expect(text.indexOf('\nf\n')).toBeLessThan(text.indexOf('BT'));
    expect(text.indexOf('BT')).toBeLessThan(text.indexOf(labelTj));
  });
});
