// SmartArt. A diagram is four parts of DESCRIPTION — data, layout, quick style,
// colours — that a reader is meant to lay out itself, and nobody outside Office
// does. The producer therefore also writes what it drew, as plain DrawingML
// under `dsp:` (MS-ODRAWXML 2.1), reachable by a `diagramDrawing` relationship.
// tdf83671_SmartArt_import.xlsx draws three nested circles there and we drew
// nothing at all between its "start" and its "end".

import { describe, expect, it } from 'vitest';

import {
  DEFAULT_THEME_PALETTE,
  applyColorMods,
  makeColorResolver,
  readColorMods,
} from '@/core/drawingml/colors';
import { parseDiagramShapes } from '@/excel/sheet-shape-parser';
import { parseXml } from '@/pptx/pptx-reader';

const COLORS = makeColorResolver(DEFAULT_THEME_PALETTE);

const FRAME = { widthPt: 200, heightPt: 200, anchorRow: 2, xPt: 50, yPt: 30 };

/** One `dsp:sp`: an ellipse with a label whose own rectangle sits above it. */
function diagram(body: string): Uint8Array {
  return new TextEncoder().encode(
    '<?xml version="1.0"?><dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree>' +
      body +
      '</dsp:spTree></dsp:drawing>',
  );
}

const SP =
  '<dsp:sp><dsp:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>' +
  '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></dsp:spPr>' +
  '<dsp:txXfrm><a:off x="914400" y="457200"/><a:ext cx="914400" cy="228600"/></dsp:txXfrm>' +
  '<dsp:txBody><a:bodyPr/><a:p><a:r><a:t>back</a:t></a:r></a:p></dsp:txBody></dsp:sp>';

describe('a SmartArt diagram’s fallback drawing', () => {
  const shapes = parseDiagramShapes(diagram(SP), FRAME, COLORS);

  it('draws the graphic at the frame’s corner plus its own offset', () => {
    // 914400 EMU = 72pt, so 50 + 72 = 122pt across and 30 + 0 = 30pt down.
    const graphic = shapes[0];
    expect(graphic?.fill).toEqual({ kind: 'solid', colorHex: 'FFC000' });
    expect(graphic?.float?.posH.offsetPt).toBeCloseTo(122, 3);
    expect(graphic?.float?.posV.offsetPt).toBeCloseTo(30, 3);
    expect(graphic?.width).toBeCloseTo(144, 3);
  });

  it('gives the label the rectangle the diagram sets aside for it', () => {
    // Left on the shape, all three of the file's labels stacked in the centre
    // of its circles instead of sitting one per band.
    const label = shapes[1];
    expect(label?.text?.content).toHaveLength(1);
    expect(label?.float?.posV.offsetPt).toBeCloseTo(30 + 36, 3);
    expect(label?.height).toBeCloseTo(18, 3);
    // The words move; the circle they came from keeps its fill and loses them.
    expect(shapes[0]?.text).toBeUndefined();
  });

  it('ignores a shape that states no size', () => {
    expect(
      parseDiagramShapes(diagram('<dsp:sp><dsp:spPr/></dsp:sp>'), FRAME, COLORS),
    ).toEqual([]);
  });
});

describe('hue and saturation offsets (§20.1.2.3.15)', () => {
  function shifted(xml: string): string {
    const node = parseXml(new TextEncoder().encode(`<a:schemeClr>${xml}</a:schemeClr>`))[0];
    return applyColorMods('FFC000', node ? readColorMods(node) : []);
  }

  it('turns one accent into a diagram’s series of colours', () => {
    // SmartArt is built on it: one colour and a `hueOff` per node. Dropping it
    // drew all three circles the same orange.
    expect(shifted('')).toBe('FFC000');
    // 4900445 sixtieth-thousandths = 81.7° round the wheel: amber → green.
    expect(shifted('<a:hueOff val="4900445"/>')).toMatch(/^[0-9A-F]{6}$/u);
    expect(shifted('<a:hueOff val="4900445"/>')).not.toBe('FFC000');
  });

  it('counts hue in degrees, not in thousandths of a percent', () => {
    // A full 360° turn is the identity; read in the units every OTHER transform
    // uses, the same value collapsed every colour onto hue zero — pure red.
    expect(shifted('<a:hueOff val="21600000"/>')).toBe('FFC000');
  });

  it('offsets saturation as well', () => {
    expect(shifted('<a:satOff val="-40777"/>')).not.toBe('FFC000');
  });
});
