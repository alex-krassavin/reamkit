// §20.5.2.1 `absoluteAnchor` and §20.1.4.1.10 `<xdr:style>`. Both of
// tdf139763ShapeAnchor.xlsx's arrows drew nothing at all on a page LibreOffice
// fills with two blue ones: their colour lives behind a gallery reference, and
// the part is saved with indentation, so the whitespace between `<a:fillRef>`
// and its `<a:schemeClr>` was read as the colour.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { ShapeBlock } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';
import { parseSheetShapes } from '@/excel/sheet-shape-parser';
import { parseWorksheet } from '@/excel/worksheet-parser';

const GALLERY =
  '<xdr:style>\n  <a:lnRef idx="2">\n    <a:schemeClr val="accent1"/>\n  </a:lnRef>\n' +
  '  <a:fillRef idx="1">\n    <a:schemeClr val="accent1"/>\n  </a:fillRef>\n' +
  '  <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>\n' +
  '  <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>\n</xdr:style>';

const SP =
  '<xdr:sp macro="" textlink="">\n  <xdr:nvSpPr><xdr:cNvPr id="2" name="Arrow"/><xdr:cNvSpPr/></xdr:nvSpPr>\n' +
  '  <xdr:spPr>\n    <a:xfrm><a:off x="0" y="0"/><a:ext cx="2293620" cy="579120"/></a:xfrm>\n' +
  '    <a:prstGeom prst="leftArrow"><a:avLst/></a:prstGeom>\n  </xdr:spPr>\n' +
  `  ${GALLERY}\n  <xdr:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></xdr:txBody>\n</xdr:sp>`;

function shapeOf(sheetShape: {
  rawShapeXml?: string;
  rawAnchorXml?: string;
}): ShapeBlock | undefined {
  const { flow } = Ream.parse(buildXlsx({ rows: [['a']], sheetShape }));
  const el = flow.body.find((b) => b.kind === 'shape');
  return el?.kind === 'shape' ? el.shape : undefined;
}

describe('a gallery-styled sheet shape', () => {
  it('finds the colour past the whitespace a pretty-printed part puts first', () => {
    const shape = shapeOf({ rawShapeXml: SP });
    expect(shape?.fill.kind).toBe('solid');
    expect(shape?.line?.fill).not.toBe('none');
  });
});

describe('absoluteAnchor (§20.5.2.1)', () => {
  const SHEET = parseWorksheet(
    new TextEncoder().encode(
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    ),
  );

  /** The shape's corner, in points from the sheet's own. */
  function cornerOf(pos: string): { x: number; y: number } | undefined {
    const xml = new TextEncoder().encode(
      '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
        `<xdr:absoluteAnchor>${pos}<xdr:ext cx="2880000" cy="720000"/>${SP}</xdr:absoluteAnchor></xdr:wsDr>`,
    );
    const shape = parseSheetShapes(xml, SHEET, () => '4472C4')[0];
    const f = shape?.float;
    return f ? { x: f.posH?.offsetPt ?? 0, y: f.posV?.offsetPt ?? 0 } : undefined;
  }

  it('places the shape at the offset it names', () => {
    // An absolute anchor names no cell: `<xdr:pos>` IS the placement, and read
    // nowhere both of tdf139763ShapeAnchor.xlsx's arrows piled into the corner.
    // 1440000 × 1080000 EMU = 113.39pt × 85.04pt.
    expect(cornerOf('<xdr:pos x="1440000" y="1080000"/>')?.x).toBeCloseTo(113.39, 1);
    expect(cornerOf('<xdr:pos x="1440000" y="1080000"/>')?.y).toBeCloseTo(85.04, 1);
  });

  it('leaves an anchor that names none at the corner', () => {
    expect(cornerOf('')).toEqual({ x: 0, y: 0 });
  });
});
