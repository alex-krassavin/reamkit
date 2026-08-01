// §20.1.7.6 `a:xfrm` on a sheet shape, and §20.1.9.18's block-arrow geometry.
// tdf135828_Shape_Rect.xlsx turns an `upArrow` 76.9° with `rot="4616172"`: we
// read the rotation nowhere and drew the arrow lying flat across the page.

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { ShapeBlock } from '@/core/document-model';
import { presetPaths } from '@/core/drawingml/preset-geometry';
import { Ream } from '@/core/converter/ream';

/** The `a:ext` of tdf135828_Shape_Rect.xlsx's arrow: 23.25pt × 156pt. */
const EXT = '<a:ext cx="295275" cy="1981200"/>';

function shapeWith(xfrm: string): ShapeBlock | undefined {
  const { flow } = Ream.parse(
    buildXlsx({
      rows: [['a']],
      sheetShape: {
        rawShapeXml:
          '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="arrow"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
          `<xdr:spPr>${xfrm}<a:prstGeom prst="upArrow"><a:avLst/></a:prstGeom>` +
          '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></xdr:spPr></xdr:sp>',
      },
    }),
  );
  const el = flow.body.find((b) => b.kind === 'shape');
  return el?.kind === 'shape' ? el.shape : undefined;
}

describe('a turned sheet shape (§20.1.7.6)', () => {
  it('reads the rotation and the flips', () => {
    expect(shapeWith(`<a:xfrm rot="4616172" flipV="1">${EXT}</a:xfrm>`)?.transform).toEqual({
      rotation60k: 4616172,
      flipV: true,
    });
  });

  it('takes its size from a:ext, which the anchor no longer describes', () => {
    // The anchor spans the ground the TURNED arrow covers, a different
    // rectangle entirely; `a:ext` is the arrow itself.
    const shape = shapeWith(`<a:xfrm rot="4616172">${EXT}</a:xfrm>`);
    expect(shape?.width).toBeCloseTo(23.25, 2);
    expect(shape?.height).toBeCloseTo(156, 2);
  });

  it('leaves an unturned shape on its anchor', () => {
    const shape = shapeWith(`<a:xfrm>${EXT}</a:xfrm>`);
    expect(shape?.transform).toBeUndefined();
    expect(shape?.height).not.toBeCloseTo(156, 2);
  });
});

describe('block arrow geometry (§20.1.9.18)', () => {
  /** Every point the path visits. */
  function points(w: number, h: number): Array<{ x: number; y: number }> {
    const path = presetPaths('upArrow', w, h, new Map())?.[0];
    return (path?.segments ?? []).flatMap((seg) =>
      'x' in seg && 'y' in seg ? [{ x: seg.x, y: seg.y }] : [],
    );
  }

  it('measures the head against the SHORTER side of the box', () => {
    // 20 wide, 200 tall. At the default 50 % the head is half of 20, not half
    // of 200 — a compact triangle rather than a spike down two thirds of the
    // shaft. The shaft's far edge therefore sits 10pt short of the tip.
    const ys = points(20, 200).map((p) => p.y);
    expect(Math.max(...ys)).toBeCloseTo(200, 5);
    expect(ys.some((y) => Math.abs(y - 190) < 1e-6)).toBe(true);
  });

  it('measures the shaft against it too', () => {
    // 10pt across, centred in a 20pt box: edges at 5 and 15.
    const xs = points(20, 200).map((p) => p.x);
    expect(xs.some((x) => Math.abs(x - 5) < 1e-6)).toBe(true);
    expect(xs.some((x) => Math.abs(x - 15) < 1e-6)).toBe(true);
  });
});
