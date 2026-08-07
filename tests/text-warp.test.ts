// §20.1.9.10 `a:prstTxWarp` — the preset warps a WordArt body is bent through.
//
// Every constant these presets carry was read off a rendered envelope: a box of
// capitals set in each warp, the top and bottom edge of the ink measured column
// by column. The tests below hold the shape of those envelopes — that a can
// swells only at its ends, that a wave is one sine period and a double wave
// two, that a ring winds once round the box — so a change to a curve has to be
// a deliberate one.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildPptx } from './fixtures/build-pptx';
import type { WarpFrame } from '@/core/drawingml/text-warp';
import { Ream } from '@/core/converter/ream';
import { FontRegistry, createFontMeasure, parseTtf } from '@/core/font';
import { flowRenderOptions } from '@/core/converter/project';
import { layoutStyledDocument } from '@/layout/styled-layout';
import { isTextWarp, textWarpEdges, warpGlyphMatrix } from '@/core/drawingml/text-warp';

/** Both edges at `u`, as a `[top, bottom]` pair rounded to three places. */
function at(preset: string, u: number, adj?: number): [number, number] {
  const e = textWarpEdges(preset, u, adj);
  if (!e) throw new Error(`no envelope for ${preset}`);
  return [Number(e.top.toFixed(3)), Number(e.bottom.toFixed(3))];
}

describe('preset text warps', () => {
  it('does not treat the enumeration\'s "no warp" member as one', () => {
    // Two thirds of the decks that mention a warp at all state only this: a
    // body under it is an ordinary text box that wraps and is not stretched.
    expect(isTextWarp('textNoShape')).toBe(false);
    expect(textWarpEdges('textNoShape', 0.5)).toBeUndefined();
    expect(isTextWarp('textWave1')).toBe(true);
  });

  it('leaves a preset it does not know unbent', () => {
    expect(isTextWarp('textArchUpPour')).toBe(false);
    expect(textWarpEdges('textNotAWarp', 0.5)).toBeUndefined();
  });

  it('fills the box: every envelope touches both its edges', () => {
    const presets = [
      'textPlain',
      'textStop',
      'textTriangle',
      'textTriangleInverted',
      'textChevron',
      'textChevronInverted',
      'textCanUp',
      'textCanDown',
      'textWave1',
      'textWave2',
      'textDoubleWave1',
      'textWave4',
      'textInflate',
      'textDeflate',
      'textInflateTop',
      'textInflateBottom',
      'textDeflateTop',
      'textDeflateBottom',
      'textFadeRight',
      'textFadeLeft',
      'textFadeUp',
      'textFadeDown',
      'textSlantUp',
      'textSlantDown',
      'textCascadeUp',
      'textCascadeDown',
      'textCurveUp',
      'textCurveDown',
    ];
    for (const preset of presets) {
      let top = Infinity;
      let bottom = -Infinity;
      for (let i = 0; i <= 200; i++) {
        const e = textWarpEdges(preset, i / 200);
        if (!e) throw new Error(`no envelope for ${preset}`);
        top = Math.min(top, e.top);
        bottom = Math.max(bottom, e.bottom);
        expect(e.bottom).toBeGreaterThanOrEqual(e.top - 1e-9);
      }
      expect(top).toBeCloseTo(0, 2);
      expect(bottom).toBeCloseTo(1, 2);
    }
  });

  it('rides a wave through one sine period and a double wave through two', () => {
    // The band keeps its height and only its position moves: a wave is a
    // TRANSLATION, which is why the letters stay their own size along it.
    const height = (u: number): number => at('textWave1', u)[1] - at('textWave1', u)[0];
    expect(height(0)).toBeCloseTo(height(0.37), 3);
    expect(height(0.62)).toBeCloseTo(height(0.85), 3);
    // Trough a quarter in, crest three quarters in, level at both ends.
    expect(at('textWave1', 0.25)[0]).toBeCloseTo(0, 3);
    expect(at('textWave1', 0.75)[1]).toBeCloseTo(1, 3);
    expect(at('textWave1', 0)[0]).toBeCloseTo(at('textWave1', 0.5)[0], 2);
    // `textWave2` is the same curve read the other way up.
    expect(at('textWave2', 0.25)[1]).toBeCloseTo(1, 3);
    expect(at('textWave2', 0.75)[0]).toBeCloseTo(0, 3);
    // Two periods: the double wave troughs an eighth in, not a quarter.
    expect(at('textDoubleWave1', 0.125)[0]).toBeCloseTo(0, 3);
    expect(at('textDoubleWave1', 0.625)[0]).toBeCloseTo(0, 3);
    expect(at('textWave4', 0.125)[1]).toBeCloseTo(1, 3);
  });

  it('swells a can only towards its ends', () => {
    // The measured curve is flat across the middle and lifts sharply at the
    // edges — a quarter of the way out it has moved a hundredth of the box.
    const [midTop] = at('textCanUp', 0.5);
    expect(midTop).toBeCloseTo(0, 3);
    expect(at('textCanUp', 0.375)[0]).toBeLessThan(0.01);
    expect(at('textCanUp', 0)[0]).toBeCloseTo(0.14, 2);
    // Down is the same swell hanging from the other edge.
    expect(at('textCanDown', 0.5)[1]).toBeCloseTo(1, 3);
    expect(at('textCanDown', 0)[1]).toBeCloseTo(0.86, 2);
  });

  it('tapers the straight-sided presets in a straight line', () => {
    // A triangle's baseline is flat and its top rises to a point; a chevron
    // keeps its height and peaks in the middle; both are linear in the
    // distance from the centre, so half way out is half the taper.
    expect(at('textTriangle', 0.25)).toEqual([0.25, 1]);
    expect(at('textTriangle', 0.5)).toEqual([0, 1]);
    expect(at('textTriangleInverted', 0.25)).toEqual([0, 0.75]);
    expect(at('textChevron', 0.5)).toEqual([0, 0.75]);
    expect(at('textChevron', 0)).toEqual([0.25, 1]);
    expect(at('textChevronInverted', 0.5)).toEqual([0.25, 1]);
  });

  it('reads the adjustment a file states, and its own default when it does not', () => {
    // §20.1.9.10 — `a:avLst`'s `adj` in hundred-thousandths. A deflate pinches
    // its middle by twice the guide, so 50 000 closes it completely.
    expect(at('textDeflate', 0.5)).toEqual([0.375, 0.625]);
    expect(at('textDeflate', 0.5, 25000)).toEqual([0.25, 0.75]);
    expect(at('textDeflate', 0.5, 50000)).toEqual([0.5, 0.5]);
    // An inflate-top lifts its top edge by the guide at the box's ends.
    expect(at('textInflateTop', 0, 50000)).toEqual([0.5, 1]);
    expect(at('textInflateTop', 0, 25000)).toEqual([0.25, 1]);
  });
});

/** A frame whose block and box are the same 400×100 rectangle at the origin. */
const FRAME: WarpFrame = {
  preset: 'textPlain',
  boxX: 0,
  boxY: 0,
  boxWidth: 400,
  boxHeight: 100,
  srcX: 0,
  srcWidth: 400,
  srcTop: 0,
  srcHeight: 100,
};

describe('warped glyph placement', () => {
  it('leaves a glyph alone when the warp is the plain stretch', () => {
    const m = warpGlyphMatrix(FRAME, 100, 20, 100);
    // Block and box are the same rectangle, so nothing scales; the baseline is
    // the block's bottom edge and the glyph's y runs against the page's.
    expect(m).toBeDefined();
    expect(m![0]).toBeCloseTo(1, 6);
    expect(m![1]).toBeCloseTo(0, 6);
    expect(m![2]).toBeCloseTo(0, 6);
    expect(m![3]).toBeCloseTo(-1, 6);
    expect(m![4]).toBeCloseTo(100, 6);
    expect(m![5]).toBeCloseTo(100, 6);
  });

  it('stretches a block onto a box it does not match', () => {
    const m = warpGlyphMatrix({ ...FRAME, boxWidth: 800, boxHeight: 300 }, 200, 20, 50);
    expect(m![0]).toBeCloseTo(2, 6); // 800 / 400
    expect(m![3]).toBeCloseTo(-3, 6); // 300 / 100
    expect(m![4]).toBeCloseTo(400, 6);
    expect(m![5]).toBeCloseTo(150, 6);
  });

  it('leans a glyph along the slope the curve runs at', () => {
    // Half way up a wave's rise the baseline is climbing, so the glyph is
    // sheared with it: a rising curve gives a NEGATIVE page-y slope.
    const rising = warpGlyphMatrix({ ...FRAME, preset: 'textWave1' }, 0, 10, 50);
    expect(rising![1]).toBeLessThan(0);
    // At the trough and the crest the curve is level and the glyph is upright.
    const trough = warpGlyphMatrix({ ...FRAME, preset: 'textWave1' }, 100, 0.01, 50);
    expect(trough![1]).toBeCloseTo(0, 2);
  });

  it('drops a glyph the warp leaves no room for', () => {
    // `textFadeUp` closes its band completely at the box's ends.
    const gone = warpGlyphMatrix({ ...FRAME, preset: 'textFadeUp' }, 400, 0, 50);
    expect(gone).toBeUndefined();
  });

  it('winds a ring once round the box, opening at its left edge', () => {
    const ring: WarpFrame = { ...FRAME, preset: 'textRingInside' };
    // The first glyph starts at the box's left edge, half way down it.
    const first = warpGlyphMatrix(ring, 0, 10, 0)!;
    expect(first[4]).toBeCloseTo(0, 6);
    expect(first[5]).toBeCloseTo(50, 6);
    // A quarter of the way round it is at the top, upright and running right.
    const quarter = warpGlyphMatrix(ring, 100, 10, 0)!;
    expect(quarter[4]).toBeCloseTo(200, 6);
    expect(quarter[5]).toBeCloseTo(0, 6);
    expect(quarter[0]).toBeGreaterThan(0);
    expect(quarter[3]).toBeLessThan(0);
    // Three quarters round it is at the bottom, upside down: it runs the other
    // way and its own y now runs WITH the page's.
    const threeQuarters = warpGlyphMatrix(ring, 300, 10, 0)!;
    expect(threeQuarters[4]).toBeCloseTo(200, 6);
    expect(threeQuarters[5]).toBeCloseTo(100, 6);
    expect(threeQuarters[0]).toBeLessThan(0);
    expect(threeQuarters[3]).toBeGreaterThan(0);
  });

  it('turns a ring the other way, and its letters the other way up', () => {
    // `textRingOutside` dips under the box first and points its letters into
    // the ring, so the quarter turn that was the top is now the bottom — and
    // the baseline that sat on the ring's OUTER edge is now the block's last
    // row, not its first.
    const ring: WarpFrame = { ...FRAME, preset: 'textRingOutside' };
    const quarter = warpGlyphMatrix(ring, 100, 10, 100)!;
    expect(quarter[5]).toBeCloseTo(100, 6);
    expect(quarter[0]).toBeGreaterThan(0);
    expect(quarter[3]).toBeLessThan(0);
  });

  it('holds a ring band at a constant depth all the way round', () => {
    // A scaled annulus would be four times deeper at the sides of a box four
    // times wider than tall, and the letters there came out as spikes.
    const ring: WarpFrame = { ...FRAME, preset: 'textRingInside' };
    const top = warpGlyphMatrix(ring, 100, 10, 0)!;
    const side = warpGlyphMatrix(ring, 0, 10, 0)!;
    expect(Math.hypot(side[2], side[3])).toBeCloseTo(Math.hypot(top[2], top[3]), 6);
  });
});

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

/**
 * One slide holding one WordArt shape, laid out.
 *
 * The box is 4 000 000 × 1 000 000 EMU = 315 × 79 pt, and the text is far too
 * wide for it at 54 pt: an ordinary body would break it over three lines.
 */
function warpedPage(prst: string, text = 'Text Wave One Two') {
  const sp =
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="W"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0">` +
    (prst === '' ? '' : `<a:prstTxWarp prst="${prst}"><a:avLst/></a:prstTxWarp>`) +
    `</a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="5400" b="1"/>` +
    `<a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const flow = Ream.parse(buildPptx([sp])).flow;
  const laid = layoutStyledDocument(flow.body, {
    registry: FontRegistry.fromBytes(FONTS),
    ...flowRenderOptions(flow),
  });
  // A slide always carries the zero-width paragraph that stands for its body;
  // only the shape's own words are of interest here.
  return laid.pages[0]!.commands.filter(
    (c) =>
      c.type === 'line' &&
      c.line.tokens.some((t) => t.kind === 'text' && t.text.trim() !== '' && t.text !== '​'),
  );
}

describe('a WordArt body through the layout', () => {
  it('sets the whole paragraph on one line, whatever the box is wide', () => {
    // §20.1.9.10 — warped text does not wrap: it is bent as one line and then
    // stretched onto the box. Broken at the box's width first, tdf114848's
    // "Text Wave 1" came out as three stacked lines the warp then bent as a
    // block.
    const bent = warpedPage('textWave1');
    expect(bent).toHaveLength(1);
    // The same body without a warp is the ordinary box it always was.
    const flat = warpedPage('');
    expect(flat.length).toBeGreaterThan(1);
  });

  it('hands each line the two frames the curve maps between', () => {
    const [line] = warpedPage('textWave1');
    if (line?.type !== 'line') throw new Error('no line');
    const warp = line.warp;
    expect(warp?.preset).toBe('textWave1');
    // The box is the shape's, insets and all: 4 000 000 EMU = 315pt across.
    expect(warp?.boxWidth).toBeCloseTo(315, 0);
    expect(warp?.boxHeight).toBeCloseTo(78.7, 0);
    // The block is the un-warped line — far wider than the box it goes onto.
    expect(warp!.srcWidth).toBeGreaterThan(warp!.boxWidth);
    // …and its height is the INK's, not the line box's: a line of capitals
    // with no descender reaches from its cap height to its baseline and no
    // further, and stretching the empty band below it would set the whole
    // block a third too small.
    expect(warp!.srcHeight).toBeLessThan(54);
    expect(warp!.srcHeight).toBeGreaterThan(30);
  });

  it('places every glyph on its own through the emitter', async () => {
    // Each letter sits at its own point on the curve, so each carries its own
    // text matrix — and on a wave most of them are sheared, which a flat line
    // never is.
    const pdf = await Ream.parse(
      buildPptx([
        `<p:sp><p:nvSpPr><p:cNvPr id="2" name="W"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></a:xfrm>` +
          `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
          `<p:txBody><a:bodyPr wrap="none">` +
          `<a:prstTxWarp prst="textWave1"><a:avLst/></a:prstTxWarp></a:bodyPr>` +
          `<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="5400" b="1"/>` +
          `<a:t>Text Wave 1</a:t></a:r></a:p></p:txBody></p:sp>`,
      ]),
    ).convert('pdf');
    const stream = Buffer.from(pdf).toString('latin1');
    const matrices = [
      ...stream.matchAll(/(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) Tm/gu),
    ];
    // One per glyph of "Text Wave 1", spaces aside — nine at the least.
    expect(matrices.length).toBeGreaterThanOrEqual(9);
    // The wave leans them: several carry a `b` term no flat line would.
    expect(matrices.filter((m) => Math.abs(Number(m[2])) > 0.05).length).toBeGreaterThanOrEqual(4);
  });
});

describe('ink measurement', () => {
  it('measures what a string DRAWS, not the band its font reserves', () => {
    // WordArt stretches its marks onto its shape, so the empty space an
    // ascender or a descender reserves must not be stretched with them.
    const measure = createFontMeasure(parseTtf(FONTS.bold));
    const caps = measure.textInkPt('HAND', 100);
    // Capitals reach a cap height and stop dead on the baseline.
    expect(caps.below).toBeCloseTo(0, 3);
    expect(caps.above).toBeGreaterThan(60);
    expect(caps.above).toBeLessThan(80);
    // A descender drops below it, and a round letter overshoots the cap line
    // by a hair — both of which the line box would have hidden.
    const tails = measure.textInkPt('gjpqy', 100);
    expect(tails.below).toBeGreaterThan(10);
    // A string with no ink at all still has to answer with something usable.
    const blank = measure.textInkPt('   ', 100);
    expect(blank.above).toBeGreaterThan(0);
  });
});
