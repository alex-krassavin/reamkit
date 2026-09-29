// §9.10.2 — punctuation read off the shape a glyph draws, for a font that
// states no character for it (Stripe writes every hyphen, colon, dash and
// bracket of its invoices as U+0000).

import { describe, expect, it } from 'vitest';

import type { GlyphShape } from '@/pdf-reader/glyph-shapes';
import type { PathSeg } from '@/pdf-reader/content';
import { punctuationOf } from '@/pdf-reader/glyph-shapes';

/** A closed box, as one contour, in em units at size 1 on a baseline at 0. */
function box(x0: number, y0: number, x1: number, y1: number): Array<PathSeg> {
  return [
    { op: 'move', x: x0, y: y0 },
    { op: 'line', x: x1, y: y0 },
    { op: 'line', x: x1, y: y1 },
    { op: 'line', x: x0, y: y1 },
    { op: 'close' },
  ];
}

function shape(...contours: Array<Array<PathSeg>>): GlyphShape {
  const segs = contours.flat();
  // The box a curve fills reaches toward its control points, as a drawn
  // glyph's does.
  const points = segs.flatMap((s) =>
    s.op === 'close'
      ? []
      : s.op === 'cubic'
        ? [
            { x: s.x1, y: s.y1 },
            { x: s.x2, y: s.y2 },
            { x: s.x, y: s.y },
          ]
        : [{ x: s.x, y: s.y }],
  );
  return {
    segs,
    minX: Math.min(...points.map((p) => p.x)),
    minY: Math.min(...points.map((p) => p.y)),
    maxX: Math.max(...points.map((p) => p.x)),
    maxY: Math.max(...points.map((p) => p.y)),
  };
}

/** A bracket's curve as a crescent: its ends at the top and bottom on `endX`, its belly at `bellyX`. */
function bracket(endX: number, bellyX: number): GlyphShape {
  return shape([
    { op: 'move', x: endX, y: 0.82 },
    { op: 'cubic', x1: bellyX, y1: 0.6, x2: bellyX, y2: 0.1, x: endX, y: -0.09 },
    { op: 'line', x: endX + (endX > bellyX ? -0.05 : 0.05), y: -0.09 },
    { op: 'cubic', x1: bellyX + 0.05, y1: 0.1, x2: bellyX + 0.05, y2: 0.6, x: endX, y: 0.8 },
    { op: 'close' },
  ]);
}

describe('punctuation a glyph draws (§9.10.2)', () => {
  it('reads a flat bar at the middle of the lower case as a dash, by its length', () => {
    // Inter's own: a hyphen 0.33 em long, an en dash half an em.
    expect(punctuationOf(shape(box(0.07, 0.31, 0.4, 0.41)), 0, 1)).toBe('-');
    expect(punctuationOf(shape(box(0, 0.32, 0.5, 0.4)), 0, 1)).toBe('–');
    expect(punctuationOf(shape(box(0, 0.32, 1, 0.4)), 0, 1)).toBe('—');
  });

  it('reads two dots stacked on the baseline as a colon, and a hanging lower one as a semicolon', () => {
    expect(punctuationOf(shape(box(0.07, 0.09, 0.2, 0.2), box(0.07, 0.48, 0.2, 0.59)), 0, 1)).toBe(
      ':',
    );
    expect(punctuationOf(shape(box(0.07, -0.15, 0.2, 0.1), box(0.07, 0.48, 0.2, 0.59)), 0, 1)).toBe(
      ';',
    );
  });

  it('reads one small mark on the baseline as a full stop, and one hanging below it as a comma', () => {
    expect(punctuationOf(shape(box(0.07, 0, 0.19, 0.12)), 0, 1)).toBe('.');
    expect(punctuationOf(shape(box(0.07, -0.15, 0.19, 0.12)), 0, 1)).toBe(',');
  });

  it('tells an opening bracket from a closing one by which way its ends turn', () => {
    // "(" bulges LEFT, so its ends stand right of its middle.
    expect(punctuationOf(bracket(0.25, 0.04), 0, 1)).toBe('(');
    expect(punctuationOf(bracket(0.04, 0.25), 0, 1)).toBe(')');
  });

  it('measures in the size the line is set at, from the line’s own baseline', () => {
    // The same hyphen at 9pt on a baseline at y=700.
    const at = (y: number, s: number): GlyphShape => ({
      ...shape(box(0.07 * s, y + 0.31 * s, 0.4 * s, y + 0.41 * s)),
    });
    expect(punctuationOf(at(700, 9), 700, 9)).toBe('-');
  });

  it('leaves alone a shape that could be more than one thing', () => {
    // A tall box is a letter (an l, an I, a 1) as much as it is anything.
    expect(punctuationOf(shape(box(0.1, 0, 0.2, 0.72)), 0, 1)).toBeUndefined();
    // A bar standing ON the baseline is an underscore or a rule, not a dash.
    expect(punctuationOf(shape(box(0, -0.1, 0.5, -0.02)), 0, 1)).toBeUndefined();
  });
});
