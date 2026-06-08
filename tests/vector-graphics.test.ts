import { describe, expect, it } from 'vitest';

import type { VectorShape } from '@/pdf/vector-graphics';
import { PathBuilder, emitVectorShape } from '@/pdf/vector-graphics';

// A 100×50 rectangle path in a local y-up frame.
function rectPath() {
  return new PathBuilder().moveTo(0, 0).lineTo(100, 0).lineTo(100, 50).lineTo(0, 50).close();
}

describe('emitVectorShape', () => {
  it('emits a filled + stroked rectangle as q…cm…path…B…Q', () => {
    const shape: VectorShape = {
      paths: [rectPath().build()],
      fillColorHex: '4472C4',
      stroke: { colorHex: '2F528F', widthPt: 1 },
      transform: [1, 0, 0, 1, 72, 200],
    };
    const ops = emitVectorShape(shape);
    const text = ops.join('\n');

    expect(ops[0]).toBe('q');
    expect(ops[ops.length - 1]).toBe('Q');
    expect(text).toContain('1 0 0 1 72 200 cm');
    expect(text).toContain('1 w');
    // path construction
    expect(text).toContain('0 0 m');
    expect(text).toContain('100 0 l');
    expect(text).toContain('100 50 l');
    expect(text).toContain('0 50 l');
    expect(ops).toContain('h');
    // colours: a non-stroking (rg) and a stroking (RG) set
    expect(ops.some((o) => o.endsWith(' rg'))).toBe(true);
    expect(ops.some((o) => o.endsWith(' RG'))).toBe(true);
    // fill + stroke ⇒ B
    expect(ops).toContain('B');
    expect(ops).not.toContain('f');
    expect(ops).not.toContain('S');
  });

  it('fill-only emits f and no stroke state', () => {
    const shape: VectorShape = {
      paths: [rectPath().build()],
      fillColorHex: 'FF0000',
      transform: [1, 0, 0, 1, 0, 0],
    };
    const ops = emitVectorShape(shape);
    expect(ops).toContain('f');
    expect(ops).not.toContain('B');
    expect(ops.some((o) => o.endsWith(' RG'))).toBe(false);
    expect(ops.some((o) => o.endsWith(' w'))).toBe(false);
  });

  it('stroke-only emits S and no fill colour', () => {
    const shape: VectorShape = {
      paths: [rectPath().build()],
      stroke: { colorHex: '000000', widthPt: 2 },
      transform: [1, 0, 0, 1, 0, 0],
    };
    const ops = emitVectorShape(shape);
    expect(ops).toContain('S');
    expect(ops).not.toContain('B');
    expect(ops).not.toContain('f');
    expect(ops.some((o) => o.endsWith(' rg'))).toBe(false);
    expect(ops).toContain('2 w');
  });

  it('evenodd fill rule selects the * painting variant', () => {
    const shape: VectorShape = {
      paths: [rectPath().build('evenodd')],
      fillColorHex: '00FF00',
      transform: [1, 0, 0, 1, 0, 0],
    };
    expect(emitVectorShape(shape)).toContain('f*');
  });

  it('emits a cubic Bézier as a c operator', () => {
    const shape: VectorShape = {
      paths: [new PathBuilder().moveTo(0, 0).cubicTo(10, 20, 30, 40, 50, 0).build()],
      stroke: { colorHex: '000000', widthPt: 1 },
      transform: [1, 0, 0, 1, 0, 0],
    };
    expect(emitVectorShape(shape)).toContain('10 20 30 40 50 0 c');
  });

  it('emits dash, cap and join state when set', () => {
    const shape: VectorShape = {
      paths: [rectPath().build()],
      stroke: { colorHex: '000000', widthPt: 1, dash: [3, 2], cap: 'round', join: 'bevel' },
      transform: [1, 0, 0, 1, 0, 0],
    };
    const ops = emitVectorShape(shape);
    expect(ops).toContain('[3 2] 0 d');
    expect(ops).toContain('1 J'); // round cap
    expect(ops).toContain('2 j'); // bevel join
  });

  it('no fill and no stroke ends the path with n (no-op paint)', () => {
    const shape: VectorShape = {
      paths: [rectPath().build()],
      transform: [1, 0, 0, 1, 0, 0],
    };
    expect(emitVectorShape(shape)).toContain('n');
  });
});
