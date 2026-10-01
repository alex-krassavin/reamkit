// ISO/IEC 14496-22 — a TrueType font built from outlines, as a document that
// embeds its faces needs one (see `src/core/font/ttf-build.ts`).

import { describe, expect, it } from 'vitest';

import type { BuiltFace, GlyphSeg } from '@/core/font';
import { buildTrueType, editableEmbedding, parseTtf } from '@/core/font';
import { outlineSource } from '@/pdf-reader/glyf-outline';

/** A 400 × 700 box standing 100 units in, drawn counter-clockwise as PostScript draws ink. */
const BOX: ReadonlyArray<GlyphSeg> = [
  { op: 'move', x: 0.1, y: 0 },
  { op: 'line', x: 0.5, y: 0 },
  { op: 'line', x: 0.5, y: 0.7 },
  { op: 'line', x: 0.1, y: 0.7 },
  { op: 'close' },
];

/** A ring of four cubic quarter-circles — no quadratic draws any of them exactly. */
function circle(cx: number, cy: number, r: number): Array<GlyphSeg> {
  const k = 0.5523 * r;
  return [
    { op: 'move', x: cx + r, y: cy },
    { op: 'cubic', x1: cx + r, y1: cy + k, x2: cx + k, y2: cy + r, x: cx, y: cy + r },
    { op: 'cubic', x1: cx - k, y1: cy + r, x2: cx - r, y2: cy + k, x: cx - r, y: cy },
    { op: 'cubic', x1: cx - r, y1: cy - k, x2: cx - k, y2: cy - r, x: cx, y: cy - r },
    { op: 'cubic', x1: cx + k, y1: cy - r, x2: cx + r, y2: cy - k, x: cx + r, y: cy },
    { op: 'close' },
  ];
}

/** A quadratic arch raised to a cubic, the way a TrueType outline reaches the reader. */
const ARCH: ReadonlyArray<GlyphSeg> = [
  { op: 'move', x: 0, y: 0 },
  // Raised from the quadratic through (0, 0.6): each control two thirds of the
  // way from its end towards that point.
  { op: 'cubic', x1: 0, y1: 0.4, x2: 0.5 / 3, y2: 0.6, x: 0.5, y: 0.6 },
  { op: 'line', x: 0.5, y: 0 },
  { op: 'close' },
];

function build(fsType = 8, kerning: BuiltFace['kerning'] = undefined): Uint8Array {
  return buildTrueType({
    ...(kerning ? { kerning } : {}),
    family: 'Test Face',
    bold: true,
    italic: false,
    postScriptName: 'TestFace-Bold',
    glyphs: [
      { codePoints: [0x20, 0xa0], outline: [], advance: 250 },
      { codePoints: [0x41], outline: BOX, advance: 600 },
      { codePoints: [0x4f], outline: circle(0.35, 0.35, 0.3), advance: 700 },
      { codePoints: [0x61], outline: ARCH, advance: 520 },
      { codePoints: [0x10c80], outline: BOX, advance: 900 },
    ],
    ascent: 800,
    descent: -200,
    fsType,
  });
}

/** Where a table starts in an sfnt, by tag. */
function tableAt(font: Uint8Array, tag: string): number {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  for (let i = 0; i < view.getUint16(4); i++) {
    const at = 12 + i * 16;
    if (String.fromCharCode(...font.subarray(at, at + 4)) === tag) return view.getUint32(at + 8);
  }
  throw new Error(`no ${tag} table`);
}

/** The points a glyph's contours pass through, in font units. */
function pointsOf(path: ReadonlyArray<GlyphSeg>): Array<[number, number]> {
  return path.flatMap(
    (seg): Array<[number, number]> =>
      seg.op === 'close' ? [] : [[Math.round(seg.x * 1000), Math.round(seg.y * 1000)]],
  );
}

describe('a TrueType font built from outlines', () => {
  it('reaches each glyph through its characters, the BMP and beyond it', () => {
    const font = parseTtf(build());
    expect(font.unitsPerEm).toBe(1000);
    expect(font.numGlyphs).toBe(6); // .notdef and five
    expect(font.glyphForCodepoint(0x20)).toBe(1);
    expect(font.glyphForCodepoint(0xa0)).toBe(1); // the no-break space shares the space
    expect(font.glyphForCodepoint(0x41)).toBe(2);
    expect(font.glyphForCodepoint(0x4f)).toBe(3);
    expect(font.glyphForCodepoint(0x10c80)).toBe(5);
    expect(font.glyphForCodepoint(0x42)).toBe(0);
  });

  it('keeps the advances the page set the glyphs with', () => {
    const font = parseTtf(build());
    expect([...font.advanceWidths]).toEqual([0, 250, 600, 700, 520, 900]);
  });

  it('turns a PostScript contour to run clockwise, as TrueType ink does', () => {
    const glyphs = outlineSource(build())!;
    // The box began at its lower left going right; clockwise it goes up first.
    expect(pointsOf(glyphs.path(2)!).slice(0, 4)).toEqual([
      [100, 700],
      [500, 700],
      [500, 0],
      [100, 0],
    ]);
  });

  it('gives back a raised quadratic as the quadratic it was', () => {
    const arch = outlineSource(build())!.path(4)!;
    const curves = arch.filter((seg) => seg.op === 'cubic');
    // One curve, not a run of approximations, through the same control.
    expect(curves).toHaveLength(1);
    const curve = curves[0]!;
    expect([curve.x1, curve.y1, curve.x2, curve.y2]).toEqual(
      [0, 0.4, 0.5 / 3, 0.6].map((v) => expect.closeTo(v, 6)),
    );
  });

  it('follows a true cubic to within a unit', () => {
    const ring = outlineSource(build())!.path(3)!;
    for (const [x, y] of pointsOf(ring)) {
      // Every point on the curve, and every control near it, stays on the circle.
      expect(Math.abs(Math.hypot(x - 350, y - 350) - 300)).toBeLessThan(40);
    }
    const onCurve = ring.flatMap((seg) => (seg.op === 'cubic' ? [[seg.x, seg.y] as const] : []));
    for (const [x, y] of onCurve) {
      expect(Math.abs(Math.hypot(x * 1000 - 350, y * 1000 - 350) - 300)).toBeLessThan(1.5);
    }
  });

  it('states the licence, the style and the names it is installed under', () => {
    const bytes = build(12);
    const view = new DataView(bytes.buffer);
    const os2 = tableAt(bytes, 'OS/2');
    expect(view.getUint16(os2 + 8)).toBe(12); // fsType
    expect(view.getUint16(os2 + 4)).toBe(700); // usWeightClass
    expect(view.getUint16(os2 + 62) & 0x20).toBe(0x20); // fsSelection BOLD
    expect(view.getUint16(tableAt(bytes, 'head') + 44) & 1).toBe(1); // macStyle bold
    const name = tableAt(bytes, 'name');
    const count = view.getUint16(name + 2);
    const storage = name + view.getUint16(name + 4);
    const names = new Map<number, string>();
    for (let i = 0; i < count; i++) {
      const at = name + 6 + i * 12;
      const length = view.getUint16(at + 8);
      const offset = view.getUint16(at + 10);
      let text = '';
      for (let k = 0; k < length; k += 2)
        text += String.fromCharCode(view.getUint16(storage + offset + k));
      names.set(view.getUint16(at + 6), text);
    }
    expect(names.get(1)).toBe('Test Face');
    expect(names.get(2)).toBe('Bold');
    expect(names.get(6)).toBe('TestFace-Bold');
  });

  it('sums to the sfnt magic number', () => {
    const bytes = build();
    let sum = 0;
    for (let i = 0; i < bytes.length; i += 4) {
      const word =
        ((bytes[i] ?? 0) << 24) | ((bytes[i + 1] ?? 0) << 16) | ((bytes[i + 2] ?? 0) << 8);
      sum = (sum + ((word | (bytes[i + 3] ?? 0)) >>> 0)) >>> 0;
    }
    expect(sum).toBe(0xb1b0afba);
  });

  it('kerns the pairs it is given, by glyph and in the order a search expects', () => {
    const font = build(8, [
      { left: 0x4f, right: 0x41, value: -40 }, // O A
      { left: 0x41, right: 0x4f, value: -60.4 }, // A O
      { left: 0x41, right: 0x42, value: -10 }, // no B in the face: nothing to kern
    ]);
    const at = tableAt(font, 'kern');
    const view = new DataView(font.buffer);
    expect(view.getUint16(at + 2)).toBe(1); // one subtable
    expect(view.getUint16(at + 8)).toBe(0x0001); // horizontal, format 0
    expect(view.getUint16(at + 10)).toBe(2); // the two pairs whose glyphs it holds
    const pairs = [0, 1].map((i) => {
      const entry = at + 18 + i * 6;
      return [view.getUint16(entry), view.getUint16(entry + 2), view.getInt16(entry + 4)];
    });
    // A is glyph 2 and O glyph 3: (2,3) sorts before (3,2).
    expect(pairs).toEqual([
      [2, 3, -60],
      [3, 2, -40],
    ]);
    // …and the same pairs as GPOS's `kern` feature, which is what a shaper reads.
    const kerning = parseTtf(font).kerning;
    expect(kerning.get('2,3')).toBe(-60);
    expect(kerning.get('3,2')).toBe(-40);
    expect(kerning.size).toBe(2);
  });

  it('joins the characters a ligature stands for into its glyph (GSUB liga)', () => {
    const font = buildTrueType({
      family: 'Test Face',
      bold: false,
      italic: false,
      postScriptName: 'TestFace',
      glyphs: [
        { codePoints: [0x66], outline: BOX, advance: 300 }, // f
        { codePoints: [0x69], outline: BOX, advance: 250 }, // i
      ],
      ligatures: [
        { codePoints: [0x66, 0x69], outline: ARCH, advance: 520 }, // fi
        { codePoints: [0x66, 0x66, 0x69], outline: ARCH, advance: 800 }, // ffi
        { codePoints: [0x66, 0x6c], outline: ARCH, advance: 520 }, // fl: no l to join
      ],
      ascent: 800,
      descent: -200,
      fsType: 8,
    });
    const parsed = parseTtf(font);
    // f is glyph 1, i glyph 2; the ligatures follow them, reachable by no character.
    expect(parsed.numGlyphs).toBe(5);
    expect(parsed.ligatures.get('1,2')).toBe(3);
    expect(parsed.ligatures.get('1,1,2')).toBe(4);
    expect(parsed.ligatures.size).toBe(2);
    expect([...parsed.advanceWidths]).toEqual([0, 300, 250, 520, 800]);
    expect(parsed.glyphForCodepoint(0xfb01)).toBe(0);
  });

  it('writes no kerning tables for a face kerned by nothing', () => {
    expect(() => tableAt(build(), 'kern')).toThrow();
    expect(() => tableAt(build(), 'GPOS')).toThrow();
    expect(() => tableAt(build(), 'GSUB')).toThrow();
  });

  it('builds the same bytes from the same face', () => {
    expect(build()).toEqual(build());
  });
});

describe('the embedding a face licence allows (OS/2 fsType)', () => {
  it('lets an installable or editable face into an editable document', () => {
    expect(editableEmbedding(undefined)).toBe(true);
    expect(editableEmbedding(0)).toBe(true);
    expect(editableEmbedding(0x0008)).toBe(true);
    // Several usage bits: the least restrictive stands.
    expect(editableEmbedding(0x000c)).toBe(true);
  });

  it('keeps out a restricted face, a print-only one, and one that may not be subset', () => {
    expect(editableEmbedding(0x0002)).toBe(false);
    expect(editableEmbedding(0x0004)).toBe(false);
    expect(editableEmbedding(0x0804)).toBe(false);
    expect(editableEmbedding(0x0100)).toBe(false);
    expect(editableEmbedding(0x0208)).toBe(false);
  });
});
