// The shape a metafile is read INTO: a page of drawing primitives in the
// metafile's own logical units, plus the box those units span.
//
// EMF (MS-EMF) and WMF (MS-WMF) are two spellings of the same device
// interface — a device context with a pen, a brush, a font and a current
// point, and records that draw through it. So both readers build this, and
// everything downstream — the layout that scales it into a picture's box, the
// writers that draw it — sees one vocabulary.
//
// Text arrives as a STRING plus the font that was selected, not as glyphs: the
// metafile names a typeface and the caller owns the font machinery.

import type { StrokeStyle, VectorPath } from '@/core/vector';

/** A path drawn by the metafile: filled, stroked, or both. */
export interface PicturePath {
  readonly kind: 'path';
  readonly paths: ReadonlyArray<VectorPath>;
  /** 6-hex, no leading `#`. Absent when the path is not filled. */
  readonly fillColorHex?: string;
  readonly stroke?: StrokeStyle;
}

/** A run of text the metafile puts at a point, in its own logical units. */
export interface PictureText {
  readonly kind: 'text';
  readonly text: string;
  /** The reference point, in logical units (y DOWN, as the metafile has it). */
  readonly x: number;
  readonly y: number;
  /** Which corner of the text the point names (MS-WMF §2.1.2.4 TextAlignment). */
  readonly alignH: 'left' | 'center' | 'right';
  /** Whether `y` is the text's TOP (the default) or its baseline. */
  readonly alignBaseline: boolean;
  /** Em size in logical units — the height the font was created with. */
  readonly sizeLu: number;
  readonly colorHex: string;
  readonly fontFamily?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** Tenths of a degree, counter-clockwise — the font's own escapement. */
  readonly escapement?: number;
}

/** A bitmap the metafile blits into a rectangle of its own logical frame. */
export interface PictureImage {
  readonly kind: 'image';
  /** The destination rectangle in logical units (y down). */
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** The bitmap as a standalone file (PNG), ready for the resource store. */
  readonly png: Uint8Array;
}

export type PicturePrim = PicturePath | PictureText | PictureImage;

/**
 * One metafile, read: its primitives and the logical box they are drawn in.
 * Coordinates are the metafile's own, with y running DOWN; the caller maps the
 * box onto the picture's frame.
 */
export interface MetaPicture {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly prims: ReadonlyArray<PicturePrim>;
  /** Records the reader knew of but did not draw, for a loss note. */
  readonly skipped: ReadonlyArray<string>;
}

/** A pen, a brush and a font, as the device context holds them. */
export interface MetaPen {
  readonly kind: 'pen';
  readonly colorHex: string;
  readonly widthLu: number;
  readonly style: 'solid' | 'dash' | 'dot' | 'dashdot' | 'dashdotdot' | 'none';
}

export interface MetaBrush {
  readonly kind: 'brush';
  readonly colorHex: string;
  /** `null` for BS_HOLLOW/BS_NULL — a brush that paints nothing. */
  readonly hollow: boolean;
}

export interface MetaFont {
  readonly kind: 'font';
  readonly heightLu: number;
  readonly family?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly escapement?: number;
}

/**
 * An object the reader does not model — a palette, a region, a pattern brush.
 * It still OCCUPIES a slot in the metafile's object table: handles are indices
 * into that table, so a creation the reader passed over silently shifted every
 * one that followed. 23884's chart selects its coloured pens by handle, and one
 * unrecorded palette put every selection one slot back — on the black pen the
 * pens had displaced.
 */
export interface MetaOther {
  readonly kind: 'other';
}

export type MetaObject = MetaPen | MetaBrush | MetaFont | MetaOther;

/** The device context: everything a record draws THROUGH. */
export interface DeviceContext {
  pen: MetaPen;
  brush: MetaBrush;
  font?: MetaFont;
  textColorHex: string;
  bkColorHex: string;
  /** MS-WMF §2.1.1.9 — TRANSPARENT leaves the background of text unpainted. */
  bkOpaque: boolean;
  alignH: 'left' | 'center' | 'right';
  alignBaseline: boolean;
  fillRule: 'nonzero' | 'evenodd';
  x: number;
  y: number;
  /** The world→page transform, as `[a, b, c, d, e, f]`. */
  transform: readonly [number, number, number, number, number, number];
}

export const DEFAULT_PEN: MetaPen = {
  kind: 'pen',
  colorHex: '000000',
  widthLu: 1,
  style: 'solid',
};
export const DEFAULT_BRUSH: MetaBrush = { kind: 'brush', colorHex: 'FFFFFF', hollow: false };

/** A device context in its documented initial state (MS-WMF §3.1.5). */
export function newDeviceContext(): DeviceContext {
  return {
    pen: DEFAULT_PEN,
    brush: DEFAULT_BRUSH,
    textColorHex: '000000',
    bkColorHex: 'FFFFFF',
    bkOpaque: true,
    alignH: 'left',
    alignBaseline: false,
    fillRule: 'evenodd', // ALTERNATE is the documented default
    x: 0,
    y: 0,
    transform: [1, 0, 0, 1, 0, 0],
  };
}

export function cloneDc(dc: DeviceContext): DeviceContext {
  return { ...dc };
}

/** A COLORREF (`0x00bbggrr`) as the 6-hex this codebase's colours use. */
export function colorRef(v: number): string {
  const r = v & 0xff;
  const g = (v >>> 8) & 0xff;
  const b = (v >>> 16) & 0xff;
  return [r, g, b]
    .map((c) => c.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

/** Apply the world transform to a point. */
export function applyTransform(
  t: DeviceContext['transform'],
  x: number,
  y: number,
): { x: number; y: number } {
  return { x: t[0] * x + t[2] * y + t[4], y: t[1] * x + t[3] * y + t[5] };
}
