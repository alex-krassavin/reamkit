// IR core — canonical length unit (ir-design.md §4).
//
// The IR trees carry exactly one length unit: PostScript points (1/72 inch),
// as the branded type `Pt`. Readers convert their format-native units (twips,
// half-points, EMU, …) at the boundary; writers convert out the same way.
//
// Branding is intentionally shallow: `Pt` is assignable to `number`, so layout
// arithmetic reads values freely. Only *constructing* a Pt requires going
// through `pt()` or a `*ToPt` converter — the brand marks API boundaries, not
// every intermediate expression.

/**
 * The IR's one length unit: PostScript points (1/72 inch), as a branded number.
 * Readers convert their format-native units (twips, half-points, EMU, …) to `Pt`
 * at the boundary, and writers convert out the same way. The brand is shallow —
 * a `Pt` is assignable to `number` so layout arithmetic reads it freely; only
 * *constructing* one requires {@link pt} or a `*ToPt` converter.
 */
export type Pt = number & { readonly __brand: 'pt' };

/** Brand a raw number as points. The number must already BE points. */
export function pt(value: number): Pt {
  return value as Pt;
}

/**
 * Twentieths of a point (OOXML `dxa` / twips): 20 twips = 1 pt.
 *
 * Deliberately `* (1/20)`, not `/ 20`: the two differ in the last ulp for ~35%
 * of integer inputs, and the layout engine historically multiplied by a
 * `TWIP_TO_PT = 1/20` constant — keeping the exact operator keeps the
 * byte-identical corpus gate meaningful across the IR migration.
 */
const TWIP_TO_PT = 1 / 20;
export function twipsToPt(twips: number): Pt {
  return (twips * TWIP_TO_PT) as Pt;
}

/** Half-points (OOXML font sizes, `w:sz`): 2 half-points = 1 pt. */
export function halfPtToPt(halfPt: number): Pt {
  return (halfPt / 2) as Pt;
}

/** Eighths of a point (OOXML border widths, `w:sz` on borders): 8 = 1 pt. */
export function eighthPtToPt(eighthPt: number): Pt {
  return (eighthPt / 8) as Pt;
}

/** English Metric Units (DrawingML): 914400 EMU = 1 inch = 72 pt → 12700 EMU = 1 pt. */
export function emuToPt(emu: number): Pt {
  return (emu / 12700) as Pt;
}

/** CSS reference pixels: 96 px = 1 inch = 72 pt. */
export function pxToPt(px: number): Pt {
  return (px * 0.75) as Pt;
}

/** Inches: 1 inch = 72 pt. */
export function inchToPt(inches: number): Pt {
  return (inches * 72) as Pt;
}

/** Millimetres: 25.4 mm = 1 inch = 72 pt. */
export function mmToPt(mm: number): Pt {
  return ((mm / 25.4) * 72) as Pt;
}
