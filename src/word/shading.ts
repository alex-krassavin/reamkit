// ECMA-376 Part 1 §17.3.5 `w:shd` — the background behind a run, a paragraph
// or a cell. One resolver for all three: the element states a PATTERN (`w:val`)
// drawn in `w:color` over a background `w:fill`, and what reaches the page is
// the two blended by the pattern's density.

/** §17.18.78 ST_Shd — the pattern densities, as the share of `color` over `fill`. */
const PATTERN_DENSITY: ReadonlyMap<string, number> = new Map([
  ['solid', 1],
  ['pct5', 0.05],
  ['pct10', 0.1],
  ['pct12', 0.125],
  ['pct15', 0.15],
  ['pct20', 0.2],
  ['pct25', 0.25],
  ['pct30', 0.3],
  ['pct35', 0.35],
  ['pct37', 0.375],
  ['pct40', 0.4],
  ['pct45', 0.45],
  ['pct50', 0.5],
  ['pct55', 0.55],
  ['pct60', 0.6],
  ['pct62', 0.625],
  ['pct65', 0.65],
  ['pct70', 0.7],
  ['pct75', 0.75],
  ['pct80', 0.8],
  ['pct85', 0.85],
  ['pct87', 0.875],
  ['pct90', 0.9],
  ['pct95', 0.95],
  // The named hatches carry no density of their own; Word draws them at a
  // quarter or so of the pattern colour, which is what these approximate.
  ['thinHorzStripe', 0.25],
  ['thinVertStripe', 0.25],
  ['thinDiagStripe', 0.25],
  ['thinHorzCross', 0.3],
  ['thinDiagCross', 0.3],
  ['horzStripe', 0.4],
  ['vertStripe', 0.4],
  ['diagStripe', 0.4],
  ['horzCross', 0.5],
  ['diagCross', 0.5],
]);

const HEX6 = /^[0-9A-Fa-f]{6}$/u;

/**
 * The colour a `w:shd` actually paints: its `w:fill` background with the
 * pattern colour blended in at the pattern's density. `clear` (and an absent
 * `w:val`) is the fill alone; `nil` paints nothing.
 *
 * @param val   `w:val` — the pattern.
 * @param color `w:color` — the pattern's own colour (`auto` ⇒ black).
 * @param fill  `w:fill` — the background (`auto` ⇒ white).
 * @returns The 6-hex colour to paint, or `undefined` when nothing is painted.
 */
export function shadingFillHex(
  val: string | undefined,
  color: string | undefined,
  fill: string | undefined,
): string | undefined {
  if (val === 'nil') return undefined;
  const bg = fill !== undefined && HEX6.test(fill) ? fill.toUpperCase() : undefined;
  const density = val !== undefined ? PATTERN_DENSITY.get(val) : undefined;
  if (density === undefined) return bg;
  // `auto` on the pattern colour is black over the fill — that is what makes a
  // `pct15` of nothing but white a light grey.
  const fg = color !== undefined && HEX6.test(color) ? color.toUpperCase() : '000000';
  return blend(fg, bg ?? 'FFFFFF', density);
}

/** `a` over `b` at `t` (0..1), per channel. */
function blend(a: string, b: string, t: number): string {
  const ch = (hex: string, i: number): number => parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  const mix = (i: number): string =>
    Math.round(ch(a, i) * t + ch(b, i) * (1 - t))
      .toString(16)
      .padStart(2, '0');
  return (mix(0) + mix(1) + mix(2)).toUpperCase();
}
