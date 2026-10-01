// ECMA-376 Part 1 §18.8.19 `tint` — how Excel lightens and darkens a colour.
//
// The standard moves the colour's luminance toward white or black over HLS
// "where HLSMAX is currently 255", and leaves the arithmetic at that. Excel
// works it in the Windows HLS model instead: a 0..240 scale, every step an
// integer, rounded as Windows' ColorRGBToHLS / ColorHLSToRGB round, and the
// tint in whole tenths of a percent. That model gives all fifty shades Excel's
// colour picker offers for the Office 2013–2022 theme, to the digit — 4472C4
// lighter 80% is D9E1F2, where floating-point HSL makes DAE3F3 — and the tint
// Excel writes for a shade (0.79998168889431442, its 80% in 32767ths) rounds
// back to the tenth of a percent it was picked as. Excel's own PDF of a grid
// of 23 colours under 21 tints (theme and `rgb` alike, 2026-10-01) agrees in
// 471 cells of 483; the twelve it does not are all one tint, 0.123, which
// Excel rounds to the nearest luminance where every other tint it was given
// floors. It writes no such tint itself.

/** The top of the HLS scale. */
const HLS_MAX = 240;
/** The top of an RGB channel. */
const RGB_MAX = 255;

/** Integer division of a non-negative numerator, as C divides unsigned values. */
const div = (a: number, b: number): number => Math.floor(a / b);

/** RGB to hue, luminance and saturation, each 0..240, rounded as Windows rounds. */
function toHls(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const sum = max + min;
  const l = div(sum * HLS_MAX + RGB_MAX, 2 * RGB_MAX);
  // A grey has no hue; Windows calls it two thirds of the way round.
  if (max === min) return [(HLS_MAX * 2) / 3, l, 0];
  const range = max - min;
  const s =
    l <= HLS_MAX / 2
      ? div(range * HLS_MAX + div(sum, 2), sum)
      : div(range * HLS_MAX + div(2 * RGB_MAX - sum, 2), 2 * RGB_MAX - sum);
  const away = (c: number): number => div((max - c) * (HLS_MAX / 6) + div(range, 2), range);
  let h =
    r === max
      ? away(b) - away(g)
      : g === max
        ? HLS_MAX / 3 + away(r) - away(b)
        : (HLS_MAX * 2) / 3 + away(g) - away(r);
  if (h < 0) h += HLS_MAX;
  if (h > HLS_MAX) h -= HLS_MAX;
  return [h, l, s];
}

/** Hue, luminance and saturation, each 0..240, back to RGB as Windows rounds it. */
function toRgb(h: number, l: number, s: number): [number, number, number] {
  const channel = (v: number): number => div(v * RGB_MAX + HLS_MAX / 2, HLS_MAX);
  if (s === 0) {
    const grey = channel(l);
    return [grey, grey, grey];
  }
  const hi =
    l <= HLS_MAX / 2
      ? div(l * (HLS_MAX + s) + HLS_MAX / 2, HLS_MAX)
      : l + s - div(l * s + HLS_MAX / 2, HLS_MAX);
  const lo = 2 * l - hi;
  const at = (hue: number): number => {
    const x = hue < 0 ? hue + HLS_MAX : hue > HLS_MAX ? hue - HLS_MAX : hue;
    if (x < HLS_MAX / 6) return lo + div((hi - lo) * x + HLS_MAX / 12, HLS_MAX / 6);
    if (x < HLS_MAX / 2) return hi;
    if (x < (HLS_MAX * 2) / 3) {
      return lo + div((hi - lo) * ((HLS_MAX * 2) / 3 - x) + HLS_MAX / 12, HLS_MAX / 6);
    }
    return lo;
  };
  return [channel(at(h + HLS_MAX / 3)), channel(at(h)), channel(at(h - HLS_MAX / 3))];
}

/**
 * Lighten a colour by a positive tint or darken it by a negative one, as Excel
 * does (§18.8.19): a tint of 0.4 takes the luminance 40% of the way to white,
 * one of -0.25 a quarter of the way to black.
 *
 * @param hex  The colour, RRGGBB.
 * @param tint The tint, -1..1.
 * @returns The tinted colour, RRGGBB in upper case; `hex` itself at no tint.
 */
export function applyTint(hex: string, tint: number): string {
  if (!Number.isFinite(tint)) return hex;
  const thousandths = Math.max(-1000, Math.min(1000, Math.round(tint * 1000)));
  if (thousandths === 0) return hex;
  const v = parseInt(hex, 16);
  const [h, l, s] = toHls((v >> 16) & 255, (v >> 8) & 255, v & 255);
  const lum =
    thousandths > 0
      ? l + div((HLS_MAX - l) * thousandths, 1000)
      : div(l * (1000 + thousandths), 1000);
  return toRgb(h, lum, s)
    .map((c) => c.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}
