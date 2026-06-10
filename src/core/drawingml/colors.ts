// DrawingML colour resolution (ECMA-376 §20.1.2.3).
//
// The shape parser stays theme-agnostic: it emits a RawColor (a direct sRGB
// value or a theme scheme-name reference) and a ColorResolver maps it to a
// concrete 6-hex string. The converter builds a theme-backed resolver from
// word/theme/theme1.xml (see theme-parser); standalone parser tests use the
// built-in default palette below.

// A colour transform child (§20.1.2.3): lumMod/lumOff modulate luminance,
// shade darkens toward black, tint lightens toward white. `val` is normalised to
// 0..1 (the XML stores thousandths of a percent). alpha is parsed but ignored
// (solid fills emit no transparency).
export interface ColorMod {
  readonly kind: 'lumMod' | 'lumOff' | 'shade' | 'tint' | 'alpha';
  readonly val: number;
}

export type RawColor =
  | { readonly srgb: string; readonly mods?: ReadonlyArray<ColorMod> }
  | { readonly scheme: string; readonly mods?: ReadonlyArray<ColorMod> };

export type ColorResolver = (raw: RawColor) => string | undefined;

function clamp01(x: number): number {
  return x < 0 ? 0 : x > 1 ? 1 : x;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) return [0, 0, l];
  const s = d / (1 - Math.abs(2 * l - 1));
  let h: number;
  if (max === r) h = ((g - b) / d) % 6;
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  h *= 60;
  if (h < 0) h += 360;
  return [h, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  let r = 0;
  let g = 0;
  let b = 0;
  if (h < 60) [r, g, b] = [c, x, 0];
  else if (h < 120) [r, g, b] = [x, c, 0];
  else if (h < 180) [r, g, b] = [0, c, x];
  else if (h < 240) [r, g, b] = [0, x, c];
  else if (h < 300) [r, g, b] = [x, 0, c];
  else [r, g, b] = [c, 0, x];
  return [r + m, g + m, b + m];
}

// Apply DrawingML colour transforms to a 6-hex value, returning a 6-hex value.
export function applyColorMods(hex: string, mods: ReadonlyArray<ColorMod>): string {
  if (mods.length === 0) return hex;
  const n = parseInt(hex, 16);
  let r = ((n >> 16) & 255) / 255;
  let g = ((n >> 8) & 255) / 255;
  let b = (n & 255) / 255;
  for (const m of mods) {
    if (m.kind === 'shade') {
      r *= m.val;
      g *= m.val;
      b *= m.val;
    } else if (m.kind === 'tint') {
      r = r * m.val + (1 - m.val);
      g = g * m.val + (1 - m.val);
      b = b * m.val + (1 - m.val);
    } else if (m.kind === 'lumMod' || m.kind === 'lumOff') {
      const [h, s, l] = rgbToHsl(r, g, b);
      const l2 = m.kind === 'lumMod' ? l * m.val : clamp01(l + m.val);
      [r, g, b] = hslToRgb(h, s, l2);
    }
  }
  const toHex = (x: number): string =>
    Math.round(clamp01(x) * 255)
      .toString(16)
      .padStart(2, '0');
  return (toHex(r) + toHex(g) + toHex(b)).toUpperCase();
}

// Office 2013 default theme palette — the colours Word assigns to the standard
// scheme slots when a document carries no custom theme part.
export const DEFAULT_THEME_PALETTE: ReadonlyMap<string, string> = new Map([
  ['dk1', '000000'],
  ['lt1', 'FFFFFF'],
  ['dk2', '44546A'],
  ['lt2', 'E7E6E6'],
  ['accent1', '4472C4'],
  ['accent2', 'ED7D31'],
  ['accent3', 'A5A5A5'],
  ['accent4', 'FFC000'],
  ['accent5', '5B9BD5'],
  ['accent6', '70AD47'],
  ['hlink', '0563C1'],
  ['folHlink', '954F72'],
]);

// schemeClr uses text/background aliases that map onto the dk/lt slots
// (§20.1.2.3.29). phClr (placeholder) is contextual and unresolved here.
const SCHEME_ALIAS: Readonly<Record<string, string>> = {
  tx1: 'dk1',
  bg1: 'lt1',
  tx2: 'dk2',
  bg2: 'lt2',
};

export function resolveSchemeName(name: string): string {
  return SCHEME_ALIAS[name] ?? name;
}

// Build a resolver over a scheme-name → hex palette. sRGB references pass
// through verbatim (upper-cased); scheme references are aliased then looked up.
export function makeColorResolver(palette: ReadonlyMap<string, string>): ColorResolver {
  return (raw) => {
    const base =
      'srgb' in raw ? raw.srgb.toUpperCase() : palette.get(resolveSchemeName(raw.scheme));
    if (base === undefined) return undefined;
    return raw.mods && raw.mods.length > 0 ? applyColorMods(base, raw.mods) : base;
  };
}

export const defaultColorResolver: ColorResolver = makeColorResolver(DEFAULT_THEME_PALETTE);
