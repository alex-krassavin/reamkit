// DrawingML colour resolution (ECMA-376 §20.1.2.3).
//
// The shape parser stays theme-agnostic: it emits a RawColor (a direct sRGB
// value or a theme scheme-name reference) and a ColorResolver maps it to a
// concrete 6-hex string. The converter builds a theme-backed resolver from
// word/theme/theme1.xml (see theme-parser); standalone parser tests use the
// built-in default palette below.

// A colour transform child (§20.1.2.3): lumMod/lumOff modulate luminance,
// shade darkens toward black, tint lightens toward white. `val` is normalised to
// 0..1 (the XML stores thousandths of a percent). alpha composites over white
// (solid fills emit no transparency).
// ---------------------------------------------------------------------------
// Shared colour-node parsing (§20.1.2.3) — one owner for chart-parser and the
// word drawing-parser, which had drifted apart as verbatim copies.
// ---------------------------------------------------------------------------

import type { PoNode } from '@/core/po-helpers';
import { poAttr, poChildren, poIntAttr, poIs } from '@/core/po-helpers';

/**
 * A colour transform child (§20.1.2.3): `lumMod`/`lumOff` modulate luminance,
 * `shade` darkens toward black, `tint` lightens toward white. `val` is normalised
 * to 0..1 (the XML stores thousandths of a percent). `alpha` composites over white
 * (solid fills emit no transparency).
 */
export interface ColorMod {
  readonly kind: 'lumMod' | 'lumOff' | 'shade' | 'tint' | 'alpha' | 'satMod' | 'hueOff' | 'satOff';
  readonly val: number;
}

/**
 * A theme-agnostic colour reference emitted by the shape parser: either a direct
 * sRGB value or a theme scheme-name reference, each with optional colour
 * transforms. A {@link ColorResolver} maps it to a concrete 6-hex string.
 */
export type RawColor =
  | { readonly srgb: string; readonly mods?: ReadonlyArray<ColorMod> }
  | { readonly scheme: string; readonly mods?: ReadonlyArray<ColorMod> };

/** Maps a {@link RawColor} to a 6-hex string, or `undefined` when unresolvable. */
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

// IEC 61966-2-1 — the sRGB transfer function and its inverse. `shade` and
// `tint` are defined on the light a colour stands for, not on the byte that
// encodes it.
function toLinear(c: number): number {
  return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
}

function toSrgb(c: number): number {
  const v = clamp01(c);
  return v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055;
}

/**
 * Apply DrawingML colour transforms ({@link ColorMod}s) to a 6-hex value,
 * returning a 6-hex value. `shade`/`tint` scale in RGB; `lumMod`/`lumOff` adjust
 * luminance in HSL. An empty `mods` returns `hex` unchanged.
 *
 * @param hex  The base colour, RRGGBB.
 * @param mods The transforms to apply, in order.
 * @returns The transformed colour, uppercase RRGGBB.
 */
export function applyColorMods(hex: string, mods: ReadonlyArray<ColorMod>): string {
  if (mods.length === 0) return hex;
  const n = parseInt(hex, 16);
  let r = ((n >> 16) & 255) / 255;
  let g = ((n >> 8) & 255) / 255;
  let b = (n & 255) / 255;
  for (const m of mods) {
    if (m.kind === 'shade') {
      // §20.1.2.3.31 — a shade is a fraction of the input colour, and the
      // fraction is taken on LINEAR light, not on the gamma-encoded byte. Done
      // on the byte, the 51 % shade the Office theme builds its gradients from
      // came out near-black where both references draw a medium red
      // (fdo87488.docx's diagram).
      r = toSrgb(toLinear(r) * m.val);
      g = toSrgb(toLinear(g) * m.val);
      b = toSrgb(toLinear(b) * m.val);
    } else if (m.kind === 'tint') {
      // §20.1.2.3.34 — its twin toward white, on the same linear light.
      r = toSrgb(toLinear(r) * m.val + (1 - m.val));
      g = toSrgb(toLinear(g) * m.val + (1 - m.val));
      b = toSrgb(toLinear(b) * m.val + (1 - m.val));
    } else if (m.kind === 'lumMod' || m.kind === 'lumOff') {
      const [h, s, l] = rgbToHsl(r, g, b);
      const l2 = m.kind === 'lumMod' ? l * m.val : clamp01(l + m.val);
      [r, g, b] = hslToRgb(h, s, l2);
    } else if (m.kind === 'hueOff') {
      // §20.1.2.3.15 — a hue OFFSET, in degrees around the wheel. SmartArt is
      // built on it: one accent colour and a `hueOff` per node is how a diagram
      // gets its series of colours, and dropping it drew all three circles of
      // tdf83671_SmartArt_import.xlsx the same orange where the reference gives
      // them orange, green and blue.
      const [h, sat, l] = rgbToHsl(r, g, b);
      [r, g, b] = hslToRgb((((h + m.val) % 360) + 360) % 360, sat, l);
    } else if (m.kind === 'satOff') {
      // §20.1.2.3.27 — the saturation twin, an offset rather than a factor.
      const [h, sat, l] = rgbToHsl(r, g, b);
      [r, g, b] = hslToRgb(h, clamp01(sat + m.val), l);
    } else if (m.kind === 'alpha') {
      // §20.1.2.3.1 — the colour is drawn at `val` opacity. Composited over a
      // white page that is a tint of itself, which is what both references
      // show: dml-groupshape-childposition.docx draws eleven of its strokes at
      // 20% and we drew them all in full dark navy.
      r = r * m.val + (1 - m.val);
      g = g * m.val + (1 - m.val);
      b = b * m.val + (1 - m.val);
    } else {
      // §20.1.2.3.32 `satMod` — saturation modulation. The Office theme's gradients are
      // built from it (`<a:shade val="51000"/><a:satMod val="130000"/>`), and
      // dropping it drew every gallery shape a shade duller than either
      // reference: 47504.xlsx's rectangle came out grey-blue.
      const [h, s, l] = rgbToHsl(r, g, b);
      [r, g, b] = hslToRgb(h, clamp01(s * m.val), l);
    }
  }
  const toHex = (x: number): string =>
    Math.round(clamp01(x) * 255)
      .toString(16)
      .padStart(2, '0');
  return (toHex(r) + toHex(g) + toHex(b)).toUpperCase();
}

/**
 * Office 2013 default theme palette — the colours Word assigns to the standard
 * scheme slots when a document carries no custom theme part.
 */
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

/**
 * Resolve a `schemeClr` text/background alias (`tx1`/`bg1`/`tx2`/`bg2`,
 * §20.1.2.3.29) to its underlying `dk`/`lt` slot name; other names pass through.
 */
export function resolveSchemeName(name: string): string {
  return SCHEME_ALIAS[name] ?? name;
}

/**
 * Build a {@link ColorResolver} over a scheme-name → hex `palette`. sRGB
 * references pass through verbatim (upper-cased); scheme references are aliased
 * (via {@link resolveSchemeName}) then looked up; colour transforms are applied.
 *
 * @param palette The scheme-name → RRGGBB map (e.g. {@link DEFAULT_THEME_PALETTE}).
 * @returns A resolver that maps a {@link RawColor} to a hex string or `undefined`.
 */
export function makeColorResolver(palette: ReadonlyMap<string, string>): ColorResolver {
  return (raw) => {
    const base =
      'srgb' in raw ? raw.srgb.toUpperCase() : palette.get(resolveSchemeName(raw.scheme));
    if (base === undefined) return undefined;
    return raw.mods && raw.mods.length > 0 ? applyColorMods(base, raw.mods) : base;
  };
}

/** A {@link ColorResolver} backed by the {@link DEFAULT_THEME_PALETTE}. */
export const defaultColorResolver: ColorResolver = makeColorResolver(DEFAULT_THEME_PALETTE);

// §20.1.10.47 ST_PresetColorVal — the handful of names a `a:prstClr` uses in
// practice. The full table runs to 140 CSS names; these are the ones Word
// writes for shadows and outlines.
const PRESET_COLORS: ReadonlyMap<string, string> = new Map([
  ['black', '000000'],
  ['white', 'FFFFFF'],
  ['gray', '808080'],
  ['grey', '808080'],
  ['darkGray'.toLowerCase(), 'A9A9A9'],
  ['lightGray'.toLowerCase(), 'D3D3D3'],
  ['red', 'FF0000'],
  ['green', '008000'],
  ['blue', '0000FF'],
  ['yellow', 'FFFF00'],
  ['cyan', '00FFFF'],
  ['magenta', 'FF00FF'],
  ['dkGray'.toLowerCase(), 'A9A9A9'],
  ['ltGray'.toLowerCase(), 'D3D3D3'],
]);

/**
 * Read the colour transform children ({@link ColorMod}s) under an `a:srgbClr` /
 * `a:schemeClr` node, normalising each `val` from thousandths-of-a-percent to 0..1.
 */
export function readColorMods(colorNode: PoNode): Array<ColorMod> {
  const mods: Array<ColorMod> = [];
  for (const c of poChildren(colorNode)) {
    for (const kind of [
      'lumMod',
      'lumOff',
      'shade',
      'tint',
      'alpha',
      'satMod',
      'satOff',
    ] as const) {
      if (poIs(c, `a:${kind}`)) {
        const v = poIntAttr(c, 'val');
        if (v !== undefined) mods.push({ kind, val: v / 100000 });
      }
    }
    // §20.1.2.3.15 `a:hueOff` counts SIXTIETH-THOUSANDTHS OF A DEGREE, not
    // thousandths of a percent — a different unit from every other transform
    // here, and the reason it gets its own line. Kept in degrees, which is what
    // the HSL helpers below speak.
    if (poIs(c, 'a:hueOff')) {
      const v = poIntAttr(c, 'val');
      if (v !== undefined) mods.push({ kind: 'hueOff', val: v / 60000 });
    }
  }
  return mods;
}

/**
 * §20.1.2.3.32 `a:sysClr` — the host system's own colour. `val` names it and
 * `lastClr` carries what the producing application resolved it to, which is
 * what a reader uses; the two names that matter without one are the window
 * background and its text.
 */
const SYSTEM_COLORS: ReadonlyMap<string, string> = new Map([
  ['window', 'FFFFFF'],
  ['windowText', '000000'],
  ['windowFrame', '000000'],
  ['btnText', '000000'],
  ['btnFace', 'F0F0F0'],
  ['highlight', '0078D7'],
  ['highlightText', 'FFFFFF'],
]);

/**
 * Resolve an `a:srgbClr` / `a:schemeClr` / `a:sysClr` node to a hex string (with
 * colour transforms applied). Returns `undefined` when the node is some other
 * element, valueless, or the resolver does not know the colour. Container
 * traversal policy (stop at the first colour node vs continue past unresolved
 * ones) stays with the callers.
 *
 * @param c            The candidate colour node.
 * @param resolveColor The resolver mapping a {@link RawColor} to hex.
 * @returns The resolved RRGGBB, or `undefined`.
 */
export function resolveColorNode(c: PoNode, resolveColor: ColorResolver): string | undefined {
  // A system colour is stated, not referenced: it carries its own value and
  // there is nothing for the theme resolver to look up. Unread, it fell through
  // to whatever the container's style said instead — ConditionalFormattingSamples
  // .xlsx spells its buttons' text `<a:sysClr val="windowText"/>` over a shape
  // style asking for `lt1`, and fifteen of them came out white on pale blue.
  if (poIs(c, 'a:sysClr')) {
    const last = poAttr(c, 'lastClr');
    if (last && /^[0-9A-Fa-f]{6}$/.test(last)) return last.toUpperCase();
    return SYSTEM_COLORS.get(poAttr(c, 'val') ?? '');
  }
  // §20.1.2.3.22 `a:prstClr` — a colour NAMED rather than referenced, out of
  // the preset table. Word writes the shadow under a picture as
  // `<a:prstClr val="black"><a:alpha val="40000"/>`, and unread the shadow had
  // no colour to be drawn in at all (imgshadow.docx's six screenshots).
  if (poIs(c, 'a:prstClr')) {
    const hex = PRESET_COLORS.get((poAttr(c, 'val') ?? '').toLowerCase());
    if (hex === undefined) return undefined;
    const presetMods = readColorMods(c);
    return presetMods.length > 0 ? applyColorMods(hex, presetMods) : hex;
  }
  const isSrgb = poIs(c, 'a:srgbClr');
  if (!isSrgb && !poIs(c, 'a:schemeClr')) return undefined;
  const v = poAttr(c, 'val');
  if (!v) return undefined;
  const mods = readColorMods(c);
  const raw = isSrgb ? { srgb: v } : { scheme: v };
  return resolveColor(mods.length > 0 ? { ...raw, mods } : raw);
}
