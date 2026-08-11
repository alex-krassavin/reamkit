// E-PDF EP16c — read a PDF axial/radial shading back into a gradient fill. A
// shape filled with a shading pattern (`/Pattern cs /Pn scn` … fill) is the
// path-bounded case the interpreter captures; this resolves the page's
// /Pattern resources into ShapeGradients keyed by resource name, evaluating the
// shading's /Function (type 2 exponential, type 3 stitching, type 0 sampled, and
// type 4 by running it through `./function`) for the colour stops. The bare `sh`
// operator (clip-bounded) is not captured.
//
// The colour SPACES a page names are read here too, including the way out of a
// `/Separation` or `/DeviceN` (§8.6.6.4): its tint transform, as something the
// interpreter can call.

import { cieToSrgb } from './cie-color';
import { readFunction } from './function';
import type { CieSpace } from './cie-color';
import type { PdfFunction } from './function';
import type { GradientStop, ShapeGradient } from '@/core/vector';
import type { PdfDict, PdfValue } from '@/pdf/objects';

import type { PdfFile } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

/**
 * Resolve a page's `/Pattern` resources into gradient fills (E-PDF EP16c, ISO
 * 32000-1 §8.7.4.5). Every `PatternType` 2 (shading) pattern is evaluated — its
 * `/Shading` type 2 (axial) or 3 (radial) plus the `/Function` colour stops —
 * and keyed by resource name; the interpreter looks the name up when a shape is
 * filled with `/Pattern cs /Pn scn`. The bare `sh` operator (clip-bounded) is
 * not captured.
 *
 * @param file      The owning file.
 * @param resources The resource dictionary in force — a page's, or the form or
 *                  annotation appearance's own.
 * @returns A map from pattern resource name to its {@link ShapeGradient}.
 */
export function buildShadingMap(
  file: PdfFile,
  resources: PdfDict | undefined,
): Map<string, ShapeGradient> {
  const out = new Map<string, ShapeGradient>();
  if (!resources) return out;
  const patterns = file.get(resources, 'Pattern');
  if (!(patterns instanceof Map)) return out;
  for (const [nm, value] of patterns) {
    const pat = file.resolve(value);
    if (!(pat instanceof Map)) continue;
    const shading = dictOf(file.resolve(pat.get('Shading') ?? PDF_NULL));
    if (!shading) continue;
    const gradient = parseShading(file, shading);
    if (gradient) out.set(nm, gradient);
  }
  return out;
}

/**
 * §8.7.4.5.2 — an axial or radial shading as a gradient, for a bare `sh`.
 *
 * A `sh` paints the CLIP rather than a path, so what it needs is not a fill for
 * a shape the page drew but the gradient itself, to fill the region with. It is
 * the same reading `buildShadingMap` does for a pattern.
 *
 * @param file The owning file.
 * @param sh   The shading dictionary.
 * @returns The gradient, or `undefined` for a type this does not read.
 */
export function gradientShading(file: PdfFile, sh: PdfDict): ShapeGradient | undefined {
  return parseShading(file, sh);
}

/**
 * §8.7.4.5 — which kind of shading this is, as the file states it.
 *
 * @param file The owning file.
 * @param sh   The shading dictionary.
 * @returns Its `/ShadingType`, or 0 where the file states none.
 */
export function shadingTypeOf(file: PdfFile, sh: PdfDict): number {
  return numOf(file.get(sh, 'ShadingType'));
}

function parseShading(file: PdfFile, sh: PdfDict): ShapeGradient | undefined {
  const type = numOf(file.get(sh, 'ShadingType'));
  if (type !== 2 && type !== 3) return undefined; // only axial (2) / radial (3)
  const stops = parseFunction(file, sh.get('Function'), shadingSpace(file, sh));
  if (!stops || stops.length === 0) return undefined;
  if (type === 3) return { kind: 'radial', stops };
  // Axial: the angle is the Coords direction, with y negated (PDF y-up → the
  // DrawingML y-down convention the model stores).
  const c = numArray(file, sh.get('Coords'));
  const angle =
    c && c.length >= 4
      ? ((((Math.atan2(-(c[3]! - c[1]!), c[2]! - c[0]!) * 180) / Math.PI) % 360) + 360) % 360
      : 0;
  return { kind: 'linear', angle, stops };
}

// A 1-in / n-out function → colour stops. Recurses for the type-3 stitching case.
function parseFunction(
  file: PdfFile,
  value: PdfValue | undefined,
  space: ColorSpaceInfo | undefined,
): Array<GradientStop> | undefined {
  const resolved = value !== undefined ? file.resolve(value) : undefined;
  const dict = dictOf(resolved);
  if (!dict) return undefined;
  const type = numOf(file.get(dict, 'FunctionType'));

  if (type === 2) {
    const c0 = colorOf(numArray(file, dict.get('C0')) ?? [0], space);
    const c1 = colorOf(numArray(file, dict.get('C1')) ?? [1], space);
    return [
      { offset: 0, colorHex: c0 },
      { offset: 1, colorHex: c1 },
    ];
  }

  if (type === 3) {
    const fns = dict.get('Functions');
    const subs = Array.isArray(fns) ? fns : undefined;
    if (!subs) return undefined;
    const bounds = numArray(file, dict.get('Bounds')) ?? [];
    const domain = numArray(file, dict.get('Domain')) ?? [0, 1];
    const d0 = domain[0] ?? 0;
    const d1 = domain[domain.length - 1] ?? 1;
    const edges = [d0, ...bounds, d1];
    const stops: Array<GradientStop> = [];
    const span = d1 - d0 || 1;
    for (let i = 0; i < subs.length; i++) {
      const sub = parseFunction(file, subs[i], space);
      if (!sub || sub.length === 0) continue;
      // Each subfunction's own 0..1 laid onto the piece of the domain it holds
      // — ALL of its stops, not just its ends. A subfunction may be a stitch
      // itself, and its inner steps are the picture: issue10572.pdf stitches
      // twelve copies of a green/blue pair whose `/Bounds [0.5 0.5]` makes a
      // hard edge, and reduced to first-and-last each pair came back a smooth
      // fade instead of two flat bands.
      const lo = ((edges[i] ?? d0) - d0) / span;
      const hi = ((edges[i + 1] ?? d1) - d0) / span;
      for (const stop of sub) pushStop(stops, lo + stop.offset * (hi - lo), stop.colorHex);
    }
    return stops.length > 0 ? stops : undefined;
  }

  if (type === 0 && resolved instanceof PdfStream) {
    return sampleFunction(file, resolved, space);
  }
  if (type === 4) {
    // §7.10.5 — a program, not a table: the only way to a stop is to RUN it,
    // which `./function` does. Sampled along the domain like any other curve.
    const fn = readFunction(file, resolved);
    if (!fn) return undefined;
    const domain = numArray(file, dict.get('Domain')) ?? [0, 1];
    const d0 = domain[0] ?? 0;
    const d1 = domain[1] ?? 1;
    const stops: Array<GradientStop> = [];
    for (let s = 0; s < PS_STOPS; s++) {
      const off = s / (PS_STOPS - 1);
      pushStop(stops, off, colorOf(fn([d0 + off * (d1 - d0)]), space));
    }
    return stops.length > 0 ? stops : undefined;
  }
  return undefined;
}

/**
 * §8.7.4.5.3 — a FUNCTION-BASED shading, sampled into a picture.
 *
 * Type 1 is not a ramp between two points: it is a function of two variables
 * over a rectangle, and no gradient can stand for one. Painted by a bare `sh`
 * it fills the clip, and nothing here lifted it at all —
 * function_based_shading.pdf is nine such squares and 43% of the page's ink,
 * and reconstructed to a blank sheet.
 *
 * Sampled it is exactly a picture, which every format downstream can show. The
 * grid is fixed: the function is smooth by construction (§8.7.4.5.3 gives it a
 * `/Domain` and nothing else), so more samples buy nothing a reader can see.
 *
 * @param file    The owning file.
 * @param shading The shading dictionary.
 * @returns The picture and the domain it covers, or `undefined` for a shading
 *          of another type or one whose function cannot be run.
 */
export function sampledShading(
  file: PdfFile,
  shading: PdfDict,
): { rgb: Uint8Array; size: number; domain: [number, number, number, number] } | undefined {
  if (numOf(file.get(shading, 'ShadingType')) !== 1) return undefined;
  const space = shadingSpace(file, shading);
  const fn = readFunction(file, shading.get('Function'));
  if (!fn) return undefined;
  const d = numArray(file, shading.get('Domain')) ?? [0, 1, 0, 1];
  if (d.length < 4) return undefined;
  const domain: [number, number, number, number] = [d[0]!, d[1]!, d[2]!, d[3]!];
  const size = SAMPLED_SIDE;
  const rgb = new Uint8Array(size * size * 3);
  for (let row = 0; row < size; row++) {
    // The raster's first row is the TOP, and the domain's y runs up.
    const y = domain[3] - ((row + 0.5) / size) * (domain[3] - domain[2]);
    for (let col = 0; col < size; col++) {
      const x = domain[0] + ((col + 0.5) / size) * (domain[1] - domain[0]);
      const hex = colorOf(fn([x, y]), space);
      const at = (row * size + col) * 3;
      rgb[at] = Number.parseInt(hex.slice(0, 2), 16);
      rgb[at + 1] = Number.parseInt(hex.slice(2, 4), 16);
      rgb[at + 2] = Number.parseInt(hex.slice(4, 6), 16);
    }
  }
  return { rgb, size, domain };
}

/**
 * §8.7.4.5 — the colour space a shading's function lands in.
 *
 * The components alone do not say: a `/Separation` function gives ONE number
 * and that number is a strength of ink, not a grey level.
 * function_based_shading_cmyk.pdf's third square is a spot colour whose full
 * tint is a warm red, and read as grey it came back a black-to-white ramp.
 */
function shadingSpace(file: PdfFile, sh: PdfDict): ColorSpaceInfo | undefined {
  const cs = sh.get('ColorSpace');
  return cs === undefined ? undefined : colorSpaceAt(file, file.resolve(cs), 0);
}

/** How many samples a side a function-based shading is drawn at. */
const SAMPLED_SIDE = 128;

/** How many places a type-4 gradient is sampled at along its domain. */
const PS_STOPS = 16;

// §7.10.2 sampled function — read the table and sample it at a few offsets.
function sampleFunction(
  file: PdfFile,
  stream: PdfStream,
  space: ColorSpaceInfo | undefined,
): Array<GradientStop> | undefined {
  const d = stream.dict;
  const size = numArray(file, d.get('Size'));
  const range = numArray(file, d.get('Range'));
  const bps = numOf(file.get(d, 'BitsPerSample'));
  if (!size || !range || size.length < 1 || range.length < 2 || bps < 1) return undefined;
  const n = size[0]!;
  const comps = range.length / 2;
  const data = file.streamData(stream);
  const maxv = 2 ** bps - 1;
  const bitAt = (sampleIdx: number, comp: number): number => {
    let bit = (sampleIdx * comps + comp) * bps;
    let v = 0;
    for (let k = 0; k < bps; k++) {
      const byte = data[bit >> 3] ?? 0;
      v = (v << 1) | ((byte >> (7 - (bit & 7))) & 1);
      bit++;
    }
    return v;
  };
  const count = Math.min(Math.max(n, 2), 16);
  const stops: Array<GradientStop> = [];
  for (let s = 0; s < count; s++) {
    const off = s / (count - 1);
    const j = Math.round(off * (n - 1));
    const c: Array<number> = [];
    for (let comp = 0; comp < comps; comp++) {
      const lo = range[comp * 2]!;
      const hi = range[comp * 2 + 1]!;
      c.push(lo + (bitAt(j, comp) / maxv) * (hi - lo));
    }
    pushStop(stops, off, colorOf(c));
  }
  return stops.length > 0 ? stops : undefined;
}

// --- helpers ----------------------------------------------------------------

function dictOf(v: PdfValue | undefined): PdfDict | undefined {
  if (v instanceof PdfStream) return v.dict;
  if (v instanceof Map) return v;
  return undefined;
}

function numOf(v: PdfValue | undefined): number {
  return typeof v === 'number' ? v : 0;
}

function numArray(file: PdfFile, v: PdfValue | undefined): Array<number> | undefined {
  const r = v !== undefined ? file.resolve(v) : undefined;
  if (!Array.isArray(r)) return undefined;
  return r.map((x) => (typeof x === 'number' ? x : 0));
}

// Colour components (0..1, in the shading colour space) → a 6-hex sRGB string.
// A shading's function output as a colour. Where the stated space cannot say
// what the numbers mean, the COUNT is the fallback — and a lone number is a
// grey level here, unlike an `sc` operand, because a shading that states no
// space at all has nothing else to go on. bug1721218_reduced.pdf paints a shape
// through a space this does not read, and defaulted to black it arrived as a
// black blob beside the router the page draws.
function colorOf(c: ReadonlyArray<number>, space?: ColorSpaceInfo): string {
  return spaceColor(c, space) ?? spaceColor(c, undefined) ?? grayHex(c[0] ?? 0);
}

/**
 * §8.6.8 — the colour a run of `sc` / `scn` components comes to.
 *
 * The space in force decides, and where it was not read the COUNT is the next
 * best witness: three numbers are RGB and four are CMYK on every device space
 * there is. One number is the ambiguous case — grey in a device space, but the
 * strength of a colorant in a Separation, where 1 is the ink at full and reads
 * dark — so a lone component is only taken where the space said what it means.
 *
 * @param nums  The components, as the page or a shading's function states them.
 * @param space The space in force, where it was read.
 * @returns The colour as 6 hex digits, or `undefined` where the numbers do not
 *          say what colour they are.
 */
export function spaceColor(
  nums: ReadonlyArray<number>,
  space: ColorSpaceInfo | undefined,
): string | undefined {
  if (nums.length === 0) return undefined;
  // §8.6.5.6/§8.6.5.7 — a CIE space's numbers mean what its own transform makes
  // of them, which is not what the same numbers mean to a device space.
  if (space?.cie) {
    const [r, g, b] = cieToSrgb(space.cie, nums);
    return rgbHex(r, g, b);
  }
  const kind = space?.kind ?? (nums.length === 3 ? 'rgb' : nums.length === 4 ? 'cmyk' : undefined);
  switch (kind) {
    case 'rgb':
      return nums.length >= 3 ? rgbHex(nums[0]!, nums[1]!, nums[2]!) : undefined;
    case 'cmyk':
      return nums.length >= 4 ? cmykHex(nums[0]!, nums[1]!, nums[2]!, nums[3]!) : undefined;
    case 'gray':
      return grayHex(nums[0]!);
    case 'tint':
      // §8.6.6.4/§8.6.6.5 — a tint is a strength of colorant, and the space's
      // own transform says what colour that strength comes to. Run it and the
      // answer is in the alternate space; devicen.pdf's three triangles are
      // green, blue and red.
      if (space?.tint) {
        const out = space.tint.transform(nums);
        const hex = spaceColor(out, space.tint.alternate);
        if (hex !== undefined) return hex;
      }
      // Without a transform this can run, the honest reading is "this much
      // ink", which is the inverse of a grey level.
      return grayHex(1 - Math.max(...nums));
    default:
      return undefined;
  }
}

// PDF colour operands (0..1 per channel) → a 6-hex sRGB string.
function clamp255(v: number): number {
  return Math.max(0, Math.min(255, Math.round(v * 255)));
}
function hex2(v: number): string {
  return clamp255(v).toString(16).padStart(2, '0');
}

/** Three channels, each 0..1, as 6 upper-case hex digits. */
export function rgbHex(r: number, g: number, b: number): string {
  return (hex2(r) + hex2(g) + hex2(b)).toUpperCase();
}

/** One grey level, 0..1, as 6 upper-case hex digits. */
export function grayHex(v: number): string {
  return rgbHex(v, v, v);
}

/** §8.6.4.4 — four inks, each 0..1, as 6 upper-case hex digits. */
export function cmykHex(c: number, m: number, y: number, k: number): string {
  return rgbHex((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k));
}

function hex255(r: number, g: number, b: number): string {
  const h = (x: number): string =>
    Math.max(0, Math.min(255, Math.round(x)))
      .toString(16)
      .padStart(2, '0');
  return (h(r) + h(g) + h(b)).toUpperCase();
}

// Append a stop, coalescing one at the same offset (keeps the gradient monotone).
function pushStop(stops: Array<GradientStop>, offset: number, colorHex: string): void {
  const o = Math.max(0, Math.min(1, offset));
  const last = stops[stops.length - 1];
  if (last && Math.abs(last.offset - o) < 1e-6) {
    if (last.colorHex === colorHex) return;
  }
  stops.push({ offset: o, colorHex });
}

/**
 * §11.6.4.4 — the constant fill alpha (`/ca`) of every `/ExtGState` the page
 * names, by name.
 *
 * `gs` sets a whole graphics state at once, and one of the things in it is how
 * opaque the paint is. 22060_A1_01_Plans.pdf marks its evacuation routes with
 * a green band at `ca` 0.6, meant to be read THROUGH: painted solid, the floor
 * plan under each band disappears.
 *
 * §11.3.5 `/BM` comes off the same dictionary. A blend nothing downstream can
 * perform is still worth knowing: `Multiply` and `Darken` both let dark ink
 * under the paint show through, which is what a highlighter IS, and a mark that
 * only darkens belongs UNDER the words rather than over them.
 *
 * @param file      The owning file.
 * @param resources The resource dictionary in force — a page's, or the form or
 *                  annotation appearance's own, since `gs` resolves against
 *                  whichever is current.
 * @returns Name → the fill alpha and blend worth carrying, for the states that
 *          state one.
 */
export function buildAlphaMap(file: PdfFile, resources: PdfDict | undefined): Map<string, GsPaint> {
  const out = new Map<string, GsPaint>();
  if (!resources) return out;
  const states = file.get(resources, 'ExtGState');
  if (!(states instanceof Map)) return out;
  for (const [name, value] of states) {
    const state = file.resolve(value);
    if (!(state instanceof Map)) continue;
    // §8.4.5 — a `gs` changes ONLY the parameters its dictionary names, so what
    // matters here is which ones it names, not only what they say.
    // alphatrans.pdf sets `ca` 0.5 in one state and then names a second that
    // carries `/CA` alone; read as "everything this one does not say is the
    // default", the second turned the fill opaque again and the blue square it
    // draws at half opacity buried what it covers.
    const ca = file.resolve(state.get('ca') ?? PDF_NULL);
    const alpha = typeof ca === 'number' && ca >= 0 && ca <= 1 ? ca : undefined;
    // `/BM` may be a name or an array of them, best first (§11.3.5).
    const bm = file.resolve(state.get('BM') ?? PDF_NULL);
    const first = Array.isArray(bm) ? file.resolve(bm[0] ?? PDF_NULL) : bm;
    const mode = first instanceof PdfName ? first.value : undefined;
    const darkens = mode === 'Multiply' || mode === 'Darken';
    // §11.3.5 — a mode that mixes two colours by a rule no anchored picture and
    // no run property can express. `Normal` and `Compatible` are the absence of
    // one; `Multiply` and `Darken` are carried as `darkens` and approximated.
    const blend =
      mode !== undefined && mode !== 'Normal' && mode !== 'Compatible' && !darkens
        ? mode
        : undefined;
    // §11.6.5 `/SMask` — a mask that varies the paint's opacity from place to
    // place, out of another group's luminosity or alpha. `/None` is the absence
    // of one, and STATING it takes a mask off.
    const statesMask = state.has('SMask');
    const masked = statesMask
      ? file.resolve(state.get('SMask') ?? PDF_NULL) instanceof Map
      : undefined;
    if (alpha === undefined && mode === undefined && !statesMask) continue;
    out.set(name, {
      ...(alpha !== undefined ? { alpha } : {}),
      ...(mode !== undefined ? { statesBlend: true, darkens } : {}),
      ...(blend !== undefined ? { blend } : {}),
      ...(masked !== undefined ? { masked } : {}),
    });
  }
  return out;
}

/**
 * What a `/ExtGState` says about paint that the reconstruction can carry.
 *
 * Every field is absent when the state does not NAME that parameter, because a
 * `gs` leaves what it does not name alone (§8.4.5).
 */
export interface GsPaint {
  /** §11.6.4.4 `/ca` — the constant fill alpha, where the state names one. */
  readonly alpha?: number;
  /** §11.3.5 — whether `/BM` is named at all, since naming `/Normal` ends a blend. */
  readonly statesBlend?: boolean;
  /** §11.3.5 `/BM` — the paint only darkens, so what it covers shows through. */
  readonly darkens?: boolean;
  /**
   * §11.3.5 `/BM` — a blend NOTHING downstream can perform, named so the loss
   * report can say which. `Normal` is no blend at all and is never named here.
   */
  readonly blend?: string;
  /**
   * §11.6.5 `/SMask` — the paint's opacity varies from place to place, out of
   * another group's luminosity or alpha. Nothing downstream has a mask like it.
   * `false` where the state names `/None`, which takes a mask off.
   */
  readonly masked?: boolean;
}

/**
 * §8.6 — how many components a colour space takes, and what they mean.
 *
 * `sc` / `scn` give bare numbers; only the space in force says whether `1 1 1`
 * is white or something else, and whether a lone `1` is white (DeviceGray) or
 * the colorant at full strength (Separation), which is usually black.
 */
export interface ColorSpaceInfo {
  readonly kind: 'gray' | 'rgb' | 'cmyk' | 'tint';
  readonly components: number;
  /**
   * §8.6.5.6/§8.6.5.7 — the CIE parameters, for a `CalGray` or `CalRGB` space.
   * Its numbers look like a device space's and are not: they mean what comes
   * out of this transform.
   */
  readonly cie?: CieSpace;
  /**
   * §8.6.6.4/§8.6.6.5 — for a `Separation` or `DeviceN`, the way OUT of it: the
   * tint transform and the space its numbers land in. Absent where the file
   * states a transform this cannot run, and then a tint is only "this much ink".
   */
  readonly tint?: {
    readonly transform: PdfFunction;
    readonly alternate: ColorSpaceInfo;
  };
}

/**
 * The spaces a `/ColorSpace` resource dictionary names, by name.
 *
 * @param file      The owning file.
 * @param resources The resource dictionary in force.
 * @returns Name → what the space comes to, for the spaces that are read.
 */
export function buildColorSpaceMap(
  file: PdfFile,
  resources: PdfDict | undefined,
): Map<string, ColorSpaceInfo> {
  const out = new Map<string, ColorSpaceInfo>();
  if (!resources) return out;
  const spaces = file.get(resources, 'ColorSpace');
  if (!(spaces instanceof Map)) return out;
  for (const [name, value] of spaces) {
    const info = colorSpaceAt(file, file.resolve(value), 0);
    if (info) out.set(name, info);
  }
  return out;
}

/** An alternate space may name another; this is where that stops. */
const MAX_ALTERNATE = 4;

/** What a colour space object comes to, or `undefined` for one not read. */
function colorSpaceAt(file: PdfFile, cs: PdfValue, depth: number): ColorSpaceInfo | undefined {
  if (cs instanceof PdfName) return byFamily(cs.value, 0);
  if (!Array.isArray(cs) || cs.length === 0) return undefined;
  const head = file.resolve(cs[0]!);
  if (!(head instanceof PdfName)) return undefined;
  if (head.value === 'ICCBased') {
    // §8.6.5.5 — the stream's `/N` is the component count, and that is all this
    // needs: an ICC profile of three components reads as RGB.
    const stream = file.resolve(cs[1] ?? PDF_NULL);
    const n = stream instanceof PdfStream ? file.get(stream.dict, 'N') : undefined;
    return byFamily('ICCBased', typeof n === 'number' ? n : 3);
  }
  if (head.value === 'Indexed') {
    // §8.6.6.3 — one operand, an index into a table. Reading the table is the
    // image path's business; a bare `sc` into one is rare and left alone.
    return undefined;
  }
  if (head.value === 'CalGray') {
    // §8.6.5.6 — one number through one gamma, which is unambiguous.
    //
    // `CalRGB` is NOT read this way, though the file states it the same way.
    // Its transform is well defined on paper and no two renderers agree on the
    // chromatic adaptation at the end of it: calrgb.pdf's neutral column comes
    // back light blue-grey from mutool, and neither adapting the stated white
    // to D65 (Bradford or von Kries) nor ignoring the white reproduces that.
    // Guessing at it moved the file 0.630 to 0.621 while making some of its
    // pages worse, so it keeps the device reading until there is something to
    // check an implementation against.
    const params = file.resolve(cs[1] ?? PDF_NULL);
    const cie = params instanceof Map ? cieParams(file, params, false) : undefined;
    const base = { kind: 'gray' as const, components: 1 };
    return cie ? { ...base, cie } : base;
  }
  if (head.value === 'Separation' || head.value === 'DeviceN') {
    const names = file.resolve(cs[1] ?? PDF_NULL);
    const n = head.value === 'Separation' ? 1 : Array.isArray(names) ? names.length : 1;
    // §8.6.6.4 — the space names its colorants, an ALTERNATE space, and the
    // transform between them; run it and the tint is a colour rather than an
    // amount of ink. devicen.pdf's three triangles are green, blue and red, and
    // read as ink at full strength all three came back black.
    const alternate =
      depth < MAX_ALTERNATE
        ? colorSpaceAt(file, file.resolve(cs[2] ?? PDF_NULL), depth + 1)
        : undefined;
    const transform = readFunction(file, cs[3]);
    if (alternate && transform && alternate.kind !== 'tint') {
      return { kind: 'tint', components: n, tint: { transform, alternate } };
    }
    return { kind: 'tint', components: n };
  }
  return byFamily(head.value, 0);
}

/** §8.6.5.6/§8.6.5.7 — the white point, gamma and matrix a CIE space states. */
function cieParams(file: PdfFile, params: PdfDict, rgb: boolean): CieSpace | undefined {
  const nums = (v: PdfValue | undefined): Array<number> =>
    Array.isArray(v)
      ? v.map((x) => file.resolve(x)).filter((x): x is number => typeof x === 'number')
      : [];
  const white = nums(file.get(params, 'WhitePoint'));
  if (white.length < 3 || !(white[1]! > 0)) return undefined;
  const gammaRaw = file.get(params, 'Gamma');
  const gamma = typeof gammaRaw === 'number' ? [gammaRaw] : nums(gammaRaw);
  const matrix = nums(file.get(params, 'Matrix'));
  return {
    white: [white[0]!, white[1]!, white[2]!],
    gamma: gamma.length > 0 ? gamma : rgb ? [1, 1, 1] : [1],
    // §8.6.5.7 — the identity is the default, which is also what a matrix of
    // the wrong length comes to.
    ...(rgb ? { matrix: matrix.length === 9 ? matrix : [1, 0, 0, 0, 1, 0, 0, 0, 1] } : {}),
  };
}

function byFamily(family: string, n: number): ColorSpaceInfo | undefined {
  switch (family) {
    case 'DeviceGray':
    case 'CalGray':
    case 'G':
      return { kind: 'gray', components: 1 };
    case 'DeviceRGB':
    case 'CalRGB':
    case 'Lab':
    case 'RGB':
      return { kind: 'rgb', components: 3 };
    case 'DeviceCMYK':
    case 'CMYK':
      return { kind: 'cmyk', components: 4 };
    case 'ICCBased':
      return n === 1
        ? { kind: 'gray', components: 1 }
        : n === 4
          ? { kind: 'cmyk', components: 4 }
          : { kind: 'rgb', components: 3 };
    default:
      return undefined;
  }
}
