// §8.6.5.6/§8.6.5.7 — the CIE-based grey and RGB spaces.
//
// `CalGray` and `CalRGB` look like `DeviceGray` and `DeviceRGB` — one number, or
// three, between zero and one — and are not them. Each carries a white point, a
// gamma per component and (for RGB) a matrix into CIE XYZ, and the colour a
// viewer shows is what comes back out of that. calgray.pdf reads 0.258 against
// its source as a device grey and 0.044 as what it is.
//
// The RGB half is written here and NOT used: see `./shading`. Its way back from
// XYZ needs a chromatic adaptation, and no two renderers agree on which —
// calrgb.pdf's neutral column comes back light blue-grey from mutool and
// neither Bradford, nor von Kries, nor ignoring the white point reproduces
// that. It waits for something to check it against.

/** §8.6.5.6/§8.6.5.7 — what a CIE-based space states about itself. */
export interface CieSpace {
  /** `/WhitePoint` — the colour of the illuminant, in XYZ. */
  readonly white: readonly [number, number, number];
  /** `/Gamma` — one per component. */
  readonly gamma: ReadonlyArray<number>;
  /** §8.6.5.7 `/Matrix` — components to XYZ, column-major as the PDF states it. */
  readonly matrix?: ReadonlyArray<number>;
}

/**
 * One colour in a CIE-based space, as sRGB in 0..1.
 *
 * @param space The space's own parameters.
 * @param abc   Its components: one for a grey space, three for an RGB one.
 * @returns The sRGB triple a viewer shows for it.
 */
export function cieToSrgb(space: CieSpace, abc: ReadonlyArray<number>): [number, number, number] {
  const clamp01 = (v: number): number => Math.min(1, Math.max(0, v));
  if (space.matrix === undefined || abc.length < 3) {
    // §8.6.5.6 — a grey space: A^G scales the white point, so the result is a
    // neutral of that luminance. The sRGB transfer alone puts it on a screen.
    const a = clamp01(abc[0] ?? 0) ** (space.gamma[0] ?? 1);
    const v = srgbTransfer(a);
    return [v, v, v];
  }
  const m = space.matrix;
  const a = clamp01(abc[0]!) ** (space.gamma[0] ?? 1);
  const b = clamp01(abc[1]!) ** (space.gamma[1] ?? 1);
  const c = clamp01(abc[2]!) ** (space.gamma[2] ?? 1);
  // §8.6.5.7 — the matrix runs [XA YA ZA XB YB ZB XC YC ZC]: each triple is one
  // component's contribution, so the product is a column at a time.
  const x = (m[0] ?? 0) * a + (m[3] ?? 0) * b + (m[6] ?? 0) * c;
  const y = (m[1] ?? 0) * a + (m[4] ?? 0) * b + (m[7] ?? 0) * c;
  const z = (m[2] ?? 0) * a + (m[5] ?? 0) * b + (m[8] ?? 0) * c;
  const [ax, ay, az] = adaptToD65(x, y, z, space.white);
  // IEC 61966-2-1 — XYZ (D65) to linear sRGB.
  const r = 3.2404542 * ax - 1.5371385 * ay - 0.4985314 * az;
  const g = -0.969266 * ax + 1.8760108 * ay + 0.041556 * az;
  const bl = 0.0556434 * ax - 0.2040259 * ay + 1.0572252 * az;
  return [srgbTransfer(clamp01(r)), srgbTransfer(clamp01(g)), srgbTransfer(clamp01(bl))];
}

/**
 * §8.6.5.8 — one colour in a `Lab` space, as sRGB in 0..1.
 *
 * Lab is the one CIE space a file states in absolute terms: `L*` from 0 to 100
 * is lightness and `a*`/`b*` are the two opponent axes, and unlike `CalRGB`
 * (see the note above) there is no argument about the way out — every renderer
 * takes it through XYZ against the stated white and adapts to the screen's.
 * issue10339_reduced.pdf paints two grids of blue swatches through an `Indexed`
 * palette whose base is one, and read as anything else the page came back
 * blank.
 *
 * @param white The space's `/WhitePoint`, in XYZ.
 * @param lab   `L*`, `a*`, `b*`.
 * @returns The sRGB triple a viewer shows for it.
 */
export function labToSrgb(
  white: readonly [number, number, number],
  lab: readonly [number, number, number],
): [number, number, number] {
  const m = (Math.min(100, Math.max(0, lab[0])) + 16) / 116;
  const l = m + lab[1] / 500;
  const n = m - lab[2] / 200;
  const x = white[0] * inverseTransfer(l);
  const y = white[1] * inverseTransfer(m);
  const z = white[2] * inverseTransfer(n);
  const [ax, ay, az] = adaptToD65(x, y, z, white);
  const r = 3.2404542 * ax - 1.5371385 * ay - 0.4985314 * az;
  const g = -0.969266 * ax + 1.8760108 * ay + 0.041556 * az;
  const b = 0.0556434 * ax - 0.2040259 * ay + 1.0572252 * az;
  const clamp01 = (v: number): number => Math.min(1, Math.max(0, v));
  return [srgbTransfer(clamp01(r)), srgbTransfer(clamp01(g)), srgbTransfer(clamp01(b))];
}

/** §8.6.5.8 — the `g` of the Lab definition, which is a cube with a linear toe. */
function inverseTransfer(v: number): number {
  return v >= 6 / 29 ? v ** 3 : (108 / 841) * (v - 4 / 29);
}

/** The D65 white every sRGB screen is referred to. */
const D65: readonly [number, number, number] = [0.9505, 1, 1.089];

/**
 * Bradford chromatic adaptation from the space's own white to D65.
 *
 * A file may state any illuminant — calrgb.pdf states several — and a colour
 * measured against one white shown against another is the wrong colour. The
 * Bradford transform is the one every colour-managed renderer uses for this.
 */
function adaptToD65(
  x: number,
  y: number,
  z: number,
  white: readonly [number, number, number],
): [number, number, number] {
  const [sx, sy, sz] = coneResponse(white[0], white[1], white[2]);
  const [dx, dy, dz] = coneResponse(D65[0], D65[1], D65[2]);
  if (sx === 0 || sy === 0 || sz === 0) return [x, y, z];
  const [cx, cy, cz] = coneResponse(x, y, z);
  return fromConeResponse((cx * dx) / sx, (cy * dy) / sy, (cz * dz) / sz);
}

/** XYZ into the Bradford cone-response space. */
function coneResponse(x: number, y: number, z: number): [number, number, number] {
  return [
    0.8951 * x + 0.2664 * y - 0.1614 * z,
    -0.7502 * x + 1.7135 * y + 0.0367 * z,
    0.0389 * x - 0.0685 * y + 1.0296 * z,
  ];
}

/** …and back out of it. */
function fromConeResponse(l: number, m: number, s: number): [number, number, number] {
  return [
    0.9869929 * l - 0.1470543 * m + 0.1599627 * s,
    0.4323053 * l + 0.5183603 * m + 0.0492912 * s,
    -0.0085287 * l + 0.0400428 * m + 0.9684867 * s,
  ];
}

/** IEC 61966-2-1 — linear light to the sRGB signal a screen takes. */
function srgbTransfer(v: number): number {
  return v <= 0.0031308 ? 12.92 * v : 1.055 * v ** (1 / 2.4) - 0.055;
}
