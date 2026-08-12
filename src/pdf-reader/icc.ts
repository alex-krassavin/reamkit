// §8.6.5.5 / ICC.1:2010 — the profile an `/ICCBased` space carries, as a
// transform to sRGB.
//
// The reader has always taken `/N` from such a space and read three components
// as RGB, which is right for the sRGB-like profiles most files carry and wrong
// for every other. franz_2.pdf is one flat rectangle painted `0.5 0.5 0.5` in
// Apple's "Generic RGB Profile" — gamma 1.8, its own primaries — and its whole
// page is the difference: 128 against the 145 every colour-managed renderer
// shows, and the file's one line of text says "The background should be gray".
//
// Only the matrix/TRC form is read, which is what a display profile is: three
// colourant tags giving the primaries in XYZ, three tone curves, and a white
// point. A profile that states its transform as a lookup table (`A2B0`) is a
// different machine and keeps the device reading.

/** How a profile turns its own components into sRGB, each 0..1. */
export type IccTransform = (comps: ReadonlyArray<number>) => [number, number, number];

/**
 * Read an ICC profile into a transform to sRGB.
 *
 * @param profile The raw profile bytes, as `/ICCBased`'s stream holds them.
 * @returns The transform, or `undefined` for a profile whose form this does not
 *          read — a CMYK one, a table-based one, or a damaged one.
 */
export function iccTransform(profile: Uint8Array): IccTransform | undefined {
  if (profile.length < HEADER_SIZE + 4) return undefined;
  const view = new DataView(profile.buffer, profile.byteOffset, profile.byteLength);
  // The profile's own colour space, at byte 16: only these two are read.
  const space = tag(profile, 16);
  if (space !== 'RGB ' && space !== 'GRAY') return undefined;
  const tags = tagTable(profile, view);
  if (!tags) return undefined;
  const white = xyzTag(profile, view, tags.get('wtpt')) ?? D50;

  if (space === 'GRAY') {
    const curve = curveTag(profile, view, tags.get('kTRC'));
    if (!curve) return undefined;
    return (comps): [number, number, number] => {
      const y = curve(clamp01(comps[0] ?? 0));
      // §6.3.4.2 — a grey profile's tone curve gives luminance against its own
      // white, which is the neutral of that luminance.
      const [x, yy, z] = adapt([white[0] * y, white[1] * y, white[2] * y], white);
      return toSrgb(x, yy, z);
    };
  }

  const r = xyzTag(profile, view, tags.get('rXYZ'));
  const g = xyzTag(profile, view, tags.get('gXYZ'));
  const b = xyzTag(profile, view, tags.get('bXYZ'));
  if (!r || !g || !b) return undefined;
  const rc = curveTag(profile, view, tags.get('rTRC'));
  const gc = curveTag(profile, view, tags.get('gTRC'));
  const bc = curveTag(profile, view, tags.get('bTRC'));
  if (!rc || !gc || !bc) return undefined;
  return (comps): [number, number, number] => {
    const lr = rc(clamp01(comps[0] ?? 0));
    const lg = gc(clamp01(comps[1] ?? 0));
    const lb = bc(clamp01(comps[2] ?? 0));
    // The colourant tags ARE the matrix: each is one primary in XYZ.
    const x = r[0] * lr + g[0] * lg + b[0] * lb;
    const y = r[1] * lr + g[1] * lg + b[1] * lb;
    const z = r[2] * lr + g[2] * lg + b[2] * lb;
    const [ax, ay, az] = adapt([x, y, z], D50);
    return toSrgb(ax, ay, az);
  };
}

/** The profile header is fixed-length; the tag table follows it. */
const HEADER_SIZE = 128;

/** ICC.1:2010 §7.2.16 — the PCS every matrix/TRC profile is referred to. */
const D50: readonly [number, number, number] = [0.9642, 1, 0.8249];

/** The white every sRGB screen is referred to. */
const D65: readonly [number, number, number] = [0.9505, 1, 1.089];

/** A four-byte signature, as the profile writes them. */
function tag(data: Uint8Array, at: number): string {
  return String.fromCharCode(
    data[at] ?? 0,
    data[at + 1] ?? 0,
    data[at + 2] ?? 0,
    data[at + 3] ?? 0,
  );
}

/** §7.3 — the tag table: a count, then twelve bytes of signature/offset/size each. */
function tagTable(
  data: Uint8Array,
  view: DataView,
): Map<string, { offset: number; size: number }> | undefined {
  const count = view.getUint32(HEADER_SIZE);
  if (count === 0 || count > MAX_TAGS) return undefined;
  const out = new Map<string, { offset: number; size: number }>();
  for (let i = 0; i < count; i++) {
    const at = HEADER_SIZE + 4 + i * 12;
    if (at + 12 > data.length) break;
    const offset = view.getUint32(at + 4);
    const size = view.getUint32(at + 8);
    if (offset + size > data.length) continue;
    out.set(tag(data, at), { offset, size });
  }
  return out.size > 0 ? out : undefined;
}

/** No profile carries more tags than this; a bigger count is a damaged file. */
const MAX_TAGS = 256;

/** §10.31 `XYZType` — a signature, four reserved bytes, then three s15Fixed16. */
function xyzTag(
  data: Uint8Array,
  view: DataView,
  where: { offset: number; size: number } | undefined,
): [number, number, number] | undefined {
  if (!where || where.size < 20 || tag(data, where.offset) !== 'XYZ ') return undefined;
  const at = where.offset + 8;
  return [s15(view, at), s15(view, at + 4), s15(view, at + 8)];
}

/** The s15Fixed16 every ICC number is written as. */
function s15(view: DataView, at: number): number {
  return view.getInt32(at) / 65536;
}

/**
 * §10.6/§10.18 — a tone curve, as the function that linearises one component.
 *
 * `curv` is a table: no entries is the identity, one is a gamma written as
 * u8Fixed8, and more is a sampled curve read by interpolation. `para` states
 * one of five formulas by their parameters, of which the sRGB-style type 3 is
 * the one every modern profile writes.
 */
function curveTag(
  data: Uint8Array,
  view: DataView,
  where: { offset: number; size: number } | undefined,
): ((v: number) => number) | undefined {
  if (!where || where.size < 12) return undefined;
  const kind = tag(data, where.offset);
  if (kind === 'curv') {
    const count = view.getUint32(where.offset + 8);
    if (count === 0) return (v): number => v;
    if (count === 1) {
      const gamma = view.getUint16(where.offset + 12) / 256;
      return (v): number => v ** gamma;
    }
    if (where.offset + 12 + count * 2 > data.length || count > MAX_CURVE_POINTS) return undefined;
    const table = new Float64Array(count);
    for (let i = 0; i < count; i++) table[i] = view.getUint16(where.offset + 12 + i * 2) / 65535;
    return (v): number => {
      const at = clamp01(v) * (count - 1);
      const i = Math.floor(at);
      const lo = table[i] ?? 0;
      const hi = table[Math.min(i + 1, count - 1)] ?? lo;
      return lo + (hi - lo) * (at - i);
    };
  }
  if (kind === 'para') {
    const type = view.getUint16(where.offset + 8);
    const p = (i: number): number => s15(view, where.offset + 12 + i * 4);
    const g = p(0);
    // §10.18 — the five formulas, each adding one parameter to the last.
    if (type === 0) return (v): number => v ** g;
    if (type === 1) {
      const [a, b] = [p(1), p(2)];
      return (v): number => (v >= -b / a ? (a * v + b) ** g : 0);
    }
    if (type === 2) {
      const [a, b, c] = [p(1), p(2), p(3)];
      return (v): number => (v >= -b / a ? (a * v + b) ** g + c : c);
    }
    if (type === 3) {
      const [a, b, c, d] = [p(1), p(2), p(3), p(4)];
      return (v): number => (v >= d ? (a * v + b) ** g : c * v);
    }
    if (type === 4) {
      const [a, b, c, d, e, f] = [p(1), p(2), p(3), p(4), p(5), p(6)];
      return (v): number => (v >= d ? (a * v + b) ** g + e : c * v + f);
    }
  }
  return undefined;
}

/** No tone curve is sampled at more points than this. */
const MAX_CURVE_POINTS = 65536;

/** Bradford adaptation from the profile's white to the screen's. */
function adapt(
  xyz: readonly [number, number, number],
  from: readonly [number, number, number],
): [number, number, number] {
  const cone = (x: number, y: number, z: number): [number, number, number] => [
    0.8951 * x + 0.2664 * y - 0.1614 * z,
    -0.7502 * x + 1.7135 * y + 0.0367 * z,
    0.0389 * x - 0.0685 * y + 1.0296 * z,
  ];
  const [sx, sy, sz] = cone(from[0], from[1], from[2]);
  const [dx, dy, dz] = cone(D65[0], D65[1], D65[2]);
  const [cx, cy, cz] = cone(xyz[0], xyz[1], xyz[2]);
  const [ax, ay, az] = [
    cx * (sx === 0 ? 1 : dx / sx),
    cy * (sy === 0 ? 1 : dy / sy),
    cz * (sz === 0 ? 1 : dz / sz),
  ];
  // Back out of cone space — the inverse of the matrix above.
  return [
    0.9869929 * ax - 0.1470543 * ay + 0.1599627 * az,
    0.4323053 * ax + 0.5183603 * ay + 0.0492912 * az,
    -0.0085287 * ax + 0.0400428 * ay + 0.9684867 * az,
  ];
}

/** IEC 61966-2-1 — XYZ (D65) to sRGB, gamma and all. */
function toSrgb(x: number, y: number, z: number): [number, number, number] {
  const r = 3.2404542 * x - 1.5371385 * y - 0.4985314 * z;
  const g = -0.969266 * x + 1.8760108 * y + 0.041556 * z;
  const b = 0.0556434 * x - 0.2040259 * y + 1.0572252 * z;
  return [transfer(clamp01(r)), transfer(clamp01(g)), transfer(clamp01(b))];
}

function transfer(v: number): number {
  return v <= 0.0031308 ? 12.92 * v : 1.055 * v ** (1 / 2.4) - 0.055;
}

function clamp01(v: number): number {
  return Math.min(1, Math.max(0, v));
}
