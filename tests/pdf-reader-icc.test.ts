// §8.6.5.5 — an `/ICCBased` space states its own transform, and for anything
// but an sRGB-like profile that transform is not what a device space would make
// of the same numbers. franz_2.pdf is one flat rectangle painted `0.5 0.5 0.5`
// in Apple's "Generic RGB Profile" — gamma 1.8, its own primaries — and its one
// line of text reads "The background should be gray".

import { describe, expect, it } from 'vitest';

import { Ream } from '@/core/converter/ream';
import { iccTransform } from '@/pdf-reader/icc';

/** The D50-adapted sRGB primaries, which sum to D50 exactly. */
const PRIMARIES: ReadonlyArray<readonly [number, number, number]> = [
  [0.4360747, 0.2225045, 0.0139322],
  [0.3850649, 0.7168786, 0.0971045],
  [0.1430804, 0.0606169, 0.7141733],
];
const D50: readonly [number, number, number] = [0.9642, 1, 0.8249];

describe('ICC profile transform (§8.6.5.5)', () => {
  it('takes a matrix/TRC profile through its own gamma and primaries', () => {
    const transform = iccTransform(rgbProfile(1.8));
    expect(transform).toBeDefined();
    if (!transform) return;
    const [r, g, b] = transform([0.5, 0.5, 0.5]);
    // Equal components through primaries that sum to the white give that white
    // at 0.5^1.8 = 0.2872 of full, and sRGB shows that as 0.570 — 145 of 255,
    // which is what every colour-managed renderer puts on the screen.
    expect(r).toBeCloseTo(0.5697, 2);
    expect(g).toBeCloseTo(0.5697, 2);
    expect(b).toBeCloseTo(0.5697, 2);
  });

  it('leaves a profile whose form it does not read alone', () => {
    // A CMYK profile, and a truncated one: both keep the device reading.
    expect(iccTransform(rgbProfile(2.2, 'CMYK'))).toBeUndefined();
    expect(iccTransform(rgbProfile(2.2).subarray(0, 100))).toBeUndefined();
  });

  it('paints a fill in the colour the profile states', () => {
    const doc = Ream.parse(iccPdf(1.8));
    const shape = doc.flow.body.find((b) => b.kind === 'shape');
    expect(shape).toBeDefined();
    if (shape?.kind !== 'shape' || shape.shape.fill.kind !== 'solid') return;
    const hex = shape.shape.fill.colorHex ?? '';
    // 145 is what mutool and every other colour-managed renderer put on the
    // screen for franz_2.pdf; a point either way is this profile's primaries
    // being written to a fixed point rather than Apple's own.
    const grey = Number.parseInt(hex.slice(0, 2), 16);
    expect(grey).toBeGreaterThanOrEqual(144);
    expect(grey).toBeLessThanOrEqual(146);
    expect(hex.slice(2)).toBe(hex.slice(0, 2).repeat(2));
  });

  it('reads a gamma of 1 as the sRGB-like profile it is', () => {
    // The common case: a profile whose numbers a device space would have got
    // right anyway. Linear light at 0.5 is 0.735 on a screen.
    const transform = iccTransform(rgbProfile(1));
    expect(transform?.([0.5, 0.5, 0.5])[0]).toBeCloseTo(0.7354, 2);
  });
});

/** A one-page PDF filling a rectangle `0.5 0.5 0.5` in an `/ICCBased` space. */
function iccPdf(gamma: number): Uint8Array {
  const profile = rgbProfile(gamma);
  const content = '/Cs cs 0.5 0.5 0.5 scn 10 10 180 80 re f';
  const encoder = new TextEncoder();
  const objects: Array<string | Uint8Array> = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R ' +
      '/Resources << /ColorSpace << /Cs [/ICCBased 5 0 R] >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    profileObject(profile),
  ];
  const parts: Array<Uint8Array> = [encoder.encode('%PDF-1.7\n')];
  const offsets: Array<number> = [];
  let at = parts[0]!.length;
  objects.forEach((body, i) => {
    offsets.push(at);
    const open = encoder.encode(`${String(i + 1)} 0 obj\n`);
    const bytes = typeof body === 'string' ? encoder.encode(body) : body;
    const close = encoder.encode('\nendobj\n');
    parts.push(open, bytes, close);
    at += open.length + bytes.length + close.length;
  });
  let tail = `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) tail += `${String(off).padStart(10, '0')} 00000 n \n`;
  tail += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(at)}\n%%EOF\n`;
  parts.push(encoder.encode(tail));
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/** The `/ICCBased` stream, whose body is the profile itself. */
function profileObject(profile: Uint8Array): Uint8Array {
  const encoder = new TextEncoder();
  const head = encoder.encode(`<< /N 3 /Length ${String(profile.length)} >>\nstream\n`);
  const tail = encoder.encode('\nendstream');
  const out = new Uint8Array(head.length + profile.length + tail.length);
  out.set(head, 0);
  out.set(profile, head.length);
  out.set(tail, head.length + profile.length);
  return out;
}

/**
 * A matrix/TRC display profile: the D50-adapted sRGB primaries and one gamma
 * shared by all three channels — the shape of every display profile there is.
 */
function rgbProfile(gamma: number, space = 'RGB '): Uint8Array {
  const tags: Array<{ sig: string; body: Uint8Array }> = [
    { sig: 'wtpt', body: xyzType(D50) },
    { sig: 'rXYZ', body: xyzType(PRIMARIES[0]!) },
    { sig: 'gXYZ', body: xyzType(PRIMARIES[1]!) },
    { sig: 'bXYZ', body: xyzType(PRIMARIES[2]!) },
    { sig: 'rTRC', body: curvType(gamma) },
    { sig: 'gTRC', body: curvType(gamma) },
    { sig: 'bTRC', body: curvType(gamma) },
  ];
  const tableSize = 4 + tags.length * 12;
  let at = 128 + tableSize;
  const total = tags.reduce((n, t) => n + t.body.length, at);
  const out = new Uint8Array(total);
  const view = new DataView(out.buffer);
  view.setUint32(0, total);
  out.set(ascii(space), 16); // the profile's own colour space
  out.set(ascii('XYZ '), 20); // the PCS
  out.set(ascii('acsp'), 36);
  view.setUint32(128, tags.length);
  tags.forEach((t, i) => {
    const entry = 132 + i * 12;
    out.set(ascii(t.sig), entry);
    view.setUint32(entry + 4, at);
    view.setUint32(entry + 8, t.body.length);
    out.set(t.body, at);
    at += t.body.length;
  });
  return out;
}

/** §10.31 `XYZType`. */
function xyzType(xyz: readonly [number, number, number]): Uint8Array {
  const out = new Uint8Array(20);
  const view = new DataView(out.buffer);
  out.set(ascii('XYZ '), 0);
  xyz.forEach((v, i) => {
    view.setInt32(8 + i * 4, Math.round(v * 65536));
  });
  return out;
}

/** §10.6 `curveType` with one entry, which states a gamma as u8Fixed8. */
function curvType(gamma: number): Uint8Array {
  const out = new Uint8Array(14);
  const view = new DataView(out.buffer);
  out.set(ascii('curv'), 0);
  view.setUint32(8, 1);
  view.setUint16(12, Math.round(gamma * 256));
  return out;
}

function ascii(text: string): Uint8Array {
  return Uint8Array.from([...text].map((c) => c.charCodeAt(0)));
}
