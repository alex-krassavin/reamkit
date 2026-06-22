// Minimal valid ICC v2 RGB display profile (ICC.1:2001-04).
//
// PDF/A-1b requires an OutputIntent whose DestOutputProfile is a valid embedded
// ICC profile. We synthesise a compact matrix/TRC RGB profile from scratch —
// no external .icc asset needed — carrying the required tags for a 'mntr' /
// 'RGB ' / 'XYZ ' profile:
//   desc, wtpt, rXYZ, gXYZ, bXYZ, rTRC, gTRC, bTRC, cprt
//
// Primaries are the standard Bradford-adapted sRGB → D50 matrix columns and a
// gamma≈2.2 tone curve. This is colorimetrically "sRGB-ish"; PDF/A validators
// check structural validity (version, class, required tags) rather than exact
// colour, so the approximation is conformant.

// s15Fixed16: signed fixed-point, 16 integer bits, 16 fraction bits.
function s15Fixed16(value: number): number {
  return Math.round(value * 65536);
}

function writeU32BE(view: DataView, offset: number, value: number): void {
  view.setUint32(offset, value >>> 0, false);
}

function tag(s: string): number {
  return (
    ((s.charCodeAt(0) << 24) |
      (s.charCodeAt(1) << 16) |
      (s.charCodeAt(2) << 8) |
      s.charCodeAt(3)) >>>
    0
  );
}

// D50 PCS illuminant (ICC reference white).
const D50 = { x: 0.9642, y: 1.0, z: 0.8249 };
// Bradford-adapted sRGB primaries to D50 (the values in the canonical sRGB
// ICC profile).
const RED = { x: 0.43607, y: 0.22249, z: 0.0139 };
const GREEN = { x: 0.38515, y: 0.71687, z: 0.09708 };
const BLUE = { x: 0.14307, y: 0.06061, z: 0.7141 };

const DESC = 'sRGB-compatible';
const COPYRIGHT = 'Public Domain';

function buildTextDescription(text: string): Uint8Array {
  // textDescriptionType (ICC v2 §6.5.17): 'desc' + reserved + ASCII section +
  // Unicode section (empty) + ScriptCode section (empty).
  const ascii = text + '\0';
  const asciiBytes = new TextEncoder().encode(ascii);
  const size = 8 + 4 + asciiBytes.length + 4 + 4 + 2 + 1 + 67;
  const buf = new Uint8Array(align4(size));
  const view = new DataView(buf.buffer);
  writeU32BE(view, 0, tag('desc'));
  writeU32BE(view, 4, 0); // reserved
  writeU32BE(view, 8, asciiBytes.length); // ASCII count incl. NUL
  buf.set(asciiBytes, 12);
  // Unicode language code (0), Unicode count (0), ScriptCode code (0),
  // Mac count (0), Mac description (67 bytes of 0) — all already zero.
  return buf;
}

function buildText(text: string): Uint8Array {
  // textType (ICC v2 §6.5.18): 'text' + reserved + NUL-terminated ASCII.
  const bytes = new TextEncoder().encode(text + '\0');
  const size = 8 + bytes.length;
  const buf = new Uint8Array(align4(size));
  const view = new DataView(buf.buffer);
  writeU32BE(view, 0, tag('text'));
  writeU32BE(view, 4, 0);
  buf.set(bytes, 8);
  return buf;
}

function buildXYZ(x: number, y: number, z: number): Uint8Array {
  // XYZType (ICC v2 §6.5.26): 'XYZ ' + reserved + one XYZNumber (3× s15Fixed16).
  const buf = new Uint8Array(20);
  const view = new DataView(buf.buffer);
  writeU32BE(view, 0, tag('XYZ '));
  writeU32BE(view, 4, 0);
  view.setInt32(8, s15Fixed16(x), false);
  view.setInt32(12, s15Fixed16(y), false);
  view.setInt32(16, s15Fixed16(z), false);
  return buf;
}

function buildCurveGamma(gamma: number): Uint8Array {
  // curveType (ICC v2 §6.5.3) with a single u8Fixed8 gamma value.
  const buf = new Uint8Array(align4(14));
  const view = new DataView(buf.buffer);
  writeU32BE(view, 0, tag('curv'));
  writeU32BE(view, 4, 0);
  writeU32BE(view, 8, 1); // entry count
  view.setUint16(12, Math.round(gamma * 256), false); // u8Fixed8
  return buf;
}

function align4(n: number): number {
  return (n + 3) & ~3;
}

interface TagEntry {
  readonly sig: string;
  readonly data: Uint8Array;
}

/**
 * Synthesise a minimal valid ICC v2 matrix/TRC RGB display profile from scratch
 * (no external `.icc` asset), for use as a PDF/A-1b `OutputIntent`'s
 * `DestOutputProfile`. Colorimetrically "sRGB-ish"; PDF/A validators check
 * structural validity (version, class, required tags) rather than exact colour.
 *
 * @returns The encoded ICC profile bytes.
 */
export function buildSrgbIccProfile(): Uint8Array {
  const desc = buildTextDescription(DESC);
  const wtpt = buildXYZ(D50.x, D50.y, D50.z);
  const rXYZ = buildXYZ(RED.x, RED.y, RED.z);
  const gXYZ = buildXYZ(GREEN.x, GREEN.y, GREEN.z);
  const bXYZ = buildXYZ(BLUE.x, BLUE.y, BLUE.z);
  const trc = buildCurveGamma(2.2);
  const cprt = buildText(COPYRIGHT);

  // The three TRC tags can share one data block (identical curves) — ICC
  // allows multiple tag entries to point at the same offset.
  const entries: Array<TagEntry> = [
    { sig: 'desc', data: desc },
    { sig: 'wtpt', data: wtpt },
    { sig: 'rXYZ', data: rXYZ },
    { sig: 'gXYZ', data: gXYZ },
    { sig: 'bXYZ', data: bXYZ },
    { sig: 'rTRC', data: trc },
    { sig: 'gTRC', data: trc },
    { sig: 'bTRC', data: trc },
    { sig: 'cprt', data: cprt },
  ];

  const headerSize = 128;
  const tagCount = entries.length;
  const tagTableSize = 4 + tagCount * 12;
  let dataOffset = align4(headerSize + tagTableSize);

  // Lay out data blocks, deduplicating identical buffers (the shared TRC).
  const blockOffsets = new Map<Uint8Array, number>();
  const orderedBlocks: Array<{ data: Uint8Array; offset: number }> = [];
  for (const e of entries) {
    if (!blockOffsets.has(e.data)) {
      blockOffsets.set(e.data, dataOffset);
      orderedBlocks.push({ data: e.data, offset: dataOffset });
      dataOffset += align4(e.data.length);
    }
  }

  const totalSize = dataOffset;
  const out = new Uint8Array(totalSize);
  const view = new DataView(out.buffer);

  // --- Header (128 bytes) ---
  writeU32BE(view, 0, totalSize); // profile size
  writeU32BE(view, 4, 0); // preferred CMM (none)
  writeU32BE(view, 8, 0x02200000); // version 2.2.0
  writeU32BE(view, 12, tag('mntr')); // device class: display
  writeU32BE(view, 16, tag('RGB ')); // data colour space
  writeU32BE(view, 20, tag('XYZ ')); // PCS
  // date/time (bytes 24..35): fixed 2000-01-01T00:00:00 for determinism.
  view.setUint16(24, 2000, false);
  view.setUint16(26, 1, false);
  view.setUint16(28, 1, false);
  writeU32BE(view, 36, tag('acsp')); // profile file signature
  // platform (40), flags (44), manufacturer (48), model (52), attributes
  // (56..63), creator (80) — all left zero.
  writeU32BE(view, 64, 0); // rendering intent: perceptual
  // PCS illuminant (bytes 68..79): D50.
  view.setInt32(68, s15Fixed16(D50.x), false);
  view.setInt32(72, s15Fixed16(D50.y), false);
  view.setInt32(76, s15Fixed16(D50.z), false);

  // --- Tag table ---
  writeU32BE(view, headerSize, tagCount);
  let p = headerSize + 4;
  for (const e of entries) {
    const off = blockOffsets.get(e.data)!;
    writeU32BE(view, p, tag(e.sig));
    writeU32BE(view, p + 4, off);
    writeU32BE(view, p + 8, e.data.length); // unpadded size
    p += 12;
  }

  // --- Tag data ---
  for (const b of orderedBlocks) out.set(b.data, b.offset);

  return out;
}
