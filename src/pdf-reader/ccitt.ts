// E-PDF EP15 — CCITT Group 3 / Group 4 fax decoding (ITU-T T.4 / T.6, as used by
// PDF's /CCITTFaxDecode filter, ISO 32000-1 §7.4.6). Decodes a bilevel scan into
// a packed 1-bit-per-pixel bitmap (bit 1 = black), which image-decode.ts turns
// into a DeviceGray raster. Supports Group 4 (K < 0, the common PDF case, pure
// two-dimensional T.6 coding) and Group 3 one-dimensional (K = 0); Group 3 mixed
// 2-D (K > 0) is not decoded (its line tag bits + EOL framing are rare in PDF).
//
// The white/black run-length codes are the fixed modified-Huffman tables from
// T.4; the two-dimensional vertical/pass/horizontal mode codes are from T.6.

/**
 * The `/CCITTFaxDecode` parameters that drive {@link decodeCcitt} (ISO 32000-1
 * §7.4.6), pulled from the filter's `/DecodeParms`.
 */
export interface CcittParams {
  /** `/K` — `<0` Group 4, `0` Group 3 1-D, `>0` Group 3 2-D (unsupported). */
  readonly k: number;
  /** `/Columns` (pixels per row). */
  readonly columns: number;
  /** `/Rows` (or the image `/Height`) — the row count to decode. */
  readonly rows: number;
  /** `/EncodedByteAlign` — pad each row to a byte boundary. */
  readonly byteAlign: boolean;
}

/**
 * Decode a CCITT Group 3 / Group 4 fax stream into a packed 1-bit-per-pixel
 * bitmap.
 *
 * @param data   The raw fax codestream (any wrapping filters already stripped).
 * @param params The `/CCITTFaxDecode` parameters.
 * @returns The packed bitmap (`rowBytes × rows`, bit 1 = black, MSB first), or
 *   `undefined` when the stream cannot be decoded (e.g. Group 3 2-D, or nothing
 *   decoded at all).
 */
export function decodeCcitt(data: Uint8Array, params: CcittParams): Uint8Array | undefined {
  const { k, columns, rows, byteAlign } = params;
  if (k > 0) return undefined; // Group 3 two-dimensional — not supported
  if (columns <= 0 || rows <= 0 || columns > 1 << 20) return undefined;

  const reader = new BitReader(data);
  const rowBytes = (columns + 7) >> 3;
  const out = new Uint8Array(rowBytes * rows);
  // The reference line for row 0 is an imaginary all-white line: its only
  // changing elements are the two sentinels at `columns`.
  let ref: Array<number> = [columns, columns];

  for (let y = 0; y < rows; y++) {
    if (byteAlign && y > 0) reader.align();
    const cur = k < 0 ? decode2D(reader, ref, columns) : decode1D(reader, columns);
    if (!cur) {
      if (y === 0) return undefined; // nothing decoded at all
      break; // EOFB / truncated — keep the rows we have (rest stay white)
    }
    expandRow(cur, columns, out, y * rowBytes);
    cur.push(columns, columns); // sentinels for the next line's b1/b2 lookups
    ref = cur;
  }
  return out;
}

// --- one row of changing elements -------------------------------------------

// Group 3 1-D: alternating white/black runs across the line. Returns the
// positions where the colour changes (starting white), or null at end-of-data.
function decode1D(reader: BitReader, columns: number): Array<number> | null {
  const cur: Array<number> = [];
  let pos = 0;
  let color = 0; // 0 white, 1 black
  while (pos < columns) {
    const run = readRun(reader, color);
    if (run < 0) return cur.length > 0 ? cur : null;
    pos = Math.min(pos + run, columns);
    cur.push(pos);
    color ^= 1;
  }
  return cur;
}

// Group 4 / T.6: two-dimensional coding against the reference line. Returns the
// coding line's changing elements, or null at EOFB / a decoding error.
function decode2D(reader: BitReader, ref: Array<number>, columns: number): Array<number> | null {
  const cur: Array<number> = [];
  let a0 = -1;
  let color = 0;
  while (a0 < columns) {
    const [b1, b2] = findB(ref, a0, color, columns);
    const mode = readMode(reader);
    if (!mode) return cur.length > 0 || a0 >= 0 ? cur : null;
    if (mode.kind === 'pass') {
      a0 = b2; // colour unchanged; pixels a0..b2 keep the current colour
    } else if (mode.kind === 'horizontal') {
      const start = a0 < 0 ? 0 : a0;
      const r1 = readRun(reader, color);
      const r2 = readRun(reader, color ^ 1);
      if (r1 < 0 || r2 < 0) return cur.length > 0 ? cur : null;
      const a1 = Math.min(start + r1, columns);
      const a2 = Math.min(a1 + r2, columns);
      cur.push(a1, a2);
      a0 = a2; // colour unchanged (two runs)
    } else {
      const a1 = Math.min(Math.max(b1 + mode.d, 0), columns); // vertical V(d)
      cur.push(a1);
      a0 = a1;
      color ^= 1;
    }
  }
  return cur;
}

// b1: the first changing element on the reference line to the right of a0 and of
// the opposite colour to a0; b2: the next changing element after b1. Reference
// elements alternate colour — element at an even index is a white→black change
// (its pixel is black), an odd index is black→white (its pixel is white).
function findB(ref: Array<number>, a0: number, color: number, columns: number): [number, number] {
  let i = 0;
  while (i < ref.length && ref[i]! <= a0) i++;
  // The element at index i is black when i is even. b1 must be the opposite
  // colour to a0: opposite-of-white is black (even i), opposite-of-black is
  // white (odd i). Skip one element if the parity is wrong.
  const wantEven = color === 0; // a0 white → want black (even) element
  if (i < ref.length && (i % 2 === 0) !== wantEven) i++;
  const b1 = i < ref.length ? ref[i]! : columns;
  const b2 = i + 1 < ref.length ? ref[i + 1]! : columns;
  return [b1, b2];
}

// Paint the black runs of a changing-element line into the packed output row.
function expandRow(cur: Array<number>, columns: number, out: Uint8Array, rowOff: number): void {
  let prev = 0;
  let color = 0;
  for (const cRaw of cur) {
    const c = Math.min(cRaw, columns);
    if (color === 1) setBlack(out, rowOff, prev, c);
    prev = c;
    color ^= 1;
    if (prev >= columns) return;
  }
  if (color === 1 && prev < columns) setBlack(out, rowOff, prev, columns);
}

function setBlack(out: Uint8Array, rowOff: number, from: number, to: number): void {
  for (let x = from; x < to; x++) out[rowOff + (x >> 3)]! |= 0x80 >> (x & 7);
}

// --- code reading -----------------------------------------------------------

// A run length: makeup codes (≥64) accumulate until a terminating code (<64).
function readRun(reader: BitReader, color: number): number {
  const table = color === 1 ? BLACK : WHITE;
  let total = 0;
  for (let guard = 0; guard < 64; guard++) {
    const run = readCode(reader, table);
    if (run < 0) return -1;
    total += run;
    if (run < 64) return total; // terminating code
  }
  return -1; // runaway makeup chain
}

/** A decoded T.6 two-dimensional mode code: pass, horizontal, or vertical V(d). */
export interface Mode {
  readonly kind: 'pass' | 'horizontal' | 'vertical';
  /** Vertical offset (−3..3); meaningful only when `kind` is `'vertical'`. */
  readonly d: number;
}

function readMode(reader: BitReader): Mode | undefined {
  let code = 0;
  for (let len = 1; len <= 7; len++) {
    const b = reader.readBit();
    if (b < 0) return undefined;
    code = (code << 1) | b;
    const m = MODES.get((len << 8) | code);
    if (m) return m;
  }
  return undefined; // EOL / extension / EOFB — treat as end of line
}

interface RunTable {
  readonly map: Map<number, number>;
  readonly maxLen: number;
}

function readCode(reader: BitReader, table: RunTable): number {
  let code = 0;
  for (let len = 1; len <= table.maxLen; len++) {
    const b = reader.readBit();
    if (b < 0) return -1;
    code = (code << 1) | b;
    const run = table.map.get((len << 16) | code);
    if (run !== undefined) return run;
  }
  return -1;
}

class BitReader {
  private bytePos = 0;
  private bitPos = 0;
  constructor(private readonly data: Uint8Array) {}
  readBit(): number {
    if (this.bytePos >= this.data.length) return -1;
    const bit = (this.data[this.bytePos]! >> (7 - this.bitPos)) & 1;
    if (++this.bitPos === 8) {
      this.bitPos = 0;
      this.bytePos++;
    }
    return bit;
  }
  align(): void {
    if (this.bitPos !== 0) {
      this.bitPos = 0;
      this.bytePos++;
    }
  }
}

// --- the fixed T.4 / T.6 code tables ----------------------------------------

/**
 * T.4 white-run modified-Huffman codes as `[run, bit-string]` (terminating runs
 * 0..63, make-up runs multiples of 64). Exported so the test can build an
 * independent encoder and validate the codes form a valid prefix set.
 */
export const WHITE_CODES: ReadonlyArray<readonly [number, string]> = [
  [0, '00110101'],
  [1, '000111'],
  [2, '0111'],
  [3, '1000'],
  [4, '1011'],
  [5, '1100'],
  [6, '1110'],
  [7, '1111'],
  [8, '10011'],
  [9, '10100'],
  [10, '00111'],
  [11, '01000'],
  [12, '001000'],
  [13, '000011'],
  [14, '110100'],
  [15, '110101'],
  [16, '101010'],
  [17, '101011'],
  [18, '0100111'],
  [19, '0001100'],
  [20, '0001000'],
  [21, '0010111'],
  [22, '0000011'],
  [23, '0000100'],
  [24, '0101000'],
  [25, '0101011'],
  [26, '0010011'],
  [27, '0100100'],
  [28, '0011000'],
  [29, '00000010'],
  [30, '00000011'],
  [31, '00011010'],
  [32, '00011011'],
  [33, '00010010'],
  [34, '00010011'],
  [35, '00010100'],
  [36, '00010101'],
  [37, '00010110'],
  [38, '00010111'],
  [39, '00101000'],
  [40, '00101001'],
  [41, '00101010'],
  [42, '00101011'],
  [43, '00101100'],
  [44, '00101101'],
  [45, '00000100'],
  [46, '00000101'],
  [47, '00001010'],
  [48, '00001011'],
  [49, '01010010'],
  [50, '01010011'],
  [51, '01010100'],
  [52, '01010101'],
  [53, '00100100'],
  [54, '00100101'],
  [55, '01011000'],
  [56, '01011001'],
  [57, '01011010'],
  [58, '01011011'],
  [59, '01001010'],
  [60, '01001011'],
  [61, '00110010'],
  [62, '00110011'],
  [63, '00110100'],
  [64, '11011'],
  [128, '10010'],
  [192, '010111'],
  [256, '0110111'],
  [320, '00110110'],
  [384, '00110111'],
  [448, '01100100'],
  [512, '01100101'],
  [576, '01101000'],
  [640, '01100111'],
  [704, '011001100'],
  [768, '011001101'],
  [832, '011010010'],
  [896, '011010011'],
  [960, '011010100'],
  [1024, '011010101'],
  [1088, '011010110'],
  [1152, '011010111'],
  [1216, '011011000'],
  [1280, '011011001'],
  [1344, '011011010'],
  [1408, '011011011'],
  [1472, '010011000'],
  [1536, '010011001'],
  [1600, '010011010'],
  [1664, '011000'],
  [1728, '010011011'],
];

export const BLACK_CODES: ReadonlyArray<readonly [number, string]> = [
  [0, '0000110111'],
  [1, '010'],
  [2, '11'],
  [3, '10'],
  [4, '011'],
  [5, '0011'],
  [6, '0010'],
  [7, '00011'],
  [8, '000101'],
  [9, '000100'],
  [10, '0000100'],
  [11, '0000101'],
  [12, '0000111'],
  [13, '00000100'],
  [14, '00000111'],
  [15, '000011000'],
  [16, '0000010111'],
  [17, '0000011000'],
  [18, '0000001000'],
  [19, '00001100111'],
  [20, '00001101000'],
  [21, '00001101100'],
  [22, '00000110111'],
  [23, '00000101000'],
  [24, '00000010111'],
  [25, '00000011000'],
  [26, '000011001010'],
  [27, '000011001011'],
  [28, '000011001100'],
  [29, '000011001101'],
  [30, '000001101000'],
  [31, '000001101001'],
  [32, '000001101010'],
  [33, '000001101011'],
  [34, '000011010010'],
  [35, '000011010011'],
  [36, '000011010100'],
  [37, '000011010101'],
  [38, '000011010110'],
  [39, '000011010111'],
  [40, '000001101100'],
  [41, '000001101101'],
  [42, '000011011010'],
  [43, '000011011011'],
  [44, '000001010100'],
  [45, '000001010101'],
  [46, '000001010110'],
  [47, '000001010111'],
  [48, '000001100100'],
  [49, '000001100101'],
  [50, '000001010010'],
  [51, '000001010011'],
  [52, '000000100100'],
  [53, '000000110111'],
  [54, '000000111000'],
  [55, '000000100111'],
  [56, '000000101000'],
  [57, '000001011000'],
  [58, '000001011001'],
  [59, '000000101011'],
  [60, '000000101100'],
  [61, '000001011010'],
  [62, '000001100110'],
  [63, '000001100111'],
  [64, '0000001111'],
  [128, '000011001000'],
  [192, '000011001001'],
  [256, '000001011011'],
  [320, '000000110011'],
  [384, '000000110100'],
  [448, '000000110101'],
  [512, '0000001101100'],
  [576, '0000001101101'],
  [640, '0000001001010'],
  [704, '0000001001011'],
  [768, '0000001001100'],
  [832, '0000001001101'],
  [896, '0000001110010'],
  [960, '0000001110011'],
  [1024, '0000001110100'],
  [1088, '0000001110101'],
  [1152, '0000001110110'],
  [1216, '0000001110111'],
  [1280, '0000001010010'],
  [1344, '0000001010011'],
  [1408, '0000001010100'],
  [1472, '0000001010101'],
  [1536, '0000001011010'],
  [1600, '0000001011011'],
  [1664, '0000001100100'],
  [1728, '0000001100101'],
];

export const SHARED_MAKEUP: ReadonlyArray<readonly [number, string]> = [
  [1792, '00000001000'],
  [1856, '00000001100'],
  [1920, '00000001101'],
  [1984, '000000010010'],
  [2048, '000000010011'],
  [2112, '000000010100'],
  [2176, '000000010101'],
  [2240, '000000010110'],
  [2304, '000000010111'],
  [2368, '000000011100'],
  [2432, '000000011101'],
  [2496, '000000011110'],
  [2560, '000000011111'],
];

export const MODE_CODES: ReadonlyArray<readonly [string, Mode]> = [
  ['0001', { kind: 'pass', d: 0 }],
  ['001', { kind: 'horizontal', d: 0 }],
  ['1', { kind: 'vertical', d: 0 }],
  ['011', { kind: 'vertical', d: 1 }],
  ['000011', { kind: 'vertical', d: 2 }],
  ['0000011', { kind: 'vertical', d: 3 }],
  ['010', { kind: 'vertical', d: -1 }],
  ['000010', { kind: 'vertical', d: -2 }],
  ['0000010', { kind: 'vertical', d: -3 }],
];

function buildRunTable(
  ...groups: ReadonlyArray<ReadonlyArray<readonly [number, string]>>
): RunTable {
  const map = new Map<number, number>();
  let maxLen = 0;
  for (const group of groups) {
    for (const [run, bits] of group) {
      map.set((bits.length << 16) | parseInt(bits, 2), run);
      maxLen = Math.max(maxLen, bits.length);
    }
  }
  return { map, maxLen };
}

const WHITE = buildRunTable(WHITE_CODES, SHARED_MAKEUP);
const BLACK = buildRunTable(BLACK_CODES, SHARED_MAKEUP);
const MODES: Map<number, Mode> = new Map(
  MODE_CODES.map(([bits, m]) => [(bits.length << 8) | parseInt(bits, 2), m]),
);
