// The MSB-first LZW that TIFF 6.0 §13 and PDF (ISO 32000-1 §7.4.4.2) share —
// the same coder, described twice: variable-width codes 9→12 bits, a clear code
// (256) and an end-of-data code (257), the table restarting at 258.
//
// TIFF's own encoders widen the code one step EARLY (the "early change" PDF
// spells out as a parameter), which is why the two specifications describe the
// identical bitstream; the flag is the only dial between them.
//
// GIF's LZW is a different animal — LSB-first, its own code size — and lives
// with the GIF decoder in images.ts.

/** How to run {@link lzwDecodeMsb}: the early-change flag and an output cap. */
export interface LzwOptions {
  /**
   * Widen the code one step BEFORE the table fills (PDF `/EarlyChange`,
   * TIFF's own convention). 1 by default, matching both specifications' norm.
   */
  readonly earlyChange?: number;
  /** Stop after this many output bytes — the guard against a decompression bomb. */
  readonly limit: number;
}

/**
 * Decode an MSB-first LZW stream (TIFF 6.0 §13 / ISO 32000-1 §7.4.4.2). The
 * KwKwK case (a code not yet in the table) repeats the previous string plus its
 * own first byte. Truncated input ends the stream where it stops rather than
 * throwing: a partly-decoded strip still prints.
 */
export function lzwDecodeMsb(data: Uint8Array, options: LzwOptions): Uint8Array {
  const earlyChange = options.earlyChange ?? 1;
  let out = new Uint8Array(1 << 12); // grows by doubling
  let outLen = 0;
  const emit = (e: Uint8Array): boolean => {
    if (outLen + e.length > options.limit) return false;
    if (outLen + e.length > out.length) {
      let cap = out.length * 2;
      while (cap < outLen + e.length) cap *= 2;
      const grown = new Uint8Array(cap);
      grown.set(out.subarray(0, outLen));
      out = grown;
    }
    out.set(e, outLen);
    outLen += e.length;
    return true;
  };

  let bitBuffer = 0;
  let bitCount = 0;
  let pos = 0;
  let codeLength = 9;
  const readCode = (): number => {
    while (bitCount < codeLength) {
      if (pos >= data.length) return -1;
      bitBuffer = ((bitBuffer << 8) | data[pos++]!) >>> 0;
      bitCount += 8;
    }
    bitCount -= codeLength;
    return (bitBuffer >>> bitCount) & ((1 << codeLength) - 1);
  };

  const dict = new Array<Uint8Array>(4096);
  let nextCode = 258;
  let prev = -1;
  const reset = (): void => {
    for (let i = 0; i < 256; i++) dict[i] = Uint8Array.of(i);
    nextCode = 258;
    codeLength = 9;
    prev = -1;
  };
  reset();

  for (;;) {
    const code = readCode();
    if (code < 0 || code === 257) break; // out of data / end-of-data
    if (code === 256) {
      reset();
      continue;
    }
    if (prev < 0) {
      const first = dict[code];
      if (!first || !emit(first)) break;
      prev = code;
      continue;
    }
    const prevEntry = dict[prev]!;
    // A known code uses its entry; an as-yet-unassigned code is KwKwK.
    let entry = code < nextCode ? dict[code] : undefined;
    if (!entry) {
      entry = new Uint8Array(prevEntry.length + 1);
      entry.set(prevEntry);
      entry[prevEntry.length] = prevEntry[0]!;
    }
    if (!emit(entry)) break;
    if (nextCode < 4096) {
      const added = new Uint8Array(prevEntry.length + 1);
      added.set(prevEntry);
      added[prevEntry.length] = entry[0]!;
      dict[nextCode++] = added;
      if (nextCode + earlyChange === 512) codeLength = 10;
      else if (nextCode + earlyChange === 1024) codeLength = 11;
      else if (nextCode + earlyChange === 2048) codeLength = 12;
    }
    prev = code;
  }
  return out.subarray(0, outLen);
}
