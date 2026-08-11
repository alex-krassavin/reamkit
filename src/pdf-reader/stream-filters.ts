// §7.4 — the general-purpose stream filters, undone for ANY stream.
//
// These four say nothing about what the bytes mean: they are transport, and a
// stream still wearing one has not been read. The image path had its own copy
// and `streamData` passed them through on the reasoning that "image-decode
// undoes them" — true for an image, and false for everything else a PDF puts in
// a stream. asciihexdecode.pdf writes its whole page as
// `42540A2F46312033302054660A…` — that is `BT /F1 30 Tf …` — and the page came
// back blank because nothing ever turned the hex into bytes.
//
// The image CODECS (`/DCTDecode`, `/JPXDecode`, `/JBIG2Decode`,
// `/CCITTFaxDecode`) are a different thing: they ARE the image, and undoing
// them is decoding a picture, which belongs to `./image-decode`.

/** §7.4.5 — a byte, then that many literal bytes, or a run of one repeated. */
export function runLengthDecode(data: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  let i = 0;
  while (i < data.length) {
    const len = data[i++]!;
    if (len === 128) break; // EOD
    if (len < 128) {
      for (let j = 0; j <= len && i < data.length; j++) out.push(data[i++]!);
    } else {
      const b = data[i++] ?? 0;
      for (let j = 0; j < 257 - len; j++) out.push(b);
    }
  }
  return Uint8Array.from(out);
}

/** §7.4.3 — four bytes as five base-85 digits, `z` for four zeroes, `~>` ends. */
export function ascii85Decode(data: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  let tuple = 0;
  let count = 0;
  for (const c of data) {
    if (c === 0x7e) break; // ~> terminator
    if (c <= 0x20) continue; // whitespace
    if (c === 0x7a && count === 0) {
      out.push(0, 0, 0, 0);
      continue;
    }
    if (c < 0x21 || c > 0x75) continue;
    tuple = tuple * 85 + (c - 0x21);
    if (++count === 5) {
      out.push((tuple >>> 24) & 0xff, (tuple >>> 16) & 0xff, (tuple >>> 8) & 0xff, tuple & 0xff);
      tuple = 0;
      count = 0;
    }
  }
  if (count > 0) {
    for (let i = count; i < 5; i++) tuple = tuple * 85 + 84;
    for (let i = 0; i < count - 1; i++) out.push((tuple >>> (24 - i * 8)) & 0xff);
  }
  return Uint8Array.from(out);
}

/** §7.4.2 — one byte as two hex digits, `>` ends, an odd last digit is halved. */
export function asciiHexDecode(data: Uint8Array): Uint8Array {
  const out: Array<number> = [];
  let hi = -1;
  for (const c of data) {
    if (c === 0x3e) break; // '>'
    const v = hexVal(c);
    if (v < 0) continue;
    if (hi < 0) hi = v;
    else {
      out.push((hi << 4) | v);
      hi = -1;
    }
  }
  if (hi >= 0) out.push(hi << 4);
  return Uint8Array.from(out);
}

function hexVal(b: number): number {
  if (b >= 0x30 && b <= 0x39) return b - 0x30;
  if (b >= 0x41 && b <= 0x46) return b - 0x41 + 10;
  if (b >= 0x61 && b <= 0x66) return b - 0x61 + 10;
  return -1;
}
