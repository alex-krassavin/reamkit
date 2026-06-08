// ISO 32000-1:2008 §D.2 — WinAnsiEncoding.
// Built-in Type 1 fonts default to StandardEncoding; we explicitly switch to
// WinAnsiEncoding (a.k.a. CP-1252) in the font dictionary so common Latin-1
// characters render without per-character mapping.

const WIN_ANSI_REMAP: Record<number, number> = {
  0x20ac: 0x80,
  0x201a: 0x82,
  0x0192: 0x83,
  0x201e: 0x84,
  0x2026: 0x85,
  0x2020: 0x86,
  0x2021: 0x87,
  0x02c6: 0x88,
  0x2030: 0x89,
  0x0160: 0x8a,
  0x2039: 0x8b,
  0x0152: 0x8c,
  0x017d: 0x8e,
  0x2018: 0x91,
  0x2019: 0x92,
  0x201c: 0x93,
  0x201d: 0x94,
  0x2022: 0x95,
  0x2013: 0x96,
  0x2014: 0x97,
  0x02dc: 0x98,
  0x2122: 0x99,
  0x0161: 0x9a,
  0x203a: 0x9b,
  0x0153: 0x9c,
  0x017e: 0x9e,
  0x0178: 0x9f,
};

const SUBSTITUTE = 0x3f;

export function encodeWinAnsi(text: string): Uint8Array {
  const codePoints = [...text];
  const out = new Uint8Array(codePoints.length);
  for (let i = 0; i < codePoints.length; i++) {
    const cp = codePoints[i]!.codePointAt(0)!;
    if (cp <= 0x7f || (cp >= 0xa0 && cp <= 0xff)) {
      out[i] = cp;
    } else if (cp in WIN_ANSI_REMAP) {
      out[i] = WIN_ANSI_REMAP[cp]!;
    } else {
      out[i] = SUBSTITUTE;
    }
  }
  return out;
}

// ISO 32000-1:2008 §7.3.4.2 — Literal strings: parentheses must be balanced
// or escaped; backslash escapes itself. Output bytes ready to be wrapped in
// "(" ... ")" within a content stream.
export function encodeLiteralBytes(text: string): Uint8Array {
  const bytes = encodeWinAnsi(text);
  let extra = 0;
  for (const b of bytes) {
    if (b === 0x28 || b === 0x29 || b === 0x5c) extra++;
  }
  if (extra === 0) return bytes;
  const out = new Uint8Array(bytes.length + extra);
  let j = 0;
  for (const b of bytes) {
    if (b === 0x28 || b === 0x29 || b === 0x5c) out[j++] = 0x5c;
    out[j++] = b;
  }
  return out;
}
