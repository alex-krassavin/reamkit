// ISO 32000-1:2008 §7.3 — Object types.

export const PDF_NULL = Symbol('PDF_NULL');
export type PdfNull = typeof PDF_NULL;

export class PdfName {
  constructor(public readonly value: string) {}
}

export class PdfRef {
  constructor(
    public readonly id: number,
    public readonly generation: number = 0,
  ) {}
}

export class PdfStream {
  constructor(
    public readonly dict: PdfDict,
    public readonly data: Uint8Array,
  ) {}
}

// PDF hex strings (§7.3.4.3) are written as `<HEXBYTES>` and decode to raw
// bytes. We use them for /Info entries containing non-ASCII text: the bytes
// are UTF-16BE with a leading BOM, the PDF-standard form for Unicode strings.
export class PdfHexString {
  constructor(public readonly bytes: Uint8Array) {}
}

// A verbatim token emitted exactly as given — no escaping, no quoting. Used for
// the signature /ByteRange placeholder, a fixed-width array overwritten in place
// after the byte offsets are known (ISO 32000 §12.8.1).
export class PdfRawToken {
  constructor(public readonly token: string) {}
}

export type PdfDict = Map<string, PdfValue>;
export type PdfArray = Array<PdfValue>;

export type PdfValue =
  | PdfNull
  | boolean
  | number
  | string
  | PdfName
  | PdfHexString
  | PdfRawToken
  | PdfArray
  | PdfDict
  | PdfRef
  | PdfStream;

export const name = (v: string): PdfName => new PdfName(v);
export const ref = (id: number, gen = 0): PdfRef => new PdfRef(id, gen);
export const dict = (entries: Record<string, PdfValue>): PdfDict =>
  new Map(Object.entries(entries));
export const stream = (entries: Record<string, PdfValue>, data: Uint8Array): PdfStream =>
  new PdfStream(dict(entries), data);

// Build a PDF Unicode string (UTF-16BE with BOM) suitable for /Info or
// document-level text metadata. Accepts arbitrary unicode input.
export const unicodeString = (s: string): PdfHexString => {
  // 2 bytes BOM + 2 bytes per UTF-16 code unit.
  const codepoints: Array<number> = [];
  for (const ch of s) codepoints.push(ch.codePointAt(0)!);
  // Re-expand to UTF-16 code units (handle surrogate pairs).
  const units: Array<number> = [];
  for (const cp of codepoints) {
    if (cp < 0x10000) {
      units.push(cp);
    } else {
      const adj = cp - 0x10000;
      units.push(0xd800 + (adj >> 10));
      units.push(0xdc00 + (adj & 0x3ff));
    }
  }
  const out = new Uint8Array(2 + units.length * 2);
  out[0] = 0xfe;
  out[1] = 0xff;
  for (let i = 0; i < units.length; i++) {
    out[2 + i * 2] = (units[i]! >> 8) & 0xff;
    out[2 + i * 2 + 1] = units[i]! & 0xff;
  }
  return new PdfHexString(out);
};
