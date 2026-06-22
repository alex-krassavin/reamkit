// ISO 32000-1:2008 §7.3 — Object types.

/** The PDF `null` object (§7.3.9), modelled as a unique symbol. */
export const PDF_NULL = Symbol('PDF_NULL');
/** The type of the PDF `null` object, {@link PDF_NULL}. */
export type PdfNull = typeof PDF_NULL;

/** A PDF name object (§7.3.5), e.g. `/Type` — the leading slash is added on serialization. */
export class PdfName {
  /** @param value The name's text, without the leading `/`. */
  constructor(public readonly value: string) {}
}

/** An indirect reference (§7.3.10), written `id gen R`. */
export class PdfRef {
  /**
   * @param id         The referenced object number.
   * @param generation The generation number (0 for freshly written objects).
   */
  constructor(
    public readonly id: number,
    public readonly generation: number = 0,
  ) {}
}

/** A PDF stream object (§7.3.8): a dictionary plus its raw byte payload. */
export class PdfStream {
  /**
   * @param dict The stream dictionary (`/Length` is filled in on serialization).
   * @param data The raw stream bytes.
   */
  constructor(
    public readonly dict: PdfDict,
    public readonly data: Uint8Array,
  ) {}
}

/**
 * A PDF hex string (§7.3.4.3), written as `<HEXBYTES>` and decoding to raw
 * bytes. Used for `/Info` entries containing non-ASCII text: the bytes are
 * UTF-16BE with a leading BOM, the PDF-standard form for Unicode strings.
 */
export class PdfHexString {
  /** @param bytes The raw bytes to emit as hex. */
  constructor(public readonly bytes: Uint8Array) {}
}

/**
 * A verbatim token emitted exactly as given — no escaping, no quoting. Used for
 * the signature `/ByteRange` placeholder, a fixed-width array overwritten in
 * place after the byte offsets are known (ISO 32000 §12.8.1).
 */
export class PdfRawToken {
  /** @param token The literal text to emit unchanged. */
  constructor(public readonly token: string) {}
}

/** A PDF dictionary (§7.3.7): an ordered map from name keys to {@link PdfValue}s. */
export type PdfDict = Map<string, PdfValue>;
/** A PDF array (§7.3.6) of {@link PdfValue}s. */
export type PdfArray = Array<PdfValue>;

/** Any PDF object value (§7.3): the union of all the primitive and composite object types. */
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

/** Construct a {@link PdfName} from its text (without the leading `/`). */
export const name = (v: string): PdfName => new PdfName(v);
/** Construct a {@link PdfRef} to object `id` at generation `gen` (default 0). */
export const ref = (id: number, gen = 0): PdfRef => new PdfRef(id, gen);
/** Construct a {@link PdfDict} from a plain object of name → value entries. */
export const dict = (entries: Record<string, PdfValue>): PdfDict =>
  new Map(Object.entries(entries));
/** Construct a {@link PdfStream} from a dictionary literal plus its byte payload. */
export const stream = (entries: Record<string, PdfValue>, data: Uint8Array): PdfStream =>
  new PdfStream(dict(entries), data);

/**
 * Build a PDF Unicode string (UTF-16BE with BOM) suitable for `/Info` or
 * document-level text metadata. Accepts arbitrary Unicode input.
 *
 * @param s The text to encode.
 * @returns A {@link PdfHexString} holding the BOM-prefixed UTF-16BE bytes.
 */
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
