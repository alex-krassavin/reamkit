// E-PDF EP1 — COS lexer (ISO 32000-1 §7.2/§7.3). Turns a PDF byte buffer into a
// stream of lexical tokens: numbers, names, strings (literal + hex), the array
// and dictionary delimiters, and bare keywords (obj / R / stream / true / …).
// The parser (parser.ts) drives this with look-ahead to recover the object
// grammar (an `N G R` reference, an `N G obj … endobj` definition, a stream).
//
// We read PDFs into the SAME object model the writer emits (src/pdf/objects.ts),
// so a parsed document is the inverse of a written one — the natural shape for
// round-tripping Ream's own tagged output back into a FlowDoc.

/**
 * One COS lexical token (ISO 32000-1 §7.2/§7.3): a number, name, string (literal
 * or hex), an array/dictionary delimiter, a bare keyword (`obj` / `R` / `stream`
 * / `true` / …), or end-of-input.
 */
export type Token =
  | { readonly kind: 'num'; readonly value: number }
  | { readonly kind: 'name'; readonly value: string }
  | { readonly kind: 'str'; readonly value: string } // literal (…) → latin1
  | { readonly kind: 'hexstr'; readonly bytes: Uint8Array } // <…>
  | { readonly kind: 'arrayOpen' } // [
  | { readonly kind: 'arrayClose' } // ]
  | { readonly kind: 'dictOpen' } // <<
  | { readonly kind: 'dictClose' } // >>
  | { readonly kind: 'keyword'; readonly value: string } // obj endobj stream R true …
  | { readonly kind: 'eof' };

// §7.2.3 — the six PDF whitespace bytes (NUL, TAB, LF, FF, CR, SP).
function isWhitespace(b: number): boolean {
  return b === 0x00 || b === 0x09 || b === 0x0a || b === 0x0c || b === 0x0d || b === 0x20;
}

// §7.2.2 — the delimiter bytes ( ) < > [ ] { } / %.
function isDelimiter(b: number): boolean {
  return (
    b === 0x28 || // (
    b === 0x29 || // )
    b === 0x3c || // <
    b === 0x3e || // >
    b === 0x5b || // [
    b === 0x5d || // ]
    b === 0x7b || // {
    b === 0x7d || // }
    b === 0x2f || // /
    b === 0x25 // %
  );
}

// A "regular" byte: anything that is neither whitespace nor a delimiter. Names
// and keywords are runs of these.
function isRegular(b: number): boolean {
  return !isWhitespace(b) && !isDelimiter(b);
}

function hexVal(b: number): number {
  if (b >= 0x30 && b <= 0x39) return b - 0x30; // 0-9
  if (b >= 0x41 && b <= 0x46) return b - 0x41 + 10; // A-F
  if (b >= 0x61 && b <= 0x66) return b - 0x61 + 10; // a-f
  return -1;
}

/**
 * COS lexer (ISO 32000-1 §7.2/§7.3): scans a PDF byte buffer into a stream of
 * {@link Token}s. The parser drives it with look-ahead to recover the object
 * grammar. `pos` is the public cursor; callers (and the parser's rewind logic)
 * read and assign it directly.
 */
export class Lexer {
  pos: number;
  /**
   * @param buf The PDF byte buffer to tokenize.
   * @param pos The starting byte offset (defaults to the start of the buffer).
   */
  constructor(
    private readonly buf: Uint8Array,
    pos = 0,
  ) {
    this.pos = pos;
  }

  /** The length of the underlying byte buffer. */
  get length(): number {
    return this.buf.length;
  }

  /** The byte at index `i`, or −1 when out of range. */
  byteAt(i: number): number {
    return i >= 0 && i < this.buf.length ? this.buf[i]! : -1;
  }

  /** §7.2.3 — skip whitespace and `%`-to-end-of-line comments. */
  skipWhitespace(): void {
    const buf = this.buf;
    while (this.pos < buf.length) {
      const b = buf[this.pos]!;
      if (isWhitespace(b)) {
        this.pos++;
      } else if (b === 0x25) {
        // comment: to end of line (CR, LF or CRLF)
        this.pos++;
        while (this.pos < buf.length && buf[this.pos] !== 0x0a && buf[this.pos] !== 0x0d) {
          this.pos++;
        }
      } else {
        break;
      }
    }
  }

  /** Read and consume the next {@link Token} from the current position. */
  nextToken(): Token {
    this.skipWhitespace();
    const buf = this.buf;
    if (this.pos >= buf.length) return { kind: 'eof' };
    const b = buf[this.pos]!;

    switch (b) {
      case 0x5b: // [
        this.pos++;
        return { kind: 'arrayOpen' };
      case 0x5d: // ]
        this.pos++;
        return { kind: 'arrayClose' };
      case 0x3c: // <  — either "<<" (dict) or a hex string
        if (buf[this.pos + 1] === 0x3c) {
          this.pos += 2;
          return { kind: 'dictOpen' };
        }
        return this.readHexString();
      case 0x3e: // >  — must be ">>"
        if (buf[this.pos + 1] === 0x3e) {
          this.pos += 2;
          return { kind: 'dictClose' };
        }
        this.pos++; // lone '>' is malformed; skip it so we make progress
        return this.nextToken();
      case 0x28: // (
        return this.readLiteralString();
      case 0x2f: // /
        return this.readName();
      case 0x7b: // {
      case 0x7d: // }  — PostScript function delimiters; surface as keywords
        this.pos++;
        return { kind: 'keyword', value: String.fromCharCode(b) };
    }

    if (b === 0x2b || b === 0x2d || b === 0x2e || (b >= 0x30 && b <= 0x39)) {
      return this.readNumber();
    }
    if (isRegular(b)) return this.readKeyword();
    // Unknown byte — skip to stay productive.
    this.pos++;
    return this.nextToken();
  }

  /** Read a numeric token (optional sign, digits and a decimal point). */
  private readNumber(): Token {
    const buf = this.buf;
    const start = this.pos;
    if (buf[this.pos] === 0x2b || buf[this.pos] === 0x2d) this.pos++;
    while (this.pos < buf.length) {
      const b = buf[this.pos]!;
      if ((b >= 0x30 && b <= 0x39) || b === 0x2e) this.pos++;
      else break;
    }
    const text = latin1(buf.subarray(start, this.pos));
    const value = Number(text);
    return { kind: 'num', value: Number.isFinite(value) ? value : 0 };
  }

  /** Read a `/Name` token, decoding `#XX` hex escapes (§7.3.5). */
  private readName(): Token {
    const buf = this.buf;
    this.pos++; // consume '/'
    const out: Array<number> = [];
    while (this.pos < buf.length) {
      const b = buf[this.pos]!;
      if (!isRegular(b)) break;
      if (b === 0x23 && this.pos + 2 < buf.length) {
        // #XX hex escape inside a name (§7.3.5)
        const hi = hexVal(buf[this.pos + 1]!);
        const lo = hexVal(buf[this.pos + 2]!);
        if (hi >= 0 && lo >= 0) {
          out.push(hi * 16 + lo);
          this.pos += 3;
          continue;
        }
      }
      out.push(b);
      this.pos++;
    }
    return { kind: 'name', value: latin1(Uint8Array.from(out)) };
  }

  /** Read a bare keyword token: a run of regular bytes (`obj`, `R`, `true`, …). */
  private readKeyword(): Token {
    const buf = this.buf;
    const start = this.pos;
    while (this.pos < buf.length && isRegular(buf[this.pos]!)) this.pos++;
    return { kind: 'keyword', value: latin1(buf.subarray(start, this.pos)) };
  }

  /** Read a `<…>` hex string (§7.3.4.3); an odd trailing digit takes a low nibble of 0. */
  private readHexString(): Token {
    const buf = this.buf;
    this.pos++; // consume '<'
    const out: Array<number> = [];
    let hi = -1;
    while (this.pos < buf.length) {
      const b = buf[this.pos]!;
      this.pos++;
      if (b === 0x3e) break; // '>'
      const v = hexVal(b);
      if (v < 0) continue; // whitespace between hex digits is allowed
      if (hi < 0) hi = v;
      else {
        out.push(hi * 16 + v);
        hi = -1;
      }
    }
    if (hi >= 0) out.push(hi * 16); // odd trailing digit → low nibble 0 (§7.3.4.3)
    return { kind: 'hexstr', bytes: Uint8Array.from(out) };
  }

  /**
   * Read a `(…)` literal string (§7.3.4.2): handles nested parentheses, backslash
   * escapes and octal codes. The bytes are decoded latin1.
   */
  private readLiteralString(): Token {
    const buf = this.buf;
    this.pos++; // consume '('
    const out: Array<number> = [];
    let depth = 1;
    while (this.pos < buf.length) {
      const b = buf[this.pos]!;
      this.pos++;
      if (b === 0x5c) {
        // backslash escape (§7.3.4.2)
        if (this.pos >= buf.length) break;
        const e = buf[this.pos]!;
        this.pos++;
        switch (e) {
          case 0x6e: // \n
            out.push(0x0a);
            break;
          case 0x72: // \r
            out.push(0x0d);
            break;
          case 0x74: // \t
            out.push(0x09);
            break;
          case 0x62: // \b
            out.push(0x08);
            break;
          case 0x66: // \f
            out.push(0x0c);
            break;
          case 0x0a: // line continuation: backslash-LF → nothing
            break;
          case 0x0d: // backslash-CR (or CRLF) → nothing
            if (buf[this.pos] === 0x0a) this.pos++;
            break;
          default:
            if (e >= 0x30 && e <= 0x37) {
              // up to three octal digits
              let oct = e - 0x30;
              for (let k = 0; k < 2 && this.pos < buf.length; k++) {
                const d = buf[this.pos]!;
                if (d < 0x30 || d > 0x37) break;
                oct = oct * 8 + (d - 0x30);
                this.pos++;
              }
              out.push(oct & 0xff);
            } else {
              out.push(e); // \( \) \\ and any other → the char itself
            }
        }
        continue;
      }
      if (b === 0x28) {
        depth++;
        out.push(b);
        continue;
      }
      if (b === 0x29) {
        depth--;
        if (depth === 0) break;
        out.push(b);
        continue;
      }
      out.push(b);
    }
    return { kind: 'str', value: latin1(Uint8Array.from(out)) };
  }

  /**
   * First index of an ASCII `needle` at or after `from` (−1 if none). Used to find
   * `endstream` when a stream's `/Length` is missing or an unresolved reference.
   */
  indexOfAscii(needle: string, from: number): number {
    const buf = this.buf;
    const n = needle.length;
    outer: for (let i = from; i <= buf.length - n; i++) {
      for (let j = 0; j < n; j++) {
        if (buf[i + j] !== needle.charCodeAt(j)) continue outer;
      }
      return i;
    }
    return -1;
  }

  /**
   * §7.3.8.1 — read a stream's raw bytes. `pos` must sit right after the `stream`
   * keyword. The keyword is followed by CRLF (or a lone LF); the data then runs
   * for `length` bytes, or — when the length is unknown — up to `endstream`.
   * Leaves `pos` at the `endstream` keyword.
   */
  readStreamBody(length: number | undefined): Uint8Array {
    const buf = this.buf;
    if (buf[this.pos] === 0x0d && buf[this.pos + 1] === 0x0a) this.pos += 2;
    else if (buf[this.pos] === 0x0a || buf[this.pos] === 0x0d) this.pos += 1;
    const start = this.pos;
    if (length !== undefined && length >= 0 && start + length <= buf.length) {
      this.pos = start + length;
      return buf.subarray(start, start + length);
    }
    const es = this.indexOfAscii('endstream', start);
    const end = es < 0 ? buf.length : es;
    let dataEnd = end;
    if (dataEnd > start && buf[dataEnd - 1] === 0x0a) {
      dataEnd--;
      if (dataEnd > start && buf[dataEnd - 1] === 0x0d) dataEnd--;
    } else if (dataEnd > start && buf[dataEnd - 1] === 0x0d) {
      dataEnd--;
    }
    this.pos = end;
    return buf.subarray(start, dataEnd);
  }
}

/**
 * Bytes → a Latin-1 (ISO-8859-1) string: each byte becomes the code point of the
 * same value, so the string round-trips back to the exact bytes. PDF text in
 * strings is decoded to Unicode later (via the font's `/ToUnicode`); at the COS
 * layer a string is just bytes.
 */
export function latin1(bytes: Uint8Array): string {
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  return s;
}
