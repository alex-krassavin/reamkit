// E-PDF EP2 — /ToUnicode CMap parser (ISO 32000-1 §9.10.3). A ToUnicode CMap is
// a small PostScript program mapping character codes to Unicode via `bfchar`
// (single codes) and `bfrange` (contiguous ranges) blocks; `codespacerange`
// declares whether codes are one or two bytes. We need only those three blocks.

import { Lexer } from './lexer';

/** A parsed `/ToUnicode` CMap: the code→Unicode mapping and the code width. */
export interface ToUnicode {
  /** Character code → Unicode string (a single code may map to a ligature). */
  readonly map: ReadonlyMap<number, string>;
  /** Whether character codes are one or two bytes wide (from `codespacerange`). */
  readonly codeBytes: 1 | 2;
}

const MAX_RANGE = 65_536; // DoS guard on a pathological bfrange

/**
 * Parse a `/ToUnicode` CMap (ISO 32000-1 §9.10.3) into a code→Unicode map. Reads
 * the `codespacerange` (code width), `bfchar` (single codes) and `bfrange`
 * (contiguous ranges, scalar or array destination) blocks; everything else is
 * ignored.
 *
 * @param bytes The decoded CMap stream.
 * @returns The mapping plus the declared code width (1 or 2 bytes).
 */
export function parseToUnicodeCMap(bytes: Uint8Array): ToUnicode {
  const lexer = new Lexer(bytes);
  const map = new Map<number, string>();
  let codeBytes: 1 | 2 = 1;

  for (;;) {
    const tok = lexer.nextToken();
    if (tok.kind === 'eof') break;
    if (tok.kind !== 'keyword') continue;

    if (tok.value === 'begincodespacerange') {
      for (;;) {
        const t = lexer.nextToken();
        if (t.kind === 'eof' || (t.kind === 'keyword' && t.value === 'endcodespacerange')) break;
        if (t.kind === 'hexstr') {
          if (t.bytes.length >= 2) codeBytes = 2;
          lexer.nextToken(); // skip the range's upper bound
        }
      }
    } else if (tok.value === 'beginbfchar') {
      for (;;) {
        const src = lexer.nextToken();
        if (src.kind === 'eof' || (src.kind === 'keyword' && src.value === 'endbfchar')) break;
        if (src.kind !== 'hexstr') continue;
        const dst = lexer.nextToken();
        if (dst.kind !== 'hexstr') continue;
        map.set(bytesToInt(src.bytes), utf16be(dst.bytes));
      }
    } else if (tok.value === 'beginbfrange') {
      for (;;) {
        const lo = lexer.nextToken();
        if (lo.kind === 'eof' || (lo.kind === 'keyword' && lo.value === 'endbfrange')) break;
        if (lo.kind !== 'hexstr') continue;
        const hi = lexer.nextToken();
        if (hi.kind !== 'hexstr') continue;
        const loN = bytesToInt(lo.bytes);
        const hiN = bytesToInt(hi.bytes);
        const dst = lexer.nextToken();
        if (dst.kind === 'hexstr') {
          const base = utf16be(dst.bytes);
          for (let i = 0; loN + i <= hiN && i < MAX_RANGE; i++)
            map.set(loN + i, incString(base, i));
        } else if (dst.kind === 'arrayOpen') {
          let i = 0;
          for (;;) {
            const el = lexer.nextToken();
            if (el.kind === 'arrayClose' || el.kind === 'eof') break;
            if (el.kind === 'hexstr') map.set(loN + i++, utf16be(el.bytes));
          }
        }
      }
    }
  }
  return { map, codeBytes };
}

// Big-endian bytes → integer (the character code).
function bytesToInt(bytes: Uint8Array): number {
  let n = 0;
  for (const b of bytes) n = n * 256 + b;
  return n;
}

// Big-endian UTF-16 bytes → string (a surrogate pair becomes its astral char).
function utf16be(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i + 1 < bytes.length; i += 2)
    s += String.fromCharCode((bytes[i]! << 8) | bytes[i + 1]!);
  if (bytes.length % 2 === 1) s += String.fromCharCode(bytes[bytes.length - 1]!);
  return s;
}

// §9.10.3 — a bfrange counts its destination up along with the code.
//
// What counts is the CHARACTER, and UTF-16 writes an astral one as two units:
// `<043B> <0441> <D835DC4E>` maps seven codes to the italic letters a..g, and
// measuring the destination in units left all seven at "𝑎" — bug1529502.pdf
// says "HKDF(s, version, e, 32)" and we read "version, a".
function incString(base: string, i: number): string {
  if (i === 0) return base;
  const chars = [...base];
  if (chars.length === 0) return base;
  // A destination of several characters is a sequence — a ligature, or a letter
  // with its mark — and it is the last of them the range counts up.
  const last = chars[chars.length - 1]!.codePointAt(0)! + i;
  if (last > 0x10ffff || (last >= 0xd800 && last <= 0xdfff)) return base;
  return chars.slice(0, -1).join('') + String.fromCodePoint(last);
}
