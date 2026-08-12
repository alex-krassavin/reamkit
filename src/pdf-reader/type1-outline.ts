// §9.6.6 / Adobe "Type 1 Font Format" — the OUTLINE of one glyph in an
// embedded Type 1 program (`/FontFile`).
//
// The third outline format, after `./glyf-outline` and `./cff-outline`, and the
// oldest. A Type 1 program is PostScript source with its interesting half
// ENCRYPTED: `eexec` hides the private dictionary, and each charstring inside
// is encrypted again on its own. Under both layers is a stack machine much like
// CFF's, with a smaller and older instruction set.
//
// Two things come out of it. The charstrings, which draw; and the program's own
// `/Encoding` array, which says what character each code stands for when the
// PDF's font dictionary does not — a `/FontFile` face with a built-in encoding
// and no `/Differences` says everything about its text in the program and
// nothing in the file.
//
// The outline comes back in a ONE-UNIT em, like the other two readers'.

import type { PathSeg } from './content';

/** What an embedded Type 1 program says about its glyphs. */
export interface Type1Font {
  /** The contours of one glyph, by NAME — a Type 1 font indexes by nothing else. */
  readonly path: (name: string) => Array<PathSeg> | undefined;
  /** The program's own `/Encoding`: code → glyph name, where it states one. */
  readonly encoding?: ReadonlyMap<number, string>;
}

/**
 * Read an embedded Type 1 program.
 *
 * @param program The raw `/FontFile` bytes — the cleartext header, the
 *                `eexec`-encrypted body, and the trailing zeros.
 * @returns A reader, or `undefined` where the program cannot be read.
 */
export function type1Font(program: Uint8Array): Type1Font | undefined {
  let parsed: Parsed;
  try {
    const read = parseType1(program);
    if (!read) return undefined;
    parsed = read;
  } catch {
    return undefined;
  }
  const cache = new Map<string, Array<PathSeg> | undefined>();
  return {
    path: (name: string): Array<PathSeg> | undefined => {
      const had = cache.get(name);
      if (had !== undefined || cache.has(name)) return had;
      let out: Array<PathSeg> | undefined;
      try {
        const charstring = parsed.charstrings.get(name);
        out = charstring ? runGlyph(parsed, charstring, 0) : undefined;
      } catch {
        out = undefined;
      }
      cache.set(name, out);
      return out;
    },
    ...(parsed.encoding ? { encoding: parsed.encoding } : {}),
  };
}

interface Parsed {
  readonly charstrings: Map<string, Uint8Array>;
  readonly subrs: Array<Uint8Array>;
  readonly encoding?: Map<number, string>;
  /** `/FontMatrix`'s scale — 1/1000 for all but a handful of faces. */
  readonly scale: number;
}

function parseType1(program: Uint8Array): Parsed | undefined {
  const body = pfbSegments(program) ?? program;
  const at = indexOf(body, 'eexec');
  if (at < 0) return undefined;
  const clear = new TextDecoder('latin1').decode(body.subarray(0, at));
  // §7 — whitespace follows `eexec`, then the ciphertext, which may be written
  // as hexadecimal instead of binary.
  let from = at + 5;
  while (from < body.length && isSpace(body[from]!)) from++;
  const cipher = looksHex(body, from) ? unhex(body, from) : body.subarray(from);
  const plain = eexec(cipher, EEXEC_R, 4);
  const text = new TextDecoder('latin1').decode(plain);
  const lenIV = readInt(text, '/lenIV') ?? 4;
  const charstrings = readCharstrings(plain, text, lenIV);
  if (charstrings.size === 0) return undefined;
  const encoding = readEncoding(clear);
  return {
    charstrings,
    subrs: readSubrs(plain, text, lenIV),
    ...(encoding ? { encoding } : {}),
    scale: readMatrixScale(clear) ?? 0.001,
  };
}

/** The `eexec` key (§7), and the one every charstring is encrypted with (§6.2). */
const EEXEC_R = 55665;
const CHARSTRING_R = 4330;

/** §7 — the decryption both layers use, less the leading random bytes. */
function eexec(data: Uint8Array, key: number, skip: number): Uint8Array {
  let r = key;
  const out = new Uint8Array(Math.max(0, data.length - skip));
  for (let i = 0; i < data.length; i++) {
    const c = data[i]!;
    const p = c ^ (r >> 8);
    r = ((c + r) * 52845 + 22719) & 0xffff;
    if (i >= skip) out[i - skip] = p & 0xff;
  }
  return out;
}

/**
 * A PFB wrapper: 0x80 0x01/0x02 then a four-byte little-endian length. The
 * segments are the file with its lengths interleaved, so they are joined back.
 */
function pfbSegments(program: Uint8Array): Uint8Array | undefined {
  if (program[0] !== 0x80) return undefined;
  const parts: Array<Uint8Array> = [];
  let at = 0;
  while (at + 6 <= program.length && program[at] === 0x80 && program[at + 1] !== 3) {
    const len =
      program[at + 2]! |
      (program[at + 3]! << 8) |
      (program[at + 4]! << 16) |
      (program[at + 5]! << 24);
    const from = at + 6;
    if (len < 0 || from + len > program.length) break;
    parts.push(program.subarray(from, from + len));
    at = from + len;
  }
  if (parts.length === 0) return undefined;
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/** §10 — `/CharStrings n dict dup begin /name len RD <bytes> ND` for each glyph. */
function readCharstrings(plain: Uint8Array, text: string, lenIV: number): Map<string, Uint8Array> {
  const out = new Map<string, Uint8Array>();
  const start = text.indexOf('/CharStrings');
  if (start < 0) return out;
  const entry = /\/([^\s/{}()[\]<>]+)\s+(\d+)\s+(-\||RD|ND)[ ]/gu;
  entry.lastIndex = start;
  for (let m = entry.exec(text); m; m = entry.exec(text)) {
    const length = Number(m[2]);
    const from = m.index + m[0].length;
    if (!Number.isFinite(length) || length <= 0 || from + length > plain.length) continue;
    if (m[3] === 'ND') continue; // `ND` ends an entry; it never opens one
    if (!out.has(m[1]!))
      out.set(m[1]!, eexec(plain.subarray(from, from + length), CHARSTRING_R, lenIV));
    entry.lastIndex = from + length;
    if (out.size >= MAX_GLYPHS) break;
  }
  return out;
}

/** How many glyphs one program may hold — a guard, not a format limit. */
const MAX_GLYPHS = 5000;

/** §8 — `/Subrs n array dup i len RD <bytes> NP`, called by index. */
function readSubrs(plain: Uint8Array, text: string, lenIV: number): Array<Uint8Array> {
  const out: Array<Uint8Array> = [];
  const start = text.indexOf('/Subrs');
  if (start < 0) return out;
  const entry = /dup\s+(\d+)\s+(\d+)\s+(-\||RD)[ ]/gu;
  entry.lastIndex = start;
  for (let m = entry.exec(text); m; m = entry.exec(text)) {
    const index = Number(m[1]);
    const length = Number(m[2]);
    const from = m.index + m[0].length;
    if (!Number.isFinite(length) || from + length > plain.length) break;
    if (index >= 0 && index < MAX_GLYPHS) {
      out[index] = eexec(plain.subarray(from, from + length), CHARSTRING_R, lenIV);
    }
    entry.lastIndex = from + length;
    if (text.startsWith('/CharStrings', entry.lastIndex)) break;
  }
  return out;
}

/**
 * §5 — the program's own `/Encoding`: either `StandardEncoding` by name, or an
 * array built up with `dup <code> /<name> put`.
 */
function readEncoding(clear: string): Map<number, string> | undefined {
  const at = clear.indexOf('/Encoding');
  if (at < 0) return undefined;
  const out = new Map<number, string>();
  const put = /dup\s+(\d+)\s*\/([^\s/{}()[\]<>]+)\s+put/gu;
  put.lastIndex = at;
  const end = clear.indexOf('readonly def', at);
  for (let m = put.exec(clear); m; m = put.exec(clear)) {
    if (end >= 0 && m.index > end) break;
    const code = Number(m[1]);
    if (Number.isFinite(code) && code >= 0 && code < 256) out.set(code, m[2]!);
  }
  return out.size > 0 ? out : undefined;
}

/** §5 — `/FontMatrix [a b c d e f]`, whose `a` is the em this program states. */
function readMatrixScale(clear: string): number | undefined {
  const m = /\/FontMatrix\s*\[\s*([-\d.eE]+)/u.exec(clear);
  const a = m ? Number(m[1]) : Number.NaN;
  return Number.isFinite(a) && a !== 0 ? a : undefined;
}

function readInt(text: string, key: string): number | undefined {
  const m = new RegExp(`${key}\\s+(-?\\d+)`, 'u').exec(text);
  const v = m ? Number(m[1]) : Number.NaN;
  return Number.isFinite(v) ? v : undefined;
}

function indexOf(data: Uint8Array, needle: string): number {
  const bytes = [...needle].map((c) => c.charCodeAt(0));
  outer: for (let i = 0; i + bytes.length <= data.length; i++) {
    for (let k = 0; k < bytes.length; k++) if (data[i + k] !== bytes[k]) continue outer;
    return i;
  }
  return -1;
}

function isSpace(b: number): boolean {
  return b === 0x20 || b === 0x0a || b === 0x0d || b === 0x09;
}

/** The ciphertext may be hexadecimal — four hex digits at the head give it away. */
function looksHex(data: Uint8Array, from: number): boolean {
  const hex = (b: number): boolean =>
    (b >= 0x30 && b <= 0x39) || (b >= 0x41 && b <= 0x46) || (b >= 0x61 && b <= 0x66);
  return [0, 1, 2, 3].every((k) => hex(data[from + k] ?? 0));
}

function unhex(data: Uint8Array, from: number): Uint8Array {
  const out: Array<number> = [];
  let half = -1;
  for (let i = from; i < data.length; i++) {
    const b = data[i]!;
    let v = -1;
    if (b >= 0x30 && b <= 0x39) v = b - 0x30;
    else if (b >= 0x41 && b <= 0x46) v = b - 0x37;
    else if (b >= 0x61 && b <= 0x66) v = b - 0x57;
    else if (isSpace(b)) continue;
    else break;
    if (half < 0) half = v;
    else {
      out.push((half << 4) | v);
      half = -1;
    }
  }
  return Uint8Array.from(out);
}

/** How deep a charstring may call, and how many operators it may run. */
const MAX_CALL_DEPTH = 10;
const MAX_STEPS = 65536;

// §6 — run one Type 1 charstring. The operators are CFF's ancestors: the same
// moves, lines and curves, with the width stated by `hsbw` rather than counted
// off the stack, and flex and hint replacement done through `callothersubr`.
function runGlyph(font: Parsed, program: Uint8Array, depth: number): Array<PathSeg> | undefined {
  const out: Array<PathSeg> = [];
  const stack: Array<number> = [];
  const ps: Array<number> = []; // the PostScript stack `pop` reads back
  const flex: Array<number> = []; // §8.3 — the points a flex collects
  let inFlex = false;
  let x = 0;
  let y = 0;
  let open = false;
  let steps = 0;

  const moveTo = (): void => {
    if (inFlex) {
      flex.push(x, y);
      return;
    }
    if (open) out.push({ op: 'close' });
    out.push({ op: 'move', x, y });
    open = true;
  };
  const curveTo = (x1: number, y1: number, x2: number, y2: number): void => {
    out.push({ op: 'cubic', x1, y1, x2, y2, x, y });
  };

  const run = (code: Uint8Array, level: number): boolean => {
    for (let i = 0; i < code.length; ) {
      if (++steps > MAX_STEPS) return false;
      const b = code[i++]!;
      if (b >= 32) {
        if (b <= 246) stack.push(b - 139);
        else if (b <= 250) stack.push((b - 247) * 256 + (code[i++] ?? 0) + 108);
        else if (b <= 254) stack.push(-(b - 251) * 256 - (code[i++] ?? 0) - 108);
        else {
          // §6.2 — 255 introduces a 32-bit integer, NOT the 16.16 fixed point a
          // Type 2 charstring writes.
          stack.push(
            (code[i]! << 24) | (code[i + 1]! << 16) | (code[i + 2]! << 8) | code[i + 3]! | 0,
          );
          i += 4;
        }
        continue;
      }
      switch (b) {
        case 13: // hsbw — the left sidebearing sets the origin, the width is not ours
          x = stack[0] ?? 0;
          y = 0;
          stack.length = 0;
          break;
        case 9: // closepath
          if (open) out.push({ op: 'close' });
          open = false;
          stack.length = 0;
          break;
        case 1: // hstem
        case 3: // vstem
          stack.length = 0;
          break;
        case 21: // rmoveto
          x += stack[stack.length - 2] ?? 0;
          y += stack[stack.length - 1] ?? 0;
          moveTo();
          stack.length = 0;
          break;
        case 22: // hmoveto
          x += stack[stack.length - 1] ?? 0;
          moveTo();
          stack.length = 0;
          break;
        case 4: // vmoveto
          y += stack[stack.length - 1] ?? 0;
          moveTo();
          stack.length = 0;
          break;
        case 5: // rlineto
          x += stack[0] ?? 0;
          y += stack[1] ?? 0;
          out.push({ op: 'line', x, y });
          stack.length = 0;
          break;
        case 6: // hlineto
          x += stack[0] ?? 0;
          out.push({ op: 'line', x, y });
          stack.length = 0;
          break;
        case 7: // vlineto
          y += stack[0] ?? 0;
          out.push({ op: 'line', x, y });
          stack.length = 0;
          break;
        case 8: {
          // rrcurveto
          const x1 = x + (stack[0] ?? 0);
          const y1 = y + (stack[1] ?? 0);
          const x2 = x1 + (stack[2] ?? 0);
          const y2 = y1 + (stack[3] ?? 0);
          x = x2 + (stack[4] ?? 0);
          y = y2 + (stack[5] ?? 0);
          curveTo(x1, y1, x2, y2);
          stack.length = 0;
          break;
        }
        case 30: {
          // vhcurveto — starts vertical, ends horizontal
          const x1 = x;
          const y1 = y + (stack[0] ?? 0);
          const x2 = x1 + (stack[1] ?? 0);
          const y2 = y1 + (stack[2] ?? 0);
          x = x2 + (stack[3] ?? 0);
          y = y2;
          curveTo(x1, y1, x2, y2);
          stack.length = 0;
          break;
        }
        case 31: {
          // hvcurveto — starts horizontal, ends vertical
          const x1 = x + (stack[0] ?? 0);
          const y1 = y;
          const x2 = x1 + (stack[1] ?? 0);
          const y2 = y1 + (stack[2] ?? 0);
          x = x2;
          y = y2 + (stack[3] ?? 0);
          curveTo(x1, y1, x2, y2);
          stack.length = 0;
          break;
        }
        case 10: {
          // callsubr
          const index = stack.pop() ?? 0;
          const sub = font.subrs[index];
          if (sub && level < MAX_CALL_DEPTH && !run(sub, level + 1)) return false;
          break;
        }
        case 11: // return
          return true;
        case 14: // endchar
          if (open) out.push({ op: 'close' });
          open = false;
          return false;
        case 12: {
          const op2 = code[i++]!;
          if (op2 === 12) {
            // div
            const den = stack.pop() ?? 1;
            const num = stack.pop() ?? 0;
            stack.push(den === 0 ? 0 : num / den);
            break;
          }
          if (op2 === 16) {
            // §8.3 `callothersubr` — flex and hint replacement, which the
            // interpreter performs itself rather than calling out.
            const which = stack.pop() ?? 0;
            const count = stack.pop() ?? 0;
            const args = stack.splice(Math.max(0, stack.length - count), count);
            if (which === 1) {
              inFlex = true;
              flex.length = 0;
            } else if (which === 0) {
              inFlex = false;
              // Seven points were collected; the first is the reference point
              // and the other six are two curves.
              if (flex.length >= 14) {
                const p = flex.slice(2);
                out.push({
                  op: 'cubic',
                  x1: p[0]!,
                  y1: p[1]!,
                  x2: p[2]!,
                  y2: p[3]!,
                  x: p[4]!,
                  y: p[5]!,
                });
                out.push({
                  op: 'cubic',
                  x1: p[6]!,
                  y1: p[7]!,
                  x2: p[8]!,
                  y2: p[9]!,
                  x: p[10]!,
                  y: p[11]!,
                });
                x = p[10]!;
                y = p[11]!;
              }
              ps.push(y, x); // the two `pop`s that follow read the end point
            } else if (which === 3) {
              ps.push(3); // hint replacement: the subr number to ignore
            } else {
              for (const a of args.reverse()) ps.push(a);
            }
            break;
          }
          if (op2 === 17) {
            // pop — take what `callothersubr` left
            stack.push(ps.pop() ?? 0);
            break;
          }
          if (op2 === 6) {
            // §8.4 seac — a standard-encoding accented character, drawn as two
            // glyphs. Only the base is taken: the accent needs the standard
            // encoding's names, and a base with no accent reads better than
            // nothing at all.
            const base = STANDARD_ENCODING[stack[3] ?? 0];
            const charstring = base !== undefined ? font.charstrings.get(base) : undefined;
            if (charstring && depth < 2) {
              const drawn = runGlyph(font, charstring, depth + 1);
              if (drawn) out.push(...drawn.map((seg) => unscale(seg, font.scale)));
            }
            stack.length = 0;
            return false;
          }
          if (op2 === 7) {
            // sbw — the sidebearing and width in both directions
            x = stack[0] ?? 0;
            y = stack[1] ?? 0;
            stack.length = 0;
            break;
          }
          if (op2 === 33) {
            // setcurrentpoint
            x = stack[0] ?? x;
            y = stack[1] ?? y;
            stack.length = 0;
            break;
          }
          stack.length = 0; // dotsection, vstem3, hstem3 — hints, nothing to draw
          break;
        }
        default:
          stack.length = 0;
          break;
      }
    }
    return true;
  };

  run(program, 0);
  if (out.length > 0 && out[out.length - 1]?.op !== 'close') out.push({ op: 'close' });
  if (out.length === 0) return undefined;
  return out.map((seg) => scaleSeg(seg, font.scale));
}

function scaleSeg(seg: PathSeg, k: number): PathSeg {
  if (seg.op === 'close') return seg;
  if (seg.op === 'cubic') {
    return {
      op: 'cubic',
      x1: seg.x1 * k,
      y1: seg.y1 * k,
      x2: seg.x2 * k,
      y2: seg.y2 * k,
      x: seg.x * k,
      y: seg.y * k,
    };
  }
  return { op: seg.op, x: seg.x * k, y: seg.y * k };
}

/** The base of a `seac` comes back already scaled; this puts it back in glyph space. */
function unscale(seg: PathSeg, k: number): PathSeg {
  return scaleSeg(seg, k === 0 ? 1 : 1 / k);
}

/** §C.1 — the names `seac` selects its two parts by, at the codes it states. */
const STANDARD_ENCODING: Readonly<Record<number, string>> = {
  32: 'space',
  33: 'exclam',
  34: 'quotedbl',
  35: 'numbersign',
  36: 'dollar',
  37: 'percent',
  38: 'ampersand',
  39: 'quoteright',
  40: 'parenleft',
  41: 'parenright',
  42: 'asterisk',
  43: 'plus',
  44: 'comma',
  45: 'hyphen',
  46: 'period',
  47: 'slash',
  48: 'zero',
  49: 'one',
  50: 'two',
  51: 'three',
  52: 'four',
  53: 'five',
  54: 'six',
  55: 'seven',
  56: 'eight',
  57: 'nine',
  58: 'colon',
  59: 'semicolon',
  60: 'less',
  61: 'equal',
  62: 'greater',
  63: 'question',
  64: 'at',
  65: 'A',
  66: 'B',
  67: 'C',
  68: 'D',
  69: 'E',
  70: 'F',
  71: 'G',
  72: 'H',
  73: 'I',
  74: 'J',
  75: 'K',
  76: 'L',
  77: 'M',
  78: 'N',
  79: 'O',
  80: 'P',
  81: 'Q',
  82: 'R',
  83: 'S',
  84: 'T',
  85: 'U',
  86: 'V',
  87: 'W',
  88: 'X',
  89: 'Y',
  90: 'Z',
  91: 'bracketleft',
  92: 'backslash',
  93: 'bracketright',
  94: 'asciicircum',
  95: 'underscore',
  96: 'quoteleft',
  97: 'a',
  98: 'b',
  99: 'c',
  100: 'd',
  101: 'e',
  102: 'f',
  103: 'g',
  104: 'h',
  105: 'i',
  106: 'j',
  107: 'k',
  108: 'l',
  109: 'm',
  110: 'n',
  111: 'o',
  112: 'p',
  113: 'q',
  114: 'r',
  115: 's',
  116: 't',
  117: 'u',
  118: 'v',
  119: 'w',
  120: 'x',
  121: 'y',
  122: 'z',
  123: 'braceleft',
  124: 'bar',
  125: 'braceright',
  126: 'asciitilde',
  193: 'grave',
  194: 'acute',
  195: 'circumflex',
  196: 'tilde',
  197: 'macron',
  198: 'breve',
  199: 'dotaccent',
  200: 'dieresis',
  202: 'ring',
  203: 'cedilla',
  205: 'hungarumlaut',
  206: 'ogonek',
  207: 'caron',
};
