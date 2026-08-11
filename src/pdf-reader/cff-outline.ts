// §9.6.6 / Adobe TN 5176, TN 5177 — the OUTLINE of one glyph in an embedded
// CFF program (`/FontFile3`).
//
// The same case `./glyf-outline` exists for, in the other outline format: a
// composite font read by glyph index whose program says nothing about what its
// characters are. A CFF glyph is not a table of points but a PROGRAM — a Type 2
// charstring, run on a stack machine with its own subroutines — so reading one
// means running it.
//
// Only what draws is run. Hints (`hstem`, `hintmask` and the rest) are counted
// for their operand widths and then dropped: they say how a rasteriser should
// snap the outline to a pixel grid, and nothing downstream rasterises here.
//
// The outline comes back in a ONE-UNIT em, like the TrueType reader's, so both
// place the same way.

import type { OutlineSource } from './glyf-outline';
import type { PathSeg } from './content';

/**
 * Prepare an embedded CFF program for outline reading.
 *
 * @param program The raw `/FontFile3` bytes (bare CFF, or the `CFF ` table of
 *                an OpenType wrapper — the caller unwraps that).
 * @returns A reader keyed by GLYPH INDEX, or `undefined` where the program
 *          cannot be read.
 */
export function cffOutlineSource(program: Uint8Array): OutlineSource | undefined {
  let font: CffFont;
  try {
    const parsed = parseCff(program);
    if (!parsed) return undefined;
    font = parsed;
  } catch {
    return undefined;
  }
  const cache = new Map<number, Array<PathSeg> | undefined>();
  return {
    path: (gid: number): Array<PathSeg> | undefined => {
      const had = cache.get(gid);
      if (had !== undefined || cache.has(gid)) return had;
      let out: Array<PathSeg> | undefined;
      try {
        out = runGlyph(font, gid);
      } catch {
        out = undefined;
      }
      cache.set(gid, out);
      return out;
    },
  };
}

/**
 * TN 5176 §10 — the CID a CID-keyed CFF gives each glyph, inverted.
 *
 * A `CIDFontType0C` program is charset-ordered by CID, not by glyph index, so a
 * composite font's code (its CID) is not the index into `CharStrings`. Where the
 * program is not CID-keyed the two are the same and this gives nothing.
 *
 * @param program The raw CFF bytes.
 * @returns CID → glyph index, or `undefined` for a program that is not
 *          CID-keyed or cannot be read.
 */
export function cffCidToGid(program: Uint8Array): Map<number, number> | undefined {
  try {
    const font = parseCff(program);
    return font?.cidToGid;
  } catch {
    return undefined;
  }
}

/** The `CFF ` table of an OpenType wrapper, where the program is one. */
export function openTypeCff(program: Uint8Array): Uint8Array | undefined {
  if (program.length < 12) return undefined;
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
  const tag = view.getUint32(0);
  // 'OTTO' is the only sfnt flavour that carries CFF outlines.
  if (tag !== 0x4f54544f) return undefined;
  const count = view.getUint16(4);
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    if (at + 16 > program.length) break;
    const name = String.fromCharCode(
      program[at]!,
      program[at + 1]!,
      program[at + 2]!,
      program[at + 3]!,
    );
    if (name !== 'CFF ') continue;
    const offset = view.getUint32(at + 8);
    const length = view.getUint32(at + 12);
    if (offset + length > program.length) return undefined;
    return program.subarray(offset, offset + length);
  }
  return undefined;
}

/** What running a charstring needs: the programs, the subroutines, the scale. */
interface CffFont {
  readonly charStrings: ReadonlyArray<Uint8Array>;
  readonly globalSubrs: ReadonlyArray<Uint8Array>;
  /** The local subroutines in force for one glyph — an FDSelect picks per glyph. */
  readonly localSubrs: (gid: number) => ReadonlyArray<Uint8Array>;
  /** `/FontMatrix`'s scale, which is 1/1000 for all but a handful of faces. */
  readonly scale: number;
  readonly cidToGid?: Map<number, number>;
}

function parseCff(program: Uint8Array): CffFont | undefined {
  if (program.length < 4) return undefined;
  const headerSize = program[2] ?? 4;
  const names = readIndex(program, headerSize);
  if (!names) return undefined;
  const tops = readIndex(program, names.end);
  if (!tops || tops.items.length === 0) return undefined;
  const strings = readIndex(program, tops.end);
  if (!strings) return undefined;
  const gsubrs = readIndex(program, strings.end);
  if (!gsubrs) return undefined;

  const top = parseDict(tops.items[0]!);
  const csOffset = top.get(17)?.[0];
  if (csOffset === undefined) return undefined;
  const charStrings = readIndex(program, csOffset);
  if (!charStrings) return undefined;

  // TN 5176 §9 — `/FontMatrix` (key 12 7) states glyph space; its default is
  // 0.001, which is every ordinary face.
  const matrix = top.get(1207);
  const scale = matrix && matrix.length >= 4 ? (matrix[0] ?? 0.001) : 0.001;

  const privateSubrs = (dict: ReadonlyMap<number, Array<number>>): ReadonlyArray<Uint8Array> => {
    const priv = dict.get(18);
    if (!priv || priv.length < 2) return [];
    const size = priv[0]!;
    const offset = priv[1]!;
    if (offset < 0 || offset + size > program.length) return [];
    const inner = parseDict(program.subarray(offset, offset + size));
    const subrsAt = inner.get(19)?.[0];
    if (subrsAt === undefined) return [];
    return readIndex(program, offset + subrsAt)?.items ?? [];
  };

  // TN 5176 §10 — a CID-keyed font has no single Private DICT: an FDSelect says
  // which of the FDArray's dictionaries each glyph belongs to.
  const isCid = top.has(1230);
  let localSubrs: (gid: number) => ReadonlyArray<Uint8Array>;
  if (isCid) {
    const fdArrayAt = top.get(1236)?.[0];
    const fdSelectAt = top.get(1237)?.[0];
    const fds = fdArrayAt !== undefined ? (readIndex(program, fdArrayAt)?.items ?? []) : [];
    const subrsPerFd = fds.map((fd) => privateSubrs(parseDict(fd)));
    const select =
      fdSelectAt !== undefined
        ? readFdSelect(program, fdSelectAt, charStrings.items.length)
        : undefined;
    localSubrs = (gid: number): ReadonlyArray<Uint8Array> =>
      subrsPerFd[select ? (select[gid] ?? 0) : 0] ?? [];
  } else {
    const only = privateSubrs(top);
    localSubrs = (): ReadonlyArray<Uint8Array> => only;
  }

  const charsetAt = top.get(15)?.[0];
  const cidToGid =
    isCid && charsetAt !== undefined && charsetAt > 2
      ? readCharset(program, charsetAt, charStrings.items.length)
      : undefined;

  return {
    charStrings: charStrings.items,
    globalSubrs: gsubrs.items,
    localSubrs,
    scale,
    ...(cidToGid ? { cidToGid } : {}),
  };
}

/** TN 5176 §5 — an INDEX: a count, an offset size, the offsets, then the data. */
function readIndex(
  data: Uint8Array,
  at: number,
): { items: Array<Uint8Array>; end: number } | undefined {
  if (at < 0 || at + 2 > data.length) return undefined;
  const count = (data[at]! << 8) | data[at + 1]!;
  if (count === 0) return { items: [], end: at + 2 };
  const offSize = data[at + 2] ?? 1;
  if (offSize < 1 || offSize > 4) return undefined;
  const offsets: Array<number> = [];
  let p = at + 3;
  for (let i = 0; i <= count; i++) {
    let v = 0;
    for (let k = 0; k < offSize; k++) v = (v << 8) | (data[p++] ?? 0);
    offsets.push(v);
  }
  const base = p - 1;
  const items: Array<Uint8Array> = [];
  for (let i = 0; i < count; i++) {
    const from = base + (offsets[i] ?? 1);
    const to = base + (offsets[i + 1] ?? 1);
    if (from < 0 || to > data.length || to < from) return undefined;
    items.push(data.subarray(from, to));
  }
  return { items, end: base + (offsets[count] ?? 1) };
}

/**
 * TN 5176 §4 — a DICT: operands then an operator. Two-byte operators (`12 x`)
 * are keyed as `1200 + x` so one map holds both.
 */
function parseDict(data: Uint8Array): Map<number, Array<number>> {
  const out = new Map<number, Array<number>>();
  let operands: Array<number> = [];
  for (let i = 0; i < data.length; ) {
    const b = data[i]!;
    if (b <= 21) {
      const key = b === 12 ? 1200 + (data[i + 1] ?? 0) : b;
      i += b === 12 ? 2 : 1;
      out.set(key, operands);
      operands = [];
    } else if (b === 28) {
      operands.push((((data[i + 1]! << 8) | data[i + 2]!) << 16) >> 16);
      i += 3;
    } else if (b === 29) {
      operands.push(
        (data[i + 1]! << 24) | (data[i + 2]! << 16) | (data[i + 3]! << 8) | data[i + 4]! | 0,
      );
      i += 5;
    } else if (b === 30) {
      // A real number, nibble-coded; only the value matters here.
      let text = '';
      i++;
      let done = false;
      while (i < data.length && !done) {
        const byte = data[i++]!;
        for (const nibble of [byte >> 4, byte & 15]) {
          if (nibble <= 9) text += String(nibble);
          else if (nibble === 10) text += '.';
          else if (nibble === 11) text += 'E';
          else if (nibble === 12) text += 'E-';
          else if (nibble === 14) text += '-';
          else if (nibble === 15) {
            done = true;
            break;
          }
        }
      }
      operands.push(Number.parseFloat(text) || 0);
    } else if (b >= 32 && b <= 246) {
      operands.push(b - 139);
      i++;
    } else if (b >= 247 && b <= 250) {
      operands.push((b - 247) * 256 + (data[i + 1] ?? 0) + 108);
      i += 2;
    } else if (b >= 251 && b <= 254) {
      operands.push(-(b - 251) * 256 - (data[i + 1] ?? 0) - 108);
      i += 2;
    } else {
      i++;
    }
  }
  return out;
}

/** TN 5176 §19 — FDSelect: which Private DICT each glyph belongs to. */
function readFdSelect(data: Uint8Array, at: number, glyphs: number): Array<number> | undefined {
  if (at < 0 || at >= data.length) return undefined;
  const out = new Array<number>(glyphs).fill(0);
  const format = data[at]!;
  if (format === 0) {
    for (let i = 0; i < glyphs; i++) out[i] = data[at + 1 + i] ?? 0;
    return out;
  }
  if (format !== 3) return undefined;
  const ranges = (data[at + 1]! << 8) | data[at + 2]!;
  let p = at + 3;
  let first = (data[p]! << 8) | data[p + 1]!;
  p += 2;
  for (let r = 0; r < ranges; r++) {
    const fd = data[p]!;
    const next = (data[p + 1]! << 8) | data[p + 2]!;
    p += 3;
    for (let g = first; g < next && g < glyphs; g++) out[g] = fd;
    first = next;
  }
  return out;
}

/** TN 5176 §13 — the charset, which for a CID-keyed font lists CIDs by glyph. */
function readCharset(
  data: Uint8Array,
  at: number,
  glyphs: number,
): Map<number, number> | undefined {
  if (at < 0 || at >= data.length) return undefined;
  const out = new Map<number, number>();
  out.set(0, 0); // glyph 0 is .notdef, and its CID is 0
  const format = data[at]!;
  let p = at + 1;
  if (format === 0) {
    for (let gid = 1; gid < glyphs && p + 1 < data.length; gid++, p += 2) {
      out.set((data[p]! << 8) | data[p + 1]!, gid);
    }
    return out;
  }
  if (format !== 1 && format !== 2) return undefined;
  const wide = format === 2;
  for (let gid = 1; gid < glyphs && p < data.length; ) {
    const first = (data[p]! << 8) | data[p + 1]!;
    const left = wide ? (data[p + 2]! << 8) | data[p + 3]! : data[p + 2]!;
    p += wide ? 4 : 3;
    for (let k = 0; k <= left && gid < glyphs; k++) out.set(first + k, gid++);
  }
  return out;
}

/** TN 5177 §4.7 — the index a subroutine call is biased by. */
function bias(subrs: ReadonlyArray<Uint8Array>): number {
  return subrs.length < 1240 ? 107 : subrs.length < 33900 ? 1131 : 32768;
}

/** How deep a charstring may call, and how many operators it may run. */
const MAX_CALL_DEPTH = 10;
const MAX_STEPS = 65536;

// TN 5177 — run one Type 2 charstring, collecting what it draws.
function runGlyph(font: CffFont, gid: number): Array<PathSeg> | undefined {
  const program = font.charStrings[gid];
  if (!program) return undefined;
  const local = font.localSubrs(gid);
  const localBias = bias(local);
  const globalBias = bias(font.globalSubrs);
  const out: Array<PathSeg> = [];
  const stack: Array<number> = [];
  const trans: Array<number> = []; // the transient array `put`/`get` use
  let x = 0;
  let y = 0;
  let stems = 0;
  let open = false;
  let widthParsed = false;
  let steps = 0;

  const moveTo = (nx: number, ny: number): void => {
    if (open) out.push({ op: 'close' });
    out.push({ op: 'move', x: nx, y: ny });
    open = true;
  };
  const lineTo = (nx: number, ny: number): void => {
    out.push({ op: 'line', x: nx, y: ny });
  };
  const curveTo = (
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    nx: number,
    ny: number,
  ): void => {
    out.push({ op: 'cubic', x1, y1, x2, y2, x: nx, y: ny });
  };
  // §4.1 — the first stack-clearing operator may be preceded by the glyph's
  // width, and the only way to tell is COUNT: one operand more than the
  // operator takes. `even` is the stem operators, which take any even number.
  const takeWidth = (arity: number | 'even'): void => {
    if (widthParsed) return;
    widthParsed = true;
    const stated = arity === 'even' ? stack.length % 2 === 1 : stack.length > arity;
    if (stated) stack.shift();
  };

  const run = (code: Uint8Array, depth: number): boolean => {
    for (let i = 0; i < code.length; ) {
      if (++steps > MAX_STEPS) return false;
      const b = code[i++]!;
      if (b >= 32 || b === 28) {
        // An operand.
        if (b === 28) {
          stack.push((((code[i]! << 8) | code[i + 1]!) << 16) >> 16);
          i += 2;
        } else if (b <= 246) {
          stack.push(b - 139);
        } else if (b <= 250) {
          stack.push((b - 247) * 256 + (code[i++] ?? 0) + 108);
        } else if (b <= 254) {
          stack.push(-(b - 251) * 256 - (code[i++] ?? 0) - 108);
        } else {
          // 255: a 16.16 fixed-point number.
          stack.push(
            ((code[i]! << 24) | (code[i + 1]! << 16) | (code[i + 2]! << 8) | code[i + 3]! | 0) /
              65536,
          );
          i += 4;
        }
        continue;
      }
      switch (b) {
        case 1: // hstem
        case 3: // vstem
        case 18: // hstemhm
        case 23: // vstemhm
          takeWidth('even');
          stems += stack.length >> 1;
          stack.length = 0;
          break;
        case 19: // hintmask
        case 20: // cntrmask
          takeWidth('even');
          stems += stack.length >> 1;
          stack.length = 0;
          i += (stems + 7) >> 3;
          break;
        case 21: // rmoveto
          takeWidth(2);
          x += stack[stack.length - 2] ?? 0;
          y += stack[stack.length - 1] ?? 0;
          moveTo(x, y);
          stack.length = 0;
          break;
        case 22: // hmoveto
          takeWidth(1);
          x += stack[stack.length - 1] ?? 0;
          moveTo(x, y);
          stack.length = 0;
          break;
        case 4: // vmoveto
          takeWidth(1);
          y += stack[stack.length - 1] ?? 0;
          moveTo(x, y);
          stack.length = 0;
          break;
        case 5: // rlineto
          for (let k = 0; k + 1 < stack.length; k += 2) {
            x += stack[k]!;
            y += stack[k + 1]!;
            lineTo(x, y);
          }
          stack.length = 0;
          break;
        case 6: // hlineto
        case 7: {
          // vlineto — the two alternate, starting on the operator's own axis
          let horizontal = b === 6;
          for (const d of stack) {
            if (horizontal) x += d;
            else y += d;
            lineTo(x, y);
            horizontal = !horizontal;
          }
          stack.length = 0;
          break;
        }
        case 8: // rrcurveto
          for (let k = 0; k + 5 < stack.length; k += 6) curve(stack, k);
          stack.length = 0;
          break;
        case 24: // rcurveline
          {
            let k = 0;
            for (; k + 5 < stack.length - 2; k += 6) curve(stack, k);
            x += stack[k] ?? 0;
            y += stack[k + 1] ?? 0;
            lineTo(x, y);
          }
          stack.length = 0;
          break;
        case 25: // rlinecurve
          {
            let k = 0;
            for (; stack.length - k > 6; k += 2) {
              x += stack[k]!;
              y += stack[k + 1]!;
              lineTo(x, y);
            }
            curve(stack, k);
          }
          stack.length = 0;
          break;
        case 26: // vvcurveto
        case 27: {
          // hhcurveto
          let k = 0;
          let d1 = 0;
          if (stack.length % 4 === 1) d1 = stack[k++]!;
          for (; k + 3 < stack.length; k += 4) {
            const vertical = b === 26;
            const x1 = vertical ? x + d1 : x + stack[k]!;
            const y1 = vertical ? y + stack[k]! : y + d1;
            const x2 = x1 + stack[k + 1]!;
            const y2 = y1 + stack[k + 2]!;
            x = vertical ? x2 : x2 + stack[k + 3]!;
            y = vertical ? y2 + stack[k + 3]! : y2;
            curveTo(x1, y1, x2, y2, x, y);
            d1 = 0;
          }
          stack.length = 0;
          break;
        }
        case 30: // vhcurveto
        case 31: {
          // hvcurveto — alternating, with a possible odd last operand
          let horizontal = b === 31;
          let k = 0;
          while (k + 3 < stack.length) {
            const last = stack.length - k === 5;
            const x1 = horizontal ? x + stack[k]! : x;
            const y1 = horizontal ? y : y + stack[k]!;
            const x2 = x1 + stack[k + 1]!;
            const y2 = y1 + stack[k + 2]!;
            if (horizontal) {
              y = y2 + stack[k + 3]!;
              x = x2 + (last ? stack[k + 4]! : 0);
            } else {
              x = x2 + stack[k + 3]!;
              y = y2 + (last ? stack[k + 4]! : 0);
            }
            curveTo(x1, y1, x2, y2, x, y);
            k += 4;
            horizontal = !horizontal;
          }
          stack.length = 0;
          break;
        }
        case 10: // callsubr
        case 29: {
          // callgsubr
          const subrs = b === 10 ? local : font.globalSubrs;
          const index = (stack.pop() ?? 0) + (b === 10 ? localBias : globalBias);
          const sub = subrs[index];
          if (sub && depth < MAX_CALL_DEPTH && !run(sub, depth + 1)) return false;
          break;
        }
        case 11: // return
          return true;
        case 14: // endchar
          // `endchar` takes nothing, or four for the deprecated `seac` form.
          takeWidth(stack.length > 4 ? 4 : 0);
          if (open) out.push({ op: 'close' });
          open = false;
          return false;
        case 12: {
          const op2 = code[i++]!;
          runEscape(op2, stack, trans);
          if (op2 === 35 || op2 === 34 || op2 === 36 || op2 === 37) {
            // The flex operators draw two curves; the stack layout differs per
            // operator, and `flexCurves` returns the six-tuples to draw.
            const drawn = flexCurves(op2, stack, x, y);
            for (const c of drawn.curves) curveTo(c[0], c[1], c[2], c[3], c[4], c[5]);
            x = drawn.x;
            y = drawn.y;
          }
          stack.length = 0;
          break;
        }
        default:
          stack.length = 0;
          break;
      }
    }
    return true;
  };

  // A relative curve from six operands starting at `k`.
  function curve(operands: ReadonlyArray<number>, k: number): void {
    const x1 = x + (operands[k] ?? 0);
    const y1 = y + (operands[k + 1] ?? 0);
    const x2 = x1 + (operands[k + 2] ?? 0);
    const y2 = y1 + (operands[k + 3] ?? 0);
    x = x2 + (operands[k + 4] ?? 0);
    y = y2 + (operands[k + 5] ?? 0);
    curveTo(x1, y1, x2, y2, x, y);
  }

  run(program, 0);
  // A charstring that ends without `endchar` leaves its last contour open.
  if (out.length > 0 && out[out.length - 1]?.op !== 'close') out.push({ op: 'close' });
  if (out.length === 0) return undefined;
  return out.map((seg) => scaleSeg(seg, font.scale));
}

// The arithmetic escapes (`12 x`) a charstring may use. Only the ones that
// touch the stack matter; the rest are hint and flex operators handled by the
// caller, and anything unknown clears the stack there.
function runEscape(op: number, stack: Array<number>, trans: Array<number>): void {
  switch (op) {
    case 3: // and
    case 4: // or
    case 10: // add
    case 11: // sub
    case 12: // div
    case 24: // mul
      {
        const b = stack.pop() ?? 0;
        const a = stack.pop() ?? 0;
        stack.push(
          op === 3
            ? a !== 0 && b !== 0
              ? 1
              : 0
            : op === 4
              ? a !== 0 || b !== 0
                ? 1
                : 0
              : op === 10
                ? a + b
                : op === 11
                  ? a - b
                  : op === 12
                    ? b === 0
                      ? 0
                      : a / b
                    : a * b,
        );
      }
      break;
    case 5: // not
      stack.push((stack.pop() ?? 0) === 0 ? 1 : 0);
      break;
    case 9: // abs
      stack.push(Math.abs(stack.pop() ?? 0));
      break;
    case 14: // neg
      stack.push(-(stack.pop() ?? 0));
      break;
    case 18: // drop
      stack.pop();
      break;
    case 20: {
      // put
      const index = stack.pop() ?? 0;
      trans[index & 31] = stack.pop() ?? 0;
      break;
    }
    case 21: // get
      stack.push(trans[(stack.pop() ?? 0) & 31] ?? 0);
      break;
    case 26: // sqrt
      stack.push(Math.sqrt(Math.abs(stack.pop() ?? 0)));
      break;
    case 27: {
      // dup
      const v = stack[stack.length - 1] ?? 0;
      stack.push(v);
      break;
    }
    case 28: {
      // exch
      const b = stack.pop() ?? 0;
      const a = stack.pop() ?? 0;
      stack.push(b, a);
      break;
    }
    default:
      break;
  }
}

/**
 * TN 5177 §4.2 — the flex operators, which draw two curves as one. `hflex` and
 * friends leave axes implicit; this expands them to the six-tuples to draw.
 */
function flexCurves(
  op: number,
  stack: ReadonlyArray<number>,
  x0: number,
  y0: number,
): { curves: Array<[number, number, number, number, number, number]>; x: number; y: number } {
  const n = (i: number): number => stack[i] ?? 0;
  const curves: Array<[number, number, number, number, number, number]> = [];
  let x = x0;
  let y = y0;
  const push = (dxs: ReadonlyArray<number>, dys: ReadonlyArray<number>): void => {
    const x1 = x + (dxs[0] ?? 0);
    const y1 = y + (dys[0] ?? 0);
    const x2 = x1 + (dxs[1] ?? 0);
    const y2 = y1 + (dys[1] ?? 0);
    x = x2 + (dxs[2] ?? 0);
    y = y2 + (dys[2] ?? 0);
    curves.push([x1, y1, x2, y2, x, y]);
  };
  if (op === 35) {
    // flex: two full curves, then a tolerance this ignores
    push([n(0), n(2), n(4)], [n(1), n(3), n(5)]);
    push([n(6), n(8), n(10)], [n(7), n(9), n(11)]);
  } else if (op === 34) {
    // hflex: the pair stays on one line, y returning to where it began
    const startY = y;
    push([n(0), n(1), n(3)], [0, n(2), 0]);
    push([n(4), n(6), 0], [0, startY - y, 0]);
    y = startY;
  } else if (op === 36) {
    // hflex1
    const startY = y;
    push([n(0), n(2), n(4)], [n(1), n(3), 0]);
    push([n(5), n(7), n(8)], [0, n(6), startY - (y + n(6))]);
    y = startY;
  } else if (op === 37) {
    // flex1 — the last point returns to the start on whichever axis moved less
    const startX = x;
    const startY = y;
    let dx = 0;
    let dy = 0;
    for (let k = 0; k < 10; k += 2) {
      dx += n(k);
      dy += n(k + 1);
    }
    push([n(0), n(2), n(4)], [n(1), n(3), n(5)]);
    const lastIsX = Math.abs(dx) > Math.abs(dy);
    const endX = lastIsX ? x + n(6) + n(8) + n(10) : startX;
    const endY = lastIsX ? startY : y + n(7) + n(9) + n(10);
    const x1 = x + n(6);
    const y1 = y + n(7);
    const x2 = x1 + n(8);
    const y2 = y1 + n(9);
    curves.push([x1, y1, x2, y2, endX, endY]);
    x = endX;
    y = endY;
  }
  return { curves, x, y };
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
