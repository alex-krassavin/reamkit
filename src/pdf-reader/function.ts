// §7.10 — a PDF function, as something that can be CALLED.
//
// A file states functions where a value has to be worked out rather than
// written down: the colour along a gradient, and — the reason this exists — the
// transform that turns a tint into a colour a device can show (§8.6.6.4,
// §8.6.6.5). Without running it a `/Separation` or `/DeviceN` fill has only its
// tint to go on, and a tint of 1 read as a grey level is BLACK: devicen.pdf's
// three triangles are green, blue and red, and all three came back black.
//
// All four kinds are here. Types 0, 2 and 3 are arithmetic on the numbers the
// dictionary states. Type 4 is a small PostScript: a program in braces, run on
// a stack, which is the one kind that cannot be read off the dictionary at all.

import { Lexer } from './lexer';

import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { PdfFile } from './document';
import { PDF_NULL, PdfStream } from '@/pdf/objects';

/** §7.10 — m numbers in, n numbers out. */
export type PdfFunction = (inputs: ReadonlyArray<number>) => Array<number>;

/** A sampled function bigger than this is not read: it is a DoS, not a table. */
const MAX_SAMPLES = 1 << 22;

/** A type-4 program deeper than this is not run. */
const MAX_PS_DEPTH = 32;

/** …nor one with more tokens than this. */
const MAX_PS_TOKENS = 1 << 16;

/** How far a stitching or array function may nest before this gives up. */
const MAX_NESTING = 8;

/**
 * Read a `/Function` entry into something callable (§7.10).
 *
 * The entry may also be an ARRAY of functions, each giving one output, which is
 * what a `/DeviceN` with a per-colorant transform states; that comes back as
 * one function returning all of them in order.
 *
 * @param file  The owning file.
 * @param value The `/Function` (or `/TintTransform`) entry, unresolved.
 * @returns The function, or `undefined` for one this cannot run.
 */
export function readFunction(file: PdfFile, value: PdfValue | undefined): PdfFunction | undefined {
  return readAt(file, value, 0);
}

function readAt(
  file: PdfFile,
  value: PdfValue | undefined,
  depth: number,
): PdfFunction | undefined {
  if (value === undefined || depth > MAX_NESTING) return undefined;
  const resolved = file.resolve(value);
  if (Array.isArray(resolved)) {
    // §7.10.1 — an array of 1-out functions, together making the n outputs.
    const parts = resolved.map((v) => readAt(file, v, depth + 1));
    if (parts.some((p) => p === undefined)) return undefined;
    return (inputs) => parts.flatMap((p) => p!(inputs));
  }
  const dict =
    resolved instanceof PdfStream ? resolved.dict : resolved instanceof Map ? resolved : undefined;
  if (!dict) return undefined;
  const type = file.get(dict, 'FunctionType');
  const domain = numbers(file, dict.get('Domain'));
  if (type === 2) return exponential(file, dict, domain);
  if (type === 3) return stitching(file, dict, domain, depth);
  if (type === 0 && resolved instanceof PdfStream) return sampled(file, resolved, domain);
  if (type === 4 && resolved instanceof PdfStream) return calculator(file, resolved, domain);
  return undefined;
}

/** §7.10.2 — a sampled function: a table, interpolated between its samples. */
function sampled(file: PdfFile, stream: PdfStream, domain: Array<number>): PdfFunction | undefined {
  const d = stream.dict;
  const size = numbers(file, d.get('Size'));
  const range = numbers(file, d.get('Range'));
  const bps = file.get(d, 'BitsPerSample');
  const m = size.length;
  const n = range.length >> 1;
  if (m === 0 || n === 0 || domain.length < m * 2 || typeof bps !== 'number') return undefined;
  if (!size.every((s) => Number.isInteger(s) && s >= 1)) return undefined;
  let total = n;
  for (const s of size) {
    total *= s;
    if (total > MAX_SAMPLES) return undefined;
  }
  const encode = orDefault(numbers(file, d.get('Encode')), m * 2, (i) =>
    i % 2 === 0 ? 0 : size[(i - 1) >> 1]! - 1,
  );
  const decode = orDefault(numbers(file, d.get('Decode')), n * 2, (i) => range[i] ?? 0);
  const data = file.streamData(stream);
  const max = 2 ** bps - 1;
  // The sample at a flat index, whose components run together: §7.10.2 orders
  // them with the FIRST input varying fastest.
  const sample = (flat: number, out: number): number => {
    let bit = (flat * n + out) * bps;
    let v = 0;
    for (let k = 0; k < bps; k++) {
      v = v * 2 + (((data[bit >> 3] ?? 0) >> (7 - (bit & 7))) & 1);
      bit++;
    }
    return v / max;
  };
  return (inputs) => {
    // Each input is clipped to its domain and encoded onto its axis of the
    // table; the fractional part is what the interpolation runs on.
    const base: Array<number> = [];
    const frac: Array<number> = [];
    for (let i = 0; i < m; i++) {
      const e = interpolate(
        clamp(inputs[i] ?? 0, domain[i * 2]!, domain[i * 2 + 1]!),
        domain[i * 2]!,
        domain[i * 2 + 1]!,
        encode[i * 2]!,
        encode[i * 2 + 1]!,
      );
      const at = clamp(e, 0, size[i]! - 1);
      const lo = Math.min(Math.floor(at), Math.max(0, size[i]! - 2));
      base.push(size[i] === 1 ? 0 : lo);
      frac.push(size[i] === 1 ? 0 : at - lo);
    }
    // Multilinear: every corner of the m-dimensional cell, weighted by how far
    // the point stands from it along each axis.
    const out = new Array<number>(n).fill(0);
    for (let corner = 0; corner < 1 << m; corner++) {
      let weight = 1;
      let flat = 0;
      let stride = 1;
      for (let i = 0; i < m; i++) {
        const up = (corner >> i) & 1;
        weight *= up ? frac[i]! : 1 - frac[i]!;
        flat += Math.min(base[i]! + up, size[i]! - 1) * stride;
        stride *= size[i]!;
      }
      if (weight === 0) continue;
      for (let j = 0; j < n; j++) out[j] = (out[j] ?? 0) + weight * sample(flat, j);
    }
    return out.map((v, j) =>
      clamp(
        interpolate(v, 0, 1, decode[j * 2]!, decode[j * 2 + 1]!),
        Math.min(range[j * 2]!, range[j * 2 + 1]!),
        Math.max(range[j * 2]!, range[j * 2 + 1]!),
      ),
    );
  };
}

/** §7.10.3 — an exponential interpolation between two colours. */
function exponential(file: PdfFile, dict: PdfDict, domain: Array<number>): PdfFunction | undefined {
  if (domain.length < 2) return undefined;
  const c0 = numbers(file, dict.get('C0'));
  const c1 = numbers(file, dict.get('C1'));
  const nExp = file.get(dict, 'N');
  if (typeof nExp !== 'number') return undefined;
  const from = c0.length > 0 ? c0 : [0];
  const to = c1.length > 0 ? c1 : [1];
  const n = Math.max(from.length, to.length);
  return (inputs) => {
    const x = clamp(inputs[0] ?? 0, domain[0]!, domain[1]!);
    // A non-integer exponent on a negative base is not a number; the domain of
    // a colour function does not reach there, and clamping is the safe reading.
    const t = nExp === 1 ? x : Math.max(0, x) ** nExp;
    const out: Array<number> = [];
    for (let i = 0; i < n; i++) out.push((from[i] ?? 0) + t * ((to[i] ?? 0) - (from[i] ?? 0)));
    return out;
  };
}

/** §7.10.4 — k functions laid end to end along one input. */
function stitching(
  file: PdfFile,
  dict: PdfDict,
  domain: Array<number>,
  depth: number,
): PdfFunction | undefined {
  if (domain.length < 2) return undefined;
  const list = file.resolve(dict.get('Functions') ?? PDF_NULL);
  if (!Array.isArray(list) || list.length === 0) return undefined;
  const parts = list.map((v) => readAt(file, v, depth + 1));
  if (parts.some((p) => p === undefined)) return undefined;
  const bounds = numbers(file, dict.get('Bounds'));
  const encode = orDefault(numbers(file, dict.get('Encode')), parts.length * 2, (i) => i % 2);
  const d0 = domain[0]!;
  const d1 = domain[1]!;
  return (inputs) => {
    const x = clamp(inputs[0] ?? 0, d0, d1);
    // The subdomain x falls in: the bounds are its interior edges.
    let k = 0;
    while (k < bounds.length && x >= bounds[k]!) k++;
    k = Math.min(k, parts.length - 1);
    const lo = k === 0 ? d0 : bounds[k - 1]!;
    const hi = k === bounds.length ? d1 : bounds[k]!;
    return parts[k]!([interpolate(x, lo, hi, encode[k * 2]!, encode[k * 2 + 1]!)]);
  };
}

/** §7.10.5 — a PostScript calculator: a program in braces, run on a stack. */
function calculator(
  file: PdfFile,
  stream: PdfStream,
  domain: Array<number>,
): PdfFunction | undefined {
  const range = numbers(file, stream.dict.get('Range'));
  const program = parseProgram(file.streamData(stream));
  if (!program || range.length < 2) return undefined;
  const n = range.length >> 1;
  return (inputs) => {
    const stack: Array<number> = [];
    for (let i = 0; i * 2 + 1 < domain.length; i++) {
      stack.push(clamp(inputs[i] ?? 0, domain[i * 2]!, domain[i * 2 + 1]!));
    }
    if (domain.length === 0) stack.push(...inputs);
    run(program, stack, 0);
    // §7.10.5 — the LAST n values on the stack are the outputs, in order.
    const out = stack.slice(-n);
    while (out.length < n) out.unshift(0);
    return out.map((v, j) =>
      clamp(
        v,
        Math.min(range[j * 2]!, range[j * 2 + 1]!),
        Math.max(range[j * 2]!, range[j * 2 + 1]!),
      ),
    );
  };
}

/** One token of a type-4 program: a number, an operator, or a nested block. */
type PsToken = number | string | Array<PsToken>;

/** Read `{ … }` into nested token lists, or `undefined` where it is not one. */
function parseProgram(bytes: Uint8Array): Array<PsToken> | undefined {
  const lexer = new Lexer(bytes);
  let count = 0;
  const read = (depth: number): Array<PsToken> | undefined => {
    if (depth > MAX_PS_DEPTH) return undefined;
    const out: Array<PsToken> = [];
    for (;;) {
      if (++count > MAX_PS_TOKENS) return undefined;
      const tok = lexer.nextToken();
      if (tok.kind === 'eof') return depth === 0 ? out : undefined;
      if (tok.kind === 'num') {
        out.push(tok.value);
      } else if (tok.kind === 'keyword') {
        if (tok.value === '{') {
          const inner = read(depth + 1);
          if (!inner) return undefined;
          out.push(inner);
        } else if (tok.value === '}') {
          return out;
        } else {
          out.push(tok.value);
        }
      }
    }
  };
  const top = read(0);
  // The whole program is itself one braced block; unwrap it.
  return top && top.length === 1 && Array.isArray(top[0]) ? top[0] : top;
}

/** Run a token list against the stack (§7.10.5, Table 42). */
function run(program: ReadonlyArray<PsToken>, stack: Array<number>, depth: number): void {
  if (depth > MAX_PS_DEPTH) return;
  const pop = (): number => stack.pop() ?? 0;
  // A block is pushed, not run: only `if` and `ifelse` run one, and they take
  // it from just behind them in the token list rather than off the stack.
  const blocks: Array<ReadonlyArray<PsToken>> = [];
  for (const tok of program) {
    if (typeof tok === 'number') {
      stack.push(tok);
      continue;
    }
    if (Array.isArray(tok)) {
      blocks.push(tok);
      continue;
    }
    switch (tok) {
      // Arithmetic.
      case 'abs':
        stack.push(Math.abs(pop()));
        break;
      case 'add': {
        const b = pop();
        stack.push(pop() + b);
        break;
      }
      case 'atan': {
        const den = pop();
        const numr = pop();
        // §7.10.5 — in DEGREES, and always positive.
        stack.push(((Math.atan2(numr, den) * 180) / Math.PI + 360) % 360);
        break;
      }
      case 'ceiling':
        stack.push(Math.ceil(pop()));
        break;
      case 'cos':
        stack.push(Math.cos((pop() * Math.PI) / 180));
        break;
      case 'cvi':
        stack.push(Math.trunc(pop()));
        break;
      case 'cvr':
        break; // already a real
      case 'div': {
        const b = pop();
        const a = pop();
        stack.push(b === 0 ? 0 : a / b);
        break;
      }
      case 'exp': {
        const b = pop();
        stack.push(pop() ** b);
        break;
      }
      case 'floor':
        stack.push(Math.floor(pop()));
        break;
      case 'idiv': {
        const b = pop();
        const a = pop();
        stack.push(b === 0 ? 0 : Math.trunc(a / b));
        break;
      }
      case 'ln': {
        const a = pop();
        stack.push(a > 0 ? Math.log(a) : 0);
        break;
      }
      case 'log': {
        const a = pop();
        stack.push(a > 0 ? Math.log10(a) : 0);
        break;
      }
      case 'mod': {
        const b = pop();
        const a = pop();
        stack.push(b === 0 ? 0 : Math.trunc(a) % Math.trunc(b));
        break;
      }
      case 'mul': {
        const b = pop();
        stack.push(pop() * b);
        break;
      }
      case 'neg':
        stack.push(-pop());
        break;
      case 'round':
        stack.push(Math.round(pop()));
        break;
      case 'sin':
        stack.push(Math.sin((pop() * Math.PI) / 180));
        break;
      case 'sqrt':
        stack.push(Math.sqrt(Math.max(0, pop())));
        break;
      case 'sub': {
        const b = pop();
        stack.push(pop() - b);
        break;
      }
      case 'truncate':
        stack.push(Math.trunc(pop()));
        break;
      // Relational, boolean and bitwise. A boolean rides the stack as 1 or 0,
      // and the bitwise operators are the same words as the boolean ones —
      // which is why `and` on two flags and `and` on two integers agree here.
      case 'and': {
        const b = pop();
        stack.push(int(pop()) & int(b));
        break;
      }
      case 'bitshift': {
        const s = int(pop());
        const a = int(pop());
        stack.push(s >= 0 ? a << s : a >> -s);
        break;
      }
      case 'eq': {
        const b = pop();
        stack.push(pop() === b ? 1 : 0);
        break;
      }
      case 'false':
        stack.push(0);
        break;
      case 'ge': {
        const b = pop();
        stack.push(pop() >= b ? 1 : 0);
        break;
      }
      case 'gt': {
        const b = pop();
        stack.push(pop() > b ? 1 : 0);
        break;
      }
      case 'le': {
        const b = pop();
        stack.push(pop() <= b ? 1 : 0);
        break;
      }
      case 'lt': {
        const b = pop();
        stack.push(pop() < b ? 1 : 0);
        break;
      }
      case 'ne': {
        const b = pop();
        stack.push(pop() !== b ? 1 : 0);
        break;
      }
      case 'not': {
        const a = pop();
        stack.push(a === 0 ? 1 : a === 1 ? 0 : ~int(a));
        break;
      }
      case 'or': {
        const b = pop();
        stack.push(int(pop()) | int(b));
        break;
      }
      case 'true':
        stack.push(1);
        break;
      case 'xor': {
        const b = pop();
        stack.push(int(pop()) ^ int(b));
        break;
      }
      // Conditional.
      case 'if': {
        const body = blocks.pop();
        if (body && pop() !== 0) run(body, stack, depth + 1);
        break;
      }
      case 'ifelse': {
        const otherwise = blocks.pop();
        const body = blocks.pop();
        const branch = pop() !== 0 ? body : otherwise;
        if (branch) run(branch, stack, depth + 1);
        break;
      }
      // Stack.
      case 'copy': {
        const k = int(pop());
        if (k > 0 && k <= stack.length) stack.push(...stack.slice(-k));
        break;
      }
      case 'dup': {
        const a = pop();
        stack.push(a, a);
        break;
      }
      case 'exch': {
        const b = pop();
        const a = pop();
        stack.push(b, a);
        break;
      }
      case 'index': {
        const k = int(pop());
        stack.push(k >= 0 && k < stack.length ? stack[stack.length - 1 - k]! : 0);
        break;
      }
      case 'pop':
        pop();
        break;
      case 'roll': {
        const j = int(pop());
        const k = int(pop());
        if (k > 0 && k <= stack.length) {
          const part = stack.splice(stack.length - k, k);
          const by = ((j % k) + k) % k;
          stack.push(...part.slice(k - by), ...part.slice(0, k - by));
        }
        break;
      }
      default:
        break; // an operator this does not know leaves the stack alone
    }
  }
}

/** A stack value as the 32-bit integer the bitwise operators work on. */
function int(v: number): number {
  return Number.isFinite(v) ? v | 0 : 0;
}

function clamp(v: number, lo: number, hi: number): number {
  if (!Number.isFinite(v)) return lo;
  return v < lo ? lo : v > hi ? hi : v;
}

/** §7.10.2 — one number carried from one interval onto another. */
function interpolate(x: number, xLo: number, xHi: number, yLo: number, yHi: number): number {
  if (xHi === xLo) return yLo;
  return yLo + ((x - xLo) * (yHi - yLo)) / (xHi - xLo);
}

function numbers(file: PdfFile, v: PdfValue | undefined): Array<number> {
  const r = v !== undefined ? file.resolve(v) : undefined;
  if (!Array.isArray(r)) return [];
  return r.map((x) => {
    const n = file.resolve(x);
    return typeof n === 'number' ? n : 0;
  });
}

/** The stated array, where it is the right length; otherwise the default. */
function orDefault(
  stated: Array<number>,
  length: number,
  fallback: (i: number) => number,
): Array<number> {
  if (stated.length === length) return stated;
  return Array.from({ length }, (_, i) => fallback(i));
}
