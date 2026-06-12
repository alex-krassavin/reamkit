// E-PDF EP2 — content-stream interpreter (ISO 32000-1 §9.4). Walks a page's
// (decoded) content stream, tracking the graphics state (CTM via q/Q/cm) and the
// text object state (the text + line matrices, font, size and spacing), and emits
// one positioned text run per show operator (Tj / TJ / ' / "). Each run carries
// the glyph-origin in page space (user space, points) and the effective font
// size, the raw material a later stage groups into lines and paragraphs.
//
// Decoding shown bytes to Unicode and the per-glyph advance widths come from the
// supplied ContentFont (built from the font dictionaries in EP2b); an unmapped
// font falls back to Latin-1 with a half-em advance so text still surfaces.

import { Lexer } from './lexer';
import type { PdfValue } from '@/pdf/objects';
import { PdfHexString, PdfName } from '@/pdf/objects';

export interface ContentFont {
  readonly bytesPerCode: 1 | 2; // simple fonts read 1 byte/code, Type0 reads 2
  decode: (codes: ReadonlyArray<number>) => string;
  width: (code: number) => number; // glyph advance in 1000-unit text space
}

export interface TextRun {
  readonly text: string;
  readonly x: number; // glyph-origin in page space (points)
  readonly y: number;
  readonly fontSizePt: number;
  readonly fontKey: string;
}

// 2D affine matrix [a b c d e f] = ⎡a b 0⎤ row-vector convention ([x y 1] · M).
//                                  ⎢c d 0⎥
//                                  ⎣e f 1⎦
type Matrix = readonly [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

// A applied first, then B.
function multiply(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[1] * b[2],
    a[0] * b[1] + a[1] * b[3],
    a[2] * b[0] + a[3] * b[2],
    a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4],
    a[4] * b[1] + a[5] * b[3] + b[5],
  ];
}

function translation(tx: number, ty: number): Matrix {
  return [1, 0, 0, 1, tx, ty];
}

const FALLBACK_FONT: ContentFont = {
  bytesPerCode: 1,
  decode: (codes) => codes.map((c) => String.fromCharCode(c)).join(''),
  width: () => 500,
};

interface TextState {
  ctm: Matrix;
  fontKey: string;
  font: ContentFont;
  fontSize: number;
  charSpacing: number; // Tc, text-space units
  wordSpacing: number; // Tw
  hScale: number; // Tz / 100
  leading: number; // TL
  rise: number; // Ts
}

function initialState(): TextState {
  return {
    ctm: IDENTITY,
    fontKey: '',
    font: FALLBACK_FONT,
    fontSize: 0,
    charSpacing: 0,
    wordSpacing: 0,
    hScale: 1,
    leading: 0,
    rise: 0,
  };
}

export function interpretContent(
  bytes: Uint8Array,
  fonts: ReadonlyMap<string, ContentFont>,
): Array<TextRun> {
  const runs: Array<TextRun> = [];
  const lexer = new Lexer(bytes);
  const stack: Array<TextState> = [];
  let state = initialState();
  let tm: Matrix = IDENTITY; // text matrix
  let tlm: Matrix = IDENTITY; // line matrix
  let operands: Array<PdfValue> = [];

  const num = (i: number): number => {
    const v = operands[i];
    return typeof v === 'number' ? v : 0;
  };

  // Advance the text matrix for one shown glyph (§9.4.4): w0·Tfs + Tc (+ Tw for
  // the single-byte space), all scaled horizontally by Th.
  const advanceGlyph = (code: number): void => {
    const w0 = state.font.width(code) / 1000;
    const isSpace = state.font.bytesPerCode === 1 && code === 0x20;
    const tx =
      (w0 * state.fontSize + state.charSpacing + (isSpace ? state.wordSpacing : 0)) * state.hScale;
    tm = multiply(translation(tx, 0), tm);
  };

  // Decode a shown string and advance the matrix glyph by glyph, returning its
  // Unicode (without emitting — Tj and TJ both build on this).
  const consume = (operand: PdfValue): string => {
    const codes = splitCodes(toBytes(operand), state.font.bytesPerCode);
    for (const code of codes) advanceGlyph(code);
    return state.font.decode(codes);
  };

  const emitAt = (origin: Matrix, text: string): void => {
    if (text.length === 0) return;
    const scaleY = Math.hypot(origin[2], origin[3]) || 1;
    runs.push({
      text,
      x: origin[4],
      y: origin[5],
      fontSizePt: state.fontSize * scaleY,
      fontKey: state.fontKey,
    });
  };

  // Tj / ' / " — one string at the current origin.
  const showString = (operand: PdfValue): void => {
    const origin = multiply(tm, state.ctm);
    emitAt(origin, consume(operand));
  };

  // TJ — an array of strings and kerning adjustments; one run at the start origin.
  const showArray = (arr: ReadonlyArray<PdfValue>): void => {
    const origin = multiply(tm, state.ctm);
    let text = '';
    for (const el of arr) {
      if (typeof el === 'number') {
        tm = multiply(translation((-el / 1000) * state.fontSize * state.hScale, 0), tm);
      } else if (typeof el === 'string' || el instanceof PdfHexString) {
        text += consume(el);
      }
    }
    emitAt(origin, text);
  };

  const exec = (op: string): void => {
    switch (op) {
      case 'q':
        stack.push({ ...state });
        break;
      case 'Q':
        state = stack.pop() ?? state;
        break;
      case 'cm':
        state.ctm = multiply(matrixFromOperands(operands), state.ctm);
        break;
      case 'BT':
        tm = IDENTITY;
        tlm = IDENTITY;
        break;
      case 'ET':
        break;
      case 'Tf': {
        const key = operands[0] instanceof PdfName ? operands[0].value : '';
        state.fontKey = key;
        state.font = fonts.get(key) ?? FALLBACK_FONT;
        state.fontSize = num(1);
        break;
      }
      case 'Td':
        tlm = multiply(translation(num(0), num(1)), tlm);
        tm = tlm;
        break;
      case 'TD':
        state.leading = -num(1);
        tlm = multiply(translation(num(0), num(1)), tlm);
        tm = tlm;
        break;
      case 'Tm':
        tlm = matrixFromOperands(operands);
        tm = tlm;
        break;
      case 'T*':
        tlm = multiply(translation(0, -state.leading), tlm);
        tm = tlm;
        break;
      case 'TL':
        state.leading = num(0);
        break;
      case 'Tc':
        state.charSpacing = num(0);
        break;
      case 'Tw':
        state.wordSpacing = num(0);
        break;
      case 'Tz':
        state.hScale = num(0) / 100;
        break;
      case 'Ts':
        state.rise = num(0);
        break;
      case 'Tj':
        if (operands.length > 0) showString(operands[operands.length - 1]!);
        break;
      case 'TJ':
        if (Array.isArray(operands[0])) showArray(operands[0]);
        break;
      case "'":
        tlm = multiply(translation(0, -state.leading), tlm);
        tm = tlm;
        if (operands.length > 0) showString(operands[operands.length - 1]!);
        break;
      case '"':
        state.wordSpacing = num(0);
        state.charSpacing = num(1);
        tlm = multiply(translation(0, -state.leading), tlm);
        tm = tlm;
        if (operands.length > 2) showString(operands[2]!);
        break;
      default:
        break; // path, colour, XObject, marked-content … ignored for text
    }
  };

  for (;;) {
    lexer.skipWhitespace();
    const tok = lexer.nextToken();
    if (tok.kind === 'eof') break;
    switch (tok.kind) {
      case 'num':
        operands.push(tok.value);
        break;
      case 'name':
        operands.push(new PdfName(tok.value));
        break;
      case 'str':
        operands.push(tok.value);
        break;
      case 'hexstr':
        operands.push(new PdfHexString(tok.bytes));
        break;
      case 'arrayOpen':
        operands.push(readArray(lexer));
        break;
      case 'dictOpen':
        skipDict(lexer);
        operands.push([]); // a marked-content property dict — value ignored
        break;
      case 'keyword':
        if (tok.value === 'BI') skipInlineImage(lexer);
        else exec(tok.value);
        operands = [];
        break;
      default:
        operands = []; // stray ] or >> — reset
        break;
    }
  }
  return runs;
}

function matrixFromOperands(operands: ReadonlyArray<PdfValue>): Matrix {
  const n = (i: number): number => (typeof operands[i] === 'number' ? operands[i] : 0);
  return [n(0), n(1), n(2), n(3), n(4), n(5)];
}

function toBytes(operand: PdfValue): Uint8Array {
  if (operand instanceof PdfHexString) return operand.bytes;
  if (typeof operand === 'string') {
    const out = new Uint8Array(operand.length);
    for (let i = 0; i < operand.length; i++) out[i] = operand.charCodeAt(i) & 0xff;
    return out;
  }
  return new Uint8Array(0);
}

function splitCodes(bytes: Uint8Array, bytesPerCode: 1 | 2): Array<number> {
  const out: Array<number> = [];
  if (bytesPerCode === 2) {
    for (let i = 0; i + 1 < bytes.length; i += 2) out.push((bytes[i]! << 8) | bytes[i + 1]!);
    if (bytes.length % 2 === 1) out.push(bytes[bytes.length - 1]!);
  } else {
    for (const b of bytes) out.push(b);
  }
  return out;
}

// Read a content-stream array operand (TJ): numbers and strings up to `]`.
function readArray(lexer: Lexer): Array<PdfValue> {
  const out: Array<PdfValue> = [];
  for (;;) {
    const tok = lexer.nextToken();
    if (tok.kind === 'arrayClose' || tok.kind === 'eof') break;
    if (tok.kind === 'num') out.push(tok.value);
    else if (tok.kind === 'str') out.push(tok.value);
    else if (tok.kind === 'hexstr') out.push(new PdfHexString(tok.bytes));
  }
  return out;
}

function skipDict(lexer: Lexer): void {
  let depth = 1;
  for (;;) {
    const tok = lexer.nextToken();
    if (tok.kind === 'eof') break;
    if (tok.kind === 'dictOpen') depth++;
    else if (tok.kind === 'dictClose' && --depth === 0) break;
  }
}

// §8.9.7 — an inline image: skip from BI past the binary data to EI.
function skipInlineImage(lexer: Lexer): void {
  for (;;) {
    const tok = lexer.nextToken();
    if (tok.kind === 'eof') return;
    if (tok.kind === 'keyword' && tok.value === 'ID') break;
  }
  const ei = lexer.indexOfAscii('EI', lexer.pos);
  lexer.pos = ei < 0 ? lexer.length : ei + 2;
}
