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
import type { ShapeGradient } from '@/core/vector';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import { PDF_NULL, PdfHexString, PdfName } from '@/pdf/objects';

/**
 * A page font as the interpreter needs it (built from the font dictionaries in
 * EP2b): how wide each code is, and how a run of codes decodes to Unicode. An
 * unmapped font falls back to Latin-1 with a half-em advance so text still
 * surfaces.
 */
export interface ContentFont {
  /** Bytes per character code: simple fonts read 1 byte/code, Type0 reads 2. */
  readonly bytesPerCode: 1 | 2;
  /** Decode a sequence of character codes to a Unicode string. */
  decode: (codes: ReadonlyArray<number>) => string;
  /** Glyph advance for one code, in 1000-unit text space. */
  width: (code: number) => number;
}

/**
 * One positioned text run emitted by a show operator: its decoded text, the
 * glyph origin in page space and the effective font size — the raw material a
 * later stage groups into lines and paragraphs.
 */
export interface TextRun {
  readonly text: string;
  /** Glyph origin x in page space (points). */
  readonly x: number;
  /** Glyph origin y in page space (points). */
  readonly y: number;
  readonly fontSizePt: number;
  readonly fontKey: string;
  /**
   * The marked-content id of the enclosing `BDC` sequence (§14.6), if any — the
   * link from this text to the structure element that owns it (E-PDF EP3).
   */
  readonly mcid?: number;
  /**
   * A `/Link` annotation whose `/Rect` covers this run's origin attaches its URI
   * here (E-PDF EP8), so the reconstructed run carries the hyperlink.
   */
  readonly href?: string;
}

/**
 * A painted XObject (`/Name Do`, §8.8) — an image or form. The CTM maps the unit
 * square to page space, so it carries both the placement and the size; the
 * `mcid` links the paint to its structure element (a `/Figure`, E-PDF EP6).
 */
export interface ImagePlacement {
  /** XObject resource name (no leading slash). */
  readonly name: string;
  readonly ctm: Matrix;
  readonly mcid?: number;
}

/**
 * One segment of a painted path (E-PDF EP10/EP11), in page space (y-up): a
 * `move`/`line`/`cubic` Bézier point or a subpath `close`.
 */
export type PathSeg =
  | { readonly op: 'move'; readonly x: number; readonly y: number }
  | { readonly op: 'line'; readonly x: number; readonly y: number }
  | {
      readonly op: 'cubic';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly op: 'close' };

/**
 * A painted path emitted by the path-painting operators (§8.5.3), captured in
 * page space (y-up). The optional fields record only what the paint mode set:
 * a fill colour/gradient (EP10/EP16c) and/or a stroke colour + width (EP11),
 * plus the enclosing structure id.
 */
export interface VectorPlacement {
  readonly segs: ReadonlyArray<PathSeg>;
  /** Fill colour (6-hex), present iff the path is filled (`f` / `F` / `f*` / `B` / `b`). */
  readonly fillHex?: string;
  /** Shading pattern, present iff filled with one (EP16c). */
  readonly gradient?: ShapeGradient;
  /** Stroke colour (6-hex), present iff the path is stroked (`S` / `s` / `B` / `b`) — EP11. */
  readonly strokeHex?: string;
  /** Stroke width in page-space points — EP11. */
  readonly lineWidth?: number;
  readonly mcid?: number;
}

/** Everything {@link interpretContent} extracts from one page's content stream. */
export interface InterpretResult {
  readonly texts: Array<TextRun>;
  readonly images: Array<ImagePlacement>;
  readonly vectors: Array<VectorPlacement>;
}

/**
 * 2D affine matrix `[a b c d e f]`, row-vector convention (`[x y 1] · M`):
 * ```
 * ⎡a b 0⎤
 * ⎢c d 0⎥
 * ⎣e f 1⎦
 * ```
 */
export type Matrix = readonly [number, number, number, number, number, number];
/** The identity {@link Matrix} (no transform). */
export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** Compose two {@link Matrix matrices}: `a` applied first, then `b`. */
export function multiply(a: Matrix, b: Matrix): Matrix {
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
  fillColor: string; // current non-stroking colour (6-hex), graphics state (EP10)
  strokeColor: string; // current stroking colour (6-hex), graphics state (EP11)
  lineWidth: number; // current line width in user-space units (EP11)
  fillGradient: ShapeGradient | undefined; // current non-stroking shading pattern (EP16c)
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
    fillColor: '000000',
    strokeColor: '000000',
    lineWidth: 1, // §8.4.3.2 default line width
    fillGradient: undefined,
  };
}

/**
 * Walk a page's decoded content stream (§9.4) and extract its positioned text,
 * painted XObjects and painted paths. Tracks the graphics state (CTM via
 * `q`/`Q`/`cm`) and text state (text/line matrices, font, size, spacing), and
 * emits one {@link TextRun} per show operator (`Tj` / `TJ` / `'` / `"`) in page
 * space (points).
 *
 * @param bytes      The page's decoded content-stream bytes.
 * @param fonts      Page fonts by resource key (`Tf` name), for decoding + widths;
 *                   an unmapped key falls back to Latin-1 with a half-em advance.
 * @param initialCtm The starting CTM mapping user space to page space.
 * @param shadings   Shading patterns by name, selected by `scn`/`sc` (EP16c).
 * @returns The extracted text runs, image placements and vector paths.
 */
export function interpretContent(
  bytes: Uint8Array,
  fonts: ReadonlyMap<string, ContentFont>,
  initialCtm: Matrix = IDENTITY,
  shadings: ReadonlyMap<string, ShapeGradient> = new Map(),
): InterpretResult {
  const runs: Array<TextRun> = [];
  const images: Array<ImagePlacement> = [];
  const vectors: Array<VectorPlacement> = []; // filled paths (EP10)
  const lexer = new Lexer(bytes);
  const stack: Array<TextState> = [];
  let state = initialState();
  state.ctm = initialCtm;
  let tm: Matrix = IDENTITY; // text matrix
  let tlm: Matrix = IDENTITY; // line matrix
  let path: Array<PathSeg> = []; // the current path under construction (page space)
  let operands: Array<PdfValue> = [];
  const mcStack: Array<number | undefined> = []; // marked-content (MCID) nesting

  // Apply the CTM to a user-space point → page space (§8.3.4).
  const toPage = (x: number, y: number): [number, number] => [
    x * state.ctm[0] + y * state.ctm[2] + state.ctm[4],
    x * state.ctm[1] + y * state.ctm[3] + state.ctm[5],
  ];
  const moveTo = (x: number, y: number): void => {
    const [px, py] = toPage(x, y);
    path.push({ op: 'move', x: px, y: py });
  };
  const lineTo = (x: number, y: number): void => {
    const [px, py] = toPage(x, y);
    path.push({ op: 'line', x: px, y: py });
  };
  const curveTo = (a: number, b: number, c: number, d: number, e: number, f: number): void => {
    const [x1, y1] = toPage(a, b);
    const [x2, y2] = toPage(c, d);
    const [x, y] = toPage(e, f);
    path.push({ op: 'cubic', x1, y1, x2, y2, x, y });
  };
  const rectTo = (x: number, y: number, w: number, h: number): void => {
    moveTo(x, y);
    lineTo(x + w, y);
    lineTo(x + w, y + h);
    lineTo(x, y + h);
    path.push({ op: 'close' });
  };
  // The line width in page space — the user-space width scaled by the CTM
  // (§8.4.3.2); a uniform scale is the geometric mean √|det| of the matrix.
  const ctmLineWidth = (): number => {
    const m = state.ctm;
    const scale = Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
    return state.lineWidth * scale;
  };
  // Emit the current path as a painted vector (§8.5.3): filled, stroked, or both.
  // `n` and clip operators paint nothing — they pass fill=stroke=false to clear.
  const paintPath = (fill: boolean, stroke: boolean): void => {
    if (path.length >= 2 && (fill || stroke)) {
      const mcid = mcStack.length > 0 ? mcStack[mcStack.length - 1] : undefined;
      vectors.push({
        segs: path,
        ...(fill ? { fillHex: state.fillColor } : {}),
        ...(fill && state.fillGradient ? { gradient: state.fillGradient } : {}),
        ...(stroke ? { strokeHex: state.strokeColor, lineWidth: ctmLineWidth() } : {}),
        ...(mcid !== undefined ? { mcid } : {}),
      });
    }
    path = [];
  };

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
    const mcid = mcStack.length > 0 ? mcStack[mcStack.length - 1] : undefined;
    runs.push({
      text,
      x: origin[4],
      y: origin[5],
      fontSizePt: state.fontSize * scaleY,
      fontKey: state.fontKey,
      ...(mcid !== undefined ? { mcid } : {}),
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
      case 'BDC': {
        // `tag props BDC` — a structure content sequence (or an artifact). Push
        // its /MCID so the runs inside it are tagged; an /Artifact has none.
        const tag = operands[operands.length - 2];
        const props = operands[operands.length - 1];
        const isArtifact = tag instanceof PdfName && tag.value === 'Artifact';
        const mcidVal = !isArtifact && props instanceof Map ? props.get('MCID') : undefined;
        mcStack.push(typeof mcidVal === 'number' ? mcidVal : undefined);
        break;
      }
      case 'BMC':
        mcStack.push(undefined); // `tag BMC` — no properties, so no MCID
        break;
      case 'EMC':
        mcStack.pop();
        break;
      case 'Do': {
        // Paint an XObject (image or form). Record its name + the CTM (which
        // already folds in the placement `cm`) so a later stage can resolve and
        // size it; tag it with the enclosing structure id (a /Figure).
        const nm = operands[0];
        if (nm instanceof PdfName) {
          const mcid = mcStack.length > 0 ? mcStack[mcStack.length - 1] : undefined;
          images.push({ name: nm.value, ctm: state.ctm, ...(mcid !== undefined ? { mcid } : {}) });
        }
        break;
      }
      // §8.6.8 non-stroking colour → the current fill colour (EP10); a solid
      // colour clears any active shading pattern.
      case 'rg':
        state.fillColor = rgbHex(num(0), num(1), num(2));
        state.fillGradient = undefined;
        break;
      case 'g':
        state.fillColor = grayHex(num(0));
        state.fillGradient = undefined;
        break;
      case 'k':
        state.fillColor = cmykHex(num(0), num(1), num(2), num(3));
        state.fillGradient = undefined;
        break;
      // §8.6.8 colour in a named space; a /Pattern name selects a shading
      // pattern (EP16c), numeric operands are a solid colour we leave as-is.
      case 'scn':
      case 'sc': {
        const last = operands[operands.length - 1];
        state.fillGradient = last instanceof PdfName ? shadings.get(last.value) : undefined;
        break;
      }
      // §8.6.8 stroking colour → the current stroke colour (EP11).
      case 'RG':
        state.strokeColor = rgbHex(num(0), num(1), num(2));
        break;
      case 'G':
        state.strokeColor = grayHex(num(0));
        break;
      case 'K':
        state.strokeColor = cmykHex(num(0), num(1), num(2), num(3));
        break;
      case 'w':
        state.lineWidth = num(0); // §8.4.3.2 line width (user space)
        break;
      // §8.5.2 path construction.
      case 'm':
        moveTo(num(0), num(1));
        break;
      case 'l':
        lineTo(num(0), num(1));
        break;
      case 'c':
        curveTo(num(0), num(1), num(2), num(3), num(4), num(5));
        break;
      case 're':
        rectTo(num(0), num(1), num(2), num(3));
        break;
      case 'h':
        path.push({ op: 'close' });
        break;
      // §8.5.3 path painting — capture FILLS (EP10) and STROKES (EP11). Clips
      // (W/W*) only mark the region; the following painting operator emits.
      case 'f':
      case 'F':
      case 'f*':
        paintPath(true, false);
        break;
      case 'S':
        paintPath(false, true);
        break;
      case 's':
        path.push({ op: 'close' });
        paintPath(false, true);
        break;
      case 'B':
      case 'B*':
        paintPath(true, true);
        break;
      case 'b':
      case 'b*':
        path.push({ op: 'close' });
        paintPath(true, true);
        break;
      case 'n':
        paintPath(false, false); // end the path with no paint
        break;
      case 'W':
      case 'W*':
        break; // intersect clip; the painting operator that follows clears the path
      default:
        break; // other graphics-state / stroking operators ignored
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
        operands.push(readDict(lexer)); // e.g. a BDC marked-content property dict
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
  return { texts: runs, images, vectors };
}

// PDF colour operands (0..1 per channel) → a 6-hex sRGB string.
function clamp255(v: number): number {
  return Math.max(0, Math.min(255, Math.round(v * 255)));
}
function hex2(v: number): string {
  return clamp255(v).toString(16).padStart(2, '0');
}
function rgbHex(r: number, g: number, b: number): string {
  return (hex2(r) + hex2(g) + hex2(b)).toUpperCase();
}
function grayHex(v: number): string {
  return rgbHex(v, v, v);
}
function cmykHex(c: number, m: number, y: number, k: number): string {
  return rgbHex((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k));
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

// Read a content-stream dictionary operand (assumes `<<` already consumed),
// capturing name → value pairs up to `>>` — needed for a BDC /MCID property.
function readDict(lexer: Lexer): PdfDict {
  const map: PdfDict = new Map<string, PdfValue>();
  for (;;) {
    const key = lexer.nextToken();
    if (key.kind === 'dictClose' || key.kind === 'eof') break;
    if (key.kind !== 'name') continue;
    map.set(key.value, readValue(lexer));
  }
  return map;
}

function readValue(lexer: Lexer): PdfValue {
  const tok = lexer.nextToken();
  switch (tok.kind) {
    case 'num':
      return tok.value;
    case 'name':
      return new PdfName(tok.value);
    case 'str':
      return tok.value;
    case 'hexstr':
      return new PdfHexString(tok.bytes);
    case 'arrayOpen':
      return readArray(lexer);
    case 'dictOpen':
      return readDict(lexer);
    case 'keyword':
      return tok.value === 'true' ? true : tok.value === 'false' ? false : PDF_NULL;
    default:
      return PDF_NULL;
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
