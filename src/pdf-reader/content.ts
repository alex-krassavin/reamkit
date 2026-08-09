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
  /** §9.6.2 — the face's own `/BaseFont` name, for a document that embeds it. */
  readonly name?: string;
  /** §9.8.1 — the face is a bold one (weight, the ForceBold flag, or its name). */
  readonly bold?: boolean;
  /** §9.8.1 — the face is slanted (`/ItalicAngle`, the Italic flag, or its name). */
  readonly italic?: boolean;
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
  /**
   * Where the pen stood after the last glyph, in page space (§9.4.4). The
   * interpreter advances the text matrix by the font's own widths, so this is
   * a measurement, not the half-em-per-character guess the reader used to make
   * — and the difference between a word space and a table column is exactly
   * the kind of thing a guess gets wrong.
   */
  readonly endX: number;
  /** The pen's y after the last glyph — with {@link endX}, the whole advance. */
  readonly endY: number;
  /**
   * The baseline's direction in page space, degrees counter-clockwise from
   * left-to-right (§9.4.2 — the text matrix may turn as well as move). Absent
   * for ordinary upright text, which is nearly all of it.
   */
  readonly angleDeg?: number;
  readonly fontSizePt: number;
  readonly fontKey: string;
  /**
   * §9.6.2 `/BaseFont` — the face's own name, subset prefix dropped and
   * lowercased, or absent when the font states none. This is what a rebuilt run
   * asks for, so a page whose faces the file EMBEDS is re-set in them rather
   * than in a substitute (see `./embedded-fonts`).
   */
  readonly fontName?: string;
  /** §9.8.1 — the face the glyphs were shown in is a bold one. */
  readonly bold?: boolean;
  /** §9.8.1 — the face the glyphs were shown in is a slanted one. */
  readonly italic?: boolean;
  /** §8.6.8 — the non-stroking colour the glyphs were painted in (6-hex). */
  readonly colorHex: string;
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
  /** Where the `Do` fell in the stream's painting order — see {@link VectorPlacement.order}. */
  readonly order: number;
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
/** A path's page-space bounding box, or `undefined` when it names no point. */
function pathBox(
  segs: ReadonlyArray<PathSeg>,
): { minX: number; minY: number; maxX: number; maxY: number } | undefined {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  const add = (x: number, y: number): void => {
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x);
    maxY = Math.max(maxY, y);
  };
  for (const seg of segs) {
    if (seg.op === 'move' || seg.op === 'line') add(seg.x, seg.y);
    else if (seg.op === 'cubic') {
      add(seg.x1, seg.y1);
      add(seg.x2, seg.y2);
      add(seg.x, seg.y);
    }
  }
  return Number.isFinite(minX) ? { minX, minY, maxX, maxY } : undefined;
}

/** The area of a box, for choosing the smaller of two clip regions. */
const area = (b: { minX: number; minY: number; maxX: number; maxY: number }): number =>
  Math.max(0, b.maxX - b.minX) * Math.max(0, b.maxY - b.minY);

/**
 * §8.5.4 — the clipping region in force when a path was painted: the path that
 * `W`/`W*` installed, plus its page-space bounding box.
 */
export interface ClipRegion {
  readonly segs: ReadonlyArray<PathSeg>;
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
}

export interface VectorPlacement {
  /**
   * Where this fell in the stream's painting order (§8.5.3): later covers
   * earlier, and a `Do` of a form is numbered here too, so a caller walking
   * into that form knows exactly where its marks belong among these.
   */
  readonly order: number;
  readonly segs: ReadonlyArray<PathSeg>;
  /** §8.5.4 — the clip in force when it was painted, when there was one. */
  readonly clip?: ClipRegion;
  /** Fill colour (6-hex), present iff the path is filled (`f` / `F` / `f*` / `B` / `b`). */
  readonly fillHex?: string;
  /** Shading pattern, present iff filled with one (EP16c). */
  readonly gradient?: ShapeGradient;
  /** §11.6.4.4 `/ca` — how opaque the fill is, when the page asked for less. */
  readonly alpha?: number;
  /**
   * §8.7.3 — the TILING pattern resource name the path is filled with. Its
   * content is a stream of its own, so what the fill actually shows is only
   * known by walking into it; the `fillHex` beside this is not the fill.
   */
  readonly patternName?: string;
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

/** Below this a baseline is upright: a rounded matrix is not a turned one. */
const UPRIGHT_TOLERANCE_DEG = 0.5;

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
  fillPattern: string | undefined; // §8.7.3 non-stroking TILING pattern resource name
  fillAlpha: number; // §11.6.4.4 `/ca` — how opaque the non-stroking paint is
  clip: ClipRegion | undefined; // §8.5.4 the clipping region in force
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
    fillPattern: undefined,
    fillAlpha: 1,
    clip: undefined,
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
 * @param alphas     Constant fill alphas by `/ExtGState` name, selected by `gs`.
 * @returns The extracted text runs, image placements and vector paths.
 */
export function interpretContent(
  bytes: Uint8Array,
  fonts: ReadonlyMap<string, ContentFont>,
  initialCtm: Matrix = IDENTITY,
  shadings: ReadonlyMap<string, ShapeGradient> = new Map(),
  alphas: ReadonlyMap<string, number> = new Map(),
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
  let pendingClip = false; // §8.5.4 `W` seen; the next painting operator installs it
  let paintOrder = 0; // §8.5.3 the sequence marks are laid down in
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
        order: paintOrder++,
        segs: path,
        ...(state.clip ? { clip: state.clip } : {}),
        ...(fill && state.fillPattern !== undefined ? { patternName: state.fillPattern } : {}),
        ...(fill ? { fillHex: state.fillColor } : {}),
        ...(fill && state.fillAlpha < 1 ? { alpha: state.fillAlpha } : {}),
        ...(fill && state.fillGradient ? { gradient: state.fillGradient } : {}),
        ...(stroke ? { strokeHex: state.strokeColor, lineWidth: ctmLineWidth() } : {}),
        ...(mcid !== undefined ? { mcid } : {}),
      });
    }
    // §8.5.4 — `W` names the clip but does not install it: the painting
    // operator that ENDS the path does, and the path it names is this one.
    if (pendingClip) {
      pendingClip = false;
      const box = pathBox(path);
      if (box) {
        const next: ClipRegion = { segs: path, ...box };
        // Clips intersect. Nested ones nest, so the smaller region stands for
        // the intersection — the whole of it, where it is the whole.
        state.clip = state.clip && area(state.clip) <= area(next) ? state.clip : next;
      }
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

  const emitAt = (origin: Matrix, text: string, end: Matrix): void => {
    if (text.length === 0) return;
    const scaleY = Math.hypot(origin[2], origin[3]) || 1;
    const mcid = mcStack.length > 0 ? mcStack[mcStack.length - 1] : undefined;
    // §9.4.2 — the text matrix turns as well as moves. The baseline's own
    // direction is the first column of it; upright text leaves this at zero.
    const angle = (Math.atan2(origin[1], origin[0]) * 180) / Math.PI;
    runs.push({
      text,
      x: origin[4],
      y: origin[5],
      endX: end[4],
      endY: end[5],
      ...(Math.abs(angle) > UPRIGHT_TOLERANCE_DEG ? { angleDeg: angle } : {}),
      fontSizePt: state.fontSize * scaleY,
      fontKey: state.fontKey,
      ...(state.font.name !== undefined ? { fontName: state.font.name } : {}),
      ...(state.font.bold ? { bold: true } : {}),
      ...(state.font.italic ? { italic: true } : {}),
      colorHex: state.fillColor,
      ...(mcid !== undefined ? { mcid } : {}),
    });
  };

  // Tj / ' / " — one string at the current origin.
  const showString = (operand: PdfValue): void => {
    const origin = multiply(tm, state.ctm);
    const text = consume(operand);
    emitAt(origin, text, multiply(tm, state.ctm));
  };

  // §9.4.3 TJ — an array of strings with the pen nudged BETWEEN them. Each
  // string is emitted where it stands, because a nudge is the page saying that
  // this piece does not go where the font's own widths would put it: read as
  // one run from the start origin, every adjustment was thrown away and the
  // error accumulated along the line. 160F-2019.pdf kerns its labels that way,
  // and "période du" drifted a point and a half by its last letter.
  const showArray = (arr: ReadonlyArray<PdfValue>): void => {
    for (const el of arr) {
      if (typeof el === 'number') {
        tm = multiply(translation((-el / 1000) * state.fontSize * state.hScale, 0), tm);
      } else if (typeof el === 'string' || el instanceof PdfHexString) {
        const origin = multiply(tm, state.ctm);
        const text = consume(el);
        emitAt(origin, text, multiply(tm, state.ctm));
      }
    }
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
          images.push({
            order: paintOrder++,
            name: nm.value,
            ctm: state.ctm,
            ...(mcid !== undefined ? { mcid } : {}),
          });
        }
        break;
      }
      // §8.6.8 non-stroking colour → the current fill colour (EP10); a solid
      // colour clears any active shading pattern.
      case 'rg':
        state.fillColor = rgbHex(num(0), num(1), num(2));
        state.fillGradient = undefined;
        state.fillPattern = undefined;
        break;
      case 'g':
        state.fillColor = grayHex(num(0));
        state.fillGradient = undefined;
        state.fillPattern = undefined;
        break;
      case 'k':
        state.fillColor = cmykHex(num(0), num(1), num(2), num(3));
        state.fillGradient = undefined;
        state.fillPattern = undefined;
        break;
      // §8.6.8 colour in a named space; a /Pattern name selects a shading
      // pattern (EP16c), numeric operands are a solid colour we leave as-is.
      case 'scn':
      case 'sc': {
        // §8.6.6.2 — in a Pattern colour space the operand is a pattern NAME,
        // not components. A shading pattern (type 2) resolves to a gradient
        // here; a tiling one (type 1) is a whole content stream and is named
        // for the image walk to enter. Neither is a colour, and taking the
        // fill colour still standing from before painted 22060_A1_01_Plans.pdf's
        // four floor plans as four black rectangles.
        const last = operands[operands.length - 1];
        const named = last instanceof PdfName ? last.value : undefined;
        state.fillGradient = named !== undefined ? shadings.get(named) : undefined;
        state.fillPattern = named !== undefined && !state.fillGradient ? named : undefined;
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
      case 'gs': {
        // §8.4.5 — a named graphics state, of which only the constant fill
        // alpha is read here. A band meant to be seen through is not the same
        // mark as one that hides what it covers.
        const nm = operands[operands.length - 1];
        if (nm instanceof PdfName) state.fillAlpha = alphas.get(nm.value) ?? 1;
        break;
      }
      case 'W':
      case 'W*':
        // §8.5.4 — the current path becomes the clip once the painting
        // operator that follows ends it.
        pendingClip = true;
        break;
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
