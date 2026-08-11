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
import { cieToSrgb } from './cie-color';
import type { ColorSpaceInfo, GsPaint } from './shading';
import type { TextMarkup } from './annot-draw';
import type { ShapeGradient } from '@/core/vector';
import type { PdfDict, PdfStream, PdfValue } from '@/pdf/objects';
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
  /**
   * §9.7.6.2 — how a shown string breaks into codes, where the CMap is not
   * fixed-width. A named CMap like `90ms-RKSJ-H` mixes one-byte and two-byte
   * codes, and split down the middle a Japanese line came apart into nonsense.
   * Absent means the fixed width above.
   */
  readonly splitCodes?: (bytes: Uint8Array) => Array<number>;
  /** Decode a sequence of character codes to a Unicode string. */
  decode: (codes: ReadonlyArray<number>) => string;
  /** Glyph advance for one code, in 1000-unit text space. */
  width: (code: number) => number;
  /**
   * §9.4.4 / §9.7.4.3 — the face sets its text DOWN the page, not across, and
   * the pen advances by the vertical displacement `w1` rather than by `w0`.
   * A `…-V` CMap asks for this; `/DW2`'s default `[880 -1000]` is one em down.
   * The number here is that displacement in 1000-unit text space, and it is
   * negative because the page's y runs up.
   */
  readonly verticalAdvance?: (code: number) => number;
  /** §9.6.2 — the face's own `/BaseFont` name, for a document that embeds it. */
  readonly name?: string;
  /** §9.8.1 — the face is a bold one (weight, the ForceBold flag, or its name). */
  readonly bold?: boolean;
  /** §9.8.1 — the face is slanted (`/ItalicAngle`, the Italic flag, or its name). */
  readonly italic?: boolean;
  /**
   * §9.6.5 — a Type 3 face, whose glyphs are content streams rather than
   * outlines. What such a font draws is not type at all: it is whatever the
   * procedure paints, in the resources the font states.
   */
  readonly type3?: Type3Face;
}

/** §9.6.5 — the parts of a Type 3 font a caller needs to run its glyphs. */
export interface Type3Face {
  /** `/FontMatrix` — glyph space to text space. */
  readonly matrix: Matrix;
  /** `/Encoding` + `/CharProcs` — the content stream one code draws. */
  readonly proc: (code: number) => PdfStream | undefined;
  /** `/Resources` the procedures draw with, when the font states its own. */
  readonly resources: PdfDict | undefined;
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
  /**
   * §9.4.4 — the advance of this face's SPACE at the size the run was shown
   * at, which is what says whether a gap between two runs was a word space.
   * Absent where the face states no width for it.
   */
  readonly spaceWidthPt?: number;
  /** §9.8.1 — the face the glyphs were shown in is a bold one. */
  readonly bold?: boolean;
  /**
   * §12.5.6.10 — a text-markup annotation marks these glyphs: highlighted,
   * underlined, struck through. Applied after extraction, from the page's
   * `/Annots` (see `./text`), because it is stated about the words rather than
   * painted among them.
   */
  readonly markup?: TextMarkup;
  /**
   * §8.6.6.2 — the glyphs are filled with a tiling PATTERN, named here for the
   * caller to resolve: a pattern is a content stream, not a colour, and the
   * fill colour still standing from before is not what the page shows.
   */
  readonly fillPatternName?: string;
  /**
   * §9.3.6 — the page painted these glyphs NOWHERE: mode 3 shows nothing and
   * mode 7 only adds to the clip. A scanned page carries its recognised words
   * that way, under the picture of the page — so the run is kept, because it
   * is the only text such a document has, and a reader reproducing the page
   * leaves it to the picture.
   */
  readonly invisible?: boolean;
  /**
   * §9.3.6 — the colour the glyphs are STROKED in, when the rendering mode
   * asks for a stroke, and how wide the pen is.
   */
  readonly outlineHex?: string;
  readonly outlineWidthPt?: number;
  /** §9.8.1 — the face the glyphs were shown in is a slanted one. */
  readonly italic?: boolean;
  /**
   * §9.6.5 — the face is a Type 3 one, so what the page SHOWS here is the
   * glyph procedures, not type. The run is kept for its words; a reader that
   * reproduces the page draws the procedures instead of re-setting it.
   */
  readonly type3?: boolean;
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
/** §8.9.7 — an image written into the content stream rather than named. */
export interface InlineImage {
  /** Its dictionary, keys as written — the abbreviated ones (`/W`, `/CS`, `/F`). */
  readonly dict: PdfDict;
  /** Its bytes, still filtered as the dictionary says. */
  readonly data: Uint8Array;
}

export interface ImagePlacement {
  /** Where the `Do` fell in the stream's painting order — see {@link VectorPlacement.order}. */
  readonly order: number;
  /** XObject resource name (no leading slash), empty for an inline image. */
  readonly name: string;
  /** §8.9.7 — the image itself, where it was written into the stream. */
  readonly inline?: InlineImage;
  readonly ctm: Matrix;
  /** §8.5.4 — the clip in force when it was painted, when there was one. */
  readonly clip?: ClipRegion;
  /**
   * §8.9.6.2 — the non-stroking colour in force. A stencil `/ImageMask` carries
   * no colour of its own: it says only WHERE to paint, and this is what.
   */
  readonly fillHex: string;
  readonly mcid?: number;
  /** §11.3.5 `/BM` — a blend the page asked for that nothing here performs. */
  readonly blend?: string;
  /** §11.6.5 `/SMask` — the paint faded from place to place; nothing here does. */
  readonly masked?: boolean;
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
   * §11.3.5 `/BM` — the fill only DARKENS what it covers (`Multiply`,
   * `Darken`), so the marks under it read through. A highlighter is this and
   * nothing else.
   */
  readonly darkens?: boolean;
  /** §11.3.5 `/BM` — a blend the page asked for that nothing here performs. */
  readonly blend?: string;
  /** §11.6.5 `/SMask` — the paint faded from place to place; nothing here does. */
  readonly masked?: boolean;
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
  /** §9.6.5 — every Type 3 glyph the stream showed, with where to run it. */
  readonly glyphs: Array<Type3Call>;
  /**
   * §8.7.4.3 — the stream painted a region with a bare `sh`, which fills the
   * CLIP rather than a path and is not lifted. Counted so the reader reports it
   * where it happened: reported unconditionally, it fired on all four hundred
   * files of the pdf.js corpus, most of which contain no `sh` at all, and a
   * loss report that cries wolf on every document tells a reader nothing.
   */
  readonly bareShadings: number;
}

/**
 * §9.6.5 — one showing of a Type 3 glyph: which procedure, and the matrix that
 * puts glyph space on the page.
 */
export interface Type3Call {
  readonly stream: PdfStream;
  readonly resources: PdfDict | undefined;
  readonly ctm: Matrix;
  /** §8.5.3 — its place in the stream's painting order, as a form call has. */
  readonly order: number;
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
  /** §8.6.8 — the space `sc` / `scn` give components in. */
  fillSpace: ColorSpaceInfo | undefined;
  strokeSpace: ColorSpaceInfo | undefined;
  ctm: Matrix;
  fontKey: string;
  font: ContentFont;
  fontSize: number;
  /** §9.3.6 `Tr` — 0 fill, 1 stroke, 2 both, 3 invisible, 4–7 the same plus clip. */
  renderMode: number;
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
  fillDarkens: boolean; // §11.3.5 `/BM` Multiply or Darken — the paint only darkens
  blendMode: string | undefined; // §11.3.5 `/BM` — a blend nothing here can perform
  softMask: boolean; // §11.6.5 `/SMask` — the paint fades from place to place
  clip: ClipRegion | undefined; // §8.5.4 the clipping region in force
}

function initialState(): TextState {
  return {
    fillSpace: undefined,
    strokeSpace: undefined,
    ctm: IDENTITY,
    fontKey: '',
    font: FALLBACK_FONT,
    fontSize: 0,
    renderMode: 0,
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
    fillDarkens: false,
    blendMode: undefined,
    softMask: false,
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

/** §8.6.8 — the space a `cs` / `CS` operand names, if it is one we read. */
function spaceOf(
  operands: ReadonlyArray<PdfValue>,
  spaces: ReadonlyMap<string, ColorSpaceInfo>,
): ColorSpaceInfo | undefined {
  const name = operands[operands.length - 1];
  if (!(name instanceof PdfName)) return undefined;
  const direct = spaces.get(name.value);
  if (direct) return direct;
  // §8.6.3 — the device families may be named inline without a resource entry.
  switch (name.value) {
    case 'DeviceGray':
      return { kind: 'gray', components: 1 };
    case 'DeviceRGB':
      return { kind: 'rgb', components: 3 };
    case 'DeviceCMYK':
      return { kind: 'cmyk', components: 4 };
    default:
      return undefined;
  }
}

/**
 * §8.6.8 — the colour a run of `sc` / `scn` components comes to.
 *
 * The space in force decides, and where it was not read the COUNT is the next
 * best witness: three numbers are RGB and four are CMYK on every device space
 * there is. One number is the ambiguous case — grey in a device space, but the
 * strength of a colorant in a Separation, where 1 is the ink at full and reads
 * dark — so a lone component is only taken where the space said what it means.
 */
function componentColor(
  operands: ReadonlyArray<PdfValue>,
  space: ColorSpaceInfo | undefined,
): string | undefined {
  const nums = operands.filter((o): o is number => typeof o === 'number');
  if (nums.length === 0) return undefined;
  // §8.6.5.6/§8.6.5.7 — a CIE space's numbers mean what its own transform makes
  // of them, which is not what the same numbers mean to a device space.
  if (space?.cie) {
    const [r, g, b] = cieToSrgb(space.cie, nums);
    return rgbHex(r, g, b);
  }
  const kind = space?.kind ?? (nums.length === 3 ? 'rgb' : nums.length === 4 ? 'cmyk' : undefined);
  switch (kind) {
    case 'rgb':
      return nums.length >= 3 ? rgbHex(nums[0]!, nums[1]!, nums[2]!) : undefined;
    case 'cmyk':
      return nums.length >= 4 ? cmykHex(nums[0]!, nums[1]!, nums[2]!, nums[3]!) : undefined;
    case 'gray':
      return grayHex(nums[0]!);
    case 'tint':
      // §8.6.6.4/§8.6.6.5 — a tint is a strength of colorant, and the space's
      // own transform says what colour that strength comes to. Run it and the
      // answer is in the alternate space; devicen.pdf's three triangles are
      // green, blue and red.
      if (space?.tint) {
        const out = space.tint.transform(nums);
        const hex = componentColor(out, space.tint.alternate);
        if (hex !== undefined) return hex;
      }
      // Without a transform this can run, the honest reading is "this much
      // ink", which is the inverse of a grey level.
      return grayHex(1 - Math.max(...nums));
    default:
      return undefined;
  }
}

export function interpretContent(
  bytes: Uint8Array,
  fonts: ReadonlyMap<string, ContentFont>,
  initialCtm: Matrix = IDENTITY,
  shadings: ReadonlyMap<string, ShapeGradient> = new Map(),
  alphas: ReadonlyMap<string, GsPaint> = new Map(),
  spaces: ReadonlyMap<string, ColorSpaceInfo> = new Map(),
  hiddenOc: ReadonlySet<string> = new Set(),
): InterpretResult {
  const runs: Array<TextRun> = [];
  const images: Array<ImagePlacement> = [];
  const vectors: Array<VectorPlacement> = []; // filled paths (EP10)
  const glyphs: Array<Type3Call> = []; // §9.6.5 Type 3 glyph procedures
  let bareShadings = 0; // §8.7.4.3 `sh` — a region painted, not a path filled
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
  // §8.11.3.2 — how deep inside a `/OC … BDC` naming a group the page does NOT
  // show. The marks are still interpreted, because the graphics state they set
  // outlives them; they are simply not emitted.
  const ocStack: Array<boolean> = [];
  let hiddenDepth = 0;
  const visible = (): boolean => hiddenDepth === 0;

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
    if (path.length >= 2 && (fill || stroke) && visible()) {
      const mcid = mcStack.length > 0 ? mcStack[mcStack.length - 1] : undefined;
      vectors.push({
        order: paintOrder++,
        segs: path,
        ...(state.clip ? { clip: state.clip } : {}),
        ...(fill && state.fillPattern !== undefined ? { patternName: state.fillPattern } : {}),
        ...(fill ? { fillHex: state.fillColor } : {}),
        ...(fill && state.fillAlpha < 1 ? { alpha: state.fillAlpha } : {}),
        ...(fill && state.fillDarkens ? { darkens: true } : {}),
        ...(state.blendMode !== undefined ? { blend: state.blendMode } : {}),
        ...(state.softMask ? { masked: true } : {}),
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
    // §9.6.5 — a Type 3 glyph is a content stream, and it draws BEFORE the pen
    // moves on. The matrix that places it is the font's own, composed onto the
    // text-space matrix the glyph would have been set at.
    const type3 = state.font.type3;
    if (type3) {
      const stream = type3.proc(code);
      if (stream && visible()) {
        const scale: Matrix = [state.fontSize * state.hScale, 0, 0, state.fontSize, 0, state.rise];
        glyphs.push({
          stream,
          resources: type3.resources,
          ctm: multiply(type3.matrix, multiply(scale, multiply(tm, state.ctm))),
          order: paintOrder++,
        });
      }
    }
    const isSpace = state.font.bytesPerCode === 1 && code === 0x20;
    const spacing = state.charSpacing + (isSpace ? state.wordSpacing : 0);
    // §9.4.4 — in VERTICAL writing the pen goes down the page and the
    // horizontal scale does not touch it: `Tz` scales the writing direction,
    // which is the other axis here.
    const down = state.font.verticalAdvance;
    if (down) {
      tm = multiply(translation(0, (down(code) / 1000) * state.fontSize - spacing), tm);
      return;
    }
    const w0 = state.font.width(code) / 1000;
    const tx = (w0 * state.fontSize + spacing) * state.hScale;
    tm = multiply(translation(tx, 0), tm);
  };

  // Decode a shown string and advance the matrix glyph by glyph, returning its
  // Unicode (without emitting — Tj and TJ both build on this).
  const consume = (operand: PdfValue): string => {
    const bytes = toBytes(operand);
    const codes = state.font.splitCodes
      ? state.font.splitCodes(bytes)
      : splitCodes(bytes, state.font.bytesPerCode);
    for (const code of codes) advanceGlyph(code);
    return state.font.decode(codes);
  };

  const emitAt = (origin: Matrix, shown: string, end: Matrix): void => {
    const text = logicalOrder(shown);
    if (text.length === 0) return;
    const scaleY = Math.hypot(origin[2], origin[3]) || 1;
    const scaleX = (Math.hypot(origin[0], origin[1]) || 1) * state.hScale;
    const mcid = mcStack.length > 0 ? mcStack[mcStack.length - 1] : undefined;
    // §9.4.2 — the text matrix turns as well as moves. The baseline's own
    // direction is the first column of it; upright text leaves this at zero.
    const angle = (Math.atan2(origin[1], origin[0]) * 180) / Math.PI;
    if (!visible()) return;
    runs.push({
      text,
      x: origin[4],
      y: origin[5],
      endX: end[4],
      endY: end[5],
      ...(Math.abs(angle) > UPRIGHT_TOLERANCE_DEG ? { angleDeg: angle } : {}),
      // §9.3.1 — `Tf` may be NEGATIVE, which flips the text rather than shrinking
      // it. bug1011159.pdf sets its line at −20, and a size below zero is not a
      // size any downstream format states: the magnitude is what was shown.
      fontSizePt: Math.abs(state.fontSize) * scaleY,
      // §9.4.4 — how far the pen moves for a SPACE here: the face's own width
      // for it, plus the character and word spacing in force, which is exactly
      // the advance `advanceGlyph` gives that code. basicapi.pdf sets its page
      // number as thirty-one spaces and "page 1 / 3" in one run, and the ink in
      // it starts only where those thirty-one advances end.
      spaceWidthPt:
        ((state.font.width(0x20) / 1000) * state.fontSize + state.charSpacing + state.wordSpacing) *
        scaleX,
      fontKey: state.fontKey,
      ...(state.font.name !== undefined ? { fontName: state.font.name } : {}),
      ...(state.font.type3 ? { type3: true } : {}),
      ...(state.renderMode === 3 || state.renderMode === 7 ? { invisible: true } : {}),
      ...(state.fillPattern !== undefined ? { fillPatternName: state.fillPattern } : {}),
      // §9.3.6 — modes 1, 2, 5 and 6 stroke the glyphs; the pen is the one the
      // graphics state holds, in the stroking colour.
      ...(strokesText(state.renderMode)
        ? {
            outlineHex: state.strokeColor,
            outlineWidthPt: ctmLineWidth(),
          }
        : {}),
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
      case 'Tr':
        // §9.3.6 — how the glyphs are painted, if at all.
        state.renderMode = num(0);
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
        // §8.11.3.2 — `/OC /Name BDC` guards everything up to its `EMC` on a
        // group the file may have turned off. issue11144_reduced.pdf keeps
        // three versions of its page this way, two of them off, and read
        // without this they were drawn over the one a viewer shows.
        const isOc =
          tag instanceof PdfName &&
          tag.value === 'OC' &&
          props instanceof PdfName &&
          hiddenOc.has(props.value);
        ocStack.push(isOc);
        if (isOc) hiddenDepth++;
        break;
      }
      case 'BMC':
        mcStack.push(undefined); // `tag BMC` — no properties, so no MCID
        ocStack.push(false);
        break;
      case 'EMC':
        mcStack.pop();
        if (ocStack.pop() === true) hiddenDepth--;
        break;
      case 'Do': {
        // Paint an XObject (image or form). Record its name + the CTM (which
        // already folds in the placement `cm`) so a later stage can resolve and
        // size it; tag it with the enclosing structure id (a /Figure).
        const nm = operands[0];
        if (nm instanceof PdfName && visible()) {
          const mcid = mcStack.length > 0 ? mcStack[mcStack.length - 1] : undefined;
          images.push({
            order: paintOrder++,
            name: nm.value,
            ctm: state.ctm,
            ...(state.clip ? { clip: state.clip } : {}),
            fillHex: state.fillColor,
            ...(mcid !== undefined ? { mcid } : {}),
            ...(state.blendMode !== undefined ? { blend: state.blendMode } : {}),
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
      // §8.6.8 — `cs` / `CS` select the space the next `sc` / `scn` gives
      // components in. Without it the components mean nothing: this file paints
      // its page with `/Cs1 cs 1 1 1 sc`, which is WHITE, and left unread the
      // fill stayed at the black it starts on and covered the whole sheet.
      case 'cs':
        state.fillSpace = spaceOf(operands, spaces);
        break;
      case 'CS':
        state.strokeSpace = spaceOf(operands, spaces);
        break;
      // §8.6.8 colour in a named space; a /Pattern name selects a shading
      // pattern (EP16c), numeric operands are components in the current space.
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
        if (named !== undefined) {
          state.fillGradient = shadings.get(named);
          state.fillPattern = state.fillGradient ? undefined : named;
          break;
        }
        const hex = componentColor(operands, state.fillSpace);
        if (hex !== undefined) {
          state.fillColor = hex;
          state.fillGradient = undefined;
          state.fillPattern = undefined;
        }
        break;
      }
      case 'SCN':
      case 'SC': {
        const last = operands[operands.length - 1];
        if (last instanceof PdfName) break;
        const hex = componentColor(operands, state.strokeSpace);
        if (hex !== undefined) state.strokeColor = hex;
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
      // §8.7.4.3 — `sh` paints the CLIP with a shading, not a path with a fill.
      // Nothing here lifts that region; counted so the reader can say so where
      // it happened rather than on every document it reads.
      case 'sh':
        bareShadings++;
        break;
      case 'n':
        paintPath(false, false); // end the path with no paint
        break;
      case 'gs': {
        // §8.4.5 — a named graphics state, of which the constant fill alpha
        // and the blend mode are read here. A band meant to be seen through is
        // not the same mark as one that hides what it covers.
        const nm = operands[operands.length - 1];
        if (nm instanceof PdfName) {
          const paint = alphas.get(nm.value);
          state.fillAlpha = paint?.alpha ?? 1;
          state.fillDarkens = paint?.darkens ?? false;
          state.blendMode = paint?.blend;
          state.softMask = paint?.masked ?? false;
        }
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
        if (tok.value === 'BI') {
          // §8.9.7 — an image written into the stream itself rather than named.
          const inline = readInlineImage(lexer);
          if (inline && visible()) {
            images.push({
              order: paintOrder++,
              name: '',
              inline,
              ctm: state.ctm,
              ...(state.clip ? { clip: state.clip } : {}),
              fillHex: state.fillColor,
              ...(mcStack.length > 0 ? { mcid: mcStack[mcStack.length - 1]! } : {}),
              ...(state.blendMode !== undefined ? { blend: state.blendMode } : {}),
            });
          }
        } else exec(tok.value);
        operands = [];
        break;
      default:
        operands = []; // stray ] or >> — reset
        break;
    }
  }
  return { texts: runs, images, vectors, glyphs, bareShadings };
}

/** §9.3.6 — the rendering modes that put a line round the glyphs. */
function strokesText(mode: number): boolean {
  return mode === 1 || mode === 2 || mode === 5 || mode === 6;
}

/**
 * §9.4 — a show operator paints its glyphs along the baseline, left to right,
 * whatever the script. For Arabic or Hebrew that means the string a PDF holds
 * is in VISUAL order: the first character of the word is the last one shown.
 *
 * Every reader turns this back into logical order before handing it on, because
 * everything downstream — a search, a markdown file, a layout engine that does
 * its own bidi — takes logical order and reverses it again for display.
 * ArabicCIDTrueType.pdf came out mirrored for exactly that reason: the reader
 * passed visual order through and the layout reversed it a second time.
 *
 * A run is reversed only when it is wholly right-to-left. Anything mixed — a
 * number inside an Arabic sentence runs left to right — needs the full bidi
 * algorithm, and guessing at it would be worse than leaving it alone.
 */
function logicalOrder(text: string): string {
  return isRightToLeft(text) ? [...text].reverse().join('') : text;
}

/**
 * Whether a string is wholly right-to-left: at least one letter of an RTL
 * script and nothing of any other, spaces and joiners aside.
 *
 * Anything mixed — a number inside an Arabic sentence runs left to right —
 * needs the full bidi algorithm, and guessing at it would be worse than
 * leaving it alone.
 *
 * @param text The string to judge.
 * @returns Whether it is one run of right-to-left script.
 */
export function isRightToLeft(text: string): boolean {
  let rtl = false;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (isRtl(cp)) {
      rtl = true;
      continue;
    }
    if (cp !== 0x20 && cp !== 0x0a && cp !== 0x200c && cp !== 0x200d) return false;
  }
  return rtl;
}

/** The right-to-left blocks: Hebrew, Arabic, Syriac, Thaana, and the forms. */
function isRtl(cp: number): boolean {
  return (
    (cp >= 0x0590 && cp <= 0x05ff) || // Hebrew
    (cp >= 0x0600 && cp <= 0x06ff) || // Arabic
    (cp >= 0x0700 && cp <= 0x074f) || // Syriac
    (cp >= 0x0780 && cp <= 0x07bf) || // Thaana
    (cp >= 0x08a0 && cp <= 0x08ff) || // Arabic Extended-A
    (cp >= 0xfb1d && cp <= 0xfdff) || // Hebrew + Arabic Presentation Forms-A
    (cp >= 0xfe70 && cp <= 0xfeff) // Arabic Presentation Forms-B
  );
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

/**
 * §8.9.7 — an inline image: its dictionary between `BI` and `ID`, then its
 * bytes up to `EI`.
 *
 * The bytes are binary and may hold `EI` themselves, so the end is found by
 * MEASURING where the dictionary says how much there is — an unfiltered image
 * is exactly `ceil(W · BPC · components / 8) · H` bytes — and only searched for
 * where a filter makes the length unknowable. images_1bit_grayscale.pdf draws
 * two of them and both were skipped over.
 */
function readInlineImage(lexer: Lexer): InlineImage | undefined {
  const dict: PdfDict = new Map<string, PdfValue>();
  for (;;) {
    const tok = lexer.nextToken();
    if (tok.kind === 'eof') return undefined;
    if (tok.kind === 'keyword' && tok.value === 'ID') break;
    if (tok.kind === 'name') dict.set(tok.value, readValue(lexer));
  }
  // §8.9.7 — exactly ONE whitespace byte separates `ID` from the data.
  const start = lexer.pos + 1;
  const measured = inlineLength(dict);
  let end = measured !== undefined ? start + measured : -1;
  if (end < 0 || end > lexer.length) {
    end = lexer.indexOfAscii('EI', start);
    if (end < 0) {
      lexer.pos = lexer.length;
      return undefined;
    }
  }
  const data = lexer.slice(start, Math.min(end, lexer.length));
  const ei = lexer.indexOfAscii('EI', end);
  lexer.pos = ei < 0 ? lexer.length : ei + 2;
  return { dict, data };
}

/** How many bytes an UNFILTERED inline image's samples take, if that is known. */
function inlineLength(dict: PdfDict): number | undefined {
  if (dict.has('F') || dict.has('Filter')) return undefined;
  const num = (...keys: ReadonlyArray<string>): number => {
    for (const k of keys) {
      const v = dict.get(k);
      if (typeof v === 'number') return v;
    }
    return 0;
  };
  const w = num('W', 'Width');
  const h = num('H', 'Height');
  if (w <= 0 || h <= 0) return undefined;
  const mask = dict.get('IM') === true || dict.get('ImageMask') === true;
  const bpc = mask ? 1 : num('BPC', 'BitsPerComponent') || 8;
  const cs = dict.get('CS') ?? dict.get('ColorSpace');
  const name = cs instanceof PdfName ? cs.value : '';
  const comps = mask
    ? 1
    : /^(RGB|DeviceRGB|CalRGB)$/u.test(name)
      ? 3
      : /^(CMYK|DeviceCMYK)$/u.test(name)
        ? 4
        : /^(G|DeviceGray|CalGray|I|Indexed)$/u.test(name) || name === ''
          ? 1
          : 0;
  if (comps === 0) return undefined; // a space this cannot size: search instead
  return (((w * comps * bpc + 7) / 8) | 0) * h;
}
