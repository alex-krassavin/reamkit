// §20.1.9.10 `a:prstTxWarp` — the preset warp a text body is bent through.
//
// A warp is TWO curves over the shape's box: the top edge the text hangs from
// and the bottom edge it stands on. A glyph at horizontal position `u` is set
// between the two, so where the pair spreads the text is tall, where it pinches
// the text is short, and where the pair rises together the text rises with it.
// The whole warped block then FILLS the box — the size the runs state stops
// mattering, which is why a piece of WordArt keeps its proportions and not its
// point size when its shape is resized.
//
// The envelopes below are stated in the box's own normalised frame: `u` runs
// 0…1 left to right, and the returned `top`/`bottom` run 0…1 top to bottom.
// Every one of them is normalised so the block exactly fills the box —
// `min top = 0`, `max bottom = 1` — which is what both references draw.
//
// Where a curve carries a shape constant it was read off the rendered envelope
// of a box of capitals: the spec states these as Bézier pairs whose control
// points are not public in a form this file could cite, so each preset is the
// closed form that reproduces the drawn edge. The families that are exact —
// the straight tapers, the sinusoids, the linear slants — say so by having
// round constants.

/** The two edges of a warp at one horizontal position, `0`…`1` down the box. */
export interface WarpEdges {
  readonly top: number;
  readonly bottom: number;
}

/** A warp: the pair of edges at any `u`, plus whether it bends at all. */
type Envelope = (u: number, adj: number) => WarpEdges;

/** Distance from the box's centre line, `0` at the middle and `1` at either edge. */
const fromCentre = (u: number): number => Math.abs(2 * u - 1);

/** A band of constant height `1 - drop` whose top edge is `top`. */
const band = (top: number, drop: number): WarpEdges => ({ top, bottom: top + (1 - drop) });

/**
 * §20.1.10.76 ST_TextShapeType — the warp each preset applies, keyed by its
 * name. `adj` arrives as the `a:avLst` adjustment in hundred-thousandths, or
 * the preset's own default when the file states none.
 */
const ENVELOPES = new Map<string, { readonly defaultAdj: number; readonly at: Envelope }>([
  // The text simply fills the box: no bend, but still the WordArt stretch.
  ['textPlain', { defaultAdj: 0, at: () => ({ top: 0, bottom: 1 }) }],
  // Square in the middle, chamfered over the outer third at each end.
  [
    'textStop',
    {
      defaultAdj: 12500,
      at: (u, adj) => {
        const t = adj * Math.max(0, (fromCentre(u) - 0.375) / 0.625);
        return { top: t, bottom: 1 - t };
      },
    },
  ],
  // A triangle: the baseline is flat and the top rises to a point at the centre.
  [
    'textTriangle',
    { defaultAdj: 50000, at: (u, adj) => ({ top: adj * fromCentre(u), bottom: 1 }) },
  ],
  [
    'textTriangleInverted',
    { defaultAdj: 50000, at: (u, adj) => ({ top: 0, bottom: 1 - adj * fromCentre(u) }) },
  ],
  // A chevron: the band keeps its height and both edges peak at the centre.
  ['textChevron', { defaultAdj: 25000, at: (u, adj) => band(adj * fromCentre(u), adj) }],
  [
    'textChevronInverted',
    { defaultAdj: 25000, at: (u, adj) => band(adj * (1 - fromCentre(u)), adj) },
  ],
  // A can: the band keeps its height and both edges swell away from the centre.
  ['textCanUp', { defaultAdj: 14000, at: (u, adj) => band(adj * fromCentre(u) ** 4, adj) }],
  ['textCanDown', { defaultAdj: 14000, at: (u, adj) => band(adj * (1 - fromCentre(u) ** 4), adj) }],
  // One sine period across the box, the band riding it at constant height.
  [
    'textWave1',
    { defaultAdj: 6250, at: (u, adj) => band(adj * (1 - Math.sin(2 * Math.PI * u)), 2 * adj) },
  ],
  [
    'textWave2',
    { defaultAdj: 6250, at: (u, adj) => band(adj * (1 + Math.sin(2 * Math.PI * u)), 2 * adj) },
  ],
  // Two periods — `textWave4` is the second of them read the other way up.
  [
    'textDoubleWave1',
    { defaultAdj: 6250, at: (u, adj) => band(adj * (1 - Math.sin(4 * Math.PI * u)), 2 * adj) },
  ],
  [
    'textWave4',
    { defaultAdj: 6250, at: (u, adj) => band(adj * (1 + Math.sin(4 * Math.PI * u)), 2 * adj) },
  ],
  // Inflate/deflate: the band's height alone changes, about a fixed centre line.
  [
    'textInflate',
    {
      defaultAdj: 13600,
      at: (u, adj) => {
        const t = adj * fromCentre(u) ** 2;
        return { top: t, bottom: 1 - t };
      },
    },
  ],
  [
    'textDeflate',
    {
      defaultAdj: 37500,
      at: (u, adj) => {
        const t = adj * (1 - fromCentre(u) ** 2.23);
        return { top: t, bottom: 1 - t };
      },
    },
  ],
  // …and the one-sided pairs, which hold one edge flat and bend the other.
  [
    'textInflateTop',
    { defaultAdj: 31800, at: (u, adj) => ({ top: adj * fromCentre(u) ** 2, bottom: 1 }) },
  ],
  [
    'textDeflateTop',
    { defaultAdj: 47000, at: (u, adj) => ({ top: adj * (1 - fromCentre(u) ** 2), bottom: 1 }) },
  ],
  [
    'textInflateBottom',
    { defaultAdj: 31800, at: (u, adj) => ({ top: 0, bottom: 1 - adj * fromCentre(u) ** 2 }) },
  ],
  [
    'textDeflateBottom',
    { defaultAdj: 47000, at: (u, adj) => ({ top: 0, bottom: 1 - adj * (1 - fromCentre(u) ** 2) }) },
  ],
  // The two stacked-band warps. A body of one line only ever reaches the first
  // band; the rest of the stack is the second and third line's, which the
  // layout does not yet split out.
  [
    'textDeflateInflate',
    { defaultAdj: 19800, at: (u, adj) => ({ top: 0, bottom: 0.268 + adj * fromCentre(u) ** 2 }) },
  ],
  [
    'textDeflateInflateDeflate',
    { defaultAdj: 4600, at: (u, adj) => ({ top: 0, bottom: 0.258 + adj * fromCentre(u) ** 2 }) },
  ],
  // Fades: the band closes towards one side of the box, or towards both ends.
  [
    'textFadeRight',
    {
      defaultAdj: 33333,
      at: (u, adj) => ({ top: adj * u, bottom: 1 - adj * u }),
    },
  ],
  [
    'textFadeLeft',
    {
      defaultAdj: 33333,
      at: (u, adj) => ({ top: adj * (1 - u), bottom: 1 - adj * (1 - u) }),
    },
  ],
  [
    'textFadeUp',
    { defaultAdj: 100000, at: (u, adj) => ({ top: adj * fromCentre(u) ** 2, bottom: 1 }) },
  ],
  [
    'textFadeDown',
    { defaultAdj: 100000, at: (u, adj) => ({ top: 0, bottom: 1 - adj * fromCentre(u) ** 2 }) },
  ],
  // Slants: the band keeps its height and slides along one diagonal.
  ['textSlantUp', { defaultAdj: 55500, at: (u, adj) => band(adj * (1 - u), adj) }],
  ['textSlantDown', { defaultAdj: 44400, at: (u, adj) => band(adj * u, adj) }],
  // Cascades: the band both slides and tapers, so the words step down in size.
  [
    'textCascadeUp',
    {
      defaultAdj: 55500,
      at: (u, adj) => ({ top: 0.138 * (1 - u), bottom: 1 - adj + adj * (1 - u) }),
    },
  ],
  [
    'textCascadeDown',
    { defaultAdj: 55500, at: (u, adj) => ({ top: 0.138 * u, bottom: 1 - adj + adj * u }) },
  ],
  // Curves: an arc across the box that also tapers towards the far end.
  ['textCurveUp', { defaultAdj: 45600, at: (u, adj) => curve(1 - u, adj) }],
  ['textCurveDown', { defaultAdj: 45600, at: (u, adj) => curve(u, adj) }],
]);

/**
 * The ring presets, which do not hang text between two edges at all: they wind
 * it once round an ellipse inscribed in the box. Both start at the box's left
 * edge; `Inside` climbs over the top with each letter's head pointing out of
 * the ring, `Outside` dips under the bottom with each letter's head pointing
 * into it — which is why half of either one reads upside down.
 *
 * The band the letters occupy is a fixed share of the ring whatever the text:
 * eight letters and thirty-two fill exactly the same annulus, so it is the
 * preset and not the string that fixes the inner radius.
 */
const RING_INNER = 0.34;
const RINGS = new Map<string, { readonly clockwise: boolean; readonly headOut: boolean }>([
  ['textRingInside', { clockwise: true, headOut: true }],
  ['textRingOutside', { clockwise: false, headOut: false }],
]);

/**
 * The `textCurve*` pair: the baseline is an arc over the box and the band
 * narrows towards the end the text runs to.
 *
 * @param u   Position across the box, already turned the way the preset runs.
 * @param adj How much of the band's height the taper takes by the far end.
 * @returns   The two edges at `u`.
 */
function curve(u: number, adj: number): WarpEdges {
  const bottom = 1 - 0.24 * (2 * u - 1) ** 2;
  return { top: bottom - (0.758 - adj * u ** 1.7), bottom };
}

/**
 * Whether a preset bends text at all. `textNoShape` is the "no warp" member of
 * the enumeration and by far its most common value — two thirds of the decks
 * that mention a warp state only this one — and a body carrying it is an
 * ordinary text box, NOT WordArt stretched to its shape.
 *
 * @param preset The `a:prstTxWarp @prst` value.
 * @returns Whether {@link textWarpEdges} can bend text through this preset.
 */
export function isTextWarp(preset: string): boolean {
  return ENVELOPES.has(preset) || RINGS.has(preset);
}

/**
 * Where a warped body's two frames sit on the page: the shape's box, which the
 * warped text is stretched onto, and the box the UN-warped block occupies.
 * Both are in the page's top-left frame, y growing down.
 */
export interface WarpFrame {
  readonly preset: string;
  readonly adjust?: number;
  readonly boxX: number;
  readonly boxY: number;
  readonly boxWidth: number;
  readonly boxHeight: number;
  readonly srcX: number;
  readonly srcWidth: number;
  readonly srcTop: number;
  readonly srcHeight: number;
}

/**
 * The matrix that places one glyph of a warped body.
 *
 * Maps the glyph's own text space — x along its advance, y UP from its
 * baseline — into the page's top-left frame, as `[a, b, c, d, e, f]` with
 * `px = a·x + c·y + e` and `py = b·x + d·y + f`. Upright text therefore has a
 * negative `d`: the glyph's y runs against the page's.
 *
 * @param frame     The two frames, from the laid-out line.
 * @param glyphX    The glyph's left edge in the un-warped block.
 * @param advance   How wide the glyph is there.
 * @param baselineY The line's baseline in the un-warped block.
 * @returns The six numbers, or `undefined` where the warp leaves no room to
 *          draw — the pinched end of a fade, or a preset this file cannot bend.
 */
export function warpGlyphMatrix(
  frame: WarpFrame,
  glyphX: number,
  advance: number,
  baselineY: number,
): readonly [number, number, number, number, number, number] | undefined {
  const u = (glyphX + advance / 2 - frame.srcX) / frame.srcWidth;
  const v = (baselineY - frame.srcTop) / frame.srcHeight;
  const ring = RINGS.get(frame.preset);
  if (ring) return ringMatrix(frame, ring, u, v, glyphX);
  return envelopeMatrix(frame, u, v, glyphX, advance);
}

/** The glyph matrix for a warp that hangs text between a top and a bottom edge. */
function envelopeMatrix(
  frame: WarpFrame,
  u: number,
  v: number,
  glyphX: number,
  advance: number,
): readonly [number, number, number, number, number, number] | undefined {
  const edges = textWarpEdges(frame.preset, u, frame.adjust);
  if (!edges) return undefined;
  const height = edges.bottom - edges.top;
  if (!(height > 0)) return undefined;
  const sx = frame.boxWidth / frame.srcWidth;
  const sy = (height * frame.boxHeight) / frame.srcHeight;
  // The slope the baseline runs at here, as page-y over page-x: the shear that
  // leans each glyph along the curve instead of leaving a rising wave as a
  // staircase of upright letters. Read off the curve either side of the glyph,
  // since a preset states its edges and not their derivative.
  const step = 0.004;
  const lo = Math.max(0, u - step);
  const hi = Math.min(1, u + step);
  const at = (q: number): number | undefined => {
    const e = textWarpEdges(frame.preset, q, frame.adjust);
    return e ? e.top + v * (e.bottom - e.top) : undefined;
  };
  const before = at(lo);
  const after = at(hi);
  const slope =
    before !== undefined && after !== undefined && hi > lo
      ? ((after - before) * frame.boxHeight) / ((hi - lo) * frame.boxWidth)
      : 0;
  const x0 = frame.boxX + ((glyphX - frame.srcX) / frame.srcWidth) * frame.boxWidth;
  const mid = edges.top + v * height;
  const y0 = frame.boxY + mid * frame.boxHeight - (slope * advance * sx) / 2;
  return [sx, slope * sx, 0, -sy, x0, y0];
}

/**
 * The glyph matrix for a ring: one turn round the ellipse inscribed in the box.
 *
 * The band is a CONSTANT inset from that ellipse, not a scaled copy of it. A
 * scaled copy would be the natural reading — a circular annulus squashed to the
 * box's proportions — but it makes the band four times deeper at the ring's
 * sides than at its top on a box four times wider than tall, and the letters
 * there come out as spikes radiating from the centre. Held at a constant depth
 * the ring's outer edge still touches the box all the way round, which is what
 * both references draw and what a box of capitals measures.
 */
function ringMatrix(
  frame: WarpFrame,
  ring: { readonly clockwise: boolean; readonly headOut: boolean },
  _u: number,
  v: number,
  glyphX: number,
): readonly [number, number, number, number, number, number] | undefined {
  const a = frame.boxWidth / 2;
  const b = frame.boxHeight / 2;
  if (!(a > 0 && b > 0 && frame.srcWidth > 0 && frame.srcHeight > 0)) return undefined;
  // Both rings open at the box's left edge; the turn runs one way or the other
  // from there. The glyph's left edge decides where it starts, so a letter is
  // set from its own corner exactly as a flat one is.
  const uLeft = (glyphX - frame.srcX) / frame.srcWidth;
  const turn = ring.clockwise ? -1 : 1;
  const psi = Math.PI + turn * 2 * Math.PI * uLeft;
  const cos = Math.cos(psi);
  const sin = Math.sin(psi);
  // The tangent the letter is set along, and how fast the ellipse runs there:
  // the letters crowd where the curve is tight and spread where it is flat, so
  // the line's own width still goes exactly once round.
  const rawX = -turn * a * sin;
  const rawY = -turn * b * cos;
  const speed = Math.hypot(rawX, rawY);
  if (!(speed > 0)) return undefined;
  const tx = rawX / speed;
  const ty = rawY / speed;
  // Inward, and then the way the letter's head points: out of the ring for
  // `Inside`, into it for `Outside`, which is why half of either reads upside
  // down.
  const inX = -turn * -ty;
  const inY = -turn * tx;
  const head = ring.headOut ? -1 : 1;
  const depth = (1 - RING_INNER) * b;
  const sT = (2 * Math.PI * speed) / frame.srcWidth;
  const sN = depth / frame.srcHeight;
  // How far in from the outer edge this line's baseline sits.
  const inset = (ring.headOut ? v : 1 - v) * depth;
  return [
    sT * tx,
    sT * ty,
    sN * head * inX,
    sN * head * inY,
    frame.boxX + a + a * cos + inX * inset,
    frame.boxY + b - b * sin + inY * inset,
  ];
}

/**
 * The two edges of a warp at one position across its box.
 *
 * @param preset The `a:prstTxWarp @prst` value.
 * @param u      Position across the box, `0` at its left edge and `1` at its right.
 * @param adj    The `a:avLst` `adj` guide in hundred-thousandths, when the file states one.
 * @returns The top and bottom edges, `0`…`1` down the box, or `undefined` for a
 *          preset this file does not bend (an unknown name, or `textNoShape`).
 */
export function textWarpEdges(preset: string, u: number, adj?: number): WarpEdges | undefined {
  const env = ENVELOPES.get(preset);
  if (!env) return undefined;
  const a = (adj ?? env.defaultAdj) / 100000;
  return env.at(Math.min(1, Math.max(0, u)), a);
}
