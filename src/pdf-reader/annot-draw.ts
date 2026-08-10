// §12.5.5 — the appearance a markup annotation does NOT carry.
//
// "If the annotation does not contain an appearance stream, the conforming
// reader shall generate one." Nine files of the pdf.js corpus are exactly that
// case — a line, a square, a circle, an ink scrawl, a polygon — and every one
// came back as a blank page while poppler and LibreOffice both drew it. The
// geometry is all there in the annotation's own dictionary; what is missing is
// only the content stream that paints it.
//
// The four TEXT-MARKUP subtypes are not among them: a highlight marks WORDS,
// not a place on the paper, so it comes back on the runs (`textMarkupOf`).
//
// So it is written here, in PAGE space, as the operators a viewer would have
// found in `/AP` `/N`. Only for the markup subtypes whose geometry says exactly
// what to draw: a WIDGET with no appearance for the state in force is a field
// that draws nothing — an unticked box — not one to invent a look for.

import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { PdfFile } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

/** The circle constant: how far a Bézier handle reaches to round a quarter. */
const KAPPA = 0.5523;

/** §12.5.2 `/Border` and §12.5.4 `/BS` `/W` both default the pen to one point. */
const DEFAULT_BORDER_PT = 1;

/** A path in page space and the operator that paints it. */
interface Drawing {
  readonly ops: ReadonlyArray<string>;
  /** `'S'` stroke, `'f'` fill, `'B'` both — §8.5.3.1. */
  readonly paint: 'S' | 'f' | 'B';
}

/**
 * A generated normal appearance for an annotation that carries none, or
 * `undefined` when its subtype or its geometry says nothing to draw.
 *
 * The stream is in PAGE space, so the caller places it with the identity: it is
 * not fitted to `/Rect` the way an authored appearance is, because what it is
 * drawn from is already stated in page coordinates.
 *
 * @param file  The owning file, for resolving references.
 * @param annot The annotation dictionary.
 * @returns The synthesized appearance stream, or `undefined`.
 */
export function drawnAppearance(file: PdfFile, annot: PdfDict): PdfStream | undefined {
  const subtype = file.get(annot, 'Subtype');
  if (!(subtype instanceof PdfName)) return undefined;
  // §12.7.3.3 — a field whose appearance was never generated still HAS its
  // value, and the value is the form's content. 160F-2019.pdf files seven of
  // them and every one was lost.
  if (subtype.value === 'Widget') return typedValue(file, annot);
  const pen = borderWidth(file, annot);
  const drawing = pathFor(file, annot, subtype.value, pen);
  if (!drawing || drawing.ops.length === 0) return undefined;
  const stroke = colorOf(file.get(annot, 'C'));
  const fillColor = colorOf(file.get(annot, 'IC'));
  const body = [
    'q',
    `${num(pen)} w`,
    ...(stroke !== undefined ? [`${stroke} RG`] : []),
    ...(fillColor !== undefined ? [`${fillColor} rg`] : []),
    ...drawing.ops,
    drawing.paint,
    'Q',
  ].join('\n');
  return new PdfStream(new Map(), new TextEncoder().encode(body));
}

/**
 * §12.7.2 `/DR` — the resources a generated field appearance draws with.
 *
 * The `/DA` string names a font (`/Helv 12 Tf`) and that name means nothing
 * without the form's own default resource dictionary, which is where every
 * field's face lives.
 *
 * @param file  The owning file.
 * @param annot The annotation the appearance was generated for.
 * @returns The form's `/DR`, or `undefined` when there is none.
 */
export function drawnResources(file: PdfFile, annot: PdfDict): PdfDict | undefined {
  const subtype = file.get(annot, 'Subtype');
  if (!(subtype instanceof PdfName) || subtype.value !== 'Widget') return undefined;
  const acro = file.get(file.catalog, 'AcroForm');
  if (!(acro instanceof Map)) return undefined;
  const dr = file.get(acro, 'DR');
  return dr instanceof Map ? dr : undefined;
}

/**
 * §12.7.3.3 — the text a form field holds, set the way its `/DA` asks.
 *
 * A widget normally carries the picture of its value in `/AP`, and a viewer
 * that fills a field regenerates it. One that does not — or a file flagged
 * `NeedAppearances` — leaves only `/V`, and reading the appearances alone gives
 * a form with every box empty. The field's own default appearance string names
 * the face, the size and the colour; the rest is arithmetic on `/Rect`.
 *
 * Only a TEXT field: a button's value is a state, and a choice field's list is
 * drawn from `/Opt` rather than typed into a box.
 */
function typedValue(file: PdfFile, annot: PdfDict): PdfStream | undefined {
  if (inherited(file, annot, 'FT') !== 'Tx') return undefined;
  const value = inheritedString(file, annot, 'V');
  if (value === undefined || value.length === 0) return undefined;
  const rect = rectangle(file.get(annot, 'Rect'));
  if (!rect) return undefined;
  const da = defaultAppearance(file, annot);
  const flags = inheritedNumber(file, annot, 'Ff') ?? 0;
  const multiline = (flags & MULTILINE) !== 0;
  const box = inset(rect, borderWidth(file, annot) + 1);
  const height = box[3] - box[1];
  const width = box[2] - box[0];
  if (!(width > 0 && height > 0)) return undefined;
  // A size of zero asks the viewer to fit the box (§12.7.3.3). One line in it,
  // or a readable twelve where there are many.
  const size = da.size > 0 ? da.size : multiline ? 12 : Math.min(12, height * 0.66);
  const lines = multiline ? value.split(/\r\n|[\r\n]/u) : [value.replace(/[\r\n]+/gu, ' ')];
  const leading = size * 1.16;
  // The first baseline: one line sits on the box's middle, many start at its top.
  const first = multiline ? box[3] - size : box[1] + (height - size) / 2 + size * 0.22;
  const ops: Array<string> = ['q', 'BT', `${da.font} ${num(size)} Tf`, `${da.color} rg`];
  lines.forEach((line, i) => {
    const y = first - i * leading;
    if (y < box[1] - size) return;
    ops.push(
      `1 0 0 1 ${num(box[0] + indent(line, size, width, quadding(file, annot)))} ${num(y)} Tm`,
    );
    ops.push(`(${escapeText(line)}) Tj`);
  });
  ops.push('ET', 'Q');
  return new PdfStream(new Map(), new TextEncoder().encode(ops.join('\n')));
}

/** §12.7.4.3 — the field is a multi-line text box (bit 13 of `/Ff`). */
const MULTILINE = 1 << 12;

/**
 * §12.7.3.1 `/Q` — 0 left, 1 centred, 2 right. Where it is not left the line
 * has to be measured, and the face is not in hand: half the size per character
 * is what Helvetica averages, and a field's value is short enough that the
 * error stays under a character.
 */
function indent(line: string, size: number, width: number, q: number): number {
  if (q !== 1 && q !== 2) return 0;
  const guess = Math.min(width, line.length * size * 0.5);
  return q === 1 ? (width - guess) / 2 : width - guess;
}

/** §12.7.3.1 `/Q`, inherited from the field's parent. */
function quadding(file: PdfFile, annot: PdfDict): number {
  return inheritedNumber(file, annot, 'Q') ?? 0;
}

/**
 * §12.7.3.3 `/DA` — the default appearance string, from the field or the form.
 * Only the `Tf` and the colour are read; the rest is a content stream fragment
 * this reader does not need to run.
 */
function defaultAppearance(
  file: PdfFile,
  annot: PdfDict,
): { font: string; size: number; color: string } {
  const acro = file.get(file.catalog, 'AcroForm');
  const own = inheritedString(file, annot, 'DA');
  const form = acro instanceof Map ? file.get(acro, 'DA') : undefined;
  const da = own ?? (typeof form === 'string' ? form : '');
  const tf = /\/([^\s/]+)\s+([\d.]+)\s+Tf/u.exec(da);
  const rgb = /([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg/u.exec(da);
  const gray = /(^|\s)([\d.]+)\s+g(\s|$)/u.exec(da);
  const color = rgb
    ? `${num(Number(rgb[1]))} ${num(Number(rgb[2]))} ${num(Number(rgb[3]))}`
    : gray
      ? `${num(Number(gray[2]))} ${num(Number(gray[2]))} ${num(Number(gray[2]))}`
      : '0 0 0';
  return { font: `/${tf?.[1] ?? 'Helv'}`, size: Number(tf?.[2] ?? 0), color };
}

/** §7.3.4.2 — the characters a literal string may not carry unescaped. */
function escapeText(s: string): string {
  return s.replace(/[\\()]/gu, (c) => `\\${c}`).replace(/[\r\n]/gu, ' ');
}

/** A field entry, from the widget or the field tree above it (§12.7.3.1). */
function inherited(file: PdfFile, annot: PdfDict, key: string): string | undefined {
  const found = climb(file, annot, key);
  return found instanceof PdfName ? found.value : undefined;
}

/** The same, for an entry stated as a string. */
function inheritedString(file: PdfFile, annot: PdfDict, key: string): string | undefined {
  const found = climb(file, annot, key);
  return typeof found === 'string' ? found : undefined;
}

/** The same, for a number. */
function inheritedNumber(file: PdfFile, annot: PdfDict, key: string): number | undefined {
  const found = climb(file, annot, key);
  return typeof found === 'number' ? found : undefined;
}

/** Walk `/Parent` until the entry is found, or the tree ends. */
function climb(file: PdfFile, annot: PdfDict, key: string): PdfValue | undefined {
  let at: PdfDict | undefined = annot;
  for (let depth = 0; at && depth < 32; depth++) {
    const here = file.get(at, key);
    if (here !== PDF_NULL) return here;
    const up: PdfValue = file.get(at, 'Parent');
    at = up instanceof Map ? up : undefined;
  }
  return undefined;
}

/**
 * §12.5.6.10 — what a TEXT-MARKUP annotation says about the words it covers.
 *
 * A highlight, an underline, a strikeout and a squiggle are not drawings: each
 * names a run of text and a way of marking it. Lifted as artwork they are a
 * band and a rule anchored to the page, which is right until the text re-sets
 * and is then a band sitting between two paragraphs it does not mark. Carried
 * on the RUNS they survive the reflow, and a .docx gets what the annotation
 * meant — `w:shd`, `w:u`, `w:strike`.
 *
 * @param file  The owning file.
 * @param annot The annotation dictionary.
 * @returns How it marks and what it marks, or `undefined` for another subtype.
 */
export function textMarkupOf(file: PdfFile, annot: PdfDict): TextMarkupAnnot | undefined {
  const subtype = file.get(annot, 'Subtype');
  if (!(subtype instanceof PdfName)) return undefined;
  const kind = subtype.value;
  if (kind !== 'Highlight' && kind !== 'Underline' && kind !== 'StrikeOut' && kind !== 'Squiggly') {
    return undefined;
  }
  const quads = eachQuad(file.get(annot, 'QuadPoints'));
  if (quads.length === 0) return undefined;
  const rgb = colorOf(file.get(annot, 'C'));
  const hex = rgb === undefined ? undefined : hexOf(rgb);
  const mark: TextMarkup =
    kind === 'Highlight'
      ? // A highlighter with no colour named is yellow, which is what one is.
        { highlightHex: hex ?? 'FFFF00' }
      : kind === 'StrikeOut'
        ? { strike: true }
        : {
            underline: kind === 'Squiggly' ? 'wave' : 'single',
            ...(hex !== undefined ? { underlineHex: hex } : {}),
          };
  return { mark, quads };
}

/** A text-markup annotation: how it marks, and the boxes it marks. */
export interface TextMarkupAnnot {
  readonly mark: TextMarkup;
  readonly quads: ReadonlyArray<Quad>;
}

/** §12.5.6.10 — how a text-markup annotation marks the words it covers. */
export interface TextMarkup {
  /** `/Highlight` — the wash painted behind them (6-hex). */
  readonly highlightHex?: string;
  /** `/Underline`, `/Squiggly` — a rule under them, straight or wavy. */
  readonly underline?: 'single' | 'wave';
  /** The rule's own colour (6-hex), when the annotation states one. */
  readonly underlineHex?: string;
  /** `/StrikeOut` — a rule through them. */
  readonly strike?: boolean;
}

/** One marked box in page space: its span of x, its foot and its height. */
export interface Quad {
  readonly x0: number;
  readonly x1: number;
  readonly y0: number;
  readonly h: number;
}

/** The `rg` operands as a 6-hex colour. */
function hexOf(operands: string): string | undefined {
  const n = operands.split(' ').map(Number);
  if (n.length !== 3 || n.some((v) => !Number.isFinite(v))) return undefined;
  return n
    .map((v) =>
      Math.round(Math.min(1, Math.max(0, v)) * 255)
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
    .toUpperCase();
}

/** What each markup subtype draws, from the geometry it states. */
function pathFor(file: PdfFile, annot: PdfDict, subtype: string, pen: number): Drawing | undefined {
  switch (subtype) {
    case 'Ink':
      // Every stroke of the scrawl is a subpath of one path, stroked once.
      return { ops: polylines(listOfLists(file.get(annot, 'InkList')), false), paint: 'S' };
    case 'Line':
      return { ops: polylines([numbers(file.get(annot, 'L'))], false), paint: 'S' };
    case 'Polygon':
    case 'PolyLine':
      return {
        ops: polylines([numbers(file.get(annot, 'Vertices'))], subtype === 'Polygon'),
        paint: filledOrStroked(file, annot),
      };
    case 'Square':
    case 'Circle': {
      const rect = rectangle(file.get(annot, 'Rect'));
      // The pen straddles the path, so the path runs half a width inside the
      // rectangle: drawn on the edge, half the border falls outside the box the
      // annotation claims for itself.
      if (!rect) return undefined;
      return {
        ops: boxOps(inset(rect, pen / 2), subtype === 'Circle'),
        paint: filledOrStroked(file, annot),
      };
    }
    default:
      // The text-markup subtypes are NOT drawn: they mark words, and a band
      // anchored where the words used to be is left behind the moment the text
      // re-sets. They come back on the runs instead — see `textMarkupOf`.
      return undefined;
  }
}

/** `/IC` fills a square, a circle or a polygon; `/C` strokes its border. */
function filledOrStroked(file: PdfFile, annot: PdfDict): 'S' | 'f' | 'B' {
  const interior = colorOf(file.get(annot, 'IC'));
  if (interior === undefined) return 'S';
  return colorOf(file.get(annot, 'C')) !== undefined ? 'B' : 'f';
}

/** Each run of x y pairs as one subpath, closed or open. */
function polylines(runs: ReadonlyArray<ReadonlyArray<number>>, close: boolean): Array<string> {
  const ops: Array<string> = [];
  for (const pts of runs) {
    if (pts.length < 4) continue;
    ops.push(`${num(pts[0]!)} ${num(pts[1]!)} m`);
    for (let i = 2; i + 1 < pts.length; i += 2) ops.push(`${num(pts[i]!)} ${num(pts[i + 1]!)} l`);
    if (close) ops.push('h');
  }
  return ops;
}

/** §12.5.6.8 — a square is its rectangle; a circle is the ellipse inside it. */
function boxOps(r: readonly [number, number, number, number], ellipse: boolean): Array<string> {
  const [x0, y0, x1, y1] = r;
  if (!(x1 > x0 && y1 > y0)) return [];
  if (!ellipse) return [`${num(x0)} ${num(y0)} ${num(x1 - x0)} ${num(y1 - y0)} re`];
  const cx = (x0 + x1) / 2;
  const cy = (y0 + y1) / 2;
  const rx = (x1 - x0) / 2;
  const ry = (y1 - y0) / 2;
  const hx = rx * KAPPA;
  const hy = ry * KAPPA;
  const c = (ax: number, ay: number, bx: number, by: number, x: number, y: number): string =>
    `${num(ax)} ${num(ay)} ${num(bx)} ${num(by)} ${num(x)} ${num(y)} c`;
  return [
    `${num(cx + rx)} ${num(cy)} m`,
    c(cx + rx, cy + hy, cx + hx, cy + ry, cx, cy + ry),
    c(cx - hx, cy + ry, cx - rx, cy + hy, cx - rx, cy),
    c(cx - rx, cy - hy, cx - hx, cy - ry, cx, cy - ry),
    c(cx + hx, cy - ry, cx + rx, cy - hy, cx + rx, cy),
    'h',
  ];
}

/**
 * §12.5.6.10 `/QuadPoints` — the run of text a text-markup annotation marks,
 * eight numbers per quad. Their stated order is notoriously not the order
 * producers write them in, so each quad is taken as the box its four corners
 * bound, which comes to the same box either way for text that is not turned.
 */
function eachQuad(quads: PdfValue | undefined): Array<Quad> {
  const n = numbers(quads);
  const out: Array<Quad> = [];
  for (let q = 0; q + 7 < n.length; q += 8) {
    const xs = [n[q]!, n[q + 2]!, n[q + 4]!, n[q + 6]!];
    const ys = [n[q + 1]!, n[q + 3]!, n[q + 5]!, n[q + 7]!];
    const x0 = Math.min(...xs);
    const x1 = Math.max(...xs);
    const y0 = Math.min(...ys);
    const h = Math.max(...ys) - y0;
    if (x1 > x0 && h > 0) out.push({ x0, x1, y0, h });
  }
  return out;
}

/** §12.5.4 `/BS` `/W`, else §12.5.2 `/Border`'s third number, else one point. */
function borderWidth(file: PdfFile, annot: PdfDict): number {
  const bs = file.get(annot, 'BS');
  if (bs instanceof Map) {
    const w = file.get(bs, 'W');
    if (typeof w === 'number' && w >= 0) return w;
  }
  const border = numbers(file.get(annot, 'Border'));
  if (border.length >= 3 && border[2]! >= 0) return border[2]!;
  return DEFAULT_BORDER_PT;
}

/**
 * §12.5.6.2 — an annotation colour is 1, 3 or 4 numbers (grey, RGB, CMYK), and
 * an EMPTY array means no colour at all. Returned as the three operands `rg`
 * and `RG` both take.
 */
function colorOf(value: PdfValue | undefined): string | undefined {
  const n = numbers(value);
  if (n.length === 1) return `${num(n[0]!)} ${num(n[0]!)} ${num(n[0]!)}`;
  if (n.length === 3) return n.map((v) => num(v)).join(' ');
  if (n.length === 4) {
    // §10.4.2.3 — the naive conversion, which is what a viewer with no colour
    // profile to hand does as well.
    const [c, m, y, k] = n as [number, number, number, number];
    return [num((1 - c) * (1 - k)), num((1 - m) * (1 - k)), num((1 - y) * (1 - k))].join(' ');
  }
  return undefined;
}

/** The arrays inside an array — `/InkList` is one run of points per stroke. */
function listOfLists(value: PdfValue | undefined): Array<Array<number>> {
  if (!Array.isArray(value)) return [];
  return value.map((v) => numbers(v));
}

/** The numeric members of an array value, or an empty list. */
function numbers(value: PdfValue | undefined): Array<number> {
  if (!Array.isArray(value)) return [];
  return value.filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
}

/** A four-number array as an ordered rectangle, or `undefined`. */
function rectangle(v: PdfValue | undefined): [number, number, number, number] | undefined {
  const n = numbers(v);
  if (n.length < 4) return undefined;
  return [
    Math.min(n[0]!, n[2]!),
    Math.min(n[1]!, n[3]!),
    Math.max(n[0]!, n[2]!),
    Math.max(n[1]!, n[3]!),
  ];
}

/** The rectangle pulled in on every side, never past its own middle. */
function inset(
  r: readonly [number, number, number, number],
  by: number,
): [number, number, number, number] {
  const dx = Math.min(by, (r[2] - r[0]) / 2);
  const dy = Math.min(by, (r[3] - r[1]) / 2);
  return [r[0] + dx, r[1] + dy, r[2] - dx, r[3] - dy];
}

/** Three decimals is finer than any pen, and keeps the stream readable. */
function num(v: number): string {
  return String(Math.round(v * 1000) / 1000);
}
