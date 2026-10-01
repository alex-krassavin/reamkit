// §12.5.5 — annotation appearance streams.
//
// A form field draws nothing in the page's content stream. Its look — the tint
// behind it, its border, the value typed into it — lives in the widget
// annotation's own appearance stream, a form XObject the viewer paints after
// the page. 160F-2019.pdf carries seventy-six of them, which is every shaded
// box on the certificate, and reading only the page content gave the grid with
// no fields in it.
//
// The stream draws in its own space and has to be fitted to the annotation's
// `/Rect`: the algorithm §12.5.5 states is to transform the `/BBox` by the
// form's `/Matrix`, take the bounding box of what comes out, and map THAT onto
// the rectangle.

import { IDENTITY, multiply } from './content';
import { drawnAppearance, drawnResources, textMarkupOf } from './annot-draw';
import type { Matrix } from './content';
import type { PdfDict, PdfValue } from '@/pdf/objects';
import type { PdfFile, PdfPage } from './document';
import { PDF_NULL, PdfName, PdfStream } from '@/pdf/objects';

/** One annotation's normal appearance, ready to interpret in page space. */
export interface Appearance {
  readonly stream: PdfStream;
  /** Maps the appearance's own space onto the page (§12.5.5). */
  readonly ctm: Matrix;
  /** The appearance's `/Resources`, when it states its own. */
  readonly resources: PdfDict | undefined;
  /** §8.10.1 — the form's `/BBox`, in its own space, which clips what it paints. */
  readonly bbox?: readonly [number, number, number, number];
}

/**
 * §8.10.1 — an appearance's content as it paints: clipped to its `/BBox`.
 *
 * The box is the form's clip, not a hint. A form-filling tool that wrote its
 * field backgrounds in PAGE coordinates inside a box forty points wide paints
 * them nowhere any viewer shows — and painted anyway, with the fill colour it
 * never set, bug1669099.pdf's form came back with three black slabs over its
 * letterhead and its terms of payment, each twice as far across the sheet as
 * the field it belonged to. The clip is installed the way the content would
 * install one, so every pass that reads the appearance sees the same page.
 *
 * @param file       The owning file.
 * @param appearance The appearance.
 * @returns Its content, preceded by the clip to its box where it has one.
 */
export function appearanceContent(file: PdfFile, appearance: Appearance): Uint8Array {
  const data = file.streamData(appearance.stream);
  const box = appearance.bbox;
  if (!box) return data;
  const [x0, y0, x1, y1] = box;
  const clip = new TextEncoder().encode(
    `${String(x0)} ${String(y0)} ${String(x1 - x0)} ${String(y1 - y0)} re W n\n`,
  );
  const out = new Uint8Array(clip.length + data.length);
  out.set(clip, 0);
  out.set(data, clip.length);
  return out;
}

/** §12.5.3 `/F` — the annotation is not painted at all. */
const FLAG_HIDDEN = 2;
const FLAG_NOVIEW = 32;

/**
 * Every annotation appearance the page shows, in `/Annots` order — which is
 * the order they paint in, over the page's own content.
 *
 * A `Popup` is a note's window and is never part of the page. An annotation
 * flagged Hidden or NoView paints nothing. An `/AP` `/N` may be a stream or a
 * dictionary of states, in which case `/AS` names the one in force. An
 * annotation with no appearance at all gets one written for it from its own
 * geometry, where its subtype says exactly what that is ({@link drawnAppearance}).
 *
 * @param file The owning file, for resolving references.
 * @param page The page whose annotations are wanted.
 * @returns The appearances, each with the matrix that places it on the page.
 */
/** Whether the page's own content stream shows any glyphs at all. */
function pageHasText(file: PdfFile, page: PdfPage): boolean {
  const had = textful.get(page.dict);
  if (had !== undefined) return had;
  // Cheap and sufficient: a page that never opens a text object shows no
  // words, and one that does is not a page a markup annotation is alone on.
  const content = file.pageContent(page);
  let found = false;
  for (let i = 0; i + 1 < content.length && !found; i++) {
    found = content[i] === 0x42 && content[i + 1] === 0x54; // "BT"
  }
  textful.set(page.dict, found);
  return found;
}

/** One answer per page: the collector is called from three passes. */
const textful = new WeakMap<PdfDict, boolean>();

export function collectPageAppearances(file: PdfFile, page: PdfPage): Array<Appearance> {
  const annots = file.get(page.dict, 'Annots');
  if (!Array.isArray(annots)) return [];
  const out: Array<Appearance> = [];
  const regenerate = needsAppearances(file);
  for (const entry of annots) {
    const annot = file.resolve(entry);
    if (!(annot instanceof Map)) continue;
    const subtype = file.get(annot, 'Subtype');
    if (subtype instanceof PdfName && subtype.value === 'Popup') continue;
    // A text-markup annotation marks WORDS, and its band comes back on the
    // runs rather than as a mark on the paper (see `textMarkupOf`). Painting
    // its appearance too would lay the band over the words a second time,
    // where the words no longer are.
    //
    // Unless the page has no words at all: then nothing carries the mark and
    // skipping it leaves the page blank. bug1538111.pdf is four markup
    // annotations over an empty page and reconstructed to nothing.
    if (textMarkupOf(file, annot) && pageHasText(file, page)) continue;
    const flags = file.get(annot, 'F');
    if (typeof flags === 'number' && (flags & FLAG_HIDDEN || flags & FLAG_NOVIEW)) continue;

    const stream = normalAppearance(file, annot);
    // §12.7.2 `/NeedAppearances` — the form says its appearances are not to be
    // trusted and the viewer is to build them from the fields' values. A text
    // field's variable text is its value (§12.7.3.3); the rest of what the
    // stale appearance draws — the frame, the tint — stands.
    const typed = regenerate && isWidget(file, annot) ? drawnAppearance(file, annot) : undefined;
    if (typed && typedField(file, annot)) {
      const kept = stream ? withoutVariableText(file, stream) : undefined;
      const rect = rectangle(file.get(annot, 'Rect'));
      if (kept && rect) out.push(fitted(file, kept, rect));
      out.push({ stream: typed, ctm: IDENTITY, resources: drawnResources(file, annot) });
      continue;
    }
    if (!stream) {
      // §12.5.5 — no appearance to paint, so one is written from the geometry
      // the annotation states. Already in page space, so it is placed as it is.
      const drawn = drawnAppearance(file, annot);
      if (drawn) out.push({ stream: drawn, ctm: IDENTITY, resources: drawnResources(file, annot) });
      continue;
    }
    const rect = rectangle(file.get(annot, 'Rect'));
    if (!rect) continue;
    out.push(fitted(file, stream, rect));
  }
  return out;
}

/** An appearance stream placed on its annotation's `/Rect` (§12.5.5). */
function fitted(
  file: PdfFile,
  stream: PdfStream,
  rect: readonly [number, number, number, number],
): Appearance {
  const matrix = matrixOf(file, stream.dict);
  const bbox = rectangle(file.get(stream.dict, 'BBox'));
  const resources = file.get(stream.dict, 'Resources');
  return {
    stream,
    ctm: multiply(matrix, fitToRect(bbox, matrix, rect)),
    resources: resources instanceof Map ? resources : undefined,
    ...(bbox ? { bbox } : {}),
  };
}

/** §12.7.2 — whether the form asks the viewer to build its fields' appearances. */
function needsAppearances(file: PdfFile): boolean {
  const acro = file.get(file.catalog, 'AcroForm');
  return acro instanceof Map && file.get(acro, 'NeedAppearances') === true;
}

function isWidget(file: PdfFile, annot: PdfDict): boolean {
  const subtype = file.get(annot, 'Subtype');
  return subtype instanceof PdfName && subtype.value === 'Widget';
}

/** Whether a widget belongs to a text field — `/FT /Tx`, on it or up its field tree. */
function typedField(file: PdfFile, annot: PdfDict): boolean {
  let node: PdfDict | undefined = annot;
  for (let depth = 0; node && depth < 32; depth++) {
    const type = file.get(node, 'FT');
    if (type instanceof PdfName) return type.value === 'Tx';
    const parent = file.get(node, 'Parent');
    node = parent instanceof Map ? parent : undefined;
  }
  return false;
}

/**
 * §12.7.3.3 — an appearance with its variable text taken out: the part it
 * marks `/Tx BMC … EMC`, which is what a viewer rebuilds from the value.
 * bug1844583.pdf's stale appearance still reads "Dlrow Olleh" there, over a
 * field whose value is "Hello World".
 */
function withoutVariableText(file: PdfFile, stream: PdfStream): PdfStream {
  const text = new TextDecoder('latin1').decode(file.streamData(stream));
  // No marked span ends past the last `EMC`; searched for in the whole stream,
  // a `/Tx BMC` that none closes read on to the end of it from every one.
  const end = text.lastIndexOf('EMC') + 'EMC'.length;
  const stripped = text.slice(0, end).replace(/\/Tx\s+BMC[\s\S]*?EMC/gu, '') + text.slice(end);
  // The bytes come back decoded, so the filters that encoded them go too.
  const dict = new Map(stream.dict);
  dict.delete('Filter');
  dict.delete('DecodeParms');
  return new PdfStream(dict, Uint8Array.from([...stripped].map((c) => c.charCodeAt(0))));
}

/** §12.5.5 `/AP` `/N` — the normal appearance, through `/AS` when it is a set. */
function normalAppearance(file: PdfFile, annot: PdfDict): PdfStream | undefined {
  const ap = file.get(annot, 'AP');
  if (!(ap instanceof Map)) return undefined;
  const normal = file.get(ap, 'N');
  if (normal instanceof PdfStream) return normal;
  if (!(normal instanceof Map)) return undefined;
  const state = file.get(annot, 'AS');
  if (state instanceof PdfName) {
    // The state in force, and only it. A set that does not carry the named
    // state draws NOTHING — a check box whose author drew only the tick has
    // nothing to draw when it is clear, and drawing the tick anyway ticks
    // every box on the form. annotation-button-widget.pdf is three boxes and
    // six radio buttons of which one box and one button are set; it came back
    // with all nine filled in.
    const picked = file.resolve(normal.get(state.value) ?? PDF_NULL);
    return picked instanceof PdfStream ? picked : undefined;
  }
  // No `/AS` at all: a set of one is unambiguous anyway.
  const only = [...normal.values()]
    .map((v) => file.resolve(v))
    .filter((v) => v instanceof PdfStream);
  return only.length === 1 ? only[0] : undefined;
}

/**
 * §12.5.5 — the matrix taking the form's `/Matrix`-transformed `/BBox` onto the
 * annotation's `/Rect`: the two are fitted corner to corner, so an appearance
 * authored at any size lands exactly in the rectangle that owns it.
 */
function fitToRect(
  bbox: readonly [number, number, number, number] | undefined,
  matrix: Matrix,
  rect: readonly [number, number, number, number],
): Matrix {
  if (!bbox) return [1, 0, 0, 1, rect[0], rect[1]];
  // The four corners through the matrix, then their bounding box.
  const xs: Array<number> = [];
  const ys: Array<number> = [];
  for (const [x, y] of [
    [bbox[0], bbox[1]],
    [bbox[2], bbox[1]],
    [bbox[2], bbox[3]],
    [bbox[0], bbox[3]],
  ] as const) {
    xs.push(matrix[0] * x + matrix[2] * y + matrix[4]);
    ys.push(matrix[1] * x + matrix[3] * y + matrix[5]);
  }
  const bw = Math.max(...xs) - Math.min(...xs);
  const bh = Math.max(...ys) - Math.min(...ys);
  const sx = bw > 0 ? (rect[2] - rect[0]) / bw : 1;
  const sy = bh > 0 ? (rect[3] - rect[1]) / bh : 1;
  return [sx, 0, 0, sy, rect[0] - Math.min(...xs) * sx, rect[1] - Math.min(...ys) * sy];
}

/** §8.10.2 `/Matrix`, or the identity when the form states none. */
function matrixOf(file: PdfFile, dict: PdfDict): Matrix {
  const m = file.resolve(dict.get('Matrix') ?? PDF_NULL);
  if (!Array.isArray(m) || m.length !== 6) return IDENTITY;
  const n = m.map((v) => {
    const r: PdfValue = file.resolve(v);
    return typeof r === 'number' ? r : 0;
  });
  return [n[0]!, n[1]!, n[2]!, n[3]!, n[4]!, n[5]!];
}

/** A four-number array as an ordered rectangle, or `undefined`. */
function rectangle(v: PdfValue | undefined): [number, number, number, number] | undefined {
  if (!Array.isArray(v) || v.length < 4) return undefined;
  const n = v.slice(0, 4).map((x) => (typeof x === 'number' ? x : NaN));
  if (n.some((x) => !Number.isFinite(x))) return undefined;
  return [
    Math.min(n[0]!, n[2]!),
    Math.min(n[1]!, n[3]!),
    Math.max(n[0]!, n[2]!),
    Math.max(n[1]!, n[3]!),
  ];
}
