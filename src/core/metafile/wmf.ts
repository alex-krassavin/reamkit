// MS-WMF — read a Windows Metafile into {@link MetaPicture}.
//
// The older sibling of EMF: the same device context, sixteen-bit coordinates,
// and records that name their function in a word rather than a longword. The
// object table is positional (a created object takes the first free slot),
// which is the one structural difference the reader has to keep in mind.
//
// Coordinates come out in the metafile's LOGICAL units, and the box is the
// window the file sets — a WMF has no device frame of its own. A placeable
// header (the `\xd7\xcd\xc6\x9a` one Word writes) states the box up front and
// is used when it is there.

import type { PathSegment, StrokeStyle, VectorPath } from '@/core/vector';
import type { DeviceContext, MetaObject, MetaPicture, PicturePrim } from '@/core/metafile/picture';
import { PathBuilder } from '@/core/vector';
import { cloneDc, colorRef, newDeviceContext } from '@/core/metafile/picture';
import { cropDib, readDib } from '@/core/metafile/dib';
import { makeBlitter } from '@/core/metafile/blit';
import { ellipseSegments } from '@/core/metafile/emf';
import { fromSymbolFont, symbolGeometryOf, symbolOutline } from '@/core/metafile/symbol-fonts';

const PLACEABLE_KEY = 0x9ac6cdd7;

/** Whether the bytes open with a WMF header — placeable or bare (§2.3.2.2). */
export function isWmf(bytes: Uint8Array): boolean {
  if (bytes.length < 18) return false;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (v.getUint32(0, true) === PLACEABLE_KEY) return true;
  const type = v.getUint16(0, true);
  const headerWords = v.getUint16(2, true);
  return (type === 1 || type === 2) && headerWords === 9;
}

const RECORD_NAMES: ReadonlyMap<number, string> = new Map([
  [0x0922, 'pie'],
  [0x0830, 'chord'],
  [0x0817, 'arc'],
  [0x0419, 'flood fill'],
  [0x0548, 'flood fill'],
  [0x0f7, 'palette'],
  [0x0626, 'escape'],
  [0x02fd, 'region'],
  [0x0228, 'region fill'],
  [0x0429, 'region frame'],
  [0x012a, 'region invert'],
  [0x012b, 'region paint'],
]);

/**
 * Read a WMF into its primitives.
 *
 * @throws Error when the bytes are not a WMF.
 */
export function readWmf(bytes: Uint8Array): MetaPicture {
  if (!isWmf(bytes)) throw new Error('WMF: not a Windows metafile');
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const placeable = v.getUint32(0, true) === PLACEABLE_KEY;
  // §2.3.2.3 — the placeable header's bounding box, in logical units.
  const frame = placeable
    ? {
        left: v.getInt16(6, true),
        top: v.getInt16(8, true),
        right: v.getInt16(10, true),
        bottom: v.getInt16(12, true),
      }
    : undefined;

  const prims: Array<PicturePrim> = [];
  const skipped = new Set<string>();
  // §3.1.4.1 — objects live in a table by INDEX; a create takes the lowest
  // free one, which is why the reader has to track the holes a delete leaves.
  const objects: Array<MetaObject | undefined> = [];
  const put = (obj: MetaObject): void => {
    const free = objects.indexOf(undefined);
    if (free >= 0) objects[free] = obj;
    else objects.push(obj);
  };
  const stack: Array<DeviceContext> = [];
  let dc = newDeviceContext();
  let win = { x: 0, y: 0, cx: 0, cy: 0 };
  // §2.3.5.13 — SETWINDOWORG moves the frame the records AFTER it are drawn in,
  // and a metafile may move it many times. The box below can describe only one
  // frame, so every primitive is re-expressed in the FIRST one as it is read.
  // Kept as the last window instead, 41246-2's coloured bands were laid out
  // against an origin set long after they were drawn and came off the slide's
  // left edge.
  let origin: { x: number; y: number } | undefined;
  const emit = (prim: PicturePrim): void => {
    origin ??= { x: win.x, y: win.y };
    const dx = origin.x - win.x;
    const dy = origin.y - win.y;
    prims.push(dx === 0 && dy === 0 ? prim : shiftPrim(prim, dx, dy));
  };

  const strokeOf = (): StrokeStyle | undefined => {
    if (dc.pen.style === 'none') return undefined;
    const w = Math.max(1, dc.pen.widthLu);
    return {
      colorHex: dc.pen.colorHex,
      widthPt: w,
      ...(dc.pen.style === 'solid'
        ? {}
        : { dash: dc.pen.style === 'dot' ? [w, 2 * w] : [4 * w, 3 * w] }),
    };
  };

  const paint = (segments: ReadonlyArray<PathSegment>, fill: boolean, stroke: boolean): void => {
    if (segments.length === 0) return;
    const path: VectorPath = { segments, fillRule: dc.fillRule };
    const st = stroke ? strokeOf() : undefined;
    const fillHex = fill && !dc.brush.hollow ? dc.brush.colorHex : undefined;
    if (fillHex === undefined && !st) return;
    emit({
      kind: 'path',
      paths: [path],
      ...(fillHex !== undefined ? { fillColorHex: fillHex } : {}),
      ...(st ? { stroke: st } : {}),
    });
  };

  const rectSegments = (l: number, t: number, r: number, b: number): Array<PathSegment> =>
    new PathBuilder().moveTo(l, t).lineTo(r, t).lineTo(r, b).lineTo(l, b).close().build()
      .segments as Array<PathSegment>;

  const blitter = makeBlitter(emit);
  /**
   * §2.3.1 — one blit record: the packed bitmap the record ends with, into the
   * destination rectangle its fields name. A WMF's coordinates are its own
   * logical units, which is the frame the primitives are already in.
   */
  const blit = (o: {
    dibAt: number;
    dest: { x: number; y: number; w: number; h: number };
    src: { x: number; y: number; w: number; h: number };
    rop: number;
  }): void => {
    const dib = readDib(bytes, off + o.dibAt);
    if (!dib) {
      skipped.add('bitmap');
      return;
    }
    // The source rectangle is stated in the BITMAP's own coordinates, which run
    // from its first row — the bottom one, the way a DIB is usually stored.
    const s = o.src;
    const part =
      s.x !== 0 || s.y !== 0 || s.w !== dib.width || s.h !== dib.height
        ? cropDib(dib, s.x, dib.bottomUp === true ? dib.height - s.y - s.h : s.y, s.w, s.h)
        : dib;
    blitter.blit(
      part,
      {
        x: Math.min(o.dest.x, o.dest.x + o.dest.w),
        y: Math.min(o.dest.y, o.dest.y + o.dest.h),
        width: Math.abs(o.dest.w),
        height: Math.abs(o.dest.h),
      },
      o.rop,
    );
  };

  let off = placeable ? 22 : 0;
  off += 18; // the META header itself
  let guard = 0;
  while (off + 6 <= bytes.length && guard++ < 500_000) {
    const size = v.getUint32(off, true); // in WORDS, the record included
    const fn = v.getUint16(off + 4, true);
    if (size < 3 || off + size * 2 > bytes.length) break;
    // Parameters start after the size and the function, and a record's own
    // words are its parameters in REVERSE for the coordinate records.
    const p = (i: number): number => v.getInt16(off + 6 + i * 2, true);
    const pu = (i: number): number => v.getUint16(off + 6 + i * 2, true);
    const p32 = (i: number): number => v.getUint32(off + 6 + i * 2, true);

    switch (fn) {
      case 0x0000: // META_EOF
        off = bytes.length;
        continue;
      case 0x020b: // META_SETWINDOWORG — y then x
        win = { ...win, x: p(1), y: p(0) };
        break;
      case 0x020c: // META_SETWINDOWEXT
        win = { ...win, cx: p(1), cy: p(0) };
        break;
      case 0x001e: // META_SAVEDC
        stack.push(cloneDc(dc));
        break;
      case 0x0127: // META_RESTOREDC
        {
          const back = Math.max(1, -p(0));
          for (let i = 0; i < back && stack.length > 0; i++) dc = stack.pop()!;
        }
        break;
      case 0x012d: // META_SELECTOBJECT
        {
          const obj = objects[pu(0)];
          if (obj?.kind === 'pen') dc.pen = obj;
          else if (obj?.kind === 'brush') dc.brush = obj;
          else if (obj?.kind === 'font') dc.font = obj;
        }
        break;
      case 0x01f0: // META_DELETEOBJECT
        objects[pu(0)] = undefined;
        break;
      // §2.3.4 — the objects the reader does not model still take a slot, or
      // every handle created after one of them points at the wrong object.
      case 0x00f7: // META_CREATEPALETTE
      case 0x0142: // META_DIBCREATEPATTERNBRUSH
      case 0x01f9: // META_CREATEPATTERNBRUSH
      case 0x06ff: // META_CREATEREGION
        put({ kind: 'other' });
        break;
      case 0x02fa: // META_CREATEPENINDIRECT — style, width (POINT16), colour
        put({
          kind: 'pen',
          style: penStyle(pu(0)),
          widthLu: p(1),
          colorHex: colorRef(p32(3)),
        });
        break;
      case 0x02fc: // META_CREATEBRUSHINDIRECT — style, colour, hatch
        put({
          kind: 'brush',
          colorHex: colorRef(p32(1)),
          hollow: pu(0) === 1,
        });
        break;
      case 0x02fb: // META_CREATEFONTINDIRECT
        put(readLogFont(v, bytes, off + 6));
        break;
      case 0x0102: // META_SETBKMODE
        dc.bkOpaque = pu(0) === 2;
        break;
      case 0x0106: // META_SETPOLYFILLMODE
        dc.fillRule = pu(0) === 2 ? 'nonzero' : 'evenodd';
        break;
      case 0x012e: // META_SETTEXTALIGN
        {
          const f = pu(0);
          dc.alignH = (f & 6) === 6 ? 'center' : (f & 2) === 2 ? 'right' : 'left';
          dc.alignBaseline = (f & 24) === 24;
        }
        break;
      case 0x0209: // META_SETTEXTCOLOR
        dc.textColorHex = colorRef(p32(0));
        break;
      case 0x0201: // META_SETBKCOLOR
        dc.bkColorHex = colorRef(p32(0));
        break;
      case 0x0214: // META_MOVETO — y then x
        dc.x = p(1);
        dc.y = p(0);
        break;
      case 0x0213: // META_LINETO
        {
          const x = p(1);
          const y = p(0);
          paint(
            [
              { op: 'move', x: dc.x, y: dc.y },
              { op: 'line', x, y },
            ],
            false,
            true,
          );
          dc.x = x;
          dc.y = y;
        }
        break;
      case 0x041b: // META_RECTANGLE — bottom, right, top, left
        paint(rectSegments(p(3), p(2), p(1), p(0)), true, true);
        break;
      case 0x0418: // META_ELLIPSE
        paint(
          ellipseSegments((x, y) => ({ x, y }), p(3), p(2), p(1), p(0)),
          true,
          true,
        );
        break;
      case 0x061c: // META_ROUNDRECT — corner size first, then the box
        paint(rectSegments(p(5), p(4), p(3), p(2)), true, true);
        break;
      case 0x061d: // META_PATBLT — the destination filled with the brush, which
        // is how a WMF paints a rule, a bar or a panel (5139 of them in the
        // corpus, more than any other drawing record).
        {
          const rop = p32(0);
          const h = p(2);
          const w = p(3);
          const y = p(4);
          const x = p(5);
          const hex = rop === 0x00000042 ? '000000' : rop === 0x00ff0062 ? 'FFFFFF' : undefined;
          const brush = dc.brush;
          if (hex !== undefined) dc.brush = { kind: 'brush', colorHex: hex, hollow: false };
          paint(rectSegments(x, y, x + w, y + h), true, false);
          dc.brush = brush;
        }
        break;
      // §2.3.1.2 / §2.3.1.3 / §2.3.1.5 — the three records that carry a packed
      // bitmap. Each states its raster operation first and the bitmap last, and
      // each may leave the bitmap OUT when its operation has no use for a
      // source: the record is then the brush painting the destination, exactly
      // as META_PATBLT is (0x00F00021 PATCOPY, 288 records of the corpus).
      case 0x0940: // META_DIBBITBLT
      case 0x0b41: // META_DIBSTRETCHBLT
      case 0x0f43: // META_STRETCHDIB
        {
          const rop = p32(0);
          // Where each record keeps its fields, as parameter words. A stretch
          // names the source extent it scales from; a plain blit takes the
          // destination's. META_STRETCHDIB spends one word more on what its
          // colour table means, which moves everything after it along.
          const f =
            fn === 0x0940
              ? { dib: 22, srcH: 4, srcW: 5, ySrc: 2, xSrc: 3, h: 4, w: 5, y: 6, x: 7 }
              : fn === 0x0b41
                ? { dib: 26, srcH: 2, srcW: 3, ySrc: 4, xSrc: 5, h: 6, w: 7, y: 8, x: 9 }
                : { dib: 28, srcH: 3, srcW: 4, ySrc: 5, xSrc: 6, h: 7, w: 8, y: 9, x: 10 };
          if (f.dib + 12 > size * 2) {
            // No bitmap: the brush (or the operation's own colour) fills the
            // destination. A META_DIBBITBLT then carries a reserved word ahead
            // of its coordinates that the form with a bitmap does not — the
            // record is 12 words, one more than its fields account for, and
            // read without the shift it draws a rectangle of no height.
            const hex = rop === 0x00000042 ? '000000' : rop === 0x00ff0062 ? 'FFFFFF' : undefined;
            const brush = dc.brush;
            if (hex !== undefined) dc.brush = { kind: 'brush', colorHex: hex, hollow: false };
            const shift = fn === 0x0940 ? 1 : 0;
            const [h, w, y, x] = [p(f.h + shift), p(f.w + shift), p(f.y + shift), p(f.x + shift)];
            paint(rectSegments(x, y, x + w, y + h), true, false);
            dc.brush = brush;
            break;
          }
          blit({
            dibAt: f.dib,
            dest: { x: p(f.x), y: p(f.y), w: p(f.w), h: p(f.h) },
            src: { x: p(f.xSrc), y: p(f.ySrc), w: p(f.srcW), h: p(f.srcH) },
            rop,
          });
        }
        break;
      case 0x0324: // META_POLYGON
      case 0x0325: // META_POLYLINE
        {
          const count = pu(0);
          const pts = readPoints(v, off + 8, count);
          const closed = fn === 0x0324;
          paint(polySegments(pts, closed), closed, true);
        }
        break;
      case 0x0538: // META_POLYPOLYGON
        {
          const polys = pu(0);
          const counts: Array<number> = [];
          for (let i = 0; i < polys; i++) counts.push(v.getUint16(off + 8 + i * 2, true));
          let at = off + 8 + polys * 2;
          const segs: Array<PathSegment> = [];
          for (const c of counts) {
            segs.push(...polySegments(readPoints(v, at, c), true));
            at += c * 4;
          }
          paint(segs, true, true);
        }
        break;
      case 0x0521: // META_TEXTOUT — count, string, y, x
        {
          const count = pu(0);
          const text = readString(bytes, off + 8, count, false);
          const words = (count + 1) >> 1;
          pushText(emit, dc, text, p(1 + words + 1), p(1 + words));
        }
        break;
      case 0x0a32: // META_EXTTEXTOUT — y, x, count, options, [rect], string
        {
          const y = p(0);
          const x = p(1);
          const count = pu(2);
          const options = pu(3);
          const hasRect = (options & 0x0006) !== 0;
          const text = readString(bytes, off + 6 + (4 + (hasRect ? 4 : 0)) * 2, count, false);
          pushText(emit, dc, text, x, y);
        }
        break;
      default:
        {
          const name = RECORD_NAMES.get(fn);
          if (name) skipped.add(name);
        }
        break;
    }
    off += size * 2;
  }

  const box = frame ?? {
    left: win.x,
    top: win.y,
    right: win.x + (win.cx || 1000),
    bottom: win.y + (win.cy || 1000),
  };
  return {
    left: box.left,
    top: box.top,
    width: Math.max(1, Math.abs(box.right - box.left)),
    height: Math.max(1, Math.abs(box.bottom - box.top)),
    prims,
    skipped: [...skipped],
  };
}

function pushText(
  emit: (prim: PicturePrim) => void,
  dc: DeviceContext,
  text: string,
  x: number,
  y: number,
): void {
  if (text.trim() === '') return;
  const em = Math.abs(dc.font?.heightLu ?? 12);
  // A symbol font's letters are not letters: Webdings `n` is a filled circle,
  // and no substitute font has one either — so the plain shapes are DRAWN.
  const drawn = symbolPrims(text, dc, x, y, em);
  if (drawn) {
    for (const prim of drawn) emit(prim);
    return;
  }
  emit({
    kind: 'text',
    // The rest are translated to the Unicode that means the same thing.
    text: fromSymbolFont(text, dc.font?.family),
    x,
    y,
    alignH: dc.alignH,
    alignBaseline: dc.alignBaseline,
    sizeLu: em,
    colorHex: dc.textColorHex,
    ...(dc.font?.family ? { fontFamily: dc.font.family } : {}),
    ...(dc.font?.bold ? { bold: true } : {}),
    ...(dc.font?.italic ? { italic: true } : {}),
    ...(dc.font?.escapement ? { escapement: dc.font.escapement } : {}),
  });
}

// A primitive moved into another frame: the metafile's own units, y down.
function shiftPrim(prim: PicturePrim, dx: number, dy: number): PicturePrim {
  if (prim.kind === 'text') return { ...prim, x: prim.x + dx, y: prim.y + dy };
  if (prim.kind === 'image') return { ...prim, x: prim.x + dx, y: prim.y + dy };
  return {
    ...prim,
    paths: prim.paths.map((path) => ({
      ...path,
      segments: path.segments.map((sg) =>
        'x' in sg
          ? {
              ...sg,
              x: sg.x + dx,
              y: sg.y + dy,
              ...('x1' in sg ? { x1: sg.x1 + dx, y1: sg.y1 + dy } : {}),
              ...('x2' in sg ? { x2: sg.x2 + dx, y2: sg.y2 + dy } : {}),
            }
          : sg,
      ),
    })),
  };
}

function penStyle(style: number): (MetaObject & { kind: 'pen' })['style'] {
  switch (style & 0xf) {
    case 1:
      return 'dash';
    case 2:
      return 'dot';
    case 3:
      return 'dashdot';
    case 4:
      return 'dashdotdot';
    case 5:
      return 'none';
    default:
      return 'solid';
  }
}

// §2.2.2.9 Font — height, width, escapement, orientation, weight, then the
// byte flags and a 32-byte face name.
function readLogFont(v: DataView, bytes: Uint8Array, at: number): MetaObject {
  const height = v.getInt16(at, true);
  const escapement = v.getInt16(at + 4, true);
  const weight = v.getInt16(at + 8, true);
  const italic = bytes[at + 10] !== 0;
  const face = readString(bytes, at + 18, 32, false).replace(/\0.*$/u, '');
  return {
    kind: 'font',
    heightLu: height,
    ...(face ? { family: face } : {}),
    ...(weight >= 600 ? { bold: true } : {}),
    ...(italic ? { italic: true } : {}),
    ...(escapement ? { escapement } : {}),
  };
}

function readString(bytes: Uint8Array, at: number, count: number, wide: boolean): string {
  let out = '';
  for (let i = 0; i < count; i++) {
    const o = at + (wide ? i * 2 : i);
    if (o >= bytes.length) break;
    const code = wide ? bytes[o]! | ((bytes[o + 1] ?? 0) << 8) : bytes[o]!;
    if (code === 0) break;
    out += String.fromCharCode(code);
  }
  return out;
}

function readPoints(v: DataView, at: number, count: number): Array<{ x: number; y: number }> {
  const out: Array<{ x: number; y: number }> = [];
  for (let i = 0; i < count; i++) {
    const o = at + i * 4;
    if (o + 4 > v.byteLength) break;
    out.push({ x: v.getInt16(o, true), y: v.getInt16(o + 2, true) });
  }
  return out;
}

function polySegments(
  pts: ReadonlyArray<{ x: number; y: number }>,
  closed: boolean,
): Array<PathSegment> {
  if (pts.length === 0) return [];
  const segs: Array<PathSegment> = [{ op: 'move', x: pts[0]!.x, y: pts[0]!.y }];
  for (let i = 1; i < pts.length; i++) segs.push({ op: 'line', x: pts[i]!.x, y: pts[i]!.y });
  if (closed) segs.push({ op: 'close' });
  return segs;
}

/**
 * The symbols a text record draws, as filled outlines, or `undefined` when the
 * string is not one the shapes cover.
 *
 * Both fonts advance one em per symbol and fill that em, sitting from 0.2 em
 * below the baseline to 0.8 em above it — measured off Webdings itself.
 *
 * @param text The string as stored.
 * @param dc   The device context (font, colour, alignment).
 * @param x    The reference point's x, in logical units.
 * @param y    Its y — the text's top unless the context says baseline.
 * @param em   The font's height in logical units.
 * @returns One filled path per character, or `undefined`.
 */
function symbolPrims(
  text: string,
  dc: DeviceContext,
  x: number,
  y: number,
  em: number,
): Array<PicturePrim> | undefined {
  const shapes = symbolGeometryOf(text, dc.font?.family);
  if (!shapes) return undefined;
  const width = em * shapes.length;
  const left = dc.alignH === 'center' ? x - width / 2 : dc.alignH === 'right' ? x - width : x;
  // The em box the glyph fills: below the baseline by a fifth, or straight down
  // from the top when the point names the top.
  const top = dc.alignBaseline ? y - em * 0.8 : y;
  return shapes.map((shape, i) => ({
    kind: 'path' as const,
    paths: [symbolOutline(shape, left + i * em, top, em)],
    fillColorHex: dc.textColorHex,
  }));
}
