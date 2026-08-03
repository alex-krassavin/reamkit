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
import { ellipseSegments } from '@/core/metafile/emf';

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
  [0x0940, 'bitmap'],
  [0x0b41, 'bitmap'],
  [0x0f43, 'bitmap'],
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
    prims.push({
      kind: 'path',
      paths: [path],
      ...(fillHex !== undefined ? { fillColorHex: fillHex } : {}),
      ...(st ? { stroke: st } : {}),
    });
  };

  const rectSegments = (l: number, t: number, r: number, b: number): Array<PathSegment> =>
    new PathBuilder().moveTo(l, t).lineTo(r, t).lineTo(r, b).lineTo(l, b).close().build()
      .segments as Array<PathSegment>;

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
          pushText(prims, dc, text, p(1 + words + 1), p(1 + words));
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
          pushText(prims, dc, text, x, y);
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
  prims: Array<PicturePrim>,
  dc: DeviceContext,
  text: string,
  x: number,
  y: number,
): void {
  if (text.trim() === '') return;
  prims.push({
    kind: 'text',
    text,
    x,
    y,
    alignH: dc.alignH,
    alignBaseline: dc.alignBaseline,
    sizeLu: Math.abs(dc.font?.heightLu ?? 12),
    colorHex: dc.textColorHex,
    ...(dc.font?.family ? { fontFamily: dc.font.family } : {}),
    ...(dc.font?.bold ? { bold: true } : {}),
    ...(dc.font?.italic ? { italic: true } : {}),
    ...(dc.font?.escapement ? { escapement: dc.font.escapement } : {}),
  });
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
