// MS-EMF — read an Enhanced Metafile into {@link MetaPicture}.
//
// An EMF is a list of records played through a device context: a pen, a brush,
// a font, a current point and a logical→device mapping. This reads the records
// that DRAW (paths, rectangles, text, bitmaps, brush blits) and the ones that
// set the state they draw through; the rest — palettes, colour management,
// regions — are named in `skipped`, which the reader exposes as a diagnostic:
// the render path it is read from carries no loss channel of its own, so
// nothing turns it into a reported Loss automatically.
//
// Coordinates come out in DEVICE units: the window→viewport mapping (§2.3.11)
// is applied as the records go by, so the picture's box is the header's own
// `rclBounds` and nothing downstream needs the mapping again.

import type { PathSegment, StrokeStyle, VectorPath } from '@/core/vector';
import type { DeviceContext, MetaObject, MetaPicture, PicturePrim } from '@/core/metafile/picture';
import { PathBuilder } from '@/core/vector';
import {
  applyTransform,
  clippedAway,
  cloneDc,
  colorRef,
  intersectClip,
  newDeviceContext,
  primBounds,
} from '@/core/metafile/picture';
import { cropDib, fadedDib, readDib } from '@/core/metafile/dib';
import { makeBlitter } from '@/core/metafile/blit';
import { fromSymbolFont } from '@/core/metafile/symbol-fonts';

/** Whether the bytes open with an EMF header (MS-EMF §2.3.4.2: type 1 + " EMF"). */
export function isEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 44) return false;
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return v.getUint32(0, true) === 1 && v.getUint32(40, true) === 0x464d4520;
}

const RECORD_NAMES: ReadonlyMap<number, string> = new Map([
  [48, 'palette'],
  [49, 'palette'],
  [50, 'palette'],
  [51, 'palette'],
  [52, 'palette'],
  [70, 'comment'],
  [71, 'region fill'],
  [72, 'region frame'],
  [73, 'region invert'],
  [74, 'region paint'],
  [78, 'masked blit'],
  [79, 'parallelogram blit'],
  [116, 'transparent blit'],
  [118, 'gradient fill'],
]);

/**
 * Read an EMF into its primitives.
 *
 * @throws Error when the bytes are not an EMF or its header is malformed.
 */
export function readEmf(bytes: Uint8Array): MetaPicture {
  if (!isEmf(bytes)) throw new Error('EMF: not an enhanced metafile');
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const bounds = {
    left: v.getInt32(8, true),
    top: v.getInt32(12, true),
    right: v.getInt32(16, true),
    bottom: v.getInt32(20, true),
  };

  const prims: Array<PicturePrim> = [];
  const skipped = new Set<string>();
  const objects = new Map<number, MetaObject>();
  const stack: Array<DeviceContext> = [];
  let dc = newDeviceContext();
  // §2.3.11 — the logical→device mapping. Absent window/viewport extents mean
  // MM_TEXT: one logical unit is one device unit.
  let win = { x: 0, y: 0, cx: 1, cy: 1 };
  let view = { x: 0, y: 0, cx: 1, cy: 1 };
  let haveWin = false;
  let haveView = false;
  const sx = (): number => (haveWin && haveView && win.cx !== 0 ? view.cx / win.cx : 1);
  const sy = (): number => (haveWin && haveView && win.cy !== 0 ? view.cy / win.cy : 1);
  /** Logical → device, through the world transform and the window mapping. */
  const px = (x: number, y: number): { x: number; y: number } => {
    const w = applyTransform(dc.transform, x, y);
    return { x: (w.x - win.x) * sx() + view.x, y: (w.y - win.y) * sy() + view.y };
  };

  // The path under construction: EMR_BEGINPATH opens one, the painting records
  // close it. Outside a path bracket the drawing records paint at once.
  let building: Array<PathSegment> | undefined;
  const open = (): Array<PathSegment> => building ?? [];

  // §2.3.2 — the clip the records after it are limited to, kept in the same
  // device space the primitives are stored in. A primitive wholly outside it is
  // dropped: an embedded workbook's preview clips its sheet to the used range,
  // and unclipped we painted the whole empty grid around it.
  const emit = (prim: PicturePrim): void => {
    if (!clippedAway(dc, primBounds(prim))) prims.push(prim);
  };
  const blitter = makeBlitter(emit);
  /**
   * §2.3.1 — one blit record: the source bitmap into the destination
   * rectangle, both stated in the record's own fields. The rectangle is mapped
   * the way every other coordinate here is, so the picture lands in the same
   * device frame as the paths around it.
   */
  const blit = (o: {
    bmiAt: number;
    cbBmi: number;
    bitsAt: number;
    cbBits: number;
    dest: { x: number; y: number; w: number; h: number };
    src?: { x: number; y: number; w: number; h: number };
    /** Whether the source rectangle is stated in the BITMAP's coordinates. */
    inBitmap?: boolean;
    /** §2.3.1.1 — the constant transparency an EMR_ALPHABLEND draws through. */
    blend?: { alpha: number; perPixel: boolean };
    rop: number;
  }): void => {
    const dib =
      o.cbBmi > 0
        ? readDib(bytes, off + o.bmiAt, {
            bitsAt: off + o.bitsAt,
            cbBits: o.cbBits,
            ...(o.blend?.perPixel === true ? { alpha: true } : {}),
          })
        : undefined;
    if (!dib) {
      skipped.add('bitmap');
      return;
    }
    // A source rectangle that is not the whole bitmap takes a part of it; the
    // record then stretches THAT part over the destination. A bottom-up bitmap
    // counts its rows from the bottom, so the part it names is measured there.
    const s = o.src;
    const part =
      s && (s.x !== 0 || s.y !== 0 || s.w !== dib.width || s.h !== dib.height)
        ? cropDib(
            dib,
            s.x,
            o.inBitmap === true && dib.bottomUp === true ? dib.height - s.y - s.h : s.y,
            s.w,
            s.h,
          )
        : dib;
    const a = px(o.dest.x, o.dest.y);
    const b = px(o.dest.x + o.dest.w, o.dest.y + o.dest.h);
    blitter.blit(
      o.blend && o.blend.alpha < 255 ? fadedDib(part, o.blend.alpha / 255) : part,
      {
        x: Math.min(a.x, b.x),
        y: Math.min(a.y, b.y),
        width: Math.abs(b.x - a.x),
        height: Math.abs(b.y - a.y),
      },
      o.rop,
    );
  };

  const strokeOf = (): StrokeStyle | undefined => {
    if (dc.pen.style === 'none') return undefined;
    // A pen one logical unit wide is a HAIRLINE: the thinnest line the device
    // draws, not one unit of a scaled-up world.
    const w = dc.pen.widthLu <= 1 ? 1 : dc.pen.widthLu * Math.abs(sx());
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

  /** Two mapped corners → the rectangle they bound, in either order. */
  function rectOf(
    a: { x: number; y: number },
    b: { x: number; y: number },
  ): { left: number; top: number; right: number; bottom: number } {
    return {
      left: Math.min(a.x, b.x),
      top: Math.min(a.y, b.y),
      right: Math.max(a.x, b.x),
      bottom: Math.max(a.y, b.y),
    };
  }

  /** A closed rectangle, as the geometry records draw one. */
  const rectSegments = (l: number, t: number, r: number, b: number): Array<PathSegment> => {
    const p0 = px(l, t);
    const p1 = px(r, t);
    const p2 = px(r, b);
    const p3 = px(l, b);
    return new PathBuilder()
      .moveTo(p0.x, p0.y)
      .lineTo(p1.x, p1.y)
      .lineTo(p2.x, p2.y)
      .lineTo(p3.x, p3.y)
      .close()
      .build().segments as Array<PathSegment>;
  };

  let off = 0;
  let guard = 0;
  while (off + 8 <= bytes.length && guard++ < 200_000) {
    const type = v.getUint32(off, true);
    const size = v.getUint32(off + 4, true);
    if (size < 8 || off + size > bytes.length) break;
    const at = (o: number): number => v.getInt32(off + o, true);
    const u = (o: number): number => v.getUint32(off + o, true);

    switch (type) {
      case 14: // EMR_EOF
        off = bytes.length;
        continue;
      case 9: // EMR_SETWINDOWEXTEX
        win = { ...win, cx: at(8), cy: at(12) };
        haveWin = true;
        break;
      case 10: // EMR_SETWINDOWORGEX
        win = { ...win, x: at(8), y: at(12) };
        haveWin = true;
        break;
      case 11: // EMR_SETVIEWPORTEXTEX
        view = { ...view, cx: at(8), cy: at(12) };
        haveView = true;
        break;
      case 12: // EMR_SETVIEWPORTORGEX
        view = { ...view, x: at(8), y: at(12) };
        haveView = true;
        break;
      // §2.3.2 clip records. Only the rectangular forms are modelled; an
      // EXTSELECTCLIPRGN whose region is one rectangle is one of them, and an
      // empty region with RGN_COPY clears the clip.
      case 30: {
        // EMR_INTERSECTCLIPRECT — a RECTL in logical units
        const a = px(at(8), at(12));
        const b = px(at(16), at(20));
        dc = { ...dc, clip: intersectClip(dc, rectOf(a, b)) };
        break;
      }
      case 75: {
        // EMR_EXTSELECTCLIPRGN — RegionData, then the mode
        const cbRgnData = u(8);
        const mode = u(12);
        if (cbRgnData === 0 && mode === 5) {
          const { clip: _drop, ...rest } = dc;
          dc = rest;
        } else if (mode === 5 && cbRgnData >= 32 && u(16 + 4) === 1) {
          // One RECTL follows the 32-byte RGNDATAHEADER.
          const base = off + 16 + 32;
          const a = px(v.getInt32(base, true), v.getInt32(base + 4, true));
          const b = px(v.getInt32(base + 8, true), v.getInt32(base + 12, true));
          dc = { ...dc, clip: rectOf(a, b) };
        }
        break;
      }
      case 33: // EMR_SAVEDC
        stack.push(cloneDc(dc));
        break;
      case 34: // EMR_RESTOREDC
        {
          const back = Math.max(1, -at(8));
          for (let i = 0; i < back && stack.length > 0; i++) dc = stack.pop()!;
        }
        break;
      case 35: // EMR_SETWORLDTRANSFORM
        dc.transform = xform(v, off + 8);
        break;
      case 36: // EMR_MODIFYWORLDTRANSFORM
        {
          const m = xform(v, off + 8);
          const mode = u(8 + 24);
          // MWT_IDENTITY 1 / MWT_LEFTMULTIPLY 2 / MWT_RIGHTMULTIPLY 3 / SET 4
          dc.transform =
            mode === 1
              ? [1, 0, 0, 1, 0, 0]
              : mode === 2
                ? multiply(m, dc.transform)
                : mode === 3
                  ? multiply(dc.transform, m)
                  : m;
        }
        break;
      case 37: // EMR_SELECTOBJECT
        {
          const ih = u(8);
          const stock = stockObject(ih);
          const obj = stock ?? objects.get(ih);
          if (obj?.kind === 'pen') dc.pen = obj;
          else if (obj?.kind === 'brush') dc.brush = obj;
          else if (obj?.kind === 'font') dc.font = obj;
        }
        break;
      case 38: // EMR_CREATEPEN — the LOGPEN's width is a POINT, so its colour
        // sits a whole point past the style, not half of one.
        objects.set(u(8), {
          kind: 'pen',
          colorHex: colorRef(u(24)),
          widthLu: at(16),
          style: penStyle(u(12)),
        });
        break;
      case 95: // EMR_EXTCREATEPEN — the EXTLOGPEN sits past the pen's own
        // bitmap fields: style, width, BRUSH style, then the colour.
        objects.set(u(8), {
          kind: 'pen',
          colorHex: colorRef(u(40)),
          widthLu: at(32),
          // A pen whose brush is BS_NULL paints nothing, whatever its style.
          style: u(36) === 1 ? 'none' : penStyle(u(28)),
        });
        break;
      case 39: // EMR_CREATEBRUSHINDIRECT
        {
          const style = u(12);
          objects.set(u(8), {
            kind: 'brush',
            colorHex: colorRef(u(16)),
            // BS_NULL (1) paints nothing; a HATCHED brush (2) is drawn as its
            // colour, which is what a hatch reads as at document scale.
            hollow: style === 1,
          });
        }
        break;
      case 82: // EMR_EXTCREATEFONTINDIRECTW
        objects.set(u(8), readLogFont(v, bytes, off + 12));
        break;
      case 40: // EMR_DELETEOBJECT
        objects.delete(u(8));
        break;
      case 18: // EMR_SETBKMODE
        dc.bkOpaque = u(8) === 2;
        break;
      case 19: // EMR_SETPOLYFILLMODE
        dc.fillRule = u(8) === 2 ? 'nonzero' : 'evenodd';
        break;
      case 22: // EMR_SETTEXTALIGN
        {
          const f = u(8);
          dc.alignH = (f & 6) === 6 ? 'center' : (f & 2) === 2 ? 'right' : 'left';
          dc.alignBaseline = (f & 24) === 24;
        }
        break;
      case 24: // EMR_SETTEXTCOLOR
        dc.textColorHex = colorRef(u(8));
        break;
      case 25: // EMR_SETBKCOLOR
        dc.bkColorHex = colorRef(u(8));
        break;
      case 27: // EMR_MOVETOEX
        dc.x = at(8);
        dc.y = at(12);
        if (building) {
          const p = px(dc.x, dc.y);
          building.push({ op: 'move', x: p.x, y: p.y });
        }
        break;
      case 54: // EMR_LINETO
        {
          const from = px(dc.x, dc.y);
          dc.x = at(8);
          dc.y = at(12);
          const to = px(dc.x, dc.y);
          if (building) {
            if (building.length === 0) building.push({ op: 'move', x: from.x, y: from.y });
            building.push({ op: 'line', x: to.x, y: to.y });
          } else {
            paint(
              [
                { op: 'move', x: from.x, y: from.y },
                { op: 'line', x: to.x, y: to.y },
              ],
              false,
              true,
            );
          }
        }
        break;
      case 59: // EMR_BEGINPATH
        building = [];
        break;
      case 61: // EMR_CLOSEFIGURE
        if (building && building.length > 0) building.push({ op: 'close' });
        break;
      case 60: // EMR_ENDPATH
        break;
      case 62: // EMR_FILLPATH
      case 63: // EMR_STROKEANDFILLPATH
      case 64: // EMR_STROKEPATH
        paint(open(), type !== 64, type !== 62);
        building = undefined;
        break;
      case 68: // EMR_ABORTPATH
        building = undefined;
        break;
      case 43: // EMR_RECTANGLE
        paint(rectSegments(at(8), at(12), at(16), at(20)), true, true);
        break;
      case 42: // EMR_ELLIPSE
        paint(ellipseSegments(px, at(8), at(12), at(16), at(20)), true, true);
        break;
      case 44: // EMR_ROUNDRECT — drawn square-cornered; the radius is smaller
        // than a point at any size a document prints one.
        paint(rectSegments(at(8), at(12), at(16), at(20)), true, true);
        break;
      case 47: // EMR_PIE
      case 46: // EMR_CHORD
      case 45: // EMR_ARC
        // §2.3.5.13/.4/.2 — a bounding box and two RADIAL points: the arc runs
        // between where the rays from the box's centre through them meet the
        // ellipse. A pie closes through that centre and a chord straight
        // across; an arc is only the curve, so it is stroked and never filled.
        paint(
          arcSegments(px, at(8), at(12), at(16), at(20), at(24), at(28), at(32), at(36), type),
          type !== 45,
          true,
        );
        break;
      case 3: // EMR_POLYGON
      case 4: // EMR_POLYLINE
      case 2: // EMR_POLYBEZIER
      case 5: // EMR_POLYBEZIERTO
      case 6: // EMR_POLYLINETO
      case 86: // EMR_POLYGON16
      case 87: // EMR_POLYLINE16
      case 85: // EMR_POLYBEZIER16
      case 88: // EMR_POLYBEZIERTO16
      case 89: // EMR_POLYLINETO16
        {
          const small = type >= 85;
          const count = u(24);
          const pts = readPoints(v, off + 28, count, small);
          const closed = type === 3 || type === 86;
          const bezier = type === 2 || type === 5 || type === 85 || type === 88;
          const continues = type === 5 || type === 6 || type === 88 || type === 89;
          const segs = polySegments(px, pts, {
            closed,
            bezier,
            ...(continues ? { from: { x: dc.x, y: dc.y } } : {}),
          });
          if (pts.length > 0) {
            dc.x = pts[pts.length - 1]!.x;
            dc.y = pts[pts.length - 1]!.y;
          }
          if (building) building.push(...segs);
          else paint(segs, closed, true);
        }
        break;
      case 90: // EMR_POLYPOLYLINE16
      case 91: // EMR_POLYPOLYGON16
      case 7: // EMR_POLYPOLYLINE
      case 8: // EMR_POLYPOLYGON
        {
          const small = type >= 90;
          const polys = u(24);
          const total = u(28);
          const counts: Array<number> = [];
          for (let i = 0; i < polys; i++) counts.push(v.getUint32(off + 32 + i * 4, true));
          const pts = readPoints(v, off + 32 + polys * 4, total, small);
          const closed = type === 91 || type === 8;
          const segs: Array<PathSegment> = [];
          let k = 0;
          for (const c of counts) {
            segs.push(...polySegments(px, pts.slice(k, k + c), { closed, bezier: false }));
            k += c;
          }
          if (building) building.push(...segs);
          else paint(segs, closed, true);
        }
        break;
      case 83: // EMR_EXTTEXTOUTA
      case 84: // EMR_EXTTEXTOUTW
        {
          // §2.3.5.7 — the EMRTEXT follows the record's own fields; its string
          // offset is measured from the START of the record.
          const refX = at(36);
          const refY = at(40);
          const chars = u(44);
          const offString = u(48);
          const wide = type === 84;
          const text = readString(bytes, off + offString, chars, wide);
          if (text.trim() !== '') {
            const p = px(refX, refY);
            const font = dc.font;
            emit({
              kind: 'text',
              // A symbol font's letters are not letters (see symbol-fonts).
              text: fromSymbolFont(text, font?.family),
              x: p.x,
              y: p.y,
              alignH: dc.alignH,
              alignBaseline: dc.alignBaseline,
              sizeLu: Math.abs(font?.heightLu ?? 12) * Math.abs(sy()),
              colorHex: dc.textColorHex,
              ...(font?.family ? { fontFamily: font.family } : {}),
              ...(font?.bold ? { bold: true } : {}),
              ...(font?.italic ? { italic: true } : {}),
              ...(font?.escapement ? { escapement: font.escapement } : {}),
            });
          }
        }
        break;
      // EMR_BITBLT / EMR_STRETCHBLT — the same fields but for the source
      // extent a stretch states. With NO source bitmap either is a BRUSH blit:
      // the destination rectangle filled with the current brush, which is how a
      // metafile paints a rule or a panel.
      case 76:
      case 77: {
        // §2.3.1.2 — offBmiSrc is at 84 and its SIZE at 88; a blit with no
        // bitmap behind it is the brush painting the destination.
        const cbBmiSrc = u(88);
        const rop = u(40);
        const x = at(24);
        const y = at(28);
        const w = at(32);
        const h = at(36);
        if (cbBmiSrc === 0) {
          // BLACKNESS / WHITENESS paint a colour of their own; PATCOPY and
          // its neighbours paint the brush.
          const hex = rop === 0x00000042 ? '000000' : rop === 0x00ff0062 ? 'FFFFFF' : undefined;
          const brush = dc.brush;
          if (hex !== undefined) {
            dc.brush = { kind: 'brush', colorHex: hex, hollow: false };
          }
          paint(rectSegments(x, y, x + w, y + h), true, false);
          dc.brush = brush;
          break;
        }
        // §2.3.1.5 — a STRETCHBLT states the source extent it scales from;
        // a plain BITBLT takes the destination's, one pixel for one.
        blit({
          bmiAt: u(84),
          cbBmi: cbBmiSrc,
          bitsAt: u(92),
          cbBits: u(96),
          dest: { x, y, w, h },
          src: {
            x: at(44),
            y: at(48),
            w: type === 77 ? at(100) : w,
            h: type === 77 ? at(104) : h,
          },
          rop,
        });
        break;
      }
      case 80: // EMR_SETDIBITSTODEVICE — the source bitmap, unscaled, at the
        // destination point: its own extent is the size it lands at.
        blit({
          bmiAt: u(48),
          cbBmi: u(52),
          bitsAt: u(56),
          cbBits: u(60),
          dest: { x: at(24), y: at(28), w: at(40), h: at(44) },
          src: { x: at(32), y: at(36), w: at(40), h: at(44) },
          inBitmap: true,
          rop: 0x00cc0020,
        });
        break;
      case 81: // EMR_STRETCHDIBITS — the source rectangle scaled onto the
        // destination one; the record that carries a picture in an EMF.
        blit({
          bmiAt: u(48),
          cbBmi: u(52),
          bitsAt: u(56),
          cbBits: u(60),
          dest: { x: at(24), y: at(28), w: at(72), h: at(76) },
          src: { x: at(32), y: at(36), w: at(40), h: at(44) },
          inBitmap: true,
          rop: u(68),
        });
        break;
      case 114: // EMR_ALPHABLEND — a STRETCHBLT whose raster operation is a
        // BLENDFUNCTION instead: a constant transparency, and a flag saying the
        // bitmap carries an alpha channel of its own.
        blit({
          bmiAt: u(84),
          cbBmi: u(88),
          bitsAt: u(92),
          cbBits: u(96),
          dest: { x: at(24), y: at(28), w: at(32), h: at(36) },
          src: { x: at(44), y: at(48), w: at(100), h: at(104) },
          blend: { alpha: (u(40) >>> 16) & 0xff, perPixel: u(40) >>> 24 === 1 },
          rop: 0x00cc0020,
        });
        break;
      default:
        {
          const name = RECORD_NAMES.get(type);
          if (name) skipped.add(name);
        }
        break;
    }
    off += size;
  }

  return {
    left: bounds.left,
    top: bounds.top,
    width: Math.max(1, bounds.right - bounds.left),
    height: Math.max(1, bounds.bottom - bounds.top),
    prims,
    skipped: [...skipped],
  };
}

function xform(v: DataView, at: number): DeviceContext['transform'] {
  return [
    v.getFloat32(at, true),
    v.getFloat32(at + 4, true),
    v.getFloat32(at + 8, true),
    v.getFloat32(at + 12, true),
    v.getFloat32(at + 16, true),
    v.getFloat32(at + 20, true),
  ];
}

function multiply(
  a: DeviceContext['transform'],
  b: DeviceContext['transform'],
): DeviceContext['transform'] {
  return [
    a[0] * b[0] + a[1] * b[2],
    a[0] * b[1] + a[1] * b[3],
    a[2] * b[0] + a[3] * b[2],
    a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4],
    a[4] * b[1] + a[5] * b[3] + b[5],
  ];
}

// §2.1.25 PenStyle: the low nibble is the line style; 5 is PS_NULL.
function penStyle(style: number): NonNullable<MetaObject & { kind: 'pen' }>['style'] {
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

// §2.1.31 — the stock objects a record may select without creating one.
function stockObject(ih: number): MetaObject | undefined {
  if ((ih & 0x80000000) === 0) return undefined;
  switch (ih & 0x7fffffff) {
    case 0: // WHITE_BRUSH
      return { kind: 'brush', colorHex: 'FFFFFF', hollow: false };
    case 1: // LTGRAY_BRUSH
      return { kind: 'brush', colorHex: 'C0C0C0', hollow: false };
    case 2: // GRAY_BRUSH
      return { kind: 'brush', colorHex: '808080', hollow: false };
    case 3: // DKGRAY_BRUSH
      return { kind: 'brush', colorHex: '404040', hollow: false };
    case 4: // BLACK_BRUSH
      return { kind: 'brush', colorHex: '000000', hollow: false };
    case 5: // NULL_BRUSH
      return { kind: 'brush', colorHex: '000000', hollow: true };
    case 6: // WHITE_PEN
      return { kind: 'pen', colorHex: 'FFFFFF', widthLu: 1, style: 'solid' };
    case 7: // BLACK_PEN
      return { kind: 'pen', colorHex: '000000', widthLu: 1, style: 'solid' };
    case 8: // NULL_PEN
      return { kind: 'pen', colorHex: '000000', widthLu: 1, style: 'none' };
    default:
      return undefined;
  }
}

// §2.2.13 LogFont — the face name is 32 UTF-16 code units at its end.
function readLogFont(v: DataView, bytes: Uint8Array, at: number): MetaObject {
  const height = v.getInt32(at, true);
  const escapement = v.getInt32(at + 8, true);
  const weight = v.getInt32(at + 16, true);
  const italic = bytes[at + 20] !== 0;
  const face = readString(bytes, at + 28, 32, true).replace(/\0.*$/u, '');
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
    if (o + (wide ? 1 : 0) >= bytes.length) break;
    const code = wide ? bytes[o]! | (bytes[o + 1]! << 8) : bytes[o]!;
    if (code === 0) break;
    out += String.fromCharCode(code);
  }
  return out;
}

function readPoints(
  v: DataView,
  at: number,
  count: number,
  small: boolean,
): Array<{ x: number; y: number }> {
  const out: Array<{ x: number; y: number }> = [];
  const step = small ? 4 : 8;
  for (let i = 0; i < count; i++) {
    const o = at + i * step;
    if (o + step > v.byteLength) break;
    out.push(
      small
        ? { x: v.getInt16(o, true), y: v.getInt16(o + 2, true) }
        : { x: v.getInt32(o, true), y: v.getInt32(o + 4, true) },
    );
  }
  return out;
}

type Mapper = (x: number, y: number) => { x: number; y: number };

function polySegments(
  map: Mapper,
  pts: ReadonlyArray<{ x: number; y: number }>,
  o: { closed: boolean; bezier: boolean; from?: { x: number; y: number } },
): Array<PathSegment> {
  const segs: Array<PathSegment> = [];
  if (pts.length === 0) return segs;
  let i = 0;
  if (o.from) {
    const s = map(o.from.x, o.from.y);
    segs.push({ op: 'move', x: s.x, y: s.y });
  } else {
    const s = map(pts[0]!.x, pts[0]!.y);
    segs.push({ op: 'move', x: s.x, y: s.y });
    i = 1;
  }
  if (o.bezier) {
    for (; i + 2 < pts.length + 1 && i + 2 <= pts.length; i += 3) {
      const c1 = map(pts[i]!.x, pts[i]!.y);
      const c2 = map(pts[i + 1]!.x, pts[i + 1]!.y);
      const to = map(pts[i + 2]!.x, pts[i + 2]!.y);
      segs.push({ op: 'cubic', x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y, x: to.x, y: to.y });
    }
  } else {
    for (; i < pts.length; i++) {
      const p = map(pts[i]!.x, pts[i]!.y);
      segs.push({ op: 'line', x: p.x, y: p.y });
    }
  }
  if (o.closed) segs.push({ op: 'close' });
  return segs;
}

// An ellipse as four cubics — the constant is the usual circle approximation.
const KAPPA = 0.5522847498307936;

/**
 * §2.3.5.2/.4/.13 — the wedge, chord or bare curve an `EMR_ARC`, `EMR_CHORD` or
 * `EMR_PIE` draws.
 *
 * All three state the same thing: an ellipse inscribed in a box, and two points
 * naming RAYS from its centre. The arc runs from where the first ray crosses
 * the ellipse to where the second does, anticlockwise on the page — GDI's
 * default direction — and the two records differ only in how they close: a pie
 * back through the centre, a chord straight across, an arc not at all.
 *
 * Built in the metafile's own coordinates and mapped point by point, exactly as
 * the ellipse and the polygons are: the mapping is affine, so a Bézier's
 * control points survive it.
 *
 * @param map        The metafile's logical → page mapping.
 * @param l,t,r,b    The box the ellipse is inscribed in.
 * @param sx,sy      The point naming the ray the arc starts at.
 * @param ex,ey      The point naming the ray it ends at.
 * @param type       The record type, which decides how the figure closes.
 * @returns The path segments, in page coordinates.
 */
function arcSegments(
  map: Mapper,
  l: number,
  t: number,
  r: number,
  b: number,
  sx: number,
  sy: number,
  ex: number,
  ey: number,
  type: number,
): Array<PathSegment> {
  const cx = (l + r) / 2;
  const cy = (t + b) / 2;
  const rx = (r - l) / 2;
  const ry = (b - t) / 2;
  if (!(rx > 0) || !(ry > 0)) return [];
  // The angle of a ray, in the frame the ellipse is parametrised in. The
  // metafile's y grows DOWN, so this angle grows clockwise on the page.
  const angleOf = (x: number, y: number): number => Math.atan2((y - cy) / ry, (x - cx) / rx);
  const start = angleOf(sx, sy);
  let sweep = angleOf(ex, ey) - start;
  // Anticlockwise on the page is the direction of DECREASING angle here; two
  // rays that coincide name the whole ellipse, not nothing.
  if (sweep >= 0) sweep -= 2 * Math.PI;
  const point = (a: number): { x: number; y: number } =>
    map(cx + rx * Math.cos(a), cy + ry * Math.sin(a));
  const segs: Array<PathSegment> = [];
  const from = point(start);
  if (type === 47) {
    // A pie opens at the centre and runs out to the arc.
    const c = map(cx, cy);
    segs.push({ op: 'move', x: c.x, y: c.y }, { op: 'line', x: from.x, y: from.y });
  } else {
    segs.push({ op: 'move', x: from.x, y: from.y });
  }
  // No cubic spans more than a quarter turn, which is where the approximation
  // stays true to within a fraction of a device pixel.
  const steps = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2)));
  const delta = sweep / steps;
  const alpha = (4 / 3) * Math.tan(delta / 4);
  for (let i = 0; i < steps; i++) {
    const a0 = start + delta * i;
    const a1 = a0 + delta;
    // The tangent at an angle, scaled by the handle length the arc needs.
    const t0 = map(
      cx + rx * Math.cos(a0) - alpha * rx * Math.sin(a0),
      cy + ry * Math.sin(a0) + alpha * ry * Math.cos(a0),
    );
    const t1 = map(
      cx + rx * Math.cos(a1) + alpha * rx * Math.sin(a1),
      cy + ry * Math.sin(a1) - alpha * ry * Math.cos(a1),
    );
    const to = point(a1);
    segs.push({ op: 'cubic', x1: t0.x, y1: t0.y, x2: t1.x, y2: t1.y, x: to.x, y: to.y });
  }
  if (type !== 45) segs.push({ op: 'close' });
  return segs;
}

export function ellipseSegments(
  map: Mapper,
  l: number,
  t: number,
  r: number,
  b: number,
): Array<PathSegment> {
  const cx = (l + r) / 2;
  const cy = (t + b) / 2;
  const rx = (r - l) / 2;
  const ry = (b - t) / 2;
  const k = KAPPA;
  const p = (x: number, y: number) => map(x, y);
  const s = p(cx + rx, cy);
  const segs: Array<PathSegment> = [{ op: 'move', x: s.x, y: s.y }];
  const arc = (
    c1x: number,
    c1y: number,
    c2x: number,
    c2y: number,
    ex: number,
    ey: number,
  ): void => {
    const a = p(c1x, c1y);
    const bb = p(c2x, c2y);
    const e = p(ex, ey);
    segs.push({ op: 'cubic', x1: a.x, y1: a.y, x2: bb.x, y2: bb.y, x: e.x, y: e.y });
  };
  arc(cx + rx, cy + ry * k, cx + rx * k, cy + ry, cx, cy + ry);
  arc(cx - rx * k, cy + ry, cx - rx, cy + ry * k, cx - rx, cy);
  arc(cx - rx, cy - ry * k, cx - rx * k, cy - ry, cx, cy - ry);
  arc(cx + rx * k, cy - ry, cx + rx, cy - ry * k, cx + rx, cy);
  segs.push({ op: 'close' });
  return segs;
}
