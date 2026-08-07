// ECMA-376 §20.1.10.55 — preset shape geometries → vector paths.
//
// Produces paths in a LOCAL y-up frame (origin bottom-left) sized w×h points,
// matching the vector-graphics emit layer. The y-down→y-up flip therefore
// lives here (and in custGeom), never at emit time. Adjust values are the raw
// a:gd numbers (Word's preset guides are mostly thousandths-of-a-percent).

import type { CustomGeometry } from '@/core/document-model';
import type { PathSegment, VectorPath } from '@/core/vector';
import { arcToBeziers, ellipseSegments, roundRectSegments } from '@/core/arc-to-bezier';
import { PathBuilder } from '@/core/vector';

/**
 * The four closed segments of the `w`×`h` bounding rectangle, in the local y-up
 * frame (origin bottom-left). The fallback geometry for unknown presets.
 */
export function rectSegments(w: number, h: number): ReadonlyArray<PathSegment> {
  return new PathBuilder().moveTo(0, 0).lineTo(w, 0).lineTo(w, h).lineTo(0, h).close().build()
    .segments;
}

/** The `w`×`h` bounding rectangle as a {@link VectorPath}. */
export function rectPath(w: number, h: number): VectorPath {
  return { segments: rectSegments(w, h) };
}

// Closed polygon from a point list (y-up, points in path order).
function polygon(points: ReadonlyArray<readonly [number, number]>): VectorPath {
  const b = new PathBuilder();
  points.forEach(([x, y], i) => (i === 0 ? b.moveTo(x, y) : b.lineTo(x, y)));
  return b.close().build();
}

const clamp = (v: number, lo: number, hi: number): number => Math.max(lo, Math.min(hi, v));

// Adjust guide as a fraction (raw a:gd val ÷ 100000), with a default.
const frac = (adjust: ReadonlyMap<string, number>, key: string, def: number): number =>
  (adjust.get(key) ?? def) / 100000;

/**
 * Map a preset shape name (§20.1.10.55) to its vector path(s) in the local y-up
 * frame, sized `w`×`h` points and shaped by the `adjust` guides. Returns `null`
 * for an unknown preset so the caller falls back to the bounding rectangle —
 * graceful degradation that keeps even unimplemented presets visible (with the
 * right fill/line) and never throws.
 *
 * @param preset The preset geometry name (e.g. `'roundRect'`, `'rightArrow'`).
 * @param w      Box width in points.
 * @param h      Box height in points.
 * @param adjust Raw `a:gd` adjust guides by name (thousandths of a percent).
 * @returns The path(s), or `null` for an unrecognised preset.
 */
export function presetPaths(
  preset: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>,
): Array<VectorPath> | null {
  switch (preset) {
    case 'rect':
      return [rectPath(w, h)];
    case 'roundRect': {
      // adj is a fraction of the shorter side (default 16667 ⇒ ~1/6).
      return [{ segments: roundRectSegments(w, h, frac(adjust, 'adj', 16667) * Math.min(w, h)) }];
    }
    case 'ellipse':
      return [{ segments: ellipseSegments(w, h) }];
    case 'triangle': {
      // adj = horizontal apex position (default centred).
      const apexX = clamp(frac(adjust, 'adj', 50000), 0, 1) * w;
      return [
        polygon([
          [0, 0],
          [w, 0],
          [apexX, h],
        ]),
      ];
    }
    case 'rtTriangle':
      return [
        polygon([
          [0, 0],
          [w, 0],
          [0, h],
        ]),
      ];
    case 'diamond':
      return [
        polygon([
          [w / 2, 0],
          [w, h / 2],
          [w / 2, h],
          [0, h / 2],
        ]),
      ];
    case 'parallelogram': {
      const s = clamp(frac(adjust, 'adj', 25000) * w, 0, w);
      return [
        polygon([
          [0, 0],
          [w - s, 0],
          [w, h],
          [s, h],
        ]),
      ];
    }
    case 'trapezoid': {
      const t = clamp(frac(adjust, 'adj', 25000) * w, 0, w / 2);
      return [
        polygon([
          [0, 0],
          [w, 0],
          [w - t, h],
          [t, h],
        ]),
      ];
    }
    case 'pentagon':
      return [regularPolygon(w, h, 5)];
    // §20.1.10.55 — the regular polygons the gallery names by their side count.
    // An odd one stands on a vertex, an even one on a flat edge, which is the
    // half-step of rotation between them.
    case 'heptagon':
      return [regularPolygon(w, h, 7)];
    case 'octagon':
      return [regularPolygon(w, h, 8, Math.PI / 8)];
    case 'decagon':
      return [regularPolygon(w, h, 10, Math.PI / 10)];
    case 'dodecagon':
      return [regularPolygon(w, h, 12, Math.PI / 12)];
    // A ring and a rectangular frame are each one shape with a HOLE, which is
    // an inner subpath wound the same way and filled even-odd.
    case 'donut': {
      const t = clamp(frac(adjust, 'adj', 25000), 0, 0.5) * Math.min(w, h);
      const b = new PathBuilder().append(ellipseSegments(w, h));
      const [ix, iy] = [w / 2 - t, h / 2 - t];
      if (ix > 0 && iy > 0) {
        b.moveTo(w / 2 + ix, h / 2).append(arcToBeziers(w / 2, h / 2, ix, iy, 0, 2 * Math.PI));
        b.close();
      }
      return [b.build('evenodd')];
    }
    case 'frame': {
      const t = clamp(frac(adjust, 'adj', 12500), 0, 0.5) * Math.min(w, h);
      const b = new PathBuilder()
        .moveTo(0, 0)
        .lineTo(w, 0)
        .lineTo(w, h)
        .lineTo(0, h)
        .close()
        .moveTo(t, t)
        .lineTo(w - t, t)
        .lineTo(w - t, h - t)
        .lineTo(t, h - t)
        .close();
      return [b.build('evenodd')];
    }
    // A band laid across the box corner to corner: the fraction it names is how
    // far along each edge it starts.
    case 'diagStripe': {
      const a = clamp(frac(adjust, 'adj', 50000), 0, 1);
      return [
        polygon([
          [0, h - h * a],
          [w * a, h],
          [w, h],
          [0, 0],
        ]),
      ];
    }
    // A circle with one quadrant pulled out to the corner of the box.
    case 'teardrop': {
      const a = clamp(frac(adjust, 'adj', 100000), 0, 2);
      const [rx, ry] = [w / 2, h / 2];
      const b = new PathBuilder()
        .moveTo(rx, h)
        .append(arcToBeziers(rx, ry, rx, ry, Math.PI / 2, (3 * Math.PI) / 2))
        .lineTo(w * (0.5 + a / 2), h * (0.5 + a / 2))
        .close();
      return [b.build()];
    }
    // A cylinder seen from the side: the body under an ellipse, and the ellipse
    // drawn again on top so its near edge shows.
    case 'can':
      return cylinder(w, h, (clamp(frac(adjust, 'adj', 25000), 0, 0.5) * h) / 2);
    // Two bars meeting at a mitred corner — the top-left half of a frame. Each
    // bar's thickness is a fraction of the SHORTER side, and the mitre runs at
    // the box's own diagonal, so it is the thickness scaled by the aspect.
    case 'halfFrame': {
      const ss = Math.min(w, h);
      const x1 = clamp(frac(adjust, 'adj2', 33333), 0, 1) * ss;
      const y1 = clamp(frac(adjust, 'adj1', 33333), 0, 1) * ss;
      const x2 = w - (h === 0 ? 0 : (y1 * w) / h);
      const y2 = h - (w === 0 ? 0 : (x1 * h) / w);
      return [
        polygon([
          [0, 0],
          [0, h],
          [w, h],
          [x2, h - y1],
          [x1, h - y1],
          [x1, h - y2],
        ]),
      ];
    }
    // A rectangle with its four edges chamfered inwards. PowerPoint shades each
    // face to make it read as a raised block; a shape here takes one fill, so
    // what carries is the STRUCTURE — the inner rectangle and the four mitres.
    case 'bevel': {
      const t = clamp(frac(adjust, 'adj', 12500), 0, 0.5) * Math.min(w, h);
      const face = (pts: ReadonlyArray<readonly [number, number]>): VectorPath => polygon(pts);
      return [
        face([
          [0, h],
          [w, h],
          [w - t, h - t],
          [t, h - t],
        ]),
        face([
          [0, h],
          [t, h - t],
          [t, t],
          [0, 0],
        ]),
        face([
          [w, 0],
          [0, 0],
          [t, t],
          [w - t, t],
        ]),
        face([
          [w, 0],
          [w - t, t],
          [w - t, h - t],
          [w, h],
        ]),
        polygon([
          [t, t],
          [w - t, t],
          [w - t, h - t],
          [t, h - t],
        ]),
      ];
    }
    // A rectangle whose corners are cut IN rather than off — the quarter circle
    // is centred outside the shape, so it bites into it.
    case 'plaque': {
      const t = clamp(frac(adjust, 'adj', 16667), 0, 0.5) * Math.min(w, h);
      // Each quarter circle is centred ON the box corner, so it bites inwards;
      // centred inside, the same radius rounds the corner off instead.
      const b = new PathBuilder()
        .moveTo(t, 0)
        .lineTo(w - t, 0)
        .append(arcToBeziers(w, 0, t, t, Math.PI, -Math.PI / 2))
        .lineTo(w, h - t)
        .append(arcToBeziers(w, h, t, t, (3 * Math.PI) / 2, -Math.PI / 2))
        .lineTo(t, h)
        .append(arcToBeziers(0, h, t, t, 0, -Math.PI / 2))
        .lineTo(0, t)
        .append(arcToBeziers(0, 0, t, t, Math.PI / 2, -Math.PI / 2))
        .close();
      return [b.build()];
    }
    case 'hexagon': {
      const inset = clamp(frac(adjust, 'adj', 25000) * w, 0, w / 2);
      return [
        polygon([
          [inset, 0],
          [w - inset, 0],
          [w, h / 2],
          [w - inset, h],
          [inset, h],
          [0, h / 2],
        ]),
      ];
    }
    // §20.1.10.55 — the rectangles whose corners are rounded or cut off. Each
    // names which corners with one or two adjust guides; the radius is that
    // fraction of the SHORTER side.
    case 'round1Rect':
    case 'round2SameRect':
    case 'round2DiagRect':
    case 'snip1Rect':
    case 'snip2SameRect':
    case 'snip2DiagRect':
    case 'snipRoundRect':
      return [cornerRect(preset, w, h, adjust)];
    case 'star4':
    case 'star5':
    case 'star6':
    case 'star7':
    case 'star8':
    case 'star10':
    case 'star12':
    case 'star16':
    case 'star24':
    case 'star32': {
      const s = STARS.get(preset);
      return s ? [star(w, h, s, adjust)] : null;
    }
    case 'line':
    case 'straightConnector1':
      // Box diagonal (top-left → bottom-right in y-down ⇒ (0,h)→(w,0) here).
      // Open path: a connector is stroked, never filled.
      return [new PathBuilder().moveTo(0, h).lineTo(w, 0).build()];
    case 'bentConnector2':
      // §20.1.10.55 — one right-angle: along the top, then down the far side.
      return [new PathBuilder().moveTo(0, h).lineTo(w, h).lineTo(w, 0).build()];
    case 'bentConnector3': {
      // …and two, with the upright standing where `adj1` puts it across the box
      // (half way by default). VML's bent connector (`o:spt="34"`) is the same
      // shape: groupshape-child-rotation.docx joins its two boxes with one.
      const bx = clamp(frac(adjust, 'adj1', 50000), 0, 1) * w;
      return [new PathBuilder().moveTo(0, h).lineTo(bx, h).lineTo(bx, 0).lineTo(w, 0).build()];
    }
    case 'rightArrow':
      return [blockArrow(w, h, adjust, 'right')];
    case 'leftArrow':
      return [blockArrow(w, h, adjust, 'left')];
    case 'upArrow':
      return [blockArrow(w, h, adjust, 'up')];
    case 'downArrow':
      return [blockArrow(w, h, adjust, 'down')];
    case 'leftRightArrow': {
      // A shaft with a head at each end, both measured against the short side.
      const ss = Math.min(w, h);
      const t = clamp(frac(adjust, 'adj1', 50000), 0, 1) * ss;
      const hl = clamp(frac(adjust, 'adj2', 50000), 0, 1) * ss;
      const [sb, st] = [(h - t) / 2, (h + t) / 2];
      return [
        polygon([
          [0, h / 2],
          [hl, h],
          [hl, st],
          [w - hl, st],
          [w - hl, h],
          [w, h / 2],
          [w - hl, 0],
          [w - hl, sb],
          [hl, sb],
          [hl, 0],
        ]),
      ];
    }
    case 'notchedRightArrow': {
      // A right arrow whose tail is cut back into a V — the notch is as deep as
      // the head is long, which is what makes it read as a ribbon.
      const ss = Math.min(w, h);
      const t = clamp(frac(adjust, 'adj1', 50000), 0, 1) * ss;
      const hl = clamp(frac(adjust, 'adj2', 50000), 0, 1) * ss;
      const [sb, st] = [(h - t) / 2, (h + t) / 2];
      const notch = Math.min(hl, w / 2);
      return [
        polygon([
          [0, sb],
          [w - hl, sb],
          [w - hl, 0],
          [w, h / 2],
          [w - hl, h],
          [w - hl, st],
          [0, st],
          [notch, h / 2],
        ]),
      ];
    }
    case 'stripedRightArrow': {
      // The same arrow behind two bars: a thin one at the very tail and a
      // wider one beside it, both a slice of the short side (§20.1.10.55).
      const ss = Math.min(w, h);
      const t = clamp(frac(adjust, 'adj1', 50000), 0, 1) * ss;
      const hl = clamp(frac(adjust, 'adj2', 50000), 0, 1) * ss;
      const [sb, st] = [(h - t) / 2, (h + t) / 2];
      const unit = ss / 12;
      return [
        polygon([
          [0, sb],
          [unit, sb],
          [unit, st],
          [0, st],
        ]),
        polygon([
          [unit * 2, sb],
          [unit * 4, sb],
          [unit * 4, st],
          [unit * 2, st],
        ]),
        polygon([
          [unit * 5, sb],
          [w - hl, sb],
          [w - hl, 0],
          [w, h / 2],
          [w - hl, h],
          [w - hl, st],
          [unit * 5, st],
        ]),
      ];
    }
    case 'bentUpArrow': {
      // An L standing on its long arm, with the head at the top of the short
      // one: the corner a flow turns at.
      const ss = Math.min(w, h);
      const t = clamp(frac(adjust, 'adj2', 25000), 0, 1) * ss;
      const hw = clamp(frac(adjust, 'adj3', 25000), 0, 1) * ss * 2;
      const hh = Math.min(hw, h);
      const cx = w - hw / 2;
      return [
        polygon([
          [0, 0],
          [cx + t / 2, 0],
          [cx + t / 2, h - hh],
          [w, h - hh],
          [cx, h],
          [cx - hw / 2, h - hh],
          [cx - t / 2, h - hh],
          [cx - t / 2, t],
          [0, t],
        ]),
      ];
    }
    case 'chevron':
    case 'homePlate': {
      // A block that points: the same pentagon, with the chevron notched at the
      // back so a row of them interlocks.
      const ss = Math.min(w, h);
      const x = Math.min(clamp(frac(adjust, 'adj', 50000), 0, 1) * ss, w);
      const points: Array<readonly [number, number]> = [
        [0, 0],
        [w - x, 0],
        [w, h / 2],
        [w - x, h],
        [0, h],
      ];
      if (preset === 'chevron') points.push([x, h / 2]);
      return [polygon(points)];
    }
    case 'corner': {
      // §20.1.10.55 — an L: `adj1` is the thickness of the horizontal arm,
      // `adj2` of the vertical one, both against the short side.
      const ss = Math.min(w, h);
      const th = clamp(frac(adjust, 'adj1', 50000), 0, 1) * ss;
      const tw = clamp(frac(adjust, 'adj2', 50000), 0, 1) * ss;
      return [
        polygon([
          [0, 0],
          [w, 0],
          [w, th],
          [tw, th],
          [tw, h],
          [0, h],
        ]),
      ];
    }
    case 'pie':
    case 'blockArc': {
      // Angles are 60 000ths of a degree, clockwise from due east in the
      // y-DOWN frame the spec measures in — negated here, as custGeom's arcs
      // are. A pie closes through its centre; a block arc closes through the
      // inner arc `adj3` names.
      const [rx, ry] = [w / 2, h / 2];
      const start = -((adjust.get('adj1') ?? 0) / 60000) * (Math.PI / 180);
      const end =
        -((adjust.get('adj2') ?? (preset === 'pie' ? 16200000 : 10800000)) / 60000) *
        (Math.PI / 180);
      let sweep = end - start;
      if (sweep > 0) sweep -= 2 * Math.PI;
      const at = (a: number, r: number): readonly [number, number] => [
        rx + rx * r * Math.cos(a),
        ry + ry * r * Math.sin(a),
      ];
      const b = new PathBuilder();
      const [sx0, sy0] = at(start, 1);
      b.moveTo(sx0, sy0).append(arcToBeziers(rx, ry, rx, ry, start, sweep));
      if (preset === 'pie') return [b.lineTo(rx, ry).close().build()];
      const inner = 1 - clamp(frac(adjust, 'adj3', 25000), 0, 1);
      const [ix, iy] = at(start + sweep, inner);
      return [
        b
          .lineTo(ix, iy)
          .append(arcToBeziers(rx, ry, rx * inner, ry * inner, start + sweep, -sweep))
          .close()
          .build(),
      ];
    }
    case 'gear6':
    case 'gear9': {
      // A gear is a circle with N square teeth — the tooth depth and half-width
      // are what `adj1`/`adj2` name, against the short side.
      const teeth = preset === 'gear6' ? 6 : 9;
      const ss = Math.min(w, h);
      const depth = clamp(frac(adjust, 'adj1', 15000), 0, 0.4) * ss;
      const [cx, cy] = [w / 2, h / 2];
      const outer = Math.min(w, h) / 2;
      const root = outer - depth;
      const half = Math.PI / teeth / 2.6; // half a tooth, in radians
      const pts: Array<readonly [number, number]> = [];
      for (let i = 0; i < teeth; i++) {
        const a = (i * 2 * Math.PI) / teeth;
        const gap = Math.PI / teeth;
        pts.push([cx + root * Math.cos(a - gap + half), cy + root * Math.sin(a - gap + half)]);
        pts.push([cx + outer * Math.cos(a - half), cy + outer * Math.sin(a - half)]);
        pts.push([cx + outer * Math.cos(a + half), cy + outer * Math.sin(a + half)]);
        pts.push([cx + root * Math.cos(a + gap - half), cy + root * Math.sin(a + gap - half)]);
      }
      return [polygon(pts)];
    }
    case 'circularArrow':
    case 'leftCircularArrow': {
      // A band of arc with a head on it — the arrow a cycle diagram turns on.
      // Angles are 60 000ths of a degree in the spec's y-DOWN frame, so they
      // negate here; `adj1` is the band's thickness against the short side.
      const ss = Math.min(w, h);
      const t = clamp(frac(adjust, 'adj1', 12500), 0.02, 0.45) * ss;
      // The head stands a thickness proud of the band on each side, so the band
      // is inset by that much: the arrow fills its box and no more.
      const [rx, ry] = [Math.max(t, w / 2 - t), Math.max(t, h / 2 - t)];
      const [cx, cy] = [w / 2, h / 2];
      const deg = (v: number): number => -(v / 60000) * (Math.PI / 180);
      const tail = deg(adjust.get('adj3') ?? (preset === 'circularArrow' ? 20457681 : 12500000));
      const head = deg(adjust.get('adj2') ?? (preset === 'circularArrow' ? 1142319 : 1142319));
      let sweep = head - tail;
      // Keep the shorter way round, in the direction the preset turns.
      const turn = preset === 'circularArrow' ? -1 : 1;
      while (sweep * turn < 0) sweep += turn * 2 * Math.PI;
      while (Math.abs(sweep) > 2 * Math.PI) sweep -= turn * 2 * Math.PI;
      // The head eats the last stretch of the band and stands t/2 proud of it.
      const headSweep = turn * Math.min(Math.abs(sweep) / 3, 0.35);
      const bandEnd = tail + sweep - headSweep;
      // Everything is clamped to the box: an arc's Bézier control points ride a
      // little outside the curve, and a head that overshoots is a head drawn
      // over its neighbour.
      const at = (a: number, r: number): readonly [number, number] => [
        clamp(cx + (rx + r) * Math.cos(a), 0, w),
        clamp(cy + (ry + r) * Math.sin(a), 0, h),
      ];
      const b = new PathBuilder();
      const [ox, oy] = at(tail, t / 2);
      b.moveTo(ox, oy).append(
        arcToBeziers(cx, cy, rx + t / 2, ry + t / 2, tail, bandEnd - tail).map((seg) =>
          seg.op === 'cubic'
            ? {
                op: 'cubic' as const,
                x1: clamp(seg.x1, 0, w),
                y1: clamp(seg.y1, 0, h),
                x2: clamp(seg.x2, 0, w),
                y2: clamp(seg.y2, 0, h),
                x: clamp(seg.x, 0, w),
                y: clamp(seg.y, 0, h),
              }
            : seg,
        ),
      );
      const [hx1, hy1] = at(bandEnd, t);
      const [tipX, tipY] = at(tail + sweep, 0);
      const [hx2, hy2] = at(bandEnd, -t);
      const [ix, iy] = at(bandEnd, -t / 2);
      return [
        b
          .lineTo(hx1, hy1)
          .lineTo(tipX, tipY)
          .lineTo(hx2, hy2)
          .lineTo(ix, iy)
          .append(arcToBeziers(cx, cy, rx - t / 2, ry - t / 2, bandEnd, tail - bandEnd))
          .close()
          .build(),
      ];
    }
    case 'actionButtonBlank':
      return [{ segments: roundRectSegments(w, h, Math.min(w, h) * 0.1) }];
    case 'flowChartMerge':
      return [
        polygon([
          [0, h],
          [w, h],
          [w / 2, 0],
        ]),
      ];
    case 'flowChartExtract':
      return [
        polygon([
          [0, 0],
          [w, 0],
          [w / 2, h],
        ]),
      ];
    case 'flowChartDelay': {
      // A rectangle whose right end is a half-circle.
      const b = new PathBuilder().moveTo(0, 0).lineTo(w / 2, 0);
      return [
        b
          .append(arcToBeziers(w / 2, h / 2, w / 2, h / 2, -Math.PI / 2, Math.PI))
          .lineTo(0, h)
          .close()
          .build(),
      ];
    }
    // The flowchart's stored data is a cylinder too, with a fixed lid.
    case 'flowChartMagneticDisk':
      return cylinder(w, h, h / 6);
    case 'flowChartMagneticTape': {
      // A circle with a short foot out of its bottom right. Drawn as the two
      // shapes it reads as, which keeps every control point inside the box.
      const foot = h / 8;
      return [
        { segments: ellipseSegments(w, h) },
        polygon([
          [w / 2, 0],
          [w, 0],
          [w, foot],
          [w / 2, foot],
        ]),
      ];
    }
    case 'flowChartPunchedTape': {
      // A rectangle with a wave along the top and the bottom.
      const wave = h / 5;
      const b = new PathBuilder().moveTo(0, wave);
      return [
        b
          .cubicTo(w / 4, wave * 2, (w * 3) / 4, 0, w, wave)
          .lineTo(w, h - wave)
          .cubicTo((w * 3) / 4, h - wave * 2, w / 4, h, 0, h - wave)
          .close()
          .build(),
      ];
    }
    case 'wedgeRoundRectCallout': {
      // A rounded rectangle with a tail run out to the point `adj1`/`adj2`
      // name, as fractions of the box measured from its CENTRE.
      const r = clamp(frac(adjust, 'adj3', 16667), 0, 0.5) * Math.min(w, h);
      const tipX = w / 2 + frac(adjust, 'adj1', -20833) * w;
      const tipY = h / 2 - frac(adjust, 'adj2', 62500) * h;
      const [cx, cy] = [w / 2, h / 2];
      const len = Math.hypot(tipX - cx, tipY - cy) || 1;
      const [ux, uy] = [(tipX - cx) / len, (tipY - cy) / len];
      const base = Math.min(w, h) / 6;
      return [
        { segments: roundRectSegments(w, h, r) },
        polygon([
          [cx - uy * base, cy + ux * base],
          [cx + uy * base, cy - ux * base],
          [tipX, tipY],
        ]),
      ];
    }
    case 'leftBracket':
    case 'rightBracket':
    case 'leftBrace':
    case 'rightBrace': {
      // Stroked, never filled: an open path. A bracket turns at its two
      // corners; a brace turns again at the middle, where its point is.
      const brace = preset === 'leftBrace' || preset === 'rightBrace';
      const r = Math.min(clamp(frac(adjust, 'adj1', 8333), 0, 0.5) * h, brace ? w / 2 : w);
      const mirrored = preset === 'rightBracket' || preset === 'rightBrace';
      const x = (v: number): number => (mirrored ? w - v : v);
      // A brace's spine stands a corner in from the edge so its middle point
      // can reach the edge; a bracket's spine IS the edge.
      const spine = brace ? r : 0;
      const b = new PathBuilder()
        .moveTo(x(w), 0)
        .lineTo(x(spine + r), 0)
        .lineTo(x(spine), r);
      if (brace) {
        const mid = clamp(frac(adjust, 'adj2', 50000), 0, 1) * h;
        b.lineTo(x(spine), Math.max(r, mid - r))
          .lineTo(x(0), mid)
          .lineTo(x(spine), Math.min(h - r, mid + r));
      }
      return [
        b
          .lineTo(x(spine), h - r)
          .lineTo(x(spine + r), h)
          .lineTo(x(w), h)
          .build(),
      ];
    }
    default:
      return null;
  }
}

/**
 * Convert a custom geometry (§20.1.9) to a vector path. Path-space coordinates
 * (y-down, origin top-left) are scaled to the shape box and flipped to the local
 * y-up frame; quadratics are elevated to cubics; `arcTo` is decomposed via the
 * shared arc helper. Angles are 1/60000° clockwise in y-down, which becomes
 * negative (CCW) once y is flipped.
 *
 * @param geom The parsed custom geometry (path size + draw commands).
 * @param wPt  Box width in points.
 * @param hPt  Box height in points.
 * @returns A single-element array holding the built path.
 */
export function customPaths(geom: CustomGeometry, wPt: number, hPt: number): Array<VectorPath> {
  const sx = geom.pathWidth > 0 ? wPt / geom.pathWidth : 1;
  const sy = geom.pathHeight > 0 ? hPt / geom.pathHeight : 1;
  const tx = (x: number): number => x * sx;
  const ty = (y: number): number => hPt - y * sy;

  const b = new PathBuilder();
  // Current point in path-space (needed to derive arcTo's ellipse centre).
  let curX = 0;
  let curY = 0;
  for (const cmd of geom.commands) {
    switch (cmd.cmd) {
      case 'move':
        b.moveTo(tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      case 'line':
        b.lineTo(tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      case 'cubic':
        b.cubicTo(tx(cmd.x1), ty(cmd.y1), tx(cmd.x2), ty(cmd.y2), tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      case 'quad': {
        // Elevate (P0=cur, ctrl, P2) to a cubic in path-space, then transform.
        const c1x = curX + (2 / 3) * (cmd.x1 - curX);
        const c1y = curY + (2 / 3) * (cmd.y1 - curY);
        const c2x = cmd.x + (2 / 3) * (cmd.x1 - cmd.x);
        const c2y = cmd.y + (2 / 3) * (cmd.y1 - cmd.y);
        b.cubicTo(tx(c1x), ty(c1y), tx(c2x), ty(c2y), tx(cmd.x), ty(cmd.y));
        curX = cmd.x;
        curY = cmd.y;
        break;
      }
      case 'arc': {
        const st = (cmd.stAng / 60000) * (Math.PI / 180);
        const sw = (cmd.swAng / 60000) * (Math.PI / 180);
        // The current point is the arc start; derive the ellipse centre in
        // path-space, then map to local (negating angles for the y-flip).
        const cxP = curX - cmd.wR * Math.cos(st);
        const cyP = curY - cmd.hR * Math.sin(st);
        b.append(arcToBeziers(tx(cxP), ty(cyP), cmd.wR * sx, cmd.hR * sy, -st, -sw));
        curX = cxP + cmd.wR * Math.cos(st + sw);
        curY = cyP + cmd.hR * Math.sin(st + sw);
        break;
      }
      case 'close':
        b.close();
        break;
    }
  }
  return [b.build()];
}

/** How one corner of a {@link cornerRect} is finished. */
type Corner = { readonly cut: number; readonly round: boolean };
const SQUARE: Corner = { cut: 0, round: false };

// The corners in path order — top-left, top-right, bottom-right, bottom-left —
// as [which adjust guide, rounded or snipped] for each preset. `undefined`
// leaves the corner square.
const CORNER_RECTS: ReadonlyMap<
  string,
  {
    readonly defaults: readonly [number, number];
    readonly corners: readonly [
      0 | 1 | undefined,
      0 | 1 | undefined,
      0 | 1 | undefined,
      0 | 1 | undefined,
    ];
    readonly round: boolean | readonly [boolean, boolean];
  }
> = new Map([
  // One corner: the top-right, the one Word's own preview shows cut.
  [
    'round1Rect',
    { defaults: [16667, 0], corners: [undefined, 0, undefined, undefined], round: true },
  ],
  [
    'snip1Rect',
    { defaults: [16667, 0], corners: [undefined, 0, undefined, undefined], round: false },
  ],
  // Two the same: adj1 the top pair, adj2 the bottom pair.
  ['round2SameRect', { defaults: [16667, 0], corners: [0, 0, 1, 1], round: true }],
  ['snip2SameRect', { defaults: [16667, 0], corners: [0, 0, 1, 1], round: false }],
  // Two diagonal: adj1 top-left and bottom-right, adj2 the other pair.
  ['round2DiagRect', { defaults: [16667, 0], corners: [0, 1, 0, 1], round: true }],
  ['snip2DiagRect', { defaults: [16667, 0], corners: [0, 1, 0, 1], round: false }],
  // One snipped (top-left) and one rounded (top-right).
  [
    'snipRoundRect',
    { defaults: [16667, 16667], corners: [0, 1, undefined, undefined], round: [false, true] },
  ],
]);

// A rectangle whose corners are rounded or cut off by the preset's guides.
function cornerRect(
  preset: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>,
): VectorPath {
  const spec = CORNER_RECTS.get(preset)!;
  const ss = Math.min(w, h);
  // The one-guide presets name it `adj`; the two-guide ones `adj1`/`adj2`.
  const guide = (i: 0 | 1): number => {
    const named =
      adjust.get(i === 0 ? 'adj1' : 'adj2') ?? (i === 0 ? adjust.get('adj') : undefined);
    return clamp((named ?? spec.defaults[i]) / 100000, 0, 0.5) * ss;
  };
  const corners = spec.corners.map((g): Corner => {
    if (g === undefined) return SQUARE;
    const round = typeof spec.round === 'boolean' ? spec.round : spec.round[g];
    return { cut: guide(g), round };
  }) as [Corner, Corner, Corner, Corner];
  return roundedCorners(w, h, corners);
}

// Traverse the box clockwise in the y-up frame — top-left, top-right,
// bottom-right, bottom-left — finishing each corner as its {@link Corner} says.
function roundedCorners(
  w: number,
  h: number,
  [tl, tr, br, bl]: readonly [Corner, Corner, Corner, Corner],
): VectorPath {
  // Corner: its own point, its centre when rounded, and the angle the incoming
  // edge arrives at (each corner sweeps 90° clockwise from there).
  const at: ReadonlyArray<{
    c: Corner;
    from: readonly [number, number];
    to: readonly [number, number];
    centre: readonly [number, number];
    start: number;
  }> = [
    { c: tl, from: [0, h - tl.cut], to: [tl.cut, h], centre: [tl.cut, h - tl.cut], start: Math.PI },
    {
      c: tr,
      from: [w - tr.cut, h],
      to: [w, h - tr.cut],
      centre: [w - tr.cut, h - tr.cut],
      start: Math.PI / 2,
    },
    { c: br, from: [w, br.cut], to: [w - br.cut, 0], centre: [w - br.cut, br.cut], start: 0 },
    { c: bl, from: [bl.cut, 0], to: [0, bl.cut], centre: [bl.cut, bl.cut], start: -Math.PI / 2 },
  ];
  const b = new PathBuilder();
  at.forEach((corner, i) => {
    if (i === 0) b.moveTo(corner.from[0], corner.from[1]);
    else b.lineTo(corner.from[0], corner.from[1]);
    if (corner.c.cut <= 0) return;
    if (corner.c.round) {
      b.append(
        arcToBeziers(
          corner.centre[0],
          corner.centre[1],
          corner.c.cut,
          corner.c.cut,
          corner.start,
          -Math.PI / 2,
        ),
      );
    } else {
      b.lineTo(corner.to[0], corner.to[1]);
    }
  });
  return b.close().build();
}

// §20.1.10.55 star4…star32 — the star family's three published guides. `adj` is
// the inner radius as a fraction of 50000; `hf`/`vf` stretch the vertex ellipse
// so the outermost points land exactly on the box, which is why an odd star
// (star5, star7) needs a different pull horizontally and vertically.
interface StarPreset {
  readonly points: number;
  readonly adj: number;
  readonly hf: number;
  readonly vf: number;
}

const STARS: ReadonlyMap<string, StarPreset> = new Map([
  ['star4', { points: 4, adj: 12500, hf: 100000, vf: 100000 }],
  ['star5', { points: 5, adj: 19098, hf: 105146, vf: 110557 }],
  ['star6', { points: 6, adj: 28868, hf: 115470, vf: 100000 }],
  ['star7', { points: 7, adj: 34601, hf: 102572, vf: 105210 }],
  ['star8', { points: 8, adj: 37500, hf: 100000, vf: 100000 }],
  ['star10', { points: 10, adj: 42533, hf: 105146, vf: 100000 }],
  ['star12', { points: 12, adj: 37500, hf: 100000, vf: 100000 }],
  ['star16', { points: 16, adj: 37500, hf: 100000, vf: 100000 }],
  ['star24', { points: 24, adj: 37500, hf: 100000, vf: 100000 }],
  ['star32', { points: 32, adj: 37500, hf: 100000, vf: 100000 }],
]);

// An n-pointed star: outer vertices every 360/n° starting at the top, inner
// vertices halfway between them at `adj/50000` of the radius. The ellipse is
// (w/2·hf, h/2·vf) about a centre pushed to (w/2, h/2·vf), so `vf` scales the
// whole figure about the box's TOP edge and the topmost point stays at y=0.
function star(
  w: number,
  h: number,
  preset: StarPreset,
  adjust: ReadonlyMap<string, number>,
): VectorPath {
  const inner = clamp(frac(adjust, 'adj', preset.adj) * 2, 0, 1);
  const hf = preset.hf / 100000;
  const vf = preset.vf / 100000;
  const pts: Array<readonly [number, number]> = [];
  for (let i = 0; i < preset.points; i++) {
    const out = -Math.PI / 2 + (i * 2 * Math.PI) / preset.points;
    for (const [angle, r] of [
      [out, 1],
      [out + Math.PI / preset.points, inner],
    ] as const) {
      const x = (w / 2) * (1 + hf * r * Math.cos(angle));
      // Path space is y-down; the local frame is y-up, hence h − y.
      pts.push([x, h - (h / 2) * vf * (1 + r * Math.sin(angle))]);
    }
  }
  return polygon(pts);
}

// Regular n-gon inscribed in the box, first vertex at the top (pointing up).
/**
 * A cylinder standing on its base: the body under its lid, and the lid drawn
 * again on top so its near edge shows.
 *
 * Both paths are wound the SAME way round. Wound against each other a nonzero
 * fill cancels the two and the top comes out hollow, which is how the flowchart
 * disk of tdf114848's fifth page had been drawn all along; and the base bulges
 * DOWN, which the same shape had going up.
 */
function cylinder(w: number, h: number, ry: number): Array<VectorPath> {
  const [rx, cx] = [w / 2, w / 2];
  const body = new PathBuilder()
    .moveTo(0, ry)
    .lineTo(0, h - ry)
    .append(arcToBeziers(cx, h - ry, rx, ry, Math.PI, -Math.PI))
    .lineTo(w, ry)
    .append(arcToBeziers(cx, ry, rx, ry, 0, -Math.PI))
    .close()
    .build();
  const lid = new PathBuilder()
    .moveTo(w, h - ry)
    .append(arcToBeziers(cx, h - ry, rx, ry, 0, -2 * Math.PI))
    .close()
    .build();
  return [body, lid];
}

function regularPolygon(w: number, h: number, n: number, turn = 0): VectorPath {
  const cx = w / 2;
  const cy = h / 2;
  const rx = w / 2;
  const ry = h / 2;
  const pts: Array<readonly [number, number]> = [];
  for (let i = 0; i < n; i++) {
    const a = Math.PI / 2 + turn + (i * 2 * Math.PI) / n;
    pts.push([cx + rx * Math.cos(a), cy + ry * Math.sin(a)]);
  }
  return polygon(pts);
}

// Single-headed block arrow. adj1 = body thickness fraction (default 0.5),
// adj2 = head length fraction (default 0.5).
function blockArrow(
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>,
  dir: 'right' | 'left' | 'up' | 'down',
): VectorPath {
  const thick = clamp(frac(adjust, 'adj1', 50000), 0, 1);
  const head = clamp(frac(adjust, 'adj2', 50000), 0, 1);
  // §20.1.9.18 measures BOTH the shaft and the head against `ss` — the SHORTER
  // side of the box — not against the direction the arrow points. A tall thin
  // `upArrow` 23pt × 156pt therefore has a 12pt head, not a 78pt one:
  // tdf135828_Shape_Rect.xlsx drew a shaft that petered out into a spike where
  // the reference draws a compact triangle on a long stem.
  const ss = Math.min(w, h);
  if (dir === 'right' || dir === 'left') {
    const sb = (h - ss * thick) / 2;
    const st = (h + ss * thick) / 2;
    const hl = head * ss;
    if (dir === 'right') {
      const bx = w - hl;
      return polygon([
        [0, sb],
        [bx, sb],
        [bx, 0],
        [w, h / 2],
        [bx, h],
        [bx, st],
        [0, st],
      ]);
    }
    const bx = hl;
    return polygon([
      [w, sb],
      [bx, sb],
      [bx, 0],
      [0, h / 2],
      [bx, h],
      [bx, st],
      [w, st],
    ]);
  }
  const sl = (w - ss * thick) / 2;
  const sr = (w + ss * thick) / 2;
  const hh = head * ss;
  if (dir === 'up') {
    const by = h - hh;
    return polygon([
      [sl, 0],
      [sr, 0],
      [sr, by],
      [w, by],
      [w / 2, h],
      [0, by],
      [sl, by],
    ]);
  }
  const ty = hh;
  return polygon([
    [sl, h],
    [sr, h],
    [sr, ty],
    [w, ty],
    [w / 2, 0],
    [0, ty],
    [sl, ty],
  ]);
}
