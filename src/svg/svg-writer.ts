// SVG writer (ir-design §7 / stage 6): the third adapter, written purely
// against the PageDoc schema (LaidOutDocument/PageItem). It deliberately
// knows nothing about OOXML or the PDF writer.
//
// PageItem coordinates are top-left/y-down (the frame the schema froze on at
// stage 6.4) — exactly SVG's own frame, so this writer emits them verbatim;
// it is the PDF emitter that converts into PDF's y-up frame at emission.
//
// Pages stack vertically in one <svg>, separated by a gap — a faithful,
// dependency-free preview of the laid-out document. Its text is drawn from
// the faces' own outlines (see svg-text.ts), so it needs no fonts to show.

import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { Loss } from '@/core/ir';
import type { ImageItem, LaidOutDocument, LaidOutPage, PageItem } from '@/layout/page-doc';
import type { PathSegment, VectorShape } from '@/core/vector';
import type { PlacedImage, SvgLineCtx } from '@/svg/svg-text';
import { svgPathData } from '@/core/vector';

import { FEATURES } from '@/core/ir';
import { toBase64 } from '@/core/bytes';
import { gradientSvgDef } from '@/core/drawingml/shape-render';
import { BORDER_DASHES, washVeil } from '@/layout/line-paint';
import { pageLayers, paintPlan } from '@/layout/page-doc';
import { SvgGlyphs, emitSvgLine } from '@/svg/svg-text';

const PAGE_GAP = 12;

/** Options for {@link writeSvg}. */
export interface SvgWriteOptions {
  /** Gap between stacked pages, in px/pt (default 12). */
  readonly pageGap?: number;
}

/**
 * Render a {@link LaidOutDocument} to a single SVG preview (ir-design §7 /
 * stage 6). Written purely against the PageDoc schema; knows nothing about OOXML
 * or the PDF writer. PageItem coordinates are already top-left/y-down — SVG's own
 * frame — so they are emitted verbatim. Pages stack vertically in one `<svg>`,
 * separated by {@link SvgWriteOptions.pageGap}, each on a white outlined rect so
 * they read as pages.
 *
 * @param laid The laid-out, paginated document from the layout engine.
 * @param opts Optional page-gap override.
 * @returns The encoded SVG bytes plus the recorded {@link Loss} list.
 */
export function writeSvg(laid: LaidOutDocument, opts: SvgWriteOptions = {}): WriteResult {
  const gap = opts.pageGap ?? PAGE_GAP;
  const losses: Array<Loss> = [];
  const width = Math.max(1, ...laid.pages.map((p) => p.width));
  const height =
    laid.pages.reduce((s, p) => s + p.height, 0) + gap * Math.max(0, laid.pages.length - 1);

  const parts: Array<string> = [];
  const ids = { n: 0 }; // unique generated-id counter across the whole document
  const glyphs = new SvgGlyphs();
  const ctx: SvgLineCtx = {
    glyphs,
    ids,
    laid,
    image: (out, placed) => emitPlaced(out, placed, laid, ids),
    shape: (out, shape) => emitShape(out, shape, ids),
  };
  let yOffset = 0;
  laid.pages.forEach((page, i) => {
    parts.push(`<g transform="translate(0 ${fmt(yOffset)})" data-page="${i + 1}">`);
    // Page background + outline so stacked pages read as pages.
    parts.push(
      `<rect x="0" y="0" width="${fmt(page.width)}" height="${fmt(page.height)}" fill="#ffffff" stroke="#cccccc"/>`,
    );
    emitPage(parts, page, ctx);
    parts.push('</g>');
    yOffset += page.height + gap;
  });
  parts.push('</svg>');
  // The glyphs the pages used, each once — known only now that they are drawn.
  parts.unshift(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(width)}" height="${fmt(height)}" viewBox="0 0 ${fmt(width)} ${fmt(height)}">`,
    glyphs.defsXml(),
  );

  return { bytes: new TextEncoder().encode(parts.join('\n')), losses };
}

/**
 * The page-medium {@link DocumentWriter} adapter (id `'svg'`), wrapping
 * {@link writeSvg}, with the set of {@link FEATURES} it renders.
 */
export const svgWriter: DocumentWriter<LaidOutDocument> = {
  id: 'svg',
  consumes: 'page',
  supports: new Set([FEATURES.text, FEATURES.tables, FEATURES.images, FEATURES.shapes]),
  write: (doc, opts) => writeSvg(doc, opts ?? {}),
};

// A page in its layers (pageLayers): its own items, then the ones that stand
// IN FRONT of its text (PageItemBase.over), every pass again over what the
// first left — and each group seen through a window of its own inside a clip.
function emitPage(out: Array<string>, page: LaidOutPage, ctx: SvgLineCtx): void {
  for (const layer of pageLayers(page.commands)) {
    const w = layer.window;
    if (!w) {
      emitLayer(out, layer.items, ctx);
      continue;
    }
    const id = `win${String(ctx.ids.n++)}`;
    out.push(
      `<clipPath id="${id}"><rect x="${fmt(w.x)}" y="${fmt(w.y)}" width="${fmt(w.width)}" height="${fmt(w.height)}"/></clipPath>`,
      `<g clip-path="url(#${id})">`,
    );
    emitLayer(out, layer.items, ctx);
    out.push('</g>');
  }
}

function emitLayer(out: Array<string>, commands: ReadonlyArray<PageItem>, ctx: SvgLineCtx): void {
  // The shared canonical paint order — one owner for every writer. PageItem
  // coordinates are already top-left/y-down: SVG's native frame.
  const plan = paintPlan(commands);

  // What the page puts behind its content, in its own order.
  for (const item of plan.behind) emitPageItem(out, item, ctx);

  for (const f of plan.fills) {
    out.push(
      `<rect x="${fmt(f.x)}" y="${fmt(f.y)}" width="${fmt(f.width)}" height="${fmt(f.height)}" fill="#${f.fillColorHex}"/>`,
    );
  }

  for (const img of plan.images) emitImage(out, img, ctx);

  for (const b of plan.borders) {
    const x2 = b.x + b.width;
    const yTop = b.y;
    const yBottom = b.y + b.height;
    const [ax, ay, bx, by] =
      b.side === 'top'
        ? [b.x, yTop, x2, yTop]
        : b.side === 'bottom'
          ? [b.x, yBottom, x2, yBottom]
          : b.side === 'left'
            ? [b.x, yTop, b.x, yBottom]
            : [x2, yTop, x2, yBottom];
    // §18.18.3's dashed/dotted are line patterns, not weights — the PDF's own.
    const dash = b.borderStyle !== undefined ? BORDER_DASHES.get(b.borderStyle) : undefined;
    out.push(
      `<line x1="${fmt(ax)}" y1="${fmt(ay)}" x2="${fmt(bx)}" y2="${fmt(by)}" stroke="#${b.borderColorHex}" stroke-width="${fmt(b.borderSizePt)}"${
        dash ? ` stroke-dasharray="${dash.map(fmt).join(' ')}"` : ''
      }/>`,
    );
  }

  for (const sh of plan.shapes) {
    emitShape(out, sh.shape, ctx.ids);
  }

  // A picture paints as one thing, in its own order: a metafile buries a label
  // under a panel it draws afterwards, and the passes above would lift it out.
  for (const picture of plan.pictures) {
    for (const item of picture) emitPageItem(out, item, ctx);
  }

  // …and a page that states its own paint order paints in it (a slide's shape
  // tree IS that order), kind by kind ignored.
  for (const run of plan.ordered) {
    for (const item of run) emitPageItem(out, item, ctx);
  }

  for (const t of plan.lines) emitSvgLine(out, t, ctx);
}

/**
 * One picture item, its placement and its effects: §20.1.8.23 `a:duotone`
 * recolours it between its dark end and its light end — luminance into every
 * channel, then a two-entry transfer table per channel — and a picture FILL
 * (§20.1.8.14) is clipped to the outline of the shape it fills.
 *
 * @param out  The SVG fragment sink.
 * @param img  The image item.
 * @param ctx  The document, its id counter, its drawing.
 */
function emitImage(out: Array<string>, img: ImageItem, ctx: SvgLineCtx): void {
  let filter: string | undefined;
  if (img.duotone) {
    const id = `duo${String(ctx.ids.n++)}`;
    const chan = (hex: string, at: number): number => parseInt(hex.slice(at, at + 2), 16) / 255;
    const table = (at: number): string =>
      `${fmt(chan(img.duotone!.shadowHex, at))} ${fmt(chan(img.duotone!.highlightHex, at))}`;
    out.push(
      `<filter id="${id}" color-interpolation-filters="sRGB">` +
        `<feColorMatrix type="matrix" values="0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0"/>` +
        `<feComponentTransfer>` +
        `<feFuncR type="table" tableValues="${table(0)}"/>` +
        `<feFuncG type="table" tableValues="${table(2)}"/>` +
        `<feFuncB type="table" tableValues="${table(4)}"/>` +
        `</feComponentTransfer></filter>`,
    );
    filter = id;
  }
  let clipped = false;
  if (img.clip) {
    const id = `pic${String(ctx.ids.n++)}`;
    const [a, b, c, d, e, f] = img.clip.transform;
    out.push(
      `<clipPath id="${id}">` +
        img.clip.paths
          .map(
            (p) =>
              `<path d="${pathData(p.segments)}" transform="matrix(${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ${fmt(e)} ${fmt(f)})"/>`,
          )
          .join('') +
        `</clipPath>`,
      `<g clip-path="url(#${id})">`,
    );
    clipped = true;
  }
  emitPlaced(
    out,
    {
      resourceName: img.imageResourceName,
      x: img.x,
      y: img.y,
      width: img.width,
      height: img.height,
      ...(img.crop ? { crop: img.crop } : {}),
      ...(img.rotationDeg !== undefined ? { rotationDeg: img.rotationDeg } : {}),
      ...(img.flipH === true ? { flipH: true } : {}),
      ...(img.flipV === true ? { flipV: true } : {}),
      ...(img.alpha !== undefined ? { alpha: img.alpha } : {}),
    },
    ctx.laid,
    ctx.ids,
    filter,
    img.wash,
  );
  if (clipped) out.push('</g>');
}

/**
 * A picture in its box, as the PDF places one: turned about the box's centre
 * (clockwise, as DrawingML measures), mirrored about it, and — cropped by
 * `a:srcRect` (§20.1.8.55) — drawn larger than the box with the box as its clip,
 * so the part that was kept fills it. A wash (§14.1.2.10) is the veil laid over
 * it.
 */
function emitPlaced(
  out: Array<string>,
  placed: PlacedImage,
  laid: LaidOutDocument,
  ids: { n: number },
  filter?: string,
  wash?: { readonly gain: number; readonly black: number },
): void {
  const href = imageHref(placed.resourceName, laid);
  if (!href) return;
  const { x, y, width, height } = placed;
  const cx = x + width / 2;
  const cy = y + height / 2;
  const turns: Array<string> = [];
  if (placed.rotationDeg) turns.push(`rotate(${fmt(placed.rotationDeg)} ${fmt(cx)} ${fmt(cy)})`);
  if (placed.flipH === true || placed.flipV === true) {
    const sx = placed.flipH === true ? -1 : 1;
    const sy = placed.flipV === true ? -1 : 1;
    turns.push(
      `matrix(${String(sx)} 0 0 ${String(sy)} ${fmt(sx === -1 ? 2 * cx : 0)} ${fmt(sy === -1 ? 2 * cy : 0)})`,
    );
  }
  const open = turns.length > 0 ? `<g transform="${turns.join(' ')}">` : '';
  if (open) out.push(open);
  const crop = placed.crop;
  let clipId: string | undefined;
  let drawn = { x, y, width, height };
  if (crop) {
    const keptW = 1 - crop.left - crop.right;
    const keptH = 1 - crop.top - crop.bottom;
    if (keptW > 0 && keptH > 0) {
      const fullW = width / keptW;
      const fullH = height / keptH;
      drawn = { x: x - crop.left * fullW, y: y - crop.top * fullH, width: fullW, height: fullH };
      clipId = `crop${String(ids.n++)}`;
      out.push(
        `<clipPath id="${clipId}"><rect x="${fmt(x)}" y="${fmt(y)}" width="${fmt(width)}" height="${fmt(height)}"/></clipPath>`,
      );
    }
  }
  const alpha =
    placed.alpha !== undefined && placed.alpha < 1 ? ` opacity="${fmt(placed.alpha)}"` : '';
  out.push(
    `<image x="${fmt(drawn.x)}" y="${fmt(drawn.y)}" width="${fmt(drawn.width)}" height="${fmt(drawn.height)}" href="${href}" preserveAspectRatio="none"` +
      `${clipId ? ` clip-path="url(#${clipId})"` : ''}${filter ? ` filter="url(#${filter})"` : ''}${alpha}/>`,
  );
  const veil = washVeil(wash);
  if (veil) {
    const grey = Math.round(veil.grey * 255)
      .toString(16)
      .padStart(2, '0');
    out.push(
      `<rect x="${fmt(x)}" y="${fmt(y)}" width="${fmt(width)}" height="${fmt(height)}" fill="#${grey}${grey}${grey}" opacity="${fmt(veil.alpha)}"/>`,
    );
  }
  if (open) out.push('</g>');
}

/**
 * One page item, whatever its kind — for the runs that paint in their own order
 * (what stands behind the content, and a picture's own primitives) rather than
 * in the by-kind passes.
 *
 * @param out  The SVG fragment sink.
 * @param item The item to draw.
 * @param ctx  The document, its glyphs, its id counter and its drawing.
 */
function emitPageItem(out: Array<string>, item: PageItem, ctx: SvgLineCtx): void {
  if (item.type === 'shape') {
    emitShape(out, item.shape, ctx.ids);
    return;
  }
  if (item.type === 'line') {
    emitSvgLine(out, item, ctx);
    return;
  }
  if (item.type === 'image') {
    emitImage(out, item, ctx);
    return;
  }
  if (item.type === 'fill') {
    out.push(
      `<rect x="${fmt(item.x)}" y="${fmt(item.y)}" width="${fmt(item.width)}" height="${fmt(item.height)}" fill="#${item.fillColorHex}"/>`,
    );
  }
}

function emitShape(out: Array<string>, shape: VectorShape, ids: { n: number }): void {
  const [a, b, c, d, e, f] = shape.transform;
  // The stored CTM maps the shape's local y-up frame straight into the
  // top-left page frame — SVG's matrix() convention verbatim.
  const transform = `matrix(${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ${fmt(e)} ${fmt(f)})`;
  let fill: string;
  if (shape.fillGradient) {
    const id = `grad${String(ids.n++)}`;
    out.push(gradientSvgDef(id, shape.fillGradient));
    fill = `url(#${id})`;
  } else {
    fill = shape.fillColorHex ? `#${shape.fillColorHex}` : 'none';
  }
  // §20.1.2.3.1 — a fill the document made transparent lets what is behind it
  // show through.
  const fillAlpha =
    shape.fillAlpha !== undefined && shape.fillAlpha < 1 && fill !== 'none'
      ? ` fill-opacity="${fmt(shape.fillAlpha)}"`
      : '';
  // The stroke in the shape's local units, as the PDF sets its width after
  // the same matrix — with its caps, joins and dashes.
  const s = shape.stroke;
  const stroke = s
    ? ` stroke="#${s.colorHex}" stroke-width="${fmt(s.widthPt)}"` +
      (s.cap !== undefined ? ` stroke-linecap="${s.cap}"` : '') +
      (s.join !== undefined ? ` stroke-linejoin="${s.join}"` : '') +
      (s.dash && s.dash.length > 0 ? ` stroke-dasharray="${s.dash.map(fmt).join(' ')}"` : '')
    : '';
  // §20.1.8.40 — the drop shadow, drawn first so the shape lands on top of it.
  // SVG can blur, so here the softness the source asked for is the softness
  // drawn: `blurRad` is the full spread, and a Gaussian's is about 2σ.
  const shadow = shape.shadow;
  if (shadow) {
    const id = `shadow${String(ids.n++)}`;
    if (shadow.blurPt > 0) {
      out.push(
        `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%">` +
          `<feGaussianBlur stdDeviation="${fmt(shadow.blurPt / 2)}"/></filter>`,
      );
    }
    // The stored CTM lands in the y-DOWN page frame, so a shadow that falls
    // down the page moves in +y.
    const shTransform =
      `matrix(${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ` +
      `${fmt(e + shadow.dxPt)} ${fmt(f + shadow.dyPt)})`;
    out.push(
      `<g opacity="${fmt(shadow.alpha)}"${shadow.blurPt > 0 ? ` filter="url(#${id})"` : ''}>`,
    );
    for (const path of shape.paths) {
      const rule = path.fillRule === 'evenodd' ? ' fill-rule="evenodd"' : '';
      out.push(
        `<path d="${pathData(path.segments)}" fill="#${shadow.colorHex}"${rule}` +
          ` transform="${shTransform}"/>`,
      );
    }
    out.push('</g>');
  }
  for (const path of shape.paths) {
    const d2 = pathData(path.segments);
    const rule = path.fillRule === 'evenodd' ? ' fill-rule="evenodd"' : '';
    out.push(
      `<path d="${d2}" fill="${fill}"${fillAlpha}${rule}${stroke} transform="${transform}"/>`,
    );
  }
}

// Local path coordinates are y-up; the transform (page flip composed in by
// layout) maps them into the y-down page frame. Emit the raw coordinates.
function pathData(segments: ReadonlyArray<PathSegment>): string {
  return svgPathData(segments, fmt);
}

function imageHref(resourceName: string, laid: LaidOutDocument): string | undefined {
  if (!resourceName) return undefined;
  // resourceName (ImN) → ResourceId → bytes from the content-addressed store;
  // the mime comes from the layout-time prepare (the image expert) instead of
  // re-sniffing the bytes here.
  for (const [resourceId, res] of laid.imageResources) {
    if (res.resourceName !== resourceName) continue;
    const bytes = laid.resources.get(resourceId);
    if (!bytes || !res.prepared) return undefined;
    return `data:${res.prepared.mimeType};base64,${toBase64(bytes)}`;
  }
  return undefined;
}

function fmt(n: number): string {
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
}
