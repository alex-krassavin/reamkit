// The text of a laid-out page, drawn as the PDF draws it: every glyph the
// shaper chose, from the face's own outline, at the point the layout put it.
//
// A `<text>` element hands the characters to the viewer's fonts instead — a
// sans for every family, at the sans's own advances — so a line set in one
// face and measured in another ran past the cell it was laid out to end in, a
// bold word left a gap after it, and an italic header came out upright. The
// outlines are the faces the layout measured: each glyph is defined once, as a
// path in a one-unit em, and placed with the matrix the PDF's text matrix is —
// size, condensing, a faked italic's shear — so the page is the same drawing
// in any viewer, with no font to find.
//
// Everything the PDF emitter draws beside the glyphs is drawn here the same
// way: a run's shading, a comment's highlight, underlines and strikes, faked
// bold, inline pictures and formulas, right-to-left order, justified spaces,
// turned lines and WordArt. The shared decisions (order, justification, one
// matrix or many) come from `@/layout/line-paint`, where both writers read them.

import type {
  FontResource,
  ImageToken,
  LaidOutDocument,
  MathToken,
  TextLineItem,
  TextToken,
} from '@/layout/page-doc';
import type { ImageCrop } from '@/core/document-model';
import type { VectorShape } from '@/core/vector';
import type { OutlineSource } from '@/pdf-reader/glyf-outline';
import { reverseByCodePoint } from '@/core/bidi';
import { warpGlyphMatrix } from '@/core/drawingml/text-warp';
import { sanitizeHref } from '@/core/links';
import { svgPathData } from '@/core/vector';
import { FAUX_ITALIC_SHEAR, computeJustifyExtra, lineVisualOrder } from '@/layout/line-paint';
import { cffOutlineSource, openTypeCff } from '@/pdf-reader/cff-outline';
import { outlineSource } from '@/pdf-reader/glyf-outline';

/** The colour of a comment range's highlight — the PDF emitter's. */
const HIGHLIGHT_HEX = 'fff3a3';

/** The glyph outlines of a page's faces, each defined once and used by id. */
export class SvgGlyphs {
  private readonly sources = new Map<FontResource, OutlineSource | null>();
  private readonly ids = new Map<string, string | null>();
  private readonly defs: Array<string> = [];
  private readonly faceIds = new Map<FontResource, number>();

  /**
   * The id of the path that draws glyph `gid` of `font`, defining it on first
   * use: `null` for a glyph with no ink (a space), `undefined` where the face
   * carries no outlines this reads — the caller then sets the token as text.
   *
   * @param font The face.
   * @param gid  The glyph.
   * @returns The `<path>` id, `null`, or `undefined`.
   */
  use(font: FontResource, gid: number): string | null | undefined {
    const source = this.sourceOf(font);
    if (!source) return undefined;
    let face = this.faceIds.get(font);
    if (face === undefined) {
      face = this.faceIds.size;
      this.faceIds.set(font, face);
    }
    const key = `g${String(face)}-${String(gid)}`;
    const known = this.ids.get(key);
    if (known !== undefined) return known;
    const path = source.path(gid);
    if (!path || path.length === 0) {
      this.ids.set(key, null);
      return null;
    }
    this.defs.push(`<path id="${key}" d="${svgPathData(path, fmtEm)}"/>`);
    this.ids.set(key, key);
    return key;
  }

  /** The `<defs>` holding every glyph used, or nothing when none was. */
  defsXml(): string {
    return this.defs.length > 0 ? `<defs>${this.defs.join('')}</defs>` : '';
  }

  private sourceOf(font: FontResource): OutlineSource | undefined {
    let source = this.sources.get(font);
    if (source === undefined) {
      const raw = font.parsed.raw;
      const cff = font.parsed.tables.has('CFF ') ? openTypeCff(raw) : undefined;
      source = (cff ? cffOutlineSource(cff) : outlineSource(raw)) ?? null;
      this.sources.set(font, source);
    }
    return source ?? undefined;
  }
}

/** What drawing a line needs from the writer around it. */
export interface SvgLineCtx {
  readonly glyphs: SvgGlyphs;
  /** The document-wide counter generated ids are numbered from. */
  readonly ids: { n: number };
  readonly laid: LaidOutDocument;
  /** A picture in a box, as the writer draws one. */
  readonly image: (out: Array<string>, placed: PlacedImage) => void;
  /** A vector shape, as the writer draws one. */
  readonly shape: (out: Array<string>, shape: VectorShape) => void;
}

/** A picture placed in a box on the page — its top-left corner and size, in points. */
export interface PlacedImage {
  readonly resourceName: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly crop?: ImageCrop;
  readonly rotationDeg?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly alpha?: number;
}

/**
 * Draw one laid-out line: what stands behind its glyphs, then the glyphs, a
 * picture or formula wherever one stands in it.
 *
 * @param out  The SVG fragment sink.
 * @param item The line, as the layout placed it.
 * @param ctx  The glyph definitions and the writer's own drawing.
 */
export function emitSvgLine(out: Array<string>, item: TextLineItem, ctx: SvgLineCtx): void {
  const clip = item.clip;
  if (clip) {
    const id = `clip${String(ctx.ids.n++)}`;
    out.push(
      `<clipPath id="${id}"><rect x="${fmt(clip.x)}" y="${fmt(clip.y)}" ` +
        `width="${fmt(clip.width)}" height="${fmt(clip.height)}"/></clipPath>`,
      `<g clip-path="url(#${id})">`,
    );
  }
  if (item.warp) {
    emitWarped(out, item, ctx);
  } else if (item.rotationDeg) {
    // A turned line is set flat inside a frame turned about its own origin —
    // its tabs, stretched spaces and rules turn with it. SVG's y runs down, so
    // a counter-clockwise turn on the page is a negative angle here.
    out.push(
      `<g transform="rotate(${fmt(-item.rotationDeg)} ${fmt(item.originX)} ${fmt(item.baselineY)})">`,
    );
    emitFlat(out, item, ctx);
    out.push('</g>');
  } else {
    emitFlat(out, item, ctx);
  }
  if (clip) out.push('</g>');
}

/** A line set along its baseline, token by token. */
function emitFlat(out: Array<string>, item: TextLineItem, ctx: SvgLineCtx): void {
  const line = item.line;
  const originX: number = item.originX;
  const baseline: number = item.baselineY;
  const extraPerSpace = computeJustifyExtra(line);
  const advance = (tok: (typeof line.tokens)[number]): number =>
    tok.widthPt + (tok.kind === 'text' && tok.isSpace ? extraPerSpace : 0);
  const lineFs = line.maxFontSizePt || 12;
  const boxAscent = Math.max(lineFs, line.mathAscentPt ?? 0);
  const boxDescent = Math.max(lineFs * 0.2, line.mathDescentPt ?? 0);

  // §17.3.2.32 — a run's own background, behind its glyphs; and a comment
  // range's highlight. Both span the line's box, as in the PDF.
  let x = originX;
  for (const tok of line.tokens) {
    const w = advance(tok);
    if (tok.kind === 'text') {
      const hex =
        tok.resolvedRun.shadingColorHex ?? (tok.highlight === true ? HIGHLIGHT_HEX : undefined);
      if (hex !== undefined) {
        out.push(
          `<rect x="${fmt(x)}" y="${fmt(baseline - boxAscent)}" width="${fmt(w)}" ` +
            `height="${fmt(boxAscent + boxDescent)}" fill="#${hex}"/>`,
        );
      }
    }
    x += w;
  }

  // §17.3.2.40 `w:u` / §17.3.2.9 `w:strike` — the rules, drawn before the
  // glyphs as the PDF draws them.
  x = originX;
  for (const tok of line.tokens) {
    const w = advance(tok);
    if (tok.kind === 'text' && (tok.resolvedRun.underline !== 'none' || tok.resolvedRun.strike)) {
      emitDecoration(out, tok, x, w, baseline);
    }
    x += w;
  }

  // The tokens, in the order they stand on the page.
  const hasRtl = line.tokens.some((t) => t.bidiLevel % 2 === 1);
  const order = hasRtl ? lineVisualOrder(line) : line.tokens.map((_, i) => i);
  x = originX;
  for (const i of order) {
    const tok = line.tokens[i]!;
    if (tok.kind === 'image') emitImageToken(out, tok, x, baseline, ctx);
    else if (tok.kind === 'math') emitMathToken(out, tok, x, baseline, ctx);
    else emitTextToken(out, tok, x, baseline, ctx);
    x += advance(tok);
  }

  // The words themselves, unseen, over the glyphs that draw them: outlines
  // are pictures of text, and a page of them can be neither searched nor
  // selected nor read aloud. One element a line, its characters in reading
  // order, stretched to the line's width so a selection lands on the glyphs —
  // the way a PDF viewer lays its text over a page it has drawn.
  const words = line.tokens.map((t) => (t.kind === 'text' ? t.text : '')).join('');
  if (words.trim().length > 0) {
    out.push(
      `<text x="${fmt(originX)}" y="${fmt(baseline)}" font-size="${fmt(lineFs)}" ` +
        `textLength="${fmt(x - originX)}" lengthAdjust="spacingAndGlyphs" fill-opacity="0" ` +
        `xml:space="preserve">${escapeXml(words)}</text>`,
    );
  }
}

/**
 * One run's glyphs from `x` along the baseline: the shaped glyphs of its text
 * (reversed for a right-to-left run, as the PDF shows it), each at the advance
 * the shaper gave it plus the run's letter spacing.
 */
function emitTextToken(
  out: Array<string>,
  tok: TextToken,
  x: number,
  baseline: number,
  ctx: SvgLineCtx,
): void {
  if (tok.text.length === 0) return;
  const size = tok.fontSizePt;
  // §9.3.3's horizontal scaling: a condensed face drawn from a regular one, and
  // a run the document sets narrower than its face (§17.3.2.43).
  const scale = (tok.synthetic?.widthScale ?? 1) * (tok.resolvedRun.widthScale ?? 1);
  const slant = tok.synthetic?.italic === true ? FAUX_ITALIC_SHEAR : 0;
  const rise = tok.risePt ?? 0;
  const text = tok.bidiLevel % 2 === 1 ? reverseByCodePoint(tok.text) : tok.text;
  const run = tok.font.measure.glyphRun(text);
  const upem = tok.font.parsed.unitsPerEm;
  const spacingEm = (tok.resolvedRun.letterSpacingPt ?? 0) / (size * scale);
  const uses: Array<string> = [];
  let pen = 0;
  for (let g = 0; g < run.gids.length; g++) {
    const id = ctx.glyphs.use(tok.font, run.gids[g]!);
    if (id === undefined) {
      emitAsText(out, tok, x, baseline);
      return;
    }
    if (id !== null) uses.push(`<use href="#${id}"${pen !== 0 ? ` x="${fmtEm(pen)}"` : ''}/>`);
    pen += (run.advances[g] ?? 0) / upem + spacingEm;
  }
  if (uses.length === 0) return;
  // ISO 32000-1 §9.3.6 — the faked bold (a stroke a thirtieth of the em wide,
  // in the text's colour) and §21.1.2.3.9's outline (one of its own).
  const outline = tok.resolvedRun.textOutline;
  const strokeWidth =
    outline !== undefined ? outline.widthPt : tok.synthetic?.bold === true ? size * 0.03 : 0;
  const stroke =
    strokeWidth > 0
      ? ` stroke="#${outline?.colorHex ?? tok.resolvedRun.colorHex}" stroke-width="${fmtEm(strokeWidth / size)}"`
      : '';
  // The PDF's text matrix in the page's y-down frame: the em scaled to the
  // size (and condensed), y flipped, leaned for a faked italic, set on the
  // baseline lifted by the run's rise.
  const group =
    `<g transform="matrix(${fmt(size * scale)} 0 ${fmt(slant * size)} ${fmt(-size)} ` +
    `${fmt(x)} ${fmt(baseline - rise)})" fill="#${tok.resolvedRun.colorHex}"${stroke}>` +
    `${uses.join('')}</g>`;
  const href = tok.href !== undefined ? sanitizeHref(tok.href) : undefined;
  out.push(href !== undefined ? `<a href="${escapeXml(href)}">${group}</a>` : group);
}

/** A run of a face whose outlines cannot be read: its characters, for the viewer's fonts to set. */
function emitAsText(out: Array<string>, tok: TextToken, x: number, baseline: number): void {
  if (tok.text.trim().length === 0) return;
  out.push(
    `<text x="${fmt(x)}" y="${fmt(baseline - (tok.risePt ?? 0))}" font-family="sans-serif" ` +
      `font-size="${fmt(tok.fontSizePt)}" fill="#${tok.resolvedRun.colorHex}">${escapeXml(tok.text)}</text>`,
  );
}

/** A run's underline and strikethrough, where the PDF emitter draws them. */
function emitDecoration(
  out: Array<string>,
  tok: TextToken,
  x: number,
  w: number,
  baseline: number,
): void {
  const size = tok.fontSizePt;
  const thickness = Math.max(0.4, size * 0.06);
  const style = tok.resolvedRun.underline;
  if (style !== 'none') {
    // §17.3.2.40 — the rule carries its own colour when the run gives one;
    // §17.18.99 — a heavy style is twice the weight, a double is two rules.
    const hex = tok.resolvedRun.underlineColorHex ?? tok.resolvedRun.colorHex;
    const heavy = style === 'thick' || style === 'dottedHeavy' || style === 'dashHeavy';
    const t = heavy ? thickness * 2 : thickness;
    const top = baseline + size * 0.12;
    const dashed =
      style === 'dotted' || style === 'dottedHeavy'
        ? [t, t * 2]
        : style === 'dash' || style === 'dashHeavy'
          ? [t * 4, t * 3]
          : undefined;
    if (dashed) {
      out.push(
        `<line x1="${fmt(x)}" y1="${fmt(top + t / 2)}" x2="${fmt(x + w)}" y2="${fmt(top + t / 2)}" ` +
          `stroke="#${hex}" stroke-width="${fmt(t)}" stroke-dasharray="${dashed.map(fmt).join(' ')}"/>`,
      );
    } else {
      out.push(
        `<rect x="${fmt(x)}" y="${fmt(top)}" width="${fmt(w)}" height="${fmt(t)}" fill="#${hex}"/>`,
      );
      if (style === 'double') {
        out.push(
          `<rect x="${fmt(x)}" y="${fmt(top + t * 2)}" width="${fmt(w)}" height="${fmt(t)}" fill="#${hex}"/>`,
        );
      }
    }
  }
  if (tok.resolvedRun.strike) {
    out.push(
      `<rect x="${fmt(x)}" y="${fmt(baseline - size * 0.26 - thickness)}" width="${fmt(w)}" ` +
        `height="${fmt(thickness)}" fill="#${tok.resolvedRun.colorHex}"/>`,
    );
  }
}

/**
 * An inline picture standing on the baseline, its bottom edge on it — or a
 * metafile, which draws itself from that corner in a frame of its own (y up).
 */
function emitImageToken(
  out: Array<string>,
  tok: ImageToken,
  x: number,
  baseline: number,
  ctx: SvgLineCtx,
): void {
  const b = tok.drawBox;
  const left = x + (b?.dxPt ?? 0);
  // The box's bottom edge: the baseline, raised by the box's own offset.
  const bottom = baseline - (b?.dyPt ?? 0);
  const width = b?.widthPt ?? tok.widthPt;
  const height = b?.heightPt ?? tok.heightPt;
  const meta = tok.metafile;
  if (meta) {
    for (const img of meta.images ?? []) {
      if (img.resourceName === '') continue;
      ctx.image(out, {
        resourceName: img.resourceName,
        x: left + img.x,
        y: bottom - img.y - img.height,
        width: img.width,
        height: img.height,
        ...(img.rotationDeg !== undefined ? { rotationDeg: img.rotationDeg } : {}),
      });
    }
    // The metafile's frame is y-up from the box's bottom-left corner.
    for (const sh of meta.shapes) {
      ctx.shape(out, { ...sh, transform: [1, 0, 0, -1, left, bottom] });
    }
    for (const t of meta.texts) {
      const first = t.line.tokens.find((k) => k.kind === 'text');
      if (first?.kind !== 'text') continue;
      emitTextToken(out, first, left + t.x, bottom - t.y, ctx);
    }
    return;
  }
  if (!tok.imageResourceName) return;
  // §20.1.8.40 — the shadow it casts, under it: the same box, offset, softened.
  const shadow = tok.shadow;
  if (shadow) {
    const id = `shadow${String(ctx.ids.n++)}`;
    const blur = shadow.blurPt > 0;
    if (blur) {
      out.push(
        `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%">` +
          `<feGaussianBlur stdDeviation="${fmt(shadow.blurPt / 2)}"/></filter>`,
      );
    }
    out.push(
      `<rect x="${fmt(left + shadow.dxPt)}" y="${fmt(bottom - height + shadow.dyPt)}" ` +
        `width="${fmt(width)}" height="${fmt(height)}" fill="#${shadow.colorHex}" ` +
        `opacity="${fmt(shadow.alpha)}"${blur ? ` filter="url(#${id})"` : ''}/>`,
    );
  }
  ctx.image(out, {
    resourceName: tok.imageResourceName,
    x: left,
    y: bottom - height,
    width,
    height,
    ...(tok.crop ? { crop: tok.crop } : {}),
    ...(tok.rotationDeg !== undefined ? { rotationDeg: tok.rotationDeg } : {}),
    ...(tok.flipH === true ? { flipH: true } : {}),
    ...(tok.flipV === true ? { flipV: true } : {}),
  });
  // §20.1.2.2.24 — the picture's own frame.
  const frame = tok.outline;
  if (frame) {
    out.push(
      `<rect x="${fmt(left)}" y="${fmt(bottom - height)}" width="${fmt(width)}" height="${fmt(height)}" ` +
        `fill="none" stroke="#${frame.colorHex}" stroke-width="${fmt(frame.widthPt)}"/>`,
    );
  }
}

/** An inline formula: its glyphs, rules and paths, placed from its origin on the baseline (y up). */
function emitMathToken(
  out: Array<string>,
  tok: MathToken,
  x: number,
  baseline: number,
  ctx: SvgLineCtx,
): void {
  for (const it of tok.items) {
    if (it.kind === 'glyph') {
      const run = it.font.measure.glyphRun(it.text);
      const upem = it.font.parsed.unitsPerEm;
      const uses: Array<string> = [];
      let pen = 0;
      for (let g = 0; g < run.gids.length; g++) {
        const id = ctx.glyphs.use(it.font, run.gids[g]!);
        if (id) uses.push(`<use href="#${id}"${pen !== 0 ? ` x="${fmtEm(pen)}"` : ''}/>`);
        pen += (run.advances[g] ?? 0) / upem;
      }
      if (uses.length > 0) {
        out.push(
          `<g transform="matrix(${fmt(it.sizePt)} 0 0 ${fmt(-it.sizePt)} ` +
            `${fmt(x + it.x)} ${fmt(baseline - it.y)})" fill="#000000">${uses.join('')}</g>`,
        );
      }
    } else if (it.kind === 'rule') {
      out.push(
        `<rect x="${fmt(x + it.x)}" y="${fmt(baseline - it.y - it.h)}" width="${fmt(it.w)}" ` +
          `height="${fmt(it.h)}" fill="#000000"/>`,
      );
    } else {
      ctx.shape(out, {
        paths: [{ segments: it.segments }],
        ...(it.fill === true ? { fillColorHex: '000000' } : {}),
        ...(it.strokeWidthPt !== undefined
          ? { stroke: { colorHex: '000000', widthPt: it.strokeWidthPt } }
          : {}),
        transform: [1, 0, 0, -1, x, baseline],
      });
    }
  }
}

/**
 * §20.1.9.10 — WordArt: every glyph at its own point on the warp's curve, each
 * with its own matrix. The token's glyphs are rescaled to the width it was
 * measured at, as the PDF does, so the block still ends where the layout said.
 */
function emitWarped(out: Array<string>, item: TextLineItem, ctx: SvgLineCtx): void {
  const warp = item.warp!;
  let cursor: number = item.originX;
  for (const tok of item.line.tokens) {
    if (tok.kind !== 'text' || tok.isSpace) {
      cursor += tok.widthPt;
      continue;
    }
    const size = tok.fontSizePt;
    const chars = [...tok.text];
    const raw = chars.map((c) => tok.font.measure.textWidthPt(c, size));
    const total = raw.reduce((a, b) => a + b, 0);
    const k = total > 0 ? tok.widthPt / total : 1;
    for (let i = 0; i < chars.length; i++) {
      const adv = raw[i]! * k;
      const m = warpGlyphMatrix(warp, cursor, adv, item.baselineY);
      cursor += adv;
      if (!m) continue;
      const run = tok.font.measure.glyphRun(chars[i]!);
      const upem = tok.font.parsed.unitsPerEm;
      const uses: Array<string> = [];
      let pen = 0;
      for (let g = 0; g < run.gids.length; g++) {
        const id = ctx.glyphs.use(tok.font, run.gids[g]!);
        if (id) uses.push(`<use href="#${id}"${pen !== 0 ? ` x="${fmtEm(pen)}"` : ''}/>`);
        pen += (run.advances[g] ?? 0) / upem;
      }
      if (uses.length === 0) continue;
      // The warp's matrix maps the glyph's text space (y up) into the page's
      // top-left frame; the em is the size.
      out.push(
        `<g transform="matrix(${fmt(m[0] * size)} ${fmt(m[1] * size)} ${fmt(m[2] * size)} ` +
          `${fmt(m[3] * size)} ${fmt(m[4])} ${fmt(m[5])})" fill="#${tok.resolvedRun.colorHex}">` +
          `${uses.join('')}</g>`,
      );
    }
  }
}

function escapeXml(s: string): string {
  return s
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

/** A page coordinate, to the hundredth of a point. */
function fmt(n: number): string {
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
}

/** A coordinate in a one-unit em, to a ten-thousandth of it. */
function fmtEm(n: number): string {
  const r = Math.round(n * 10000) / 10000;
  return Object.is(r, -0) ? '0' : String(r);
}
