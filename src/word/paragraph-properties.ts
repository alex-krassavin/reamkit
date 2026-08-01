// ECMA-376 Part 1 §17.3.1 — Paragraph Properties (pPr).

import type {
  Alignment,
  Border,
  BorderStyle,
  CellBorders,
  CellShading,
  FrameProperties,
  ParagraphProperties,
  TabStop,
} from '@/core/document-model';
import type { Pt } from '@/core/ir';
import { eighthPtToPt, pt, twipsToPt } from '@/core/ir';

import { parseRunProperties } from '@/word/run-properties';
import { asArray, asElement, getAttr, getVal, parseIntAttr, parseToggle } from '@/word/xml-helpers';

const ALIGNMENTS = new Set<Alignment>(['left', 'right', 'center', 'both', 'distribute']);
const LINE_RULES = new Set<'auto' | 'exact' | 'atLeast'>(['auto', 'exact', 'atLeast']);

const X_ALIGNS = new Set(['left', 'center', 'right', 'inside', 'outside']);
const Y_ALIGNS = new Set(['top', 'center', 'bottom', 'inside', 'outside']);
const ANCHORS = new Set(['text', 'margin', 'page']);
const WRAPS = new Set(['auto', 'around', 'none', 'notBeside', 'tight', 'through']);

/**
 * §17.3.1.11 `w:framePr` — the box a floating text frame takes and where it
 * hangs. Lengths are twips; the alignments and anchors are read only when they
 * name a value the spec defines.
 *
 * @param node The `w:framePr` element in flat shape.
 * @returns The frame, or `undefined` when the element says nothing usable.
 */
function parseFramePr(node: unknown): FrameProperties | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  const twips = (name: string): Pt | undefined => {
    const v = parseIntAttr(node, name);
    return v === undefined ? undefined : twipsToPt(v);
  };
  const pick = <T extends string>(name: string, allowed: ReadonlySet<string>): T | undefined => {
    const v = getAttr(node, name);
    return v !== undefined && allowed.has(v) ? (v as T) : undefined;
  };
  const rule = getAttr(node, 'hRule');
  const out: Mutable<FrameProperties> = {};
  const w = twips('w');
  const h = twips('h');
  const x = twips('x');
  const y = twips('y');
  const hSpace = twips('hSpace');
  const vSpace = twips('vSpace');
  if (w !== undefined) out.widthPt = w;
  if (h !== undefined) out.heightPt = h;
  if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') out.heightRule = rule;
  if (x !== undefined) out.xPt = x;
  if (y !== undefined) out.yPt = y;
  if (hSpace !== undefined) out.hSpacePt = hSpace;
  if (vSpace !== undefined) out.vSpacePt = vSpace;
  const xAlign = pick<NonNullable<FrameProperties['xAlign']>>('xAlign', X_ALIGNS);
  const yAlign = pick<NonNullable<FrameProperties['yAlign']>>('yAlign', Y_ALIGNS);
  const hAnchor = pick<NonNullable<FrameProperties['hAnchor']>>('hAnchor', ANCHORS);
  const vAnchor = pick<NonNullable<FrameProperties['vAnchor']>>('vAnchor', ANCHORS);
  const wrap = pick<NonNullable<FrameProperties['wrap']>>('wrap', WRAPS);
  if (xAlign) out.xAlign = xAlign;
  if (yAlign) out.yAlign = yAlign;
  if (hAnchor) out.hAnchor = hAnchor;
  if (vAnchor) out.vAnchor = vAnchor;
  if (wrap) out.wrap = wrap;
  // §17.3.1.11 — a frame is a BOX, and one of zero width and zero height is
  // not a frame at all. fdo70812.docx writes exactly that in its document
  // defaults (`w:w="0" w:h="0" w:hRule="exact"`), which every paragraph then
  // inherited: each was laid out as a frame with no room in it, and "Sample
  // pages document." came out one word per line.
  const sized = (w !== undefined && w > 0) || (h !== undefined && h > 0);
  const placed = x !== undefined || y !== undefined || xAlign !== undefined || yAlign !== undefined;
  if (!sized && !(placed && (x ?? 0) + (y ?? 0) !== 0)) return undefined;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Parse a `w:pPr` element (ECMA-376 Part 1 §17.3.1) into {@link ParagraphProperties}.
 * Reads the style reference, alignment, spacing, indentation, page-break-before,
 * bidi, outline level, numbering reference (`w:numPr`) and paragraph-mark run
 * properties (`w:rPr`); unrecognized or out-of-range values are skipped.
 *
 * @param pPr The `w:pPr` element in flat (fast-xml-parser) shape, or anything
 *   non-element (yielding an empty result).
 * @returns The extracted properties; an empty object when `pPr` is absent.
 */
export function parseParagraphProperties(pPr: unknown): ParagraphProperties {
  const el = asElement(pPr);
  if (!el) return {};

  const out: Mutable<ParagraphProperties> = {};

  if ('w:pStyle' in el) {
    const v = getVal(el['w:pStyle']);
    if (v) out.styleId = v;
  }

  if ('w:jc' in el) {
    const v = getVal(el['w:jc']);
    if (v && ALIGNMENTS.has(v as Alignment)) {
      out.alignment = v as Alignment;
    }
  }

  if ('w:spacing' in el) {
    const node = el['w:spacing'];
    const before = parseIntAttr(node, 'before');
    const after = parseIntAttr(node, 'after');
    const line = parseIntAttr(node, 'line');
    const lineRule = getAttr(node, 'lineRule');
    if (before !== undefined) out.spacingBefore = twipsToPt(before);
    if (after !== undefined) out.spacingAfter = twipsToPt(after);
    if (line !== undefined) out.spacingLine = twipsToPt(line);
    if (lineRule && LINE_RULES.has(lineRule as 'auto' | 'exact' | 'atLeast')) {
      out.spacingLineRule = lineRule as 'auto' | 'exact' | 'atLeast';
    }
  }

  if ('w:ind' in el) {
    const node = el['w:ind'];
    const left = parseIntAttr(node, 'left');
    const right = parseIntAttr(node, 'right');
    const firstLine = parseIntAttr(node, 'firstLine');
    const hanging = parseIntAttr(node, 'hanging');
    if (left !== undefined) out.indentLeft = twipsToPt(left);
    if (right !== undefined) out.indentRight = twipsToPt(right);
    if (firstLine !== undefined) out.indentFirstLine = twipsToPt(firstLine);
    else if (hanging !== undefined) out.indentFirstLine = twipsToPt(-hanging);
  }

  // §17.3.1.11 — the paragraph is a floating text frame, not part of the flow.
  if ('w:framePr' in el) {
    const frame = parseFramePr(el['w:framePr']);
    if (frame) out.frame = frame;
  }

  if ('w:tabs' in el) {
    const stops = parseTabs(el['w:tabs']);
    if (stops.length > 0) out.tabs = stops;
  }

  // §17.3.1.24 `w:pBdr` — rules around the paragraph, spelled exactly as a
  // cell's are. Read nowhere, Test_ThemeBorderColor.docx lost the two coloured
  // rules that are the whole of its page.
  if ('w:pBdr' in el) {
    const borders = parseParagraphBorders(el['w:pBdr']);
    if (borders) out.borders = borders;
  }

  // §17.3.1.31 — the paragraph's own background. A direct `@w:fill` hex is
  // honoured, the way a cell's `w:shd` already is; `auto` and pattern-only
  // shading leave it unfilled.
  if ('w:shd' in el) {
    const fill = getAttr(el['w:shd'], 'fill');
    if (fill && fill !== 'auto' && /^[0-9A-Fa-f]{6}$/u.test(fill)) {
      out.shading = { colorHex: fill.toUpperCase() };
    }
  }

  // §17.3.1.41 `w:textDirection` — the paragraph reads bottom-to-top (`btLr`)
  // or top-to-bottom (`tbRl`). fdo76979.docx runs its side tab that way, in a
  // frame in the header, and we set it flat across the page.
  if ('w:textDirection' in el) {
    const v = getVal(el['w:textDirection']);
    if (v === 'btLr' || v === 'tbRl') out.textDirection = v;
  }

  // §17.3.1.32 `w:snapToGrid` — whether this paragraph stands on the section's
  // document grid (§17.6.5). Word's own header and footer styles turn it off,
  // so a gridded document's running text would otherwise be spaced like its
  // body: cjklist30.docx says so in both.
  if ('w:snapToGrid' in el) {
    const v = parseToggle(el['w:snapToGrid']);
    if (v !== undefined) out.snapToGrid = v;
  }

  if ('w:contextualSpacing' in el) {
    const v = parseToggle(el['w:contextualSpacing']);
    if (v !== undefined) out.contextualSpacing = v;
  }

  if ('w:pageBreakBefore' in el) {
    const v = parseToggle(el['w:pageBreakBefore']);
    if (v !== undefined) out.pageBreakBefore = v;
  }

  // ECMA-376 §17.3.1.6 — w:bidi is a toggle setting the paragraph base
  // direction to RTL.
  if ('w:bidi' in el) {
    const v = parseToggle(el['w:bidi']);
    if (v !== undefined) out.bidi = v;
  }

  // ECMA-376 §17.3.1.20 — w:outlineLvl (0–8 = Heading 1–9; 9 = body text).
  if ('w:outlineLvl' in el) {
    const v = parseIntAttr(el['w:outlineLvl'], 'val');
    if (v !== undefined) out.outlineLevel = v;
  }

  if ('w:numPr' in el) {
    const numPr = asElement(el['w:numPr']);
    if (numPr) {
      const numIdVal = numPr['w:numId'];
      const ilvlVal = numPr['w:ilvl'];
      const numIdAttr =
        typeof numIdVal === 'object' && numIdVal !== null ? getAttr(numIdVal, 'val') : undefined;
      const ilvlAttr =
        typeof ilvlVal === 'object' && ilvlVal !== null ? getAttr(ilvlVal, 'val') : undefined;
      if (numIdAttr !== undefined) {
        const ilvlNum = ilvlAttr !== undefined ? Number(ilvlAttr) : 0;
        out.numbering = {
          numId: numIdAttr,
          ilvl: Number.isFinite(ilvlNum) ? ilvlNum : 0,
        };
      }
    }
  }

  if ('w:rPr' in el) {
    const rPr = parseRunProperties(el['w:rPr']);
    if (Object.keys(rPr).length > 0) out.runProperties = rPr;
  }

  return out;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

// §17.18.90 ST_TabJc. `bar` draws a rule and advances nothing, and `clear`
// removes an inherited stop; neither places text, so neither becomes a stop.
const TAB_ALIGNMENTS: ReadonlyMap<string, TabStop['alignment']> = new Map([
  ['left', 'left'],
  ['start', 'left'],
  ['center', 'center'],
  ['right', 'right'],
  ['end', 'right'],
  ['decimal', 'decimal'],
]);

const TAB_LEADERS: ReadonlySet<string> = new Set(['dot', 'hyphen', 'underscore', 'middleDot']);

const BORDER_STYLES = new Set<BorderStyle>([
  'none',
  'single',
  'double',
  'thick',
  'dotted',
  'dashed',
]);

/**
 * §17.18.2 ST_Border — the rule's pattern, of which the standard names some
 * hundred and eighty. The handful we draw pass through; `nil` and `none` are no
 * rule at all; everything else is a rule we cannot draw exactly, and a solid
 * one of the stated width and colour is far closer than none. Rejected
 * outright, SdtContent.docx lost the `thickThinSmallGap` rule under its header.
 *
 * @param val The `w:val` attribute.
 * @returns The style to draw, or `undefined` when there is nothing to draw.
 */
function borderStyleOf(val: string | undefined): BorderStyle | undefined {
  if (!val) return undefined;
  // `nil` and `none` are a rule that is explicitly ABSENT: recorded as such, so
  // it overrides the one a style would otherwise lend the edge.
  if (val === 'nil' || val === 'none') return 'none';
  return BORDER_STYLES.has(val as BorderStyle) ? (val as BorderStyle) : 'single';
}

/**
 * §17.3.1.24 `w:pBdr` — the rules around a paragraph. Spelled exactly as a
 * cell's `w:tcBorders` is, but reached through the flat parse shape this module
 * works in, and with the `w:space` a cell border does not have.
 *
 * @param node The `w:pBdr` element.
 * @returns The edges that name a rule, or `undefined` when none do.
 */
function parseParagraphBorders(node: unknown): CellBorders | undefined {
  const el = asElement(node);
  if (!el) return undefined;
  const edge = (...names: Array<string>): Border | undefined => {
    const found = names.map((n) => el[n]).find((v) => v !== undefined);
    const b = asElement(found);
    if (!b) return undefined;
    const style = borderStyleOf(getVal(b));
    if (!style) return undefined;
    const sz = parseIntAttr(b, 'sz');
    const space = parseIntAttr(b, 'space');
    const color = getAttr(b, 'color');
    return {
      style,
      ...(sz !== undefined ? { width: eighthPtToPt(sz) } : {}),
      // §17.3.1.24 — `w:space` is in POINTS, not twips or eighths.
      ...(space !== undefined ? { spacePt: pt(space) } : {}),
      ...(color && color !== 'auto' && /^[0-9A-Fa-f]{6}$/u.test(color)
        ? { colorHex: color.toUpperCase() }
        : {}),
    };
  };
  const out: Mutable<CellBorders> = {};
  const top = edge('w:top');
  const bottom = edge('w:bottom');
  const left = edge('w:left', 'w:start');
  const right = edge('w:right', 'w:end');
  if (top) out.top = top;
  if (bottom) out.bottom = bottom;
  if (left) out.left = left;
  if (right) out.right = right;
  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * §17.3.1.37 `w:tabs` — the paragraph's own stops, in ascending position.
 *
 * @param node The `w:tabs` element.
 * @returns The stops that place text, sorted; `[]` when it declares none.
 */
function parseTabs(node: unknown): Array<TabStop> {
  const out: Array<TabStop> = [];
  for (const raw of asArray(asElement(node)?.['w:tab'])) {
    const el = asElement(raw);
    if (!el) continue;
    const pos = parseIntAttr(el, 'pos');
    if (pos === undefined) continue;
    const alignment = TAB_ALIGNMENTS.get(getAttr(el, 'val') ?? 'left');
    if (!alignment) continue;
    const leader = getAttr(el, 'leader');
    out.push({
      positionPt: twipsToPt(pos),
      alignment,
      ...(leader !== undefined && TAB_LEADERS.has(leader)
        ? { leader: leader as NonNullable<TabStop['leader']> }
        : {}),
    });
  }
  return out.sort((a, b) => a.positionPt - b.positionPt);
}
