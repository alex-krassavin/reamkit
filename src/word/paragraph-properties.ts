// ECMA-376 Part 1 §17.3.1 — Paragraph Properties (pPr).

import type {
  Alignment,
  Border,
  BorderStyle,
  CellBorders,
  CellShading,
  ParagraphProperties,
  TabStop,
} from '@/core/document-model';
import { eighthPtToPt, pt, twipsToPt } from '@/core/ir';

import { parseRunProperties } from '@/word/run-properties';
import { asArray, asElement, getAttr, getVal, parseIntAttr, parseToggle } from '@/word/xml-helpers';

const ALIGNMENTS = new Set<Alignment>(['left', 'right', 'center', 'both', 'distribute']);
const LINE_RULES = new Set<'auto' | 'exact' | 'atLeast'>(['auto', 'exact', 'atLeast']);

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
