// ECMA-376 Part 1 §17.3.1 — Paragraph Properties (pPr).

import type { Alignment, ParagraphProperties } from '@/document-model';

import { parseRunProperties } from '@/ooxml/wordproc/run-properties';
import {
  asElement,
  getAttr,
  getVal,
  parseIntAttr,
  parseToggle,
} from '@/ooxml/wordproc/xml-helpers';

const ALIGNMENTS = new Set<Alignment>(['left', 'right', 'center', 'both', 'distribute']);
const LINE_RULES = new Set<'auto' | 'exact' | 'atLeast'>(['auto', 'exact', 'atLeast']);

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
    if (before !== undefined) out.spacingBeforeTwips = before;
    if (after !== undefined) out.spacingAfterTwips = after;
    if (line !== undefined) out.spacingLineTwips = line;
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
    if (left !== undefined) out.indentLeftTwips = left;
    if (right !== undefined) out.indentRightTwips = right;
    if (firstLine !== undefined) out.indentFirstLineTwips = firstLine;
    else if (hanging !== undefined) out.indentFirstLineTwips = -hanging;
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
