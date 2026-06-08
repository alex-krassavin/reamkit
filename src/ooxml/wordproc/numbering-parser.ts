// ECMA-376 Part 1 §17.9 — Numbering.xml parser.
//
// Numbering is a two-level indirection:
//   abstractNum (defines levels: numFmt, lvlText, start, indent, …)
//   num         (instance binding a numId → abstractNumId, with optional
//                per-level overrides — overrides not yet implemented).
// A paragraph references a list via <w:numPr> { numId, ilvl } in its pPr.

import { XMLParser } from 'fast-xml-parser';

import type {
  AbstractNumbering,
  Numbering,
  NumberingFormat,
  NumberingInstance,
  NumberingLevel,
} from '@/document-model';

import { parseParagraphProperties } from '@/ooxml/wordproc/paragraph-properties';
import { parseRunProperties } from '@/ooxml/wordproc/run-properties';
import { asArray, asElement, getAttr } from '@/ooxml/wordproc/xml-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

const FORMATS = new Set<NumberingFormat>([
  'decimal',
  'lowerLetter',
  'upperLetter',
  'lowerRoman',
  'upperRoman',
  'bullet',
  'none',
]);

export const EMPTY_NUMBERING: Numbering = {
  abstractNums: new Map(),
  numInstances: new Map(),
};

export function parseNumbering(data: Uint8Array): Numbering {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const root = asElement(tree['w:numbering']);
  if (!root) return EMPTY_NUMBERING;

  const abstractNums = new Map<string, AbstractNumbering>();
  for (const a of asArray(root['w:abstractNum'])) {
    const el = asElement(a);
    if (!el) continue;
    const id = getAttr(el, 'abstractNumId');
    if (!id) continue;
    const levels = new Map<number, NumberingLevel>();
    for (const lvlNode of asArray(el['w:lvl'])) {
      const lvlEl = asElement(lvlNode);
      if (!lvlEl) continue;
      const ilvlStr = getAttr(lvlEl, 'ilvl');
      if (!ilvlStr) continue;
      const ilvl = Number(ilvlStr);
      if (!Number.isFinite(ilvl)) continue;

      const startAttr = getValVal(lvlEl['w:start']);
      const start = startAttr !== undefined ? Number(startAttr) : 1;
      const fmtStr = getValVal(lvlEl['w:numFmt']) ?? 'decimal';
      const format: NumberingFormat = FORMATS.has(fmtStr as NumberingFormat)
        ? (fmtStr as NumberingFormat)
        : 'decimal';
      const lvlText = getValVal(lvlEl['w:lvlText']) ?? '';

      levels.set(ilvl, {
        ilvl,
        start: Number.isFinite(start) ? start : 1,
        format,
        lvlText,
        paragraphProperties: parseParagraphProperties(lvlEl['w:pPr']),
        runProperties: parseRunProperties(lvlEl['w:rPr']),
      });
    }
    abstractNums.set(id, { id, levels });
  }

  const numInstances = new Map<string, NumberingInstance>();
  for (const n of asArray(root['w:num'])) {
    const el = asElement(n);
    if (!el) continue;
    const numId = getAttr(el, 'numId');
    if (!numId) continue;
    const abstractNumId = getValVal(el['w:abstractNumId']);
    if (!abstractNumId) continue;
    numInstances.set(numId, { numId, abstractNumId });
  }

  return { abstractNums, numInstances };
}

function getValVal(node: unknown): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const v =
    (node as Record<string, unknown>)['@_w:val'] ?? (node as Record<string, unknown>)['@_val'];
  return typeof v === 'string' ? v : undefined;
}
