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
} from '@/core/document-model';

import { parseParagraphProperties } from '@/word/paragraph-properties';
import { parseRunProperties } from '@/word/run-properties';
import { asArray, asElement, getAttr } from '@/word/xml-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

const FORMATS = new Set<NumberingFormat>([
  'decimal',
  'decimalZero',
  'decimalFullWidth',
  'ordinal',
  'lowerLetter',
  'upperLetter',
  'lowerRoman',
  'upperRoman',
  'ideographTraditional',
  'ideographZodiac',
  'ideographLegalTraditional',
  'ideographDigital',
  'koreanDigital2',
  'chineseCounting',
  'chineseCountingThousand',
  'japaneseCounting',
  'koreanCounting',
  'taiwaneseCounting',
  'taiwaneseCountingThousand',
  'hebrew1',
  'hebrew2',
  'decimalEnclosedCircle',
  'bullet',
  'none',
]);

/** The empty {@link Numbering} returned when a document has no `numbering.xml`. */
export const EMPTY_NUMBERING: Numbering = {
  abstractNums: new Map(),
  numInstances: new Map(),
};

/**
 * Parse `word/numbering.xml` (ECMA-376 Part 1 §17.9) into a {@link Numbering}:
 * the abstract numbering definitions (each level's format, start, text template
 * and indent/run properties) plus the num instances that bind a `numId` to an
 * `abstractNumId`. `basedOn` inheritance and per-level overrides are not resolved
 * here; an unparseable format defaults to `decimal`.
 *
 * @param data The raw `numbering.xml` bytes.
 * @returns The parsed numbering, or {@link EMPTY_NUMBERING} when the root is absent.
 */
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

      // §17.9.25 — a level that states no `w:start` starts at ZERO, not one.
      // FDO74105.docx omits it and LibreOffice numbers its list from 0.
      const startAttr = getValVal(lvlEl['w:start']);
      const start = startAttr !== undefined ? Number(startAttr) : 0;
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
