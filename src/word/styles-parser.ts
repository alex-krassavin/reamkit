// ECMA-376 Part 1 §17.7 — Styles.xml parser.
//
// Produces a StyleSheet of document defaults + named styles. Inheritance
// (basedOn) is recorded but not resolved here — that happens in the
// style-cascade module so callers can decide which paragraph/run a style
// applies to.

import { XMLParser } from 'fast-xml-parser';

import type { Style, StyleSheet, StyleType } from '@/core/document-model';

import { parseParagraphProperties } from '@/word/paragraph-properties';
import { parseRunProperties } from '@/word/run-properties';
import { asArray, asElement, getAttr, getVal } from '@/word/xml-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

const STYLE_TYPES = new Set<StyleType>(['paragraph', 'character', 'table', 'numbering']);

export const EMPTY_STYLE_SHEET: StyleSheet = {
  defaultRunProperties: {},
  defaultParagraphProperties: {},
  styles: new Map(),
};

export function parseStyles(stylesXml: Uint8Array): StyleSheet {
  const xml = decoder.decode(stylesXml);
  const tree = parser.parse(xml) as Record<string, unknown>;

  const root = asElement(tree['w:styles']);
  if (!root) return EMPTY_STYLE_SHEET;

  const docDefaults = asElement(root['w:docDefaults']);
  let defaultRunProperties = {};
  let defaultParagraphProperties = {};
  if (docDefaults) {
    const rPrDef = asElement(docDefaults['w:rPrDefault']);
    if (rPrDef) defaultRunProperties = parseRunProperties(rPrDef['w:rPr']);
    const pPrDef = asElement(docDefaults['w:pPrDefault']);
    if (pPrDef) defaultParagraphProperties = parseParagraphProperties(pPrDef['w:pPr']);
  }

  const styles = new Map<string, Style>();
  for (const s of asArray(root['w:style'])) {
    const el = asElement(s);
    if (!el) continue;
    const id = getAttr(el, 'styleId');
    const typeAttr = getAttr(el, 'type');
    if (!id || !typeAttr) continue;
    if (!STYLE_TYPES.has(typeAttr as StyleType)) continue;
    const type = typeAttr as StyleType;

    const basedOnVal = getVal(el['w:basedOn']);
    const isDefault = getAttr(el, 'default') === '1';

    const style: Style = {
      id,
      type,
      ...(basedOnVal ? { basedOn: basedOnVal } : {}),
      isDefault,
      runProperties: parseRunProperties(el['w:rPr']),
      paragraphProperties: parseParagraphProperties(el['w:pPr']),
    };
    styles.set(id, style);
  }

  return { defaultRunProperties, defaultParagraphProperties, styles };
}
