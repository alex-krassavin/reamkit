// ECMA-376 Part 1 §17.3.2 — Run Properties (rPr).

import type { FontFamilyMap, RunProperties, UnderlineStyle, VerticalAlign } from '@/document-model';

import {
  asElement,
  getAttr,
  getVal,
  parseIntAttr,
  parseToggle,
} from '@/ooxml/wordproc/xml-helpers';

const UNDERLINE_STYLES = new Set<UnderlineStyle>([
  'none',
  'single',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashHeavy',
  'wave',
]);

const VERTICAL_ALIGNS = new Set<VerticalAlign>(['baseline', 'superscript', 'subscript']);

export function parseRunProperties(rPr: unknown): RunProperties {
  const el = asElement(rPr);
  if (!el) return {};

  const out: Mutable<RunProperties> = {};

  if ('w:rStyle' in el) {
    const v = getVal(el['w:rStyle']);
    if (v) out.styleId = v;
  }

  if ('w:b' in el) {
    const v = parseToggle(el['w:b']);
    if (v !== undefined) out.bold = v;
  }
  if ('w:i' in el) {
    const v = parseToggle(el['w:i']);
    if (v !== undefined) out.italic = v;
  }
  if ('w:strike' in el) {
    const v = parseToggle(el['w:strike']);
    if (v !== undefined) out.strike = v;
  }

  if ('w:u' in el) {
    const v = getVal(el['w:u']);
    if (v && UNDERLINE_STYLES.has(v as UnderlineStyle)) {
      out.underline = v as UnderlineStyle;
    }
  }

  if ('w:sz' in el) {
    const v = parseIntAttr(el['w:sz'], 'val');
    if (v !== undefined) out.fontSizeHalfPoints = v;
  }

  if ('w:color' in el) {
    const v = getVal(el['w:color']);
    if (v && /^[0-9A-Fa-f]{6}$/.test(v)) {
      out.colorHex = v.toUpperCase();
    }
  }

  if ('w:rFonts' in el) {
    const ff = parseFontFamily(el['w:rFonts']);
    if (ff) out.fontFamily = ff;
  }

  if ('w:vertAlign' in el) {
    const v = getVal(el['w:vertAlign']);
    if (v && VERTICAL_ALIGNS.has(v as VerticalAlign)) {
      out.verticalAlign = v as VerticalAlign;
    }
  }

  // ECMA-376 §17.3.2.30 — w:rtl is a toggle property.
  if ('w:rtl' in el) {
    const v = parseToggle(el['w:rtl']);
    if (v !== undefined) out.rtl = v;
  }

  // ECMA-376 §17.3.2.20 — w:lang @w:val (the Latin language, e.g. "en-US").
  // Surfaced for the tagged-PDF per-element /Lang (a paragraph whose language
  // differs from the document default is tagged so AT switches pronunciation).
  if ('w:lang' in el) {
    const v = getAttr(el['w:lang'], 'val');
    if (v) out.lang = v;
  }

  return out;
}

function parseFontFamily(node: unknown): FontFamilyMap | undefined {
  const ff: Mutable<FontFamilyMap> = {};
  const ascii = getAttr(node, 'ascii');
  const hAnsi = getAttr(node, 'hAnsi');
  const cs = getAttr(node, 'cs');
  const eastAsia = getAttr(node, 'eastAsia');
  if (ascii) ff.ascii = ascii;
  if (hAnsi) ff.hAnsi = hAnsi;
  if (cs) ff.cs = cs;
  if (eastAsia) ff.eastAsia = eastAsia;
  return Object.keys(ff).length > 0 ? ff : undefined;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };
