// Style cascade resolver.
//
// ECMA-376 Part 1 §17.7.2 — order of resolution (lowest priority first):
//   1. Document defaults (docDefaults.rPrDefault / pPrDefault)
//   2. Paragraph style chain (resolved via basedOn) — runProperties become
//      the implicit default for runs in that paragraph; paragraphProperties
//      apply to the paragraph itself.
//   3. Character (run) style chain (basedOn) — runProperties only.
//   4. Direct properties on the paragraph / run.
//
// Higher priority overrides lower; an `undefined` field at the higher tier
// inherits from the lower one.

import type { ParagraphProperties, RunProperties, StyleSheet } from '@/core/document-model';

import type {
  ResolvedParagraphProperties,
  ResolvedRunProperties,
} from '@/core/style-cascade/types';
import { DEFAULT_RESOLVED_PARAGRAPH, DEFAULT_RESOLVED_RUN } from '@/core/style-cascade/types';

interface StyleChainResult {
  readonly rPr: RunProperties;
  readonly pPr: ParagraphProperties;
}

// Null Object for StyleSheet — parsers fall back to it when a package has no
// styles part, and grid-based formats (xlsx) render with it outright.
export const EMPTY_STYLE_SHEET: StyleSheet = {
  defaultRunProperties: {},
  defaultParagraphProperties: {},
  styles: new Map(),
};

// The renderer resolves the same (run, paragraph) pair in several phases
// (font collection, tokenization, table measuring) — memoize by identity.
// Model objects are readonly and rebuilt per parse, and the outer key is the
// sheet itself, so entries can never leak across documents (oop-design §4.2).
const runCascadeCache = new WeakMap<
  StyleSheet,
  WeakMap<RunProperties, WeakMap<ParagraphProperties, ResolvedRunProperties>>
>();

export function resolveRunProperties(
  runDirect: RunProperties,
  paragraphDirect: ParagraphProperties,
  sheet: StyleSheet,
): ResolvedRunProperties {
  let bySheet = runCascadeCache.get(sheet);
  if (!bySheet) {
    bySheet = new WeakMap();
    runCascadeCache.set(sheet, bySheet);
  }
  let byRun = bySheet.get(runDirect);
  if (!byRun) {
    byRun = new WeakMap();
    bySheet.set(runDirect, byRun);
  }
  const hit = byRun.get(paragraphDirect);
  if (hit) return hit;
  const resolved = computeRunProperties(runDirect, paragraphDirect, sheet);
  byRun.set(paragraphDirect, resolved);
  return resolved;
}

function computeRunProperties(
  runDirect: RunProperties,
  paragraphDirect: ParagraphProperties,
  sheet: StyleSheet,
): ResolvedRunProperties {
  let acc = mergeRun(DEFAULT_RESOLVED_RUN, sheet.defaultRunProperties);

  if (paragraphDirect.styleId) {
    const chain = resolveStyleChain(paragraphDirect.styleId, sheet);
    acc = mergeRun(acc, chain.rPr);
  }
  if (runDirect.styleId) {
    const chain = resolveStyleChain(runDirect.styleId, sheet);
    acc = mergeRun(acc, chain.rPr);
  }
  return mergeRun(acc, runDirect);
}

export function resolveParagraphProperties(
  paragraphDirect: ParagraphProperties,
  sheet: StyleSheet,
): ResolvedParagraphProperties {
  let acc = mergePar(DEFAULT_RESOLVED_PARAGRAPH, sheet.defaultParagraphProperties);
  if (paragraphDirect.styleId) {
    const chain = resolveStyleChain(paragraphDirect.styleId, sheet);
    acc = mergePar(acc, chain.pPr);
  }
  return mergePar(acc, paragraphDirect);
}

function resolveStyleChain(styleId: string, sheet: StyleSheet): StyleChainResult {
  const chain: Array<{ rPr: RunProperties; pPr: ParagraphProperties }> = [];
  const visited = new Set<string>();
  let cursor: string | undefined = styleId;
  while (cursor && !visited.has(cursor)) {
    visited.add(cursor);
    const style = sheet.styles.get(cursor);
    if (!style) break;
    chain.unshift({ rPr: style.runProperties, pPr: style.paragraphProperties });
    cursor = style.basedOn;
  }
  let rPr: RunProperties = {};
  let pPr: ParagraphProperties = {};
  for (const link of chain) {
    rPr = mergeRunPartial(rPr, link.rPr);
    pPr = mergeParPartial(pPr, link.pPr);
  }
  return { rPr, pPr };
}

function mergeRun(base: ResolvedRunProperties, override: RunProperties): ResolvedRunProperties {
  const lang = override.lang ?? base.lang;
  return {
    bold: override.bold ?? base.bold,
    italic: override.italic ?? base.italic,
    underline: override.underline ?? base.underline,
    strike: override.strike ?? base.strike,
    fontSizePt: override.fontSizePt ?? base.fontSizePt,
    colorHex: override.colorHex ?? base.colorHex,
    fontFamily: override.fontFamily ?? base.fontFamily,
    verticalAlign: override.verticalAlign ?? base.verticalAlign,
    rtl: override.rtl ?? base.rtl,
    ...(lang !== undefined ? { lang } : {}),
  };
}

function mergePar(
  base: ResolvedParagraphProperties,
  override: ParagraphProperties,
): ResolvedParagraphProperties {
  // outlineLevel / styleId are optional (most paragraphs have neither); carry
  // the higher-priority one through, omitting the key entirely when undefined
  // (exactOptionalPropertyTypes).
  const outlineLevel = override.outlineLevel ?? base.outlineLevel;
  const styleId = override.styleId ?? base.styleId;
  return {
    alignment: override.alignment ?? base.alignment,
    spacingBefore: override.spacingBefore ?? base.spacingBefore,
    spacingAfter: override.spacingAfter ?? base.spacingAfter,
    spacingLine: override.spacingLine ?? base.spacingLine,
    spacingLineRule: override.spacingLineRule ?? base.spacingLineRule,
    indentLeft: override.indentLeft ?? base.indentLeft,
    indentRight: override.indentRight ?? base.indentRight,
    indentFirstLine: override.indentFirstLine ?? base.indentFirstLine,
    pageBreakBefore: override.pageBreakBefore ?? base.pageBreakBefore,
    bidi: override.bidi ?? base.bidi,
    ...(outlineLevel !== undefined ? { outlineLevel } : {}),
    ...(styleId !== undefined ? { styleId } : {}),
  };
}

function mergeRunPartial(base: RunProperties, override: RunProperties): RunProperties {
  return copyDefined(base, override);
}

function mergeParPartial(
  base: ParagraphProperties,
  override: ParagraphProperties,
): ParagraphProperties {
  return copyDefined(base, override);
}

function copyDefined<T extends object>(base: T, override: T): T {
  const out = { ...base };
  for (const key of Object.keys(override) as Array<keyof T>) {
    const v = override[key];
    if (v !== undefined) out[key] = v;
  }
  return out;
}
