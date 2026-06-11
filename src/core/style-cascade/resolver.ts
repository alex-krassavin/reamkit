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

import type {
  BodyElement,
  Paragraph,
  ParagraphProperties,
  RunProperties,
  StyleSheet,
} from '@/core/document-model';

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

// Paragraph-side memo, the symmetric half of runCascadeCache: keyed by the
// direct-properties object identity. Producers that share properties objects
// across paragraphs (the xlsx grid mapper) collapse to one resolved object
// per distinct input instead of one per paragraph.
const paragraphCascadeCache = new WeakMap<
  StyleSheet,
  WeakMap<ParagraphProperties, ResolvedParagraphProperties>
>();

export function resolveParagraphProperties(
  paragraphDirect: ParagraphProperties,
  sheet: StyleSheet,
): ResolvedParagraphProperties {
  let bySheet = paragraphCascadeCache.get(sheet);
  if (!bySheet) {
    bySheet = new WeakMap();
    paragraphCascadeCache.set(sheet, bySheet);
  }
  const hit = bySheet.get(paragraphDirect);
  if (hit) return hit;
  const resolved = computeParagraphProperties(paragraphDirect, sheet);
  bySheet.set(paragraphDirect, resolved);
  return resolved;
}

function computeParagraphProperties(
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
  const numbering = override.numbering ?? base.numbering;
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
    ...(numbering !== undefined ? { numbering } : {}),
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

// ---------------------------------------------------------------------------
// FlowDoc transform (ir-design stage 6, variant A): resolve the cascade for an
// entire body so the tree carries final effective properties. Readers run this
// after numbering; writers and the layout see resolved values. Resolving again
// over EMPTY_STYLE_SHEET is the identity (defaults are fully overwritten and
// style chains are empty), which keeps direct raw-body callers working.
//
// The transform mutates the tree IN PLACE and returns it. Readers own the
// freshly-parsed trees they pass in (nothing else holds a reference), and
// rebuilding the tree immutably doubles peak memory — a real 874 KB workbook
// with a huge grid (POI bug62181.xlsx) OOMed a 512 MB heap on the rebuild
// version. Only the `properties` fields are overwritten; the resolved objects
// themselves stay memo-shared via the resolveRunProperties WeakMap cache.
// ---------------------------------------------------------------------------

// Pre-seed the cascade cache with a resolved pair's fixpoint. Resolving an
// already-resolved (run, paragraph) pair over the empty sheet is the identity
// by value (the stage-6 contract) — registering it here makes the renderer's
// idempotent re-resolve a cache HIT returning the very same object, instead
// of allocating an equal copy per unique pair (on grid-shaped documents those
// copies double the resolved-property population).
function primeResolvedFixpoint(run: RunProperties, para: ParagraphProperties): void {
  let bySheet = runCascadeCache.get(EMPTY_STYLE_SHEET);
  if (!bySheet) {
    bySheet = new WeakMap();
    runCascadeCache.set(EMPTY_STYLE_SHEET, bySheet);
  }
  let byRun = bySheet.get(run);
  if (!byRun) {
    byRun = new WeakMap();
    bySheet.set(run, byRun);
  }
  if (!byRun.has(para)) byRun.set(para, run as ResolvedRunProperties);
}

// The paragraph-side fixpoint: same idea as primeResolvedFixpoint, for the
// renderer's re-resolve of an already-resolved paragraph over the empty sheet.
function primeParagraphFixpoint(para: ParagraphProperties): void {
  let bySheet = paragraphCascadeCache.get(EMPTY_STYLE_SHEET);
  if (!bySheet) {
    bySheet = new WeakMap();
    paragraphCascadeCache.set(EMPTY_STYLE_SHEET, bySheet);
  }
  if (!bySheet.has(para)) bySheet.set(para, para as ResolvedParagraphProperties);
}

export function resolveBodyStyles(
  body: ReadonlyArray<BodyElement>,
  sheet: StyleSheet,
): ReadonlyArray<BodyElement> {
  const visitParagraph = (p: Paragraph): void => {
    // Run resolution sees the RAW paragraph properties (its styleId drives
    // the paragraph-style rPr layer) — so resolve every run first, then
    // overwrite the paragraph's own properties. Same order the renderer used.
    for (const r of p.runs) {
      (r as { properties: RunProperties }).properties = resolveRunProperties(
        r.properties,
        p.properties,
        sheet,
      );
    }
    (p as { properties: ParagraphProperties }).properties = resolveParagraphProperties(
      p.properties,
      sheet,
    );
    primeParagraphFixpoint(p.properties);
    for (const r of p.runs) primeResolvedFixpoint(r.properties, p.properties);
  };

  const visit = (el: BodyElement): void => {
    if (el.kind === 'paragraph') {
      visitParagraph(el.paragraph);
    } else if (el.kind === 'table') {
      for (const row of el.table.rows) {
        for (const cell of row.cells) {
          for (const child of cell.content) visit(child);
        }
      }
    } else if (el.kind === 'shape' && el.shape.text) {
      for (const child of el.shape.text.content) visit(child);
    }
    // image, chart, textless shape: nothing to resolve
  };

  for (const el of body) visit(el);
  return body;
}

export function resolveHeadersFootersStyles(
  hf: ReadonlyMap<string, ReadonlyArray<BodyElement>>,
  sheet: StyleSheet,
): ReadonlyMap<string, ReadonlyArray<BodyElement>> {
  for (const value of hf.values()) resolveBodyStyles(value, sheet);
  return hf;
}
