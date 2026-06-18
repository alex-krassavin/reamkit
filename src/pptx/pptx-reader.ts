// E-PPTX PX0–PX5 — PresentationML (.pptx) reader: bytes → FlowDoc. A presentation
// is a positioned canvas, which maps cleanly onto the existing IR: each slide is a
// section at the deck's page size, and its shapes become absolutely positioned
// floating elements (Route A, epics.md). PX0 established the seam — sniff, slide
// size from p:sldSz, slide count from p:sldIdLst, one page per slide; the slide's
// shapes (slide-parser) fill each page. This module owns the part graph: it
// resolves each slide's layout → master → theme chain into the placeholder
// cascade (PX2), the deck colour resolver (PX5a) and the inherited background
// (PX5b), all memoized by layout path, plus the per-slide image (PX3a) and chart
// (PX4a) resolvers.

import { XMLParser } from 'fast-xml-parser';

import type { BodyElement, Chart, SectionProperties, ShapeFill } from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss, Pt, ResourceId } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import type { Relationship } from '@/core/opc';
import type { PlaceholderCascade } from '@/pptx/placeholder-cascade';
import type { SlideContext } from '@/pptx/slide-parser';

import type { ColorResolver } from '@/core/drawingml/colors';

import { bytesInclude } from '@/core/bytes';
import { parseChart, withChartColorStyle } from '@/core/drawingml/chart-parser';
import {
  DEFAULT_THEME_PALETTE,
  defaultColorResolver,
  makeColorResolver,
} from '@/core/drawingml/colors';
import { parseTheme } from '@/core/drawingml/theme-parser';
import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { OpcPackage } from '@/core/opc';
import { poAttr, poChildren, poFindDescendant, poIntAttr, poIs } from '@/core/po-helpers';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { buildPlaceholderCascade } from '@/pptx/placeholder-cascade';
import { backdropElement, parseBackgroundFill, parseSlideShapes } from '@/pptx/slide-parser';

const EMU_PER_PT = 12700;
// §19.2.1.39 sldSz default — a 4:3 deck (10" × 7.5"); real decks always declare it.
const DEFAULT_CX = 9144000;
const DEFAULT_CY = 6858000;

const decoder = new TextDecoder();
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

// Whether a slide is hidden (p:sld@show="0"). The attribute lives on the slide's
// root element; scan a bounded prefix (the root tag is at the very top) rather
// than parse the whole slide a second time just for this. If the root tag were
// improbably long and fell outside the window we simply render the slide — a
// harmless over-render, never a wrong omission.
function isSlideHidden(data: Uint8Array): boolean {
  const head = decoder.decode(data.subarray(0, 4096));
  const root = head.match(/<p:sld\b[^>]*>/);
  return root ? /\bshow\s*=\s*["']0["']/.test(root[0]) : false;
}

export function readPptx(bytes: Uint8Array): ReadResult<FlowDoc> {
  const losses: Array<Loss> = [];
  const pkg = OpcPackage.open(bytes);
  const presPath = pkg.getMainDocumentPath();
  const presData = pkg.getPart(presPath);
  const resources = new ResourceStore();
  const charts = new Map<string, Chart>();

  let cx = DEFAULT_CX;
  let cy = DEFAULT_CY;
  const slideParts: Array<{ path: string; data: Uint8Array }> = [];
  if (presData) {
    const tree = parser.parse(decoder.decode(presData)) as Array<PoNode>;
    const pres = tree.find((n) => poIs(n, 'p:presentation'));
    const kids = pres ? poChildren(pres) : [];
    const sldSz = kids.find((c) => poIs(c, 'p:sldSz'));
    cx = (sldSz ? poIntAttr(sldSz, 'cx') : undefined) || DEFAULT_CX;
    cy = (sldSz ? poIntAttr(sldSz, 'cy') : undefined) || DEFAULT_CY;
    // Slide order is p:sldIdLst/p:sldId@r:id, each resolving to a /slide rel of
    // the presentation part. Resolve them in declared order to the slide parts.
    const slideRelById = new Map<string, Relationship>(
      pkg
        .getPartRelationships(presPath)
        .filter((r) => r.type.endsWith('/slide'))
        .map((r) => [r.id, r] as const),
    );
    const lst = kids.find((c) => poIs(c, 'p:sldIdLst'));
    const ids = lst ? poChildren(lst).filter((c) => poIs(c, 'p:sldId')) : [];
    let hidden = 0;
    for (const sldId of ids) {
      const rid = poAttr(sldId, 'r:id');
      const rel = rid !== undefined ? slideRelById.get(rid) : undefined;
      const part = rel ? pkg.resolveRelatedPart(presPath, rel) : undefined;
      if (!part) continue;
      // p:sld@show="0" marks a hidden slide; PowerPoint and LibreOffice both omit
      // hidden slides from a printed/exported deck, so we skip them too.
      if (isSlideHidden(part.data)) {
        hidden++;
        continue;
      }
      slideParts.push(part);
    }
    if (hidden > 0) {
      losses.push({
        severity: 'dropped',
        feature: FEATURES.text,
        detail: `${hidden} hidden slide(s) omitted (p:sld@show="0")`,
      });
    }
  }
  if (slideParts.length === 0) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail: 'presentation has no slides',
    });
  }

  // One page per slide at the deck size. Each slide is anchored by an in-flow
  // paragraph that forces a page (pageBreakBefore after the first; a single
  // zero-width-space run gives the otherwise-empty page one line so it is
  // actually emitted). The slide's text-bearing shapes are floating elements
  // placed on that page at their EMU positions.
  const slideCount = Math.max(1, slideParts.length);
  const pageW = pt(cx / EMU_PER_PT);
  const pageH = pt(cy / EMU_PER_PT);
  const stylesByLayout = new Map<string, SlideStyles>();
  const body: Array<BodyElement> = [];
  for (let i = 0; i < slideCount; i++) {
    body.push({
      kind: 'paragraph',
      paragraph: {
        properties: i > 0 ? { pageBreakBefore: true } : {},
        runs: [{ text: '​', properties: {} }],
      },
    });
    const part = slideParts[i];
    if (part) {
      const styles = slideStylesFor(pkg, part.path, stylesByLayout);
      const ctx: SlideContext = {
        ...(styles.cascade ? { cascade: styles.cascade } : {}),
        colors: styles.colors,
        resolveImage: makeSlideImageResolver(pkg, part.path, resources),
        resolveChart: makeSlideChartResolver(pkg, part.path, charts, styles.colors),
        resolveHyperlink: makeHyperlinkResolver(pkg, part.path),
        resolveDiagram: makeSlideDiagramResolver(pkg, part.path),
        onLoss: (loss) => losses.push(loss.where ? loss : { ...loss, where: `slide ${i + 1}` }),
      };
      body.push(...parseSlide(part.data, ctx, styles.background, pageW, pageH));
    }
  }

  const section: SectionProperties = {
    pageSize: { width: pageW, height: pageH },
    // A slide is a margin-less canvas: floating shapes position from the page
    // edge (relativeFrom:'page'), and in-flow content (a table — PX4b) sits at
    // the top-left rather than inside a default print margin.
    margins: { top: pt(0), right: pt(0), bottom: pt(0), left: pt(0) },
    headers: [],
    footers: [],
  };
  const doc: FlowDoc = {
    kind: 'flow',
    body: resolveBodyStyles([...body], EMPTY_STYLE_SHEET),
    sections: [],
    section,
    styles: EMPTY_STYLE_SHEET,
    resources,
    ...(charts.size > 0 ? { charts } : {}),
  };
  return { doc, losses };
}

export function parseXml(data: Uint8Array): Array<PoNode> {
  return parser.parse(decoder.decode(data)) as Array<PoNode>;
}

// A slide part's bytes → a full-slide backdrop (PX5b) followed by its floating
// shapes/images/frames. The background is the slide's own p:bg, else the
// inherited layout/master one.
function parseSlide(
  data: Uint8Array,
  ctx: SlideContext,
  inheritedBg: ShapeFill | undefined,
  pageW: Pt,
  pageH: Pt,
): Array<BodyElement> {
  const tree = parseXml(data);
  const sld = tree.find((n) => poIs(n, 'p:sld'));
  const cSld = sld ? poChildren(sld).find((c) => poIs(c, 'p:cSld')) : undefined;
  const colors = ctx.colors ?? defaultColorResolver;
  const bgNode = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:bg')) : undefined;
  const bg = (bgNode ? parseBackgroundFill(bgNode, colors) : undefined) ?? inheritedBg;
  const spTree = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:spTree')) : undefined;

  const out: Array<BodyElement> = [];
  if (bg) out.push(backdropElement(bg, pageW, pageH));
  if (spTree) out.push(...parseSlideShapes(spTree, ctx));
  return out;
}

// A layout's or master's p:cSld/p:bg → its background fill (PX5b).
function partBackground(
  tree: ReadonlyArray<PoNode>,
  root: 'p:sldLayout' | 'p:sldMaster',
  colors: ColorResolver,
): ShapeFill | undefined {
  const sld = tree.find((n) => poIs(n, root));
  const cSld = sld ? poChildren(sld).find((c) => poIs(c, 'p:cSld')) : undefined;
  const bg = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:bg')) : undefined;
  return bg ? parseBackgroundFill(bg, colors) : undefined;
}

// An image resolver scoped to one slide: a blip relationship id (a:blip
// @r:embed) → the media bytes, stored (content-addressed, deduped) in the
// document's ResourceStore. Relationship ids are scoped to their owning part,
// so this resolves against the slide's own .rels (mirrors docx).
function makeSlideImageResolver(
  pkg: OpcPackage,
  slidePath: string,
  resources: ResourceStore,
): (relId: string) => ResourceId | undefined {
  const cache = new Map<string, ResourceId | undefined>();
  return (relId) => {
    if (cache.has(relId)) return cache.get(relId);
    const rel = pkg.getPartRelationships(slidePath).find((r) => r.id === relId);
    const resolved = rel ? pkg.resolveRelatedPart(slidePath, rel) : undefined;
    const id = resolved ? resources.put(resolved.data) : undefined;
    cache.set(relId, id);
    return id;
  };
}

// A hyperlink resolver scoped to one slide: a run's a:hlinkClick @r:id → the
// external target URL from the slide's .rels (PX6). Internal links (to another
// slide) have no URL and resolve to undefined.
function makeHyperlinkResolver(
  pkg: OpcPackage,
  slidePath: string,
): (relId: string) => string | undefined {
  return (relId) => {
    const rel = pkg.getPartRelationships(slidePath).find((r) => r.id === relId);
    return rel && rel.targetMode === 'External' ? rel.target : undefined;
  };
}

// A chart resolver scoped to one slide: a c:chart relationship id → the parsed
// chart, stored in the document's charts map under a globally-unique key
// (relationship ids are part-scoped, so two slides can reuse the same id). The
// ChartBlock carries that key as its chartRelId. Chart colours resolve through
// the deck's theme palette (PX5).
function makeSlideChartResolver(
  pkg: OpcPackage,
  slidePath: string,
  charts: Map<string, Chart>,
  colors: ColorResolver,
): (relId: string) => string | undefined {
  const cache = new Map<string, string | undefined>();
  return (relId) => {
    if (cache.has(relId)) return cache.get(relId);
    const rel = pkg.getPartRelationships(slidePath).find((r) => r.id === relId);
    const resolved = rel ? pkg.resolveRelatedPart(slidePath, rel) : undefined;
    const chart = resolved ? parseChart(resolved.data, colors) : null;
    let key: string | undefined;
    if (chart && resolved) {
      key = `${slidePath}!${relId}`;
      charts.set(key, withChartColorStyle(chart, pkg, resolved.path, colors));
    }
    cache.set(relId, key);
    return key;
  };
}

// SmartArt: a data relationship id (dgm:relIds @r:dm) → the diagram's
// pre-rendered drawing override (its dsp:spTree). Follows slide →
// diagrams/data#.xml → (rel type .../diagramDrawing) → diagrams/drawing#.xml.
// Undefined when the file ships no drawing override (E-SMARTART SA0).
function makeSlideDiagramResolver(
  pkg: OpcPackage,
  slidePath: string,
): (relId: string) => PoNode | undefined {
  const cache = new Map<string, PoNode | undefined>();
  return (relId) => {
    if (cache.has(relId)) return cache.get(relId);
    let spTree: PoNode | undefined;
    const dataRel = pkg.getPartRelationships(slidePath).find((r) => r.id === relId);
    const data = dataRel ? pkg.resolveRelatedPart(slidePath, dataRel) : undefined;
    if (data) {
      const drawRel = pkg
        .getPartRelationships(data.path)
        .find((r) => r.type.endsWith('/diagramDrawing'));
      const draw = drawRel ? pkg.resolveRelatedPart(data.path, drawRel) : undefined;
      if (draw) {
        for (const root of parseXml(draw.data)) {
          const found = poFindDescendant(root, 'dsp:spTree');
          if (found) {
            spTree = found;
            break;
          }
        }
      }
    }
    cache.set(relId, spTree);
    return spTree;
  };
}

// The placeholder cascade + colour resolver for a slide, derived from its
// slideLayout → slideMaster (→ theme) chain and memoized by layout path (slides
// that share a layout share both).
interface SlideStyles {
  readonly cascade?: PlaceholderCascade;
  readonly colors: ColorResolver;
  // The inherited background fill (layout, else master) for slides that have no
  // p:bg of their own (PX5b).
  readonly background?: ShapeFill;
}

function slideStylesFor(
  pkg: OpcPackage,
  slidePath: string,
  cache: Map<string, SlideStyles>,
): SlideStyles {
  const layoutRel = pkg
    .getPartRelationships(slidePath)
    .find((r) => r.type.endsWith('/slideLayout'));
  const layout = layoutRel ? pkg.resolveRelatedPart(slidePath, layoutRel) : undefined;
  if (!layout) return { colors: defaultColorResolver };
  const cached = cache.get(layout.path);
  if (cached) return cached;

  const masterRel = pkg
    .getPartRelationships(layout.path)
    .find((r) => r.type.endsWith('/slideMaster'));
  const master = masterRel ? pkg.resolveRelatedPart(layout.path, masterRel) : undefined;
  const colors = master ? deckColorResolver(pkg, master.path) : defaultColorResolver;
  const layoutTree = parseXml(layout.data);
  const masterTree = master ? parseXml(master.data) : undefined;
  const cascade = buildPlaceholderCascade(layoutTree, masterTree, colors);
  const background =
    partBackground(layoutTree, 'p:sldLayout', colors) ??
    (masterTree ? partBackground(masterTree, 'p:sldMaster', colors) : undefined);
  const styles: SlideStyles = { cascade, colors, ...(background ? { background } : {}) };
  cache.set(layout.path, styles);
  return styles;
}

// The slide master's theme part (a:clrScheme) → a ColorResolver: the built-in
// Office palette with the deck's scheme colours merged over it (mirrors the
// docx/xlsx readers). The Office palette stands in when there is no theme.
function deckColorResolver(pkg: OpcPackage, masterPath: string): ColorResolver {
  const themeRel = pkg.getPartRelationships(masterPath).find((r) => r.type.endsWith('/theme'));
  const theme = themeRel ? pkg.resolveRelatedPart(masterPath, themeRel) : undefined;
  if (!theme) return defaultColorResolver;
  const palette = new Map(DEFAULT_THEME_PALETTE);
  for (const [slot, hex] of parseTheme(theme.data)) palette.set(slot, hex);
  return makeColorResolver(palette);
}

export const pptxReader: DocumentReader<FlowDoc> = {
  id: 'pptx',
  produces: 'flow',
  supports: new Set([
    FEATURES.text,
    FEATURES.images,
    FEATURES.shapes,
    FEATURES.charts,
    FEATURES.tables,
  ]),
  // A pptx is a ZIP whose parts include ppt/presentation.xml — a cheap substring
  // probe of the container bytes, no unzip needed (mirrors the docx/xlsx sniff).
  sniff: (bytes) =>
    bytes[0] === 0x50 && bytes[1] === 0x4b && bytesInclude(bytes, 'ppt/presentation.xml'),
  read: (bytes) => readPptx(bytes),
};
