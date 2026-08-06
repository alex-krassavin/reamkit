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
import type { ThemeFonts } from '@/core/drawingml/theme-parser';
import type { PlaceholderCascade } from '@/pptx/placeholder-cascade';
import type { SlideContext, ThemeFillStyles } from '@/pptx/slide-parser';

import type { ColorResolver, SchemeAliases } from '@/core/drawingml/colors';

import type { FontRegistry } from '@/core/font';
import { packageHasPart } from '@/core/bytes';
import { DiagramData } from '@/core/drawingml/diagram/data-model';
import { diagramDrawingXml } from '@/core/drawingml/diagram/to-drawing';
import { layoutDiagram } from '@/core/drawingml/diagram/layout-engine';
import { parseChart, withChartColorStyle } from '@/core/drawingml/chart-parser';
import {
  DEFAULT_SCHEME_ALIAS,
  DEFAULT_THEME_PALETTE,
  defaultColorResolver,
  makeColorResolver,
} from '@/core/drawingml/colors';
import {
  parseTheme,
  parseThemeBgFillStyles,
  parseThemeFillStyles,
  parseThemeFonts,
  parseThemeLineWidths,
} from '@/core/drawingml/theme-parser';
import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { OpcPackage } from '@/core/opc';
import { loadPptxEmbeddedFonts } from '@/pptx/embedded-fonts';
import { presetTableStyle } from '@/pptx/preset-table-styles';
import { resolveAlternateContent } from '@/core/opc/alternate-content';
import { poAttr, poChildren, poFindDescendant, poIntAttr, poIs } from '@/core/po-helpers';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { buildPlaceholderCascade, parseLevelStyles } from '@/pptx/placeholder-cascade';
import { normalizeSpid, parseVmlImageRels } from '@/pptx/ole-preview';
import {
  asBackdrop,
  backdropElement,
  parseBackgroundFill,
  parseSlideShapes,
} from '@/pptx/slide-parser';

const EMU_PER_PT = 12700;
// §19.2.1.39 sldSz default — a 4:3 deck (10" × 7.5"); real decks always declare it.
const DEFAULT_CX = 9144000;
const DEFAULT_CY = 6858000;

const decoder = new TextDecoder();
const encoder = new TextEncoder();
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
  // The parser's runaway-nesting guard defaults to 100 tags, which a slide of
  // nested groups reaches (see word/document-parser).
  maxNestedTags: 1000,
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

/**
 * Read a PresentationML (`.pptx`) package into a {@link FlowDoc}: one page per
 * slide at the deck size (`p:sldSz`), each slide's shapes placed as absolutely
 * positioned floating elements. Hidden slides (`p:sld@show="0"`) are omitted (and
 * recorded as a {@link Loss}, mirroring PowerPoint/LibreOffice export). Resolves
 * each slide's layout → master → theme chain into the placeholder cascade, colour
 * resolver and inherited background, plus its images and charts.
 *
 * @param bytes The raw `.pptx` (OPC ZIP) bytes.
 * @returns The {@link FlowDoc} plus the accumulated loss report.
 */
export function readPptx(bytes: Uint8Array): ReadResult<FlowDoc> {
  const losses: Array<Loss> = [];
  const pkg = OpcPackage.open(bytes);
  const presPath = pkg.getMainDocumentPath();
  const tableStyles = makeTableStyleResolver(pkg, presPath);
  const presData = pkg.getPart(presPath);
  const resources = new ResourceStore();
  const charts = new Map<string, Chart>();

  let cx = DEFAULT_CX;
  let cy = DEFAULT_CY;
  // §19.2.1.8 — the deck's own default text style, which is what a plain text
  // box on a slide is written in when it says nothing itself.
  let defaultTextStyle: PoNode | undefined;
  // §19.2.1.13 — the faces the deck brings with it, so it reads the same where
  // none of them are installed.
  let embeddedFonts = new Map<string, FontRegistry>();
  const slideParts: Array<{ path: string; data: Uint8Array }> = [];
  if (presData) {
    const tree = parser.parse(decoder.decode(presData)) as Array<PoNode>;
    const pres = tree.find((n) => poIs(n, 'p:presentation'));
    embeddedFonts = loadPptxEmbeddedFonts(pkg, presPath, pres, (loss: Loss) => losses.push(loss));
    const kids = pres ? poChildren(pres) : [];
    defaultTextStyle = kids.find((c) => poIs(c, 'p:defaultTextStyle'));
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
      const slideTree = parseXml(part.data);
      const styles = slideStylesFor(
        pkg,
        part.path,
        stylesByLayout,
        resources,
        charts,
        (loss) => losses.push(loss.where ? loss : { ...loss, where: `slide ${i + 1}` }),
        defaultTextStyle,
        { widthPt: pageW, heightPt: pageH },
        overrideAlias(slideTree, 'p:sld'),
      );
      const ctx: SlideContext = {
        ...(styles.cascade ? { cascade: styles.cascade } : {}),
        colors: styles.colors,
        ...(styles.background ? { backgroundFill: styles.background } : {}),
        slideSize: { widthPt: pageW, heightPt: pageH },
        ...(styles.themeFonts ? { themeFonts: styles.themeFonts } : {}),
        ...(styles.themeFills ? { themeFills: styles.themeFills } : {}),
        ...(styles.themeLineWidths ? { themeLineWidths: styles.themeLineWidths } : {}),
        resolveImage: makeSlideImageResolver(pkg, part.path, resources),
        resolveChart: makeSlideChartResolver(pkg, part.path, charts, styles.colors, resources),
        resolveHyperlink: makeHyperlinkResolver(pkg, part.path),
        resolveDiagram: makeSlideDiagramResolver(pkg, part.path),
        resolveOlePreview: makeOlePreviewResolver(pkg, part.path, resources),
        resolveTableStyle: tableStyles,
        onLoss: (loss) => losses.push(loss.where ? loss : { ...loss, where: `slide ${i + 1}` }),
      };
      // The deck's own decoration goes under the slide's content, over its
      // background — unless this slide says it shows none (§19.3.1.38).
      const inherited = showMasterShapes(slideTree, 'p:sld') ? (styles.inheritedShapes ?? []) : [];
      body.push(...parseSlide(slideTree, ctx, styles.background, pageW, pageH, inherited));
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
    ...(embeddedFonts.size > 0 ? { embeddedFonts } : {}),
  };
  return { doc, losses };
}

/**
 * Parse an OOXML part's bytes into preserve-order {@link PoNode} roots with the
 * module's shared, presentation-tuned {@link XMLParser} (attributes kept, values
 * not coerced, whitespace preserved). Shared with the slide-style resolvers.
 *
 * ISO/IEC 29500-3 §10.2 — `mc:AlternateContent` is resolved first, to the
 * `mc:Fallback` a reader takes when it implements none of the namespaces the
 * choices require. A deck writes the same shape twice this way, and reading the
 * CHOICE means reading markup written for someone else: tdf143222's whole slide
 * is an embedded worksheet whose preview picture lives in the fallback alone,
 * and the choice — written for VML — carries no picture at all, so the slide
 * came out blank. 82 of the corpus's decks carry such a block.
 */
export function parseXml(data: Uint8Array): Array<PoNode> {
  return parser.parse(resolveAlternateContent(decoder.decode(data))) as Array<PoNode>;
}

// A slide part's bytes → a full-slide backdrop (PX5b) followed by its floating
// shapes/images/frames. The background is the slide's own p:bg, else the
// inherited layout/master one.
function parseSlide(
  tree: ReadonlyArray<PoNode>,
  ctx: SlideContext,
  inheritedBg: ShapeFill | undefined,
  pageW: Pt,
  pageH: Pt,
  inheritedShapes: ReadonlyArray<BodyElement> = [],
): Array<BodyElement> {
  const sld = tree.find((n) => poIs(n, 'p:sld'));
  const cSld = sld ? poChildren(sld).find((c) => poIs(c, 'p:cSld')) : undefined;
  const colors = ctx.colors ?? defaultColorResolver;
  const bgNode = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:bg')) : undefined;
  const own = bgNode
    ? parseBackgroundFill(bgNode, colors, ctx.resolveImage, ctx.themeFills)
    : undefined;
  const bg = own ?? inheritedBg;
  const spTree = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:spTree')) : undefined;
  const shapeCtx: SlideContext = own ? { ...ctx, backgroundFill: own } : ctx;

  const out: Array<BodyElement> = [];
  if (bg) out.push(backdropElement(bg, pageW, pageH));
  // The deck's decoration paints UNDER everything the slide itself puts on the
  // page, whatever KIND each of them is (§19.3.1).
  out.push(...inheritedShapes.map(asBackdrop));
  if (spTree) out.push(...parseSlideShapes(spTree, shapeCtx));
  return out;
}

// A layout's or master's p:cSld/p:bg → its background fill (PX5b). A picture
// background names its blip through the relationships of the part it is
// written in, which is why the resolver is bound to that part and not the slide.
function partBackground(
  tree: ReadonlyArray<PoNode>,
  root: 'p:sldLayout' | 'p:sldMaster',
  colors: ColorResolver,
  resolveImage: (relId: string) => ResourceId | undefined,
  themeFills: ThemeFillStyles | undefined,
): ShapeFill | undefined {
  const sld = tree.find((n) => poIs(n, root));
  const cSld = sld ? poChildren(sld).find((c) => poIs(c, 'p:cSld')) : undefined;
  const bg = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:bg')) : undefined;
  return bg ? parseBackgroundFill(bg, colors, resolveImage, themeFills) : undefined;
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

/**
 * The preview picture resolver for one slide's embedded objects: a `p:oleObj`
 * `@spid` → the media the slide's legacy VML drawing hangs off that shape.
 *
 * Two relationship hops, and neither is the slide's own: the slide names the
 * VML part, and the VML part names the picture. The whole part is read once,
 * lazily — most slides have no embedded object at all.
 */
function makeOlePreviewResolver(
  pkg: OpcPackage,
  slidePath: string,
  resources: ResourceStore,
): (spid: string) => ResourceId | undefined {
  let byShape: Map<string, string> | undefined;
  let vmlPath: string | undefined;
  const cache = new Map<string, ResourceId | undefined>();
  return (spid) => {
    const key = normalizeSpid(spid);
    if (cache.has(key)) return cache.get(key);
    if (byShape === undefined) {
      const rel = pkg
        .getPartRelationships(slidePath)
        .find((r) => r.type.endsWith('/vmlDrawing') || r.type.endsWith('/legacyDrawing'));
      const part = rel ? pkg.resolveRelatedPart(slidePath, rel) : undefined;
      byShape = part ? parseVmlImageRels(part.data) : new Map();
      vmlPath = part?.path;
    }
    const relId = byShape.get(key);
    const imgRel =
      relId !== undefined && vmlPath !== undefined
        ? pkg.getPartRelationships(vmlPath).find((r) => r.id === relId)
        : undefined;
    const media =
      imgRel && vmlPath !== undefined ? pkg.resolveRelatedPart(vmlPath, imgRel) : undefined;
    const id = media ? resources.put(media.data) : undefined;
    cache.set(key, id);
    return id;
  };
}

/**
 * §20.1.4.2.24 — the deck's table styles, by the GUID a table names.
 *
 * The list's own `@def` is NOT a fallback: it is the style PowerPoint applies
 * when a table is INSERTED, and the table then records that GUID itself. A
 * table that names none wears none — table-with-no-theme's two rows are bare
 * text in a plain frame, not the blue banding `@def` points at.
 *
 * A GUID the part does not hold may still be one of the GALLERY's own, which
 * every deck names and none ships (see `preset-table-styles`).
 *
 * The part is read once, on the first table in the deck.
 */
function makeTableStyleResolver(
  pkg: OpcPackage,
  presPath: string,
): (styleId: string | undefined) => PoNode | undefined {
  let byId: Map<string, PoNode> | undefined;
  return (styleId) => {
    if (byId === undefined) {
      byId = new Map();
      const rel = pkg.getPartRelationships(presPath).find((r) => r.type.endsWith('/tableStyles'));
      const part = rel ? pkg.resolveRelatedPart(presPath, rel) : undefined;
      for (const root of part ? parseXml(part.data) : []) {
        const list = poIs(root, 'a:tblStyleLst') ? root : poFindDescendant(root, 'a:tblStyleLst');
        if (!list) continue;
        for (const style of poChildren(list)) {
          const id = poIs(style, 'a:tblStyle') ? poAttr(style, 'styleId') : undefined;
          if (id !== undefined) byId.set(id, style);
        }
      }
    }
    if (styleId === undefined) return undefined;
    return byId.get(styleId) ?? presetTableStyle(styleId, (xml) => parseXml(encoder.encode(xml)));
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
  resources: ResourceStore,
): (relId: string) => string | undefined {
  const cache = new Map<string, string | undefined>();
  return (relId) => {
    if (cache.has(relId)) return cache.get(relId);
    const rel = pkg.getPartRelationships(slidePath).find((r) => r.id === relId);
    const resolved = rel ? pkg.resolveRelatedPart(slidePath, rel) : undefined;
    // A picture the chart paints itself with is named through the CHART part's
    // own relationships, not the slide's (chart-texture-bg.pptx).
    const chart = resolved
      ? parseChart(resolved.data, colors, makeSlideImageResolver(pkg, resolved.path, resources))
      : null;
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
    const draw = data ? drawingPart(pkg, slidePath, data) : undefined;
    if (draw) {
      for (const root of parseXml(draw.data)) {
        const found = poFindDescendant(root, 'dsp:spTree');
        if (found) {
          spTree = found;
          break;
        }
      }
    } else if (data) {
      // No cached drawing: run the layout the file DOES carry. A generator
      // writes data, layout, colours and style and leaves the picture to the
      // reader, which is the whole reason this engine exists.
      spTree = laidOutDiagram(pkg, slidePath, data);
    }
    cache.set(relId, spTree);
    return spTree;
  };
}

// §21.4.3 — the layout part beside the data, run to the boxes it describes.
// The frame is the diagram's own EMU box; the caller scales it to the shape's,
// so a square frame keeps the proportions the layout was written for.
const DIAGRAM_FRAME = { cx: 5486400, cy: 3200400 };

function laidOutDiagram(
  pkg: OpcPackage,
  slidePath: string,
  data: { readonly path: string; readonly data: Uint8Array },
): PoNode | undefined {
  const layoutRel = pkg
    .getPartRelationships(slidePath)
    .find((r) => r.type.endsWith('/diagramLayout'));
  const layout = layoutRel ? pkg.resolveRelatedPart(slidePath, layoutRel) : undefined;
  if (!layout) return undefined;
  const model = new DiagramData(parseXml(data.data));
  const nodes = layoutDiagram(parseXml(layout.data), model, DIAGRAM_FRAME.cx, DIAGRAM_FRAME.cy);
  if (nodes.length === 0) return undefined;
  for (const root of parseXml(new TextEncoder().encode(diagramDrawingXml(nodes)))) {
    const found = poFindDescendant(root, 'dsp:spTree');
    if (found) return found;
  }
  return undefined;
}

/**
 * The pre-rendered drawing a diagram's data part points at.
 *
 * PowerPoint names it from INSIDE the data: `<dsp:dataModelExt relId>` in the
 * data's extension list, and that id is a relationship of the SLIDE, not of the
 * data part. Looked for on the data part alone — where some producers do put it
 * — two corpus decks with a `drawing1.xml` sitting right there reported having
 * no drawing at all and rendered nothing (smartart-missing-bullet,
 * tdf145528_SmartArt_Matrix).
 *
 * @param pkg       The package.
 * @param slidePath The slide the diagram is on.
 * @param data      The resolved `dgm:relIds@r:dm` data part.
 * @returns The drawing part, or undefined when the deck ships none.
 */
function drawingPart(
  pkg: OpcPackage,
  slidePath: string,
  data: { readonly path: string; readonly data: Uint8Array },
): { readonly path: string; readonly data: Uint8Array } | undefined {
  // Every root, not just the first: with `preserveOrder` the XML declaration is
  // a node of its own, so a search that starts at [0] finds nothing and every
  // diagram on the slide fell through to the same fallback drawing —
  // tdf125551 has four and drew one of them four times.
  const extRelId = parseXml(data.data)
    .map((root) => poFindDescendant(root, 'dsp:dataModelExt'))
    .map((ext) => (ext ? poAttr(ext, 'relId') : undefined))
    .find((id) => id !== undefined);
  const bySlide =
    extRelId !== undefined
      ? pkg.getPartRelationships(slidePath).find((r) => r.id === extRelId)
      : undefined;
  if (bySlide) {
    const part = pkg.resolveRelatedPart(slidePath, bySlide);
    if (part) return part;
  }
  const own = pkg.getPartRelationships(data.path).find((r) => r.type.endsWith('/diagramDrawing'));
  if (own) return pkg.resolveRelatedPart(data.path, own);
  // Last resort: the slide's own drawing relationship — but only when there is
  // exactly ONE, since with several there is nothing to say which is whose.
  const onSlide = pkg
    .getPartRelationships(slidePath)
    .filter((r) => r.type.endsWith('/diagramDrawing'));
  return onSlide.length === 1 && onSlide[0]
    ? pkg.resolveRelatedPart(slidePath, onSlide[0])
    : undefined;
}

// The placeholder cascade + colour resolver for a slide, derived from its
// slideLayout → slideMaster (→ theme) chain and memoized by layout path (slides
// that share a layout share both).
interface SlideStyles {
  readonly cascade?: PlaceholderCascade;
  readonly colors: ColorResolver;
  /**
   * §19.3.1.38 — the shapes a slide inherits: the master's, then the layout's,
   * both without their placeholders, in the order they are drawn UNDER the
   * slide's own content.
   */
  readonly inheritedShapes?: ReadonlyArray<BodyElement>;
  /** The deck theme's `a:fillStyleLst`/`a:bgFillStyleLst`, for a `p:bgRef`. */
  readonly themeFills?: ThemeFillStyles;
  /** §20.1.4.1.16 — the two typefaces a `+mn-lt`/`+mj-lt` token stands for. */
  readonly themeFonts?: ThemeFonts;
  /** §20.1.4.1.21 — the widths an `a:lnRef` indexes, in points. */
  readonly themeLineWidths?: ReadonlyArray<number>;
  // The inherited background fill (layout, else master) for slides that have no
  // p:bg of their own (PX5b).
  readonly background?: ShapeFill;
}

function slideStylesFor(
  pkg: OpcPackage,
  slidePath: string,
  cache: Map<string, SlideStyles>,
  resources: ResourceStore,
  charts: Map<string, Chart>,
  onLoss: (loss: Loss) => void,
  deckDefaultTextStyle: PoNode | undefined,
  slideSize: { readonly widthPt: Pt; readonly heightPt: Pt },
  slideAlias?: SchemeAliases,
): SlideStyles {
  const layoutRel = pkg
    .getPartRelationships(slidePath)
    .find((r) => r.type.endsWith('/slideLayout'));
  const layout = layoutRel ? pkg.resolveRelatedPart(slidePath, layoutRel) : undefined;
  if (!layout) return { colors: defaultColorResolver };
  // Slides sharing a layout share everything below — unless one of them states
  // a colour map of its own, which makes every colour in the chain read
  // differently, so that slide gets its own entry.
  const key = slideAlias ? `${layout.path}|${aliasKey(slideAlias)}` : layout.path;
  const cached = cache.get(key);
  if (cached) return cached;

  const masterRel = pkg
    .getPartRelationships(layout.path)
    .find((r) => r.type.endsWith('/slideMaster'));
  const master = masterRel ? pkg.resolveRelatedPart(layout.path, masterRel) : undefined;
  const layoutTree = parseXml(layout.data);
  const masterTree = master ? parseXml(master.data) : undefined;
  // §19.3.1.6/§19.3.1.7 — the master states which theme slot each of bg1/tx1/
  // bg2/tx2 means, and a layout or a slide may override that mapping. The
  // nearest one wins.
  //
  // A slide's override governs the SLIDE, though, not the design it sits on:
  // what the master and the layout draw keeps reading under their own map.
  // chart_pt_color_bg1 flips bg1 to dk1 for its own content, and applied to
  // the master's background as well it turned a white deck black.
  const inheritedAlias =
    overrideAlias(layoutTree, 'p:sldLayout') ?? (masterTree ? masterAlias(masterTree) : undefined);
  const alias = slideAlias ?? inheritedAlias;
  const theme = master ? themePart(pkg, master.path) : undefined;
  const colors = master ? deckColorResolver(theme?.data, alias) : defaultColorResolver;
  const inheritedColors =
    slideAlias === undefined || !master ? colors : deckColorResolver(theme?.data, inheritedAlias);
  const themeFills = theme
    ? {
        fills: parseThemeFillStyles(theme.data),
        backgrounds: parseThemeBgFillStyles(theme.data),
        resolveImage: makeSlideImageResolver(pkg, theme.path, resources),
      }
    : undefined;
  const themeLineWidths = theme ? parseThemeLineWidths(theme.data) : undefined;
  // §20.1.4.1.16 — the two typefaces the deck names once; a slide refers to
  // them by token (`+mn-lt`), which is what most of its runs actually say.
  const themeFonts = theme ? parseThemeFonts(theme.data) : undefined;
  const cascade = buildPlaceholderCascade(
    layoutTree,
    masterTree,
    colors,
    parseLevelStyles(deckDefaultTextStyle, colors, themeFonts),
    themeFonts,
  );
  const background =
    partBackground(
      layoutTree,
      'p:sldLayout',
      inheritedColors,
      makeSlideImageResolver(pkg, layout.path, resources),
      themeFills,
    ) ??
    (masterTree && master
      ? partBackground(
          masterTree,
          'p:sldMaster',
          inheritedColors,
          makeSlideImageResolver(pkg, master.path, resources),
          themeFills,
        )
      : undefined);
  // The background is resolved FIRST: a shape among the inherited ones may be
  // painted with it (§19.3.1.43 `useBgFill`).
  const deps: PartDeps = {
    colors: inheritedColors,
    resources,
    charts,
    onLoss,
    slideSize,
    ...(themeFills ? { themeFills } : {}),
    ...(themeLineWidths && themeLineWidths.length > 0 ? { themeLineWidths } : {}),
    ...(background ? { background } : {}),
  };
  // §19.3.1.38 `@showMasterSp` — the decoration a deck states once and every
  // slide carries: rules, bands, a logo. A layout may refuse the master's, and
  // a slide may refuse both (read at the slide, below).
  const inheritedShapes: Array<BodyElement> = [];
  if (masterTree && master && showMasterShapes(layoutTree, 'p:sldLayout')) {
    inheritedShapes.push(...partShapes(pkg, master.path, masterTree, 'p:sldMaster', deps));
  }
  inheritedShapes.push(...partShapes(pkg, layout.path, layoutTree, 'p:sldLayout', deps));
  const styles: SlideStyles = {
    cascade,
    colors,
    ...(themeFonts ? { themeFonts } : {}),
    ...(themeFills ? { themeFills } : {}),
    ...(themeLineWidths && themeLineWidths.length > 0 ? { themeLineWidths } : {}),
    ...(background ? { background } : {}),
    ...(inheritedShapes.length > 0 ? { inheritedShapes } : {}),
  };
  cache.set(key, styles);
  return styles;
}

/**
 * A master's or layout's own drawn shapes — everything in its `p:spTree` that
 * is not a placeholder, read with resolvers bound to THAT part: a logo's image
 * relationship, a chart's, a hyperlink's are all scoped to the part that names
 * them.
 */
function partShapes(
  pkg: OpcPackage,
  path: string,
  tree: ReadonlyArray<PoNode>,
  root: 'p:sldLayout' | 'p:sldMaster',
  deps: PartDeps,
): Array<BodyElement> {
  const part = tree.find((n) => poIs(n, root));
  const cSld = part ? poChildren(part).find((c) => poIs(c, 'p:cSld')) : undefined;
  const spTree = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:spTree')) : undefined;
  if (!spTree) return [];
  const ctx: SlideContext = {
    colors: deps.colors,
    ...(deps.themeFills ? { themeFills: deps.themeFills } : {}),
    ...(deps.themeLineWidths ? { themeLineWidths: deps.themeLineWidths } : {}),
    ...(deps.background ? { backgroundFill: deps.background } : {}),
    slideSize: deps.slideSize,
    resolveImage: makeSlideImageResolver(pkg, path, deps.resources),
    resolveChart: makeSlideChartResolver(pkg, path, deps.charts, deps.colors, deps.resources),
    resolveHyperlink: makeHyperlinkResolver(pkg, path),
    resolveDiagram: makeSlideDiagramResolver(pkg, path),
    onLoss: deps.onLoss,
  };
  return parseSlideShapes(spTree, ctx, undefined, true);
}

/** What {@link partShapes} needs from the document being read. */
interface PartDeps {
  readonly colors: ColorResolver;
  readonly themeFills?: ThemeFillStyles;
  readonly themeLineWidths?: ReadonlyArray<number>;
  /** The background these shapes sit on, for a `useBgFill` one among them. */
  readonly background?: ShapeFill;
  /** The slide's size, which says where that background lies under a shape. */
  readonly slideSize: { readonly widthPt: Pt; readonly heightPt: Pt };
  readonly resources: ResourceStore;
  readonly charts: Map<string, Chart>;
  readonly onLoss: (loss: Loss) => void;
}

/**
 * §19.3.1.38/§19.3.1.39 `@showMasterSp` — whether this part draws the shapes it
 * inherits. Absent, it does.
 */
function showMasterShapes(tree: ReadonlyArray<PoNode>, root: 'p:sldLayout' | 'p:sld'): boolean {
  const part = tree.find((n) => poIs(n, root));
  return part === undefined || poAttr(part, 'showMasterSp') !== '0';
}

/** The master's `p:clrMap` (§19.3.1.6) as a scheme-alias table. */
function masterAlias(masterTree: ReadonlyArray<PoNode>): SchemeAliases | undefined {
  const master = masterTree.find((n) => poIs(n, 'p:sldMaster'));
  const map = master ? poChildren(master).find((c) => poIs(c, 'p:clrMap')) : undefined;
  return map ? aliasFromAttrs(map) : undefined;
}

/**
 * A layout's or slide's `p:clrMapOvr` (§19.3.1.7) as a scheme-alias table, or
 * `undefined` when it holds `a:masterClrMapping` — which says "the master's",
 * i.e. nothing of its own.
 */
function overrideAlias(
  tree: ReadonlyArray<PoNode>,
  root: 'p:sldLayout' | 'p:sld',
): SchemeAliases | undefined {
  const part = tree.find((n) => poIs(n, root));
  const ovr = part ? poChildren(part).find((c) => poIs(c, 'p:clrMapOvr')) : undefined;
  const own = ovr ? poChildren(ovr).find((c) => poIs(c, 'a:overrideClrMapping')) : undefined;
  return own ? aliasFromAttrs(own) : undefined;
}

// Both spellings of the map carry the same twelve attributes: name → theme slot.
function aliasFromAttrs(node: PoNode): SchemeAliases {
  const alias: Record<string, string> = {};
  for (const name of ['bg1', 'tx1', 'bg2', 'tx2']) {
    const slot = poAttr(node, name);
    if (slot !== undefined) alias[name] = slot;
  }
  return alias;
}

/** A stable key for an alias table, so the per-layout cache can include it. */
function aliasKey(alias: SchemeAliases): string {
  return Object.entries(alias)
    .sort(([a], [b]) => (a < b ? -1 : 1))
    .map(([k, v]) => `${k}=${v}`)
    .join(',');
}

// The slide master's theme part (a:clrScheme) → a ColorResolver: the built-in
// Office palette with the deck's scheme colours merged over it (mirrors the
// docx/xlsx readers). The Office palette stands in when there is no theme, and
// `alias` is the deck's own reading of the bg/tx names.
function deckColorResolver(theme: Uint8Array | undefined, alias: SchemeAliases | undefined) {
  const palette = new Map(DEFAULT_THEME_PALETTE);
  if (theme) for (const [slot, hex] of parseTheme(theme)) palette.set(slot, hex);
  if (!theme && !alias) return defaultColorResolver;
  return makeColorResolver(palette, { ...DEFAULT_SCHEME_ALIAS, ...alias });
}

/** The master's theme part, if it has one — its path names its own pictures. */
function themePart(
  pkg: OpcPackage,
  masterPath: string,
): { readonly path: string; readonly data: Uint8Array } | undefined {
  const rel = pkg.getPartRelationships(masterPath).find((r) => r.type.endsWith('/theme'));
  return rel ? pkg.resolveRelatedPart(masterPath, rel) : undefined;
}

/**
 * The {@link DocumentReader} for PresentationML (`.pptx`): sniffs the OPC ZIP for
 * `ppt/presentation.xml` (a cheap substring probe, no unzip) and reads it via
 * {@link readPptx}.
 */
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
  // A pptx is a ZIP whose parts include ppt/presentation.xml — read off the
  // archive's own directory, no unzip needed (mirrors the docx/xlsx sniff).
  sniff: (bytes) =>
    bytes[0] === 0x50 && bytes[1] === 0x4b && packageHasPart(bytes, 'ppt/presentation.xml'),
  read: (bytes) => readPptx(bytes),
};
