// E-PPTX PX0–PX2 — PresentationML (.pptx) reader: bytes → FlowDoc. A presentation
// is a positioned canvas, which maps cleanly onto the existing IR: each slide is a
// section at the deck's page size, and its shapes become absolutely positioned
// floating elements (Route A, epics.md). PX0 established the seam — sniff, slide
// size from p:sldSz, slide count from p:sldIdLst, one page per slide. PX1 fills
// each page: a slide's text-bearing shapes (p:sp/p:txBody) become positioned
// floating text boxes. PX2 resolves placeholders against the slide's layout and
// master (the cascade), so inherited geometry + text sizing land too.

import { XMLParser } from 'fast-xml-parser';

import type { BodyElement, SectionProperties } from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import type { Relationship } from '@/core/opc';
import type { PlaceholderCascade } from '@/pptx/placeholder-cascade';

import { bytesInclude } from '@/core/bytes';
import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { OpcPackage } from '@/core/opc';
import { poAttr, poChildren, poIntAttr, poIs } from '@/core/po-helpers';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
import { buildPlaceholderCascade } from '@/pptx/placeholder-cascade';
import { parseSlideShapes } from '@/pptx/slide-parser';

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

export function readPptx(bytes: Uint8Array): ReadResult<FlowDoc> {
  const losses: Array<Loss> = [];
  const pkg = OpcPackage.open(bytes);
  const presPath = pkg.getMainDocumentPath();
  const presData = pkg.getPart(presPath);

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
    for (const sldId of ids) {
      const rid = poAttr(sldId, 'r:id');
      const rel = rid !== undefined ? slideRelById.get(rid) : undefined;
      const part = rel ? pkg.resolveRelatedPart(presPath, rel) : undefined;
      if (part) slideParts.push(part);
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
  const cascadeByLayout = new Map<string, PlaceholderCascade | undefined>();
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
      const cascade = cascadeForSlide(pkg, part.path, cascadeByLayout);
      body.push(...parseSlide(part.data, cascade));
    }
  }

  const section: SectionProperties = {
    pageSize: { width: pt(cx / EMU_PER_PT), height: pt(cy / EMU_PER_PT) },
    headers: [],
    footers: [],
  };
  const doc: FlowDoc = {
    kind: 'flow',
    body: resolveBodyStyles([...body], EMPTY_STYLE_SHEET),
    sections: [],
    section,
    styles: EMPTY_STYLE_SHEET,
    resources: new ResourceStore(),
  };
  return { doc, losses };
}

function parseXml(data: Uint8Array): Array<PoNode> {
  return parser.parse(decoder.decode(data)) as Array<PoNode>;
}

// A slide part's bytes → its floating shapes. Navigates p:sld/p:cSld/p:spTree
// and hands the shape tree to the slide parser, with the placeholder cascade.
function parseSlide(data: Uint8Array, cascade?: PlaceholderCascade): Array<BodyElement> {
  const tree = parseXml(data);
  const sld = tree.find((n) => poIs(n, 'p:sld'));
  const cSld = sld ? poChildren(sld).find((c) => poIs(c, 'p:cSld')) : undefined;
  const spTree = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:spTree')) : undefined;
  return spTree ? parseSlideShapes(spTree, cascade) : [];
}

// Resolve (and memoize) the placeholder cascade for a slide: its slideLayout
// rel, then the layout's slideMaster rel. Cached by layout path — slides that
// share a layout share the cascade. Undefined when the slide has no layout.
function cascadeForSlide(
  pkg: OpcPackage,
  slidePath: string,
  cache: Map<string, PlaceholderCascade | undefined>,
): PlaceholderCascade | undefined {
  const layoutRel = pkg
    .getPartRelationships(slidePath)
    .find((r) => r.type.endsWith('/slideLayout'));
  const layout = layoutRel ? pkg.resolveRelatedPart(slidePath, layoutRel) : undefined;
  if (!layout) return undefined;
  if (cache.has(layout.path)) return cache.get(layout.path);

  const masterRel = pkg
    .getPartRelationships(layout.path)
    .find((r) => r.type.endsWith('/slideMaster'));
  const master = masterRel ? pkg.resolveRelatedPart(layout.path, masterRel) : undefined;
  const cascade = buildPlaceholderCascade(
    parseXml(layout.data),
    master ? parseXml(master.data) : undefined,
  );
  cache.set(layout.path, cascade);
  return cascade;
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
