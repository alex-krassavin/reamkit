// E-PPTX PX0/PX1 — PresentationML (.pptx) reader: bytes → FlowDoc. A presentation
// is a positioned canvas, which maps cleanly onto the existing IR: each slide is a
// section at the deck's page size, and its shapes become absolutely positioned
// floating elements (Route A, epics.md). PX0 established the seam — sniff, slide
// size from p:sldSz, slide count from p:sldIdLst, one page per slide. PX1 fills
// each page: a slide's text-bearing shapes (p:sp/p:txBody) become positioned
// floating text boxes (slide-parser.ts).

import { XMLParser } from 'fast-xml-parser';

import type { BodyElement, SectionProperties } from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';
import type { Relationship } from '@/core/opc';

import { bytesInclude } from '@/core/bytes';
import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { OpcPackage } from '@/core/opc';
import { poAttr, poChildren, poIntAttr, poIs } from '@/core/po-helpers';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';
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
  const slideParts: Array<Uint8Array> = [];
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
      if (part) slideParts.push(part.data);
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
  const body: Array<BodyElement> = [];
  for (let i = 0; i < slideCount; i++) {
    body.push({
      kind: 'paragraph',
      paragraph: {
        properties: i > 0 ? { pageBreakBefore: true } : {},
        runs: [{ text: '​', properties: {} }],
      },
    });
    const data = slideParts[i];
    if (data) body.push(...parseSlide(data));
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

// A slide part's bytes → its floating shapes. Navigates p:sld/p:cSld/p:spTree
// and hands the shape tree to the slide parser.
function parseSlide(data: Uint8Array): Array<BodyElement> {
  const tree = parser.parse(decoder.decode(data)) as Array<PoNode>;
  const sld = tree.find((n) => poIs(n, 'p:sld'));
  const cSld = sld ? poChildren(sld).find((c) => poIs(c, 'p:cSld')) : undefined;
  const spTree = cSld ? poChildren(cSld).find((c) => poIs(c, 'p:spTree')) : undefined;
  return spTree ? parseSlideShapes(spTree) : [];
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
