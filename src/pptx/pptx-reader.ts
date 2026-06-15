// E-PPTX PX0 — PresentationML (.pptx) reader: bytes → FlowDoc. A presentation is
// a positioned canvas, which maps cleanly onto the existing IR: each slide is a
// section at the deck's page size, and (from PX1 on) its shapes become absolutely
// positioned floating elements. PX0 establishes the seam — sniff, the slide size
// from p:sldSz, the slide count from p:sldIdLst, and one page per slide.

import { XMLParser } from 'fast-xml-parser';

import type { BodyElement, SectionProperties } from '@/core/document-model';
import type { DocumentReader, ReadResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss } from '@/core/ir';
import type { PoNode } from '@/core/po-helpers';

import { bytesInclude } from '@/core/bytes';
import { FEATURES, ResourceStore, pt } from '@/core/ir';
import { OpcPackage } from '@/core/opc';
import { poAttr, poChildren, poIntAttr, poIs } from '@/core/po-helpers';
import { EMPTY_STYLE_SHEET, resolveBodyStyles } from '@/core/style-cascade';

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
  let slideCount = 0;
  if (presData) {
    const tree = parser.parse(decoder.decode(presData)) as Array<PoNode>;
    const pres = tree.find((n) => poIs(n, 'p:presentation'));
    const kids = pres ? poChildren(pres) : [];
    const sldSz = kids.find((c) => poIs(c, 'p:sldSz'));
    cx = (sldSz ? poIntAttr(sldSz, 'cx') : undefined) || DEFAULT_CX;
    cy = (sldSz ? poIntAttr(sldSz, 'cy') : undefined) || DEFAULT_CY;
    // Slide order is p:sldIdLst/p:sldId@r:id, each resolving to a /slide rel of
    // the presentation part.
    const slideRelIds = new Set(
      pkg
        .getPartRelationships(presPath)
        .filter((r) => r.type.endsWith('/slide'))
        .map((r) => r.id),
    );
    const lst = kids.find((c) => poIs(c, 'p:sldIdLst'));
    const ids = lst ? poChildren(lst).filter((c) => poIs(c, 'p:sldId')) : [];
    slideCount = ids.filter((s) => {
      const rid = poAttr(s, 'r:id');
      return rid !== undefined && slideRelIds.has(rid);
    }).length;
  }
  if (slideCount === 0) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.text,
      detail: 'presentation has no slides',
    });
    slideCount = 1;
  }

  // One page per slide at the deck size. Each slide is a paragraph that starts a
  // new page (pageBreakBefore after the first); a single zero-width-space run
  // gives the otherwise-blank slide one line so the page is actually emitted. The
  // slide's shape tree replaces this placeholder from PX1 on.
  const body: Array<BodyElement> = [];
  for (let i = 0; i < slideCount; i++) {
    body.push({
      kind: 'paragraph',
      paragraph: {
        properties: i > 0 ? { pageBreakBefore: true } : {},
        runs: [{ text: '​', properties: {} }],
      },
    });
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
