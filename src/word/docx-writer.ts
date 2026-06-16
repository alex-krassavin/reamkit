// docx writer (E-DOCX): FlowDoc → WordprocessingML package — the inverse of
// the docx reader, and the fifth adapter overall. A flow medium with zero
// layout and zero I/O, like the HTML writer.
//
// v1 contract (epics.md, variant A): the writer emits a DENORMALIZED but
// valid document. FlowDoc's body carries RESOLVED properties (the stage-6
// cascade is already collapsed), so what we write is direct formatting — no
// named styles. The round-trip guarantee is therefore semantic, not textual:
// readDocx(writeDocx(flow)) yields an equivalent FlowDoc, never the original
// bytes. Anything the writer does not serialize yet is reported as a loss,
// exactly like the other writers.
//
// Coverage (the docx-writer epic, D1–D7 + T1–T3): paragraphs and runs with
// full formatting, page breaks, numbered lists, hyperlinks and bookmarks,
// tables (spans, borders, shading, nesting), images of every format the
// package carries (raster PNG/JPEG/JPEG2000/GIF/BMP/TIFF and vector EMF/WMF,
// plus an embedded PDF picture — see mediaInfo), DrawingML shapes (preset and
// custom geometry, fill, line, text body — see shapeDrawingXml),
// headers/footers, and multi-section geometry (per-section sectPr —
// mid-document breaks ride the section's last paragraph's pPr, the final
// section a body-level sectPr; page size/margins, columns, titlePg). The
// round-trip gate proves zero writer failures across 1100 corpus documents,
// 1099 of them a full IR identity (POI 110/110, LibreOffice 989/990 — the one
// miss is an input whose referenced image part was stripped from the package,
// so the bytes do not exist to carry). Footnotes/endnotes (WT2), charts and
// OfficeMath (WT3) all write back; a shape round-trips as inline (floating
// placement is dropped).

import { omathXml } from './omml-serializer';
import type {
  BodyElement,
  CellBorders,
  CellMargins,
  CellProperties,
  Chart,
  ChartBlock,
  Comment,
  FontFamilyMap,
  Numbering,
  NumberingLevel,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  SectionColumns,
  SectionProperties,
  ShapeBlock,
  ShapeFill,
  ShapeGeometry,
  ShapeLine,
  ShapeTextBody,
  ShapeTransform,
  Table,
  TableCell,
  TableProperties,
  TableRow,
} from '@/core/document-model';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';
import type { ShapeGradient } from '@/core/vector';
import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss, ResourceId, ResourceStore } from '@/core/ir';
import type { OpcPart, Relationship } from '@/core/opc';

import { FEATURES } from '@/core/ir';
import { chartSpaceXml } from '@/core/drawingml/chart-serializer';
import { detectImageFormat } from '@/core/images';
import { buildOpcPackage } from '@/core/opc';
import {
  EMPTY_STYLE_SHEET,
  resolveParagraphProperties,
  resolveRunProperties,
} from '@/core/style-cascade';

const encoder = new TextEncoder();

// The reader stored RESOLVED properties back onto each run/paragraph (stage
// 6). The defaults are what the same resolver yields for empty input over the
// empty sheet — a field equal to these is implicit and is NOT serialized, so a
// re-read materializes the same value. This delta keeps the emitted rPr/pPr
// minimal and the round-trip an IR identity.
const DEFAULT_RUN = resolveRunProperties({}, {}, EMPTY_STYLE_SHEET);
const DEFAULT_PARA = resolveParagraphProperties({}, EMPTY_STYLE_SHEET);

const DOC_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const NUMBERING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const REL_OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_NUMBERING =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const REL_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const NUMBERING_PART = 'word/numbering.xml';
const REL_FOOTNOTES =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes';
const REL_ENDNOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes';
const FOOTNOTES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml';
const ENDNOTES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml';
const FOOTNOTES_PART = 'word/footnotes.xml';
const ENDNOTES_PART = 'word/endnotes.xml';
const REL_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_PART = 'word/comments.xml';
// Microsoft commentsExtended (w15) — the reply/resolved thread map (CM4).
const REL_COMMENTS_EXTENDED =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const COMMENTS_EXTENDED_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const COMMENTS_EXTENDED_PART = 'word/commentsExtended.xml';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

// 1 pt = 12700 EMU (English Metric Units, the DrawingML coordinate).
const EMU_PER_PT = 12700;

// The raster formats the PDF path embeds (detectImageFormat) → media naming.
const RASTER_MEDIA: Readonly<Record<string, { ext: string; contentType: string }>> = {
  png: { ext: 'png', contentType: 'image/png' },
  jpeg: { ext: 'jpeg', contentType: 'image/jpeg' },
  jpeg2000: { ext: 'jp2', contentType: 'image/jp2' },
};

// The writer round-trips a docx; it transfers image bytes verbatim, so it
// names media files for EVERY OOXML image format — including the vector /
// legacy ones the PDF path cannot render (GIF, BMP, TIFF, EMF, WMF). The
// reader stores the bytes regardless of format, so this is the only place
// format knowledge is needed on the write side.
function mediaInfo(bytes: Uint8Array): { ext: string; contentType: string } | undefined {
  const raster = detectImageFormat(bytes);
  if (raster) return RASTER_MEDIA[raster];
  const b = (i: number): number => bytes[i] ?? -1;
  // GIF — "GIF8".
  if (b(0) === 0x47 && b(1) === 0x49 && b(2) === 0x46 && b(3) === 0x38) {
    return { ext: 'gif', contentType: 'image/gif' };
  }
  // BMP — "BM".
  if (b(0) === 0x42 && b(1) === 0x4d) return { ext: 'bmp', contentType: 'image/bmp' };
  // TIFF — "II*\0" (little-endian) or "MM\0*" (big-endian).
  if ((b(0) === 0x49 && b(1) === 0x49 && b(2) === 0x2a) || (b(0) === 0x4d && b(1) === 0x4d)) {
    return { ext: 'tiff', contentType: 'image/tiff' };
  }
  // EMF — EMR_HEADER record (iType=1) with the " EMF" signature at byte 40.
  if (
    b(0) === 0x01 &&
    b(1) === 0 &&
    b(2) === 0 &&
    b(3) === 0 &&
    b(40) === 0x20 &&
    b(41) === 0x45 &&
    b(42) === 0x4d &&
    b(43) === 0x46
  ) {
    return { ext: 'emf', contentType: 'image/x-emf' };
  }
  // WMF — Aldus placeable header (D7 CD C6 9A) or a standard metafile header.
  if (
    (b(0) === 0xd7 && b(1) === 0xcd && b(2) === 0xc6 && b(3) === 0x9a) ||
    (b(0) === 0x01 && b(1) === 0x00 && b(2) === 0x09 && b(3) === 0x00)
  ) {
    return { ext: 'wmf', contentType: 'image/x-wmf' };
  }
  // PDF — "%PDF". Word/LibreOffice embed a PDF as a picture (with a raster
  // fallback for display); we carry the bytes for the round-trip.
  if (b(0) === 0x25 && b(1) === 0x50 && b(2) === 0x44 && b(3) === 0x46) {
    return { ext: 'pdf', contentType: 'application/pdf' };
  }
  return undefined;
}

const HEADER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FOOTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';

// Document-global state (shared across every part): the resource store, the
// shared word/media parts (content-addressed file dedup), a document-wide
// bookmark and drawing id counter.
interface WriteState {
  readonly resources: ResourceStore;
  readonly mediaParts: Array<OpcPart>;
  // ResourceId → its shared media target (relative to word/, e.g.
  // 'media/image1.png'); a resource yields one file regardless of who uses it.
  readonly mediaFileByResource: Map<ResourceId, string>;
  bookmarkSeq: number;
  drawingSeq: number;
  // §21.2 charts (WT3): the parsed chart data by part path, plus the chart parts
  // emitted while serializing the body, and a global chart-part counter.
  readonly charts?: ReadonlyMap<string, Chart>;
  readonly chartParts: Array<OpcPart>;
  chartSeq: number;
}

// Per-PART relationship scope (OPC §9.3 — rIds are scoped to their owning
// part). document.xml has one; each header/footer part has its own, so a media
// reference resolves against the right .rels.
interface PartScope {
  readonly rels: Array<Relationship>;
  relSeq: number;
  // Set while emitting a footnotes/endnotes part, so a note-number run (WT2)
  // emits the right §17.11.13/.5 mark.
  noteKind?: 'footnote' | 'endnote';
  // ResourceId → the rId allocated FOR THIS PART (distinct from the shared file).
  readonly relIdByResource: Map<ResourceId, string>;
}

function newScope(): PartScope {
  return { rels: [], relSeq: 0, relIdByResource: new Map() };
}

export function writeDocx(flow: FlowDoc): WriteResult {
  const losses: Array<Loss> = [];
  const body: Array<string> = [];
  const state: WriteState = {
    resources: flow.resources,
    mediaParts: [],
    mediaFileByResource: new Map(),
    bookmarkSeq: 0,
    drawingSeq: 0,
    chartParts: [],
    chartSeq: 0,
    ...(flow.charts ? { charts: flow.charts } : {}),
  };
  const docScope = newScope();
  const extraParts: Array<OpcPart> = [];
  const extraPartRels: Array<{ sourcePart: string; relationships: Array<Relationship> }> = [];
  // Header/footer parts referenced by more than one section are emitted once
  // (original relationship id → the document rId reused everywhere).
  const hfCache = new Map<string, string>();

  // §17.6.17 — sections. Each section's sectPr carries its own headers/footers.
  // A mid-document section's sectPr lives in the pPr of its LAST paragraph
  // (body[endIndex-1], the carrier the reader keeps); the final section's
  // sectPr is a direct body child.
  const sections =
    flow.sections.length > 0
      ? flow.sections
      : flow.section
        ? [{ properties: flow.section, endIndex: flow.body.length }]
        : [];
  const sectPrByClosingIndex = new Map<number, string>();
  let finalSectPr = '';
  sections.forEach((sec, i) => {
    const refs = emitHeadersFooters(
      flow,
      sec.properties,
      state,
      docScope,
      extraParts,
      extraPartRels,
      hfCache,
      losses,
    );
    const sp = sectPrXml(sec.properties, refs);
    if (i === sections.length - 1) finalSectPr = sp;
    else sectPrByClosingIndex.set(sec.endIndex - 1, sp);
  });

  flow.body.forEach((el, idx) => {
    emitBlock(body, el, losses, state, docScope, sectPrByClosingIndex.get(idx));
  });
  if (finalSectPr) body.push(finalSectPr);

  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:body>${body.join('')}</w:body>` +
    '</w:document>';

  // §17.9 numbering: re-emit the raw definitions whenever a paragraph carries
  // a list reference (the markers were stripped above — re-read regenerates
  // them). Lives at the fixed word/numbering.xml path the reader expects.
  const usesNumbering = flow.body.some(
    (el) => el.kind === 'paragraph' && el.paragraph.properties.numbering !== undefined,
  );
  const numberingPart =
    usesNumbering && flow.numbering
      ? {
          path: NUMBERING_PART,
          data: encoder.encode(numberingXml(flow.numbering)),
          contentType: NUMBERING_CONTENT_TYPE,
        }
      : undefined;
  if (numberingPart) {
    docScope.rels.push({
      id: `rId${++docScope.relSeq}`,
      type: REL_NUMBERING,
      target: 'numbering.xml',
      targetMode: 'Internal',
    });
  }

  // §17.11 footnotes / endnotes (WT2): emit the parts + a document relationship
  // whenever the document carries note content.
  emitNotes(
    flow.footnotes,
    {
      noteKind: 'footnote',
      partPath: FOOTNOTES_PART,
      rootTag: 'w:footnotes',
      noteTag: 'w:footnote',
      contentType: FOOTNOTES_CONTENT_TYPE,
      relType: REL_FOOTNOTES,
      target: 'footnotes.xml',
    },
    state,
    losses,
    docScope,
    extraParts,
    extraPartRels,
  );
  emitNotes(
    flow.endnotes,
    {
      noteKind: 'endnote',
      partPath: ENDNOTES_PART,
      rootTag: 'w:endnotes',
      noteTag: 'w:endnote',
      contentType: ENDNOTES_CONTENT_TYPE,
      relType: REL_ENDNOTES,
      target: 'endnotes.xml',
    },
    state,
    losses,
    docScope,
    extraParts,
    extraPartRels,
  );
  // §17.13.4 review comments → word/comments.xml + a document relationship (CM3),
  // plus commentsExtended.xml for reply threads and resolved flags (CM4).
  const commentParaIds = emitComments(
    flow.comments,
    state,
    losses,
    docScope,
    extraParts,
    extraPartRels,
  );
  emitCommentsExtended(flow.comments, commentParaIds, docScope, extraParts);

  const partRelationships = [
    ...(docScope.rels.length > 0
      ? [{ sourcePart: 'word/document.xml', relationships: docScope.rels }]
      : []),
    ...extraPartRels,
  ];

  const bytes = buildOpcPackage({
    parts: [
      {
        path: 'word/document.xml',
        data: encoder.encode(documentXml),
        contentType: DOC_CONTENT_TYPE,
      },
      ...(numberingPart ? [numberingPart] : []),
      ...extraParts,
      ...state.chartParts,
      ...state.mediaParts,
    ],
    rootRelationships: [
      {
        id: 'rId1',
        type: REL_OFFICE_DOCUMENT,
        target: 'word/document.xml',
        targetMode: 'Internal',
      },
    ],
    ...(partRelationships.length > 0 ? { partRelationships } : {}),
  });

  return { bytes, losses };
}

interface NoteConfig {
  readonly noteKind: 'footnote' | 'endnote';
  readonly partPath: string;
  readonly rootTag: string;
  readonly noteTag: string;
  readonly contentType: string;
  readonly relType: string;
  readonly target: string;
}

// §17.11 — emit a footnotes.xml / endnotes.xml part from the note content by id,
// prefixed with the separator / continuationSeparator stubs Word expects (the
// reader skips those on re-read). A note's blocks go through emitBlock with a
// scope flagged so its number-mark run emits w:footnoteRef / w:endnoteRef.
function emitNotes(
  notes: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined,
  cfg: NoteConfig,
  state: WriteState,
  losses: Array<Loss>,
  docScope: PartScope,
  extraParts: Array<OpcPart>,
  extraPartRels: Array<{ sourcePart: string; relationships: Array<Relationship> }>,
): void {
  if (!notes || notes.size === 0) return;
  const scope = newScope();
  scope.noteKind = cfg.noteKind;
  const stub = (type: string, id: number, mark: string): string =>
    `<${cfg.noteTag} w:type="${type}" w:id="${id}"><w:p><w:r>${mark}</w:r></w:p></${cfg.noteTag}>`;
  const noteXmls: Array<string> = [];
  for (const [id, content] of notes) {
    const inner: Array<string> = [];
    for (const el of content) emitBlock(inner, el, losses, state, scope);
    noteXmls.push(
      `<${cfg.noteTag} w:id="${escapeAttr(id)}">${inner.join('') || '<w:p/>'}</${cfg.noteTag}>`,
    );
  }
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<${cfg.rootTag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    stub('separator', -1, '<w:separator/>') +
    stub('continuationSeparator', 0, '<w:continuationSeparator/>') +
    noteXmls.join('') +
    `</${cfg.rootTag}>`;
  extraParts.push({ path: cfg.partPath, data: encoder.encode(xml), contentType: cfg.contentType });
  if (scope.rels.length > 0) {
    extraPartRels.push({ sourcePart: cfg.partPath, relationships: scope.rels });
  }
  docScope.rels.push({
    id: `rId${++docScope.relSeq}`,
    type: cfg.relType,
    target: cfg.target,
    targetMode: 'Internal',
  });
}

// A deterministic 8-hex w14:paraId for a comment id (FNV-1a). Threads link by
// these ids, so only their internal consistency matters — not the originals.
function paraIdFor(commentId: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < commentId.length; i++) {
    h ^= commentId.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0 || 1).toString(16).toUpperCase().padStart(8, '0');
}

// Stamp a w14:paraId onto a single paragraph's opening tag (paragraphXml always
// emits a bare `<w:p>`); leaves any other shape untouched.
const PARA_OPEN = '<w:p>';
function injectParaId(paragraphXmlString: string, paraId: string): string {
  return paragraphXmlString.startsWith(PARA_OPEN)
    ? `<w:p w14:paraId="${paraId}">` + paragraphXmlString.slice(PARA_OPEN.length)
    : paragraphXmlString;
}

// §17.13.4 — emit word/comments.xml from the comments by id. Unlike notes a
// comment carries author/date/initials attributes (and has no separator stubs);
// the body's commentReference runs (emitted inline) point back by id (CM3). Each
// comment's last paragraph gets a w14:paraId so commentsExtended can thread it
// (CM4); the assigned ids are returned for emitCommentsExtended.
function emitComments(
  comments: ReadonlyMap<string, Comment> | undefined,
  state: WriteState,
  losses: Array<Loss>,
  docScope: PartScope,
  extraParts: Array<OpcPart>,
  extraPartRels: Array<{ sourcePart: string; relationships: Array<Relationship> }>,
): Map<string, string> {
  const paraIds = new Map<string, string>();
  if (!comments || comments.size === 0) return paraIds;
  const scope = newScope();
  const commentXmls: Array<string> = [];
  for (const [id, c] of comments) {
    let lastParaIdx = -1;
    for (let i = 0; i < c.content.length; i++) {
      if (c.content[i]!.kind === 'paragraph') lastParaIdx = i;
    }
    const paraId = lastParaIdx >= 0 ? paraIdFor(id) : undefined;
    const inner: Array<string> = [];
    for (let i = 0; i < c.content.length; i++) {
      if (i === lastParaIdx && paraId !== undefined) {
        const buf: Array<string> = [];
        emitBlock(buf, c.content[i]!, losses, state, scope);
        inner.push(injectParaId(buf.join(''), paraId));
      } else {
        emitBlock(inner, c.content[i]!, losses, state, scope);
      }
    }
    if (paraId !== undefined) paraIds.set(id, paraId);
    const attrs =
      `w:id="${escapeAttr(id)}"` +
      (c.author !== undefined ? ` w:author="${escapeAttr(c.author)}"` : '') +
      (c.date !== undefined ? ` w:date="${escapeAttr(c.date)}"` : '') +
      (c.initials !== undefined ? ` w:initials="${escapeAttr(c.initials)}"` : '');
    commentXmls.push(`<w:comment ${attrs}>${inner.join('') || '<w:p/>'}</w:comment>`);
  }
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ` xmlns:w14="${W14_NS}">` +
    commentXmls.join('') +
    '</w:comments>';
  extraParts.push({
    path: COMMENTS_PART,
    data: encoder.encode(xml),
    contentType: COMMENTS_CONTENT_TYPE,
  });
  if (scope.rels.length > 0) {
    extraPartRels.push({ sourcePart: COMMENTS_PART, relationships: scope.rels });
  }
  docScope.rels.push({
    id: `rId${++docScope.relSeq}`,
    type: REL_COMMENTS,
    target: 'comments.xml',
    targetMode: 'Internal',
  });
  return paraIds;
}

// §commentsEx (w15) — emit word/commentsExtended.xml linking replies to parents
// (paraIdParent) and flagging resolved threads (done), keyed by the paraIds
// emitComments stamped (CM4). Emitted only when there is thread info to carry.
function emitCommentsExtended(
  comments: ReadonlyMap<string, Comment> | undefined,
  paraIds: Map<string, string>,
  docScope: PartScope,
  extraParts: Array<OpcPart>,
): void {
  if (!comments || paraIds.size === 0) return;
  const hasThreadInfo = [...comments].some(
    ([id, c]) =>
      paraIds.has(id) && ((c.parentId !== undefined && paraIds.has(c.parentId)) || c.done === true),
  );
  if (!hasThreadInfo) return;
  const rows: Array<string> = [];
  for (const [id, c] of comments) {
    const pid = paraIds.get(id);
    if (pid === undefined) continue;
    const parentPid = c.parentId !== undefined ? paraIds.get(c.parentId) : undefined;
    rows.push(
      `<w15:commentEx w15:paraId="${pid}"` +
        (parentPid !== undefined ? ` w15:paraIdParent="${parentPid}"` : '') +
        ` w15:done="${c.done === true ? '1' : '0'}"/>`,
    );
  }
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w15:commentsEx xmlns:w15="${W15_NS}">` +
    rows.join('') +
    '</w15:commentsEx>';
  extraParts.push({
    path: COMMENTS_EXTENDED_PART,
    data: encoder.encode(xml),
    contentType: COMMENTS_EXTENDED_CONTENT_TYPE,
  });
  docScope.rels.push({
    id: `rId${++docScope.relSeq}`,
    type: REL_COMMENTS_EXTENDED,
    target: 'commentsExtended.xml',
    targetMode: 'Internal',
  });
}

interface HeaderFooterRefs {
  readonly headers: Array<{ type: string; relId: string }>;
  readonly footers: Array<{ type: string; relId: string }>;
}

// Emit the header/footer parts the section references, returning the sectPr
// references. Each part is parsed back at the fixed path the reader resolves
// via the document relationship; images inside it use the part's own scope.
function emitHeadersFooters(
  flow: FlowDoc,
  section: SectionProperties | undefined,
  state: WriteState,
  docScope: PartScope,
  extraParts: Array<OpcPart>,
  extraPartRels: Array<{ sourcePart: string; relationships: Array<Relationship> }>,
  // Original relationship id → emitted document rId, so a header/footer part
  // shared by several sections is emitted once.
  hfCache: Map<string, string>,
  losses: Array<Loss>,
): HeaderFooterRefs {
  const refs: HeaderFooterRefs = { headers: [], footers: [] };
  if (!section || !flow.headersFooters) return refs;

  const emitOne = (
    kind: 'header' | 'footer',
    relationshipId: string,
    type: string,
  ): string | undefined => {
    const cached = hfCache.get(relationshipId);
    if (cached !== undefined) return cached;
    const content = flow.headersFooters!.get(relationshipId);
    if (!content) return undefined;
    const n = extraParts.filter((p) => p.path.includes(`/${kind}`)).length + 1;
    const path = `word/${kind}${n}.xml`;
    const root = kind === 'header' ? 'w:hdr' : 'w:ftr';
    const scope = newScope();
    const inner: Array<string> = [];
    for (const el of content) emitBlock(inner, el, losses, state, scope);
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `${inner.join('')}</${root}>`;
    extraParts.push({
      path,
      data: encoder.encode(xml),
      contentType: kind === 'header' ? HEADER_CONTENT_TYPE : FOOTER_CONTENT_TYPE,
    });
    if (scope.rels.length > 0) extraPartRels.push({ sourcePart: path, relationships: scope.rels });
    const relId = `rId${++docScope.relSeq}`;
    docScope.rels.push({
      id: relId,
      type: kind === 'header' ? REL_HEADER : REL_FOOTER,
      target: `${kind}${n}.xml`,
      targetMode: 'Internal',
    });
    hfCache.set(relationshipId, relId);
    return relId;
  };

  for (const h of section.headers) {
    const relId = emitOne('header', h.relationshipId, h.type);
    if (relId) refs.headers.push({ type: h.type, relId });
  }
  for (const f of section.footers) {
    const relId = emitOne('footer', f.relationshipId, f.type);
    if (relId) refs.footers.push({ type: f.type, relId });
  }
  return refs;
}

export const docxWriter: DocumentWriter<FlowDoc> = {
  id: 'docx',
  consumes: 'flow',
  supports: new Set([FEATURES.text]),
  write: (doc) => writeDocx(doc),
};

function emitBlock(
  out: Array<string>,
  el: BodyElement,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
  // §17.6.17 — a mid-document section's sectPr to attach as this block's
  // section break (it closes the section at this body element).
  closingSectPr?: string,
): void {
  if (el.kind === 'paragraph') {
    out.push(paragraphXml(el.paragraph, state, scope, closingSectPr));
    return;
  }
  if (el.kind === 'table') {
    out.push(tableXml(el.table, losses, state, scope));
    if (closingSectPr) out.push(`<w:p><w:pPr>${closingSectPr}</w:pPr></w:p>`);
    return;
  }
  if (el.kind === 'image') {
    const drawing = drawingXml(
      el.image.resource,
      el.image.width,
      el.image.height,
      el.image.altText,
      state,
      scope,
    );
    if (drawing) {
      // An image is emitted as a paragraph, so a closing section break rides
      // its pPr (like a text paragraph) rather than a separate carrier — which
      // would otherwise re-read as an extra empty paragraph.
      out.push(
        `<w:p>${pPrWithSect(el.image.paragraphProperties, closingSectPr)}<w:r>${drawing}</w:r></w:p>`,
      );
    } else {
      losses.push({ severity: 'dropped', feature: FEATURES.images, detail: 'image bytes missing' });
      if (closingSectPr) out.push(`<w:p><w:pPr>${closingSectPr}</w:pPr></w:p>`);
    }
    return;
  }
  if (el.kind === 'shape') {
    // §20.4 — a DrawingML shape as its own paragraph (the re-read collapses it
    // back to a ShapeBlock); a closing section break rides its pPr, no carrier.
    const drawing = shapeDrawingXml(el.shape, losses, state, scope);
    out.push(
      `<w:p>${pPrWithSect(el.shape.paragraphProperties, closingSectPr)}<w:r>${drawing}</w:r></w:p>`,
    );
    return;
  }
  // §21.2 — a chart block (WT3), emitted as its own paragraph like an image.
  const chartDrawing = chartBlockXml(el.chart, state, scope, losses);
  if (chartDrawing) {
    out.push(
      `<w:p>${pPrWithSect(el.chart.paragraphProperties, closingSectPr)}<w:r>${chartDrawing}</w:r></w:p>`,
    );
  } else if (closingSectPr) {
    out.push(`<w:p><w:pPr>${closingSectPr}</w:pPr></w:p>`);
  }
}

// §21.2 — a chart block as an inline w:drawing referencing a serialized chart
// part (the shared chart-serializer, also used by the xlsx writer). Returns ''
// when the chart data is missing (the caller already drops the block).
function chartBlockXml(
  chart: ChartBlock,
  state: WriteState,
  scope: PartScope,
  losses: Array<Loss>,
): string {
  const data = state.charts?.get(chart.chartRelId);
  if (!data) {
    losses.push({ severity: 'dropped', feature: FEATURES.charts, detail: 'chart data missing' });
    return '';
  }
  const cid = ++state.chartSeq;
  state.chartParts.push({
    path: `word/charts/chart${cid}.xml`,
    data: encoder.encode(chartSpaceXml(data)),
    contentType: CHART_CONTENT_TYPE,
  });
  const relId = `rId${++scope.relSeq}`;
  scope.rels.push({
    id: relId,
    type: REL_CHART,
    target: `charts/chart${cid}.xml`,
    targetMode: 'Internal',
  });
  const cx = Math.round(chart.width * EMU_PER_PT);
  const cy = Math.round(chart.height * EMU_PER_PT);
  const id = ++state.drawingSeq;
  const descr = chart.altText ? ` descr="${escapeAttr(chart.altText)}"` : '';
  return (
    '<w:drawing>' +
    '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="Chart ${id}"${descr}/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing>'
  );
}

// Allocate (or reuse) a media part + image relationship for a resource, then
// emit the w:drawing/wp:inline markup the reader round-trips. Returns '' when
// the resource has no bytes (an unresolved image — the caller drops it).
function drawingXml(
  resource: ResourceId | undefined,
  widthPt: number,
  heightPt: number,
  altText: string | undefined,
  state: WriteState,
  scope: PartScope,
): string {
  if (resource === undefined) return '';
  const relId = mediaRelId(resource, state, scope);
  if (relId === undefined) return '';
  const cx = Math.round(widthPt * EMU_PER_PT);
  const cy = Math.round(heightPt * EMU_PER_PT);
  const id = ++state.drawingSeq;
  const descr = altText ? ` descr="${escapeAttr(altText)}"` : '';
  return (
    '<w:drawing>' +
    '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="Image ${id}"${descr}/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Image ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/>' +
    `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
  );
}

// §20.4 wp:inline holding a wps:wsp — the inverse of drawing-parser's parseWsp.
// The reader collapses a lone shape paragraph to a ShapeBlock, so emitting the
// shape (rather than dropping it and leaving an empty carrier paragraph) keeps
// the round-trip's block structure stable. Floating placement (wp:anchor) is
// not re-emitted yet — a shape round-trips as inline.
const WPS_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

function shapeDrawingXml(
  shape: ShapeBlock,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
): string {
  const cx = Math.round(shape.width * EMU_PER_PT);
  const cy = Math.round(shape.height * EMU_PER_PT);
  const id = ++state.drawingSeq;
  const descr = shape.altText ? ` descr="${escapeAttr(shape.altText)}"` : '';
  const spPr =
    `<wps:spPr>${xfrmXml(shape.transform, cx, cy)}${geomXml(shape.geometry)}` +
    `${fillXml(shape.fill)}${shape.line ? lineXml(shape.line) : ''}</wps:spPr>`;
  const txbx = shape.text ? txbxXml(shape.text, losses, state, scope) : '';
  return (
    '<w:drawing>' +
    '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="Shape ${id}"${descr}/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    `<a:graphicData uri="${WPS_URI}">` +
    '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:cNvSpPr/>' +
    spPr +
    txbx +
    bodyPrXml(shape.text) +
    '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>'
  );
}

// §20.1.7.6 a:xfrm — rotation/flips plus the off+ext the reader uses as a size
// fallback. Always emitted so a re-read recovers the box even without extent.
function xfrmXml(t: ShapeTransform | undefined, cx: number, cy: number): string {
  const rot = t?.rotation60k !== undefined ? ` rot="${t.rotation60k}"` : '';
  const flipH = t?.flipH ? ' flipH="1"' : '';
  const flipV = t?.flipV ? ' flipV="1"' : '';
  return `<a:xfrm${rot}${flipH}${flipV}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
}

function geomXml(g: ShapeGeometry): string {
  if (g.kind === 'custom' && g.custom) {
    const cmds = g.custom.commands
      .map((c) => {
        switch (c.cmd) {
          case 'move':
            return `<a:moveTo><a:pt x="${c.x}" y="${c.y}"/></a:moveTo>`;
          case 'line':
            return `<a:lnTo><a:pt x="${c.x}" y="${c.y}"/></a:lnTo>`;
          case 'cubic':
            return (
              '<a:cubicBezTo>' +
              `<a:pt x="${c.x1}" y="${c.y1}"/><a:pt x="${c.x2}" y="${c.y2}"/><a:pt x="${c.x}" y="${c.y}"/>` +
              '</a:cubicBezTo>'
            );
          case 'quad':
            return `<a:quadBezTo><a:pt x="${c.x1}" y="${c.y1}"/><a:pt x="${c.x}" y="${c.y}"/></a:quadBezTo>`;
          case 'arc':
            return `<a:arcTo wR="${c.wR}" hR="${c.hR}" stAng="${c.stAng}" swAng="${c.swAng}"/>`;
          case 'close':
            return '<a:close/>';
        }
      })
      .join('');
    const { pathWidth: w, pathHeight: h } = g.custom;
    return (
      '<a:custGeom><a:avLst/><a:gdLst/>' +
      `<a:rect l="0" t="0" r="${w}" b="${h}"/>` +
      `<a:pathLst><a:path w="${w}" h="${h}">${cmds}</a:path></a:pathLst></a:custGeom>`
    );
  }
  const gds =
    g.adjust && g.adjust.size > 0
      ? [...g.adjust].map(([n, v]) => `<a:gd name="${escapeAttr(n)}" fmla="val ${v}"/>`).join('')
      : '';
  return `<a:prstGeom prst="${escapeAttr(g.preset ?? 'rect')}"><a:avLst>${gds}</a:avLst></a:prstGeom>`;
}

function fillXml(f: ShapeFill): string {
  if (f.kind === 'solid' && f.colorHex)
    return `<a:solidFill><a:srgbClr val="${f.colorHex}"/></a:solidFill>`;
  if (f.kind === 'gradient' && f.gradient) return gradFillXml(f.gradient);
  return '<a:noFill/>';
}

// A gradient fill → a:gradFill (EP16): stops as a:gs (@pos in 1000ths of a
// percent), direction as a:lin (@ang in 60000ths of a degree) or a:path (radial).
function gradFillXml(g: ShapeGradient): string {
  const stops = g.stops
    .map((s) => {
      const pos = Math.round(Math.max(0, Math.min(1, s.offset)) * 100000);
      return `<a:gs pos="${pos}"><a:srgbClr val="${s.colorHex}"/></a:gs>`;
    })
    .join('');
  const dir =
    g.kind === 'radial'
      ? '<a:path path="circle"/>'
      : `<a:lin ang="${Math.round(((((g.angle ?? 0) % 360) + 360) % 360) * 60000)}" scaled="1"/>`;
  return `<a:gradFill><a:gsLst>${stops}</a:gsLst>${dir}</a:gradFill>`;
}

function lineXml(l: ShapeLine): string {
  const w = l.width !== undefined ? ` w="${Math.round(l.width * EMU_PER_PT)}"` : '';
  const cap =
    l.cap === 'round'
      ? ' cap="rnd"'
      : l.cap === 'square'
        ? ' cap="sq"'
        : l.cap === 'flat'
          ? ' cap="flat"'
          : '';
  const inner: Array<string> = [];
  if (l.fill === 'none') inner.push('<a:noFill/>');
  else if (l.colorHex) inner.push(`<a:solidFill><a:srgbClr val="${l.colorHex}"/></a:solidFill>`);
  if (l.dash) inner.push(`<a:prstDash val="${l.dash}"/>`);
  return `<a:ln${w}${cap}>${inner.join('')}</a:ln>`;
}

function txbxXml(
  text: ShapeTextBody,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
): string {
  const inner: Array<string> = [];
  for (const el of text.content) emitBlock(inner, el, losses, state, scope);
  return `<wps:txbx><w:txbxContent>${inner.join('')}</w:txbxContent></wps:txbx>`;
}

function bodyPrXml(text: ShapeTextBody | undefined): string {
  if (!text) return '<wps:bodyPr/>';
  const ins = (v: number | undefined, name: string): string =>
    v !== undefined ? ` ${name}="${Math.round(v * EMU_PER_PT)}"` : '';
  const anchor = text.anchor ? ` anchor="${text.anchor}"` : '';
  return (
    '<wps:bodyPr' +
    ins(text.insetLeft, 'lIns') +
    ins(text.insetTop, 'tIns') +
    ins(text.insetRight, 'rIns') +
    ins(text.insetBottom, 'bIns') +
    anchor +
    '/>'
  );
}

// The media FILE is content-addressed and shared across parts; the rId is
// allocated within the CURRENT part's scope (OPC §9.3).
function mediaRelId(resource: ResourceId, state: WriteState, scope: PartScope): string | undefined {
  // Per-part rId reuse: same resource referenced twice in one part → one rId.
  const existingRel = scope.relIdByResource.get(resource);
  if (existingRel !== undefined) return existingRel;

  // Shared media file (content-addressed): create once per distinct resource.
  let target = state.mediaFileByResource.get(resource);
  if (target === undefined) {
    const bytes = state.resources.get(resource);
    if (!bytes) return undefined;
    const info = mediaInfo(bytes);
    if (!info) return undefined;
    const n = state.mediaParts.length + 1;
    target = `media/image${n}.${info.ext}`;
    state.mediaParts.push({ path: `word/${target}`, data: bytes, contentType: info.contentType });
    state.mediaFileByResource.set(resource, target);
  }

  const relId = `rId${++scope.relSeq}`;
  scope.rels.push({ id: relId, type: REL_IMAGE, target, targetMode: 'Internal' });
  scope.relIdByResource.set(resource, relId);
  return relId;
}

// §17.4 — w:tbl: properties, the column grid, then rows. Cell content recurses
// through emitBlock, so nested tables and per-cell paragraphs round-trip.
function tableXml(table: Table, losses: Array<Loss>, state: WriteState, scope: PartScope): string {
  const grid = table.grid.map((w) => `<w:gridCol w:w="${twips(w)}"/>`).join('');
  const rows = table.rows.map((row) => rowXml(row, losses, state, scope)).join('');
  return `<w:tbl>${tblPrXml(table.properties)}<w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

function tblPrXml(p: TableProperties): string {
  const out: Array<string> = [];
  if (p.widthType !== undefined) {
    const w =
      p.widthType === 'pct'
        ? Math.round((p.widthFraction ?? 0) * 5000)
        : p.widthPt !== undefined
          ? twips(p.widthPt)
          : 0;
    out.push(`<w:tblW w:w="${w}" w:type="${p.widthType}"/>`);
  }
  if (p.alignment && p.alignment !== 'left') out.push(`<w:jc w:val="${p.alignment}"/>`);
  const borders = bordersXml('w:tblBorders', p.borders);
  if (borders) out.push(borders);
  const margins = cellMarginsXml('w:tblCellMar', p.defaultCellMargins);
  if (margins) out.push(margins);
  return out.length > 0 ? `<w:tblPr>${out.join('')}</w:tblPr>` : '';
}

function rowXml(row: TableRow, losses: Array<Loss>, state: WriteState, scope: PartScope): string {
  const trPr: Array<string> = [];
  if (row.properties.height !== undefined) {
    const rule = row.properties.heightRule ?? 'atLeast';
    trPr.push(`<w:trHeight w:val="${twips(row.properties.height)}" w:hRule="${rule}"/>`);
  }
  if (row.properties.isHeader) trPr.push('<w:tblHeader/>');
  if (row.properties.cantSplit) trPr.push('<w:cantSplit/>');
  const trPrXml = trPr.length > 0 ? `<w:trPr>${trPr.join('')}</w:trPr>` : '';
  const cells = row.cells.map((cell) => cellXml(cell, losses, state, scope)).join('');
  return `<w:tr>${trPrXml}${cells}</w:tr>`;
}

function cellXml(
  cell: TableCell,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
): string {
  const content: Array<string> = [];
  for (const child of cell.content) emitBlock(content, child, losses, state, scope);
  // §17.4.66 — a w:tc must contain at least one block, ending in a paragraph.
  if (content.length === 0) content.push('<w:p/>');
  return `<w:tc>${tcPrXml(cell.properties)}${content.join('')}</w:tc>`;
}

function tcPrXml(p: CellProperties): string {
  const out: Array<string> = [];
  if (p.width !== undefined) out.push(`<w:tcW w:w="${twips(p.width)}" w:type="dxa"/>`);
  if (p.colSpan !== undefined && p.colSpan > 1) {
    out.push(`<w:gridSpan w:val="${p.colSpan}"/>`);
  }
  if (p.merge !== undefined) {
    // §17.4.85 — 'start' restarts a vertical merge; 'middle'/'end' continue it.
    out.push(p.merge === 'start' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge w:val="continue"/>');
  }
  const borders = bordersXml('w:tcBorders', p.borders);
  if (borders) out.push(borders);
  const margins = cellMarginsXml('w:tcMar', p.margins);
  if (margins) out.push(margins);
  if (p.shading) out.push(`<w:shd w:val="clear" w:color="auto" w:fill="${p.shading.colorHex}"/>`);
  return out.length > 0 ? `<w:tcPr>${out.join('')}</w:tcPr>` : '';
}

const BORDER_SIDES: ReadonlyArray<[keyof CellBorders, string]> = [
  ['top', 'w:top'],
  ['left', 'w:left'],
  ['bottom', 'w:bottom'],
  ['right', 'w:right'],
  ['insideH', 'w:insideH'],
  ['insideV', 'w:insideV'],
];

function bordersXml(tag: 'w:tblBorders' | 'w:tcBorders', borders: CellBorders | undefined): string {
  if (!borders) return '';
  const sides = BORDER_SIDES.map(([key, el]) => {
    const b = borders[key];
    if (!b) return '';
    // §17.4.x — w:sz in eighths of a point; the reader divides by 8.
    const sz = b.width !== undefined ? ` w:sz="${Math.round(b.width * 8)}"` : '';
    const color = b.colorHex !== undefined ? ` w:color="${b.colorHex}"` : '';
    return `<${el} w:val="${b.style}"${sz}${color}/>`;
  }).join('');
  return sides ? `<${tag}>${sides}</${tag}>` : '';
}

function cellMarginsXml(tag: 'w:tblCellMar' | 'w:tcMar', margins: CellMargins | undefined): string {
  if (!margins) return '';
  const sides: Array<[keyof CellMargins, string]> = [
    ['top', 'w:top'],
    ['left', 'w:left'],
    ['bottom', 'w:bottom'],
    ['right', 'w:right'],
  ];
  const inner = sides
    .map(([key, el]) => {
      const v = margins[key];
      return v !== undefined ? `<${el} w:w="${twips(v)}" w:type="dxa"/>` : '';
    })
    .join('');
  return inner ? `<${tag}>${inner}</${tag}>` : '';
}

function paragraphXml(
  p: Paragraph,
  state: WriteState,
  scope: PartScope,
  closingSectPr?: string,
): string {
  // The runs the reader actually kept: list markers re-materialize from
  // numbering.xml, math runs are not written yet, and a run is visible if it
  // has text or an inline image.
  const visible = p.runs.filter(
    (run) =>
      !run.listMarker &&
      (run.math !== undefined ||
        run.text !== '' ||
        run.inlineImage !== undefined ||
        run.pageBreak ||
        // §17.11 — a note reference / in-note number mark (WT2).
        run.footnoteRef !== undefined ||
        run.endnoteRef !== undefined ||
        run.noteNumber === true ||
        // §17.13.4.1 — a comment reference (CM3): an empty run that anchors a
        // comment must survive the round-trip.
        run.commentRef !== undefined ||
        // An empty run that still carries a link target keeps the hyperlink
        // alive (e.g. a TOC field whose page number a tracked change deleted).
        run.href !== undefined ||
        run.anchor !== undefined),
  );

  // §17.16.22 — group adjacent runs sharing a hyperlink target back into one
  // w:hyperlink container (the reader stamped href/anchor onto every run
  // inside it; this is the inverse).
  const inner: Array<string> = [];
  let i = 0;
  while (i < visible.length) {
    const run = visible[i]!;
    const key = run.href ?? run.anchor;
    if (key === undefined) {
      inner.push(runXml(run, state, scope));
      i++;
      continue;
    }
    let j = i + 1;
    while (j < visible.length && (visible[j]!.href ?? visible[j]!.anchor) === key) j++;
    const group = visible
      .slice(i, j)
      .map((r) => runXml(r, state, scope))
      .join('');
    // A hyperlink whose runs are all empty still carries its target — give it a
    // single empty run so the link survives the round-trip rather than emitting
    // an empty <w:hyperlink/> that a re-read would discard.
    inner.push(hyperlinkXml(run, group === '' ? '<w:r/>' : group, state, scope));
    i = j;
  }

  // §17.13.6.2 bookmarks: opened at the paragraph (start + end with a unique
  // id each); the reader reads the start, the end keeps the markup valid.
  const bookmarks = (p.bookmarks ?? [])
    .map((name) => {
      const id = state.bookmarkSeq++;
      return `<w:bookmarkStart w:id="${id}" w:name="${escapeAttr(name)}"/><w:bookmarkEnd w:id="${id}"/>`;
    })
    .join('');

  // §17.6.17 — a mid-document section break: its sectPr is appended inside the
  // pPr of this (the section's last) paragraph.
  const pPrInner = pPrBody(p.properties as ResolvedParagraphProperties) + (closingSectPr ?? '');
  const pPr = pPrInner !== '' ? `<w:pPr>${pPrInner}</w:pPr>` : '';
  return `<w:p>${pPr}${bookmarks}${inner.join('')}</w:p>`;
}

// w:hyperlink container: @r:id for an external target (allocates a rel),
// @w:anchor for an internal bookmark reference.
function hyperlinkXml(run: Run, inner: string, _state: WriteState, scope: PartScope): string {
  if (run.href !== undefined) {
    const id = `rId${++scope.relSeq}`;
    scope.rels.push({ id, type: REL_HYPERLINK, target: run.href, targetMode: 'External' });
    return `<w:hyperlink r:id="${id}">${inner}</w:hyperlink>`;
  }
  return `<w:hyperlink w:anchor="${escapeAttr(run.anchor!)}">${inner}</w:hyperlink>`;
}

function runXml(run: Run, state: WriteState, scope: PartScope): string {
  // §22 — a math run is an <m:oMath> (the m: namespace declared here), not a
  // w:r; its run properties do not apply (WT3).
  if (run.math !== undefined) {
    return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${omathXml(run.math)}</m:oMath>`;
  }
  const rPr = rPrXml(run.properties as ResolvedRunProperties);
  if (run.inlineImage !== undefined) {
    const img = run.inlineImage;
    const drawing = drawingXml(img.resource, img.width, img.height, undefined, state, scope);
    if (drawing) return `<w:r>${rPr}${drawing}</w:r>`;
    // Unresolved inline image with no text and no break: nothing to emit.
    if (run.text === '' && !run.pageBreak) return '';
  }
  // §17.11 — note references and the in-note number mark (WT2).
  if (run.footnoteRef !== undefined) {
    return `<w:r>${rPr}<w:footnoteReference w:id="${escapeAttr(run.footnoteRef)}"/></w:r>`;
  }
  if (run.endnoteRef !== undefined) {
    return `<w:r>${rPr}<w:endnoteReference w:id="${escapeAttr(run.endnoteRef)}"/></w:r>`;
  }
  if (run.noteNumber) {
    return `<w:r>${rPr}<${scope.noteKind === 'endnote' ? 'w:endnoteRef' : 'w:footnoteRef'}/></w:r>`;
  }
  // §17.13.4.1 — a review comment reference (CM3).
  if (run.commentRef !== undefined) {
    return `<w:r>${rPr}<w:commentReference w:id="${escapeAttr(run.commentRef)}"/></w:r>`;
  }
  // §17.3.3.1 — a page break is a run-level <w:br w:type="page"/>; emit it so a
  // run that is ONLY a break (no text, no image) survives the round-trip.
  const brk = run.pageBreak ? '<w:br w:type="page"/>' : '';
  if (run.text === '') return brk ? `<w:r>${rPr}${brk}</w:r>` : '';
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(run.text)}</w:t>${brk}</w:r>`;
}

// §17.3.2 — run properties as a delta from the resolved defaults.
function rPrXml(r: ResolvedRunProperties): string {
  const out: Array<string> = [];
  if (r.bold !== DEFAULT_RUN.bold) out.push(toggle('w:b', r.bold));
  if (r.italic !== DEFAULT_RUN.italic) out.push(toggle('w:i', r.italic));
  if (r.strike !== DEFAULT_RUN.strike) out.push(toggle('w:strike', r.strike));
  if (r.underline !== DEFAULT_RUN.underline) out.push(`<w:u w:val="${r.underline}"/>`);
  const fonts = rFontsXml(r.fontFamily);
  if (fonts) out.push(fonts);
  if (r.fontSizePt !== DEFAULT_RUN.fontSizePt) {
    // §17.3.2.38 w:sz — half-points.
    out.push(`<w:sz w:val="${Math.round(r.fontSizePt * 2)}"/>`);
  }
  if (r.colorHex !== DEFAULT_RUN.colorHex) out.push(`<w:color w:val="${r.colorHex}"/>`);
  if (r.verticalAlign !== DEFAULT_RUN.verticalAlign) {
    out.push(`<w:vertAlign w:val="${r.verticalAlign}"/>`);
  }
  if (r.rtl !== DEFAULT_RUN.rtl) out.push(toggle('w:rtl', r.rtl));
  if (r.lang !== undefined) out.push(`<w:lang w:val="${escapeAttr(r.lang)}"/>`);
  return out.length > 0 ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

// §17.3.1 — paragraph properties as a delta from the resolved defaults, the
// INNER content of w:pPr (no wrapper, so a section break can be appended).
function pPrBody(p: ResolvedParagraphProperties): string {
  const out: Array<string> = [];
  if (p.numbering) {
    // §17.3.1.19 — list membership; the marker itself comes from numbering.xml.
    out.push(
      `<w:numPr><w:ilvl w:val="${p.numbering.ilvl}"/><w:numId w:val="${escapeAttr(p.numbering.numId)}"/></w:numPr>`,
    );
  }
  if (p.outlineLevel !== undefined) out.push(`<w:outlineLvl w:val="${p.outlineLevel}"/>`);
  if (p.pageBreakBefore) out.push('<w:pageBreakBefore/>');
  if (p.bidi !== DEFAULT_PARA.bidi) out.push(toggle('w:bidi', p.bidi));
  const ind = indXml(p);
  if (ind) out.push(ind);
  const spacing = spacingXml(p);
  if (spacing) out.push(spacing);
  if (p.alignment !== DEFAULT_PARA.alignment) out.push(`<w:jc w:val="${p.alignment}"/>`);
  return out.join('');
}

function pPrXml(p: ResolvedParagraphProperties): string {
  const inner = pPrBody(p);
  return inner !== '' ? `<w:pPr>${inner}</w:pPr>` : '';
}

// A pPr with an optional mid-document section break (§17.6.17) appended — for
// block elements that emit as a paragraph (image/shape), so a closing sectPr
// rides their own pPr instead of a separate carrier paragraph.
function pPrWithSect(p: ParagraphProperties, closingSectPr?: string): string {
  const inner = pPrBody(p as ResolvedParagraphProperties) + (closingSectPr ?? '');
  return inner !== '' ? `<w:pPr>${inner}</w:pPr>` : '';
}

function indXml(p: ResolvedParagraphProperties): string {
  const attrs: Array<string> = [];
  if (p.indentLeft !== DEFAULT_PARA.indentLeft) attrs.push(`w:left="${twips(p.indentLeft)}"`);
  if (p.indentRight !== DEFAULT_PARA.indentRight) attrs.push(`w:right="${twips(p.indentRight)}"`);
  if (p.indentFirstLine !== DEFAULT_PARA.indentFirstLine) {
    // A negative first-line indent is a hanging indent (§17.3.1.12).
    if (p.indentFirstLine < 0) attrs.push(`w:hanging="${twips(-p.indentFirstLine)}"`);
    else attrs.push(`w:firstLine="${twips(p.indentFirstLine)}"`);
  }
  return attrs.length > 0 ? `<w:ind ${attrs.join(' ')}/>` : '';
}

function spacingXml(p: ResolvedParagraphProperties): string {
  const attrs: Array<string> = [];
  if (p.spacingBefore !== DEFAULT_PARA.spacingBefore) {
    attrs.push(`w:before="${twips(p.spacingBefore)}"`);
  }
  if (p.spacingAfter !== DEFAULT_PARA.spacingAfter) {
    attrs.push(`w:after="${twips(p.spacingAfter)}"`);
  }
  if (
    (p.spacingLineRule !== DEFAULT_PARA.spacingLineRule ||
      p.spacingLine !== DEFAULT_PARA.spacingLine) &&
    p.spacingLine > 0
  ) {
    // §17.3.1.33: 'auto' line spacing is in 240ths (line units); exact/atLeast
    // in twips. The reader stores spacingLine in points either way.
    const lineVal =
      p.spacingLineRule === 'auto' ? Math.round(p.spacingLine * 12) : twips(p.spacingLine);
    attrs.push(`w:line="${lineVal}"`, `w:lineRule="${p.spacingLineRule}"`);
  }
  return attrs.length > 0 ? `<w:spacing ${attrs.join(' ')}/>` : '';
}

// §17.3.2.26 w:rFonts — only the slots that differ from the resolved default.
function rFontsXml(fonts: FontFamilyMap): string {
  const d = DEFAULT_RUN.fontFamily;
  const attrs: Array<string> = [];
  if (fonts.ascii && fonts.ascii !== d.ascii) attrs.push(`w:ascii="${escapeAttr(fonts.ascii)}"`);
  if (fonts.hAnsi && fonts.hAnsi !== d.hAnsi) attrs.push(`w:hAnsi="${escapeAttr(fonts.hAnsi)}"`);
  if (fonts.cs && fonts.cs !== d.cs) attrs.push(`w:cs="${escapeAttr(fonts.cs)}"`);
  return attrs.length > 0 ? `<w:rFonts ${attrs.join(' ')}/>` : '';
}

// A boolean toggle property (§17.3.2.x): present-true is bare; present-false is
// w:val="false" (overrides an inherited true — exact on re-read here).
function toggle(tag: string, on: boolean): string {
  return on ? `<${tag}/>` : `<${tag} w:val="false"/>`;
}

// §17.9.1 numbering.xml: every abstractNum (levels with start/numFmt/lvlText
// and the level's raw pPr/rPr), then the num instances binding numId →
// abstractNumId. Re-emitted from the FlowDoc's raw `numbering` round-trip
// material, so re-read regenerates identical markers.
function numberingXml(numbering: Numbering): string {
  const out: Array<string> = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
  ];
  for (const abstractNum of numbering.abstractNums.values()) {
    out.push(`<w:abstractNum w:abstractNumId="${escapeAttr(abstractNum.id)}">`);
    for (const level of [...abstractNum.levels.values()].sort((a, b) => a.ilvl - b.ilvl)) {
      out.push(levelXml(level));
    }
    out.push('</w:abstractNum>');
  }
  for (const inst of numbering.numInstances.values()) {
    out.push(
      `<w:num w:numId="${escapeAttr(inst.numId)}">` +
        `<w:abstractNumId w:val="${escapeAttr(inst.abstractNumId)}"/></w:num>`,
    );
  }
  out.push('</w:numbering>');
  return out.join('');
}

function levelXml(level: NumberingLevel): string {
  const inner: Array<string> = [
    `<w:start w:val="${level.start}"/>`,
    `<w:numFmt w:val="${level.format}"/>`,
    `<w:lvlText w:val="${escapeAttr(level.lvlText)}"/>`,
  ];
  const pPr = rawParaPrXml(level.paragraphProperties);
  if (pPr) inner.push(pPr);
  const rPr = rawRunPrXml(level.runProperties);
  if (rPr) inner.push(rPr);
  return `<w:lvl w:ilvl="${level.ilvl}">${inner.join('')}</w:lvl>`;
}

// A numbering level's RAW (sparse) paragraph props — present fields only,
// no delta-against-defaults (unlike the resolved-body serializer above).
function rawParaPrXml(p: ParagraphProperties): string {
  const attrs: Array<string> = [];
  if (p.indentLeft !== undefined) attrs.push(`w:left="${twips(p.indentLeft)}"`);
  if (p.indentRight !== undefined) attrs.push(`w:right="${twips(p.indentRight)}"`);
  if (p.indentFirstLine !== undefined) {
    if (p.indentFirstLine < 0) attrs.push(`w:hanging="${twips(-p.indentFirstLine)}"`);
    else attrs.push(`w:firstLine="${twips(p.indentFirstLine)}"`);
  }
  const ind = attrs.length > 0 ? `<w:ind ${attrs.join(' ')}/>` : '';
  const jc = p.alignment !== undefined ? `<w:jc w:val="${p.alignment}"/>` : '';
  return ind || jc ? `<w:pPr>${jc}${ind}</w:pPr>` : '';
}

function rawRunPrXml(r: RunProperties): string {
  const out: Array<string> = [];
  if (r.bold !== undefined) out.push(toggle('w:b', r.bold));
  if (r.italic !== undefined) out.push(toggle('w:i', r.italic));
  if (r.strike !== undefined) out.push(toggle('w:strike', r.strike));
  if (r.underline !== undefined) out.push(`<w:u w:val="${r.underline}"/>`);
  const fonts = r.fontFamily ? rawRFontsXml(r.fontFamily) : '';
  if (fonts) out.push(fonts);
  if (r.fontSizePt !== undefined) out.push(`<w:sz w:val="${Math.round(r.fontSizePt * 2)}"/>`);
  if (r.colorHex !== undefined) out.push(`<w:color w:val="${r.colorHex}"/>`);
  if (r.verticalAlign !== undefined) out.push(`<w:vertAlign w:val="${r.verticalAlign}"/>`);
  return out.length > 0 ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

function rawRFontsXml(fonts: FontFamilyMap): string {
  const attrs: Array<string> = [];
  if (fonts.ascii) attrs.push(`w:ascii="${escapeAttr(fonts.ascii)}"`);
  if (fonts.hAnsi) attrs.push(`w:hAnsi="${escapeAttr(fonts.hAnsi)}"`);
  if (fonts.cs) attrs.push(`w:cs="${escapeAttr(fonts.cs)}"`);
  return attrs.length > 0 ? `<w:rFonts ${attrs.join(' ')}/>` : '';
}

// §17.6.17 — the section. Header/footer references first (Word's child order),
// then page size/margins, columns and the titlePg toggle.
function sectPrXml(s: SectionProperties, hf: HeaderFooterRefs): string {
  const parts: Array<string> = [];
  for (const h of hf.headers) {
    parts.push(`<w:headerReference w:type="${h.type}" r:id="${h.relId}"/>`);
  }
  for (const f of hf.footers) {
    parts.push(`<w:footerReference w:type="${f.type}" r:id="${f.relId}"/>`);
  }
  if (s.pageSize) {
    const orient = s.pageSize.orientation === 'landscape' ? ' w:orient="landscape"' : '';
    parts.push(
      `<w:pgSz w:w="${twips(s.pageSize.width)}" w:h="${twips(s.pageSize.height)}"${orient}/>`,
    );
  }
  if (s.margins) {
    const m = s.margins;
    const header = m.header !== undefined ? ` w:header="${twips(m.header)}"` : '';
    const footer = m.footer !== undefined ? ` w:footer="${twips(m.footer)}"` : '';
    parts.push(
      `<w:pgMar w:top="${twips(m.top)}" w:right="${twips(m.right)}"` +
        ` w:bottom="${twips(m.bottom)}" w:left="${twips(m.left)}"${header}${footer}/>`,
    );
  }
  if (s.columns) parts.push(colsXml(s.columns));
  if (s.titlePg) parts.push('<w:titlePg/>');
  if (parts.length === 0) return '';
  return `<w:sectPr>${parts.join('')}</w:sectPr>`;
}

// §17.6.4 w:cols: explicit per-column widths when present, else N equal columns
// with the shared gutter.
function colsXml(cols: SectionColumns): string {
  if (cols.explicit && cols.explicit.length > 0) {
    const inner = cols.explicit
      .map((c) => `<w:col w:w="${twips(c.widthPt)}" w:space="${twips(c.spacePt)}"/>`)
      .join('');
    return `<w:cols w:num="${cols.explicit.length}" w:equalWidth="0">${inner}</w:cols>`;
  }
  return `<w:cols w:num="${cols.count}" w:space="${twips(cols.spacePt)}"/>`;
}

const twips = (pt: number): number => Math.round(pt * 20);

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(s: string): string {
  return escapeXml(s).replace(/"/g, '&quot;');
}
