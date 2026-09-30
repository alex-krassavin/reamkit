// docx writer (E-DOCX): FlowDoc → WordprocessingML package — the inverse of
// the docx reader, and the fifth adapter overall. A flow medium with zero
// layout and zero I/O, like the HTML writer.
//
// v1 contract (epics.md, variant A): the writer emits a DENORMALIZED but
// valid document. FlowDoc's body carries RESOLVED properties (the stage-6
// cascade is already collapsed), so what we write is direct formatting — no
// named styles, only the defaults a property stated nowhere takes (see
// `stylesXml`). The round-trip guarantee is therefore semantic, not textual:
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
  FloatAnchor,
  FontFamilyMap,
  ImageBlock,
  ImageCrop,
  Numbering,
  NumberingLevel,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  Section,
  SectionColumns,
  SectionProperties,
  ShapeBlock,
  ShapeFill,
  ShapeGeometry,
  ShapeGroupChild,
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
import type { FaceFamily, FaceOutlines, FlowDoc } from '@/core/ir/flow';
import type { Loss, ResourceId, ResourceStore } from '@/core/ir';
import type { OpcPart, Relationship } from '@/core/opc';

import { FEATURES } from '@/core/ir';
import { chartSpaceXml } from '@/core/drawingml/chart-serializer';
import { detectImageFormat } from '@/core/images';
import { buildOpcPackage } from '@/core/opc';
import { OBFUSCATED_FONT_CONTENT_TYPE, embedFaces } from '@/word/font-embed';
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
const REL_FONT_TABLE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable';
const FONT_TABLE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml';
const FONT_TABLE_PART = 'word/fontTable.xml';
const REL_SETTINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const SETTINGS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
const SETTINGS_PART = 'word/settings.xml';
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const STYLES_PART = 'word/styles.xml';
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
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/**
 * The namespaces a part's root declares: WordprocessingML's and the
 * relationships', and — where the part uses one — Word 2010's `w14`, which a
 * reader that does not know it is told to pass over (ECMA-376 Part 3
 * `mc:Ignorable`) rather than to refuse the file.
 *
 * @param inner The part's content, to see whether it uses `w14`.
 * @returns The attributes, each with its leading space.
 */
function rootNamespaces(inner: string): string {
  const base =
    ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  return inner.includes('<w14:')
    ? `${base} xmlns:w14="${W14_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w14"`
    : base;
}
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

// 1 pt = 12700 EMU (English Metric Units, the DrawingML coordinate).
const EMU_PER_PT = 12700;

// The raster formats the PDF path embeds (detectImageFormat) → media naming.
//
// JPEG 2000 is deliberately absent. §15.2.14 lists the image parts a
// WordprocessingML package may carry and it is not among them: no consumer of a
// .docx displays one. Written anyway — which it was — the picture is a hole in
// the page and the file carries the bytes for nothing. S2.pdf put six of them
// in, and both of its plates came back blank in Word and in LibreOffice alike.
const RASTER_MEDIA: Readonly<Record<string, { ext: string; contentType: string }>> = {
  png: { ext: 'png', contentType: 'image/png' },
  jpeg: { ext: 'jpeg', contentType: 'image/jpeg' },
  gif: { ext: 'gif', contentType: 'image/gif' },
  bmp: { ext: 'bmp', contentType: 'image/bmp' },
};

/**
 * Why an image was not written, in terms of the image — for the loss report,
 * which used to say "image bytes missing" whatever the reason, including for
 * bytes that were all there and in a format the format does not admit.
 */
function imageRefusal(resource: ResourceId | undefined, state: WriteState): string {
  const bytes = resource === undefined ? undefined : state.resources.get(resource);
  if (!bytes) return 'image bytes missing';
  const format = detectImageFormat(bytes);
  if (format === 'jpeg2000') {
    return 'JPEG 2000 picture dropped — §15.2.14 admits no such image part, and no reader of a .docx shows one';
  }
  return `picture dropped — ${format ?? 'an unrecognised'} is not an image part a .docx may carry`;
}

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
  // The family each face a run names belongs to (a PDF's `Inter-SemiBold` is
  // Word's `Inter`), and the families written, for the font table.
  readonly faceFamilies?: ReadonlyMap<string, FaceFamily>;
  readonly familiesUsed: Map<string, FaceFamily>;
  // The outlines of the faces a run names, where the source carried them, and
  // the faces the runs written name — the ones the package embeds.
  readonly faceOutlines?: ReadonlyMap<string, FaceOutlines>;
  readonly facesUsed: Set<string>;
  // Every z-order the document's floats state, by rank (see `relativeHeight`).
  readonly zRanks: ReadonlyMap<number, number>;
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

/**
 * Serialize a {@link FlowDoc} to a WordprocessingML package (E-DOCX) — the
 * inverse of the docx reader. A flow medium with zero layout and zero I/O: the
 * body's resolved properties are written as direct (denormalized) formatting, so
 * the round-trip is semantic, not byte-for-byte. Emits the main document plus
 * numbering, footnotes/endnotes, comments (+ commentsExtended), per-section
 * headers/footers, charts and media parts; anything not yet serialized is
 * reported as a {@link Loss}.
 *
 * @param flow The interlayer to write back.
 * @returns The encoded `.docx` bytes and the loss report.
 */
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
    ...(flow.faceFamilies ? { faceFamilies: flow.faceFamilies } : {}),
    familiesUsed: new Map(),
    ...(flow.faceOutlines ? { faceOutlines: flow.faceOutlines } : {}),
    facesUsed: new Set(),
    zRanks: zRanksOf(flow),
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

  emitBody(body, flow.body, losses, state, docScope, sectPrByClosingIndex);
  if (finalSectPr) body.push(finalSectPr);

  const bodyXml = body.join('');
  const documentXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document${rootNamespaces(bodyXml)}>` +
    `<w:body>${bodyXml}</w:body>` +
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

  // §17.8.1 — the faces the runs name, embedded where the source carried
  // their outlines and their licence lets them travel.
  const embedded =
    state.faceOutlines && state.faceFamilies
      ? embedFaces(state.facesUsed, state.faceOutlines, state.faceFamilies, losses)
      : undefined;
  if (embedded && embedded.relationships.length > 0) {
    extraPartRels.push({ sourcePart: FONT_TABLE_PART, relationships: [...embedded.relationships] });
  }

  // §17.8.3 — the font table: every family the runs name, with the kind of
  // face it is (§17.8.3.10), so a reader without it substitutes a face of the
  // same kind instead of its default serif.
  const fontTablePart =
    state.familiesUsed.size > 0
      ? {
          path: FONT_TABLE_PART,
          data: encoder.encode(fontTableXml(state.familiesUsed, embedded?.elements)),
          contentType: FONT_TABLE_CONTENT_TYPE,
        }
      : undefined;
  if (fontTablePart) {
    docScope.rels.push({
      id: `rId${++docScope.relSeq}`,
      type: REL_FONT_TABLE,
      target: 'fontTable.xml',
      targetMode: 'Internal',
    });
  }

  // §17.15.1 — the settings the document is set by, where it states any. Left
  // out, a document with headers for its even pages printed its odd pages'
  // on every page once saved, and one bound along its head was bound along
  // its side.
  const settings = settingsXml(flow, sections, (embedded?.parts.length ?? 0) > 0);
  const settingsPart =
    settings !== undefined
      ? { path: SETTINGS_PART, data: encoder.encode(settings), contentType: SETTINGS_CONTENT_TYPE }
      : undefined;
  if (settingsPart) {
    docScope.rels.push({
      id: `rId${++docScope.relSeq}`,
      type: REL_SETTINGS,
      target: 'settings.xml',
      targetMode: 'Internal',
    });
  }

  // §17.7.5 — what a property stated nowhere is, stated (see `stylesXml`).
  const stylesPart = {
    path: STYLES_PART,
    data: encoder.encode(stylesXml()),
    contentType: STYLES_CONTENT_TYPE,
  };
  docScope.rels.push({
    id: `rId${++docScope.relSeq}`,
    type: REL_STYLES,
    target: 'styles.xml',
    targetMode: 'Internal',
  });

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
      stylesPart,
      ...(numberingPart ? [numberingPart] : []),
      ...(fontTablePart ? [fontTablePart] : []),
      ...(settingsPart ? [settingsPart] : []),
      ...extraParts,
      ...state.chartParts,
      ...state.mediaParts,
      ...(embedded?.parts ?? []),
    ],
    ...((embedded?.parts.length ?? 0) > 0
      ? { defaultsByExtension: { odttf: OBFUSCATED_FONT_CONTENT_TYPE } }
      : {}),
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

/**
 * §17.15.1.78 `w:settings` — the document-wide settings the model carries, in
 * the order CT_Settings declares them: `w:embedTrueTypeFonts` and
 * `w:saveSubsetFonts` (§17.15.1.42, .74) where the package embeds its faces,
 * `w:gutterAtTop` (§17.15.1.49), `w:evenAndOddHeaders` (§17.15.1.36, which a
 * section carries in the model), and inside `w:compat` §17.15.3.4's
 * `w:doNotExpandShiftReturn` before the version of Word the document is laid
 * out for ([MS-DOCX] `w:compatSetting` `compatibilityMode`), which CT_Compat
 * keeps last.
 *
 * @param flow        The document.
 * @param sections    The sections being written.
 * @param embedsFonts Whether the package embeds fonts — which then stay
 *                    embedded, as subsets, when the document is saved again.
 * @returns The part's XML, or undefined where the document states none.
 */
function settingsXml(
  flow: FlowDoc,
  sections: ReadonlyArray<Section>,
  embedsFonts: boolean,
): string | undefined {
  const parts: Array<string> = [];
  if (embedsFonts) parts.push('<w:embedTrueTypeFonts/><w:saveSubsetFonts/>');
  if (flow.gutterAtTop === true) parts.push('<w:gutterAtTop/>');
  if (sections.some((sec) => sec.properties.evenAndOddHeaders === true)) {
    parts.push('<w:evenAndOddHeaders/>');
  }
  const compat = [
    ...(flow.doNotExpandShiftReturn === true ? ['<w:doNotExpandShiftReturn/>'] : []),
    ...(flow.compatibilityMode !== undefined
      ? [
          '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word"' +
            ` w:val="${String(flow.compatibilityMode)}"/>`,
        ]
      : []),
  ];
  if (compat.length > 0) parts.push(`<w:compat>${compat.join('')}</w:compat>`);
  if (parts.length === 0) return undefined;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${parts.join('')}</w:settings>`
  );
}

/**
 * §17.7.2 `w:styles` — what a property the document states nowhere IS: the
 * values the model resolves an empty sheet to (`DEFAULT_RUN`, `DEFAULT_PARA`
 * — the very values a run or paragraph is written without), as §17.7.5
 * `w:docDefaults`, under an empty default paragraph style (§17.7.4.17).
 *
 * A package with no styles part leaves them to its reader, and Word fills
 * them from its own template's Normal: 8pt after every paragraph, lines 1.08
 * apart, runs in 12pt — so every paragraph of a reconstruction stood 8pt
 * further down than its page set it, and a run of the model's own 11pt, which
 * states no size, was set in 12.
 */
function stylesXml(): string {
  const size = Math.round(DEFAULT_RUN.fontSizePt * 2);
  const widowControl = DEFAULT_PARA.widowControl
    ? '<w:widowControl/>'
    : '<w:widowControl w:val="0"/>';
  const spacing =
    `<w:spacing w:before="${twips(DEFAULT_PARA.spacingBefore)}"` +
    ` w:after="${twips(DEFAULT_PARA.spacingAfter)}"` +
    ` w:line="${twips(DEFAULT_PARA.spacingLine)}" w:lineRule="${DEFAULT_PARA.spacingLineRule}"/>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults>' +
    `<w:rPrDefault><w:rPr><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr>${widowControl}${spacing}</w:pPr></w:pPrDefault>` +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '</w:styles>'
  );
}

// §17.8.3.9 CT_Font — family (§17.8.3.10) before pitch (§17.8.3.13), and the
// embedded faces (§17.8.3.3–6) after both: the schema's order.
function fontTableXml(
  families: ReadonlyMap<string, FaceFamily>,
  embeds?: ReadonlyMap<string, string>,
): string {
  const fonts = [...families.values()]
    .map(
      (f) =>
        `<w:font w:name="${escapeAttr(f.family)}"><w:family w:val="${f.generic}"/>` +
        `<w:pitch w:val="${f.generic === 'modern' ? 'fixed' : 'variable'}"/>` +
        `${embeds?.get(f.family) ?? ''}</w:font>`,
    )
    .join('');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${fonts}</w:fonts>`
  );
}

/**
 * §20.4.2.3 `relativeHeight` — a float's z-order, written the way Word writes
 * it: a rank over every z-order the document states, counted up from Word's
 * own floor.
 *
 * The model's z-order is an ORDER, and may be any number. Written as it stood,
 * a PDF's first two marks got 0 and 1 — which Word, and LibreOffice after it,
 * read as "above everything": an "APPROVED" stamp's blue box came back over
 * its own lettering. A tagged reading counts up from minus a million, and a
 * negative number is not a value the attribute admits at all.
 */
function relativeHeight(zOrder: number | undefined, state: WriteState): number {
  const rank = zOrder !== undefined ? state.zRanks.get(zOrder) : undefined;
  // A float that states no order stands over those that do, in document order.
  const at = rank ?? state.zRanks.size + ++state.drawingSeq;
  const step = Math.max(1, Math.min(Z_STEP, Math.floor((Z_CEILING - Z_FLOOR) / Z_ROOM)));
  return Math.min(Z_FLOOR + at * step, Z_CEILING);
}

/** Word's first float, and the highest value §20.4.2.3 lets a float take. */
const Z_FLOOR = 0xf000000;
const Z_CEILING = 0x1dffffff;
/** …and the step Word counts up in, while there is room for it. */
const Z_STEP = 1024;
const Z_ROOM = 1 << 17;

/** Every z-order the document's floats state, and each one's rank among them. */
function zRanksOf(flow: FlowDoc): Map<number, number> {
  const seen = new Set<number>();
  const visit = (blocks: ReadonlyArray<BodyElement>): void => {
    for (const el of blocks) {
      const z =
        el.kind === 'image'
          ? el.image.float?.zOrder
          : el.kind === 'shape'
            ? el.shape.float?.zOrder
            : el.kind === 'chart'
              ? el.chart.float?.zOrder
              : undefined;
      if (z !== undefined && Number.isFinite(z)) seen.add(z);
      if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
      }
    }
  };
  visit(flow.body);
  for (const band of flow.headersFooters?.values() ?? []) visit(band);
  for (const note of flow.footnotes?.values() ?? []) visit(note);
  for (const note of flow.endnotes?.values() ?? []) visit(note);
  return new Map([...seen].sort((a, b) => a - b).map((z, i) => [z, i] as const));
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
  const notesXml = noteXmls.join('');
  const xml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<${cfg.rootTag}${rootNamespaces(notesXml)}>` +
    stub('separator', -1, '<w:separator/>') +
    stub('continuationSeparator', 0, '<w:continuationSeparator/>') +
    notesXml +
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
    const bandXml = inner.join('');
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<${root}${rootNamespaces(bandXml)}>${bandXml}</${root}>`;
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

/**
 * The {@link DocumentWriter} registration for `.docx`: its id, the medium it
 * consumes (`flow`), the feature set it supports, and the {@link writeDocx}
 * entry point.
 */
export const docxWriter: DocumentWriter<FlowDoc> = {
  id: 'docx',
  consumes: 'flow',
  supports: new Set([FEATURES.text, FEATURES.fontsEmbedding]),
  write: (doc) => writeDocx(doc),
};

/**
 * The document body, with consecutive floating drawings gathered into one
 * carrier paragraph.
 *
 * §20.4.2.3 — a floating drawing is placed by its anchor and not by where its
 * paragraph lands, so the paragraph is only somewhere for the anchor to hang
 * from. Given one paragraph EACH, a page of placed artwork is that many
 * paragraphs of flow: 160F-2019.pdf's one-page form is 355 drawings, and Word
 * ran it to sixteen pages of empty lines with the artwork anchored to them.
 * Gathered, the page costs one paragraph however much is placed on it.
 *
 * A block that closes a section keeps its own paragraph — the section break
 * rides that paragraph's properties, and it must be the last one in the
 * section.
 */
function emitBody(
  out: Array<string>,
  blocks: ReadonlyArray<BodyElement>,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
  sectPrByClosingIndex: ReadonlyMap<number, string>,
): void {
  let carried: Array<string> = [];
  // The paragraph the drawings are anchored in is the first one's own: a mark
  // the page drew where it stands comes with a carrier that takes no room, and
  // written bare the carrier took the reader's default line — a blank line in
  // the flow for every run of rules and fills an invoice draws.
  let carrier: ParagraphProperties | undefined;
  const flush = (): void => {
    if (carried.length > 0) out.push(`<w:p>${pPrWithSect(carrier ?? {})}${carried.join('')}</w:p>`);
    carried = [];
    carrier = undefined;
  };
  blocks.forEach((el, idx) => {
    const closing = sectPrByClosingIndex.get(idx);
    const anchored =
      closing === undefined ? floatingDrawingRun(el, losses, state, scope) : undefined;
    if (anchored !== undefined) {
      carrier ??=
        el.kind === 'image'
          ? el.image.paragraphProperties
          : el.kind === 'shape'
            ? el.shape.paragraphProperties
            : undefined;
      carried.push(anchored);
      return;
    }
    flush();
    emitBlock(out, el, losses, state, scope, closing);
  });
  flush();
}

/**
 * The run holding one block's drawing, when that block states where on the page
 * it goes — `undefined` for everything that flows, which the ordinary block
 * emitter handles.
 */
function floatingDrawingRun(
  el: BodyElement,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
): string | undefined {
  if (el.kind === 'image' && el.image.float) {
    const drawing = drawingXml(
      el.image.resource,
      el.image.width,
      el.image.height,
      el.image.altText,
      state,
      scope,
      el.image.float,
      el.image,
    );
    if (drawing === '') {
      losses.push({
        severity: 'dropped',
        feature: FEATURES.images,
        detail: imageRefusal(el.image.resource, state),
      });
      return '';
    }
    return `<w:r>${drawing}</w:r>`;
  }
  if (el.kind === 'shape' && el.shape.float) {
    return `<w:r>${shapeDrawingXml(el.shape, losses, state, scope)}</w:r>`;
  }
  return undefined;
}

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
      el.image.float,
      el.image,
    );
    if (drawing) {
      // An image is emitted as a paragraph, so a closing section break rides
      // its pPr (like a text paragraph) rather than a separate carrier — which
      // would otherwise re-read as an extra empty paragraph.
      out.push(
        `<w:p>${pPrWithSect(el.image.paragraphProperties, closingSectPr)}<w:r>${drawing}</w:r></w:p>`,
      );
    } else {
      losses.push({
        severity: 'dropped',
        feature: FEATURES.images,
        detail: imageRefusal(el.image.resource, state),
      });
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
    '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="Chart ${id}"${descr}/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"' +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing>'
  );
}

/**
 * §20.4.2.3 — the frame a drawing is placed by: `wp:anchor` where the block
 * states where on the page it belongs, `wp:inline` where it belongs in the run
 * of text.
 *
 * Everything used to be inline. The document model has carried the anchor since
 * the reader learned to read one, and the PDF reconstruction fills it in for
 * every line and rule it lifts off a page — and all of it was thrown away here,
 * so a converted page came out as a column of drawings stacked in document
 * order. S2.pdf's two colour wheels marched down the left margin as eighteen
 * separate wedges; 160F-2019.pdf's one-page form ran to sixteen pages, because
 * each of its 355 drawings took a paragraph of its own in the flow.
 *
 * The attribute list is not decoration: CT_Anchor requires `simplePos`,
 * `relativeHeight`, `behindDoc`, `locked`, `layoutInCell` and `allowOverlap`,
 * and the children are a SEQUENCE — position, extent, wrap, docPr, graphic —
 * that Word refuses out of order.
 *
 * @param float The placement, or `undefined` for a drawing that flows.
 * @param cx    Width in EMU.
 * @param cy    Height in EMU.
 * @param head  The `wp:docPr` (and anything else) that precedes the graphic.
 * @param graphic The `a:graphic` element the frame carries.
 */
function drawingFrame(
  float: FloatAnchor | undefined,
  cx: number,
  cy: number,
  head: string,
  graphic: string,
  state: WriteState,
  rotation60k?: number,
): string {
  const WP_NS =
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  if (!float) {
    // §20.4.2.8 — the four distances a drawing keeps from the text around it.
    // Absent, Word reads them as zero and LibreOffice supplies its own frame
    // spacing: bug1708040.pdf's logo, an inline picture at the margin, came
    // back nine points right of where the page draws it.
    return `<w:drawing><wp:inline${WP_NS} distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>${head}${graphic}</wp:inline></w:drawing>`;
  }
  const emu = (pt: number | undefined): number =>
    Number.isFinite(pt) ? Math.round((pt ?? 0) * EMU_PER_PT) : 0;
  const d = float.wrapDist;
  const dist =
    ` distT="${String(emu(d?.topPt))}" distB="${String(emu(d?.bottomPt))}"` +
    ` distL="${String(emu(d?.leftPt))}" distR="${String(emu(d?.rightPt))}"`;
  // §20.4.2.3 — the z-order among the page's floats. A drawing that states none
  // still needs a number, and one that rises with document order keeps the
  // painting order the source had.
  const z = relativeHeight(float.zOrder, state);
  const attrs =
    `${dist} simplePos="0" relativeHeight="${String(z)}"` +
    ` behindDoc="${float.behind ? '1' : '0'}" locked="0"` +
    ` layoutInCell="${float.inCell === false ? '0' : '1'}" allowOverlap="1"`;
  const pos = (
    tag: 'wp:positionH' | 'wp:positionV',
    p: { relativeFrom: string; offsetPt?: number; align?: string } | undefined,
    fallback: string,
  ): string => {
    const from = p?.relativeFrom ?? fallback;
    // An offset is the precise answer and an alignment the coarse one; a
    // drawing that states neither sits at the origin of what it is measured in.
    const body =
      p?.align !== undefined
        ? `<wp:align>${p.align}</wp:align>`
        : `<wp:posOffset>${String(emu(p?.offsetPt))}</wp:posOffset>`;
    return `<${tag} relativeFrom="${escapeAttr(from)}">${body}</${tag}>`;
  };
  // `notBeside` is the FRAME's mode — the drawing keeps its place and no text
  // may stand beside it, which of the anchor's wraps is topAndBottom.
  const side = ` wrapText="${float.wrapSide ?? 'bothSides'}"`;
  const wrap =
    float.wrap === 'square'
      ? `<wp:wrapSquare${side}/>`
      : float.wrap === 'tight'
        ? `<wp:wrapTight${side}/>`
        : float.wrap === 'through'
          ? `<wp:wrapThrough${side}/>`
          : float.wrap === 'topAndBottom' || float.wrap === 'notBeside'
            ? '<wp:wrapTopAndBottom/>'
            : '<wp:wrapNone/>';
  return (
    `<w:drawing><wp:anchor${WP_NS}${attrs}>` +
    '<wp:simplePos x="0" y="0"/>' +
    pos('wp:positionH', float.posH, 'column') +
    pos('wp:positionV', float.posV, 'paragraph') +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    effectExtentXml(cx, cy, rotation60k) +
    wrap +
    head +
    graphic +
    '</wp:anchor></w:drawing>'
  );
}

/**
 * §20.4.2.3 `wp:effectExtent` — how far the drawing reaches PAST the extent.
 *
 * `wp:extent` is the shape's own size, unrotated: fdo75722-dml.docx states
 * 1964690 × 1240790 there and repeats it in `a:ext`, then carries the turn's
 * overhang here (l 198120, t 433705, r 199390, b 430530). Written as zeroes —
 * which they always were — a renderer reserves the flat box for a shape that
 * draws across the turned one, and 160F-2019.pdf's "Nature", set on its side
 * down the middle of a column, came back lying flat across it.
 *
 * A box `w × h` turned by θ about its centre spans `|w·cosθ| + |h·sinθ|`
 * across and `|w·sinθ| + |h·cosθ|` down; half of each growth hangs off each
 * side. A turn that makes the box no wider hangs off by nothing.
 */
function effectExtentXml(cx: number, cy: number, rotation60k: number | undefined): string {
  const none = '<wp:effectExtent l="0" t="0" r="0" b="0"/>';
  if (rotation60k === undefined || !Number.isFinite(rotation60k) || rotation60k % 21600000 === 0) {
    return none;
  }
  const rad = ((rotation60k / 60000) * Math.PI) / 180;
  const c = Math.abs(Math.cos(rad));
  const s = Math.abs(Math.sin(rad));
  const overhang = (span: number, turned: number): number =>
    Math.max(0, Math.round((turned - span) / 2));
  const h = overhang(cx, cx * c + cy * s);
  const v = overhang(cy, cx * s + cy * c);
  return h === 0 && v === 0
    ? none
    : `<wp:effectExtent l="${String(h)}" t="${String(v)}" r="${String(h)}" b="${String(v)}"/>`;
}

// Allocate (or reuse) a media part + image relationship for a resource, then
// emit the w:drawing markup the reader round-trips. Returns '' when
// the resource has no bytes (an unresolved image — the caller drops it).
function drawingXml(
  resource: ResourceId | undefined,
  widthPt: number,
  heightPt: number,
  altText: string | undefined,
  state: WriteState,
  scope: PartScope,
  float?: FloatAnchor,
  look?: PictureLook,
): string {
  if (resource === undefined) return '';
  const relId = mediaRelId(resource, state, scope);
  if (relId === undefined) return '';
  const cx = Math.round(widthPt * EMU_PER_PT);
  const cy = Math.round(heightPt * EMU_PER_PT);
  const id = ++state.drawingSeq;
  const descr = altText ? ` descr="${escapeAttr(altText)}"` : '';
  // §20.1.8.4 / §20.1.8.55 / §20.1.7.6 — how the picture is drawn into its
  // frame: how opaque, which part of the source, turned or mirrored. Read
  // from a .docx and never written back, a cropped picture came back whole
  // and squeezed into the frame its crop was sized for.
  const alpha =
    look?.alpha !== undefined && look.alpha < 1
      ? `<a:alphaModFix amt="${String(Math.round(Math.max(0, look.alpha) * 100000))}"/>`
      : '';
  const graphic =
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Image ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}">${alpha}</a:blip>${srcRectXml(look?.crop)}` +
    '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    `<pic:spPr>${xfrmXml(look, cx, cy)}` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic>';
  return drawingFrame(
    float,
    cx,
    cy,
    `<wp:docPr id="${id}" name="Image ${id}"${descr}/>`,
    graphic,
    state,
    look?.rotation60k,
  );
}

/** What of a picture's own look the frame writes: its crop, turn, mirror and opacity. */
type PictureLook = Pick<ImageBlock, 'crop' | 'rotation60k' | 'flipH' | 'flipV' | 'alpha'>;

/** §20.1.8.55 `a:srcRect` — each edge cut away, in thousandths of a percent. */
function srcRectXml(crop: ImageCrop | undefined): string {
  if (!crop) return '';
  const edge = (name: string, v: number): string =>
    v > 0 ? ` ${name}="${String(Math.round(v * 100000))}"` : '';
  const attrs =
    edge('l', crop.left) + edge('t', crop.top) + edge('r', crop.right) + edge('b', crop.bottom);
  return attrs === '' ? '' : `<a:srcRect${attrs}/>`;
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
  const members = shape.children ?? [];
  const graphic =
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    (members.length > 0
      ? `<a:graphicData uri="${WPG_URI}">` +
        groupXml('wpg:wgp', members, { x: 0, y: 0, cx, cy }, losses, state, scope)
      : `<a:graphicData uri="${WPS_URI}">` +
        wspXml(shape, cx, cy, undefined, losses, state, scope)) +
    '</a:graphicData></a:graphic>';
  return drawingFrame(
    shape.float,
    cx,
    cy,
    `<wp:docPr id="${id}" name="Shape ${id}"${descr}/>`,
    graphic,
    state,
    shape.transform?.rotation60k,
  );
}

/**
 * `wps:wsp` — one shape: its box, geometry, fill and outline, and the text it
 * holds. `id` is its own `wps:cNvPr` inside a group, where the members name
 * themselves; a shape alone is named by the frame's `wp:docPr`.
 */
function wspXml(
  shape: ShapeBlock,
  cx: number,
  cy: number,
  member: { readonly id: number; readonly x: number; readonly y: number } | undefined,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
): string {
  const spPr =
    `<wps:spPr>${xfrmXml(shape.transform, cx, cy, member)}${geomXml(shape.geometry)}` +
    `${fillXml(shape.fill)}${shape.line ? lineXml(shape.line) : ''}</wps:spPr>`;
  const txbx = shape.text ? txbxXml(shape.text, losses, state, scope) : '';
  return (
    `<wps:wsp xmlns:wps="${WPS_URI}">` +
    (member ? `<wps:cNvPr id="${member.id}" name="Shape ${member.id}"/>` : '') +
    '<wps:cNvSpPr/>' +
    spPr +
    txbx +
    bodyPrXml(shape.text) +
    '</wps:wsp>'
  );
}

const WPG_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

/**
 * §20.5.2.17 `wpg:wgp` / `wpg:grpSp` — a group: its box, and its members in
 * the order they are painted. The members are placed in a child space the
 * same as the box, so a member's offset from the group's corner is its offset
 * in the file. The inverse of drawing-parser's `groupChildren`.
 *
 * Written as one shape, the group came back an empty box: a figure's paths
 * and labels were thrown away with it.
 */
function groupXml(
  tag: 'wpg:wgp' | 'wpg:grpSp',
  members: ReadonlyArray<ShapeGroupChild>,
  box: { readonly x: number; readonly y: number; readonly cx: number; readonly cy: number },
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
): string {
  const ns = tag === 'wpg:wgp' ? ` xmlns:wpg="${WPG_URI}"` : '';
  const xfrm =
    `<a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${box.cx}" cy="${box.cy}"/></a:xfrm>`;
  const inner = members.map((m) => memberXml(m, losses, state, scope)).join('');
  return `<${tag}${ns}><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrm}</wpg:grpSpPr>${inner}</${tag}>`;
}

/**
 * One member of a group: a group of its own, a picture (`pic:pic`, a shape
 * whose fill is the picture — how the reader brings one in), or a shape.
 */
function memberXml(
  member: ShapeGroupChild,
  losses: Array<Loss>,
  state: WriteState,
  scope: PartScope,
): string {
  const s = member.shape;
  const x = Math.round(member.xPt * EMU_PER_PT);
  const y = Math.round(member.yPt * EMU_PER_PT);
  const cx = Math.round(s.width * EMU_PER_PT);
  const cy = Math.round(s.height * EMU_PER_PT);
  if (s.children !== undefined && s.children.length > 0) {
    return groupXml('wpg:grpSp', s.children, { x, y, cx, cy }, losses, state, scope);
  }
  const id = ++state.drawingSeq;
  if (s.fill.kind !== 'picture') {
    return wspXml(s, cx, cy, { id, x, y }, losses, state, scope);
  }
  const relId =
    s.fill.imageResource !== undefined ? mediaRelId(s.fill.imageResource, state, scope) : undefined;
  if (relId === undefined) {
    losses.push({
      severity: 'dropped',
      feature: FEATURES.images,
      detail: imageRefusal(s.fill.imageResource, state),
    });
    return '';
  }
  const alpha =
    s.fill.alpha !== undefined && s.fill.alpha < 1
      ? `<a:alphaModFix amt="${String(Math.round(Math.max(0, s.fill.alpha) * 100000))}"/>`
      : '';
  return (
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Image ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}">${alpha}</a:blip>${srcRectXml(s.fill.imageCrop)}` +
    '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    `<pic:spPr>${xfrmXml(s.transform, cx, cy, { x, y })}` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
  );
}

// §20.1.7.6 a:xfrm — rotation/flips plus the off+ext the reader uses as a size
// fallback. Always emitted so a re-read recovers the box even without extent;
// a group's member states where in the group it stands.
function xfrmXml(
  t: ShapeTransform | undefined,
  cx: number,
  cy: number,
  at?: { readonly x: number; readonly y: number },
): string {
  const rot = t?.rotation60k !== undefined ? ` rot="${t.rotation60k}"` : '';
  const flipH = t?.flipH ? ' flipH="1"' : '';
  const flipV = t?.flipV ? ' flipV="1"' : '';
  return (
    `<a:xfrm${rot}${flipH}${flipV}><a:off x="${at?.x ?? 0}" y="${at?.y ?? 0}"/>` +
    `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
  );
}

/**
 * The factor that carries a path into whole units, and the space it lands in.
 *
 * §20.1.10.16 — ST_Coordinate is a LONG: a path states its geometry in whole
 * units, and a fractional one is not a value the attribute admits. Word does
 * not round it, and does not ignore the shape: it refuses the document, saying
 * only that it "experienced an error trying to open the file".
 *
 * A path space is arbitrary — `a:path w`/`h` declare the coordinate system the
 * shape's extent is mapped onto — so a path that does not land on whole units
 * is restated in a finer space rather than rounded where it stands. A path read
 * off a PDF is in points, and a point split ten thousand ways is finer than any
 * renderer draws. The factor is bounded so the space stays far inside the
 * coordinate range whatever the path's own size, and a path already on whole
 * units is left exactly as it is.
 */
const PATH_SPACE = 1e7;

function pathScale(g: NonNullable<ShapeGeometry['custom']>): number {
  const values = [g.pathWidth, g.pathHeight];
  for (const c of g.commands) {
    if (c.cmd === 'move' || c.cmd === 'line') values.push(c.x, c.y);
    else if (c.cmd === 'cubic') values.push(c.x1, c.y1, c.x2, c.y2, c.x, c.y);
    else if (c.cmd === 'quad') values.push(c.x1, c.y1, c.x, c.y);
    else if (c.cmd === 'arc') values.push(c.wR, c.hR);
  }
  if (values.every((v) => Number.isInteger(v))) return 1;
  const span = Math.max(...values.map((v) => (Number.isFinite(v) ? Math.abs(v) : 0)), 1);
  return Math.max(1, Math.floor(PATH_SPACE / span));
}

function geomXml(g: ShapeGeometry): string {
  if (g.kind === 'custom' && g.custom) {
    const k = pathScale(g.custom);
    // A coordinate that is no number states nothing, and the origin is the one
    // place a path can start from that draws nothing wrong.
    const c1 = (v: number): number => (Number.isFinite(v) ? Math.round(v * k) : 0);
    const cmds = g.custom.commands
      .map((c) => {
        switch (c.cmd) {
          case 'move':
            return `<a:moveTo><a:pt x="${c1(c.x)}" y="${c1(c.y)}"/></a:moveTo>`;
          case 'line':
            return `<a:lnTo><a:pt x="${c1(c.x)}" y="${c1(c.y)}"/></a:lnTo>`;
          case 'cubic':
            return (
              '<a:cubicBezTo>' +
              `<a:pt x="${c1(c.x1)}" y="${c1(c.y1)}"/>` +
              `<a:pt x="${c1(c.x2)}" y="${c1(c.y2)}"/>` +
              `<a:pt x="${c1(c.x)}" y="${c1(c.y)}"/>` +
              '</a:cubicBezTo>'
            );
          case 'quad':
            return (
              `<a:quadBezTo><a:pt x="${c1(c.x1)}" y="${c1(c.y1)}"/>` +
              `<a:pt x="${c1(c.x)}" y="${c1(c.y)}"/></a:quadBezTo>`
            );
          case 'arc':
            // The radii are coordinates and scale with the space; the angles are
            // in 60000ths of a degree (§20.1.10.3) and do not.
            return (
              `<a:arcTo wR="${c1(c.wR)}" hR="${c1(c.hR)}" ` +
              `stAng="${Math.round(c.stAng)}" swAng="${Math.round(c.swAng)}"/>`
            );
          case 'close':
            return '<a:close/>';
        }
      })
      .join('');
    // §20.1.9.15 `a:path @w/@h` — the coordinate space the shape's extent is
    // mapped onto, so a zero is a space with no size in that direction and
    // nothing can be mapped into it: a horizontal rule written `h="0"` is in
    // the package and on no page, which is what became of
    // annotation-line-without-appearance-empty-Rect.pdf's red line and of
    // every flat rule drawn as a path.
    const w = Math.max(1, c1(g.custom.pathWidth));
    const h = Math.max(1, c1(g.custom.pathHeight));
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
    return `<a:solidFill>${srgbXml(f.colorHex, f.alpha)}</a:solidFill>`;
  if (f.kind === 'gradient' && f.gradient) return gradFillXml(f.gradient, f.alpha);
  return '<a:noFill/>';
}

/**
 * §20.1.2.3.19 `a:srgbClr`, with the §20.1.2.3.1 `a:alpha` of a fill that is
 * seen through. The model carried the opacity and the writer dropped it:
 * bug1755507.pdf lays a card on a shadow painted at a fifth of full strength,
 * and the shadow came back as a solid black slab around the card.
 */
function srgbXml(colorHex: string, alpha: number | undefined): string {
  if (alpha === undefined || !(alpha < 1)) return `<a:srgbClr val="${colorHex}"/>`;
  const val = Math.round(Math.max(0, alpha) * 100000);
  return `<a:srgbClr val="${colorHex}"><a:alpha val="${String(val)}"/></a:srgbClr>`;
}

// A gradient fill → a:gradFill (EP16): stops as a:gs (@pos in 1000ths of a
// percent), direction as a:lin (@ang in 60000ths of a degree) or a:path (radial).
function gradFillXml(g: ShapeGradient, alpha?: number): string {
  const stops = g.stops
    .map((s) => {
      const pos = Math.round(Math.max(0, Math.min(1, s.offset)) * 100000);
      return `<a:gs pos="${pos}">${srgbXml(s.colorHex, alpha)}</a:gs>`;
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
  else if (l.colorHex) inner.push(`<a:solidFill>${srgbXml(l.colorHex, l.alpha)}</a:solidFill>`);
  // §20.1.8.21 — the author's own pattern wins over a preset beside it, in
  // thousandths of a percent of the line's width, a dash and a space a pair.
  const custom = l.customDash ?? [];
  if (custom.length >= 2) {
    const pairs: Array<string> = [];
    for (let i = 0; i + 1 < custom.length; i += 2)
      pairs.push(
        `<a:ds d="${Math.round(custom[i]! * 100000)}" sp="${Math.round(custom[i + 1]! * 100000)}"/>`,
      );
    inner.push(`<a:custDash>${pairs.join('')}</a:custDash>`);
  } else if (l.dash) inner.push(`<a:prstDash val="${l.dash}"/>`);
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
  const wrap = text.noWrap === true ? ' wrap="none"' : '';
  return (
    '<wps:bodyPr' +
    wrap +
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

/**
 * §17.4.60 — `w:tblPr`. CT_Tbl declares it `minOccurs="1"`: a table states its
 * properties even when it has none to state, and Word refuses to open a file
 * whose `w:tbl` opens straight into its grid. Nothing else in the wild writes
 * one without it — 5432 tables across the LibreOffice, Word and POI corpora,
 * every one of them with a `w:tblPr` — and it went unnoticed here because a
 * table read from a real document always carries SOMETHING (a width, a border,
 * an alignment). One reconstructed from a PDF's glyph positions carries none of
 * it, so `Ream.parse(pdf).convert('docx')` wrote a document Word turned away.
 */
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
  // §17.4.65 — how far in from the margin the table stands, which the model
  // carried and the writer never wrote: every table a PDF set in from its
  // margin came back against it.
  if (p.indentPt !== undefined && p.indentPt !== 0) {
    out.push(`<w:tblInd w:w="${twips(p.indentPt)}" w:type="dxa"/>`);
  }
  const borders = bordersXml('w:tblBorders', p.borders);
  if (borders) out.push(borders);
  // §17.4.53 — a FIXED table is laid out by its grid and nothing else. Left
  // unsaid, a reader sizes the columns to their contents, and a receipt's
  // last column came back too narrow for the number it was ruled to hold.
  if (p.layout === 'fixed') out.push('<w:tblLayout w:type="fixed"/>');
  const margins = cellMarginsXml('w:tblCellMar', p.defaultCellMargins);
  if (margins) out.push(margins);
  return `<w:tblPr>${out.join('')}</w:tblPr>`;
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
  // §17.4.72 — a percentage width goes back as one (fiftieths of a percent);
  // re-spelling it in twips would move the cell on the next read.
  if (p.widthFraction !== undefined) {
    out.push(`<w:tcW w:w="${Math.round(p.widthFraction * 5000)}" w:type="pct"/>`);
  } else if (p.width !== undefined) {
    out.push(`<w:tcW w:w="${twips(p.width)}" w:type="dxa"/>`);
  }
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

function bordersXml(
  tag: 'w:tblBorders' | 'w:tcBorders' | 'w:pBdr',
  borders: CellBorders | undefined,
): string {
  if (!borders) return '';
  const sides = BORDER_SIDES.map(([key, name]) => {
    const b = borders[key];
    if (!b) return '';
    // §17.3.1.24 CT_PBdr has no inside edges: the one between two paragraphs
    // of a bordered set is §17.3.1.5 `w:between`, and a paragraph has no
    // vertical inside edge at all. Written as `w:insideH` Word refuses the file.
    if (tag === 'w:pBdr' && key === 'insideV') return '';
    const el = tag === 'w:pBdr' && key === 'insideH' ? 'w:between' : name;
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
  const rPr = rPrXml(run.properties as ResolvedRunProperties, state);
  if (run.inlineImage !== undefined) {
    const img = run.inlineImage;
    const drawing = drawingXml(
      img.resource,
      img.width,
      img.height,
      undefined,
      state,
      scope,
      undefined,
      img,
    );
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
  // §17.16.19 `w:fldSimple` — a page number is not the text "1", it is the
  // number of the sheet it stands on. Written as text, the foot a PDF's every
  // page shares came back saying "Page 1 of 2" on the second page as well.
  //
  // §17.16.18 — written as a COMPLEX field, every piece of it a run in the
  // run's own properties. Inside `w:fldSimple` LibreOffice sets the result in
  // the paragraph's default size, and a receipt's "Page 1 of 2" came back with
  // its two numbers half as large again as the words around them.
  if (run.field !== undefined && run.text !== '') {
    const piece = (inner: string): string => `<w:r>${rPr}${inner}</w:r>`;
    return (
      piece('<w:fldChar w:fldCharType="begin"/>') +
      piece(`<w:instrText xml:space="preserve"> ${run.field} </w:instrText>`) +
      piece('<w:fldChar w:fldCharType="separate"/>') +
      piece(`<w:t xml:space="preserve">${escapeXml(run.text)}</w:t>`) +
      piece('<w:fldChar w:fldCharType="end"/>') +
      brk
    );
  }
  if (run.text === '') return brk ? `<w:r>${rPr}${brk}</w:r>` : '';
  // §17.3.3.30 — a TAB is an ELEMENT, not a character: written inside `w:t` it
  // is whitespace, and Word draws it as nothing at all. The reconstruction of a
  // PDF sets a line the page laid out on stops with them — an invoice's "Bill
  // to" block stands beside the address it belongs to — and written as text the
  // two ran together.
  //
  // §17.3.3.1 — and so is a line BREAK: the reader reads `w:br` as a newline in
  // the run's text, and written back inside `w:t` a newline is whitespace, so
  // every address, verse and signature read and written again ran its lines
  // together. A run the reader marked as a column break breaks to the next
  // column there instead.
  const lineBreak = run.columnBreak ? '<w:br w:type="column"/>' : '<w:br/>';
  const body = run.text
    .split('\t')
    .map((piece) =>
      piece
        .split('\n')
        .map((line) => (line === '' ? '' : `<w:t xml:space="preserve">${escapeXml(line)}</w:t>`))
        .join(lineBreak),
    )
    .join('<w:tab/>');
  return `<w:r>${rPr}${body}${brk}</w:r>`;
}

// §17.3.2 — run properties as a delta from the resolved defaults.
function rPrXml(r: ResolvedRunProperties, state?: WriteState): string {
  // §17.3.2.28 CT_RPr is a SEQUENCE, and a reader may drop what arrives out of
  // it: rFonts, b, i, strike, color, kern, sz, u, shd, vertAlign, rtl, lang, and
  // Word 2010's own after them. Written
  // in the old order — `w:u` ahead of `w:rFonts` — LibreOffice ignored the
  // underline outright, so annotation-squiggly.pdf's wavy blue rule was in the
  // package and on no page.
  const out: Array<string> = [];
  // A property that is not THERE is not a property with a different value.
  // The header and footer parts are written from raw properties, not resolved
  // ones, and compared straight against the defaults every absent field came
  // out as the string "undefined": `<w:color w:val="undefined"/>` in the foot
  // of every reconstructed PDF, which is not a colour and not valid markup.
  const states = <TKey extends keyof ResolvedRunProperties>(key: TKey): boolean =>
    r[key] !== undefined && r[key] !== DEFAULT_RUN[key];
  const fonts = rFontsXml(r.fontFamily, state);
  if (fonts) out.push(fonts);
  // §17.3.2.2/§17.3.2.17/§17.3.2.39 — Word sets a complex script (Arabic,
  // Hebrew, Thai) by its OWN weight, slant and size, and the run has one of
  // each: stated for Latin text alone, ArabicCIDTrueType.pdf's 36pt lines came
  // back at the reader's default ten.
  if (states('bold')) out.push(toggle('w:b', r.bold), toggle('w:bCs', r.bold));
  if (states('italic')) out.push(toggle('w:i', r.italic), toggle('w:iCs', r.italic));
  if (states('strike')) out.push(toggle('w:strike', r.strike));
  if (states('colorHex')) out.push(`<w:color w:val="${r.colorHex}"/>`);
  // §17.3.2.35 `w:spacing` — how much wider or tighter the characters stand,
  // in twips, between the colour and the kerning. Read from a .docx and never
  // written back: text set expanded or condensed came back at the face's own
  // spacing.
  if (r.letterSpacingPt !== undefined && r.letterSpacingPt !== 0) {
    const twips = Math.round(r.letterSpacingPt * 20);
    if (twips !== 0) out.push(`<w:spacing w:val="${String(twips)}"/>`);
  }
  // §17.3.2.43 `w:w` — the share of its width each glyph is set at, in whole
  // percent (ST_TextScale, up to 600), after the spacing.
  if (r.widthScale !== undefined) {
    const percent = Math.min(600, Math.max(1, Math.round(r.widthScale * 100)));
    if (percent !== 100) out.push(`<w:w w:val="${String(percent)}"/>`);
  }
  // §17.3.2.19 `w:kern` — after the colour and before the size, the schema's
  // place for it; half-points, like the size.
  if (r.kerningMinPt !== undefined && states('kerningMinPt')) {
    out.push(`<w:kern w:val="${Math.round(r.kerningMinPt * 2)}"/>`);
  }
  if (states('fontSizePt')) {
    // §17.3.2.38 w:sz — half-points.
    const half = Math.round(r.fontSizePt * 2);
    out.push(`<w:sz w:val="${half}"/>`, `<w:szCs w:val="${half}"/>`);
  }
  // §17.3.2.40 — `w:u @w:color`, the rule's own colour where it has one: a
  // PDF's `/Underline` annotation states its colour and nothing else does.
  if (states('underline')) out.push(underlineXml(r));
  // §17.3.2.32 — the wash behind the glyphs, which is what a PDF's `/Highlight`
  // annotation marks its words with.
  if (r.shadingColorHex !== undefined) out.push(runShdXml(r.shadingColorHex));
  if (states('verticalAlign')) {
    out.push(`<w:vertAlign w:val="${r.verticalAlign}"/>`);
  }
  if (states('rtl')) out.push(toggle('w:rtl', r.rtl));
  if (r.lang !== undefined) out.push(`<w:lang w:val="${escapeAttr(r.lang)}"/>`);
  // [MS-DOCX] `w14:ligatures` — Word 2010's own, after every element of the
  // base schema (see `rootNamespaces` for the namespace it is declared in).
  if (r.ligatures !== undefined && states('ligatures')) {
    out.push(`<w14:ligatures w14:val="${r.ligatures}"/>`);
  }
  return out.length > 0 ? `<w:rPr>${out.join('')}</w:rPr>` : '';
}

// §17.3.2.40 `w:u` — the style, and the colour when the run states one.
function underlineXml(r: { underline: string; underlineColorHex?: string }): string {
  const color = r.underlineColorHex !== undefined ? ` w:color="${r.underlineColorHex}"` : '';
  return `<w:u w:val="${r.underline}"${color}/>`;
}

// §17.3.2.32 `w:rPr/w:shd` — a solid wash behind the run's own glyphs.
function runShdXml(fill: string): string {
  return `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
}

// §17.3.1 — paragraph properties as a delta from the resolved defaults, the
// INNER content of w:pPr (no wrapper, so a section break can be appended).
function pPrBody(p: ResolvedParagraphProperties): string {
  // §17.3.1.26 CT_PPrBase is a SEQUENCE, and Word enforces it: a child out of
  // order is a file it refuses or a property it drops on the floor. The order
  // below is the schema's — keepNext, keepLines, pageBreakBefore, widowControl,
  // numPr, pBdr, shd, tabs, bidi, spacing, ind, jc, outlineLvl — and it is not
  // the order these were written in until now: `w:pBdr` and `w:tabs` were
  // appended after `w:jc`, which LibreOffice reads anyway and Word does not.
  const out: Array<string> = [];
  if (p.keepNext) out.push('<w:keepNext/>');
  if (p.keepLines) out.push('<w:keepLines/>');
  if (p.pageBreakBefore) out.push('<w:pageBreakBefore/>');
  // §17.3.1.44 — on is what the styles part states, so only off is said here.
  if (p.widowControl === false) out.push('<w:widowControl w:val="0"/>');
  if (p.numbering) {
    // §17.3.1.19 — list membership; the marker itself comes from numbering.xml.
    out.push(
      `<w:numPr><w:ilvl w:val="${p.numbering.ilvl}"/><w:numId w:val="${escapeAttr(p.numbering.numId)}"/></w:numPr>`,
    );
  }
  // §17.3.1.24 `w:pBdr` — the rules the paragraph is drawn with. A PDF has no
  // paragraph borders and draws lines; the reconstruction gives the line to the
  // paragraph it separates (see `pdf-reader/layout`).
  const pBdr = bordersXml('w:pBdr', p.borders);
  if (pBdr) out.push(pBdr);
  // §17.3.1.38 `w:tabs` — the stops the paragraph's own tabs stand on. Without
  // them a tab falls to the default half-inch grid, which is not where the page
  // that was read set its second column.
  //
  // The header and footer path hands this raw properties rather than resolved
  // ones, so the field the type promises may not be there.
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition
  const tabs = p.tabs ?? [];
  if (tabs.length > 0) {
    const stops = tabs
      .map((t) => {
        const pos = twips(t.positionPt);
        const leader = t.leader !== undefined ? ` w:leader="${escapeAttr(t.leader)}"` : '';
        // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition
        return `<w:tab w:val="${escapeAttr(t.alignment ?? 'left')}" w:pos="${String(pos)}"${leader}/>`;
      })
      .join('');
    out.push(`<w:tabs>${stops}</w:tabs>`);
  }
  // A drawing's carrier and a band's lines come with RAW properties, where a
  // direction nobody stated is absent — not the opposite of the default.
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition
  if (p.bidi !== undefined && p.bidi !== DEFAULT_PARA.bidi) out.push(toggle('w:bidi', p.bidi));
  const spacing = spacingXml(p);
  if (spacing) out.push(spacing);
  const ind = indXml(p);
  if (ind) out.push(ind);
  if (JC.has(p.alignment) && p.alignment !== DEFAULT_PARA.alignment) {
    out.push(`<w:jc w:val="${p.alignment}"/>`);
  }
  if (p.outlineLevel !== undefined) out.push(`<w:outlineLvl w:val="${p.outlineLevel}"/>`);
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

/**
 * Whether a measurement is one the file can state.
 *
 * A property that was COMPUTED rather than read can arrive `NaN` — an
 * `undefined` that went through arithmetic, a share of a width that was zero —
 * and `NaN` is not a value ST_SignedTwipsMeasure admits. Word does not ignore
 * the attribute: it refuses the whole document, with no indication of what it
 * objected to. `Ream.parse(pdf).convert('docx')` wrote 355 of them on one form.
 *
 * So a measurement is written only when there IS one, and a paragraph that
 * cannot say how it is indented says nothing and inherits, which is what an
 * absent attribute means (§17.3.1.12).
 */
const stated = (pt: number): boolean => Number.isFinite(pt);

/** §17.18.44 ST_Jc — the tokens the attribute admits, so nothing else is written. */
const JC: ReadonlySet<string> = new Set(['left', 'right', 'center', 'both', 'distribute']);

function indXml(p: ResolvedParagraphProperties): string {
  const attrs: Array<string> = [];
  if (stated(p.indentLeft) && p.indentLeft !== DEFAULT_PARA.indentLeft) {
    attrs.push(`w:left="${twips(p.indentLeft)}"`);
  }
  if (stated(p.indentRight) && p.indentRight !== DEFAULT_PARA.indentRight) {
    attrs.push(`w:right="${twips(p.indentRight)}"`);
  }
  if (stated(p.indentFirstLine) && p.indentFirstLine !== DEFAULT_PARA.indentFirstLine) {
    // A negative first-line indent is a hanging indent (§17.3.1.12).
    if (p.indentFirstLine < 0) attrs.push(`w:hanging="${twips(-p.indentFirstLine)}"`);
    else attrs.push(`w:firstLine="${twips(p.indentFirstLine)}"`);
  }
  return attrs.length > 0 ? `<w:ind ${attrs.join(' ')}/>` : '';
}

function spacingXml(p: ResolvedParagraphProperties): string {
  const attrs: Array<string> = [];
  if (stated(p.spacingBefore) && p.spacingBefore !== DEFAULT_PARA.spacingBefore) {
    attrs.push(`w:before="${twips(p.spacingBefore)}"`);
  }
  if (stated(p.spacingAfter) && p.spacingAfter !== DEFAULT_PARA.spacingAfter) {
    attrs.push(`w:after="${twips(p.spacingAfter)}"`);
  }
  if (
    (p.spacingLineRule !== DEFAULT_PARA.spacingLineRule ||
      p.spacingLine !== DEFAULT_PARA.spacingLine) &&
    p.spacingLine > 0
  ) {
    // §17.3.1.33: 'auto' line spacing is in 240ths of a line, exact/atLeast
    // in twips — and the reader reads either as twips, so an 'auto' line is
    // twelve points a single line (see the HTML writer's `line-height`), and
    // the number goes back as it came. Written as twelfths, Word's own 1.08
    // lines (259) came back as 0.65 of one (155), each line over the last.
    attrs.push(`w:line="${twips(p.spacingLine)}"`, `w:lineRule="${p.spacingLineRule}"`);
  }
  return attrs.length > 0 ? `<w:spacing ${attrs.join(' ')}/>` : '';
}

// §17.3.2.26 w:rFonts — only the slots that differ from the resolved default,
// each named by its FAMILY where the source named a face (see FlowDoc
// `faceFamilies`): a reader looks a font up by the name it was installed under.
function rFontsXml(fonts: FontFamilyMap | undefined, state?: WriteState): string {
  // A band's runs come with RAW properties, and a run that names no face has
  // no font map at all: ZapfDingbats.pdf's foot threw here, and the whole
  // package with it.
  if (fonts === undefined) return '';
  const d = DEFAULT_RUN.fontFamily;
  const family = (name: string): string => {
    const known = state?.faceFamilies?.get(name);
    if (known === undefined) return name;
    state?.familiesUsed.set(known.family, known);
    if (state?.faceOutlines?.has(name) === true) state.facesUsed.add(name);
    return known.family;
  };
  const attrs: Array<string> = [];
  if (fonts.ascii && fonts.ascii !== d.ascii)
    attrs.push(`w:ascii="${escapeAttr(family(fonts.ascii))}"`);
  if (fonts.hAnsi && fonts.hAnsi !== d.hAnsi)
    attrs.push(`w:hAnsi="${escapeAttr(family(fonts.hAnsi))}"`);
  // A face a PDF drew in drew EVERY character of the run — the Arabic and the
  // Han as well as the Latin — so every slot a reader picks by script names it.
  const face =
    fonts.ascii !== undefined && state?.faceFamilies?.has(fonts.ascii) === true
      ? family(fonts.ascii)
      : undefined;
  if (face !== undefined && fonts.eastAsia === undefined) {
    attrs.push(`w:eastAsia="${escapeAttr(face)}"`);
  }
  if (fonts.cs && fonts.cs !== d.cs) attrs.push(`w:cs="${escapeAttr(family(fonts.cs))}"`);
  else if (face !== undefined) attrs.push(`w:cs="${escapeAttr(face)}"`);
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
  // `!== undefined` is not enough: `NaN` is defined, and no measurement.
  if (p.indentLeft !== undefined && stated(p.indentLeft)) {
    attrs.push(`w:left="${twips(p.indentLeft)}"`);
  }
  if (p.indentRight !== undefined && stated(p.indentRight)) {
    attrs.push(`w:right="${twips(p.indentRight)}"`);
  }
  if (p.indentFirstLine !== undefined && stated(p.indentFirstLine)) {
    if (p.indentFirstLine < 0) attrs.push(`w:hanging="${twips(-p.indentFirstLine)}"`);
    else attrs.push(`w:firstLine="${twips(p.indentFirstLine)}"`);
  }
  const ind = attrs.length > 0 ? `<w:ind ${attrs.join(' ')}/>` : '';
  const jc =
    p.alignment !== undefined && JC.has(p.alignment) ? `<w:jc w:val="${p.alignment}"/>` : '';
  return ind || jc ? `<w:pPr>${jc}${ind}</w:pPr>` : '';
}

function rawRunPrXml(r: RunProperties): string {
  // §17.3.2.28 — the schema's own order; see `rPrXml`.
  const out: Array<string> = [];
  const fonts = r.fontFamily ? rawRFontsXml(r.fontFamily) : '';
  if (fonts) out.push(fonts);
  if (r.bold !== undefined) out.push(toggle('w:b', r.bold));
  if (r.italic !== undefined) out.push(toggle('w:i', r.italic));
  if (r.strike !== undefined) out.push(toggle('w:strike', r.strike));
  if (r.colorHex !== undefined) out.push(`<w:color w:val="${r.colorHex}"/>`);
  if (r.fontSizePt !== undefined) out.push(`<w:sz w:val="${Math.round(r.fontSizePt * 2)}"/>`);
  if (r.underline !== undefined) out.push(underlineXml({ ...r, underline: r.underline }));
  if (r.shadingColorHex !== undefined) out.push(runShdXml(r.shadingColorHex));
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
// then page size/margins, columns, the titlePg toggle and the text direction.
function sectPrXml(s: SectionProperties, hf: HeaderFooterRefs): string {
  const parts: Array<string> = [];
  for (const h of hf.headers) {
    parts.push(`<w:headerReference w:type="${h.type}" r:id="${h.relId}"/>`);
  }
  for (const f of hf.footers) {
    parts.push(`<w:footerReference w:type="${f.type}" r:id="${f.relId}"/>`);
  }
  // §17.6.22 — where the section starts, which Word writes between the
  // references and the page size. Left out, a section that asked for the next
  // odd sheet came back asking for the next sheet.
  if (s.sectionStart !== undefined && s.sectionStart !== 'nextPage') {
    parts.push(`<w:type w:val="${s.sectionStart}"/>`);
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
  // §17.6.12 — how the section numbers its pages, and from what. Left out,
  // front matter numbered i, ii, iii printed 1, 2, 3 and the body went on
  // counting from there instead of starting again at 1.
  if (s.pageNumberFormat !== undefined || s.pageNumberStart !== undefined) {
    const fmt = s.pageNumberFormat !== undefined ? ` w:fmt="${s.pageNumberFormat}"` : '';
    const start = s.pageNumberStart !== undefined ? ` w:start="${String(s.pageNumberStart)}"` : '';
    parts.push(`<w:pgNumType${fmt}${start}/>`);
  }
  if (s.columns) parts.push(colsXml(s.columns));
  if (s.titlePg) parts.push('<w:titlePg/>');
  // §17.6.20 — which way the lines run, after the title-page toggle as
  // CT_SectPr orders them.
  if (s.textDirection) parts.push(`<w:textDirection w:val="${s.textDirection}"/>`);
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

/**
 * XML 1.0 §2.2 — the characters a document may contain at all.
 *
 * Everything below U+0020 except tab, newline and return is FORBIDDEN, and no
 * escape exists for them: `&#2;` is as ill-formed as the byte. A package with
 * one in it is not a document — LibreOffice says "source file could not be
 * loaded" and Word says nothing useful either.
 *
 * They reach here from the PDF reader. A subset font that states neither a
 * `/ToUnicode` nor an `/Encoding /Differences` says nothing about what its
 * codes mean, and the codes of a subset start at 1, 2, 3 — so the last-resort
 * Latin-1 reading turns a page of text into control characters. Reading it
 * better is `font.ts`'s business; what this must guarantee is that nothing the
 * reader believes can produce a package that will not open.
 *
 * A lone surrogate is the same case: a half of a pair is not a character.
 */
/**
 * XML 1.0 §2.2 — whether a code point may appear in a document at all.
 *
 * Everything below U+0020 except tab, newline and return is forbidden, and no
 * escape exists for them: `&#2;` is as ill-formed as the byte itself. So is a
 * lone surrogate — half a pair is not a character — and so are U+FFFE/U+FFFF.
 */
function xmlAllows(cp: number): boolean {
  if (cp === 0x9 || cp === 0xa || cp === 0xd) return true;
  if (cp < 0x20) return false;
  if (cp >= 0xd800 && cp <= 0xdfff) return false;
  if (cp === 0xfffe || cp === 0xffff) return false;
  return cp <= 0x10ffff;
}

/**
 * Escape text for XML, and drop what XML cannot carry.
 *
 * The dropping is not fussiness: a package with one control character in it is
 * not a document. LibreOffice answers "source file could not be loaded" and
 * Word refuses it too, so the whole conversion is lost over one byte.
 *
 * They reach here from the PDF reader. A subset font that states neither a
 * `/ToUnicode` nor an `/Encoding /Differences` says nothing about what its
 * codes mean, and a subset's codes start at 1, 2, 3 — so the last-resort
 * Latin-1 reading turns a page of prose into control characters. Reading such
 * a font better is `font.ts`'s business; what this guarantees is that nothing
 * the reader believes can produce a package that will not open.
 */
function escapeXml(s: string): string {
  let out = '';
  for (const ch of s) {
    const cp = ch.codePointAt(0) ?? 0;
    if (!xmlAllows(cp)) continue;
    out += ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch;
  }
  return out;
}

function escapeAttr(s: string): string {
  return escapeXml(s).replace(/"/g, '&quot;');
}
