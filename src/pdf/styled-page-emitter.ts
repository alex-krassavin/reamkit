// Emit phase: LaidOutDocument → PDF objects (ISO 32000).
//
// The counterpart of layoutStyledDocument (the seam of ir-design §7 /
// oop-design §4.1). Input is the laid-out document plus the output-side
// EmitOptions only — fonts and images embed first (the object order the
// pre-split renderer produced), then pages replay their PageItems, then the
// catalog assembles OutputIntent/XMP/struct-tree/attachments as required.

import type { DocumentInfo } from '@/core/document-model';
import type { ResourceId } from '@/core/ir';
import type { EmbeddedFont } from '@/pdf/cid-font';
import type { PdfDict, PdfRef, PdfValue } from '@/pdf/objects';
import type {
  DrawCommand,
  FontResource,
  ImageItem,
  ImageResource,
  ImageToken,
  LaidOutDocument,
  LaidOutPage,
  Line,
  MathToken,
  PageDimensions,
  PageTagging,
  PdfAProfile,
  SectionRenderCtx,
  StyledRenderOptions,
  TextLineItem,
  TextToken,
} from '@/pdf/styled-page-renderer';
import type { VectorShape } from '@/pdf/vector-graphics';
import type { PdfDocument } from '@/pdf/writer';
import { A4_HEIGHT, A4_WIDTH, paintPlan } from '@/pdf/styled-page-renderer';
import { emitVectorShape } from '@/pdf/vector-graphics';
import { reorderVisual, reverseByCodePoint } from '@/core/bidi';
import { embedTtfFont } from '@/pdf/cid-font';
import { embedAssociatedFile } from '@/pdf/embedded-file';
import { buildSrgbIccProfile } from '@/pdf/icc-profile';
import { addImage } from '@/pdf/image-xobject';
import { dict, name, ref, stream, unicodeString } from '@/pdf/objects';
import { addSignaturePlaceholder } from '@/pdf/signature';
import { buildXmpPacket } from '@/pdf/xmp';

const encoder = new TextEncoder();

// The emit phase sees only the laid-out document plus these output-side
// options — never the layout options (oop-design §4.1: the seam must not leak).
type EmitOptions = Pick<
  StyledRenderOptions,
  'attachments' | 'info' | 'language' | 'signaturePlaceholder'
>;

export function emitStyledPdf(
  laid: LaidOutDocument,
  options: EmitOptions,
  doc: PdfDocument,
): Uint8Array {
  const { pages: renderedPages, structBuilder } = laid;
  const { sectionCtxs, pdfaProfile, tagged } = laid;

  // Create the font/image PDF objects first — the same object order the
  // pre-split renderer produced (fonts, then images, then pages).
  const embeddedFonts = embedFontResources(doc, laid);
  const embeddedImages = embedImageResources(doc, laid);

  const pagesDict: PdfDict = dict({ Type: name('Pages'), Count: 0, Kids: [] });
  const pagesRef = doc.add(pagesDict);
  // First page's dict object, kept mutable so a signature widget can be added to
  // its /Annots after the catalog is assembled.
  let firstPageDict: PdfDict | undefined;
  const fontResourceDict = buildFontResourceDict(laid.fontResources, embeddedFonts);
  const xobjectResourceDict = buildXObjectResourceDict(laid.imageResources, embeddedImages);
  const resourcesDict = dict({
    Font: fontResourceDict,
    ...(xobjectResourceDict ? { XObject: xobjectResourceDict } : {}),
  });

  // PDF/A: build the sRGB ICC stream once — reused by the catalog OutputIntent
  // and, for PDF/A-2/3, the page transparency-group blend colour space.
  let pdfaIccRef: PdfRef | undefined;
  let transparencyGroup: PdfDict | undefined;
  if (pdfaProfile) {
    pdfaIccRef = doc.add(stream({ N: 3, Alternate: name('DeviceRGB') }, buildSrgbIccProfile()));
    if (pdfaProfile.part >= 2) {
      // PDF/A-2/3 §6.2.4.3 — a page that uses transparency (here, a soft-masked
      // image) must carry a transparency group with a device-independent blend
      // colour space; reuse the OutputIntent's ICCBased sRGB.
      transparencyGroup = dict({
        S: name('Transparency'),
        CS: [name('ICCBased'), ref(pdfaIccRef.id)],
      });
    }
  }

  const pageRefs: Array<PdfRef> = [];
  renderedPages.forEach((page, pageIndex) => {
    let pageTagging: PageTagging | undefined;
    if (structBuilder) {
      const builder = structBuilder;
      pageTagging = {
        next: 0,
        assigned: false,
        record: (structId, mcid) => builder.addMcref(structId, pageIndex, mcid),
        tagFor: (structId) => builder.node(structId).type,
      };
    }
    const contentBytes = emitPageContent(page, pageTagging);
    const contentsRef = doc.add(stream({}, contentBytes));
    const pageEntries: Record<string, PdfValue> = {
      Type: name('Page'),
      Parent: ref(pagesRef.id),
      MediaBox: [0, 0, page.width, page.height],
      Resources: resourcesDict,
      Contents: ref(contentsRef.id),
    };
    if (transparencyGroup) pageEntries['Group'] = transparencyGroup;
    if (pageTagging?.assigned) {
      // §14.7.4.4 — this page's key into /ParentTree; §14.8.4.2 — structure tab
      // order so AT navigates in logical, not geometric, order.
      pageEntries['StructParents'] = pageIndex;
      pageEntries['Tabs'] = name('S');
    }
    const pageDictObj = dict(pageEntries);
    if (firstPageDict === undefined) firstPageDict = pageDictObj;
    pageRefs.push(doc.add(pageDictObj));
  });

  if (pageRefs.length === 0) {
    // Always emit at least one page (fallback geometry: first section's dims
    // or A4 if no sections at all).
    const fallback = sectionCtxs[0] ?? defaultPageCtx();
    const contentsRef = doc.add(stream({}, new Uint8Array(0)));
    const fallbackDict = dict({
      Type: name('Page'),
      Parent: ref(pagesRef.id),
      MediaBox: [0, 0, fallback.pageWidth, fallback.pageHeight],
      Resources: resourcesDict,
      Contents: ref(contentsRef.id),
      ...(transparencyGroup ? { Group: transparencyGroup } : {}),
    });
    firstPageDict = fallbackDict;
    pageRefs.push(doc.add(fallbackDict));
  }

  pagesDict.set('Count', pageRefs.length);
  pagesDict.set('Kids', pageRefs);

  const catalogEntries: Record<string, PdfValue> = {
    Type: name('Catalog'),
    Pages: ref(pagesRef.id),
  };
  if (pdfaProfile) {
    // OutputIntent with the embedded sRGB ICC profile (PDF/A §6.2.2). The same
    // ICC stream (pdfaIccRef, built above) backs any page transparency group.
    const outputIntentRef = doc.add(
      dict({
        Type: name('OutputIntent'),
        S: name('GTS_PDFA1'),
        OutputConditionIdentifier: 'sRGB',
        Info: 'sRGB IEC61966-2.1',
        DestOutputProfile: ref(pdfaIccRef!.id),
      }),
    );
    catalogEntries['OutputIntents'] = [ref(outputIntentRef.id)];

    // Document-level XMP metadata carrying the PDF/A identifier (§6.7).
    const xmpRef = doc.add(
      stream(
        { Type: name('Metadata'), Subtype: name('XML') },
        buildXmpPacket(xmpFromInfo(options.info, pdfaProfile)),
      ),
    );
    catalogEntries['Metadata'] = ref(xmpRef.id);
  }
  if (tagged && structBuilder) {
    // Logical structure (§14.8): emit the tree now that every page ref exists,
    // then point the catalog at it and mark the document as tagged.
    const structRootRef = structBuilder.emit(doc, pageRefs);
    catalogEntries['MarkInfo'] = dict({ Marked: true });
    catalogEntries['StructTreeRoot'] = ref(structRootRef.id);
    // §14.9.2 — document natural language (from the docx default, else en-US).
    catalogEntries['Lang'] = options.language ?? 'en-US';
    // DisplayDocTitle — AT should announce the document title, not the file name
    // (Matterhorn 06-003 / PDF/UA). Only meaningful when a title is present.
    if (options.info?.title) {
      catalogEntries['ViewerPreferences'] = dict({ DisplayDocTitle: true });
    }
  }
  // Associated files (§7.11 / PDF/A-3 §6.8) — allowed for plain PDF and PDF/A-3
  // only; PDF/A-1/2 forbid arbitrary embedded files, so skip them there.
  if (options.attachments?.length && (!pdfaProfile || pdfaProfile.part === 3)) {
    const embedded = options.attachments.map((file) => ({
      file,
      ref: embedAssociatedFile(doc, file),
    }));
    catalogEntries['AF'] = embedded.map((e) => ref(e.ref.id));
    // /Names /EmbeddedFiles name tree, keyed by file name (sorted → deterministic).
    const sorted = [...embedded].sort((a, b) =>
      a.file.name < b.file.name ? -1 : a.file.name > b.file.name ? 1 : 0,
    );
    const names: Array<PdfValue> = [];
    for (const e of sorted) names.push(e.file.name, ref(e.ref.id));
    const embeddedFilesRef = doc.add(dict({ Names: names }));
    catalogEntries['Names'] = dict({ EmbeddedFiles: ref(embeddedFilesRef.id) });
  }
  // Digital signature placeholder (§12.8): an invisible /Sig widget on page 1,
  // an /AcroForm in the catalog, and a signature dict with placeholder
  // ByteRange/Contents that signPdf() fills. The unsigned placeholder is emitted
  // here; the crypto happens afterwards (async) so the writer stays sync.
  if (options.signaturePlaceholder && firstPageDict) {
    const { fieldRef, acroForm } = addSignaturePlaceholder(
      doc,
      options.signaturePlaceholder,
      pageRefs[0]!,
    );
    firstPageDict.set('Annots', [ref(fieldRef.id)]);
    catalogEntries['AcroForm'] = acroForm;
  }

  const catalogRef = doc.add(dict(catalogEntries));
  const infoRef = buildInfoDict(doc, options.info);
  return doc.build(catalogRef, infoRef, {
    ...(pdfaProfile ? { version: pdfaProfile.version, id: true } : {}),
  });
}

const DEFAULT_PRODUCER = 'Ream';

function xmpFromInfo(
  info: DocumentInfo | undefined,
  p: PdfAProfile | undefined,
): Parameters<typeof buildXmpPacket>[0] {
  return {
    pdfaPart: p ? (String(p.part) as '1' | '2' | '3') : '1',
    pdfaConformance: p ? (p.level.toUpperCase() as 'A' | 'B' | 'U') : 'B',
    producer: info?.producer ?? DEFAULT_PRODUCER,
    ...(info?.title ? { title: info.title } : {}),
    ...(info?.author ? { author: info.author } : {}),
    ...(info?.subject ? { subject: info.subject } : {}),
    ...(info?.keywords ? { keywords: info.keywords } : {}),
    ...(info?.creator ? { creator: info.creator } : {}),
    ...(info?.creationDate ? { createDate: info.creationDate } : {}),
    ...(info?.modificationDate ? { modifyDate: info.modificationDate } : {}),
  };
}

function buildInfoDict(doc: PdfDocument, info: DocumentInfo | undefined): PdfRef | undefined {
  const merged: DocumentInfo = {
    ...info,
    producer: info?.producer ?? DEFAULT_PRODUCER,
  };
  const entries: Record<string, PdfValue> = {};
  if (merged.title) entries['Title'] = unicodeString(merged.title);
  if (merged.author) entries['Author'] = unicodeString(merged.author);
  if (merged.subject) entries['Subject'] = unicodeString(merged.subject);
  if (merged.keywords) entries['Keywords'] = unicodeString(merged.keywords);
  if (merged.creator) entries['Creator'] = unicodeString(merged.creator);
  // Producer always emitted (defaulted above).
  entries['Producer'] = unicodeString(merged.producer!);
  if (merged.creationDate) entries['CreationDate'] = formatPdfDate(merged.creationDate);
  if (merged.modificationDate) entries['ModDate'] = formatPdfDate(merged.modificationDate);
  return doc.add(dict(entries));
}

// ISO 32000-1 §7.9.4 — PDF date string `D:YYYYMMDDHHmmSSOHH'mm'` where O is
// +/-/Z. We always emit UTC ("Z") so the formatter stays deterministic and
// the output is timezone-independent.
function formatPdfDate(d: Date): string {
  const pad = (n: number, w = 2) => String(n).padStart(w, '0');
  const yyyy = pad(d.getUTCFullYear(), 4);
  const MM = pad(d.getUTCMonth() + 1);
  const DD = pad(d.getUTCDate());
  const hh = pad(d.getUTCHours());
  const mm = pad(d.getUTCMinutes());
  const ss = pad(d.getUTCSeconds());
  return `D:${yyyy}${MM}${DD}${hh}${mm}${ss}Z`;
}

function defaultPageCtx(): SectionRenderCtx {
  return {
    endIndex: 0,
    properties: { headers: [], footers: [] },
    pageWidth: A4_WIDTH,
    pageHeight: A4_HEIGHT,
    marginLeft: 72,
    marginTop: 72,
    marginBottom: 72,
    contentWidth: A4_WIDTH - 144,
    pageContentHeight: A4_HEIGHT - 144,
    headerSet: { default: [], first: [], even: [] },
    footerSet: { default: [], first: [], even: [] },
    titlePg: false,
    evenAndOddHeaders: false,
  };
}

// Emit-phase counterpart: embed the probed images in collection order.
function embedImageResources(doc: PdfDocument, laid: LaidOutDocument): Map<ResourceId, PdfRef> {
  const out = new Map<ResourceId, PdfRef>();
  for (const [resourceId, res] of laid.imageResources) {
    out.set(resourceId, addImage(doc, res.prepared).ref);
  }
  return out;
}

// Emit-phase counterpart: create the PDF font objects (subset to the collected
// glyphs) for every laid-out font resource, in collection order — keeping the
// object numbering identical to the pre-split renderer.
function embedFontResources(doc: PdfDocument, laid: LaidOutDocument): Map<string, EmbeddedFont> {
  // PDF/A-1 requires a /CIDSet; PDF/A-2/3 and non-PDF/A omit it.
  const cidSet = laid.pdfaProfile?.part === 1;
  const out = new Map<string, EmbeddedFont>();
  for (const [variant, res] of laid.fontResources) {
    out.set(variant, embedTtfFont(doc, res.parsed, { usedGids: res.gids, cidSet }));
  }
  return out;
}

function buildFontResourceDict(
  resources: ReadonlyMap<string, FontResource>,
  embedded: ReadonlyMap<string, EmbeddedFont>,
): PdfDict {
  const entries: Record<string, PdfRef> = {};
  for (const [variant, res] of resources) {
    const emb = embedded.get(variant);
    if (emb) entries[res.resourceName] = ref(emb.fontRef.id);
  }
  return dict(entries);
}

function buildXObjectResourceDict(
  resources: ReadonlyMap<ResourceId, ImageResource>,
  embedded: ReadonlyMap<ResourceId, PdfRef>,
): PdfDict | undefined {
  if (embedded.size === 0) return undefined;
  const entries: Record<string, PdfRef> = {};
  for (const [resourceId, res] of resources) {
    const r = embedded.get(resourceId);
    if (r) entries[res.resourceName] = ref(r.id);
  }
  return dict(entries);
}

// Emit a run of same-typed draw commands (the image or shape pass), opening one
// marked-content sequence per contiguous group with the same owner: a structId
// → its struct type (Figure) with a fresh MCID; a pagination artifact; or a bare
// layout artifact. `body(cmd)` emits each command's own operators. In non-tagged
// mode it just emits the bodies — byte-identical to before tagging existed.
function emitTaggedRuns<T extends DrawCommand>(
  out: Array<string>,
  cmds: ReadonlyArray<T>,
  tagging: PageTagging | undefined,
  body: (cmd: T) => void,
): void {
  if (!tagging) {
    for (const c of cmds) body(c);
    return;
  }
  let openKey: string | null = null;
  const close = () => {
    if (openKey !== null) {
      out.push('EMC');
      openKey = null;
    }
  };
  for (const c of cmds) {
    const key =
      c.structId !== undefined ? `f${c.structId}` : c.artifact === 'pagination' ? 'p' : 'a';
    if (key !== openKey) {
      close();
      if (c.structId !== undefined) {
        const mcid = tagging.next++;
        tagging.assigned = true;
        tagging.record(c.structId, mcid);
        out.push(`/${tagging.tagFor(c.structId)} <</MCID ${mcid}>> BDC`);
      } else if (c.artifact === 'pagination') {
        out.push('/Artifact <</Type /Pagination>> BDC');
      } else {
        out.push('/Artifact BMC');
      }
      openKey = key;
    }
    body(c);
  }
  close();
}

function emitPageContent(page: LaidOutPage, tagging?: PageTagging): Uint8Array {
  const commands = page.commands;
  if (commands.length === 0) return new Uint8Array(0);
  // PageDoc coordinates are top-left/y-down (the frozen schema); PDF paints in
  // a y-up bottom-left frame, so every page-frame y converts as `H - y` here.
  const H = page.height;
  const out: Array<string> = [];

  // Cell backgrounds (fills) first — they sit underneath text and borders.
  // In tagged PDFs they are layout decoration → wrapped as an /Artifact so no
  // page content sits outside the structure tree (PDF/A-1a §6.3.2).
  const plan = paintPlan(commands);
  const fills = plan.fills;
  if (fills.length > 0) {
    if (tagging) out.push('/Artifact BMC');
    out.push('q');
    let lastColor = '';
    for (const f of fills) {
      const color = f.fillColorHex;
      if (color !== lastColor) {
        const [r, g, b] = hexToRgb01(color);
        out.push(`${formatNumber(r)} ${formatNumber(g)} ${formatNumber(b)} rg`);
        lastColor = color;
      }
      out.push(
        `${formatNumber(f.x)} ${formatNumber(H - f.y - f.height)} ${formatNumber(f.width)} ${formatNumber(f.height)} re`,
      );
      out.push('f');
    }
    out.push('Q');
    if (tagging) out.push('EMC');
  }

  // Block images. In tagged mode each carries its owning Figure (structId) → a
  // /Figure marked-content sequence; header/footer images are pagination
  // artifacts; anything else a bare artifact.
  // Drop images whose resource failed to embed (unsupported/corrupt) — they have
  // no XObject, so a `/<name> Do` would be a dangling reference.
  const images = plan.images.filter((c) => c.imageResourceName !== '');
  emitTaggedRuns(out, images, tagging, (img) => {
    // ISO 32000-1 §8.9.5.1 — Image XObject placement.
    // q             save graphics state
    // w 0 0 h x y cm   scale unit square to (w,h) and translate to (x,y)
    // /Im1 Do       paint the XObject
    // Q             restore
    out.push('q');
    out.push(
      `${formatNumber(img.width)} 0 0 ${formatNumber(img.height)} ${formatNumber(img.x)} ${formatNumber(H - img.y - img.height)} cm`,
    );
    out.push(`/${img.imageResourceName} Do`);
    out.push('Q');
  });

  const borders = plan.borders;
  if (borders.length > 0) {
    if (tagging) out.push('/Artifact BMC');
    out.push('q');
    let lastWidth = -1;
    let lastColor = '';
    for (const b of borders) {
      const width = b.borderSizePt;
      const color = b.borderColorHex;
      if (width !== lastWidth) {
        out.push(`${formatNumber(width)} w`);
        lastWidth = width;
      }
      if (color !== lastColor) {
        const [r, g, bl] = hexToRgb01(color);
        out.push(`${formatNumber(r)} ${formatNumber(g)} ${formatNumber(bl)} RG`);
        lastColor = color;
      }
      const x = b.x;
      const y = H - b.y - b.height; // box bottom edge in PDF's y-up frame
      const w = b.width;
      const h = b.height;
      switch (b.side) {
        case 'top':
          out.push(`${formatNumber(x)} ${formatNumber(y + h)} m`);
          out.push(`${formatNumber(x + w)} ${formatNumber(y + h)} l`);
          break;
        case 'bottom':
          out.push(`${formatNumber(x)} ${formatNumber(y)} m`);
          out.push(`${formatNumber(x + w)} ${formatNumber(y)} l`);
          break;
        case 'left':
          out.push(`${formatNumber(x)} ${formatNumber(y)} m`);
          out.push(`${formatNumber(x)} ${formatNumber(y + h)} l`);
          break;
        case 'right':
          out.push(`${formatNumber(x + w)} ${formatNumber(y)} m`);
          out.push(`${formatNumber(x + w)} ${formatNumber(y + h)} l`);
          break;
      }
      out.push('S');
    }
    out.push('Q');
    if (tagging) out.push('EMC');
  }

  // Vector shapes (DrawingML §20). Each shape is a self-contained q…cm…Q block
  // (emitVectorShape). Drawn after table fills/images/borders but before the
  // text pass, so both body text and the shape's own text (emitted as 'line'
  // commands) land on top of shape fills. In tagged mode a shape/chart carries
  // its Figure structId → /Figure marked content (contiguous chart shapes share
  // one MCID); decorative shapes fall back to a bare artifact.
  const shapes = plan.shapes;
  emitTaggedRuns(out, shapes, tagging, (sh) => {
    // The stored transform targets the top-left page frame; conjugating with
    // the page flip (an involution — same operation layout applied) recovers
    // the y-up CTM. The linear part negates exactly; only f re-rounds.
    const t = sh.shape.transform;
    const shape: VectorShape = {
      ...sh.shape,
      transform: [t[0], -t[1], t[2], -t[3], t[4], H - t[5]],
    };
    for (const op of emitVectorShape(shape)) out.push(op);
  });

  const lines = plan.lines;
  if (lines.length > 0) {
    let inBT = false;
    let lastFont = '';
    let lastSize = -1;
    let lastColor = '';
    const switchFontIfNeeded = (tok: TextToken) => {
      const fontKey = tok.font.resourceName;
      if (fontKey !== lastFont || tok.fontSizePt !== lastSize) {
        out.push(`/${fontKey} ${formatNumber(tok.fontSizePt)} Tf`);
        lastFont = fontKey;
        lastSize = tok.fontSizePt;
      }
      if (tok.resolvedRun.colorHex !== lastColor) {
        const [r, g, b] = hexToRgb01(tok.resolvedRun.colorHex);
        out.push(`${formatNumber(r)} ${formatNumber(g)} ${formatNumber(b)} rg`);
        lastColor = tok.resolvedRun.colorHex;
      }
    };
    const emitImageToken = (tok: ImageToken, x: number, baselineY: number) => {
      // Skip an inline image whose resource failed to embed (the caller still
      // advances x, so its box stays reserved).
      if (!tok.imageResourceName) return;
      // ET out of text mode, place the image XObject, then BT back in. Image
      // bottom-left sits on the text baseline so the image hangs above the
      // baseline like an inline glyph.
      if (inBT) {
        out.push('ET');
        inBT = false;
      }
      out.push('q');
      out.push(
        `${formatNumber(tok.widthPt)} 0 0 ${formatNumber(tok.heightPt)} ${formatNumber(x)} ${formatNumber(baselineY)} cm`,
      );
      out.push(`/${tok.imageResourceName} Do`);
      out.push('Q');
      // Text state is reset by ET; force re-emit on the next text token.
      lastFont = '';
      lastSize = -1;
      lastColor = '';
    };
    // Emit an inline math box: glyph items in text mode, rule/path items in
    // graphics mode. All positions are box-local, offset by the box origin.
    const emitMathToken = (tok: MathToken, originX: number, baselineY: number) => {
      for (const it of tok.items) {
        if (it.kind === 'glyph') {
          if (!inBT) {
            out.push('BT');
            inBT = true;
          }
          if (it.font.resourceName !== lastFont || it.sizePt !== lastSize) {
            out.push(`/${it.font.resourceName} ${formatNumber(it.sizePt)} Tf`);
            lastFont = it.font.resourceName;
            lastSize = it.sizePt;
          }
          if (lastColor !== '000000') {
            out.push('0 0 0 rg');
            lastColor = '000000';
          }
          out.push(`1 0 0 1 ${formatNumber(originX + it.x)} ${formatNumber(baselineY + it.y)} Tm`);
          out.push(`<${it.font.measure.encodeTextAsCidHex(it.text)}> Tj`);
        } else if (it.kind === 'rule') {
          if (inBT) {
            out.push('ET');
            inBT = false;
          }
          out.push('q');
          out.push('0 0 0 rg');
          out.push(
            `${formatNumber(originX + it.x)} ${formatNumber(baselineY + it.y)} ${formatNumber(it.w)} ${formatNumber(it.h)} re`,
          );
          out.push('f');
          out.push('Q');
          lastFont = '';
          lastSize = -1;
          lastColor = '';
        } else {
          if (inBT) {
            out.push('ET');
            inBT = false;
          }
          const shape: VectorShape = {
            paths: [{ segments: it.segments }],
            ...(it.fill ? { fillColorHex: '000000' } : {}),
            ...(it.strokeWidthPt !== undefined
              ? { stroke: { colorHex: '000000', widthPt: it.strokeWidthPt } }
              : {}),
            transform: [1, 0, 0, 1, originX, baselineY],
          };
          for (const op of emitVectorShape(shape)) out.push(op);
          lastFont = '';
          lastSize = -1;
          lastColor = '';
        }
      }
    };

    // Emit one line command's glyphs/images/math. Manages BT/ET through the
    // shared `inBT` state; produces operator-for-operator the same output as
    // before tagging existed, so the non-tagged path stays byte-identical.
    const emitOneLine = (cmd: TextLineItem) => {
      const line = cmd.line;
      const originX = cmd.originX;
      const baselineY = H - cmd.baselineY; // top-left frame → PDF y-up
      const extraPerSpace = computeJustifyExtra(line);
      const hasImageToken = line.tokens.some((t) => t.kind === 'image');
      const hasMathToken = line.tokens.some((t) => t.kind === 'math');
      const hasRtl = line.tokens.some((t) => t.bidiLevel % 2 === 1);

      // Encode a token's text, reversing code points for RTL (odd-level) runs
      // so glyphs lay out right-to-left as our cursor advances left-to-right.
      const encodeToken = (tok: TextToken): string => {
        const text = tok.bidiLevel % 2 === 1 ? reverseByCodePoint(tok.text) : tok.text;
        return tok.font.measure.encodeTextAsCidHex(text);
      };

      if (extraPerSpace > 0 || hasImageToken || hasMathToken || hasRtl) {
        // Per-token absolute positioning. Required for justify (inter-word
        // slack), inline images (text-mode exits), and BiDi (visual order
        // differs from logical order). Tokens are emitted in visual order.
        const order = hasRtl ? lineVisualOrder(line) : line.tokens.map((_, i) => i);
        let x = originX;
        for (const ti of order) {
          const tok = line.tokens[ti]!;
          if (tok.kind === 'image') {
            emitImageToken(tok, x, baselineY);
            x += tok.widthPt;
            continue;
          }
          if (tok.kind === 'math') {
            emitMathToken(tok, x, baselineY);
            x += tok.widthPt;
            continue;
          }
          if (!inBT) {
            out.push('BT');
            inBT = true;
          }
          switchFontIfNeeded(tok);
          out.push(`1 0 0 1 ${formatNumber(x)} ${formatNumber(baselineY)} Tm`);
          out.push(`<${encodeToken(tok)}> Tj`);
          x += tok.widthPt;
          if (tok.isSpace) x += extraPerSpace;
        }
      } else {
        if (!inBT) {
          out.push('BT');
          inBT = true;
        }
        out.push(`1 0 0 1 ${formatNumber(originX)} ${formatNumber(baselineY)} Tm`);
        for (const tok of line.tokens) {
          if (tok.kind !== 'text') continue; // unreachable here, but TS-narrowed
          switchFontIfNeeded(tok);
          const hex = tok.font.measure.encodeTextAsCidHex(tok.text);
          out.push(`<${hex}> Tj`);
        }
      }
    };

    if (tagging) {
      // Tagged PDF: each line is its own marked-content sequence. A body line
      // (carrying a structId) becomes /P <</MCID n>> BDC … EMC and registers its
      // MCID with the owning structure node; a line without a structId (header/
      // footer text) is a pagination artifact. The BDC/EMC bracket cleanly wraps
      // the line's own BT…ET (and any inline-image/math q…Q), which is legal —
      // marked-content brackets nest independently of BT/ET and q/Q (§14.6.1).
      for (const cmd of lines) {
        // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition -- inBT is toggled inside the emit* closures; the type-checker's flow analysis can't see those mutations and treats it as constant.
        if (inBT) {
          out.push('ET');
          inBT = false;
        }
        const sid = cmd.structId;
        if (sid !== undefined) {
          const mcid = tagging.next++;
          tagging.assigned = true;
          tagging.record(sid, mcid);
          out.push(`/${tagging.tagFor(sid)} <</MCID ${mcid}>> BDC`);
        } else if (cmd.artifact === 'pagination') {
          out.push('/Artifact <</Type /Pagination>> BDC');
        } else {
          out.push('/Artifact BMC');
        }
        emitOneLine(cmd);
        // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition -- emitOneLine may leave text mode open (inBT true) via a closure; flow analysis can't track it.
        if (inBT) {
          out.push('ET');
          inBT = false;
        }
        out.push('EMC');
      }
    } else {
      out.push('BT');
      inBT = true;
      for (const cmd of lines) emitOneLine(cmd);
      // Ensure we exit text mode if the last line ended on an image token.
      // eslint-disable-next-line @typescript-eslint/no-unnecessary-condition -- emitOneLine toggles inBT via a closure; flow analysis can't see it and assumes it stays true.
      if (inBT) out.push('ET');
    }
  }

  return encoder.encode(out.join('\n'));
}

// UAX #9 rule L2 over a line's tokens: returns token indices in visual order.
// Each token carries a single embedding level (the tokenizer split runs at
// level boundaries), so reordering tokens is equivalent to reordering their
// constituent characters.
function lineVisualOrder(line: Line): Array<number> {
  return reorderVisual(line.tokens.map((t) => t.bidiLevel));
}

// Extra width per space token when justifying. 0 for non-justify lines or
// the last line of a paragraph (last line stays left-aligned by convention).
function computeJustifyExtra(line: Line): number {
  if (line.resolved.alignment !== 'both') return 0;
  if (line.isLastInParagraph) return 0;
  let spaces = 0;
  for (const tok of line.tokens) if (tok.isSpace) spaces++;
  if (spaces === 0) return 0;
  const extra = line.availableWidthPt - line.contentWidthPt;
  if (extra <= 0) return 0;
  return extra / spaces;
}

function hexToRgb01(hex: string): readonly [number, number, number] {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  return [r, g, b];
}

function formatNumber(n: number): string {
  if (Number.isInteger(n)) return String(n);
  return n.toFixed(3).replace(/0+$/, '').replace(/\.$/, '');
}
