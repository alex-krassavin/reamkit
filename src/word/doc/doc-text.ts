// Legacy `.doc` text + formatting (DOC-1..4) — pulls the document out of a Word
// 97–2003 binary file as paragraphs of formatted runs. A `.doc` is a CFB (the
// same OLE2 container as `.xls`) holding a `WordDocument` stream and a table
// stream (`0Table` / `1Table`). Three separate structures have to be joined,
// all keyed by file character position (FC):
//   - the piece table (CLX) maps character positions (CPs) to FCs and tells
//     whether each piece is 16-bit Unicode or 8-bit Windows-1252 ("compressed");
//   - CHPX runs (in FKP pages, found through PlcBteChpx) carry run formatting;
//   - PAPX runs (in FKP pages, found through PlcBtePapx) carry paragraph
//     formatting, one run per paragraph.
// This walks the text character by character, attaches each character's CHPX,
// splits paragraphs at the CR mark (and the 0x07 table cell mark, and a section's
// closing mark), attaches each paragraph's PAPX, and emits maximal same-formatting
// runs.
//
// Spec: [MS-DOC] §2.5 (FIB), §2.8.35 (PlcPcd), §2.9.38 (Clx), §2.9.73
// (FcCompressed), §2.8.6 (PlcBteChpx), §2.8.5 (PlcBtePapx), §2.9.55 (ChpxFkp),
// §2.9.23 (PapxFkp), §2.8.26 (PlcfSed), §2.6.1 (sprm). Every offset is bounds-checked; a
// structurally odd file yields no text rather than throwing.

import { openCfb } from '@/core/ole/cfb';

const WIDENT = 0xa5ec; // FibBase.wIdent — the "this is a Word doc" magic
const OFF_FLAGS1 = 0x0a; // FibBase bit field (fEncrypted, fWhichTblStm, …)
const OFF_CCPTEXT = 0x4c; // fibRgLw97.ccpText — main-document char count
const OFF_CCPFTN = 0x50; // fibRgLw97.ccpFtn — footnote story char count
const OFF_CCPHDD = 0x54; // fibRgLw97.ccpHdd — header/footer story char count
const OFF_FCPLCFHDD = 0xf2; // fibRgFcLcb97.fcPlcfHdd — header/footer story divisions
const OFF_LCBPLCFHDD = 0xf6;
const OFF_FCDOP = 0x192; // fibRgFcLcb97.fcDop — the document properties (Dop)
const OFF_LCBDOP = 0x196;
const OFF_FCSTSHF = 0xa2; // fibRgFcLcb97.fcStshf — the style sheet (STSH)
const OFF_LCBSTSHF = 0xa6;
const OFF_FCPLCFBTECHPX = 0xfa; // fibRgFcLcb97.fcPlcfBteChpx — CHPX bin table
const OFF_LCBPLCFBTECHPX = 0xfe;
const OFF_FCPLCFBTEPAPX = 0x102; // fibRgFcLcb97.fcPlcfBtePapx — PAPX bin table
const OFF_LCBPLCFBTEPAPX = 0x106;
const OFF_FCCLX = 0x1a2; // fibRgFcLcb97.fcClx — CLX offset in the table stream
const OFF_LCBCLX = 0x1a6; // fibRgFcLcb97.lcbClx — CLX byte length
const OFF_FCPLCFSED = 0xca; // fibRgFcLcb97.fcPlcfSed — section descriptors (PlcfSed)
const OFF_LCBPLCFSED = 0xce;
const OFF_FCPLFLST = 0x2e2; // fibRgFcLcb97.fcPlfLst — the list table (PlfLst)
const OFF_LCBPLFLST = 0x2e6; // (excludes the appended LVL array)
const OFF_FCPLFLFO = 0x2ea; // fibRgFcLcb97.fcPlfLfo — the list-format-override table
const OFF_LCBPLFLFO = 0x2ee;
const FLAG_ENCRYPTED = 0x0100; // FibBase.fEncrypted
const FLAG_WHICH_TBL = 0x0200; // FibBase.fWhichTblStm — 1 ⇒ 1Table, 0 ⇒ 0Table

const FC_COMPRESSED = 0x40000000; // FcCompressed.fCompressed — 8-bit piece
const FC_MASK = 0x3fffffff; // FcCompressed.fc — the 30-bit offset field

const FKP_SIZE = 512; // an FKP is one 512-byte page
const PN_MASK = 0x3fffff; // PnFkp.pn — the low 22 bits address the page
const CR = 0x0d; // paragraph mark

// sprm opcodes (§2.6.1). Operand size is `SPRA_LEN[sprm >> 13]`; spra 6 is
// variable (length-prefixed); other sprms are skipped by that size.
const SPRM_C_FBOLD = 0x0835;
const SPRM_C_FITALIC = 0x0836;
const SPRM_C_KUL = 0x2a3e; // underline kind
const SPRM_C_HPS = 0x4a43; // font size, half-points
const SPRM_C_FSPEC = 0x0855; // fSpec — special character (e.g. a picture placeholder)
const SPRM_C_PIC_LOCATION = 0x6a03; // picture data offset into the Data stream
const SPRM_C_ISTD = 0x4a30; // the run's character style
const PIC_CHAR = 0x01; // the special character a picture occupies in the text
const SPRM_P_JC = 0x2403; // paragraph justification
const SPRM_P_DXA_RIGHT = 0x840e; // right indent (twips)
const SPRM_P_DXA_LEFT = 0x840f; // left indent (twips)
const SPRM_P_DXA_LEFT1 = 0x8411; // first-line indent (twips, signed)
const SPRM_P_DYA_BEFORE = 0xa413; // space before (twips)
const SPRM_P_DYA_AFTER = 0xa414; // space after (twips)
const SPRM_P_F_IN_TABLE = 0x2416; // paragraph is in a table
const SPRM_P_F_TTP = 0x2417; // table terminating paragraph (end of row)
const SPRM_P_F_WIDOW_CONTROL = 0x2431; // widow and orphan control
const SPRM_P_F_KEEP = 0x2405; // keep the paragraph's lines together
const SPRM_P_F_KEEP_FOLLOW = 0x2406; // keep the paragraph with the next
const SPRM_P_F_PAGE_BREAK_BEFORE = 0x2407; // the paragraph starts a new page
const SPRM_T_DEF_TABLE = 0xd608; // table definition (cell boundaries + descriptors)
const SPRM_T_DEF_TABLE_SHD = 0xd612; // per-cell shading array (Shd: cvFore/cvBack/ipat)
const MAX_ROW_CELLS = 256; // a row's cell count is small; guard a crafted huge array
const SPRM_P_ILVL = 0x260a; // list level (0–8)
const SPRM_P_ILFO = 0x460b; // list-format override index (>0 ⇒ the paragraph is a list item)
const SPRM_S_BKC = 0x3009; // where the section starts (bkc)
const SPRM_S_F_TITLE_PAGE = 0x300a; // the section's first page has its own header/footer
const SPRM_S_XA_PAGE = 0xb01f; // section page width in twips (xaPage)
const SPRM_S_YA_PAGE = 0xb020; // section page height in twips (yaPage)
const SPRM_S_DXA_LEFT = 0xb021; // left margin (twips)
const SPRM_S_DXA_RIGHT = 0xb022; // right margin (twips)
const SPRM_S_DYA_TOP = 0x9023; // top margin (twips, signed: < 0 ⇒ exactly that)
const SPRM_S_DYA_BOTTOM = 0x9024; // bottom margin (twips, signed likewise)
const SPRM_S_DYA_HDR_TOP = 0xb017; // header's distance from the top edge (twips)
const SPRM_S_DYA_HDR_BOTTOM = 0xb018; // footer's distance from the bottom edge (twips)
const SPRM_S_DZA_GUTTER = 0xb025; // binding space beside the left margin (twips)
const MAX_PAGE_TWIPS = 31680; // Word's max page dimension (22 inches) — sanity bound
const SPRA_LEN = [1, 1, 2, 4, 2, 2, 0, 3];
// spra-6 sprms whose operand is prefixed by a 2-byte (not 1-byte) length.
const SPRM_LONG_OPERAND = new Set([0xd608, 0xc60d]); // sprmTDefTable, sprmPChgTabs

const CELL_MARK = 0x07; // ends a table cell (and, on the TTP, a row)
const PAGE_BREAK = 0x0c; // a manual page break — or, closing a section, its section mark
const COLUMN_BREAK = 0x0e; // a manual column break
const LINE_BREAK = 0x0b; // a manual line break, within the paragraph
const FIELD_BEGIN = 0x13; // a field begins; its code (suppressed) follows
const FIELD_SEPARATOR = 0x14; // the code ends; the cached result (kept) follows
const FIELD_END = 0x15; // the field ends

const MAX_TEXT = 1 << 24; // 16M-char guard against a crafted piece table

const STK_PARAGRAPH = 1; // a paragraph style (Stdf.stk)
const STK_CHARACTER = 2; // a character style
const ISTD_NIL = 0x0fff; // no style: the istdBase of a style based on none
const MAX_STYLE_DEPTH = 16; // a chain of bases longer than this is a loop

/** Run (character) formatting decoded from a CHPX grpprl. */
export interface DocCharProps {
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** Word `kul` underline code (0 = none). */
  readonly underlineKul?: number;
  /** Font size in half-points. */
  readonly sizeHalfPts?: number;
  /** The character is special — `sprmCFSpec` (e.g. a picture placeholder). */
  readonly fSpec?: boolean;
  /** `sprmCPicLocation` — offset into the `Data` stream of the picture's PICF. */
  readonly picOffset?: number;
}

/** An embedded picture: the extracted image bytes and its display size (twips). */
export interface DocPicture {
  readonly bytes: Uint8Array;
  readonly widthTwips: number;
  readonly heightTwips: number;
}

/**
 * One cell border (DOC-11) from a Brc80: its width in points, colour and whether
 * it is a double line. Absent when the cell edge has no border.
 */
export interface DocBorder {
  readonly widthPt: number;
  readonly colorHex: string;
  /** The border is a double line (`brcType` 0x03). */
  readonly double?: boolean;
}

/**
 * A table cell's TC80 descriptor (DOC-11): its four edge borders and its
 * vertical-merge marker (`restart` = the top of a vertical span, `continue` = a
 * cell merged into the one above). Carried per cell on the row's TTP.
 */
export interface DocTc {
  readonly borders?: {
    readonly top?: DocBorder;
    readonly left?: DocBorder;
    readonly bottom?: DocBorder;
    readonly right?: DocBorder;
  };
  readonly vMerge?: 'restart' | 'continue';
}

/** Paragraph formatting decoded from a PAPX grpprl. */
export interface DocParaProps {
  /** Word justification (0 left, 1 center, 2 right, 3 both, 4 distribute). */
  readonly jc?: number;
  readonly indentLeftTwips?: number;
  readonly indentRightTwips?: number;
  /** First-line indent, signed (negative = hanging). */
  readonly indentFirstTwips?: number;
  readonly spaceBeforeTwips?: number;
  readonly spaceAfterTwips?: number;
  /** The paragraph is in a table — `sprmPFInTable`. */
  readonly inTable?: boolean;
  /** This is the row's terminating paragraph (TTP) — `sprmPFTtp`. */
  readonly rowEnd?: boolean;
  /**
   * `sprmTDefTable` rgdxaCenter — the (cols+1) cell-boundary x-positions (twips);
   * present on the TTP. Cell `i` width = `edge[i+1] − edge[i]`.
   */
  readonly cellEdgesTwips?: ReadonlyArray<number>;
  /**
   * `sprmTDefTable` TC80 per-cell descriptors (borders + vertical merge), one per
   * cell, present on the TTP alongside `cellEdgesTwips` (DOC-11).
   */
  readonly cellDescriptors?: ReadonlyArray<DocTc>;
  /**
   * `sprmTDefTableShd` — each cell's background fill hex (cvBack), or `undefined`
   * for an auto (no-fill) cell. Indexed parallel to `cellDescriptors`.
   */
  readonly cellShadings?: ReadonlyArray<string | undefined>;
  /** `sprmPIlfo` — list-format override index (>0 ⇒ the paragraph is a list item). */
  readonly listIlfo?: number;
  /** `sprmPIlvl` — list level (0–8). */
  readonly listIlvl?: number;
  /** `sprmPFWidowControl` — no first or last line left alone on a page. */
  readonly widowControl?: boolean;
  /** `sprmPFKeepFollow` — on the same page as the start of the next paragraph. */
  readonly keepNext?: boolean;
  /** `sprmPFKeep` — every line on one page. */
  readonly keepLines?: boolean;
  /** `sprmPFPageBreakBefore` — the paragraph starts a new page. */
  readonly pageBreakBefore?: boolean;
}

/** One maximal same-formatting run of text, or an embedded picture. */
export interface DocRun {
  readonly text: string;
  readonly props: DocCharProps;
  /** Set instead of `text` when the run is an embedded picture. */
  readonly picture?: DocPicture;
  /** The run is a manual page break (0x0C), alone in a run with no text. */
  readonly pageBreak?: boolean;
  /** The run is a manual column break (0x0E), alone in a run whose text is `\n`. */
  readonly columnBreak?: boolean;
}

/** One paragraph: its runs plus the decoded paragraph properties. */
export interface DocParagraph {
  readonly runs: ReadonlyArray<DocRun>;
  readonly props: DocParaProps;
  /**
   * The paragraph was ended by a cell mark (0x07) rather than a CR — i.e. it
   * closes a table cell. With `props.rowEnd` it also closes the row.
   */
  readonly cellMark?: boolean;
  /**
   * The paragraph was ended by the section mark of the (zero-based) section it
   * is the last of — the index into {@link DocContent.sections}.
   */
  readonly sectionEnd?: number;
  /** The paragraph mark's own character properties: what an empty line is set in. */
  readonly mark?: DocCharProps;
}

/** One section of the main document (§2.8.26 PlcfSed): what its SEP says. */
export interface DocSection {
  /**
   * `sprmSBkc` — where the section starts: 0 on the page in hand, 1 in the next
   * column, 2 on a new page, 3 on an even page, 4 on an odd page. Absent when
   * the SEP says nothing, which is a new page.
   */
  readonly bkc?: number;
  /** The page size in points (xaPage/yaPage), when present and sane. */
  readonly pageSize?: { readonly widthPt: number; readonly heightPt: number };
  /** `sprmSFTitlePage` — the first page shows the first-page header and footer. */
  readonly titlePage?: boolean;
  /** The margins the SEP states, in twips; each absent one is Word's default. */
  readonly margins?: DocMargins;
}

/**
 * A section's margins in twips (§2.6.4 sprmSDxaLeft … sprmSDzaGutter). A top or
 * bottom margin is stored signed, negative for one the header or footer may not
 * push; its size is the same either way.
 */
export interface DocMargins {
  readonly leftTwips?: number;
  readonly rightTwips?: number;
  readonly topTwips?: number;
  readonly bottomTwips?: number;
  readonly headerTwips?: number;
  readonly footerTwips?: number;
  readonly gutterTwips?: number;
}

/**
 * One section's header/footer stories ([MS-DOC] Plcfhdd: six per section, after
 * the six separator stories). `default` is the odd-page (primary) one. A story
 * the section leaves empty is the previous section's, the very same array; one
 * that holds no text is absent.
 */
export interface DocHeaderFooters {
  readonly defaultHeader?: ReadonlyArray<DocParagraph>;
  readonly defaultFooter?: ReadonlyArray<DocParagraph>;
  readonly firstHeader?: ReadonlyArray<DocParagraph>;
  readonly firstFooter?: ReadonlyArray<DocParagraph>;
  readonly evenHeader?: ReadonlyArray<DocParagraph>;
  readonly evenFooter?: ReadonlyArray<DocParagraph>;
}

/**
 * One list level (LVL) — the format of its number (DOC-10). `nfc` is the MSONFC
 * code (0 arabic, 1/2 upper/lower roman, 3/4 upper/lower letter, 23 bullet, 0xFF
 * none); `xst` is the number-text template as UTF-16 code units, where a code unit
 * ≤ 8 is a placeholder for the number of that (zero-based) level.
 */
export interface DocLvl {
  readonly nfc: number;
  readonly iStartAt: number;
  readonly xst: ReadonlyArray<number>;
}

/**
 * The document's list tables (DOC-10): an LFO → lsid map (a paragraph's 1-based
 * `sprmPIlfo` indexes it) and, per lsid, the list's levels.
 */
export interface DocListTables {
  readonly lfoLsids: ReadonlyArray<number>;
  readonly lists: ReadonlyMap<
    number,
    { readonly simple: boolean; readonly levels: ReadonlyArray<DocLvl> }
  >;
}

/** Everything extracted from a `.doc`'s `WordDocument` stream. */
export interface DocContent {
  readonly paragraphs: ReadonlyArray<DocParagraph>;
  /** Each section's header/footer stories, in section order. */
  readonly headerFooters?: ReadonlyArray<DocHeaderFooters>;
  /** `Dop.fFacingPages` — even pages show the even header and footer. */
  readonly facingPages?: boolean;
  /** The list tables resolving a paragraph's (ilfo, ilvl) to its number format. */
  readonly listTables?: DocListTables;
  /**
   * Every section of the main document, in order, when the PlcfSed is usable. A
   * section with no page size takes the reader's default (US Letter).
   */
  readonly sections?: ReadonlyArray<DocSection>;
  /** The file is encrypted/obfuscated — text cannot be read without the key. */
  readonly encrypted: boolean;
}

interface Piece {
  readonly cpStart: number;
  readonly cpEnd: number;
  readonly fc: number; // byte offset of the piece's first character
  readonly compressed: boolean;
}

interface BteRun<TProps> {
  readonly fcStart: number;
  readonly fcEnd: number;
  readonly props: TProps;
  /** The grpprl `props` were read from, read again over a style's. */
  readonly grpprl?: Uint8Array;
  /** A PAPX's paragraph style. */
  readonly istd?: number;
}

/**
 * The style sheet (§2.9.271 STSH), resolved along each style's bases. A
 * paragraph style gives its paragraphs paragraph properties and their runs
 * character properties; a character style lays its own over a run's.
 */
interface DocStyles {
  readonly paragraph: (istd: number) => {
    readonly para: DocParaProps;
    readonly chars: DocCharProps;
  };
  readonly character: (istd: number, base: DocCharProps) => DocCharProps;
}

const EMPTY: DocContent = { paragraphs: [], encrypted: false };

// The main document's sections (§2.8.26 PlcfSed): each one's SEP (§2.6.4), and
// where each but the last closes — the CP just past its last character, which is
// its section mark, mapped to the section's index. The mark is the same
// character as a manual page break (0x0C) and is told from one only by standing
// there. It ends the section's last paragraph the way a CR would; whether the
// next section starts a page is that section's own bkc.
//
// Each SED's fcSepx points at a Sepx in the WordDocument stream, whose grpprl
// carries sprmSBkc, sprmSXaPage/sprmSYaPage (page width/height in twips) and the
// margins.
// Word stores already-swapped dimensions for a landscape section (xaPage >
// yaPage), so they map straight to the page size — no orientation fix-up. A
// SED with no Sepx (fcSepx 0xFFFFFFFF) is a section with every default.
function readSections(
  wd: Uint8Array,
  table: Uint8Array,
): { sections: Array<DocSection>; ends: Map<number, number> } {
  const sections: Array<DocSection> = [];
  const ends = new Map<number, number>();
  const fc = u32(wd, OFF_FCPLCFSED);
  const lcb = u32(wd, OFF_LCBPLCFSED);
  // PlcfSed = (n+1) CPs (4 bytes each) then n SED structures (12 bytes); n ≥ 1.
  if (lcb < 4 + 16 || fc + lcb > table.length) return { sections, ends };
  const plc = table.subarray(fc, fc + lcb);
  const n = Math.floor((lcb - 4) / 16);
  for (let i = 0; i < n; i++) {
    // aCP[n] is no section's end of text: it may lie past the main document.
    if (i + 1 < n) ends.set(u32(plc, (i + 1) * 4), i);
    const fcSepx = u32(plc, (n + 1) * 4 + i * 12 + 2); // SED.fcSepx (offset 2 in the SED)
    if (fcSepx === 0xffffffff || fcSepx + 2 > wd.length) {
      sections.push({});
      continue;
    }
    const cb = u16(wd, fcSepx);
    const grpprl = wd.subarray(fcSepx + 2, Math.min(wd.length, fcSepx + 2 + cb));
    let bkc: number | undefined;
    let titlePage: boolean | undefined;
    let xa = 0;
    let ya = 0;
    const margins: { -readonly [K in keyof DocMargins]: number } = {};
    // A margin past the largest page is no margin: it is left to the default.
    const margin = (key: keyof DocMargins, twips: number): void => {
      if (twips <= MAX_PAGE_TWIPS) margins[key] = twips;
    };
    for (const { sprm, op } of sprms(grpprl)) {
      if (sprm === SPRM_S_BKC) bkc = grpprl[op];
      else if (sprm === SPRM_S_F_TITLE_PAGE) titlePage = grpprl[op] !== 0;
      else if (sprm === SPRM_S_XA_PAGE) xa = u16(grpprl, op);
      else if (sprm === SPRM_S_YA_PAGE) ya = u16(grpprl, op);
      else if (sprm === SPRM_S_DXA_LEFT) margin('leftTwips', u16(grpprl, op));
      else if (sprm === SPRM_S_DXA_RIGHT) margin('rightTwips', u16(grpprl, op));
      else if (sprm === SPRM_S_DYA_TOP) margin('topTwips', Math.abs(i16(grpprl, op)));
      else if (sprm === SPRM_S_DYA_BOTTOM) margin('bottomTwips', Math.abs(i16(grpprl, op)));
      else if (sprm === SPRM_S_DYA_HDR_TOP) margin('headerTwips', u16(grpprl, op));
      else if (sprm === SPRM_S_DYA_HDR_BOTTOM) margin('footerTwips', u16(grpprl, op));
      else if (sprm === SPRM_S_DZA_GUTTER) margin('gutterTwips', u16(grpprl, op));
    }
    const sane = xa >= 1 && ya >= 1 && xa <= MAX_PAGE_TWIPS && ya <= MAX_PAGE_TWIPS;
    sections.push({
      ...(bkc !== undefined ? { bkc } : {}),
      ...(sane ? { pageSize: { widthPt: xa / 20, heightPt: ya / 20 } } : {}),
      ...(titlePage ? { titlePage } : {}),
      ...(Object.keys(margins).length > 0 ? { margins } : {}),
    });
  }
  return { sections, ends };
}

/**
 * Extract a `.doc`'s `WordDocument` stream into its main-document paragraphs plus
 * the section's header/footer stories, list tables and sections.
 * Joins the piece table (CLX → CPs/FCs), CHPX run formatting and PAPX paragraph
 * formatting, splitting at the CR / cell marks. Every offset is bounds-checked, so
 * a structurally odd file yields no text rather than throwing.
 *
 * @param bytes The `.doc` (OLE2/CFB) bytes.
 * @returns The extracted {@link DocContent}; `encrypted` is set (with no
 *   paragraphs) for an encrypted file, and an empty result is returned when the
 *   stream is missing, not a Word file, or has no piece table (Word 6/95 is out of
 *   scope).
 */
export function extractDocContent(bytes: Uint8Array): DocContent {
  const cfb = openCfb(bytes);
  const wd = cfb.readStream('WordDocument');
  if (!wd || wd.length < OFF_LCBCLX + 4) return EMPTY;
  if (u16(wd, 0) !== WIDENT) return EMPTY;

  const flags = u16(wd, OFF_FLAGS1);
  if ((flags & FLAG_ENCRYPTED) !== 0) return { paragraphs: [], encrypted: true };

  const primary = (flags & FLAG_WHICH_TBL) !== 0 ? '1Table' : '0Table';
  const table =
    cfb.readStream(primary) ?? cfb.readStream(primary === '1Table' ? '0Table' : '1Table');
  if (!table) return EMPTY;

  const fcClx = u32(wd, OFF_FCCLX);
  const lcbClx = u32(wd, OFF_LCBCLX);
  if (lcbClx === 0 || fcClx + lcbClx > table.length) return EMPTY;

  const plcPcd = findPlcPcd(table.subarray(fcClx, fcClx + lcbClx));
  if (!plcPcd) return EMPTY;

  const pieces = parsePieces(plcPcd);
  const chpx = parsePlcBte(wd, table, OFF_FCPLCFBTECHPX, OFF_LCBPLCFBTECHPX, parseChpxFkp);
  const papx = parsePlcBte(wd, table, OFF_FCPLCFBTEPAPX, OFF_LCBPLCFBTEPAPX, parsePapxFkp);
  const ccpText = u32(wd, OFF_CCPTEXT);
  const data = cfb.readStream('Data'); // embedded picture store (PICF + image)
  const { sections, ends } = readSections(wd, table);
  const styles = parseStyles(wd, table);
  const main = buildParagraphs(wd, pieces, chpx, papx, 0, ccpText, data, ends, styles);
  const headerFooters = extractHeaderFooters(
    wd,
    table,
    pieces,
    chpx,
    papx,
    data,
    ccpText,
    Math.max(1, sections.length),
    styles,
  );
  const listTables = parseListTables(wd, table);
  // Dop.fFacingPages, the first bit of the Dop.
  const fcDop = u32(wd, OFF_FCDOP);
  const facingPages =
    u32(wd, OFF_LCBDOP) > 0 && fcDop < table.length && (table[fcDop]! & 0x01) !== 0;
  return {
    paragraphs: main,
    ...(headerFooters ? { headerFooters } : {}),
    ...(facingPages ? { facingPages } : {}),
    ...(listTables ? { listTables } : {}),
    ...(sections.length > 0 ? { sections } : {}),
    encrypted: false,
  };
}

// One style of the STSH, as far as the reader takes it: its kind, the style it
// is based on and the grpprls it lays over that one's.
interface Std {
  readonly stk: number;
  readonly base: number;
  readonly papx?: Uint8Array;
  readonly chpx?: Uint8Array;
}

// §2.9.271 STSH — the style sheet in the table stream: a u16 cbStshi and the
// Stshi (cstd, the count of styles, and cbSTDBaseInFile, the size of each
// STD's Stdf), then one LPStd per istd: a u16 cbStd and the STD, none when 0.
// An STD is its Stdf (stk and istdBase in its second u16, cupx in its third),
// its name (an Xstz: u16 cch, the characters, a u16 terminator) and cupx
// LPUpx: a u16 cbUpx, the UPX, a pad byte when cbUpx is odd. A paragraph
// style's UPXs are its PAPX (the istd, then a grpprl) and its CHPX; a character
// style's, its CHPX.
//
// Without it every paragraph read as if styled by nothing: the heading styles'
// size and weight, and the Normal style's size, never reached a run, and a
// style's spacing, indents and keeps never reached a paragraph.
function parseStyles(wd: Uint8Array, table: Uint8Array): DocStyles | undefined {
  const fc = u32(wd, OFF_FCSTSHF);
  const lcb = u32(wd, OFF_LCBSTSHF);
  if (lcb < 6 || fc + lcb > table.length) return undefined;
  const stsh = table.subarray(fc, fc + lcb);
  const cbStshi = u16(stsh, 0);
  if (cbStshi < 4 || 2 + cbStshi > stsh.length) return undefined;
  const cstd = u16(stsh, 2);
  const cbStdBase = u16(stsh, 4);
  const stds: Array<Std | undefined> = [];
  let p = 2 + cbStshi;
  for (let istd = 0; istd < cstd && p + 2 <= stsh.length; istd++) {
    const cbStd = u16(stsh, p);
    p += 2;
    stds.push(
      cbStd > 0 && p + cbStd <= stsh.length
        ? parseStd(stsh.subarray(p, p + cbStd), cbStdBase)
        : undefined,
    );
    p += cbStd;
  }

  const none = { para: {}, chars: {} };
  const memo = new Map<number, { readonly para: DocParaProps; readonly chars: DocCharProps }>();
  const paragraph = (istd: number, depth = 0): { para: DocParaProps; chars: DocCharProps } => {
    const hit = memo.get(istd);
    if (hit) return hit;
    const std = stds[istd];
    let out: { para: DocParaProps; chars: DocCharProps };
    if (depth >= MAX_STYLE_DEPTH) out = none;
    else if (std?.stk !== STK_PARAGRAPH)
      out = istd !== 0 ? paragraph(0, depth + 1) : none; // read as Normal
    else {
      const base =
        std.base !== ISTD_NIL && std.base !== istd ? paragraph(std.base, depth + 1) : none;
      out = {
        para: std.papx ? { ...base.para, ...decodePapxGrpprl(std.papx) } : base.para,
        chars: std.chpx ? { ...base.chars, ...decodeChpxGrpprl(std.chpx, base.chars) } : base.chars,
      };
    }
    memo.set(istd, out);
    return out;
  };
  const character = (istd: number, under: DocCharProps, depth = 0): DocCharProps => {
    const std = stds[istd];
    if (std?.stk !== STK_CHARACTER || depth >= MAX_STYLE_DEPTH) return under;
    const base =
      std.base !== ISTD_NIL && std.base !== istd ? character(std.base, under, depth + 1) : under;
    return std.chpx ? { ...base, ...decodeChpxGrpprl(std.chpx, base) } : base;
  };
  return { paragraph: (istd) => paragraph(istd), character: (istd, base) => character(istd, base) };
}

function parseStd(std: Uint8Array, cbStdBase: number): Std | undefined {
  if (std.length < 6) return undefined;
  const stk = u16(std, 2) & 0x0f;
  const base = u16(std, 2) >> 4;
  const cupx = u16(std, 4) & 0x0f;
  let q = cbStdBase;
  if (q + 2 > std.length) return { stk, base };
  q += 2 + u16(std, q) * 2 + 2; // the name and its terminator
  const upx: Array<Uint8Array> = [];
  for (let k = 0; k < cupx && q + 2 <= std.length; k++) {
    const cb = u16(std, q);
    q += 2;
    if (q + cb > std.length) break;
    upx.push(std.subarray(q, q + cb));
    q += cb + (cb & 1);
  }
  if (stk === STK_PARAGRAPH) {
    const papx = upx[0] && upx[0].length >= 2 ? upx[0].subarray(2) : undefined; // past its istd
    return { stk, base, ...(papx ? { papx } : {}), ...(upx[1] ? { chpx: upx[1] } : {}) };
  }
  if (stk === STK_CHARACTER) return { stk, base, ...(upx[0] ? { chpx: upx[0] } : {}) };
  return { stk, base };
}

// The character style a run's CHPX names (sprmCIstd), if it names one.
function charStyleOf(grpprl: Uint8Array): number | undefined {
  let istd: number | undefined;
  for (const s of sprms(grpprl)) if (s.sprm === SPRM_C_ISTD) istd = u16(grpprl, s.op);
  return istd;
}

// The list tables (DOC-10): PlfLfo (LFO → lsid) and PlfLst (the LSTFs + the LVL
// array appended after them). Bounds-checked throughout — structural doubt yields
// no/fewer lists (so a list item falls back to a generic bullet), never a wrong
// number. [MS-DOC] §2.9.{129,130,131,150}; the LVL array starts at fcPlfLst +
// lcbPlfLst (lcbPlfLst covers only the LSTFs).
function parseListTables(wd: Uint8Array, table: Uint8Array): DocListTables | undefined {
  const lstFc = u32(wd, OFF_FCPLFLST);
  const lstLcb = u32(wd, OFF_LCBPLFLST);
  if (lstLcb === 0 || lstFc + 2 > table.length) return undefined;

  // PlfLfo: lfoMac (u32) then lfoMac × 16-byte LFO (lsid at offset 0).
  const lfoLsids: Array<number> = [];
  const lfoFc = u32(wd, OFF_FCPLFLFO);
  if (lfoFc + 4 <= table.length) {
    const lfoMac = u32(table, lfoFc);
    for (let i = 0; i < lfoMac && lfoFc + 4 + i * 16 + 4 <= table.length; i++) {
      lfoLsids.push(u32(table, lfoFc + 4 + i * 16) | 0);
    }
  }

  // PlfLst: cLst (i16) then cLst × 28-byte LSTF (lsid @0, flags @26 fSimpleList b0).
  const cLst = i16(table, lstFc);
  if (cLst <= 0 || cLst > 0xffff) return undefined;
  const lstfs: Array<{ lsid: number; simple: boolean }> = [];
  for (let i = 0; i < cLst && lstFc + 2 + i * 28 + 28 <= table.length; i++) {
    const base = lstFc + 2 + i * 28;
    lstfs.push({ lsid: u32(table, base) | 0, simple: (table[base + 26]! & 0x01) !== 0 });
  }

  // The LVL array follows the LSTFs, in list order; each list contributes 1 LVL
  // (fSimpleList) or 9, in level order. Each LVL is variable-length, so walk it.
  const lists = new Map<number, { simple: boolean; levels: Array<DocLvl> }>();
  let p = lstFc + lstLcb;
  for (const lstf of lstfs) {
    const levels: Array<DocLvl> = [];
    for (let l = 0; l < (lstf.simple ? 1 : 9); l++) {
      const lvl = readLvl(table, p);
      if (!lvl) {
        p = table.length;
        break;
      }
      levels.push(lvl.info);
      p = lvl.next;
    }
    if (!lists.has(lstf.lsid)) lists.set(lstf.lsid, { simple: lstf.simple, levels });
  }
  return lfoLsids.length > 0 || lists.size > 0 ? { lfoLsids, lists } : undefined;
}

// One LVL = LVLF (28 bytes: iStartAt i32 @0, nfc u8 @4, cbGrpprlChpx @24,
// cbGrpprlPapx @25) + grpprlPapx + grpprlChpx + the Xst number-text template
// (u16 cch then cch UTF-16LE code units). Returns the level + the next offset.
function readLvl(d: Uint8Array, off: number): { info: DocLvl; next: number } | undefined {
  if (off + 28 > d.length) return undefined;
  const iStartAt = u32(d, off) | 0;
  const nfc = d[off + 4]!;
  const cbChpx = d[off + 24]!;
  const cbPapx = d[off + 25]!;
  let p = off + 28 + cbPapx + cbChpx;
  if (p + 2 > d.length) return undefined;
  const cch = u16(d, p);
  p += 2;
  const xst: Array<number> = [];
  for (let i = 0; i < cch && p + 2 <= d.length; i++) {
    xst.push(u16(d, p));
    p += 2;
  }
  return { info: { nfc, iStartAt, xst }, next: p };
}

// The [MS-DOC] Ico colour index (§2.9.59) → "RRGGBB" (0 = auto → black for a
// border). 0x0D is dark-red #800000 (the spec HTML's duplicate of 0x0C is errata).
const ICO_RGB: ReadonlyArray<string> = [
  '000000',
  '000000',
  '0000FF',
  '00FFFF',
  '00FF00',
  'FF00FF',
  'FF0000',
  'FFFF00',
  'FFFFFF',
  '000080',
  '008080',
  '008000',
  '800080',
  '800000',
  '808000',
  '808080',
  'C0C0C0',
];

// One TC80 (§2.9.328): tcgrf (u16) + wWidth (i16) + 4 × Brc80 (top, left, bottom,
// right). The 2-bit vertMerge field (tcgrf bits 5–6): 1 = continuation, 3 = restart.
function parseTc80(d: Uint8Array, off: number): DocTc {
  const vm = (u16(d, off) >> 5) & 0x3;
  const top = parseBrc80(d, off + 4);
  const left = parseBrc80(d, off + 8);
  const bottom = parseBrc80(d, off + 12);
  const right = parseBrc80(d, off + 16);
  const borders = {
    ...(top ? { top } : {}),
    ...(left ? { left } : {}),
    ...(bottom ? { bottom } : {}),
    ...(right ? { right } : {}),
  };
  return {
    ...(Object.keys(borders).length > 0 ? { borders } : {}),
    ...(vm === 3 ? { vMerge: 'restart' as const } : {}),
    ...(vm === 1 ? { vMerge: 'continue' as const } : {}),
  };
}

// One Brc80 (§2.9.18): dptLineWidth (u8, ⅛ pt), brcType (u8), ico (u8). An "on"
// border has a real line style; nil (0x00 / 0xFF) or zero width → no border.
function parseBrc80(d: Uint8Array, off: number): DocBorder | undefined {
  if (off + 4 > d.length) return undefined;
  const width = d[off]!;
  const brcType = d[off + 1]!;
  if (brcType === 0x00 || brcType === 0xff || width === 0) return undefined;
  return {
    widthPt: Math.max(2, width) / 8,
    colorHex: ICO_RGB[d[off + 2]!] ?? '000000',
    ...(brcType === 0x03 ? { double: true } : {}),
  };
}

// An (r, g, b) byte triple → an upper-case 6-hex-digit colour (a COLORREF stores
// its channels as R, G, B in that order).
function colorHex(r: number, g: number, b: number): string {
  return [r, g, b]
    .map((c) => c.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

// [MS-DOC] Plcfhdd — the header/footer stories live in the same piece-table
// stream, after the main text and footnotes; Plcfhdd is a plex of CPs (relative
// to that start) dividing it into stories. The first 6 are doc-wide separators;
// then 6 per section: even/odd page header, even/odd page footer, first-page
// header, first-page footer. A story with no characters at all is the previous
// section's, as Word's "same as previous" writes it (53379.doc leaves two of its
// eight sections so); one whose characters hold no text has no header or footer.
// Each story ends with one paragraph mark more than its paragraphs, which no
// page shows: LibreOffice writes a one-line header as "text", CR, CR.
function extractHeaderFooters(
  wd: Uint8Array,
  table: Uint8Array,
  pieces: ReadonlyArray<Piece>,
  chpx: ReadonlyArray<BteRun<DocCharProps>>,
  papx: ReadonlyArray<BteRun<DocParaProps>>,
  data: Uint8Array | undefined,
  ccpText: number,
  sectionCount: number,
  styles: DocStyles | undefined,
): ReadonlyArray<DocHeaderFooters> | undefined {
  if (u32(wd, OFF_CCPHDD) === 0) return undefined;
  const fc = u32(wd, OFF_FCPLCFHDD);
  const lcb = u32(wd, OFF_LCBPLCFHDD);
  if (lcb < 8 || fc + lcb > table.length) return undefined;
  const plc = table.subarray(fc, fc + lcb);
  const n = Math.floor(lcb / 4); // CP entries
  const hddStart = ccpText + u32(wd, OFF_CCPFTN); // the header story's first CP

  // Story i spans CP[i]..CP[i+1] (relative to hddStart): `inherit` when it has
  // no characters (or the plex ends before it), else its paragraphs, or
  // undefined when they hold no text.
  const story = (i: number): ReadonlyArray<DocParagraph> | undefined | 'inherit' => {
    if (i + 1 >= n) return 'inherit';
    const a = hddStart + u32(plc, i * 4);
    const b = hddStart + u32(plc, (i + 1) * 4);
    if (b <= a) return 'inherit';
    const paras = buildParagraphs(wd, pieces, chpx, papx, a, b, data, undefined, styles);
    const last = paras[paras.length - 1];
    if (last && last.runs.length === 0) paras.pop(); // the story's closing mark
    const hasContent = paras.some((p) => p.runs.some((r) => r.text.length > 0 || r.picture));
    return hasContent ? paras : undefined;
  };

  // Each section's six stories begin at index 6 + 6 × section, in the order of
  // SLOTS; an inherited one is the previous section's very array.
  const SLOTS = [
    'evenHeader',
    'defaultHeader', // odd-page = the primary header
    'evenFooter',
    'defaultFooter', // odd-page = the primary footer
    'firstHeader',
    'firstFooter',
  ] as const;
  const out: Array<DocHeaderFooters> = [];
  let prev: DocHeaderFooters = {};
  for (let sec = 0; sec < sectionCount; sec++) {
    const hf: { -readonly [K in keyof DocHeaderFooters]: DocHeaderFooters[K] } = {};
    SLOTS.forEach((slot, k) => {
      const got = story(6 + sec * 6 + k);
      const paras = got === 'inherit' ? prev[slot] : got;
      if (paras) hf[slot] = paras;
    });
    out.push(hf);
    prev = hf;
  }
  return out.some((hf) => Object.keys(hf).length > 0) ? out : undefined;
}

// §2.9.38 Clx — zero or more Prc (each a 0x01 byte + i16 length + grpprl) then
// the Pcdt: a 0x02 byte, a u32 length, and the PlcPcd of that length.
function findPlcPcd(clx: Uint8Array): Uint8Array | undefined {
  let p = 0;
  while (p < clx.length) {
    const clxt = clx[p];
    if (clxt === 0x01) {
      if (p + 3 > clx.length) return undefined;
      p += 3 + u16(clx, p + 1);
    } else if (clxt === 0x02) {
      if (p + 5 > clx.length) return undefined;
      const lcb = u32(clx, p + 1);
      const start = p + 5;
      if (start + lcb > clx.length) return undefined;
      return clx.subarray(start, start + lcb);
    } else {
      return undefined;
    }
  }
  return undefined;
}

// §2.8.35 PlcPcd — (n+1) CPs then n Pcds (8 bytes). Piece i spans CP[i]..CP[i+1];
// its Pcd.fc gives the byte offset and 8-vs-16-bit encoding.
function parsePieces(plc: Uint8Array): Array<Piece> {
  if (plc.length < 4 + 8) return [];
  const n = Math.floor((plc.length - 4) / 12);
  if (n <= 0) return [];
  const pcdBase = (n + 1) * 4;
  const pieces: Array<Piece> = [];
  for (let i = 0; i < n; i++) {
    const fcRaw = u32(plc, pcdBase + i * 8 + 2);
    const compressed = (fcRaw & FC_COMPRESSED) !== 0;
    pieces.push({
      cpStart: u32(plc, i * 4),
      cpEnd: u32(plc, (i + 1) * 4),
      fc: compressed ? (fcRaw & FC_MASK) >>> 1 : fcRaw & FC_MASK,
      compressed,
    });
  }
  return pieces;
}

// §2.8.5/§2.8.6 PlcBte{Papx,Chpx} — a PLC of (n+1) FCs then n PnFkp (u32, page in
// the low 22 bits). Each names a 512-byte FKP page; `parsePage` reads its runs.
function parsePlcBte<TProps>(
  wd: Uint8Array,
  table: Uint8Array,
  fcOff: number,
  lcbOff: number,
  parsePage: (page: Uint8Array, out: Array<BteRun<TProps>>) => void,
): Array<BteRun<TProps>> {
  const fcPlc = u32(wd, fcOff);
  const lcbPlc = u32(wd, lcbOff);
  if (lcbPlc < 4 + 4 || fcPlc + lcbPlc > table.length) return [];
  const plc = table.subarray(fcPlc, fcPlc + lcbPlc);
  const n = Math.floor((plc.length - 4) / 8);
  if (n <= 0) return [];
  const pnBase = (n + 1) * 4;
  const out: Array<BteRun<TProps>> = [];
  for (let i = 0; i < n; i++) {
    const pn = u32(plc, pnBase + i * 4) & PN_MASK;
    const pageOff = pn * FKP_SIZE;
    if (pageOff + FKP_SIZE > wd.length) continue;
    parsePage(wd.subarray(pageOff, pageOff + FKP_SIZE), out);
  }
  out.sort((a, b) => a.fcStart - b.fcStart);
  return out;
}

// §2.9.55 ChpxFkp — (crun+1) FCs, crun 1-byte word-offsets to each CHPX (0 = no
// formatting), and crun in byte 511. A CHPX is a u8 length then grpprl bytes.
function parseChpxFkp(page: Uint8Array, out: Array<BteRun<DocCharProps>>): void {
  const crun = page[FKP_SIZE - 1]!;
  if (crun === 0 || (crun + 1) * 4 + crun > FKP_SIZE - 1) return;
  for (let i = 0; i < crun; i++) {
    const fcStart = u32(page, i * 4);
    const fcEnd = u32(page, (i + 1) * 4);
    if (fcEnd <= fcStart) continue;
    const wordOff = page[(crun + 1) * 4 + i]!;
    if (wordOff !== 0) {
      const chpxOff = wordOff * 2;
      const cb = chpxOff < FKP_SIZE ? page[chpxOff]! : 0;
      if (cb > 0 && chpxOff + 1 + cb <= FKP_SIZE) {
        const grpprl = page.subarray(chpxOff + 1, chpxOff + 1 + cb);
        out.push({ fcStart, fcEnd, props: decodeChpxGrpprl(grpprl), grpprl });
        continue;
      }
    }
    out.push({ fcStart, fcEnd, props: {} });
  }
}

// §2.9.23 PapxFkp — (cpara+1) FCs, then cpara 13-byte BX (first byte = word-offset
// of the PapxInFkp, the rest is the ignored PHE), and cpara in byte 511. A
// PapxInFkp is a u8 cb (cb≠0 ⇒ GrpPrlAndIstd is 2·cb−1 bytes; cb=0 ⇒ a second u8
// gives 2·cb2 bytes), whose first 2 bytes are the istd and the rest the grpprl.
function parsePapxFkp(page: Uint8Array, out: Array<BteRun<DocParaProps>>): void {
  const cpara = page[FKP_SIZE - 1]!;
  if (cpara === 0 || (cpara + 1) * 4 + cpara * 13 > FKP_SIZE - 1) return;
  const bxBase = (cpara + 1) * 4;
  for (let i = 0; i < cpara; i++) {
    const fcStart = u32(page, i * 4);
    const fcEnd = u32(page, (i + 1) * 4);
    if (fcEnd <= fcStart) continue;
    const papxOff = page[bxBase + i * 13]! * 2;
    if (papxOff !== 0 && papxOff < FKP_SIZE) {
      const cb = page[papxOff]!;
      const size = cb !== 0 ? 2 * cb - 1 : 2 * (page[papxOff + 1] ?? 0);
      const start = cb !== 0 ? papxOff + 1 : papxOff + 2;
      if (size >= 2 && start + size <= FKP_SIZE) {
        const grpprl = page.subarray(start + 2, start + size);
        const istd = u16(page, start);
        out.push({ fcStart, fcEnd, props: decodePapxGrpprl(grpprl), grpprl, istd });
        continue;
      }
    }
    out.push({ fcStart, fcEnd, props: {} });
  }
}

// §2.6.1 — pick the character sprms out of a grpprl, skipping the rest by their
// spra-derived operand size so the walk stays aligned.
// `style` is what the run's style gives it: a toggle (0x80, 0x81) is read against it.
function decodeChpxGrpprl(d: Uint8Array, style: DocCharProps = {}): DocCharProps {
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let underlineKul: number | undefined;
  let sizeHalfPts: number | undefined;
  let fSpec: boolean | undefined;
  let picOffset: number | undefined;
  for (const s of sprms(d)) {
    switch (s.sprm) {
      case SPRM_C_FBOLD:
        bold = toggleBool(d[s.op]!, style.bold);
        break;
      case SPRM_C_FITALIC:
        italic = toggleBool(d[s.op]!, style.italic);
        break;
      case SPRM_C_KUL:
        underlineKul = d[s.op]!;
        break;
      case SPRM_C_HPS:
        sizeHalfPts = u16(d, s.op);
        break;
      case SPRM_C_FSPEC:
        fSpec = d[s.op] !== 0;
        break;
      case SPRM_C_PIC_LOCATION:
        picOffset = u32(d, s.op);
        break;
    }
  }
  return {
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underlineKul !== undefined ? { underlineKul } : {}),
    ...(sizeHalfPts !== undefined ? { sizeHalfPts } : {}),
    ...(fSpec ? { fSpec } : {}),
    ...(picOffset !== undefined ? { picOffset } : {}),
  };
}

// The paragraph sprms — justification, indents (signed twips) and the before/
// after spacing (twips).
function decodePapxGrpprl(d: Uint8Array): DocParaProps {
  let jc: number | undefined;
  let indentLeftTwips: number | undefined;
  let indentRightTwips: number | undefined;
  let indentFirstTwips: number | undefined;
  let spaceBeforeTwips: number | undefined;
  let spaceAfterTwips: number | undefined;
  let inTable: boolean | undefined;
  let rowEnd: boolean | undefined;
  let cellEdgesTwips: Array<number> | undefined;
  let cellDescriptors: Array<DocTc> | undefined;
  let cellShadings: Array<string | undefined> | undefined;
  let listIlfo: number | undefined;
  let listIlvl: number | undefined;
  let widowControl: boolean | undefined;
  let keepNext: boolean | undefined;
  let keepLines: boolean | undefined;
  let pageBreakBefore: boolean | undefined;
  for (const s of sprms(d)) {
    switch (s.sprm) {
      case SPRM_P_JC:
        jc = d[s.op]!;
        break;
      case SPRM_P_DXA_LEFT:
        indentLeftTwips = i16(d, s.op);
        break;
      case SPRM_P_DXA_RIGHT:
        indentRightTwips = i16(d, s.op);
        break;
      case SPRM_P_DXA_LEFT1:
        indentFirstTwips = i16(d, s.op);
        break;
      case SPRM_P_DYA_BEFORE:
        spaceBeforeTwips = u16(d, s.op);
        break;
      case SPRM_P_DYA_AFTER:
        spaceAfterTwips = u16(d, s.op);
        break;
      case SPRM_P_F_IN_TABLE:
        inTable = d[s.op] !== 0;
        break;
      case SPRM_P_F_TTP:
        rowEnd = d[s.op] !== 0;
        break;
      case SPRM_P_F_WIDOW_CONTROL:
        widowControl = d[s.op] !== 0;
        break;
      case SPRM_P_F_KEEP_FOLLOW:
        keepNext = d[s.op] !== 0;
        break;
      case SPRM_P_F_KEEP:
        keepLines = d[s.op] !== 0;
        break;
      case SPRM_P_F_PAGE_BREAK_BEFORE:
        pageBreakBefore = d[s.op] !== 0;
        break;
      case SPRM_T_DEF_TABLE: {
        // The operand is itcMac (u8), (itcMac+1) cell-boundary positions, then the
        // itcMac × TC80 cell descriptors (borders + vertical merge) — DOC-11.
        const cols = d[s.op] ?? 0;
        const edges: Array<number> = [];
        for (let i = 0; i <= cols && s.op + 1 + i * 2 + 2 <= d.length; i++) {
          edges.push(i16(d, s.op + 1 + i * 2));
        }
        if (edges.length >= 2) cellEdgesTwips = edges;
        const tcBase = s.op + 1 + (cols + 1) * 2;
        const descriptors: Array<DocTc> = [];
        for (let i = 0; i < cols && tcBase + i * 20 + 20 <= d.length; i++) {
          descriptors.push(parseTc80(d, tcBase + i * 20));
        }
        if (descriptors.length > 0) cellDescriptors = descriptors;
        break;
      }
      case SPRM_P_ILVL:
        // A level past the ninth is none a list has: Word 97's "List Bullet"
        // says 10, over a list of the old kind (ilfo 2047).
        listIlvl = d[s.op]! <= 8 ? d[s.op]! : 0;
        break;
      case SPRM_T_DEF_TABLE_SHD: {
        // §2.6.3 sprmTDefTableShd — an array of Shd (cvFore[4] cvBack[4] ipat[2]),
        // one per cell. cvBack is the cell's background fill (= the OOXML w:shd
        // @fill); a COLORREF whose 4th byte is 0x00 is a literal sRGB, 0xFF is auto.
        const shd: Array<string | undefined> = [];
        for (let i = 0; s.op + i * 10 + 10 <= d.length && i < MAX_ROW_CELLS; i++) {
          const o = s.op + i * 10 + 4; // cvBack
          shd.push(d[o + 3] === 0x00 ? colorHex(d[o]!, d[o + 1]!, d[o + 2]!) : undefined);
        }
        if (shd.length > 0) cellShadings = shd;
        break;
      }
      case SPRM_P_ILFO:
        listIlfo = u16(d, s.op);
        break;
    }
  }
  return {
    ...(jc !== undefined ? { jc } : {}),
    ...(indentLeftTwips !== undefined ? { indentLeftTwips } : {}),
    ...(indentRightTwips !== undefined ? { indentRightTwips } : {}),
    ...(indentFirstTwips !== undefined ? { indentFirstTwips } : {}),
    ...(spaceBeforeTwips !== undefined ? { spaceBeforeTwips } : {}),
    ...(spaceAfterTwips !== undefined ? { spaceAfterTwips } : {}),
    ...(inTable ? { inTable } : {}),
    ...(rowEnd ? { rowEnd } : {}),
    ...(cellEdgesTwips ? { cellEdgesTwips } : {}),
    ...(cellDescriptors ? { cellDescriptors } : {}),
    ...(cellShadings ? { cellShadings } : {}),
    ...(listIlfo !== undefined ? { listIlfo } : {}),
    ...(listIlvl !== undefined ? { listIlvl } : {}),
    ...(widowControl !== undefined ? { widowControl } : {}),
    ...(keepNext !== undefined ? { keepNext } : {}),
    ...(keepLines !== undefined ? { keepLines } : {}),
    ...(pageBreakBefore !== undefined ? { pageBreakBefore } : {}),
  };
}

// Iterate a grpprl, yielding each sprm opcode and the offset of its operand.
function* sprms(d: Uint8Array): Generator<{ sprm: number; op: number }> {
  let p = 0;
  while (p + 2 <= d.length) {
    const sprm = u16(d, p);
    p += 2;
    const spra = sprm >> 13;
    let op = p;
    let len = SPRA_LEN[spra]!;
    if (spra === 6) {
      if (SPRM_LONG_OPERAND.has(sprm)) {
        if (p + 2 > d.length) break;
        const cb = u16(d, p); // a 2-byte count, including the count field itself
        if (cb < 1) break;
        len = cb - 1;
        op = p + 2;
      } else {
        if (p >= d.length) break;
        len = d[p]!;
        op = p + 1;
      }
    }
    if (op + len > d.length) break;
    yield { sprm, op };
    p = op + len;
  }
}

// A toggle sprm operand: 0 off, 1 on, 0x80 what the style says, 0x81 the
// opposite of what it says (a style that says nothing is off).
function toggleBool(v: number, style: boolean | undefined): boolean | undefined {
  if (v === 0) return false;
  if (v === 1) return true;
  if (v === 0x80) return style;
  if (v === 0x81) return style !== true;
  return undefined;
}

// §2.9.158 PICF — the picture descriptor at a Data-stream offset: a header
// (cbHeader bytes) then the image. dxaGoal/dyaGoal are the natural size in twips,
// scaled by mx/my (per-mille). We scan the image region for a raster magic (an
// OfficeArt blip header precedes the PNG/JPEG); a metafile (WMF/EMF) has no such
// magic and is skipped.
function extractPicture(data: Uint8Array, fcPic: number): DocPicture | undefined {
  if (fcPic < 0 || fcPic + 0x24 > data.length) return undefined;
  const lcb = u32(data, fcPic);
  const cbHeader = u16(data, fcPic + 4);
  if (cbHeader < 0x2e || lcb <= cbHeader || fcPic + lcb > data.length) return undefined;
  const dxaGoal = u16(data, fcPic + 0x1c);
  const dyaGoal = u16(data, fcPic + 0x1e);
  const mx = u16(data, fcPic + 0x20) || 1000;
  const my = u16(data, fcPic + 0x22) || 1000;
  const bytes = scanImage(data.subarray(fcPic + cbHeader, fcPic + lcb));
  if (!bytes) return undefined;
  return {
    bytes,
    widthTwips: Math.round((dxaGoal * mx) / 1000),
    heightTwips: Math.round((dyaGoal * my) / 1000),
  };
}

const IMAGE_MAGICS: ReadonlyArray<ReadonlyArray<number>> = [
  [0x89, 0x50, 0x4e, 0x47], // PNG
  [0xff, 0xd8, 0xff], // JPEG
  [0x47, 0x49, 0x46, 0x38], // GIF
  [0x42, 0x4d], // BMP
  [0x49, 0x49, 0x2a, 0x00], // TIFF (LE)
  [0x4d, 0x4d, 0x00, 0x2a], // TIFF (BE)
];

// The first known raster magic in the region → the bytes from there to the end
// (the OfficeArt blip header precedes the image payload).
function scanImage(region: Uint8Array): Uint8Array | undefined {
  const limit = Math.min(region.length, 512);
  for (let off = 0; off < limit; off++) {
    for (const magic of IMAGE_MAGICS) {
      if (off + magic.length <= region.length && magicAt(region, off, magic)) {
        return region.subarray(off);
      }
    }
  }
  return undefined;
}

function magicAt(d: Uint8Array, off: number, magic: ReadonlyArray<number>): boolean {
  for (let i = 0; i < magic.length; i++) if (d[off + i] !== magic[i]) return false;
  return true;
}

// Walk the pieces character by character: attach each character's CHPX (by FC),
// split paragraphs at the CR mark (attaching each paragraph's PAPX, also by FC),
// clean the in-paragraph control characters, and coalesce same-formatting runs.
// `sectionEnds` maps the CP past each section mark of the main document to the
// index of the section the mark closes. With `styles`, a paragraph's properties
// are its style's with its PAPX over them, and a run's are its paragraph style's
// with its character style's and its CHPX over them.
function buildParagraphs(
  wd: Uint8Array,
  pieces: ReadonlyArray<Piece>,
  chpx: ReadonlyArray<BteRun<DocCharProps>>,
  papx: ReadonlyArray<BteRun<DocParaProps>>,
  cpStart: number,
  cpEnd: number,
  data: Uint8Array | undefined,
  sectionEnds: ReadonlyMap<number, number> = new Map(),
  styles?: DocStyles,
): Array<DocParagraph> {
  const end = Math.min(cpEnd, cpStart + MAX_TEXT);
  const paragraphs: Array<DocParagraph> = [];
  // The paragraph's runs so far, each beside the CHPX it was read from: what
  // they look like depends on the paragraph's style, known only at its end.
  let runs: Array<DocRun> = [];
  let chpxOf: Array<BteRun<DocCharProps> | undefined> = [];
  let cur = '';
  let curChpx: BteRun<DocCharProps> | undefined;
  let lastFc = 0;
  // Field nesting: a field is `0x13 code 0x14 result 0x15`; the code is dropped,
  // the cached result kept. codeDepth counts fields currently inside their code
  // portion (which may nest), so any text there is suppressed.
  let fieldDepth = 0;
  let codeDepth = 0;

  const push = (run: DocRun, from: BteRun<DocCharProps> | undefined): void => {
    runs.push(run);
    chpxOf.push(from);
  };
  const flushRun = (): void => {
    if (cur.length > 0) {
      push({ text: cur, props: curChpx?.props ?? {} }, curChpx);
      cur = '';
    }
  };
  // A run's properties over what its paragraph's style gives it: its character
  // style's (sprmCIstd), then its own CHPX.
  const runProps = (from: BteRun<DocCharProps> | undefined, base: DocCharProps): DocCharProps => {
    if (!from?.grpprl) return base;
    const cistd = charStyleOf(from.grpprl);
    const under = cistd !== undefined && styles ? styles.character(cistd, base) : base;
    return { ...under, ...decodeChpxGrpprl(from.grpprl, under) };
  };
  const endParagraph = (fc: number, cellMark: boolean, sectionEnd?: number): void => {
    flushRun();
    const pap = lookupRun(papx, fc);
    const style = styles?.paragraph(pap?.istd ?? 0);
    // Each run read over the style, and runs of text that come out alike as one.
    const settled: Array<DocRun> = [];
    runs.forEach((run, i) => {
      const props = style && !run.picture ? runProps(chpxOf[i], style.chars) : run.props;
      const prev = settled[settled.length - 1];
      if (prev && plainText(prev) && plainText(run) && sameProps(prev.props, props)) {
        settled[settled.length - 1] = { text: prev.text + run.text, props: prev.props };
      } else settled.push({ ...run, props });
    });
    const markChpx = lookupRun(chpx, fc);
    const mark = style ? runProps(markChpx, style.chars) : markChpx?.props;
    paragraphs.push({
      runs: settled,
      props: style ? { ...style.para, ...(pap?.props ?? {}) } : (pap?.props ?? {}),
      ...(cellMark ? { cellMark: true } : {}),
      ...(sectionEnd !== undefined ? { sectionEnd } : {}),
      ...(mark && Object.keys(mark).length > 0 ? { mark } : {}),
    });
    runs = [];
    chpxOf = [];
  };

  for (const piece of pieces) {
    if (piece.cpStart >= end) break;
    if (piece.cpEnd <= cpStart) continue; // the piece is entirely before the window
    const step = piece.compressed ? 1 : 2;
    const kStart = Math.max(0, cpStart - piece.cpStart);
    const kEnd = Math.min(piece.cpEnd - piece.cpStart, end - piece.cpStart);
    for (let k = kStart; k < kEnd; k++) {
      const fc = piece.fc + k * step;
      lastFc = fc;
      const code = piece.compressed ? byteCode(wd, fc) : unitCode(wd, fc);
      if (code < 0) continue;
      // A CR ends a paragraph; a cell mark (0x07) ends a table cell (and, on the
      // TTP, a row); a section mark ends the section's last paragraph — all three
      // terminate the current paragraph. Word's PAPX runs end at a section mark,
      // after text or alone in a paragraph, and never at a page break.
      const sectionEnd = code === PAGE_BREAK ? sectionEnds.get(piece.cpStart + k + 1) : undefined;
      if (code === CR || code === CELL_MARK || sectionEnd !== undefined) {
        endParagraph(fc, code === CELL_MARK, sectionEnd);
        continue;
      }
      // Field delimiters: enter the code (suppress), reach the separator (the
      // result follows, kept), or end the field. The markers themselves are dropped.
      if (code === FIELD_BEGIN) {
        fieldDepth++;
        codeDepth++;
        continue;
      }
      if (code === FIELD_SEPARATOR) {
        if (codeDepth > 0) codeDepth--;
        continue;
      }
      if (code === FIELD_END) {
        if (fieldDepth > 0) fieldDepth--;
        if (codeDepth > fieldDepth) codeDepth = fieldDepth; // a field with no separator
        continue;
      }
      if (codeDepth > 0) continue; // inside a field's code — suppress everything
      // A picture occupies a special 0x01 character whose CHPX points (via
      // sprmCPicLocation) at a PICF in the Data stream.
      if (code === PIC_CHAR) {
        const props = lookup(chpx, fc) ?? {};
        const picture =
          data && props.fSpec && props.picOffset !== undefined
            ? extractPicture(data, props.picOffset)
            : undefined;
        if (picture) {
          flushRun();
          push({ text: '', props: {}, picture }, undefined);
        }
        continue; // the 0x01 placeholder itself is not text
      }
      // A manual page or column break stands where it is in the paragraph: what
      // comes before it ends its line there, what comes after starts the next
      // page or column. Dropped with the other control characters, a page break
      // printed "Latin", the break and "Second" as "LatinSecond" on one page. It
      // is a run of its own, the way the docx reader cuts a run at a `w:br`.
      if (code === PAGE_BREAK || code === COLUMN_BREAK) {
        flushRun();
        const from = lookupRun(chpx, fc);
        const props = from?.props ?? {};
        push(
          code === PAGE_BREAK
            ? { text: '', props, pageBreak: true }
            : { text: '\n', props, columnBreak: true },
          from,
        );
        continue;
      }
      // A line break ends its line and goes on in the paragraph, as `\n` does in a
      // docx run: read as a space, "Latin", a break and "Second" set one line.
      const ch =
        code === LINE_BREAK
          ? '\n'
          : code === 0x09
            ? ' '
            : code < 0x20
              ? ''
              : String.fromCharCode(code);
      if (ch.length === 0) continue; // drop the other control characters
      const from = lookupRun(chpx, fc);
      if (cur.length === 0) curChpx = from;
      else if (from !== curChpx) {
        flushRun();
        curChpx = from;
      }
      cur += ch;
    }
  }
  // Flush a final paragraph only when text didn't end on a terminator — otherwise
  // the trailing mark would add a spurious empty paragraph (and a phantom row in a
  // table that ends, as it must, on the TTP).
  if (cur.length > 0 || runs.length > 0) endParagraph(lastFc, false);
  return paragraphs;
}

// Binary search the FC-sorted runs for the one covering `fc`.
function lookup<TProps>(runs: ReadonlyArray<BteRun<TProps>>, fc: number): TProps | undefined {
  return lookupRun(runs, fc)?.props;
}

function lookupRun<TProps>(
  runs: ReadonlyArray<BteRun<TProps>>,
  fc: number,
): BteRun<TProps> | undefined {
  let lo = 0;
  let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const r = runs[mid]!;
    if (fc < r.fcStart) hi = mid - 1;
    else if (fc >= r.fcEnd) lo = mid + 1;
    else return r;
  }
  return undefined;
}

// A run of nothing but text, which may join the one beside it.
function plainText(run: DocRun): boolean {
  return !run.picture && !run.pageBreak && !run.columnBreak;
}

function sameProps(a: DocCharProps, b: DocCharProps): boolean {
  return (
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underlineKul === b.underlineKul &&
    a.sizeHalfPts === b.sizeHalfPts
  );
}

// Windows-1252 high range (0x80–0x9F); the rest of 0x00–0xFF maps to U+0000–00FF.
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152,
  0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122,
  0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

function byteCode(d: Uint8Array, fc: number): number {
  const b = d[fc];
  if (b === undefined) return -1;
  return b >= 0x80 && b <= 0x9f ? CP1252_HIGH[b - 0x80]! : b;
}

function unitCode(d: Uint8Array, fc: number): number {
  if (fc + 1 >= d.length) return -1;
  return d[fc]! | (d[fc + 1]! << 8);
}

function u16(d: Uint8Array, off: number): number {
  return (d[off] ?? 0) | ((d[off + 1] ?? 0) << 8);
}
function i16(d: Uint8Array, off: number): number {
  const v = u16(d, off);
  return v >= 0x8000 ? v - 0x10000 : v;
}
function u32(d: Uint8Array, off: number): number {
  return (
    ((d[off] ?? 0) |
      ((d[off + 1] ?? 0) << 8) |
      ((d[off + 2] ?? 0) << 16) |
      ((d[off + 3] ?? 0) << 24)) >>>
    0
  );
}
