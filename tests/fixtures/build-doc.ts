// Minimal Word 97–2003 `.doc` builder for tests — packs a `WordDocument` stream
// (the FIB fields the reader needs, the piece text and optional CHPX/PAPX FKP
// pages) and a table stream (the CLX piece table plus optional PlcBteChpx /
// PlcBtePapx bin tables) into a CFB container, so the `.doc` reader has a
// deterministic input without a checked-in binary. Mirrors build-xls.ts: it
// writes exactly the bytes doc-text.ts reads back (FibBase offsets, PlcPcd,
// FcCompressed, FKP, sprm), so the round trip validates the offset / compression
// / formatting logic rather than restating it.
//
// CHPX/PAPX formatting is synthesised against the FIRST piece's FC mapping, so
// tests that pass `formatRuns` / `paraRuns` use a single piece (enough to
// exercise the reader's FKP → sprm → FC→CP path; the reader handles multi-piece).

import { buildCfb } from './build-cfb';

export interface DocPiece {
  readonly text: string;
  // true → stored as 8-bit Windows-1252 ("compressed"); false → 16-bit UTF-16LE.
  readonly compressed: boolean;
}

// A run of character formatting over `length` characters of the document text.
export interface DocFormatRun {
  readonly length: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underlineKul?: number; // Word `kul` code (0 = none, 1 = single, …)
  readonly sizeHalfPts?: number; // font size in half-points
  readonly picOffset?: number; // a picture: sprmCFSpec + sprmCPicLocation → Data offset
}

// A paragraph's formatting, over its `length` characters (including the CR / cell
// mark that terminates it).
export interface DocParaFormat {
  readonly length: number;
  readonly jc?: number; // 0 left, 1 center, 2 right, 3 both, 4 distribute
  readonly indentLeftTwips?: number;
  readonly indentRightTwips?: number;
  readonly indentFirstTwips?: number; // signed (negative = hanging)
  readonly spaceBeforeTwips?: number;
  readonly spaceAfterTwips?: number;
  readonly inTable?: boolean; // sprmPFInTable — the paragraph is a table cell
  readonly rowEnd?: boolean; // sprmPFTtp — the row's terminating paragraph
  readonly cellEdgesTwips?: ReadonlyArray<number>; // sprmTDefTable cell boundaries
  // Per-cell TC80 descriptors (DOC-11): borders (each a Brc80) + vertical merge.
  readonly cellTc?: ReadonlyArray<{
    readonly borders?: {
      readonly top?: DocBrcInput;
      readonly left?: DocBrcInput;
      readonly bottom?: DocBrcInput;
      readonly right?: DocBrcInput;
    };
    readonly vMerge?: 'restart' | 'continue';
  }>;
  readonly listIlfo?: number; // sprmPIlfo — list override (>0 ⇒ a list item)
  readonly listIlvl?: number; // sprmPIlvl — list level
}

// One Brc80 edge for the fixture: width in eighths of a point, brcType (1 single /
// 3 double), and the Ico colour index.
interface DocBrcInput {
  readonly widthEighthPt?: number;
  readonly brcType?: number;
  readonly ico?: number;
}

// Where piece text starts in the WordDocument stream — past the FIB fields the
// reader probes (fcClx/lcbClx sit at 0x1A2/0x1A6).
const TEXT_BASE = 0x200;
const FKP_SIZE = 512;

export function buildDoc(
  pieces: ReadonlyArray<DocPiece>,
  opts: {
    readonly whichTable?: '0Table' | '1Table';
    readonly encrypted?: boolean;
    readonly formatRuns?: ReadonlyArray<DocFormatRun>;
    readonly paraRuns?: ReadonlyArray<DocParaFormat>;
    readonly data?: Uint8Array; // the Data stream (embedded pictures: see buildPicf)
    readonly headerFooter?: {
      readonly defaultHeader?: string;
      readonly defaultFooter?: string;
      readonly firstHeader?: string;
      readonly firstFooter?: string;
      readonly evenHeader?: string;
      readonly evenFooter?: string;
    };
    // List tables (DOC-10): `lfos` maps each 1-based ilfo to a list's lsid; `lstfs`
    // gives each list its levels (nfc + iStartAt + the xst number-text code units).
    readonly lists?: {
      readonly lfos: ReadonlyArray<number>;
      readonly lstfs: ReadonlyArray<{
        readonly lsid: number;
        readonly simple?: boolean;
        readonly levels: ReadonlyArray<{
          readonly nfc: number;
          readonly iStartAt?: number;
          readonly xst: ReadonlyArray<number>;
        }>;
      }>;
    };
  } = {},
): Uint8Array {
  const whichTable = opts.whichTable ?? '1Table';

  // Lay each piece's bytes into the WordDocument stream from TEXT_BASE, keeping
  // every piece 2-byte aligned (UTF-16 needs it; the compressed offset is halved).
  let cursor = TEXT_BASE;
  const placed: Array<{ offset: number; bytes: Uint8Array; chars: number; compressed: boolean }> =
    [];
  for (const p of pieces) {
    if (cursor % 2 !== 0) cursor++;
    const bytes = p.compressed ? encodeCp1252(p.text) : encodeUtf16(p.text);
    placed.push({ offset: cursor, bytes, chars: p.text.length, compressed: p.compressed });
    cursor += bytes.length;
  }
  const mainChars = placed.reduce((sum, p) => sum + p.chars, 0);

  // Optional header/footer stories: appended after the main text so they sit at
  // CP ≥ ccpText, with a PlcfHdd dividing them.
  let headerCps: Array<number> | undefined;
  if (opts.headerFooter) {
    const story = buildHeaderStream(opts.headerFooter);
    headerCps = story.cps;
    if (cursor % 2 !== 0) cursor++;
    const bytes = encodeUtf16(story.text);
    placed.push({ offset: cursor, bytes, chars: story.text.length, compressed: false });
    cursor += bytes.length;
  }
  const totalChars = placed.reduce((sum, p) => sum + p.chars, 0);

  // PlcPcd: (n+1) character positions then n 8-byte Pcds.
  const n = placed.length;
  const plc = new Uint8Array((n + 1) * 4 + n * 8);
  const pv = new DataView(plc.buffer);
  let cpAcc = 0;
  pv.setUint32(0, 0, true);
  for (let i = 0; i < n; i++) {
    cpAcc += placed[i]!.chars;
    pv.setUint32((i + 1) * 4, cpAcc, true);
  }
  const pcdBase = (n + 1) * 4;
  for (let i = 0; i < n; i++) {
    const pc = placed[i]!;
    // FcCompressed: bit 30 = fCompressed; for a compressed piece the stored fc is
    // twice the byte offset (the reader halves it).
    const fc = pc.compressed ? (0x40000000 | (pc.offset * 2)) >>> 0 : pc.offset;
    pv.setUint16(pcdBase + i * 8, 0, true); // Pcd flags
    pv.setUint32(pcdBase + i * 8 + 2, fc, true); // FcCompressed
    pv.setUint16(pcdBase + i * 8 + 6, 0, true); // prm
  }

  // CLX: a lone Pcdt (clxt 0x02, u32 length, PlcPcd) — no property modifiers.
  const clx = new Uint8Array(5 + plc.length);
  clx[0] = 0x02;
  new DataView(clx.buffer).setUint32(1, plc.length, true);
  clx.set(plc, 5);

  // Append FKP pages (in the WordDocument stream, 512-aligned, after the text)
  // and bin tables (in the table stream, 4-aligned, after the CLX). Each FKP page
  // is named by a one-entry PlcBte that points at it.
  const first = placed[0];
  const step = first?.compressed ? 1 : 2;
  const fkpPages: Array<{ offset: number; page: Uint8Array }> = [];
  const binTables: Array<{ offset: number; data: Uint8Array }> = [];
  const fibFields: Array<{ fc: number; off: number; lcb: number }> = [];
  let wdEnd = cursor;
  let tableEnd = clx.length;

  const addFkp = (page: Uint8Array): number => {
    const offset = Math.ceil(wdEnd / FKP_SIZE) * FKP_SIZE;
    fkpPages.push({ offset, page });
    wdEnd = offset + FKP_SIZE;
    return offset / FKP_SIZE;
  };
  const addBinTable = (data: Uint8Array): number => {
    const offset = Math.ceil(tableEnd / 4) * 4;
    binTables.push({ offset, data });
    tableEnd = offset + data.length;
    return offset;
  };
  const addBte = (page: Uint8Array, fcFieldOff: number): void => {
    const pageNumber = addFkp(page);
    const bte = buildPlcBte(first!.offset, first!.offset + totalChars * step, pageNumber);
    const tableOff = addBinTable(bte);
    fibFields.push({ fc: fcFieldOff, off: tableOff, lcb: bte.length });
  };

  if (first && opts.formatRuns && opts.formatRuns.length > 0) {
    addBte(buildChpxFkp(first.offset, step, opts.formatRuns), 0xfa);
  }
  if (first && opts.paraRuns && opts.paraRuns.length > 0) {
    addBte(buildPapxFkp(first.offset, step, opts.paraRuns), 0x102);
  }
  if (headerCps) {
    const plcfHdd = new Uint8Array(headerCps.length * 4);
    const hv = new DataView(plcfHdd.buffer);
    headerCps.forEach((cp, i) => hv.setUint32(i * 4, cp, true));
    fibFields.push({ fc: 0xf2, off: addBinTable(plcfHdd), lcb: plcfHdd.length }); // PlcfHdd
  }
  if (opts.lists) {
    const { plfLst, lcbLstfs, plfLfo } = buildListTables(opts.lists);
    fibFields.push({ fc: 0x2e2, off: addBinTable(plfLst), lcb: lcbLstfs }); // fcPlfLst/lcbPlfLst
    fibFields.push({ fc: 0x2ea, off: addBinTable(plfLfo), lcb: plfLfo.length }); // fcPlfLfo/lcb
  }

  // Assemble the table stream: CLX then the bin tables.
  const table = new Uint8Array(tableEnd);
  table.set(clx, 0);
  for (const b of binTables) table.set(b.data, b.offset);

  // Assemble the WordDocument stream: FIB fields, piece bytes, FKP pages.
  const wd = new Uint8Array(wdEnd);
  const wv = new DataView(wd.buffer);
  wv.setUint16(0x00, 0xa5ec, true); // wIdent
  wv.setUint16(0x02, 0x00c1, true); // nFib (Word 97)
  wv.setUint16(
    0x0a,
    (whichTable === '1Table' ? 0x0200 : 0x0000) | (opts.encrypted ? 0x0100 : 0x0000),
    true,
  ); // fWhichTblStm | fEncrypted
  wv.setUint32(0x4c, mainChars, true); // ccpText (the header story sits past it)
  if (headerCps) {
    wv.setUint32(0x50, 0, true); // ccpFtn (no footnotes)
    wv.setUint32(0x54, totalChars - mainChars, true); // ccpHdd
  }
  for (const f of fibFields) {
    wv.setUint32(f.fc, f.off, true); // fcPlcfBte{Chpx,Papx}
    wv.setUint32(f.fc + 4, f.lcb, true); // lcbPlcfBte{Chpx,Papx}
  }
  wv.setUint32(0x1a2, 0, true); // fcClx — CLX at offset 0 of the table stream
  wv.setUint32(0x1a6, clx.length, true); // lcbClx
  for (const p of placed) wd.set(p.bytes, p.offset);
  for (const f of fkpPages) wd.set(f.page, f.offset);

  return buildCfb([
    { name: 'WordDocument', data: wd },
    { name: whichTable, data: table },
    ...(opts.data ? [{ name: 'Data', data: opts.data }] : []),
  ]);
}

// The list tables (DOC-10): PlfLfo (lfoMac + 16-byte LFOs, lsid @0) and PlfLst
// (cLst + 28-byte LSTFs, then the appended LVL array — each LVL = a 28-byte LVLF +
// the xst). lcbPlfLst covers only the LSTFs; the LVLs follow contiguously.
function buildListTables(lists: {
  readonly lfos: ReadonlyArray<number>;
  readonly lstfs: ReadonlyArray<{
    readonly lsid: number;
    readonly simple?: boolean;
    readonly levels: ReadonlyArray<{
      readonly nfc: number;
      readonly iStartAt?: number;
      readonly xst: ReadonlyArray<number>;
    }>;
  }>;
}): { plfLst: Uint8Array; lcbLstfs: number; plfLfo: Uint8Array } {
  const plfLfo = new Uint8Array(4 + lists.lfos.length * 16);
  const fv = new DataView(plfLfo.buffer);
  fv.setUint32(0, lists.lfos.length, true); // lfoMac
  lists.lfos.forEach((lsid, i) => fv.setInt32(4 + i * 16, lsid, true)); // LFO.lsid

  const cLst = lists.lstfs.length;
  const lcbLstfs = 2 + cLst * 28;
  const lstfPart = new Uint8Array(lcbLstfs);
  const lv = new DataView(lstfPart.buffer);
  lv.setInt16(0, cLst, true); // cLst
  lists.lstfs.forEach((lst, i) => {
    const base = 2 + i * 28;
    lv.setInt32(base, lst.lsid, true); // LSTF.lsid
    lstfPart[base + 26] = lst.simple ? 0x01 : 0x00; // flags: fSimpleList
  });

  const parts: Array<Uint8Array> = [lstfPart];
  for (const lst of lists.lstfs) {
    const count = lst.simple ? 1 : 9;
    for (let l = 0; l < count; l++) {
      const lvl = lst.levels[l] ?? lst.levels[lst.levels.length - 1] ?? { nfc: 23, xst: [0x2022] };
      const lvlf = new Uint8Array(28); // cbGrpprlChpx @24 / cbGrpprlPapx @25 = 0
      new DataView(lvlf.buffer).setInt32(0, lvl.iStartAt ?? 1, true);
      lvlf[4] = lvl.nfc;
      const xst = new Uint8Array(2 + lvl.xst.length * 2);
      const xv = new DataView(xst.buffer);
      xv.setUint16(0, lvl.xst.length, true);
      lvl.xst.forEach((cu, j) => xv.setUint16(2 + j * 2, cu, true));
      parts.push(lvlf, xst);
    }
  }

  const total = parts.reduce((n, p) => n + p.length, 0);
  const plfLst = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    plfLst.set(p, off);
    off += p.length;
  }
  return { plfLst, lcbLstfs, plfLfo };
}

// §2.8.26 PlcfHdd — lay the 12 header/footer stories (indices 0–5 are doc-wide
// separators, left empty; 6–11 are section 0's even/odd/first header and footer)
// as CR-terminated text, returning the concatenated story stream and the 13 CP
// boundaries that divide it.
function buildHeaderStream(hf: {
  readonly defaultHeader?: string;
  readonly defaultFooter?: string;
  readonly firstHeader?: string;
  readonly firstFooter?: string;
  readonly evenHeader?: string;
  readonly evenFooter?: string;
}): { text: string; cps: Array<number> } {
  const cr = String.fromCharCode(0x0d);
  const slots = [
    '',
    '',
    '',
    '',
    '',
    '', // doc-wide separators
    hf.evenHeader, // 6
    hf.defaultHeader, // 7 (odd-page = primary)
    hf.evenFooter, // 8
    hf.defaultFooter, // 9 (odd-page = primary)
    hf.firstHeader, // 10
    hf.firstFooter, // 11
  ].map((t) => (t ? t + cr : ''));
  const cps = [0];
  let text = '';
  for (const s of slots) {
    text += s;
    cps.push(text.length);
  }
  return { text, cps };
}

// §2.9.158 PICF — a picture descriptor (68-byte header) then the image bytes, as
// stored in the Data stream. The image sits right after the header so the reader's
// magic scan finds it at the region start.
export function buildPicf(image: Uint8Array, widthTwips: number, heightTwips: number): Uint8Array {
  const cbHeader = 0x44; // 68
  const out = new Uint8Array(cbHeader + image.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, out.length, true); // lcb
  dv.setUint16(4, cbHeader, true); // cbHeader
  dv.setUint16(6, 0x0064, true); // mm — not a metafile (an OfficeArt blip)
  dv.setUint16(0x1c, widthTwips, true); // dxaGoal
  dv.setUint16(0x1e, heightTwips, true); // dyaGoal
  dv.setUint16(0x20, 1000, true); // mx — 100%
  dv.setUint16(0x22, 1000, true); // my — 100%
  out.set(image, cbHeader);
  return out;
}

// §2.9.55 ChpxFkp — (crun+1) FCs, crun word-offsets to each CHPX (0 = none), the
// CHPX blobs (u8 length + grpprl) at word-aligned offsets, and crun in byte 511.
function buildChpxFkp(fcBase: number, step: number, runs: ReadonlyArray<DocFormatRun>): Uint8Array {
  const crun = runs.length;
  const page = new Uint8Array(FKP_SIZE);
  const dv = new DataView(page.buffer);
  dv.setUint32(0, fcBase, true);
  let cum = 0;
  for (let i = 0; i < crun; i++) {
    cum += runs[i]!.length;
    dv.setUint32((i + 1) * 4, fcBase + cum * step, true);
  }
  const rgbBase = (crun + 1) * 4;
  let blobOff = rgbBase + crun;
  for (let i = 0; i < crun; i++) {
    const grpprl = buildChpxGrpprl(runs[i]!);
    if (grpprl.length === 0) {
      page[rgbBase + i] = 0; // no formatting
      continue;
    }
    if (blobOff % 2 !== 0) blobOff++; // CHPX offsets are in 2-byte words
    page[blobOff] = grpprl.length; // cb
    page.set(grpprl, blobOff + 1);
    page[rgbBase + i] = blobOff / 2;
    blobOff += 1 + grpprl.length;
  }
  page[FKP_SIZE - 1] = crun;
  return page;
}

// §2.9.23 PapxFkp — (cpara+1) FCs, cpara 13-byte BX (first byte = word-offset of
// the PapxInFkp), the PapxInFkp blobs at word-aligned offsets, and cpara in byte
// 511. Every paragraph has a PAPX (unlike CHPX, where 0 means "no formatting").
function buildPapxFkp(
  fcBase: number,
  step: number,
  runs: ReadonlyArray<DocParaFormat>,
): Uint8Array {
  const cpara = runs.length;
  const page = new Uint8Array(FKP_SIZE);
  const dv = new DataView(page.buffer);
  dv.setUint32(0, fcBase, true);
  let cum = 0;
  for (let i = 0; i < cpara; i++) {
    cum += runs[i]!.length;
    dv.setUint32((i + 1) * 4, fcBase + cum * step, true);
  }
  const bxBase = (cpara + 1) * 4;
  let blobOff = bxBase + cpara * 13;
  for (let i = 0; i < cpara; i++) {
    const blob = buildPapxInFkp(buildPapxGrpprl(runs[i]!));
    if (blobOff % 2 !== 0) blobOff++; // PapxInFkp offsets are in 2-byte words
    page.set(blob, blobOff);
    page[bxBase + i * 13] = blobOff / 2; // BX bOffset; the rest of the BX is PHE (0)
    blobOff += blob.length;
  }
  page[FKP_SIZE - 1] = cpara;
  return page;
}

// §2.9.24 PapxInFkp — a u8 cb then GrpPrlAndIstd (istd + grpprl). cb≠0 ⇒ size is
// 2·cb−1 (odd); cb=0 ⇒ a second u8 gives 2·cb2 (even). istd is always 0 here.
function buildPapxInFkp(grpprl: Uint8Array): Uint8Array {
  const size = 2 + grpprl.length; // GrpPrlAndIstd: istd (2 bytes) + grpprl
  if (size % 2 === 1) {
    return Uint8Array.from([(size + 1) / 2, 0, 0, ...grpprl]); // cb, istd lo/hi, grpprl
  }
  return Uint8Array.from([0, size / 2, 0, 0, ...grpprl]); // cb=0, cb2, istd lo/hi, grpprl
}

// §2.6.1 — the character sprm list for a run.
function buildChpxGrpprl(run: DocFormatRun): Uint8Array {
  const parts: Array<number> = [];
  const sprm = (op: number, ...operand: Array<number>): void => {
    parts.push(op & 0xff, (op >> 8) & 0xff, ...operand);
  };
  if (run.bold) sprm(0x0835, 0x01); // sprmCFBold
  if (run.italic) sprm(0x0836, 0x01); // sprmCFItalic
  if (run.underlineKul) sprm(0x2a3e, run.underlineKul & 0xff); // sprmCKul
  if (run.sizeHalfPts) sprm(0x4a43, run.sizeHalfPts & 0xff, (run.sizeHalfPts >> 8) & 0xff); // sprmCHps
  if (run.picOffset !== undefined) {
    sprm(0x0855, 0x01); // sprmCFSpec — special character (picture)
    const fc = run.picOffset >>> 0;
    sprm(0x6a03, fc & 0xff, (fc >> 8) & 0xff, (fc >> 16) & 0xff, (fc >> 24) & 0xff); // sprmCPicLocation
  }
  return Uint8Array.from(parts);
}

// §2.6.2 — the paragraph sprm list. Indents are signed 2-byte twips.
function buildPapxGrpprl(run: DocParaFormat): Uint8Array {
  const parts: Array<number> = [];
  const sprm = (op: number, ...operand: Array<number>): void => {
    parts.push(op & 0xff, (op >> 8) & 0xff, ...operand);
  };
  const i16 = (v: number): Array<number> => [v & 0xff, (v >> 8) & 0xff];
  if (run.jc !== undefined) sprm(0x2403, run.jc & 0xff); // sprmPJc
  if (run.indentLeftTwips !== undefined) sprm(0x840f, ...i16(run.indentLeftTwips)); // sprmPDxaLeft
  if (run.indentRightTwips !== undefined) sprm(0x840e, ...i16(run.indentRightTwips)); // sprmPDxaRight
  if (run.indentFirstTwips !== undefined) sprm(0x8411, ...i16(run.indentFirstTwips)); // sprmPDxaLeft1
  if (run.spaceBeforeTwips !== undefined) sprm(0xa413, ...i16(run.spaceBeforeTwips)); // sprmPDyaBefore
  if (run.spaceAfterTwips !== undefined) sprm(0xa414, ...i16(run.spaceAfterTwips)); // sprmPDyaAfter
  // sprmTDefTable (a long-operand sprm) goes before the table flags so the reader
  // must skip its 2-byte-length operand correctly to still reach sprmPFTtp.
  if (run.cellEdgesTwips) {
    const content = [(run.cellEdgesTwips.length - 1) & 0xff, ...run.cellEdgesTwips.flatMap(i16)];
    // After the boundaries, the itcMac × TC80 cell descriptors (DOC-11).
    if (run.cellTc) {
      const brc = (e?: DocBrcInput): Array<number> =>
        e ? [e.widthEighthPt ?? 8, e.brcType ?? 1, e.ico ?? 1, 0] : [0, 0, 0, 0];
      for (const tc of run.cellTc) {
        const tcgrf = tc.vMerge === 'restart' ? 0x60 : tc.vMerge === 'continue' ? 0x20 : 0;
        content.push(
          ...i16(tcgrf),
          ...i16(0), // wWidth
          ...brc(tc.borders?.top),
          ...brc(tc.borders?.left),
          ...brc(tc.borders?.bottom),
          ...brc(tc.borders?.right),
        );
      }
    }
    const cb = content.length + 1; // a 2-byte count that includes the count field
    sprm(0xd608, cb & 0xff, (cb >> 8) & 0xff, ...content); // sprmTDefTable
  }
  if (run.listIlvl !== undefined) sprm(0x260a, run.listIlvl & 0xff); // sprmPIlvl
  if (run.listIlfo !== undefined) sprm(0x460b, run.listIlfo & 0xff, (run.listIlfo >> 8) & 0xff); // sprmPIlfo
  if (run.inTable) sprm(0x2416, 0x01); // sprmPFInTable
  if (run.rowEnd) sprm(0x2417, 0x01); // sprmPFTtp
  return Uint8Array.from(parts);
}

// §2.8.5/§2.8.6 PlcBte — a one-entry PLC: two FCs (the covered range) and one
// page number naming the FKP.
function buildPlcBte(fcStart: number, fcEnd: number, pageNumber: number): Uint8Array {
  const plc = new Uint8Array(12);
  const dv = new DataView(plc.buffer);
  dv.setUint32(0, fcStart, true);
  dv.setUint32(4, fcEnd, true);
  dv.setUint32(8, pageNumber, true);
  return plc;
}

function encodeUtf16(s: string): Uint8Array {
  const out = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = (c >> 8) & 0xff;
  }
  return out;
}

// Windows-1252 high range (0x80–0x9F); the rest is Latin-1 (== Unicode 0–0xFF).
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152,
  0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122,
  0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

function encodeCp1252(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) out[i] = c;
    else {
      const idx = CP1252_HIGH.indexOf(c);
      out[i] = idx >= 0 ? 0x80 + idx : 0x3f; // unmappable → '?'
    }
  }
  return out;
}
