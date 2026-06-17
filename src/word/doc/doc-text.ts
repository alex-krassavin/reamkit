// Legacy `.doc` text + formatting (DOC-1/2/3) — pulls the document out of a Word
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
// splits paragraphs at the CR mark, attaches each paragraph's PAPX, and emits
// maximal same-formatting runs.
//
// Spec: [MS-DOC] §2.5 (FIB), §2.8.35 (PlcPcd), §2.9.38 (Clx), §2.9.73
// (FcCompressed), §2.8.6 (PlcBteChpx), §2.8.5 (PlcBtePapx), §2.9.55 (ChpxFkp),
// §2.9.23 (PapxFkp), §2.6.1 (sprm). Every offset is bounds-checked; a
// structurally odd file yields no text rather than throwing.

import { openCfb } from '@/core/ole/cfb';

const WIDENT = 0xa5ec; // FibBase.wIdent — the "this is a Word doc" magic
const OFF_FLAGS1 = 0x0a; // FibBase bit field (fEncrypted, fWhichTblStm, …)
const OFF_CCPTEXT = 0x4c; // fibRgLw97.ccpText — main-document char count
const OFF_FCPLCFBTECHPX = 0xfa; // fibRgFcLcb97.fcPlcfBteChpx — CHPX bin table
const OFF_LCBPLCFBTECHPX = 0xfe;
const OFF_FCPLCFBTEPAPX = 0x102; // fibRgFcLcb97.fcPlcfBtePapx — PAPX bin table
const OFF_LCBPLCFBTEPAPX = 0x106;
const OFF_FCCLX = 0x1a2; // fibRgFcLcb97.fcClx — CLX offset in the table stream
const OFF_LCBCLX = 0x1a6; // fibRgFcLcb97.lcbClx — CLX byte length
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
const SPRM_P_JC = 0x2403; // paragraph justification
const SPRM_P_DXA_RIGHT = 0x840e; // right indent (twips)
const SPRM_P_DXA_LEFT = 0x840f; // left indent (twips)
const SPRM_P_DXA_LEFT1 = 0x8411; // first-line indent (twips, signed)
const SPRM_P_DYA_BEFORE = 0xa413; // space before (twips)
const SPRM_P_DYA_AFTER = 0xa414; // space after (twips)
const SPRA_LEN = [1, 1, 2, 4, 2, 2, 0, 3];

const MAX_TEXT = 1 << 24; // 16M-char guard against a crafted piece table

export interface DocCharProps {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underlineKul?: number; // Word `kul` underline code (0 = none)
  readonly sizeHalfPts?: number; // font size in half-points
}

export interface DocParaProps {
  readonly jc?: number; // Word justification (0 left, 1 center, 2 right, 3 both, 4 distribute)
  readonly indentLeftTwips?: number;
  readonly indentRightTwips?: number;
  readonly indentFirstTwips?: number; // signed (negative = hanging)
  readonly spaceBeforeTwips?: number;
  readonly spaceAfterTwips?: number;
}

export interface DocRun {
  readonly text: string;
  readonly props: DocCharProps;
}

export interface DocParagraph {
  readonly runs: ReadonlyArray<DocRun>;
  readonly props: DocParaProps;
}

export interface DocContent {
  readonly paragraphs: ReadonlyArray<DocParagraph>;
  // The file is encrypted/obfuscated — text cannot be read without the key.
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
}

const EMPTY: DocContent = { paragraphs: [], encrypted: false };

// The `WordDocument` stream of a `.doc` → its main-document paragraphs. Returns
// none when the stream is missing, not a Word file, encrypted, or has no piece
// table (older Word 6/95 without a CLX is out of scope).
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
  return { paragraphs: buildParagraphs(wd, pieces, chpx, papx, ccpText), encrypted: false };
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
    let props: DocCharProps = {};
    const wordOff = page[(crun + 1) * 4 + i]!;
    if (wordOff !== 0) {
      const chpxOff = wordOff * 2;
      const cb = chpxOff < FKP_SIZE ? page[chpxOff]! : 0;
      if (cb > 0 && chpxOff + 1 + cb <= FKP_SIZE) {
        props = decodeChpxGrpprl(page.subarray(chpxOff + 1, chpxOff + 1 + cb));
      }
    }
    out.push({ fcStart, fcEnd, props });
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
    let props: DocParaProps = {};
    const papxOff = page[bxBase + i * 13]! * 2;
    if (papxOff !== 0 && papxOff < FKP_SIZE) {
      const cb = page[papxOff]!;
      const size = cb !== 0 ? 2 * cb - 1 : 2 * (page[papxOff + 1] ?? 0);
      const start = cb !== 0 ? papxOff + 1 : papxOff + 2;
      if (size >= 2 && start + size <= FKP_SIZE) {
        props = decodePapxGrpprl(page.subarray(start + 2, start + size)); // skip the istd
      }
    }
    out.push({ fcStart, fcEnd, props });
  }
}

// §2.6.1 — pick the character sprms out of a grpprl, skipping the rest by their
// spra-derived operand size so the walk stays aligned.
function decodeChpxGrpprl(d: Uint8Array): DocCharProps {
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let underlineKul: number | undefined;
  let sizeHalfPts: number | undefined;
  for (const s of sprms(d)) {
    switch (s.sprm) {
      case SPRM_C_FBOLD:
        bold = toggleBool(d[s.op]!);
        break;
      case SPRM_C_FITALIC:
        italic = toggleBool(d[s.op]!);
        break;
      case SPRM_C_KUL:
        underlineKul = d[s.op]!;
        break;
      case SPRM_C_HPS:
        sizeHalfPts = u16(d, s.op);
        break;
    }
  }
  return {
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underlineKul !== undefined ? { underlineKul } : {}),
    ...(sizeHalfPts !== undefined ? { sizeHalfPts } : {}),
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
    }
  }
  return {
    ...(jc !== undefined ? { jc } : {}),
    ...(indentLeftTwips !== undefined ? { indentLeftTwips } : {}),
    ...(indentRightTwips !== undefined ? { indentRightTwips } : {}),
    ...(indentFirstTwips !== undefined ? { indentFirstTwips } : {}),
    ...(spaceBeforeTwips !== undefined ? { spaceBeforeTwips } : {}),
    ...(spaceAfterTwips !== undefined ? { spaceAfterTwips } : {}),
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
      if (p >= d.length) break;
      len = d[p]!;
      op = p + 1;
    }
    if (op + len > d.length) break;
    yield { sprm, op };
    p = op + len;
  }
}

// A toggle sprm operand: 0 off, 1 on, 0x81 toggle (no base ⇒ on), 0x80 inherit.
function toggleBool(v: number): boolean | undefined {
  if (v === 0) return false;
  if (v === 1 || v === 0x81) return true;
  return undefined;
}

// Walk the pieces character by character: attach each character's CHPX (by FC),
// split paragraphs at the CR mark (attaching each paragraph's PAPX, also by FC),
// clean the in-paragraph control characters, and coalesce same-formatting runs.
function buildParagraphs(
  wd: Uint8Array,
  pieces: ReadonlyArray<Piece>,
  chpx: ReadonlyArray<BteRun<DocCharProps>>,
  papx: ReadonlyArray<BteRun<DocParaProps>>,
  ccpText: number,
): Array<DocParagraph> {
  const limit = ccpText > 0 ? Math.min(ccpText, MAX_TEXT) : MAX_TEXT;
  const paragraphs: Array<DocParagraph> = [];
  let runs: Array<DocRun> = [];
  let cur = '';
  let curProps: DocCharProps = {};
  let cp = 0;
  let lastFc = 0;

  const flushRun = (): void => {
    if (cur.length > 0) {
      runs.push({ text: cur, props: curProps });
      cur = '';
    }
  };
  const endParagraph = (fc: number): void => {
    flushRun();
    paragraphs.push({ runs, props: lookup(papx, fc) ?? {} });
    runs = [];
  };

  for (const piece of pieces) {
    if (cp >= limit) break;
    const step = piece.compressed ? 1 : 2;
    const nChars = piece.cpEnd - piece.cpStart;
    for (let k = 0; k < nChars && cp < limit; k++, cp++) {
      const fc = piece.fc + k * step;
      lastFc = fc;
      const code = piece.compressed ? byteCode(wd, fc) : unitCode(wd, fc);
      if (code < 0) continue;
      if (code === CR) {
        endParagraph(fc);
        continue;
      }
      const ch =
        code === 0x09 || code === 0x0b ? ' ' : code < 0x20 ? '' : String.fromCharCode(code);
      if (ch.length === 0) continue; // drop the other control characters
      const props = lookup(chpx, fc) ?? {};
      if (cur.length === 0) curProps = props;
      else if (!sameProps(props, curProps)) {
        flushRun();
        curProps = props;
      }
      cur += ch;
    }
  }
  endParagraph(lastFc); // the final paragraph (empty if the text ended with CR)
  return paragraphs;
}

// Binary search the FC-sorted runs for the one covering `fc`.
function lookup<TProps>(runs: ReadonlyArray<BteRun<TProps>>, fc: number): TProps | undefined {
  let lo = 0;
  let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const r = runs[mid]!;
    if (fc < r.fcStart) hi = mid - 1;
    else if (fc >= r.fcEnd) lo = mid + 1;
    else return r.props;
  }
  return undefined;
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
