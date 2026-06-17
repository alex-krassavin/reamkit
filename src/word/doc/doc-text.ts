// Legacy `.doc` text + character formatting (DOC-1/DOC-2) — pulls the document
// text out of a Word 97–2003 binary file together with its run-level formatting.
// A `.doc` is a CFB (the same OLE2 container as `.xls`) holding a `WordDocument`
// stream and a table stream (`0Table` / `1Table`). Text is not stored
// contiguously: a piece table (the CLX) maps character positions (CPs) to byte
// offsets (FCs) in the `WordDocument` stream, each piece either 16-bit Unicode or
// 8-bit Windows-1252 ("compressed"). Character formatting lives elsewhere again —
// CHPX runs packed into 512-byte FKP pages, located through the PlcfBteChpx bin
// table — and is keyed by FC. This walks both and emits maximal same-formatting
// runs of text.
//
// Spec: [MS-DOC] §2.5 (FIB), §2.8.35 (PlcPcd), §2.9.38 (Clx), §2.9.73
// (FcCompressed), §2.9.177 (Pcd), §2.8.6 (PlcBteChpx), §2.9.55 (ChpxFkp), §2.6.1
// (sprm). Every offset is bounds-checked; a structurally odd file yields empty
// text rather than throwing — the reader degrades, never crashes.

import { openCfb } from '@/core/ole/cfb';

const WIDENT = 0xa5ec; // FibBase.wIdent — the "this is a Word doc" magic
const OFF_FLAGS1 = 0x0a; // FibBase bit field (fEncrypted, fWhichTblStm, …)
const OFF_CCPTEXT = 0x4c; // fibRgLw97.ccpText — main-document char count
const OFF_FCPLCFBTECHPX = 0xfa; // fibRgFcLcb97.fcPlcfBteChpx — CHPX bin table
const OFF_LCBPLCFBTECHPX = 0xfe; // fibRgFcLcb97.lcbPlcfBteChpx
const OFF_FCCLX = 0x1a2; // fibRgFcLcb97.fcClx — CLX offset in the table stream
const OFF_LCBCLX = 0x1a6; // fibRgFcLcb97.lcbClx — CLX byte length
const FLAG_ENCRYPTED = 0x0100; // FibBase.fEncrypted
const FLAG_WHICH_TBL = 0x0200; // FibBase.fWhichTblStm — 1 ⇒ 1Table, 0 ⇒ 0Table

const FC_COMPRESSED = 0x40000000; // FcCompressed.fCompressed — 8-bit piece
const FC_MASK = 0x3fffffff; // FcCompressed.fc — the 30-bit offset field

const FKP_SIZE = 512; // an FKP is one 512-byte page
const PN_MASK = 0x3fffff; // PnFkpChpx.pn — the low 22 bits address the page

// sprm opcodes (§2.6.1) — the character-property modifiers we read.
const SPRM_C_FBOLD = 0x0835;
const SPRM_C_FITALIC = 0x0836;
const SPRM_C_KUL = 0x2a3e; // underline kind
const SPRM_C_HPS = 0x4a43; // font size, half-points
// Operand size in bytes by `spra` (sprm >> 13); 6 is variable (length-prefixed).
const SPRA_LEN = [1, 1, 2, 4, 2, 2, 0, 3];

const MAX_TEXT = 1 << 24; // 16M-char guard against a crafted piece table

// Raw character properties read off the CHPX sprms (engine-neutral; the reader
// maps them to the document-model RunProperties).
export interface DocCharProps {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underlineKul?: number; // Word `kul` underline code (0 = none)
  readonly sizeHalfPts?: number; // font size in half-points
}

export interface DocRun {
  readonly text: string;
  readonly props: DocCharProps;
}

export interface DocContent {
  readonly runs: ReadonlyArray<DocRun>;
  // The file is encrypted/obfuscated — text cannot be read without the key.
  readonly encrypted: boolean;
}

interface Piece {
  readonly cpStart: number;
  readonly cpEnd: number;
  readonly fc: number; // byte offset of the piece's first character
  readonly compressed: boolean;
}

interface ChpxRun {
  readonly fcStart: number;
  readonly fcEnd: number;
  readonly props: DocCharProps;
}

// The `WordDocument` stream of a `.doc` → its main-document text as formatted
// runs. Returns no runs when the stream is missing, not a Word file, encrypted,
// or has no piece table (older Word 6/95 without a CLX is out of scope).
export function extractDocContent(bytes: Uint8Array): DocContent {
  const cfb = openCfb(bytes);
  const wd = cfb.readStream('WordDocument');
  if (!wd || wd.length < OFF_LCBCLX + 4) return EMPTY;
  if (u16(wd, 0) !== WIDENT) return EMPTY;

  const flags = u16(wd, OFF_FLAGS1);
  if ((flags & FLAG_ENCRYPTED) !== 0) return { runs: [], encrypted: true };

  // The CLX + CHPX bin table live in whichever table stream fWhichTblStm selects.
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
  const chpx = parseChpx(wd, table);
  const ccpText = u32(wd, OFF_CCPTEXT);
  return { runs: buildRuns(wd, pieces, chpx, ccpText), encrypted: false };
}

const EMPTY: DocContent = { runs: [], encrypted: false };

// §2.9.38 Clx — zero or more Prc (each a 0x01 byte + i16 length + grpprl) then
// the Pcdt: a 0x02 byte, a u32 length, and the PlcPcd of that length.
function findPlcPcd(clx: Uint8Array): Uint8Array | undefined {
  let p = 0;
  while (p < clx.length) {
    const clxt = clx[p];
    if (clxt === 0x01) {
      if (p + 3 > clx.length) return undefined;
      p += 3 + u16(clx, p + 1); // skip the property modifier
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

// §2.8.35 PlcPcd — a PLC: (n+1) CPs (u32) then n Pcds (8 bytes). Piece i spans
// CP[i]..CP[i+1]; its Pcd.fc gives the byte offset and 8-vs-16-bit encoding.
function parsePieces(plc: Uint8Array): Array<Piece> {
  if (plc.length < 4 + 8) return [];
  const n = Math.floor((plc.length - 4) / 12);
  if (n <= 0) return [];
  const pcdBase = (n + 1) * 4;
  const pieces: Array<Piece> = [];
  for (let i = 0; i < n; i++) {
    const fcRaw = u32(plc, pcdBase + i * 8 + 2); // Pcd: u16 flags, then FcCompressed
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

// §2.8.6 PlcBteChpx — a PLC of (n+1) FCs then n PnFkpChpx; each names a 512-byte
// FKP page in the WordDocument stream that holds the CHPX runs for an FC range.
function parseChpx(wd: Uint8Array, table: Uint8Array): Array<ChpxRun> {
  const fcPlc = u32(wd, OFF_FCPLCFBTECHPX);
  const lcbPlc = u32(wd, OFF_LCBPLCFBTECHPX);
  if (lcbPlc < 4 + 4 || fcPlc + lcbPlc > table.length) return [];
  const plc = table.subarray(fcPlc, fcPlc + lcbPlc);
  const n = Math.floor((plc.length - 4) / 8);
  if (n <= 0) return [];
  const pnBase = (n + 1) * 4;
  const out: Array<ChpxRun> = [];
  for (let i = 0; i < n; i++) {
    const pn = u32(plc, pnBase + i * 4) & PN_MASK;
    const pageOff = pn * FKP_SIZE;
    if (pageOff + FKP_SIZE > wd.length) continue;
    parseChpxFkp(wd.subarray(pageOff, pageOff + FKP_SIZE), out);
  }
  out.sort((a, b) => a.fcStart - b.fcStart);
  return out;
}

// §2.9.55 ChpxFkp — (crun+1) FCs, then crun 1-byte word-offsets to each run's
// CHPX (0 = no formatting), and crun in the last byte. A CHPX is a u8 length then
// that many grpprl bytes.
function parseChpxFkp(page: Uint8Array, out: Array<ChpxRun>): void {
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

// §2.6.1 — a grpprl is a list of sprms; each is a u16 opcode whose `spra` bits
// give the operand size. We pick out the character-formatting ones and skip the
// rest (advancing by their operand size keeps the walk aligned).
function decodeChpxGrpprl(d: Uint8Array): DocCharProps {
  let bold: boolean | undefined;
  let italic: boolean | undefined;
  let underlineKul: number | undefined;
  let sizeHalfPts: number | undefined;
  let p = 0;
  while (p + 2 <= d.length) {
    const sprm = u16(d, p);
    p += 2;
    const spra = sprm >> 13;
    let opStart = p;
    let opLen = SPRA_LEN[spra]!;
    if (spra === 6) {
      if (p >= d.length) break;
      opLen = d[p]!;
      opStart = p + 1;
    }
    if (opStart + opLen > d.length) break;
    switch (sprm) {
      case SPRM_C_FBOLD:
        bold = toggleBool(d[opStart]!);
        break;
      case SPRM_C_FITALIC:
        italic = toggleBool(d[opStart]!);
        break;
      case SPRM_C_KUL:
        underlineKul = d[opStart]!;
        break;
      case SPRM_C_HPS:
        sizeHalfPts = u16(d, opStart);
        break;
    }
    p = opStart + opLen;
  }
  return {
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(underlineKul !== undefined ? { underlineKul } : {}),
    ...(sizeHalfPts !== undefined ? { sizeHalfPts } : {}),
  };
}

// A toggle sprm operand: 0 off, 1 on, 0x81 toggle (no base ⇒ on), 0x80 inherit
// (no change ⇒ undefined).
function toggleBool(v: number): boolean | undefined {
  if (v === 0) return false;
  if (v === 1 || v === 0x81) return true;
  return undefined;
}

// Walk the pieces character by character, attach each character's CHPX props
// (looked up by its FC), and coalesce runs of identical formatting. Stops at
// ccpText (the main document; the piece table also covers footnotes/headers/…).
function buildRuns(
  wd: Uint8Array,
  pieces: ReadonlyArray<Piece>,
  chpx: ReadonlyArray<ChpxRun>,
  ccpText: number,
): Array<DocRun> {
  const limit = ccpText > 0 ? Math.min(ccpText, MAX_TEXT) : MAX_TEXT;
  const runs: Array<DocRun> = [];
  let cur = '';
  let curProps: DocCharProps = {};
  let cp = 0;
  for (const piece of pieces) {
    if (cp >= limit) break;
    const step = piece.compressed ? 1 : 2;
    const nChars = piece.cpEnd - piece.cpStart;
    for (let k = 0; k < nChars && cp < limit; k++, cp++) {
      const fc = piece.fc + k * step;
      const ch = piece.compressed ? decodeByte(wd, fc) : decodeUnit(wd, fc);
      const props = lookupChpx(chpx, fc);
      if (cur.length === 0) {
        curProps = props;
      } else if (!sameProps(props, curProps)) {
        runs.push({ text: cur, props: curProps });
        cur = '';
        curProps = props;
      }
      cur += ch;
    }
  }
  if (cur.length > 0) runs.push({ text: cur, props: curProps });
  return runs;
}

// Binary search the FC-sorted CHPX runs for the one covering `fc`.
function lookupChpx(chpx: ReadonlyArray<ChpxRun>, fc: number): DocCharProps {
  let lo = 0;
  let hi = chpx.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const r = chpx[mid]!;
    if (fc < r.fcStart) hi = mid - 1;
    else if (fc >= r.fcEnd) lo = mid + 1;
    else return r.props;
  }
  return {};
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

function decodeByte(d: Uint8Array, fc: number): string {
  const b = d[fc];
  if (b === undefined) return '';
  return String.fromCharCode(b >= 0x80 && b <= 0x9f ? CP1252_HIGH[b - 0x80]! : b);
}

function decodeUnit(d: Uint8Array, fc: number): string {
  if (fc + 1 >= d.length) return '';
  return String.fromCharCode(d[fc]! | (d[fc + 1]! << 8));
}

function u16(d: Uint8Array, off: number): number {
  return (d[off] ?? 0) | ((d[off + 1] ?? 0) << 8);
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
