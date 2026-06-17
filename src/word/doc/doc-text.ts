// Legacy `.doc` text retrieval (DOC-1) — pulls the document text out of a Word
// 97–2003 binary file. A `.doc` is a CFB (the same OLE2 container as `.xls`)
// holding a `WordDocument` stream and a table stream (`0Table` or `1Table`). The
// text is NOT stored contiguously: a piece table (the CLX, in the table stream)
// maps character positions to byte offsets in the `WordDocument` stream, and each
// piece is either 16-bit Unicode or 8-bit Windows-1252 ("compressed"). This walks
// the FIB → piece table → pieces and returns the concatenated main-document text.
//
// Spec: [MS-DOC] §2.5 (FIB), §2.8.35 (PlcPcd), §2.9.38 (Clx), §2.9.73
// (FcCompressed), §2.9.177 (Pcd). Every offset is bounds-checked; a structurally
// odd file yields empty text rather than throwing — the reader degrades, never
// crashes.

import { openCfb } from '@/core/ole/cfb';

const WIDENT = 0xa5ec; // FibBase.wIdent — the "this is a Word doc" magic
const OFF_FLAGS1 = 0x0a; // FibBase bit field (fEncrypted, fWhichTblStm, …)
const OFF_CCPTEXT = 0x4c; // fibRgLw97.ccpText — main-document char count
const OFF_FCCLX = 0x1a2; // fibRgFcLcb97.fcClx — CLX offset in the table stream
const OFF_LCBCLX = 0x1a6; // fibRgFcLcb97.lcbClx — CLX byte length
const FLAG_ENCRYPTED = 0x0100; // FibBase.fEncrypted
const FLAG_WHICH_TBL = 0x0200; // FibBase.fWhichTblStm — 1 ⇒ 1Table, 0 ⇒ 0Table

const FC_COMPRESSED = 0x40000000; // FcCompressed.fCompressed — 8-bit piece
const FC_MASK = 0x3fffffff; // FcCompressed.fc — the 30-bit offset field

const MAX_TEXT = 1 << 24; // 16M-char guard against a crafted piece table

export interface DocText {
  readonly text: string;
  // The file is encrypted/obfuscated — text cannot be read without the key.
  readonly encrypted: boolean;
}

// The `WordDocument` stream of a `.doc` → its main-document text. Returns empty
// text when the stream is missing, not a Word file, encrypted, or has no piece
// table (older Word 6/95 without a CLX is out of scope).
export function extractDocText(bytes: Uint8Array): DocText {
  const cfb = openCfb(bytes);
  const wd = cfb.readStream('WordDocument');
  if (!wd || wd.length < OFF_LCBCLX + 4) return { text: '', encrypted: false };
  if (u16(wd, 0) !== WIDENT) return { text: '', encrypted: false };

  const flags = u16(wd, OFF_FLAGS1);
  if ((flags & FLAG_ENCRYPTED) !== 0) return { text: '', encrypted: true };

  // The CLX lives in whichever table stream fWhichTblStm selects.
  const primary = (flags & FLAG_WHICH_TBL) !== 0 ? '1Table' : '0Table';
  const table =
    cfb.readStream(primary) ?? cfb.readStream(primary === '1Table' ? '0Table' : '1Table');
  if (!table) return { text: '', encrypted: false };

  const fcClx = u32(wd, OFF_FCCLX);
  const lcbClx = u32(wd, OFF_LCBCLX);
  if (lcbClx === 0 || fcClx + lcbClx > table.length) return { text: '', encrypted: false };

  const plcPcd = findPlcPcd(table.subarray(fcClx, fcClx + lcbClx));
  if (!plcPcd) return { text: '', encrypted: false };

  const ccpText = u32(wd, OFF_CCPTEXT);
  return { text: readPieces(wd, plcPcd, ccpText), encrypted: false };
}

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

// §2.8.35 PlcPcd — a PLC: (n+1) character positions (u32) then n Pcds (8 bytes).
// Piece i spans CP[i]..CP[i+1]; its Pcd.fc gives the byte offset and 8-vs-16-bit
// encoding of those characters in the WordDocument stream.
function readPieces(wd: Uint8Array, plc: Uint8Array, ccpText: number): string {
  if (plc.length < 4 + 8) return '';
  const n = Math.floor((plc.length - 4) / 12);
  if (n <= 0) return '';
  const pcdBase = (n + 1) * 4;
  let out = '';
  for (let i = 0; i < n && out.length < MAX_TEXT; i++) {
    const cpStart = u32(plc, i * 4);
    const cpEnd = u32(plc, (i + 1) * 4);
    let nChars = cpEnd - cpStart;
    if (nChars <= 0) continue;
    if (out.length + nChars > MAX_TEXT) nChars = MAX_TEXT - out.length;

    const fc = u32(plc, pcdBase + i * 8 + 2); // Pcd: u16 flags, then FcCompressed
    const compressed = (fc & FC_COMPRESSED) !== 0;
    const offset = compressed ? (fc & FC_MASK) >>> 1 : fc & FC_MASK;
    out += compressed ? decodeCp1252(wd, offset, nChars) : decodeUtf16(wd, offset, nChars);
  }
  // The piece table also covers footnotes/headers/… after the main text; keep
  // only the first ccpText characters (the main document).
  return ccpText > 0 && ccpText < out.length ? out.slice(0, ccpText) : out;
}

// Windows-1252 high range (0x80–0x9F); the rest of 0x00–0xFF maps to U+0000–00FF.
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152,
  0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122,
  0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178,
];

function decodeCp1252(d: Uint8Array, off: number, nChars: number): string {
  const end = Math.min(d.length, off + nChars);
  let s = '';
  for (let i = off; i < end; i++) {
    const b = d[i]!;
    s += String.fromCharCode(b >= 0x80 && b <= 0x9f ? CP1252_HIGH[b - 0x80]! : b);
  }
  return s;
}

function decodeUtf16(d: Uint8Array, off: number, nChars: number): string {
  const end = Math.min(d.length - 1, off + nChars * 2 - 1);
  let s = '';
  for (let i = off; i < end; i += 2) s += String.fromCharCode(d[i]! | (d[i + 1]! << 8));
  return s;
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
