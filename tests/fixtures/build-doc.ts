// Minimal Word 97–2003 `.doc` builder for tests — packs a `WordDocument` stream
// (the FIB fields the reader needs, the piece text and, optionally, a CHPX FKP
// page) and a table stream (the CLX piece table plus, optionally, the PlcfBteChpx
// bin table) into a CFB container, so the `.doc` reader has a deterministic input
// without a checked-in binary. Mirrors build-xls.ts: it writes exactly the bytes
// doc-text.ts reads back (FibBase offsets, PlcPcd, FcCompressed, FKP, sprm), so
// the round trip validates the offset/compression/formatting logic rather than
// restating it.
//
// CHPX formatting is synthesised against the FIRST piece's FC mapping, so tests
// that pass `formatRuns` use a single piece (enough to exercise the reader's
// FKP → sprm → FC→CP path; the reader itself handles the multi-piece case).

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
  const totalChars = placed.reduce((sum, p) => sum + p.chars, 0);

  // PlcPcd: (n+1) character positions then n 8-byte Pcds.
  const n = placed.length;
  const plc = new Uint8Array((n + 1) * 4 + n * 8);
  const pv = new DataView(plc.buffer);
  let cp = 0;
  pv.setUint32(0, 0, true);
  for (let i = 0; i < n; i++) {
    cp += placed[i]!.chars;
    pv.setUint32((i + 1) * 4, cp, true);
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

  // Optional CHPX formatting: an FKP page in the WordDocument stream (after the
  // text, 512-aligned) plus a PlcfBteChpx that points at it from the table stream.
  let wdLength = cursor;
  let fkpPage: Uint8Array | undefined;
  let fkpOffset = 0;
  let plcfBteChpx: Uint8Array | undefined;
  const first = placed[0];
  if (opts.formatRuns && opts.formatRuns.length > 0 && first) {
    const step = first.compressed ? 1 : 2;
    fkpOffset = Math.ceil(cursor / FKP_SIZE) * FKP_SIZE;
    fkpPage = buildChpxFkp(first.offset, step, opts.formatRuns);
    plcfBteChpx = buildPlcfBteChpx(
      first.offset,
      first.offset + totalChars * step,
      fkpOffset / FKP_SIZE,
    );
    wdLength = fkpOffset + FKP_SIZE;
  }

  // Table stream: the CLX, then (4-aligned) the PlcfBteChpx when present.
  let tableLength = clx.length;
  let plcfOffset = 0;
  if (plcfBteChpx) {
    plcfOffset = Math.ceil(clx.length / 4) * 4;
    tableLength = plcfOffset + plcfBteChpx.length;
  }
  const table = new Uint8Array(tableLength);
  table.set(clx, 0);
  if (plcfBteChpx) table.set(plcfBteChpx, plcfOffset);

  // WordDocument stream: the FIB fields the reader probes, the piece bytes and
  // the optional FKP page.
  const wd = new Uint8Array(wdLength);
  const wv = new DataView(wd.buffer);
  wv.setUint16(0x00, 0xa5ec, true); // wIdent
  wv.setUint16(0x02, 0x00c1, true); // nFib (Word 97)
  wv.setUint16(
    0x0a,
    (whichTable === '1Table' ? 0x0200 : 0x0000) | (opts.encrypted ? 0x0100 : 0x0000),
    true,
  ); // fWhichTblStm | fEncrypted
  wv.setUint32(0x4c, totalChars, true); // ccpText
  if (plcfBteChpx) {
    wv.setUint32(0xfa, plcfOffset, true); // fcPlcfBteChpx
    wv.setUint32(0xfe, plcfBteChpx.length, true); // lcbPlcfBteChpx
  }
  wv.setUint32(0x1a2, 0, true); // fcClx — CLX at offset 0 of the table stream
  wv.setUint32(0x1a6, clx.length, true); // lcbClx
  for (const p of placed) wd.set(p.bytes, p.offset);
  if (fkpPage) wd.set(fkpPage, fkpOffset);

  return buildCfb([
    { name: 'WordDocument', data: wd },
    { name: whichTable, data: table },
  ]);
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

// §2.6.1 — the sprm list for a character run.
function buildChpxGrpprl(run: DocFormatRun): Uint8Array {
  const parts: Array<number> = [];
  const sprm = (op: number, ...operand: Array<number>): void => {
    parts.push(op & 0xff, (op >> 8) & 0xff, ...operand);
  };
  if (run.bold) sprm(0x0835, 0x01); // sprmCFBold
  if (run.italic) sprm(0x0836, 0x01); // sprmCFItalic
  if (run.underlineKul) sprm(0x2a3e, run.underlineKul & 0xff); // sprmCKul
  if (run.sizeHalfPts) sprm(0x4a43, run.sizeHalfPts & 0xff, (run.sizeHalfPts >> 8) & 0xff); // sprmCHps
  return Uint8Array.from(parts);
}

// §2.8.6 PlcBteChpx — a one-entry PLC: two FCs (the covered range) and one page
// number naming the FKP.
function buildPlcfBteChpx(fcStart: number, fcEnd: number, pageNumber: number): Uint8Array {
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
