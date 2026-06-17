// Minimal Word 97–2003 `.doc` builder for tests — packs a `WordDocument` stream
// (the FIB fields the reader needs + the piece text) and a table stream (the CLX
// piece table) into a CFB container, so the `.doc` reader has a deterministic
// input without a checked-in binary. Mirrors build-xls.ts: it writes exactly the
// bytes doc-text.ts reads back (FibBase offsets, PlcPcd, FcCompressed), so the
// round trip validates the offset/compression logic rather than restating it.

import { buildCfb } from './build-cfb';

export interface DocPiece {
  readonly text: string;
  // true → stored as 8-bit Windows-1252 ("compressed"); false → 16-bit UTF-16LE.
  readonly compressed: boolean;
}

// Where piece text starts in the WordDocument stream — past the FIB fields the
// reader probes (fcClx/lcbClx sit at 0x1A2/0x1A6).
const TEXT_BASE = 0x200;

export function buildDoc(
  pieces: ReadonlyArray<DocPiece>,
  opts: { readonly whichTable?: '0Table' | '1Table'; readonly encrypted?: boolean } = {},
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

  // WordDocument stream: the FIB fields the reader probes + the piece bytes.
  const wd = new Uint8Array(cursor);
  const wv = new DataView(wd.buffer);
  wv.setUint16(0x00, 0xa5ec, true); // wIdent
  wv.setUint16(0x02, 0x00c1, true); // nFib (Word 97)
  wv.setUint16(
    0x0a,
    (whichTable === '1Table' ? 0x0200 : 0x0000) | (opts.encrypted ? 0x0100 : 0x0000),
    true,
  ); // fWhichTblStm | fEncrypted
  wv.setUint32(0x4c, totalChars, true); // ccpText
  wv.setUint32(0x1a2, 0, true); // fcClx — CLX at offset 0 of the table stream
  wv.setUint32(0x1a6, clx.length, true); // lcbClx
  for (const p of placed) wd.set(p.bytes, p.offset);

  return buildCfb([
    { name: 'WordDocument', data: wd },
    { name: whichTable, data: clx },
  ]);
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
