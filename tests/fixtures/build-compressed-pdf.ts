// Build a minimal PDF that stores its objects in an OBJECT STREAM (/Type/ObjStm)
// and its cross-reference as an XREF STREAM (/Type/XRef) — the compressed,
// PDF-1.5+ form that Ream's own writer never emits, so it exercises the reader's
// EP7 path. Both streams are stored uncompressed (a legal encoding), so the
// fixture needs no zlib and its byte offsets stay easy to compute.
//
// Objects 1 (Catalog), 2 (Pages), 3 (Page) live inside object stream 4; object 5
// is the xref stream. `brokenStartxref` points startxref at junk so the reader
// must fall back to the brute-force scan (which then indexes the object stream).

const enc = new TextEncoder();

export function buildCompressedPdf(opts: { brokenStartxref?: boolean } = {}): Uint8Array {
  const obj1 = '<</Type/Catalog/Pages 2 0 R>>';
  const obj2 = '<</Type/Pages/Kids[3 0 R]/Count 1>>';
  const obj3 = '<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 200]>>';
  const objData = `${obj1}\n${obj2}\n${obj3}`;
  const o1 = 0;
  const o2 = obj1.length + 1;
  const o3 = obj1.length + 1 + obj2.length + 1;
  const objHeader = `1 ${o1} 2 ${o2} 3 ${o3} `;
  const first = objHeader.length;
  const streamData = objHeader + objData;
  const objStm =
    `4 0 obj\n<</Type/ObjStm/N 3/First ${first}/Length ${streamData.length}>>\n` +
    `stream\n${streamData}\nendstream\nendobj\n`;

  const pdfHeader = '%PDF-1.5\n';
  const offObjStm = pdfHeader.length; // byte offset of `4 0 obj`
  const offXref = offObjStm + objStm.length; // byte offset of `5 0 obj`

  // Six 5-byte rows (/W [1 2 2]): type, 2-byte field2, 2-byte field3.
  const rows = new Uint8Array(30);
  const setRow = (i: number, type: number, f2: number, f3: number): void => {
    const o = i * 5;
    rows[o] = type;
    rows[o + 1] = (f2 >> 8) & 0xff;
    rows[o + 2] = f2 & 0xff;
    rows[o + 3] = (f3 >> 8) & 0xff;
    rows[o + 4] = f3 & 0xff;
  };
  setRow(0, 0, 0, 0xffff); // free-list head
  setRow(1, 2, 4, 0); // obj 1 → object stream 4, index 0
  setRow(2, 2, 4, 1);
  setRow(3, 2, 4, 2);
  setRow(4, 1, offObjStm, 0); // obj 4 → byte offset
  setRow(5, 1, offXref, 0); // obj 5 → byte offset

  const xrefHead = `5 0 obj\n<</Type/XRef/Size 6/Root 1 0 R/W[1 2 2]/Index[0 6]/Length 30>>\nstream\n`;
  const xrefTail = `\nendstream\nendobj\n`;
  const tail = `startxref\n${opts.brokenStartxref ? 999999 : offXref}\n%%EOF`;

  return concat([
    enc.encode(pdfHeader),
    enc.encode(objStm),
    enc.encode(xrefHead),
    rows,
    enc.encode(xrefTail),
    enc.encode(tail),
  ]);
}

function concat(parts: Array<Uint8Array>): Uint8Array {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}
