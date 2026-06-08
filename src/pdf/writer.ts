// ISO 32000-1:2008 §7.5 — File Structure.
// PdfDocument collects indirect objects and emits a complete PDF file:
// header (§7.5.2), body of indirect objects (§7.5.3), cross-reference table
// (§7.5.4), trailer (§7.5.5).

import type { PdfValue } from '@/pdf/objects';
import { PdfRef } from '@/pdf/objects';
import { serializeIndirectObject } from '@/pdf/serialize';

const encoder = new TextEncoder();

const BINARY_MARKER = new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]);

interface IndirectObject {
  readonly id: number;
  readonly value: PdfValue;
}

export interface BuildOptions {
  // PDF header version, e.g. "1.7" (default) or "1.4" (PDF/A-1).
  readonly version?: string;
  // When true, emit a /ID array in the trailer (required by PDF/A). The two
  // identifiers are a deterministic hash of the file body — no Date/random,
  // so the same input always yields the same bytes.
  readonly id?: boolean;
}

export class PdfDocument {
  private readonly objects: Array<IndirectObject> = [];

  add(value: PdfValue): PdfRef {
    const id = this.objects.length + 1;
    this.objects.push({ id, value });
    return new PdfRef(id);
  }

  build(root: PdfRef, info?: PdfRef, options: BuildOptions = {}): Uint8Array {
    const parts: Array<Uint8Array> = [];
    const offsets = new Array<number>(this.objects.length + 1).fill(0);
    let pos = 0;

    const push = (bytes: Uint8Array) => {
      parts.push(bytes);
      pos += bytes.length;
    };

    push(encoder.encode(`%PDF-${options.version ?? '1.7'}\n`));
    push(BINARY_MARKER);

    for (const obj of this.objects) {
      offsets[obj.id] = pos;
      push(serializeIndirectObject(obj.id, 0, obj.value));
    }

    const xrefOffset = pos;
    const totalObjects = this.objects.length + 1;
    let xref = `xref\n0 ${totalObjects}\n0000000000 65535 f \n`;
    for (let i = 1; i < totalObjects; i++) {
      const off = offsets[i]!.toString().padStart(10, '0');
      xref += `${off} 00000 n \n`;
    }
    push(encoder.encode(xref));

    const infoEntry = info ? ` /Info ${info.id} ${info.generation} R` : '';
    let idEntry = '';
    if (options.id) {
      // Hash the body emitted so far (everything before the trailer) to derive
      // a stable 16-byte file identifier. PDF/A wants both array elements
      // identical for a freshly-created file.
      const hashHex = hash128Hex(parts);
      idEntry = ` /ID [<${hashHex}> <${hashHex}>]`;
    }
    const trailer =
      `trailer\n<</Size ${totalObjects} /Root ${root.id} ${root.generation} R${infoEntry}${idEntry}>>\n` +
      `startxref\n${xrefOffset}\n%%EOF\n`;
    push(encoder.encode(trailer));

    const out = new Uint8Array(pos);
    let off = 0;
    for (const p of parts) {
      out.set(p, off);
      off += p.length;
    }
    return out;
  }
}

// Deterministic 128-bit identifier built from four independently-seeded
// FNV-1a passes over the body bytes. Not cryptographic — PDF /ID only needs to
// be a stable, well-distributed file fingerprint.
function hash128Hex(parts: ReadonlyArray<Uint8Array>): string {
  const seeds = [0x811c9dc5, 0x01000193, 0x9e3779b9, 0x85ebca6b];
  const acc = new Uint32Array(seeds);
  for (const part of parts) {
    for (let i = 0; i < part.length; i++) {
      const b = part[i]!;
      for (let s = 0; s < 4; s++) {
        acc[s] = (acc[s]! ^ b) >>> 0;
        // FNV-1a multiply by the 32-bit prime, kept in range via Math.imul.
        acc[s] = Math.imul(acc[s]!, 0x01000193) >>> 0;
      }
    }
  }
  let hex = '';
  for (let s = 0; s < 4; s++) hex += acc[s]!.toString(16).padStart(8, '0').toUpperCase();
  return hex;
}
