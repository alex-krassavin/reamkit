// ISO 32000-1:2008 §7.3 — Serialization of PDF objects to bytes.

import type { PdfDict, PdfValue } from '@/pdf/objects';
import { PDF_NULL, PdfHexString, PdfName, PdfRawToken, PdfRef, PdfStream } from '@/pdf/objects';

const encoder = new TextEncoder();

const NAME_ESCAPE = /[^!-~]|[#%()/<>[\]{}]/g;
const escapeName = (s: string): string =>
  s.replace(NAME_ESCAPE, (c) => `#${c.charCodeAt(0).toString(16).padStart(2, '0').toUpperCase()}`);

const escapeLiteralString = (s: string): string =>
  s.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');

const formatNumber = (n: number): string => {
  if (Number.isInteger(n)) return String(n);
  return n.toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
};

/**
 * Serialize a single PDF value to its textual representation (ISO 32000-1 §7.3).
 *
 * @param v The value to serialize.
 * @returns The serialized string.
 * @throws Error on a non-finite number, or a {@link PdfStream} (which must be an
 *   indirect object, not inlined).
 */
export function serializeValue(v: PdfValue): string {
  if (v === PDF_NULL) return 'null';
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) throw new Error(`Non-finite number in PDF: ${v}`);
    return formatNumber(v);
  }
  if (typeof v === 'string') return `(${escapeLiteralString(v)})`;
  if (v instanceof PdfHexString) {
    let hex = '';
    for (const b of v.bytes) hex += b.toString(16).padStart(2, '0').toUpperCase();
    return `<${hex}>`;
  }
  if (v instanceof PdfRawToken) return v.token;
  if (v instanceof PdfName) return `/${escapeName(v.value)}`;
  if (Array.isArray(v)) return `[${v.map(serializeValue).join(' ')}]`;
  if (v instanceof PdfRef) return `${v.id} ${v.generation} R`;
  if (v instanceof Map) return serializeDict(v);
  if (v instanceof PdfStream) {
    throw new Error('PdfStream must be emitted as an indirect object, not inlined');
  }
  throw new Error(`Unknown PDF value: ${String(v)}`);
}

function serializeDict(d: PdfDict): string {
  const entries: Array<string> = [];
  for (const [k, v] of d) {
    entries.push(`/${escapeName(k)} ${serializeValue(v)}`);
  }
  return `<<${entries.join(' ')}>>`;
}

/**
 * Serialize a value as a complete `id gen obj … endobj` indirect object
 * (§7.3.10). Streams are emitted with their raw payload; everything else is
 * inlined via {@link serializeValue}.
 *
 * @param id         The object number.
 * @param generation The generation number.
 * @param value      The object value.
 * @returns The encoded indirect-object bytes.
 */
export function serializeIndirectObject(
  id: number,
  generation: number,
  value: PdfValue,
): Uint8Array {
  if (value instanceof PdfStream) {
    return serializeStreamObject(id, generation, value);
  }
  const body = `${id} ${generation} obj\n${serializeValue(value)}\nendobj\n`;
  return encoder.encode(body);
}

function serializeStreamObject(id: number, generation: number, s: PdfStream): Uint8Array {
  const length = s.data.byteLength;
  if (!s.dict.has('Length')) {
    s.dict.set('Length', length);
  }
  const header = encoder.encode(`${id} ${generation} obj\n${serializeDict(s.dict)}\nstream\n`);
  const footer = encoder.encode('\nendstream\nendobj\n');
  const out = new Uint8Array(header.length + s.data.length + footer.length);
  out.set(header, 0);
  out.set(s.data, header.length);
  out.set(footer, header.length + s.data.length);
  return out;
}
