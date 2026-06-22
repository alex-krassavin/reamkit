// E-PDF EP1 — COS object parser (ISO 32000-1 §7.3). Drives the lexer to recover
// the PDF object grammar into the writer's object model (src/pdf/objects.ts):
// numbers, names, strings, arrays, dictionaries, streams, the null/boolean
// keywords, indirect references (`N G R`) and indirect object definitions
// (`N G obj … endobj`).

import type { Lexer, Token } from './lexer';
import type { PdfArray, PdfDict, PdfValue } from '@/pdf/objects';

import { PDF_NULL, PdfHexString, PdfName, PdfRef, PdfStream } from '@/pdf/objects';

/**
 * Resolves an indirect-reference `/Length` to its numeric value. When a stream's
 * `/Length` is an indirect reference it cannot be resolved while parsing the
 * object in isolation; the document layer passes a resolver. Without one the
 * parser falls back to scanning for `endstream`.
 */
export type LengthResolver = (ref: PdfRef) => number | undefined;

/** A parsed `N G obj … endobj` definition: its id, generation and contained value. */
export interface IndirectObject {
  readonly id: number;
  readonly generation: number;
  readonly value: PdfValue;
}

/** Parse one object value at the lexer's current position. */
export function parseObject(lexer: Lexer, resolveLength?: LengthResolver): PdfValue {
  return parseValue(lexer, lexer.nextToken(), resolveLength);
}

/**
 * Parse an `N G obj … endobj` definition (the lexer must sit at the leading
 * integer).
 *
 * @returns The {@link IndirectObject} (id, generation and contained value), or
 *          `undefined` when the `N G obj` header does not parse.
 */
export function parseIndirectObject(
  lexer: Lexer,
  resolveLength?: LengthResolver,
): IndirectObject | undefined {
  const idTok = lexer.nextToken();
  if (idTok.kind !== 'num') return undefined;
  const genTok = lexer.nextToken();
  if (genTok.kind !== 'num') return undefined;
  const objTok = lexer.nextToken();
  if (objTok.kind !== 'keyword' || objTok.value !== 'obj') return undefined;
  const value = parseObject(lexer, resolveLength);
  // Consume the trailing `endobj` if present (tolerated if missing).
  const save = lexer.pos;
  const end = lexer.nextToken();
  if (!(end.kind === 'keyword' && end.value === 'endobj')) lexer.pos = save;
  return { id: idTok.value, generation: genTok.value, value };
}

function parseValue(lexer: Lexer, tok: Token, resolveLength?: LengthResolver): PdfValue {
  switch (tok.kind) {
    case 'num':
      return parseNumberOrRef(lexer, tok.value);
    case 'name':
      return new PdfName(tok.value);
    case 'str':
      return tok.value;
    case 'hexstr':
      return new PdfHexString(tok.bytes);
    case 'arrayOpen':
      return parseArray(lexer, resolveLength);
    case 'dictOpen':
      return parseDictOrStream(lexer, resolveLength);
    case 'keyword':
      if (tok.value === 'true') return true;
      if (tok.value === 'false') return false;
      // `null`, `endobj`, stray operators, EOF — all surface as null.
      return PDF_NULL;
    default:
      return PDF_NULL;
  }
}

// A bare integer may begin an indirect reference `N G R`. Look ahead two tokens;
// if they are not `<int> R`, rewind to just after the first number.
function parseNumberOrRef(lexer: Lexer, first: number): PdfValue {
  if (!Number.isInteger(first) || first < 0) return first;
  const save = lexer.pos;
  const gen = lexer.nextToken();
  if (gen.kind === 'num' && Number.isInteger(gen.value)) {
    const r = lexer.nextToken();
    if (r.kind === 'keyword' && r.value === 'R') {
      return new PdfRef(first, gen.value);
    }
  }
  lexer.pos = save;
  return first;
}

function parseArray(lexer: Lexer, resolveLength?: LengthResolver): PdfArray {
  const out: PdfArray = [];
  for (;;) {
    const tok = lexer.nextToken();
    if (tok.kind === 'arrayClose' || tok.kind === 'eof') break;
    out.push(parseValue(lexer, tok, resolveLength));
  }
  return out;
}

function parseDictOrStream(lexer: Lexer, resolveLength?: LengthResolver): PdfValue {
  const map: PdfDict = new Map<string, PdfValue>();
  for (;;) {
    const keyTok = lexer.nextToken();
    if (keyTok.kind === 'dictClose' || keyTok.kind === 'eof') break;
    if (keyTok.kind !== 'name') {
      // Malformed entry (key must be a name) — skip a value to stay in step.
      if (keyTok.kind === 'arrayOpen' || keyTok.kind === 'dictOpen') {
        parseValue(lexer, keyTok, resolveLength);
      }
      continue;
    }
    const value = parseObject(lexer, resolveLength);
    map.set(keyTok.value, value);
  }

  // A dictionary directly followed by the `stream` keyword is a stream object.
  const save = lexer.pos;
  const next = lexer.nextToken();
  if (next.kind === 'keyword' && next.value === 'stream') {
    const lengthVal = map.get('Length');
    let length: number | undefined;
    if (typeof lengthVal === 'number') length = lengthVal;
    else if (lengthVal instanceof PdfRef && resolveLength) length = resolveLength(lengthVal);
    const data = lexer.readStreamBody(length);
    const endTok = lexer.nextToken(); // expect `endstream`
    if (!(endTok.kind === 'keyword' && endTok.value === 'endstream')) {
      // tolerate a missing/teleported endstream — readStreamBody already
      // positioned us; do not rewind.
    }
    return new PdfStream(map, data);
  }
  lexer.pos = save;
  return map;
}
