// E-PDF EP1 — COS lexer + object parser. Parses PDF object syntax into the
// writer's object model (src/pdf/objects.ts), so a parse is the inverse of a
// serialize.

import { describe, expect, it } from 'vitest';

import type { PdfDict, PdfValue } from '@/pdf/objects';
import { PDF_NULL, PdfHexString, PdfName, PdfRef, PdfStream } from '@/pdf/objects';
import { Lexer } from '@/pdf-reader/lexer';
import { parseIndirectObject, parseObject } from '@/pdf-reader/parser';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
const parse = (s: string): PdfValue => parseObject(new Lexer(enc(s)));
const dec = (b: Uint8Array): string => new TextDecoder().decode(b);

describe('COS object parser — primitives (E-PDF EP1)', () => {
  it('parses integers and reals, signed', () => {
    expect(parse('42')).toBe(42);
    expect(parse('3.14')).toBeCloseTo(3.14);
    expect(parse('-.5')).toBeCloseTo(-0.5);
    expect(parse('+2')).toBe(2);
  });

  it('parses names, decoding #XX escapes', () => {
    expect(parse('/Type')).toEqual(new PdfName('Type'));
    expect(parse('/A#20B')).toEqual(new PdfName('A B'));
  });

  it('parses booleans and null', () => {
    expect(parse('true')).toBe(true);
    expect(parse('false')).toBe(false);
    expect(parse('null')).toBe(PDF_NULL);
  });

  it('parses literal strings with escapes and nested parens', () => {
    expect(parse('(hello)')).toBe('hello');
    expect(parse('(a\\)b)')).toBe('a)b'); // escaped close paren
    expect(parse('(nested (parens) ok)')).toBe('nested (parens) ok');
    expect(parse('(\\101)')).toBe('A'); // octal escape \101 = 'A'
    expect(parse('(tab\\there)')).toBe('tab\there');
  });

  it('parses hex strings, padding an odd final digit', () => {
    const h = parse('<48656C6C6F>');
    expect(h).toBeInstanceOf(PdfHexString);
    expect(dec((h as PdfHexString).bytes)).toBe('Hello');
    expect(Array.from((parse('<4>') as PdfHexString).bytes)).toEqual([0x40]);
  });
});

describe('COS object parser — composites (E-PDF EP1)', () => {
  it('parses arrays, including nested', () => {
    expect(parse('[1 2 3]')).toEqual([1, 2, 3]);
    expect(parse('[1 [2 3] /X]')).toEqual([1, [2, 3], new PdfName('X')]);
  });

  it('parses dictionaries', () => {
    const d = parse('<< /Type /Page /Count 3 >>') as PdfDict;
    expect(d).toBeInstanceOf(Map);
    expect(d.get('Type')).toEqual(new PdfName('Page'));
    expect(d.get('Count')).toBe(3);
  });

  it('parses an indirect reference, not a number triple', () => {
    expect(parse('5 0 R')).toEqual(new PdfRef(5, 0));
    expect(parse('[10 0 R 2 0 R]')).toEqual([new PdfRef(10, 0), new PdfRef(2, 0)]);
    // bare numbers must stay numbers
    expect(parse('[1 2 3]')).toEqual([1, 2, 3]);
  });

  it('parses a stream object with a numeric /Length', () => {
    const s = parse('<< /Length 5 >>\nstream\nABCDE\nendstream') as PdfStream;
    expect(s).toBeInstanceOf(PdfStream);
    expect(s.dict.get('Length')).toBe(5);
    expect(dec(s.data)).toBe('ABCDE');
  });

  it('falls back to scanning for endstream when /Length is absent', () => {
    const s = parse('<< /Filter /FlateDecode >>\nstream\nrawbytes\nendstream') as PdfStream;
    expect(s).toBeInstanceOf(PdfStream);
    expect(dec(s.data)).toBe('rawbytes');
  });

  it('parses an indirect object definition (N G obj … endobj)', () => {
    const obj = parseIndirectObject(new Lexer(enc('12 0 obj\n<< /Type /Catalog >>\nendobj')));
    expect(obj?.id).toBe(12);
    expect(obj?.generation).toBe(0);
    expect((obj?.value as PdfDict).get('Type')).toEqual(new PdfName('Catalog'));
  });
});
