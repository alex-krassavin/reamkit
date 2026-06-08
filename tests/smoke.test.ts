import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocx } from './fixtures/build-docx';
import { convertDocxToPdfSync } from '@/converter';
import { parseTtf } from '@/font';

const here = dirname(fileURLToPath(import.meta.url));
const ROBOTO = new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf')));

const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

describe('M1 smoke: docx → pdf with embedded TTF', () => {
  it('emits a Type0 + CIDFontType2 font structure', () => {
    const docx = buildDocx(['Hello, world']);
    const pdf = convertDocxToPdfSync(docx, { fontBytes: ROBOTO });

    const text = asLatin1(pdf);
    expect(text.startsWith('%PDF-1.7\n')).toBe(true);
    expect(text.trimEnd().endsWith('%%EOF')).toBe(true);
    expect(text).toContain('/Type /Catalog');
    expect(text).toContain('/Subtype /Type0');
    expect(text).toContain('/Encoding /Identity-H');
    expect(text).toContain('/Subtype /CIDFontType2');
    expect(text).toContain('/CIDToGIDMap /Identity');
    expect(text).toContain('/Type /FontDescriptor');
    expect(text).toContain('/FontFile2');
    expect(text).toContain('/ToUnicode');
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto/);
  });

  it('encodes Cyrillic text as Identity-H glyph IDs (no .notdef)', () => {
    const docx = buildDocx(['Привет, мир!']);
    const pdf = convertDocxToPdfSync(docx, { fontBytes: ROBOTO });
    const text = asLatin1(pdf);

    const showMatches = [...text.matchAll(/<([0-9A-Fa-f]+)> Tj/g)];
    expect(showMatches.length).toBeGreaterThan(0);
    const hex = showMatches.map((m) => m[1]!).join('');

    expect(hex.length % 4).toBe(0);
    const gids: Array<number> = [];
    for (let i = 0; i < hex.length; i += 4) {
      gids.push(parseInt(hex.slice(i, i + 4), 16));
    }
    expect(gids.every((g) => g !== 0)).toBe(true);

    const parsed = parseTtf(ROBOTO);
    const expected = [...'Привет, мир!'].map((ch) => parsed.glyphForCodepoint(ch.codePointAt(0)!));
    expect(gids).toEqual(expected);
  });

  it('xref table has correct entry count and offsets land on objects', () => {
    const docx = buildDocx(['One']);
    const pdf = convertDocxToPdfSync(docx, { fontBytes: ROBOTO });
    const text = asLatin1(pdf);

    const sxMatch = text.match(/startxref\n(\d+)\n%%EOF/);
    expect(sxMatch).not.toBeNull();
    const xrefStart = Number(sxMatch![1]);
    const xrefSection = text.slice(xrefStart);
    const headerMatch = xrefSection.match(/^xref\n0 (\d+)\n/);
    expect(headerMatch).not.toBeNull();
    const numEntries = Number(headerMatch![1]);

    const entryRegex = /(\d{10}) (\d{5}) ([fn]) /g;
    const entries: Array<{ offset: number; gen: number; type: string }> = [];
    let m: RegExpExecArray | null;
    while ((m = entryRegex.exec(xrefSection)) !== null) {
      entries.push({ offset: Number(m[1]), gen: Number(m[2]), type: m[3]! });
    }
    expect(entries.length).toBe(numEntries);
    expect(entries[0]).toEqual({ offset: 0, gen: 65535, type: 'f' });

    for (let i = 1; i < entries.length; i++) {
      const off = entries[i]!.offset;
      const objectHeader = text.slice(off, off + 30);
      expect(objectHeader).toMatch(new RegExp(`^${i} 0 obj`));
    }
  });

  it('paginates multi-paragraph docx into multiple Tj operators', () => {
    const docx = buildDocx(['First.', 'Second.', 'Third.']);
    const pdf = convertDocxToPdfSync(docx, { fontBytes: ROBOTO });
    const text = asLatin1(pdf);

    const tjMatches = text.match(/<[0-9A-Fa-f]+> Tj/g);
    expect(tjMatches).not.toBeNull();
    expect(tjMatches!.length).toBe(3);
  });

  it('does not output WinAnsi literal-string Tj operators', () => {
    const docx = buildDocx(['Hello']);
    const pdf = convertDocxToPdfSync(docx, { fontBytes: ROBOTO });
    const text = asLatin1(pdf);

    expect(text).not.toMatch(/\([^)]*\) Tj/);
  });

  it('subsets the embedded font to well below the full TTF size', () => {
    const docx = buildDocx(['Hi']);
    const pdf = convertDocxToPdfSync(docx, { fontBytes: ROBOTO });
    expect(pdf.byteLength).toBeLessThan(ROBOTO.byteLength / 2);
  });
});
