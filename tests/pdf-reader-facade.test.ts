// E-PDF EP5 — pdfReader wired into the facade. Ream.parse now sniffs %PDF- and
// reconstructs a FlowDoc, so a PDF is a first-class input: parse it, inspect the
// interlayer, and convert it onward to HTML or docx.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { pdfReader } from '@/pdf-reader/reader';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

// Produce a PDF from a one-paragraph docx (tagged so structure survives).
async function pdfOf(text: string): Promise<Uint8Array> {
  return Ream.parse(buildDocxFromBody(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)).convert('pdf', {
    fonts: FONTS,
    tagged: true,
  });
}

const bodyText = (flow: { body: ReadonlyArray<{ kind: string }> }): string =>
  flow.body
    .filter(
      (b): b is { kind: 'paragraph'; paragraph: { runs: ReadonlyArray<{ text: string }> } } =>
        b.kind === 'paragraph',
    )
    .map((p) => p.paragraph.runs.map((r) => r.text).join(''))
    .join('')
    .replace(/\s+/g, '');

describe('pdfReader facade wiring (E-PDF EP5)', () => {
  it('sniffs %PDF- and rejects other formats', () => {
    expect(pdfReader.sniff(enc('%PDF-1.7\n1 0 obj'))).toBe(true);
    expect(pdfReader.sniff(enc('   \n%PDF-1.4'))).toBe(true); // tolerates leading junk
    expect(pdfReader.sniff(enc('PK'))).toBe(false); // a ZIP (docx/xlsx)
    expect(pdfReader.sniff(enc('plain text'))).toBe(false);
  });

  it('Ream.parse sniffs a PDF and reports format "pdf"', async () => {
    const doc = Ream.parse(await pdfOf('InterlayerText'));
    expect(doc.format).toBe('pdf');
    expect(bodyText(doc.flow)).toContain('InterlayerText');
    expect(doc.losses.some((l) => l.feature === 'images')).toBe(true);
  });

  it('converts a parsed PDF to HTML carrying its text', async () => {
    const html = new TextDecoder().decode(
      await Ream.parse(await pdfOf('ExportedToHtml')).convert('html'),
    );
    expect(html.replace(/\s+/g, '')).toContain('ExportedToHtml');
  });

  it('converts a parsed PDF to a valid docx that re-parses to the same text', async () => {
    const out = await Ream.parse(await pdfOf('BackToDocx')).convert('docx');
    expect(out[0]).toBe(0x50); // 'P'
    expect(out[1]).toBe(0x4b); // 'K' — a ZIP/OOXML package
    expect(bodyText(Ream.parse(out).flow)).toContain('BackToDocx');
  });
});
