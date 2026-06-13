// E-PDF EP9 — encrypted-PDF reading. Ream writes AES-256 (V5/R6) encrypted PDFs
// (P2); reading one back with the empty user password must derive the file key
// (Algorithm 2.A/2.B), decrypt the content + font streams, and recover the text.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

describe('encrypted PDF read (E-PDF EP9)', () => {
  it('reads an AES-256 (R6) encrypted PDF with the empty user password', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>TopSecretContents</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', {
      fonts: FONTS,
      encrypt: { userPassword: '' },
    });
    // The PDF really is encrypted (a /Encrypt entry in the trailer).
    expect(PdfFile.parse(pdf).trailer.has('Encrypt')).toBe(true);

    const html = new TextDecoder().decode(await Ream.parse(pdf).convert('html'));
    expect(html.replace(/\s+/g, '')).toContain('TopSecretContents');
  });

  it('round-trips an encrypted tagged PDF', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>EncryptedTagged</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', {
      fonts: FONTS,
      tagged: true,
      encrypt: { userPassword: '' },
    });
    const html = new TextDecoder().decode(await Ream.parse(pdf).convert('html'));
    expect(html.replace(/\s+/g, '')).toContain('EncryptedTagged');
  });
});
