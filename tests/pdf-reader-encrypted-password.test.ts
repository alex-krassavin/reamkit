// E-PDF EP14 — encrypted-PDF reading with a NON-EMPTY user password. EP9 covered
// the permissions-only (empty-password) case; here Ream writes an AES-256 (V5/R6)
// PDF locked with a real user password, and the reader must validate that exact
// password (Algorithm 2.A), recover the file key, and decrypt the content — while
// a wrong or missing password fails validation and is reported as a loss.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const PASSWORD = 'Corre©t-Horse-9';

async function lockedPdf(text: string): Promise<Uint8Array> {
  const docx = buildDocxFromBody(`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`);
  return Ream.parse(docx).convert('pdf', {
    fonts: FONTS,
    encrypt: { userPassword: PASSWORD },
  });
}

async function html(pdf: Uint8Array, password?: string): Promise<string> {
  const doc = password === undefined ? Ream.parse(pdf) : Ream.parse(pdf, { password });
  return new TextDecoder().decode(await doc.convert('html'));
}

function needsPassword(losses: Ream['losses']): boolean {
  return losses.some((l) => l.detail.includes('user password'));
}

describe('encrypted PDF read with a user password (E-PDF EP14)', () => {
  it('recovers content when the correct user password is supplied', async () => {
    const pdf = await lockedPdf('UnlockedWithPassword');

    const parsed = Ream.parse(pdf, { password: PASSWORD });
    expect(needsPassword(parsed.losses)).toBe(false);

    const out = new TextDecoder().decode(await parsed.convert('html'));
    expect(out.replace(/\s+/g, '')).toContain('UnlockedWithPassword');
  });

  it('reports a loss and recovers nothing when the password is wrong', async () => {
    const pdf = await lockedPdf('UnlockedWithPassword');

    const parsed = Ream.parse(pdf, { password: 'not-the-password' });
    expect(needsPassword(parsed.losses)).toBe(true);

    const out = new TextDecoder().decode(await parsed.convert('html'));
    expect(out.replace(/\s+/g, '')).not.toContain('UnlockedWithPassword');
  });

  it('reports a loss when no password is supplied for a locked document', async () => {
    const pdf = await lockedPdf('UnlockedWithPassword');
    expect(needsPassword(Ream.parse(pdf).losses)).toBe(true);
  });

  it('still opens the permissions-only (empty-password) case by default', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>PermissionsOnly</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', {
      fonts: FONTS,
      encrypt: { userPassword: '' },
    });

    const parsed = Ream.parse(pdf);
    expect(needsPassword(parsed.losses)).toBe(false);
    expect((await html(pdf)).replace(/\s+/g, '')).toContain('PermissionsOnly');
  });
});
