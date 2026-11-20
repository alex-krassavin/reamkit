import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { iriToUri } from '@/core/links';
import { PdfName } from '@/pdf/objects';
import { PdfFile } from '@/pdf-reader/document';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const pdfText = (bytes: Uint8Array) => Buffer.from(bytes).toString('latin1');

describe('PDF link annotations (ISO 32000 §12.5.6.5)', () => {
  it('emits a /Link annotation with a /URI action for an external hyperlink', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>before </w:t></w:r>' +
        '<w:hyperlink r:id="rId30"><w:r><w:t>visit reamkit</w:t></w:r></w:hyperlink>' +
        '<w:r><w:t> after</w:t></w:r></w:p>',
      { hyperlinks: { rId30: 'https://reamkit.dev' } },
    );
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS }));
    expect(pdf).toContain('/Subtype /Link');
    expect(pdf).toContain('/URI (https://reamkit.dev)');
    expect(pdf).toContain('/Annots');
    expect(pdf.match(/\/Subtype \/Link/g)!.length).toBe(1);
  });

  it('one rect per line: a wrapping link produces an annotation per line', async () => {
    const words = Array.from({ length: 12 }, (_, i) => `linkword${i}`).join(' ');
    const docx = buildDocxFromBody(
      `<w:p><w:hyperlink r:id="rId30"><w:r><w:t>${words}</w:t></w:r></w:hyperlink></w:p>`,
      { hyperlinks: { rId30: 'https://reamkit.dev/long' } },
    );
    // A narrow page forces the link text across several lines.
    const pdf = pdfText(
      await Ream.parse(docx).convert('pdf', { fonts: FONTS, pageWidth: 220, pageHeight: 400 }),
    );
    expect(pdf.match(/\/Subtype \/Link/g)!.length).toBeGreaterThan(1);
  });

  it('disallowed schemes produce no annotation (text still renders)', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:hyperlink r:id="rId31"><w:r><w:t>not clickable</w:t></w:r></w:hyperlink></w:p>',
      { hyperlinks: { rId31: 'javascript:alert(1)' } },
    );
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS }));
    expect(pdf).not.toContain('/Subtype /Link');
    expect(pdf).not.toContain('javascript:');
  });

  it('tagged mode encloses the annotation in a Link StructElem (OBJR + StructParent)', async () => {
    const docx = buildDocxFromBody(
      '<w:p><w:hyperlink r:id="rId30"><w:r><w:t>tagged link</w:t></w:r></w:hyperlink></w:p>',
      { hyperlinks: { rId30: 'https://reamkit.dev' } },
    );
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS, tagged: true }));
    expect(pdf).toContain('/S /Link');
    expect(pdf).toContain('/Type /OBJR');
    expect(pdf).toContain('/StructParent ');
    expect(pdf).toContain('/F 4');
  });

  it('documents without links carry no /Annots at all', async () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>plain</w:t></w:r></w:p>');
    const pdf = pdfText(await Ream.parse(docx).convert('pdf', { fonts: FONTS }));
    expect(pdf).not.toContain('/Annots');
  });
});

// ISO 32000-1 §12.6.4.7: a URI action's /URI is 7-bit ASCII. A viewer takes
// the string's bytes one character each, so the IRI a document links to goes
// out as the URI it stands for (RFC 3987 §3.1) — the one a browser follows
// from the HTML writer's href.
describe('PDF /URI is the 7-bit ASCII URI of the hyperlink target', () => {
  const targets = {
    rId30: 'https://kk.wikipedia.org/wiki/Қазақ_тілі',
    rId31: 'http://пример.рф/путь?q=тест#якорь',
    rId32: 'mailto:әлия@мысал.қаз?subject=Сәлем',
    rId33: 'https://reamkit.dev/a%20b?x=(1)',
  };
  const docx = () =>
    buildDocxFromBody(
      Object.keys(targets)
        .map(
          (id, i) =>
            `<w:p><w:hyperlink r:id="${id}"><w:r><w:t>link ${i}</w:t></w:r></w:hyperlink></w:p>`,
        )
        .join(''),
      { hyperlinks: targets },
    );

  // Every page's Link annotations, in order, with their URI action's /URI as
  // the bytes it holds (a literal string reads back as latin1).
  const readUris = (bytes: Uint8Array): Array<unknown> => {
    const file = PdfFile.parse(bytes);
    const out: Array<unknown> = [];
    for (const page of file.pages()) {
      const annots = file.get(page.dict, 'Annots');
      for (const a of Array.isArray(annots) ? annots : []) {
        const annot = file.resolve(a);
        if (!(annot instanceof Map)) continue;
        const action = file.get(annot, 'A');
        if (!(action instanceof Map)) continue;
        const kind = file.get(action, 'S');
        if (kind instanceof PdfName && kind.value === 'URI') out.push(file.get(action, 'URI'));
      }
    }
    return out;
  };

  it('percent-encodes the UTF-8 of path, query and fragment and IDNA-encodes the host', async () => {
    const uris = readUris(await Ream.parse(docx()).convert('pdf', { fonts: FONTS }));
    expect(uris).toEqual([
      'https://kk.wikipedia.org/wiki/%D2%9A%D0%B0%D0%B7%D0%B0%D2%9B_%D1%82%D1%96%D0%BB%D1%96',
      'http://xn--e1afmkfd.xn--p1ai/%D0%BF%D1%83%D1%82%D1%8C?q=%D1%82%D0%B5%D1%81%D1%82#%D1%8F%D0%BA%D0%BE%D1%80%D1%8C',
      'mailto:%D3%99%D0%BB%D0%B8%D1%8F@%D0%BC%D1%8B%D1%81%D0%B0%D0%BB.%D2%9B%D0%B0%D0%B7?subject=%D0%A1%D3%99%D0%BB%D0%B5%D0%BC',
      // An ASCII target goes out byte for byte, its %XX and parentheses alike.
      'https://reamkit.dev/a%20b?x=(1)',
    ]);
  });

  it('lands where a browser lands from the HTML writer’s href, which keeps the IRI', async () => {
    const html = new TextDecoder().decode(
      await Ream.parse(docx()).convert('html', { fonts: FONTS }),
    );
    const hrefs = [...html.matchAll(/<a href="([^"]*)"/g)].map((m) => m[1]);
    expect(hrefs).toEqual(Object.values(targets));
    const uris = readUris(await Ream.parse(docx()).convert('pdf', { fonts: FONTS }));
    expect(uris).toEqual(hrefs.map((h) => new URL(h!).href));
  });
});

describe('iriToUri (RFC 3987 §3.1)', () => {
  it('returns an ASCII target unchanged, whatever it holds', () => {
    for (const ascii of [
      'https://reamkit.dev',
      'https://x.dev/a b?q=%zz#f',
      'mailto:a@b.c?subject=hi',
    ]) {
      expect(iriToUri(ascii)).toBe(ascii);
    }
  });

  it('is idempotent: a mapped URI maps to itself', () => {
    for (const iri of ['https://例え.テスト/パス?ä=ö#ü', 'mailto:ә@ә.kz']) {
      const uri = iriToUri(iri);
      expect(uri).toMatch(/^[\x20-\x7e]*$/);
      expect(iriToUri(uri)).toBe(uri);
    }
  });

  it('encodes userinfo, keeps the port, and sees "@" and ":" in the right places', () => {
    expect(iriToUri('http://юзер@пример.рф:8080/ä')).toBe(
      'http://%D1%8E%D0%B7%D0%B5%D1%80@xn--e1afmkfd.xn--p1ai:8080/%C3%A4',
    );
    // An IP literal is ASCII by grammar; only the path changes.
    expect(iriToUri('http://[::1]:80/ä')).toBe('http://[::1]:80/%C3%A4');
  });

  it('percent-encodes a host that is no valid domain name instead of dropping the link', () => {
    expect(iriToUri('http://пр имер.рф/')).toBe(
      'http://%D0%BF%D1%80 %D0%B8%D0%BC%D0%B5%D1%80.%D1%80%D1%84/',
    );
    // "\" splits a host for the URL parser; the host stays whole here.
    expect(iriToUri('http://a\\b.рф/')).toBe('http://a\\b.%D1%80%D1%84/');
  });

  it('encodes astral characters as four octets and a lone surrogate as U+FFFD', () => {
    expect(iriToUri('https://x.dev/\u{1F600}?\uD800')).toBe('https://x.dev/%F0%9F%98%80?%EF%BF%BD');
  });
});
