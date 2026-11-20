// User text in annotations and the signature dictionary is a PDF text string
// (ISO 32000-1 §7.9.2.2): a literal is read as PDFDocEncoding, so anything past
// printable ASCII has to go out as UTF-16BE behind FE FF. Written as bare JS
// strings these went out as UTF-8 bytes, and macOS PDFKit showed the Kazakh link
// "Қазақ тілі 中文" as "ÒıÐ°Ð·Ð°Òł Ñ‡ÑŒÐ»ÑŒ ä¸ æŒ⁄" (probe, 2026-11-20). Every
// document here is synthetic.

import { webcrypto } from 'node:crypto';
import { readFileSync } from 'node:fs';

import { zipSync } from 'fflate';

import { describe, expect, it } from 'vitest';

import type { BodyElement } from '@/core/document-model';
import type { PdfDict, PdfName, PdfRef, PdfValue } from '@/pdf/objects';
import { Ream } from '@/core/converter/ream';
import { FontRegistry } from '@/core/font';
import { EMPTY_STYLE_SHEET } from '@/core/style-cascade';
import { renderStyledPdf, signPdf } from '@/pdf';
import { PdfHexString } from '@/pdf/objects';
import { PdfFile } from '@/pdf-reader/document';
import { pdfTextString } from '@/pdf-reader/encodings';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const enc = new TextEncoder();
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

const esc = (t: string): string =>
  t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');

/**
 * A DOCX with one external hyperlink per entry of `links` (each its own
 * paragraph) and, when `comment` is given, one review comment on a word.
 */
function docx(
  links: ReadonlyArray<string>,
  comment?: { readonly author: string; readonly paragraphs: ReadonlyArray<string> },
): Uint8Array {
  const body =
    links
      .map(
        (text, i) =>
          `<w:p><w:hyperlink r:id="rIdL${String(i)}">` +
          `<w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:hyperlink></w:p>`,
      )
      .join('') +
    (comment
      ? `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Reviewed</w:t></w:r>` +
        `<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`
      : '');
  const rels =
    links
      .map(
        (_, i) =>
          `<Relationship Id="rIdL${String(i)}" Type="${R_NS}/hyperlink" ` +
          `Target="https://reamkit.dev/${String(i)}" TargetMode="External"/>`,
      )
      .join('') +
    (comment ? `<Relationship Id="rIdC" Type="${R_NS}/comments" Target="comments.xml"/>` : '');
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT}">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        (comment
          ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
          : '') +
        `</Types>`,
    ),
    '_rels/.rels': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    'word/document.xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`,
    ),
    'word/_rels/document.xml.rels': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">${rels}</Relationships>`,
    ),
  };
  if (comment) {
    files['word/comments.xml'] = enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="${W_NS}">` +
        `<w:comment w:id="0" w:author="${esc(comment.author)}" w:date="2026-11-20T10:00:00Z">` +
        comment.paragraphs.map((p) => `<w:p><w:r><w:t>${esc(p)}</w:t></w:r></w:p>`).join('') +
        `</w:comment></w:comments>`,
    );
  }
  return zipSync(files);
}

/** A string value as a reader decodes it: the bytes as written, then §7.9.2.2. */
function decoded(raw: PdfValue | undefined): string {
  if (raw instanceof PdfHexString) return pdfTextString(raw.bytes);
  if (typeof raw === 'string') {
    // The reader hands a literal back as one char per byte.
    return pdfTextString(Uint8Array.from(raw, (c) => c.charCodeAt(0) & 0xff));
  }
  throw new Error(`not a string: ${String(raw)}`);
}

/** Every annotation on the first page, resolved. */
function annotations(bytes: Uint8Array): Array<PdfDict> {
  const file = PdfFile.parse(bytes);
  const annots = file.get(file.pages()[0]!.dict, 'Annots');
  if (!Array.isArray(annots)) return [];
  return annots.map((a) => file.resolve(a)).filter((a): a is PdfDict => a instanceof Map);
}

const subtype = (a: PdfDict): string | undefined =>
  (a.get('Subtype') as PdfName | undefined)?.value;

/** UTF-16BE with its byte-order mark — the non-literal form of a text string. */
function expectUtf16(raw: PdfValue | undefined): void {
  expect(raw).toBeInstanceOf(PdfHexString);
  const bytes = (raw as PdfHexString).bytes;
  expect([bytes[0], bytes[1]]).toEqual([0xfe, 0xff]);
}

describe('link annotation /Contents as a text string', () => {
  it('carries Kazakh, CJK and astral link text exactly', async () => {
    const texts = ['Қазақ тілі — ә ғ қ ң ө ұ ү һ і', '中文链接 日本語のリンク', 'Emoji 😀 and 𠜎'];
    const pdf = await Ream.parse(docx(texts)).convert('pdf', { fonts: FONTS });
    const links = annotations(pdf).filter((a) => subtype(a) === 'Link');
    expect(links.map((a) => decoded(a.get('Contents')))).toEqual(texts);
    for (const a of links) expectUtf16(a.get('Contents'));
  });

  it('keeps printable ASCII a literal, byte for byte as before', async () => {
    const texts = ['visit reamkit', 'Plain (with parens) \\ and a backslash'];
    const pdf = await Ream.parse(docx(texts)).convert('pdf', { fonts: FONTS });
    const links = annotations(pdf).filter((a) => subtype(a) === 'Link');
    expect(links.map((a) => decoded(a.get('Contents')))).toEqual(texts);
    for (const a of links) expect(typeof a.get('Contents')).toBe('string');
    const latin1 = Buffer.from(pdf).toString('latin1');
    expect(latin1).toContain('/Contents (visit reamkit)');
    expect(latin1).toContain('/Contents (Plain \\(with parens\\) \\\\ and a backslash)');
  });
});

describe('comment sticky note /Contents and /T as text strings', () => {
  it('carries a Cyrillic author and body exactly', async () => {
    const author = 'Әлия Сәрсенова';
    const body = 'Уточните это предложение: «қате» — ошибка.';
    const pdf = await Ream.parse(docx([], { author, paragraphs: [body] })).convert('pdf', {
      fonts: FONTS,
      commentAnnotations: true,
    });
    const notes = annotations(pdf).filter((a) => subtype(a) === 'Text');
    expect(notes).toHaveLength(1);
    const note = notes[0]!;
    expect(decoded(note.get('T'))).toBe(author);
    expect(decoded(note.get('Contents'))).toBe(body);
    expectUtf16(note.get('T'));
    expectUtf16(note.get('Contents'));
  });

  it('keeps an ASCII author and one-line body literals', async () => {
    const pdf = await Ream.parse(
      docx([], { author: 'Alice Reviewer', paragraphs: ['Please clarify.'] }),
    ).convert('pdf', { fonts: FONTS, commentAnnotations: true });
    const note = annotations(pdf).find((a) => subtype(a) === 'Text')!;
    expect(note.get('T')).toBe('Alice Reviewer');
    expect(note.get('Contents')).toBe('Please clarify.');
  });

  it('keeps the line break of a body of two paragraphs', async () => {
    // The body joins its paragraphs with LF, which is not printable ASCII: the
    // string goes out as UTF-16BE and the break survives as U+000A.
    const pdf = await Ream.parse(
      docx([], { author: 'Alice Reviewer', paragraphs: ['First.', 'Second.'] }),
    ).convert('pdf', { fonts: FONTS, commentAnnotations: true });
    const note = annotations(pdf).find((a) => subtype(a) === 'Text')!;
    expect(decoded(note.get('Contents'))).toBe('First.\nSecond.');
  });
});

describe('signature dictionary text strings (ISO 32000-1 Table 252)', () => {
  const registry = FontRegistry.fromBytes({ regular: FONTS.regular });
  const BODY: Array<BodyElement> = [
    {
      kind: 'paragraph',
      paragraph: { properties: {}, runs: [{ text: 'Sign me.', properties: {} }] },
    },
  ];
  const placeholder = {
    reason: 'Келісемін',
    location: 'Алматы',
    name: 'Әлия Сәрсенова',
    contactInfo: '電話 +7 727',
    fieldName: 'Қолтаңба',
  };
  const read = (bytes: Uint8Array) => {
    const file = PdfFile.parse(bytes);
    const form = file.get(file.catalog, 'AcroForm') as PdfDict;
    const field = file.resolve((form.get('Fields') as Array<PdfValue>)[0]!) as PdfDict;
    const sig = file.resolve(field.get('V') as PdfRef) as PdfDict;
    return { field, sig };
  };

  it('carries the signer, reason, place, contact and field name exactly', () => {
    const pdf = renderStyledPdf(BODY, {
      registry,
      styles: EMPTY_STYLE_SHEET,
      signaturePlaceholder: placeholder,
    });
    const { field, sig } = read(pdf);
    expect(decoded(field.get('T'))).toBe(placeholder.fieldName);
    expect(decoded(sig.get('Name'))).toBe(placeholder.name);
    expect(decoded(sig.get('Reason'))).toBe(placeholder.reason);
    expect(decoded(sig.get('Location'))).toBe(placeholder.location);
    expect(decoded(sig.get('ContactInfo'))).toBe(placeholder.contactInfo);
    expectUtf16(field.get('T'));
    expectUtf16(sig.get('Name'));
    expectUtf16(sig.get('Reason'));
  });

  it('still signs: the scan meets the /Contents hole before the hex text strings', async () => {
    const pdf = renderStyledPdf(BODY, {
      registry,
      styles: EMPTY_STYLE_SHEET,
      signaturePlaceholder: placeholder,
    });
    const key = await webcrypto.subtle.importKey(
      'pkcs8',
      new Uint8Array(readFileSync('tests/fixtures/sign/signer-key.pkcs8')),
      { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
      false,
      ['sign'],
    );
    const signed = await signPdf(pdf, {
      certificate: new Uint8Array(readFileSync('tests/fixtures/sign/signer-cert.der')),
      privateKey: key,
    });
    const { sig } = read(signed);
    const [a, b, c, d] = sig.get('ByteRange') as Array<number>;
    expect(a).toBe(0);
    expect(c! + d!).toBe(signed.length);
    // The hole the range leaves out is exactly the signature dictionary's /Contents.
    const hole = Buffer.from(signed.subarray(b, c)).toString('latin1');
    expect(hole).toMatch(/^[0-9A-F]+$/);
    expect(hole).not.toMatch(/^0+$/);
    expect(decoded(sig.get('Reason'))).toBe(placeholder.reason);
  });

  it('keeps ASCII values literals', () => {
    const latin1 = Buffer.from(
      renderStyledPdf(BODY, {
        registry,
        styles: EMPTY_STYLE_SHEET,
        signaturePlaceholder: { reason: 'I approve', name: 'Test Signer' },
      }),
    ).toString('latin1');
    expect(latin1).toContain('/Reason (I approve)');
    expect(latin1).toContain('/Name (Test Signer)');
    expect(latin1).toContain('/T (Signature1)');
  });
});
