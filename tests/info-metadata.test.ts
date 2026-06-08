import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { convertDocxToPdfSync } from '@/converter';
import { parseCoreProperties } from '@/opc';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
};
const encoder = new TextEncoder();
const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

describe('parseCoreProperties', () => {
  it('extracts Dublin Core fields from docProps/core.xml', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Quarterly Report</dc:title>
  <dc:creator>Alice</dc:creator>
  <dc:subject>Finance Q1</dc:subject>
  <cp:keywords>q1 finance report</cp:keywords>
  <dcterms:created>2026-01-15T10:30:00Z</dcterms:created>
  <dcterms:modified>2026-04-01T12:00:00Z</dcterms:modified>
</cp:coreProperties>`;
    const core = parseCoreProperties(encoder.encode(xml));
    expect(core.title).toBe('Quarterly Report');
    expect(core.creator).toBe('Alice');
    expect(core.subject).toBe('Finance Q1');
    expect(core.keywords).toBe('q1 finance report');
    expect(core.created?.toISOString()).toBe('2026-01-15T10:30:00.000Z');
    expect(core.modified?.toISOString()).toBe('2026-04-01T12:00:00.000Z');
  });

  it('returns empty result when the XML is missing or malformed', () => {
    expect(parseCoreProperties(encoder.encode('<?xml version="1.0"?><other/>'))).toEqual({});
  });
});

describe('PDF /Info dictionary integration', () => {
  it('emits /Info with default Producer when no metadata is provided', () => {
    const docx = buildDocxFromBody('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);
    // Trailer must reference Info.
    expect(text).toMatch(/\/Info \d+ \d+ R/);
    // Producer must be the bundled default ("Ream") in UTF-16BE
    // hex string starting with the FE FF BOM marker.
    expect(text).toContain('/Producer <FEFF');
    // Hex-encoded "Ream".
    const expected = utf16BeHex('Ream');
    expect(text).toContain(`/Producer <FEFF${expected}>`);
  });

  it('reads core properties from docProps/core.xml and surfaces them in /Info', () => {
    const docxBytes = buildDocxWithCoreProps('<w:p><w:r><w:t>x</w:t></w:r></w:p>', {
      title: 'My Title',
      creator: 'Alice',
      subject: 'Subject',
      keywords: 'one two',
      created: '2026-01-15T10:30:00Z',
      modified: '2026-04-01T12:00:00Z',
    });
    const pdf = convertDocxToPdfSync(docxBytes, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text).toContain(`/Title <FEFF${utf16BeHex('My Title')}>`);
    expect(text).toContain(`/Author <FEFF${utf16BeHex('Alice')}>`);
    expect(text).toContain(`/Subject <FEFF${utf16BeHex('Subject')}>`);
    expect(text).toContain(`/Keywords <FEFF${utf16BeHex('one two')}>`);
    expect(text).toContain('/CreationDate (D:20260115103000Z)');
    expect(text).toContain('/ModDate (D:20260401120000Z)');
  });

  it('caller-supplied options.info overrides values from core.xml', () => {
    const docxBytes = buildDocxWithCoreProps('<w:p><w:r><w:t>x</w:t></w:r></w:p>', {
      title: 'From Core',
      creator: 'Core Author',
    });
    const pdf = convertDocxToPdfSync(docxBytes, {
      fonts: FONTS,
      info: { title: 'Override', author: 'Override Author', producer: 'Custom Producer' },
    });
    const text = asLatin1(pdf);
    expect(text).toContain(`/Title <FEFF${utf16BeHex('Override')}>`);
    expect(text).toContain(`/Author <FEFF${utf16BeHex('Override Author')}>`);
    expect(text).toContain(`/Producer <FEFF${utf16BeHex('Custom Producer')}>`);
    expect(text).not.toContain(`/Title <FEFF${utf16BeHex('From Core')}>`);
  });

  it('formats arbitrary Unicode (Cyrillic) into UTF-16BE hex correctly', () => {
    const docxBytes = buildDocxWithCoreProps('<w:p><w:r><w:t>x</w:t></w:r></w:p>', {
      title: 'Привет мир',
      creator: 'Аноним',
    });
    const pdf = convertDocxToPdfSync(docxBytes, { fonts: FONTS });
    const text = asLatin1(pdf);
    expect(text).toContain(`/Title <FEFF${utf16BeHex('Привет мир')}>`);
    expect(text).toContain(`/Author <FEFF${utf16BeHex('Аноним')}>`);
  });
});

function utf16BeHex(s: string): string {
  let hex = '';
  for (const ch of s) {
    const cp = ch.codePointAt(0)!;
    if (cp < 0x10000) {
      hex += cp.toString(16).padStart(4, '0').toUpperCase();
    } else {
      const adj = cp - 0x10000;
      const hi = 0xd800 + (adj >> 10);
      const lo = 0xdc00 + (adj & 0x3ff);
      hex += hi.toString(16).padStart(4, '0').toUpperCase();
      hex += lo.toString(16).padStart(4, '0').toUpperCase();
    }
  }
  return hex;
}

// Reuses buildDocxFromBody but also adds a docProps/core.xml part.
function buildDocxWithCoreProps(
  bodyInner: string,
  props: {
    title?: string;
    creator?: string;
    subject?: string;
    keywords?: string;
    created?: string;
    modified?: string;
  },
): Uint8Array {
  // Build the standard docx first, then surgically add core.xml to the zip.
  const docx = buildDocxFromBody(bodyInner);
  const existing = unzipSync(docx);
  const escape = (s: string) =>
    s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const coreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/">
  ${props.title ? `<dc:title>${escape(props.title)}</dc:title>` : ''}
  ${props.creator ? `<dc:creator>${escape(props.creator)}</dc:creator>` : ''}
  ${props.subject ? `<dc:subject>${escape(props.subject)}</dc:subject>` : ''}
  ${props.keywords ? `<cp:keywords>${escape(props.keywords)}</cp:keywords>` : ''}
  ${props.created ? `<dcterms:created>${props.created}</dcterms:created>` : ''}
  ${props.modified ? `<dcterms:modified>${props.modified}</dcterms:modified>` : ''}
</cp:coreProperties>`;
  existing['docProps/core.xml'] = encoder.encode(coreXml);
  return zipSync(existing);
}
