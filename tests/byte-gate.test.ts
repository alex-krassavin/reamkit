// Byte gate — the project's core invariant, automated (oop-design.md §7).
//
// A PURE refactoring must keep PDF output byte-for-byte identical. The corpus
// pipeline diffs against LibreOffice through rasterization, and the other
// "byte-identical" tests check determinism within one revision — neither
// catches object renumbering BETWEEN revisions. These snapshots do.
//
// If this test fails on your change:
//   - refactoring? → you changed the bytes; find out why and undo it;
//   - deliberate output change? → review the diff, then `vitest run
//     tests/byte-gate.test.ts -u` and call out the gate update in the commit.

import { createHash } from 'node:crypto';
import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { convertDocxToPdfSync, convertXlsxToPdfSync } from '@/core/converter';

const FIXTURE_DIR = 'tests/fixtures/byte-gate';

const fonts = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
  italic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Italic.ttf')),
  boldItalic: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-BoldItalic.ttf')),
};

const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');

describe('byte gate: static fixtures', () => {
  // Synthetic documents checked into the repo (text, tables, lists,
  // headers/footers, sheet grids and number formats).
  for (const name of readdirSync(FIXTURE_DIR).sort()) {
    it(name, () => {
      const bytes = new Uint8Array(readFileSync(join(FIXTURE_DIR, name)));
      const pdf = name.endsWith('.xlsx')
        ? convertXlsxToPdfSync(bytes, { fonts })
        : convertDocxToPdfSync(bytes, { fonts });
      expect(sha256(pdf)).toMatchSnapshot();
    });
  }
});

describe('byte gate: images and PDF/A emit paths', () => {
  // Opaque PNG + semi-transparent PNG (SMask) inline in one paragraph; the
  // image pipeline is the most order-sensitive part of the emit phase.
  const opaque = buildTinyPng(4, 4, [0, 200, 0, 255]);
  const translucent = buildTinyPng(4, 4, [200, 0, 0, 128]);
  const body =
    '<w:p><w:r><w:t>images</w:t></w:r>' +
    '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
    '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId21"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  const docx = buildDocxFromBody(body, {
    images: {
      rId20: { contentType: 'image/png', bytes: opaque, extension: 'png' },
      rId21: { contentType: 'image/png', bytes: translucent, extension: 'png' },
    },
  });

  it('images.docx (plain: SMask kept)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts }))).toMatchSnapshot();
  });

  it('images.docx (PDF/A-1b: alpha flattened)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts, pdfA: 'PDF/A-1b' }))).toMatchSnapshot();
  });

  it('images.docx (PDF/A-2b: transparency group)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts, pdfA: 'PDF/A-2b' }))).toMatchSnapshot();
  });

  it('images.docx (PDF/A-1a: tagged structure tree)', () => {
    expect(sha256(convertDocxToPdfSync(docx, { fonts, pdfA: 'PDF/A-1a' }))).toMatchSnapshot();
  });
});
