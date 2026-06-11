// Formal PDF/A conformance via veraPDF (https://verapdf.org). This is an
// opt-in integration gate: it runs only when a veraPDF CLI is available (set the
// VERAPDF env var, or install to ~/verapdf/verapdf), and is skipped otherwise —
// veraPDF is an external Java tool not present in every environment (e.g. CI).
//
// It exercises the conformance-critical paths that structural tests can miss and
// that real validation has caught: subset /CIDSet (§6.3.5), ToUnicode coverage
// of ligature and list-marker-TAB glyphs (§6.3.8), tagged structure (Level A).

import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { buildTinyPng } from './fixtures/build-png';
import { convertDocxToPdfSync } from '@/core/converter';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
};

function findVeraPdf(): string | null {
  const candidates = [process.env.VERAPDF, resolve(homedir(), 'verapdf/verapdf')].filter(
    (c): c is string => Boolean(c),
  );
  return candidates.find((c) => existsSync(c)) ?? null;
}
const VP = findVeraPdf();

// Run veraPDF and return its one-line text verdict ("PASS <file> <flavour>" or
// "FAIL …"). veraPDF exits non-zero on a validation failure, so capture stdout.
type Flavour = '1a' | '1b' | '2a' | '2b' | '2u' | '3a' | '3b' | '3u';
function validate(vp: string, pdf: Uint8Array, flavour: Flavour): string {
  const file = resolve(mkdtempSync(resolve(tmpdir(), 'vpdf-')), 'out.pdf');
  writeFileSync(file, pdf);
  try {
    return execFileSync(vp, ['--flavour', flavour, '--format', 'text', file], {
      encoding: 'utf8',
    }).trim();
  } catch (e) {
    const stdout = (e as { stdout?: Buffer | string }).stdout;
    return (stdout ? stdout.toString() : '').trim();
  }
}

// A document covering the conformance-critical paths: ligatures ("file"/"flow"),
// a numbered list (marker TAB glyph), a table, and a figure with alt text.
function richDocx(): Uint8Array {
  const numberingXml =
    '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/></w:lvl>' +
    '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const li = (l: number, t: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${l}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const pic =
    '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Pic" descr="A red square"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  const tbl =
    '<w:tbl><w:tblPr><w:tblBorders>' +
    '<w:top w:val="single" w:sz="6"/><w:bottom w:val="single" w:sz="6"/>' +
    '<w:left w:val="single" w:sz="6"/><w:right w:val="single" w:sz="6"/>' +
    '<w:insideH w:val="single" w:sz="6"/><w:insideV w:val="single" w:sz="6"/></w:tblBorders></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:shd w:fill="D9E2F3"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Feature</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>Done</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const body =
    '<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Heading</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>The file flows efficiently through the office.</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>See </w:t></w:r><w:hyperlink r:id="rId30"><w:r><w:t>the project site</w:t></w:r></w:hyperlink><w:r><w:t> for details.</w:t></w:r>' +
    '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>' +
    tbl +
    li(0, 'First item') +
    li(1, 'Nested item') +
    pic;
  return buildDocxFromBody(body, {
    numberingXml,
    hyperlinks: { rId30: 'https://reamkit.dev' },
    footnotesXml:
      '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> A footnote, flowing efficiently.</w:t></w:r></w:p></w:footnote>',
    images: {
      rId20: {
        contentType: 'image/png' as const,
        bytes: buildTinyPng(4, 4, [200, 50, 50, 255]),
        extension: 'png' as const,
      },
    },
  });
}

// A document with a semi-transparent PNG (soft mask) — PDF/A-2/3 keep the
// transparency, so the page carries a transparency group.
function translucentDocx(): Uint8Array {
  const drawing =
    '<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
    '<wp:extent cx="457200" cy="457200"/><wp:docPr id="1" name="Pic" descr="A translucent red square"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId20"/></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  return buildDocxFromBody(
    `<w:p><w:r><w:t>Translucent figure follows.</w:t></w:r></w:p><w:p>${drawing}</w:p>`,
    {
      images: {
        rId20: {
          contentType: 'image/png' as const,
          bytes: buildTinyPng(4, 4, [200, 50, 50, 128]),
          extension: 'png' as const,
        },
      },
    },
  );
}

describe.skipIf(!VP)('veraPDF formal conformance', () => {
  it('validates a plain document (with ligatures) as PDF/A-1b', () => {
    const pdf = convertDocxToPdfSync(
      buildDocxFromBody(
        '<w:p><w:r><w:t>Plain body with file (fi) and flow (fl).</w:t></w:r></w:p>',
      ),
      { fonts: FONTS, pdfA: 'PDF/A-1b' },
    );
    expect(validate(VP!, pdf, '1b')).toMatch(/^PASS/);
  });

  it('validates a rich tagged document as PDF/A-1a (Level A)', () => {
    const pdf = convertDocxToPdfSync(richDocx(), { fonts: FONTS, pdfA: 'PDF/A-1a' });
    expect(validate(VP!, pdf, '1a')).toMatch(/^PASS/);
  });

  it('validates a translucent-image document as PDF/A-2b (transparency kept)', () => {
    const pdf = convertDocxToPdfSync(translucentDocx(), { fonts: FONTS, pdfA: 'PDF/A-2b' });
    expect(validate(VP!, pdf, '2b')).toMatch(/^PASS/);
  });

  it('validates a rich tagged document as PDF/A-2a', () => {
    const pdf = convertDocxToPdfSync(richDocx(), { fonts: FONTS, pdfA: 'PDF/A-2a' });
    expect(validate(VP!, pdf, '2a')).toMatch(/^PASS/);
  });

  it('validates a document with an embedded source as PDF/A-3b', () => {
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>PDF/A-3b with an embedded source file.</w:t></w:r></w:p>',
    );
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS, pdfA: 'PDF/A-3b', embedSource: true });
    expect(validate(VP!, pdf, '3b')).toMatch(/^PASS/);
  });

  it('validates a rich tagged document with an embedded source as PDF/A-3a', () => {
    const pdf = convertDocxToPdfSync(richDocx(), {
      fonts: FONTS,
      pdfA: 'PDF/A-3a',
      embedSource: true,
    });
    expect(validate(VP!, pdf, '3a')).toMatch(/^PASS/);
  });
});
