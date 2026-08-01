import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody, buildRichDocx } from './fixtures/build-docx';
import { countShown, showPattern } from './fixtures/pdf-show';
import type { FamilyKey } from '@/core/fonts';
import { convertDocxToPdfSync } from '@/core/converter';
import { FontRegistry, parseTtf } from '@/core/font';

const here = dirname(fileURLToPath(import.meta.url));
const FONTS = {
  regular: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Regular.ttf'))),
  bold: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Bold.ttf'))),
  italic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-Italic.ttf'))),
  boldItalic: new Uint8Array(readFileSync(resolve(here, 'fixtures/fonts/Roboto-BoldItalic.ttf'))),
};

const latin1 = new TextDecoder('latin1');
const asLatin1 = (b: Uint8Array): string => latin1.decode(b);

describe('Styled rendering: rPr + pPr → PDF', () => {
  it('embeds a separate font resource for each variant used', () => {
    const docx = buildRichDocx([
      {
        runs: [
          { text: 'Reg ' },
          { text: 'Bold ', rPrXml: '<w:rPr><w:b/></w:rPr>' },
          { text: 'Italic ', rPrXml: '<w:rPr><w:i/></w:rPr>' },
          { text: 'BoldItalic', rPrXml: '<w:rPr><w:b/><w:i/></w:rPr>' },
        ],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    // Subsetted fonts carry a 6-letter subset tag prefix (PDF §9.6.4).
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-Bold\b/);
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-Italic\b/);
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-BoldItalic\b/);
    expect(text.match(/\/Subtype \/Type0/g)?.length).toBeGreaterThanOrEqual(4);
  });

  it('emits an rg color operator with the run colour', () => {
    const docx = buildRichDocx([
      {
        runs: [
          { text: 'Plain ' },
          { text: 'Red', rPrXml: '<w:rPr><w:color w:val="ff0000"/></w:rPr>' },
        ],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);
    expect(text).toMatch(/\b1 0 0 rg\b/);
  });

  it('falls back to bold-only when boldItalic font is missing', () => {
    const docx = buildRichDocx([{ runs: [{ text: 'BI', rPrXml: '<w:rPr><w:b/><w:i/></w:rPr>' }] }]);
    const pdf = convertDocxToPdfSync(docx, {
      fonts: { regular: FONTS.regular, bold: FONTS.bold },
    });
    const text = asLatin1(pdf);
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-Bold\b/);
    expect(text).not.toMatch(/\/BaseFont \/[A-Z]{6}\+Roboto-BoldItalic\b/);
  });

  it('emits different /Tf font sizes for runs with different sizes', () => {
    const docx = buildRichDocx([
      {
        runs: [
          { text: 'Small ', rPrXml: '<w:rPr><w:sz w:val="20"/></w:rPr>' },
          { text: 'Big', rPrXml: '<w:rPr><w:sz w:val="48"/></w:rPr>' },
        ],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);
    expect(text).toMatch(/\/F\d+ 10 Tf/);
    expect(text).toMatch(/\/F\d+ 24 Tf/);
  });

  it('renders table borders as path operators and places cell text', () => {
    const body = `
      <w:p><w:r><w:t>Before</w:t></w:r></w:p>
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="3000"/>
          <w:gridCol w:w="3000"/>
        </w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
      <w:p><w:r><w:t>After</w:t></w:r></w:p>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    expect(text).toMatch(/\d+(\.\d+)? w/);
    expect(text).toMatch(/m\n[^\n]*l\nS/);

    const tjMatches = [...text.matchAll(/<[0-9A-Fa-f]+> Tj/g)];
    expect(tjMatches.length).toBeGreaterThanOrEqual(6);
  });

  it('draws an internal separator defined on a cell right border (neighbour fallback)', () => {
    // No <w:tblBorders>; the table's only border is A1's RIGHT (the A|B
    // separator). Internal verticals are drawn on the right cell's LEFT side, so
    // without the neighbour fallback (cell.left ?? leftNeighbor.right) this
    // separator would vanish — the real-world bug where a contract table lost all
    // its column gridlines.
    const body = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:tcBorders><w:right w:val="single" w:sz="8" w:color="000000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
    // The separator is emitted as a stroked path (m … l … S).
    expect(text).toMatch(/m\n[^\n]*l\nS/);
  });

  it('resolves a shared border by weight — the heavier one wins (§17.4)', () => {
    // The A|B edge is specified on BOTH cells: A.right = 2pt (sz 16), B.left =
    // 0.5pt (sz 4). Border-conflict resolution keeps the heavier (2pt) and
    // discards the lighter.
    const body = `
      <w:tbl>
        <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:tcBorders><w:right w:val="single" w:sz="16" w:color="000000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:tcBorders><w:left w:val="single" w:sz="4" w:color="000000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
    expect(text).toMatch(/(^|\s)2 w(\s|$)/); // the 2pt edge is drawn
    expect(text).not.toMatch(/(^|\s)0\.5 w(\s|$)/); // the 0.5pt one is discarded
  });

  it('justify lines use per-token Tm positioning (more Tms than left-aligned)', () => {
    const longText =
      'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor ' +
      'incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam quis ' +
      'nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo.';
    const left = buildRichDocx([
      { pPrXml: '<w:pPr><w:jc w:val="left"/></w:pPr>', runs: [{ text: longText }] },
    ]);
    const both = buildRichDocx([
      { pPrXml: '<w:pPr><w:jc w:val="both"/></w:pPr>', runs: [{ text: longText }] },
    ]);
    const pdfLeft = asLatin1(convertDocxToPdfSync(left, { fonts: FONTS }));
    const pdfBoth = asLatin1(convertDocxToPdfSync(both, { fonts: FONTS }));

    const tmLeft = (pdfLeft.match(/Tm/g) ?? []).length;
    const tmBoth = (pdfBoth.match(/Tm/g) ?? []).length;
    expect(tmBoth).toBeGreaterThan(tmLeft);
  });

  it('renders a gridSpan=2 cell with the combined width of two columns', () => {
    const body = `
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="2000"/>
          <w:gridCol w:w="2000"/>
        </w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Merged header that spans both columns</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const tjMatches = [...text.matchAll(/<[0-9A-Fa-f]+> Tj/g)];
    expect(tjMatches.length).toBeGreaterThanOrEqual(3);
  });

  it('does not render content for vMerge=continue cells', () => {
    const body = `
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="2000"/>
          <w:gridCol w:w="2000"/>
        </w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr><w:vMerge w:val="restart"/></w:tcPr>
            <w:p><w:r><w:t>StartCell</w:t></w:r></w:p>
          </w:tc>
          <w:tc><w:p><w:r><w:t>RightTop</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc>
            <w:tcPr><w:vMerge/></w:tcPr>
            <w:p><w:r><w:t>ShouldBeHidden</w:t></w:r></w:p>
          </w:tc>
          <w:tc><w:p><w:r><w:t>RightBottom</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });

    const parsed = parseTtf(FONTS.regular);
    const text = asLatin1(pdf);

    expect(text).toMatch(showPattern(parsed, 'StartCell'));
    expect(text).not.toMatch(showPattern(parsed, 'ShouldBeHidden'));
  });

  it('does not over-justify a short single-line paragraph (last line stays left)', () => {
    const docx = buildRichDocx([
      { pPrXml: '<w:pPr><w:jc w:val="both"/></w:pPr>', runs: [{ text: 'Short.' }] },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const btMatch = text.match(/BT\n([\s\S]*?)\nET/);
    expect(btMatch).not.toBeNull();
    const tmInBt = (btMatch![1]!.match(/Tm/g) ?? []).length;
    expect(tmInBt).toBe(1);
  });

  it('pulls an over-full justified line back to the measure', () => {
    // The breaker weighs a line knowing it may SHRINK each space by up to 30%
    // of its width, so it will choose one whose natural width is over the
    // measure. Drawn at that natural width the line ran into the right margin —
    // IllustrativeCases.docx put three of its four opening lines 1 to 10pt past
    // it while the fourth sat short, which reads as no justification at all.
    // Every full line of a justified paragraph is placed token by token, the
    // slack (of either sign) shared out between its spaces; a line emitted with
    // a single Tm is one that was drawn at its natural width.
    const words = 'alpha be gamma d epsilon zeta et theta iota kappa lam mu nu xi'.split(' ');
    const body = Array.from({ length: 200 }, (_, i) => words[i % words.length]!).join(' ');
    const docx = buildRichDocx([
      { pPrXml: '<w:pPr><w:jc w:val="both"/></w:pPr>', runs: [{ text: body }] },
    ]);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));

    // One BT spans the whole page, so lines are told apart by their baseline —
    // the y of each `Tm`. A line placed token by token has several.
    const perBaseline = new Map<string, number>();
    for (const m of text.matchAll(/1 0 0 1 [\d.]+ ([\d.]+) Tm/gu)) {
      perBaseline.set(m[1]!, (perBaseline.get(m[1]!) ?? 0) + 1);
    }
    const counts = [...perBaseline.values()];
    expect(counts.length).toBeGreaterThan(8);
    // All but the last line of the paragraph — the last one stays left-aligned.
    expect(counts.slice(0, -1).every((tms) => tms > 1)).toBe(true);
  });

  it('embeds the glyph a tab leader draws with (§17.3.1.38)', () => {
    // A leader's characters are made by the layout, from the stop, long after
    // the subset is chosen from the runs. Left out of it, the subset had no
    // glyph for them: TOC_field_b.docx drew its dot leader as a row of
    // missing-glyph boxes running past the right margin.
    const docx = buildRichDocx([
      {
        pPrXml:
          '<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs></w:pPr>',
        runs: [{ text: 'Heading\t1' }],
      },
    ]);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    const parsed = parseTtf(FONTS.regular);
    // The leader is drawn as a long run of the dot's own glyph. Pruned from the
    // subset, the same run came out as the missing-glyph id 0.
    const dot = parsed.glyphForCodepoint('.'.codePointAt(0)!).toString(16).padStart(4, '0');
    expect(dot).not.toBe('0000');
    expect(text).toMatch(new RegExp(`(?:${dot}){40,}`, 'iu'));
  });

  it('draws the rules a paragraph asks for around itself', () => {
    // §17.3.1.24 — one stroked rule per declared edge, each in its own colour.
    const docx = buildRichDocx([
      {
        pPrXml:
          '<w:pPr><w:pBdr>' +
          '<w:top w:val="single" w:sz="48" w:space="1" w:color="DE81E1"/>' +
          '<w:bottom w:val="single" w:sz="48" w:space="1" w:color="90ABF0"/>' +
          '</w:pBdr></w:pPr>',
        runs: [{ text: 'Sample Text' }],
      },
    ]);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // 0xDE/255 = 0.871, 0x81/255 = 0.506, 0xE1/255 = 0.882 — the top rule.
    expect(text).toMatch(/0\.871 0\.506 0\.882 RG/u);
    // 0x90/255 = 0.565, 0xAB/255 = 0.671, 0xF0/255 = 0.941 — the bottom one.
    expect(text).toMatch(/0\.565 0\.671 0\.941 RG/u);
    // 48 eighths of a point = 6pt wide.
    expect(text).toMatch(/\n6 w\n/u);
  });

  it('paints the background a paragraph asks for behind it', () => {
    const docx = buildRichDocx([
      { pPrXml: '<w:pPr><w:shd w:val="clear" w:fill="F4E7D3"/></w:pPr>', runs: [{ text: 'Band' }] },
    ]);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // 0xF4/255 = 0.957, 0xE7/255 = 0.906, 0xD3/255 = 0.827.
    expect(text).toMatch(/0\.957 0\.906 0\.827 rg/u);
    // …and it is a fill under the text, not a stroke around it.
    expect(text).toMatch(/0\.957 0\.906 0\.827 rg\n[\d.]+ [\d.]+ [\d.]+ [\d.]+ re\nf/u);
  });

  it('sends an absolute-position tab to the middle and the far side (§17.3.3.15)', () => {
    // A `w:ptab` has no distance of its own: it reaches for an edge of the
    // column. Read nowhere, SimpleHeadThreeColFoot.docx printed its three
    // footer regions as "Footer LeftFooter MiddleFooter Right".
    const docx = buildDocxFromBody(
      '<w:p>' +
        '<w:r><w:t>L</w:t></w:r>' +
        '<w:r><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/></w:r>' +
        '<w:r><w:t>M</w:t></w:r>' +
        '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>' +
        '<w:r><w:t>R</w:t></w:r>' +
        '</w:p>',
    );
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    const xs = [...text.matchAll(/1 0 0 1 ([\d.]+) [\d.]+ Tm/gu)].map((m) => Number(m[1]));
    expect(xs.length).toBeGreaterThan(2);
    // A4 less 1in margins is 451pt of column, so the centre tab lands near
    // 72 + 451/2 and the right one carries its letter to the far margin.
    expect(Math.max(...xs)).toBeGreaterThan(500);
    expect(xs.some((x) => x > 280 && x < 310)).toBe(true);
  });

  it('draws a run set in capitals in capitals (§17.3.2.5)', () => {
    // capitalized.docx prints its word in lower case where every other reader
    // shouts it — and the subset has to carry the glyphs that are DRAWN, not
    // the ones the run stores, or the same word comes out as missing glyphs.
    const docx = buildRichDocx([{ runs: [{ text: 'shout', rPrXml: '<w:rPr><w:caps/></w:rPr>' }] }]);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    const parsed = parseTtf(FONTS.regular);
    expect(text).toMatch(showPattern(parsed, 'SHOUT'));
    expect(text).not.toMatch(showPattern(parsed, 'shout'));
  });

  it('sets the lower case of a small-capitals run smaller (§17.3.2.33)', () => {
    const docx = buildRichDocx([
      { runs: [{ text: 'Ab', rPrXml: '<w:rPr><w:smallCaps/><w:sz w:val="40"/></w:rPr>' }] },
    ]);
    const text = asLatin1(convertDocxToPdfSync(docx, { fonts: FONTS }));
    // 20pt for the letter that was already a capital, four fifths for the other.
    expect(text).toMatch(/\/F\d+ 20 Tf/u);
    expect(text).toMatch(/\/F\d+ 16 Tf/u);
  });

  it('splits a table row taller than the page into chunks across pages', () => {
    // 80 paragraphs in one cell exceeds A4 content height (~698pt) at typical
    // 14pt line height. Expect at least two pages with continuation borders.
    const paragraphs = Array.from(
      { length: 80 },
      (_, i) => `<w:p><w:r><w:t>Line${i}</w:t></w:r></w:p>`,
    ).join('');
    const body = `
      <w:tbl>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
        <w:tr>
          <w:tc>${paragraphs}</w:tc>
        </w:tr>
      </w:tbl>`;
    const docx = buildDocxFromBody(body);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const pageCount = (text.match(/\/Type \/Page\b/g) ?? []).filter(
      (m) => !m.includes('Pages'),
    ).length;
    expect(pageCount).toBeGreaterThan(1);

    const parsed = parseTtf(FONTS.regular);
    // First and last lines must both render (row split, not clipping).
    expect(text).toMatch(showPattern(parsed, 'Line0'));
    expect(text).toMatch(showPattern(parsed, 'Line79'));
  });

  it('honors paragraph alignment center and right', () => {
    const docx = buildRichDocx([
      {
        pPrXml: '<w:pPr><w:jc w:val="center"/></w:pPr>',
        runs: [{ text: 'Centered' }],
      },
      {
        pPrXml: '<w:pPr><w:jc w:val="right"/></w:pPr>',
        runs: [{ text: 'Right' }],
      },
    ]);
    const pdf = convertDocxToPdfSync(docx, { fonts: FONTS });
    const text = asLatin1(pdf);

    const tmMatches = [...text.matchAll(/1 0 0 1 ([\d.]+) [\d.]+ Tm/g)];
    expect(tmMatches.length).toBeGreaterThanOrEqual(2);
    const xPositions = tmMatches.map((m) => Number(m[1]));
    // First line is centered → x > marginLeft (72)
    expect(xPositions[0]!).toBeGreaterThan(72);
    // Second line is right-aligned → even further right than centered
    expect(xPositions[1]!).toBeGreaterThan(xPositions[0]!);
  });

  it('renders a nested table inside a cell (table-in-cell)', () => {
    // A w:tbl nested in a w:tc. Previously the cell layout skipped non-paragraph
    // content, so the nested table (and all its text) was dropped — the POI
    // 60329.docx pattern (0/4812 chars). Now it lays out and renders.
    const body = `
      <w:tbl>
        <w:tblPr><w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:color="000000"/>
        </w:tblBorders></w:tblPr>
        <w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
        <w:tr><w:tc>
          <w:p><w:r><w:t>OUTERCELL</w:t></w:r></w:p>
          <w:tbl>
            <w:tblPr><w:tblBorders>
              <w:top w:val="single" w:sz="4" w:color="000000"/>
              <w:insideV w:val="single" w:sz="4" w:color="000000"/>
            </w:tblBorders></w:tblPr>
            <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
            <w:tr>
              <w:tc><w:p><w:r><w:t>NESTEDA</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>NESTEDB</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
        </w:tc></w:tr>
      </w:tbl>`;
    const text = asLatin1(convertDocxToPdfSync(buildDocxFromBody(body), { fonts: FONTS }));
    const parsed = parseTtf(FONTS.regular);
    expect(text).toMatch(showPattern(parsed, 'OUTERCELL')); // outer cell paragraph
    expect(text).toMatch(showPattern(parsed, 'NESTEDA')); // nested cell 1 (was lost)
    expect(text).toMatch(showPattern(parsed, 'NESTEDB')); // nested cell 2 (was lost)
  });

  it('measures table auto-layout with per-family fonts (was: bare-variant lookup crash)', () => {
    const reg = (b: Uint8Array): FontRegistry => FontRegistry.fromBytes({ regular: b });
    const registriesByFamily: ReadonlyMap<FamilyKey, FontRegistry> = new Map([
      ['arimo', reg(FONTS.regular)],
      ['tinos', reg(FONTS.bold)],
    ]);
    // Auto-layout table (no explicit grid widths) forces measureSingleLine,
    // which used to look fontResources up by bare variant and crash when the
    // keys are per-family ('roboto:regular', …).
    const body =
      `<w:tbl><w:tr>` +
      `<w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>CELLSANS</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/></w:rPr><w:t>CELLSERIF</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`;
    const pdf = convertDocxToPdfSync(buildDocxFromBody(body), {
      fonts: { regular: FONTS.regular },
      registriesByFamily,
    });
    // Pre-fix this crashed (fontResources.get('regular') is undefined when
    // keys are per-family). Surviving + embedding BOTH family stand-ins
    // proves each cell measured and rendered with its own family.
    const baseFonts = new Set(
      [...asLatin1(pdf).matchAll(/\/BaseFont \/[A-Z]{6}\+([A-Za-z-]+)/g)].map((m) => m[1]),
    );
    expect(baseFonts.has('Roboto-Regular')).toBe(true); // Arial → arimo
    expect(baseFonts.has('Roboto-Bold')).toBe(true); // Times → tinos stand-in
  });

  it('resolves the substitute font per run (sans / serif / mono families)', () => {
    // Distinct stand-in fonts per family (Roboto variants carry distinct
    // BaseFont names) so we can prove each run picked its OWN family by w:ascii.
    const reg = (b: Uint8Array): FontRegistry => FontRegistry.fromBytes({ regular: b });
    const registriesByFamily: ReadonlyMap<FamilyKey, FontRegistry> = new Map([
      ['arimo', reg(FONTS.regular)],
      ['tinos', reg(FONTS.bold)],
      ['cousine', reg(FONTS.italic)],
    ]);
    const body =
      `<w:p>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>SANS</w:t></w:r>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman"/></w:rPr><w:t>SERIF</w:t></w:r>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr><w:t>MONO</w:t></w:r>` +
      `</w:p>`;
    const pdf = convertDocxToPdfSync(buildDocxFromBody(body), {
      fonts: { regular: FONTS.regular },
      registriesByFamily,
    });
    const baseFonts = new Set(
      [...asLatin1(pdf).matchAll(/\/BaseFont \/[A-Z]{6}\+([A-Za-z-]+)/g)].map((m) => m[1]),
    );
    // Each run resolved to its family's stand-in → three distinct fonts embedded.
    expect(baseFonts.has('Roboto-Regular')).toBe(true); // SANS → arimo
    expect(baseFonts.has('Roboto-Bold')).toBe(true); // SERIF → tinos stand-in
    expect(baseFonts.has('Roboto-Italic')).toBe(true); // MONO → cousine stand-in
  });
});
describe('super/subscript (§17.3.2.42 w:vertAlign)', () => {
  it('draws smaller and off the baseline, not full size on it', () => {
    // The model carried the flag and the HTML writer honoured it; the PDF
    // layout did neither, so a footnote marker and a cell's "Salary⁽²⁾" came out
    // full size on the line (45540_classic_Header.xlsx).
    const pdf = convertDocxToPdfSync(
      buildRichDocx([
        {
          runs: [
            { text: 'x' },
            { text: '2', rPrXml: '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' },
            { text: '3', rPrXml: '<w:rPr><w:vertAlign w:val="subscript"/></w:rPr>' },
          ],
        },
      ]),
      { fonts: FONTS },
    );
    const text = new TextDecoder('latin1').decode(pdf);
    // Three Tm placements on one line — a line of plain text emits one.
    const yOf = [...text.matchAll(/1 0 0 1 [\d.-]+ ([\d.-]+) Tm/g)].map((m) => Number(m[1]));
    expect(yOf.length).toBeGreaterThanOrEqual(3);
    const [base, sup, sub] = yOf;
    expect(sup!).toBeGreaterThan(base!);
    expect(sub!).toBeLessThan(base!);
  });
});
