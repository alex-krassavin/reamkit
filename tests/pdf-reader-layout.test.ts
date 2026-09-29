// E-PDF EP4 — heuristic reconstruction for untagged PDFs. Convert a docx to a
// plain (untagged) PDF, then rebuild a FlowDoc from the positioned text alone:
// lines clustered by baseline, paragraphs by vertical spacing, headings by a
// font size above the median.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import type { BodyElement } from '@/core/document-model';
import { Ream } from '@/core/converter/ream';
import { OpcPackage } from '@/core/opc';
import { PdfFile } from '@/pdf-reader/document';
import { BASELINE_AT, FLOAT_CARRIER, positionedText } from '@/pdf-reader/flow-build';
import { drawnWords, endedParagraph, reconstructByLayout } from '@/pdf-reader/layout';
import { extractPageText } from '@/pdf-reader/text';
import { writeDocx } from '@/word/docx-writer';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const spaced = (text: string): string =>
  `<w:p><w:pPr><w:spacing w:after="200"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function layoutFlow(body: string) {
  const pdf = await Ream.parse(buildDocxFromBody(body)).convert('pdf', { fonts: FONTS });
  return reconstructByLayout(PdfFile.parse(pdf)).doc;
}

const paragraphs = (flow: { body: ReadonlyArray<{ kind: string }> }) =>
  flow.body.filter(
    (
      b,
    ): b is {
      kind: 'paragraph';
      paragraph: {
        properties: {
          outlineLevel?: number;
          tabs?: ReadonlyArray<unknown>;
          borders?: { top?: { style?: string } };
        };
        runs: ReadonlyArray<{ text: string }>;
      };
    } => b.kind === 'paragraph',
  );

describe('a heuristic line keeps how it looked (E-PDF EP4)', () => {
  it('carries each run’s size and colour, not just its letters', async () => {
    // The tagged path has carried these since it learned to; this one never
    // did, so every line came back at the 11pt default in black. Placed, that
    // is not a wrong shade but a wrong SHAPE: 160F-2019.pdf's footnotes are set
    // in 7pt nine and a half apart, and drawn at eleven they climbed over each
    // other.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:rPr><w:sz w:val="14"/><w:color w:val="FF0000"/></w:rPr>' +
        '<w:t>SmallRedText</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const flow = reconstructByLayout(PdfFile.parse(pdf)).doc;
    const runs = flow.body.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : []));
    const small = runs.find((r) => r.text.includes('SmallRed'));
    expect(small).toBeDefined();
    expect(small!.properties.fontSizePt).toBeCloseTo(7, 0);
    expect(small!.properties.colorHex).toBe('FF0000');
  });
});

describe('a multi-page PDF keeps its pages (E-PDF EP4)', () => {
  it('opens an output page for each source page after the first', async () => {
    // Flowed, the layout repaginates and this hardly shows. PLACED, every mark
    // is anchored to "the page", so without a break all twenty-five pages of
    // Brotli-Prototype-FileA.pdf stacked onto one.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>PageOne</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>PageTwo</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const file = PdfFile.parse(pdf);
    expect(file.pages().length).toBe(2);
    const placed = reconstructByLayout(file, 'positional').doc;
    const breaks = placed.body.filter(
      (b) => b.kind === 'paragraph' && b.paragraph.properties.pageBreakBefore === true,
    );
    expect(breaks).toHaveLength(file.pages().length - 1);
  });

  it('turns widow control off, so a column the page broke is not broken again (§17.3.1.44)', async () => {
    // Every source page opens a page of its own; where a reconstructed column
    // runs a line long, widow control would carry a second line along.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>PageOne</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>PageTwo</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const doc = reconstructByLayout(PdfFile.parse(pdf)).doc;
    const found = paragraphs(doc);
    expect(found.length).toBeGreaterThan(0);
    for (const p of found) {
      expect((p.paragraph.properties as { widowControl?: boolean }).widowControl).toBe(false);
    }
    const body = new TextDecoder().decode(
      OpcPackage.open(writeDocx(doc).bytes).getMainDocument().data,
    );
    expect(body).toContain('<w:widowControl w:val="0"/>');
  });

  it('opens a SECTION where the page size changes', () => {
    // §17.6 — a section is what carries a page size, so a document whose pages
    // differ in size is several of them. function_based_shading_cmyk.pdf is
    // 290×290 and then 1880×1260, and read as one size the second sheet's six
    // squares were cut down to the one that fitted.
    const file = PdfFile.parse(twoSizePdf());
    const doc = reconstructByLayout(file, 'positional').doc;
    expect(doc.sections).toHaveLength(2);
    expect(doc.sections[0]?.properties.pageSize?.width).toBeCloseTo(200, 1);
    expect(doc.sections[1]?.properties.pageSize?.width).toBeCloseTo(600, 1);
    // The break the pages would otherwise carry is the section's own: two
    // sections, and no page-break paragraph between them.
    const breaks = doc.body.filter(
      (b) => b.kind === 'paragraph' && b.paragraph.properties.pageBreakBefore === true,
    );
    expect(breaks).toHaveLength(0);
    // And a document of ONE size states no sections at all.
    expect(
      reconstructByLayout(
        PdfFile.parse(onePagePdf('/MediaBox [0 0 200 100]', 'BT ET')),
        'positional',
      ).doc.sections,
    ).toHaveLength(0);
  });
});

describe('a page of turned words is a page, not prose (§9.4.2)', () => {
  it('reads it placed, whatever else is on the sheet', () => {
    // The placement IS the content: re-set flat, the words come back in an
    // order the page never had. bug946506.pdf runs every line of its lorem
    // ipsum down the sheet at twenty degrees, and read as prose its lines
    // interleaved — "adipiscinnon luctus eleipsum dolor sit".
    const turned = Array.from(
      { length: 10 },
      (_, i) =>
        `BT /F0 12 Tf 0.94 0.34 -0.34 0.94 ${String(30 + i * 8)} ${String(40 + i * 18)} Tm (word${String(i)}) Tj ET`,
    ).join('\n');
    const doc = Ream.parse(
      onePagePdf('/MediaBox [0 0 300 300] /Resources << /Font << /F0 5 0 R >> >>', turned, [
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
      ]),
    );
    expect(doc.losses.some((l) => /read as a PAGE/u.test(l.detail))).toBe(true);

    // An upright page of the same size is prose and keeps the flowing reading.
    const upright = Array.from(
      { length: 10 },
      (_, i) => `BT /F0 12 Tf 1 0 0 1 30 ${String(40 + i * 18)} Tm (word${String(i)}) Tj ET`,
    ).join('\n');
    const flowed = Ream.parse(
      onePagePdf('/MediaBox [0 0 300 300] /Resources << /Font << /F0 5 0 R >> >>', upright, [
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
      ]),
    );
    expect(flowed.losses.some((l) => /read as a PAGE/u.test(l.detail))).toBe(false);
  });

  it('reads a GRID of boxes with a label in each as the page it is', () => {
    // calgray.pdf is five rows of four grey swatches, each labelled "A = 0.75"
    // and the like. Nineteen of the twenty boxes are a mark anybody can see —
    // the twentieth is painted white — and the count had to reach twenty before
    // the ratio was consulted at all. One short, the page was read as prose: the
    // four labels of each row ran together into a line and the sheet spilled
    // onto a second page.
    const cells: Array<string> = [];
    for (let row = 0; row < 5; row++) {
      for (let col = 0; col < 4; col++) {
        const x = 20 + col * 65;
        const y = 30 + row * 50;
        cells.push(`0.${String(row + 3)} g ${String(x)} ${String(y)} 60 45 re f`);
        cells.push(
          `0 g BT /F0 8 Tf 1 0 0 1 ${String(x + 4)} ${String(y + 6)} Tm (A=0.${String(row)}${String(col)}) Tj ET`,
        );
      }
    }
    const doc = Ream.parse(
      onePagePdf(
        '/MediaBox [0 0 300 300] /Resources << /Font << /F0 5 0 R >> >>',
        cells.join('\n'),
        ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
      ),
    );
    expect(doc.losses.some((l) => /read as a PAGE/u.test(l.detail))).toBe(true);
  });
});

describe('a paragraph keeps the indent the page set it with (§17.3.1.12)', () => {
  const parasOf = (content: string) => {
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf('/MediaBox [0 0 400 400] /Resources << /Font << /F0 5 0 R >> >>', content, [
          '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        ]),
      ),
    ).doc;
    return doc.body.flatMap((b) => (b.kind === 'paragraph' ? [b.paragraph] : []));
  };
  const line = (x: number, y: number, text: string): string =>
    `BT /F0 10 Tf 1 0 0 1 ${String(x)} ${String(y)} Tm (${text}) Tj ET`;

  it('reads a list item’s indent and its hanging marker', () => {
    // bug1997343.pdf sets "• They may be unordered bullet lists" ten points in
    // and its nested "1. lists may also be nested" twenty more, and every one
    // of them came back flush left against the column.
    const paras = parasOf(
      [
        line(40, 360, 'A line of body text right across the measure'),
        line(40, 346, 'and it ends here.'),
        line(55, 332, '- an item whose marker hangs to the left'),
        line(65, 318, 'and the item runs on under its own text'),
      ].join('\n'),
    );
    const item = paras.find((p) =>
      p.runs
        .map((r) => r.text)
        .join('')
        .startsWith('-'),
    );
    expect(item).toBeDefined();
    // The BODY of the item is 25pt in; its marker hangs 10pt out of that.
    expect(item?.properties.indentLeft).toBeCloseTo(25, 0);
    expect(item?.properties.indentFirstLine).toBeCloseTo(-10, 0);
    // …and the body it follows is not indented at all.
    expect(paras[0]?.properties.indentLeft ?? 0).toBe(0);
  });

  it('starts a paragraph where a short line is followed by an indented one', () => {
    // The oldest mark in typography. It used to CANCEL the test that ends a
    // paragraph — the two lines "start at different edges" — so bug1997343.pdf
    // read "…figures and mathematics. Apart from two commands at the start…"
    // as one paragraph where the file sets two.
    const paras = parasOf(
      [
        line(40, 360, 'A line of body text right across the measure'),
        line(40, 346, 'and it ends.'),
        line(55, 332, 'Apart from that, a new paragraph opens set in'),
        line(40, 318, 'and runs on to its second line at the measure'),
      ].join('\n'),
    );
    expect(paras).toHaveLength(2);
    expect(paras[1]?.runs.map((r) => r.text).join('')).toContain('Apart from that');
    expect(paras[1]?.properties.indentFirstLine).toBeCloseTo(15, 0);
  });

  it('keeps a list item whose marker line runs the full measure', () => {
    // A FULL line followed by an indented one is an item and its continuation,
    // not two paragraphs. Full across the sheet: a sheet of two lines shows no
    // measure of its own, and is re-set across the rest of the paper.
    const paras = parasOf(
      [
        line(40, 360, 'A line of body text right across the whole of the measure'),
        line(55, 346, 'and its own second line, set in under it'),
      ].join('\n'),
    );
    expect(paras).toHaveLength(1);
  });
});

describe('a word broken across a line comes back together', () => {
  it('joins on the discretionary hyphen and drops it', () => {
    // A line that ends in a hyphen was broken THERE. Read as prose with a
    // space between every line, bug1997343.pdf came back "typical two-column
    // docu ment incorporating tables, figures and mathemat ics".
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf(
          '/MediaBox [0 0 400 400] /Resources << /Font << /F0 5 0 R >> >>',
          [
            'BT /F0 10 Tf 1 0 0 1 40 360 Tm (A line that ends in a docu\\255) Tj ET',
            'BT /F0 10 Tf 1 0 0 1 40 346 Tm (ment and a two\\055) Tj ET',
            'BT /F0 10 Tf 1 0 0 1 40 332 Tm (column word after it) Tj ET',
          ].join('\n'),
          // WinAnsiEncoding, which is what a producer writing 0xAD for a
          // discretionary hyphen means by it: StandardEncoding — the built-in
          // encoding of the face, and its reading with no /Encoding at all —
          // has a single guillemet at that code (Annex D.2).
          [
            '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica ' +
              '/Encoding /WinAnsiEncoding >>',
          ],
        ),
      ),
    ).doc;
    const text = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join('');
    // The soft hyphen goes with the break…
    expect(text).toContain('document and');
    // …and the plain one belongs to the word it ends.
    expect(text).toContain('two-column word');
  });
});

describe('mathematics is set the way the page sets it', () => {
  const spansOf = (content: string) => {
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf('/MediaBox [0 0 300 200] /Resources << /Font << /F0 5 0 R >> >>', content, [
          '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        ]),
      ),
    ).doc;
    return doc.body.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : []));
  };

  it('reads a raised smaller run as a superscript (§17.3.2.42)', () => {
    // A PDF states no such property: an exponent is a smaller face set a little
    // higher. Read flat, bug1997343.pdf's "n^p = n mod p" came back "np", and
    // every prime on the page landed beside its letter instead of over it.
    const runs = spansOf(
      'BT /F0 10 Tf 1 0 0 1 40 100 Tm (n) Tj ET\n' +
        'BT /F0 7 Tf 1 0 0 1 46 103.6 Tm (p) Tj ET\n' +
        'BT /F0 10 Tf 1 0 0 1 52 100 Tm ( = n mod p) Tj ET',
    );
    const raised = runs.find((r) => r.text.trim() === 'p');
    expect(raised?.properties.verticalAlign).toBe('superscript');
    // …at the LINE's size: a document states the nominal size and the layout
    // shrinks a script, so the drawn seven points would come out at five.
    expect(raised?.properties.fontSizePt).toBeCloseTo(10, 1);
    // A run ON the baseline is not a script however small it is.
    expect(runs.find((r) => r.text.includes('mod'))?.properties.verticalAlign).toBe('baseline');
  });

  it('steps between the words of a line that holds no space', () => {
    // TeX's thin space is a sixth of an em and its medium one two ninths, both
    // under the quarter a page that draws its own spaces needs — and a LaTeX
    // document does both: prose with spaces in it, mathematics by stepping.
    // bug1997343.pdf sets "f(x) = sin x + cos x" and we read "sinx+cosx".
    const runs = spansOf(
      'BT /F0 10 Tf 1 0 0 1 40 100 Tm (sin) Tj ET\n' +
        'BT /F0 10 Tf 1 0 0 1 54.7 100 Tm (x) Tj ET\n' +
        'BT /F0 10 Tf 1 0 0 1 62 100 Tm (+) Tj ET\n' +
        'BT /F0 10 Tf 1 0 0 1 71 100 Tm (cos) Tj ET',
    );
    expect(runs.map((r) => r.text).join('')).toBe('sin x + cos');
  });

  it('leaves a page that writes its own spaces alone', () => {
    // The same gaps inside a line that HAS a space in it are kerning, not
    // words: a producer that splits a word for kerning leaves eight hundredths
    // of an em between the halves.
    const runs = spansOf(
      'BT /F0 10 Tf 1 0 0 1 40 100 Tm (Con) Tj ET\n' +
        'BT /F0 10 Tf 1 0 0 1 56.5 100 Tm (tents ) Tj ET\n' +
        'BT /F0 10 Tf 1 0 0 1 81.5 100 Tm (here) Tj ET',
    );
    expect(runs.map((r) => r.text).join('')).toBe('Contents here');
  });
});

describe('a matrix is a matrix, not three lines (§22.1.2.68)', () => {
  /** A page of prose with a 2×2 matrix set in brackets in the middle of it. */
  const withMatrix = (brackets = true): Uint8Array => {
    const ops: Array<string> = [];
    for (let i = 0; i < 5; i++)
      ops.push(`BT /F0 10 Tf 1 0 0 1 40 ${String(370 - i * 14)} Tm (a line of prose here) Tj ET`);
    // The numbers stand on two baselines six points apart, the brackets on the
    // baseline between them — which is where a stretched bracket sits.
    ops.push('BT /F0 10 Tf 1 0 0 1 60 300 Tm (1) Tj ET');
    ops.push('BT /F0 10 Tf 1 0 0 1 80 300 Tm (2) Tj ET');
    if (brackets) {
      ops.push('BT /F0 10 Tf 1 0 0 1 50 294 Tm (\\() Tj ET');
      ops.push('BT /F0 10 Tf 1 0 0 1 95 294 Tm (\\)) Tj ET');
    }
    ops.push('BT /F0 10 Tf 1 0 0 1 60 288 Tm (3) Tj ET');
    ops.push('BT /F0 10 Tf 1 0 0 1 80 288 Tm (4) Tj ET');
    for (let i = 0; i < 5; i++)
      ops.push(
        `BT /F0 10 Tf 1 0 0 1 40 ${String(260 - i * 14)} Tm (and more prose after it) Tj ET`,
      );
    return onePagePdf(
      '/MediaBox [0 0 400 400] /Resources << /Font << /F0 5 0 R >> >>',
      ops.join('\n'),
      ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
    );
  };

  it('reads the rows and the brackets as the object they are', () => {
    // A PDF states no mathematics: a matrix reaches the page as numbers on two
    // baselines with stretched brackets drawn between them. Read line by line,
    // bug1997343.pdf's product of three matrices came back as three lines of
    // prose — "1 2 1 1 1 3", "( )( ) = ( )", "3 4 0 1 3 7".
    const doc = reconstructByLayout(PdfFile.parse(withMatrix())).doc;
    const math = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : []))
      .find((r) => r.math !== undefined)?.math;
    expect(math).toEqual({
      type: 'row',
      children: [
        {
          type: 'delimiter',
          begChr: '(',
          endChr: ')',
          children: [
            {
              type: 'matrix',
              rows: [
                [
                  { type: 'run', text: '1' },
                  { type: 'run', text: '2' },
                ],
                [
                  { type: 'run', text: '3' },
                  { type: 'run', text: '4' },
                ],
              ],
            },
          ],
        },
      ],
    });
    // …and its numbers are not read a second time as prose.
    const text = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join(' ');
    expect(text).toContain('a line of prose here');
    expect(text).not.toMatch(/[1234]/u);
  });

  it('leaves close-set lines that are NOT a matrix as the prose they are', () => {
    // The brackets are what say "matrix"; without them these are three short
    // lines, and read as a matrix a table of figures would lose its columns.
    const doc = reconstructByLayout(PdfFile.parse(withMatrix(false))).doc;
    const runs = doc.body.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : []));
    expect(runs.some((r) => r.math !== undefined)).toBe(false);
    expect(runs.map((r) => r.text).join(' ')).toMatch(/1|2|3|4/u);
  });
});

describe('a running foot is a foot, not a paragraph (§17.6.13)', () => {
  /**
   * Three pages of body with the same line standing alone at the bottom, and —
   * with `signed` — a publisher's line between it and the text.
   */
  const paper = (signed = false, foot?: string): Uint8Array => {
    const page = (n: number): string => {
      const ops: Array<string> = [];
      for (let i = 0; i < 8; i++)
        ops.push(
          `BT /F0 10 Tf 1 0 0 1 40 ${String(360 - i * 14)} Tm (body line ${String(i)}) Tj ET`,
        );
      // Alone at the foot, a long way below the text block. `foot` replaces the
      // line that carries the page's own number with one that repeats.
      if (signed) ops.push('BT /F0 8 Tf 1 0 0 1 40 40 Tm (Thing Press) Tj ET');
      const last = foot ?? `The Journal of Things ${String(n)}`;
      ops.push(`BT /F0 8 Tf 1 0 0 1 40 20 Tm (${last} 2000) Tj ET`);
      return ops.join('\n');
    };
    return pages([page(1), page(2), page(3)]);
  };

  /** Those three pages assembled into a file. */
  const pages = (contents: ReadonlyArray<string>): Uint8Array => {
    const kids = contents.map((_, i) => `${String(3 + i * 2)} 0 R`).join(' ');
    const objects: Array<string> = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      `<< /Type /Pages /Kids [${kids}] /Count ${String(contents.length)} >>`,
    ];
    const fontAt = 3 + contents.length * 2;
    contents.forEach((content, i) => {
      objects.push(
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Contents ${String(4 + i * 2)} 0 R ` +
          `/Resources << /Font << /F0 ${String(fontAt)} 0 R >> >> >>`,
        `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
      );
    });
    objects.push('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    let pdf = '%PDF-1.7\n';
    const offsets: Array<number> = [];
    objects.forEach((body, i) => {
      offsets.push(pdf.length);
      pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
    });
    const xref = pdf.length;
    pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
    for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
    pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
    return new TextEncoder().encode(pdf);
  };

  it('lifts the repeated line into the section’s footer', () => {
    // Read as body it goes wherever the reflow puts it: bug1997343.pdf's page
    // number came out on a sheet of its own between the two the paper has, and
    // TAMReview.pdf's "Sprouts — http://…" in the middle of the abstract.
    const doc = reconstructByLayout(PdfFile.parse(paper())).doc;
    const body = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join(' ');
    expect(body).toContain('body line 0');
    expect(body).not.toContain('The Journal of Things');
    // …and the band itself, with the number in it made a field.
    const part = doc.section?.footers[0]?.relationshipId;
    expect(part).toBeDefined();
    const band = part !== undefined ? doc.headersFooters?.get(part) : undefined;
    const runs = band?.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : [])) ?? [];
    expect(runs.map((r) => r.text).join('')).toContain('The Journal of Things');
    expect(runs.some((r) => r.field === 'PAGE')).toBe(true);
  });

  it('lifts a foot of TWO lines, in the order the page shows them', () => {
    // A foot need not be one line. ZapfDingbats.pdf signs each sheet twice —
    // the publisher's line, and the build stamp thirty points under it — and
    // taking only the bottom line left the other in the body, where, after a
    // table that fills the sheet, it had nowhere to go but a page of its own:
    // a two-page document came out as four.
    const doc = reconstructByLayout(PdfFile.parse(paper(true))).doc;
    const body = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join(' ');
    expect(body).not.toContain('Thing Press');
    const part = doc.section?.footers[0]?.relationshipId;
    const band = part !== undefined ? doc.headersFooters?.get(part) : undefined;
    const lines =
      band?.map((b) =>
        b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text).join('') : '',
      ) ?? [];
    expect(lines).toHaveLength(2);
    expect(lines[0]).toContain('Thing Press');
    expect(lines[1]).toContain('The Journal of Things');
  });

  it('sets a foot written in REGIONS on the stops it was written at', () => {
    // A foot is written the way a spreadsheet's is: something at the left,
    // something against the far edge. ZapfDingbats.pdf signs each sheet
    // "© RenderX 2000" at the left and "XSL Formatting Objects Test Suite" at
    // the right, and the two hundred points between them came back as one word
    // space — the two crowding each other at the left.
    const page = (n: number): string =>
      [
        ...Array.from(
          { length: 8 },
          (_, i) =>
            `BT /F0 10 Tf 1 0 0 1 40 ${String(360 - i * 14)} Tm (body line ${String(i)}) Tj ET`,
        ),
        `BT /F0 8 Tf 1 0 0 1 40 20 Tm (Thing Press ${String(n)}) Tj ET`,
        'BT /F0 8 Tf 1 0 0 1 200 20 Tm (The Journal of Things) Tj ET',
      ].join('\n');
    const doc = reconstructByLayout(PdfFile.parse(pages([page(1), page(2), page(3)]))).doc;
    const part = doc.section?.footers[0]?.relationshipId;
    const band = part !== undefined ? doc.headersFooters?.get(part) : undefined;
    const line = band?.[0];
    if (line?.kind !== 'paragraph') throw new Error('the band has a line');
    expect(line.paragraph.runs.map((r) => r.text).join('')).toContain('\t');
    expect(line.paragraph.properties.tabs?.[0]).toMatchObject({ relativeTo: 'right' });
  });

  it('sets a one-line foot against the far edge, where the page set it', () => {
    // "Page 1 of 2" stands at the right margin of every page of a receipt. A
    // paragraph is read as set to the right only when its left edge is ragged,
    // which one line's is not, and the foot came back at the left margin.
    const page = (n: number): string =>
      [
        ...Array.from(
          { length: 8 },
          (_, i) =>
            `BT /F0 10 Tf 1 0 0 1 40 ${String(360 - i * 14)} Tm (body line ${String(i)}) Tj ET`,
        ),
        `BT /F0 8 Tf 1 0 0 1 220 20 Tm (Page ${String(n)} of 3) Tj ET`,
      ].join('\n');
    const doc = reconstructByLayout(PdfFile.parse(pages([page(1), page(2), page(3)]))).doc;
    const part = doc.section?.footers[0]?.relationshipId;
    const band = part !== undefined ? doc.headersFooters?.get(part) : undefined;
    const line = band?.[0];
    if (line?.kind !== 'paragraph') throw new Error('the band has a line');
    expect(line.paragraph.properties.alignment).toBe('right');
  });

  it('leaves the number alone where the foot says the SAME thing on every page', () => {
    // A page number is a number that CHANGES from page to page. ZapfDingbats.pdf
    // signs each sheet "© RenderX 2000", and read as a page number the year came
    // out as "© RenderX 1".
    const doc = reconstructByLayout(PdfFile.parse(paper(true))).doc;
    const part = doc.section?.footers[0]?.relationshipId;
    const band = part !== undefined ? doc.headersFooters?.get(part) : undefined;
    const runs = band?.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : [])) ?? [];
    // "Thing Press" repeats, "The Journal of Things 1|2|3" does not — so the
    // page's own number is still a field.
    expect(runs.some((r) => r.field === 'PAGE')).toBe(true);
    const same = paper(true, 'Thing Press');
    const only = reconstructByLayout(PdfFile.parse(same)).doc;
    const id = only.section?.footers[0]?.relationshipId;
    const kept = (id !== undefined ? only.headersFooters?.get(id) : undefined)?.flatMap((b) =>
      b.kind === 'paragraph' ? b.paragraph.runs : [],
    );
    expect(kept?.map((r) => r.text).join('')).toContain('2000');
    expect(kept?.some((r) => r.field === 'PAGE')).toBe(false);
  });

  /** A sheet of eight long lines, and two that open a chapter with its heading alone. */
  const chapters = (footed = true): Uint8Array => {
    const foot = (n: number): string =>
      footed ? `BT /F0 10 Tf 1 0 0 1 200 20 Tm (page ${String(n)} / 3) Tj ET` : '';
    const body = Array.from(
      { length: 8 },
      (_, i) =>
        `BT /F0 10 Tf 1 0 0 1 40 ${String(360 - i * 14)} Tm (body line ${String(i)} runs on across the whole of the sheet) Tj ET`,
    ).join('\n');
    const heading = 'BT /F0 16 Tf 1 0 0 1 40 360 Tm (Chapter) Tj ET';
    return pages([`${body}\n${foot(1)}`, `${heading}\n${foot(2)}`, `${heading}\n${foot(3)}`]);
  };

  it('finds the foot under a sheet of one line', () => {
    // A chapter opens on a sheet of its own: its heading, and the page number
    // under it. Asked for more lines than that, basicapi.pdf's "page 2 / 3"
    // stayed in the body and came back half way up the sheet.
    const doc = reconstructByLayout(PdfFile.parse(chapters())).doc;
    const body = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join(' ');
    expect(body).toContain('Chapter');
    expect(body).not.toContain('page');
    expect(doc.section?.footers).toHaveLength(1);
  });

  it('keeps the bottom margin down to the foot, and the foot where the page set it', () => {
    // The foot stands IN the bottom margin, so the text block may reach down
    // to the white above it however early the sheets at hand end. Measured to
    // where the three sheets' text stops, the margin came out at half the paper.
    const margins = reconstructByLayout(PdfFile.parse(chapters())).doc.section?.margins;
    // The foot's box starts three tenths of its size under the baseline at 20.
    expect(margins?.footer).toBeCloseTo(17, 0);
    // …and the text may come down to two ems above its top (20 + 9.6 + 20).
    expect(margins?.bottom).toBeLessThan(50);
  });

  it('takes the right margin from the sheets full enough to reach it', () => {
    // A sheet of one heading says where the heading ends. Two of them voted
    // the right margin in to a third of the paper, and the first sheet's long
    // lines could not be set in what was left.
    const margins = reconstructByLayout(PdfFile.parse(chapters(false))).doc.section?.margins;
    expect(margins?.right).toBeLessThan(60);
  });

  it('leaves a last paragraph where the page put it', () => {
    // One page proves nothing, and a page whose last line is a line's gap from
    // the one above it is a paragraph, not a foot.
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf(
          '/MediaBox [0 0 300 400] /Resources << /Font << /F0 5 0 R >> >>',
          Array.from(
            { length: 9 },
            (_, i) =>
              `BT /F0 10 Tf 1 0 0 1 40 ${String(360 - i * 14)} Tm (line ${String(i)}) Tj ET`,
          ).join('\n'),
          ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
        ),
      ),
    ).doc;
    expect(doc.section?.footers ?? []).toHaveLength(0);
    const body = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join(' ');
    expect(body).toContain('line 8');
  });
});

describe('the measure a narrow column shows (§17.6.11)', () => {
  /** One page, 300 wide, in a face whose every glyph is half an em. */
  const page = (lines: ReadonlyArray<string>): Uint8Array => {
    const widths = Array.from({ length: 91 }, () => 500).join(' ');
    const shown = lines
      .map((text, i) => `BT /F0 10 Tf 20 ${String(360 - i * 12)} Td (${text}) Tj ET`)
      .join('\n');
    return onePagePdf('/MediaBox [0 0 300 400] /Resources << /Font << /F0 5 0 R >> >>', shown, [
      '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FirstChar 32 /LastChar 122 ' +
        `/Widths [${widths}] >>`,
    ]);
  };
  const right = (lines: ReadonlyArray<string>): number | undefined =>
    reconstructByLayout(PdfFile.parse(page(lines))).doc.section?.margins?.right;

  it('takes a column’s measure where its lines run to one edge', () => {
    // bug1057544.pdf sets a paragraph in a column a quarter of the sheet wide;
    // held to a third of the sheet, it came back in two lines where the page
    // has four. Three lines here run to x = 120, a hundred points in.
    const column = ['aaaaaaaaaaaaaaaaaaaa', 'bbbbbbbbbbbbbbbbbbbb', 'cccccccccccccccccccc', 'ddd'];
    expect(right(column)).toBeGreaterThan(150);
  });

  it('keeps to a third of the sheet where one line alone reaches that far', () => {
    // issue10529.pdf's one long line, measured to, wrapped in the wider face
    // it is re-set in: a line no other runs to is no measure.
    const loose = ['aaaaaaaaaaaaaaaaaaaa', 'bbb', 'ccc', 'ddd'];
    expect(right(loose)).toBeLessThanOrEqual(100);
  });
});

describe('the last text of a sheet, set far below the rest (§17.3.1.33)', () => {
  const sheet = (lines: ReadonlyArray<[number, string]>): ReturnType<typeof reconstructByLayout> =>
    reconstructByLayout(
      PdfFile.parse(
        onePagePdf(
          '/MediaBox [0 0 300 400] /Resources << /Font << /F0 5 0 R >> >>',
          lines
            .map(([y, text]) => `BT /F0 10 Tf 1 0 0 1 40 ${String(y)} Tm (${text}) Tj ET`)
            .join('\n'),
          ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
        ),
      ),
    );
  const words = (els: ReadonlyArray<BodyElement>): string =>
    els
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join(' ');

  it('stands where the page set it, not a third of the sheet below the rest', () => {
    // bug1989304.pdf signs its sheet "World" at the foot; held to the third of
    // the sheet a paragraph's spacing may say, it came back half way up.
    const { doc } = sheet([
      [360, 'first line'],
      [346, 'second line'],
      [20, 'signed at the foot'],
    ]);
    expect(words(doc.body)).not.toContain('signed');
    const placed = doc.body.find(
      (b) =>
        b.kind === 'shape' && words(b.shape.text?.content ?? []).includes('signed at the foot'),
    );
    if (placed?.kind !== 'shape') throw new Error('the last line is placed');
    // Its box stands on the baseline at 20, a quarter of the size below it.
    expect(placed.shape.float?.posV?.offsetPt).toBeCloseTo(400 - 17.5 - 12.5, 0);
  });

  it('keeps text that has more after it in the flow', () => {
    const { doc } = sheet([
      [360, 'first line'],
      [100, 'far below'],
      [86, 'and more after it'],
    ]);
    expect(words(doc.body)).toContain('far below');
  });
});

describe('a line the page set in one piece (§17.3.1.12)', () => {
  /** A page 300 wide, in a face whose every glyph is half an em. */
  const laidOut = (content: string): Array<{ text: string; right?: number }> => {
    const widths = Array.from({ length: 91 }, () => 500).join(' ');
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf('/MediaBox [0 0 300 400] /Resources << /Font << /F0 5 0 R >> >>', content, [
          '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FirstChar 32 /LastChar 122 ' +
            `/Widths [${widths}] >>`,
        ]),
      ),
    ).doc;
    return doc.body.flatMap((b) =>
      b.kind === 'paragraph'
        ? [
            {
              text: b.paragraph.runs.map((r) => r.text).join(''),
              ...(b.paragraph.properties.indentRight !== undefined
                ? { right: b.paragraph.properties.indentRight }
                : {}),
            },
          ]
        : [],
    );
  };

  it('may run on into the margin rather than wrap when re-set', () => {
    // bug1108301.pdf's one line ran nearly to the margin; re-set in a wider
    // face its last word wrapped, and on a sheet fifty points tall it fell off
    // the paper.
    const [line] = laidOut(
      'BT /F0 10 Tf 20 360 Td (aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa) Tj ET',
    );
    expect(line?.right).toBeLessThan(0);
    // …a sixth of its measure at most, and never past the edge of the sheet.
    expect(line?.right).toBeGreaterThanOrEqual(-200 / 6 - 0.01);
  });

  it('keeps a paragraph the page wrapped, and a short line, to the measure', () => {
    const content = [
      'BT /F0 10 Tf 20 360 Td (aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa) Tj ET',
      'BT /F0 10 Tf 20 348 Td (bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb) Tj ET',
      'BT /F0 10 Tf 20 336 Td (cccc) Tj ET',
      'BT /F0 10 Tf 20 300 Td (dddd) Tj ET',
    ].join('\n');
    for (const para of laidOut(content)) expect(para.right ?? 0).toBeGreaterThanOrEqual(0);
  });
});

describe('the spaces a line has are the ones the page shows, once each', () => {
  /**
   * The text of a page set in Helvetica, whose space is `space` thousandths of
   * an em and every other glyph half an em.
   */
  const line = (content: string, space = 0): string => {
    const widths = Array.from({ length: 91 }, (_, i) => (i === 0 ? space : 500)).join(' ');
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf(
          '/MediaBox [0 0 300 100] /Resources << /Font << /F0 5 0 R >> >>',
          `BT /F0 20 Tf 10 60 Td ${content} ET`,
          [
            '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /FirstChar 32 /LastChar 122 ' +
              `/Widths [${widths}] >>`,
          ],
        ),
      ),
    ).doc;
    return doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join('');
  };

  it('is no word space where the ink on either side of it closes up', () => {
    // bug1046314.pdf maps a Thai mark of no width to U+0020, and "(คำแปล)" —
    // one word — came back "(คำ แปล)".
    expect(line('(ab) Tj ( ) Tj (cd) Tj')).toBe('abcd');
  });

  it('is a word space where it stands in a gap the page steps across', () => {
    expect(line('[(ab) ( ) -200 (cd)] TJ')).toBe('ab cd');
  });

  it('joins a line the page ended with a space without a second one', () => {
    // Twenty glyphs of a full first line, and a word too long for the room
    // left on it: one paragraph, broken where the page broke it.
    expect(line('(aaaa bbbb cccc dddd ) Tj 0 -24 Td (eeeeeeee) Tj', 250)).toBe(
      'aaaa bbbb cccc dddd eeeeeeee',
    );
  });
});

describe('type too small to read is a mark on the sheet, not a line of it', () => {
  // TCPDF signs the last page of everything it makes in one-point type, three
  // points from the corner of the paper.
  const signed = (): Uint8Array =>
    onePagePdf(
      '/MediaBox [0 0 300 400] /Resources << /Font << /F0 5 0 R >> >>',
      [
        ...Array.from(
          { length: 8 },
          (_, i) => `BT /F0 10 Tf 1 0 0 1 40 ${String(360 - i * 14)} Tm (line ${String(i)}) Tj ET`,
        ),
        'BT /F0 1 Tf 1 0 0 1 3 1 Tm (Powered by TCPDF) Tj ET',
      ].join('\n'),
      ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
    );

  it('measures the margins without it', () => {
    // Taken for text it was the leftmost and lowest thing on the page, and
    // every line of basicapi.pdf was set against the edge of the paper.
    const margins = reconstructByLayout(PdfFile.parse(signed())).doc.section?.margins;
    expect(margins?.left).toBeCloseTo(40, 0);
    expect(margins?.bottom).toBeGreaterThan(100);
  });

  it('keeps it where the page put it, out of the flow', () => {
    const doc = reconstructByLayout(PdfFile.parse(signed())).doc;
    const text = (els: ReadonlyArray<BodyElement>): string =>
      els
        .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
        .join('');
    expect(text(doc.body)).not.toContain('Powered');
    const stamp = doc.body.find(
      (b) => b.kind === 'shape' && text(b.shape.text?.content ?? []).includes('Powered by TCPDF'),
    );
    if (stamp?.kind !== 'shape') throw new Error('the stamp is placed');
    expect(stamp.shape.float?.posH?.offsetPt).toBeCloseTo(3, 0);
  });
});

describe('a crop box cuts the line it crosses (§14.11.2)', () => {
  it('keeps the letters the page shows and drops the rest', () => {
    // endchar.pdf is one line of a poster — "LE HOLD-UP PLANÉTAIRE" — cropped
    // to the fourteen points that hold its É, which is all any viewer shows.
    // A run that reached into the shown page was kept whole, so the line was
    // re-set into a column fourteen points wide: four pages of one letter.
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf(
          '/MediaBox [0 0 300 300] /CropBox [200 90 260 120] ' +
            '/Resources << /Font << /F0 5 0 R >> >>',
          'BT /F0 12 Tf 1 0 0 1 20 100 Tm (ABCDEFGHIJKLMNOPQRSTUVWXYZ) Tj ET',
          ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
        ),
      ),
    ).doc;
    const text = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join('');
    // Helvetica sets those capitals about 8.4pt apart from x=20, so the line
    // ends around 240 and the crop's 200..260 holds its last few letters. Which
    // few is an estimate — the run states its width, not its every letter — and
    // the answer is a letter either way.
    expect(text.length).toBeGreaterThan(2);
    expect(text.length).toBeLessThan(12);
    expect('ABCDEFGHIJKLMNOPQRSTUVWXYZ'.endsWith(text)).toBe(true);
  });

  it('leaves a line the crop does not reach', () => {
    const doc = reconstructByLayout(
      PdfFile.parse(
        onePagePdf(
          '/MediaBox [0 0 300 300] /CropBox [0 0 300 300] ' +
            '/Resources << /Font << /F0 5 0 R >> >>',
          'BT /F0 12 Tf 1 0 0 1 20 100 Tm (ABCDEFGHIJ) Tj ET',
          ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
        ),
      ),
    ).doc;
    const text = doc.body
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join('');
    expect(text).toBe('ABCDEFGHIJ');
  });
});

describe('a leader is not a row of spaced dots (§17.3.1.25)', () => {
  it('joins the dots and ends the entry at the line', () => {
    // A leader is drawn one character at a time with a step about as wide as a
    // word space, so every threshold that tells a space from a kern says
    // "space" between every dot. bug886717.pdf's contents came back as
    // "Abstract . . . . . . . . 3", four times as long as the page sets it, and
    // its forty entries reflowed into one paragraph across two pages.
    const line = (y: number, word: string, page: string): string => {
      const ops = [`BT /F0 12 Tf 1 0 0 1 40 ${String(y)} Tm (${word}) Tj ET`];
      for (let i = 0; i < 30; i++) {
        ops.push(`BT /F0 12 Tf 1 0 0 1 ${String(100 + i * 5.3)} ${String(y)} Tm (.) Tj ET`);
      }
      ops.push(`BT /F0 12 Tf 1 0 0 1 262 ${String(y)} Tm (${page}) Tj ET`);
      return ops.join('\n');
    };
    const doc = Ream.parse(
      onePagePdf(
        '/MediaBox [0 0 300 200] /Resources << /Font << /F0 5 0 R >> >>',
        `${line(150, 'Abstract', '3')}\n${line(135, 'Foreword', '5')}`,
        ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
      ),
    ).flow;
    const texts = doc.body.flatMap((b) =>
      b.kind === 'paragraph' ? [b.paragraph.runs.map((r) => r.text).join('')] : [],
    );
    // Two entries, two paragraphs — not one paragraph of both.
    expect(texts).toHaveLength(2);
    expect(texts[0]).toMatch(/^Abstract \.{30} 3$/u);
    expect(texts[1]).toMatch(/^Foreword \.{30} 5$/u);
  });
});

/** Two pages, 200×100 then 600×400, each with one word on it. */
function twoSizePdf(): Uint8Array {
  const content = 'BT /F0 12 Tf 20 40 Td (Word) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R ' +
      '/Resources << /Font << /F0 6 0 R >> >> >>',
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 400] /Contents 4 0 R ' +
      '/Resources << /Font << /F0 6 0 R >> >> >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

/** A one-page PDF of hand-written objects; `page` is the page dict's body. */
function onePagePdf(page: string, content: string, extra: ReadonlyArray<string> = []): Uint8Array {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /Contents 4 0 R ${page} >>`,
    `<< /Length ${String(content.length)} >>\nstream\n${content}\nendstream`,
    ...extra,
  ];
  let pdf = '%PDF-1.7\n';
  const offsets: Array<number> = [];
  objects.forEach((body, i) => {
    offsets.push(pdf.length);
    pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}

/** Two words on one baseline, three hundred points apart — a table row. */
const twoColumnLinePdf = (): Uint8Array =>
  onePagePdf('/MediaBox [0 0 400 800]', 'BT /F1 10 Tf 20 700 Td (Left) Tj 300 0 Td (Right) Tj ET');

/** One word stamped twice, the second set down over the first's second half. */
const stampedWordPdf = (): Uint8Array =>
  onePagePdf('/MediaBox [0 0 400 800]', 'BT /F1 20 Tf 20 700 Td (Word) Tj 15 0 Td (Word) Tj ET');

/** A word and a footnote mark set a size smaller and three quarters of an em up. */
const superscriptPdf = (): Uint8Array =>
  onePagePdf(
    '/MediaBox [0 0 400 800]',
    'BT /F1 10 Tf 20 700 Td (Word) Tj /F1 6 Tf 45 7.5 Td (1\\)) Tj ET',
  );

/**
 * A landscape sheet drawn sideways in a portrait box and stood up by `/Rotate`
 * — the shape all twenty-five pages of Brotli-Prototype-FileA.pdf take. The
 * text matrix turns the words a quarter the other way, so the page's own turn
 * is what sets them level.
 */
const turnedPagePdf = (rotate: number): Uint8Array =>
  onePagePdf(
    `/MediaBox [0 0 400 800] /Rotate ${String(rotate)}`,
    'BT /F1 10 Tf 0 -1 1 0 100 600 Tm (Side) Tj ET',
  );

/** A page whose `/MediaBox` starts twelve points off the origin, as 160F's does. */
const offsetBoxPdf = (): Uint8Array =>
  onePagePdf(
    '/MediaBox [-12 12 388 812]',
    'BT /F1 10 Tf 20 700 Td (Hello) Tj ET 0 0 1 rg 20 100 50 30 re f',
  );

/**
 * A flat line and, a quarter turn from it, a label set on its side — the shape
 * 160F-2019.pdf's "Nature" takes down the middle of a column. The two share a
 * baseline y within a line's height, which is exactly the trap.
 */
const turnedLabelPdf = (): Uint8Array =>
  onePagePdf(
    '/MediaBox [0 0 400 800]',
    'BT /F1 10 Tf 20 700 Td (Flat) Tj ET BT /F1 10 Tf 0 1 -1 0 200 698 Tm (Side) Tj ET',
  );

describe('placed reconstruction (E-PDF EP4)', () => {
  it('anchors every line where its glyphs stand, instead of flowing them', async () => {
    // A form is a grid of ruled boxes with a label in each: flowed, the labels
    // land an inch from the boxes they label, because the artwork is placed and
    // the words are not. 160F-2019.pdf is that document.
    const docx = buildDocxFromBody(
      '<w:p><w:r><w:t>FirstLine</w:t></w:r></w:p><w:p><w:r><w:t>SecondLine</w:t></w:r></w:p>',
    );
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const placed = reconstructByLayout(PdfFile.parse(pdf), 'positional');

    const shapes = placed.doc.body.filter((b) => b.kind === 'shape').map((b) => b.shape);
    expect(shapes.length).toBeGreaterThanOrEqual(2);
    // Each carries its words and an anchor of its own on the page.
    for (const shape of shapes) {
      expect(shape.float?.posV?.relativeFrom).toBe('page');
      expect(shape.text?.content.length).toBeGreaterThan(0);
    }
    // The first line sits higher on the page, so its offset from the top is less.
    const offsets = shapes.map((shape) => shape.float?.posV?.offsetPt ?? 0);
    expect(offsets[0]!).toBeLessThan(offsets[1]!);
  });

  it('splits a baseline where a column gap opens, so each piece keeps its x', () => {
    // 160F-2019.pdf sets a line number, a label and a right-hand column on one
    // baseline. Read as a single line, the right-hand text was dragged left
    // against its neighbour with one space between: "sous-total: n° dossier:".
    const placed = reconstructByLayout(PdfFile.parse(twoColumnLinePdf()), 'positional');
    const shapes = placed.doc.body.filter((b) => b.kind === 'shape').map((b) => b.shape);
    expect(shapes).toHaveLength(2);
    const offsets = shapes.map((s) => s.float?.posH?.offsetPt ?? 0).sort((a, b) => a - b);
    expect(offsets[0]).toBeCloseTo(20, 0);
    expect(offsets[1]).toBeCloseTo(320, 0);
    // Both stand on the same baseline, so neither moved vertically.
    const tops = shapes.map((s) => s.float?.posV?.offsetPt ?? 0);
    expect(tops[0]).toBeCloseTo(tops[1]!, 5);
  });

  it('splits a baseline where two runs OVERLAP, which is a placement too', () => {
    // ContentStream*Type3.pdf stamps one word three times at half its own
    // width. Read as a single line the three were flowed end to end, and the
    // line came out half again as wide as the page sets it — the same error as
    // the column gap, in the other direction.
    const placed = reconstructByLayout(PdfFile.parse(stampedWordPdf()), 'positional');
    const shapes = placed.doc.body.filter((b) => b.kind === 'shape').map((b) => b.shape);
    expect(shapes).toHaveLength(2);
    const offsets = shapes.map((s) => s.float?.posH?.offsetPt ?? 0).sort((a, b) => a - b);
    // Each stands where it was stamped: the second 15pt on, under a 40pt word.
    expect(offsets[0]).toBeCloseTo(20, 0);
    expect(offsets[1]).toBeCloseTo(35, 0);
  });

  it('keeps a mark set above the line above it, at its own size', () => {
    // 160F-2019.pdf sets its footnote marks a size smaller and three quarters
    // of an em up. Read as one line they came down flat onto the words.
    const placed = reconstructByLayout(PdfFile.parse(superscriptPdf()), 'positional');
    const shapes = placed.doc.body.filter((b) => b.kind === 'shape').map((b) => b.shape);
    expect(shapes).toHaveLength(2);
    const words = (s: (typeof shapes)[number]): string => {
      const first = s.text?.content[0];
      return first?.kind === 'paragraph' ? first.paragraph.runs.map((r) => r.text).join('') : '';
    };
    const word = shapes.find((s) => words(s) === 'Word');
    const mark = shapes.find((s) => words(s) === '1)');
    expect(word).toBeDefined();
    expect(mark).toBeDefined();
    // The mark stands higher on the page, so its offset from the top is less.
    const top = (s: (typeof shapes)[number]): number => s.float?.posV?.offsetPt ?? 0;
    expect(top(mark!)).toBeLessThan(top(word!));
    // And it is a six-point mark, not a ten-point one.
    expect(mark!.height).toBeLessThan(word!.height);
  });

  it('reads one flowing line across the same gap', () => {
    // A paragraph is meant to be read across: only the placed reading splits.
    // Across the gap stands a TAB, not a space — the page SET the second piece
    // out there, and a space closes the two up (§17.3.1.38: the stop that keeps
    // it out is written with the paragraph).
    const flowed = reconstructByLayout(PdfFile.parse(twoColumnLinePdf()));
    const paras = paragraphs(flowed.doc);
    expect(paras).toHaveLength(1);
    expect(paras[0]!.paragraph.runs.map((r) => r.text).join('')).toBe('Left\tRight');
    expect((paras[0]!.paragraph.properties.tabs ?? []).length).toBe(1);
  });

  it('stands a turned page up, and its words with it (§14.11.1)', () => {
    // Brotli-Prototype-FileA.pdf is twenty-five landscape sheets drawn sideways
    // in portrait boxes with /Rotate 270. Read as the box says, every one came
    // back portrait with its words running down the page.
    const placed = reconstructByLayout(PdfFile.parse(turnedPagePdf(270)), 'positional');
    expect(placed.doc.section?.pageSize?.width).toBeCloseTo(800, 5);
    expect(placed.doc.section?.pageSize?.height).toBeCloseTo(400, 5);
    expect(placed.doc.section?.pageSize?.orientation).toBe('landscape');
    const shape = placed.doc.body.find((b) => b.kind === 'shape');
    expect(shape?.kind).toBe('shape');
    if (shape?.kind !== 'shape') return;
    // The matrix turns the words a quarter one way and the page the other, so
    // what is left is level type.
    expect(shape.shape.transform?.rotation60k).toBeUndefined();
    // (100, 600) on a 400×800 box, turned 270°: x = 800 − 600, y = 100.
    expect(shape.shape.float?.posH?.offsetPt).toBeCloseTo(200, 5);
  });

  it('turns the words with a page whose turn leaves them running down it (§17.6.20)', () => {
    // hello_world_rotated.pdf sets its words upright in a portrait box and
    // turns the page by /Rotate 90: every viewer shows them running down a
    // landscape sheet. Read on the sheet, they came back flat across it.
    const pdf = onePagePdf(
      '/MediaBox [0 0 400 800] /Rotate 90',
      'BT /F1 20 Tf 100 600 Td (Hello world) Tj ET 0 0 1 rg 100 300 50 30 re f',
    );
    const flowed = reconstructByLayout(PdfFile.parse(pdf));
    const section = flowed.doc.section;
    expect(section?.textDirection).toBe('tbRl');
    expect([section?.pageSize?.width, section?.pageSize?.height]).toEqual([800, 400]);
    expect(section?.pageSize?.orientation).toBe('landscape');
    // Read in the frame where the words stand upright, they are one line, and
    // the line starts where the page starts it: at x=100 of the box, which is
    // the sheet's top margin once the section runs down it.
    const paras = paragraphs(flowed.doc);
    expect(paras.map((p) => p.paragraph.runs.map((r) => r.text).join(''))).toContain('Hello world');
    expect(section?.margins?.top).toBeCloseTo(100, 0);
    // The box drawn on the page stands on the SHEET, turned with the page, as
    // Word stands an anchored drawing: (100, 300) 50×30 on the box is x 300–330,
    // y 100–150 on the sheet, and a 50×30 box turned a quarter about its
    // centre covers exactly that.
    const shape = flowed.doc.body.find((b) => b.kind === 'shape');
    if (shape?.kind !== 'shape') throw new Error('expected the box');
    expect(shape.shape.transform?.rotation60k).toBe(90 * 60000);
    expect(shape.shape.float?.posH?.offsetPt).toBeCloseTo(290, 3);
    expect(shape.shape.float?.posV?.offsetPt).toBeCloseTo(110, 3);
  });

  it('gives a page whose media box bounds nothing the size a page has when it states none', () => {
    // boundingBox_invalid.pdf's first page is `/MediaBox [0 0 0 0]`. Every
    // viewer shows it as a Letter sheet with its words on it; taken at its
    // word, the sheet had no size and every word stood outside it.
    const pdf = onePagePdf('/MediaBox [0 0 0 0]', 'BT /F1 20 Tf 72 700 Td (Empty) Tj ET');
    const flowed = reconstructByLayout(PdfFile.parse(pdf));
    expect(
      paragraphs(flowed.doc).map((p) => p.paragraph.runs.map((r) => r.text).join('')),
    ).toContain('Empty');
    expect([flowed.doc.section?.pageSize?.width, flowed.doc.section?.pageSize?.height]).toEqual([
      612, 792,
    ]);
  });

  it('places what stands off the sheet as a mark, and measures nothing by it', () => {
    // freeculture.pdf sets a printer's bar of ZapfDingbats nine points under
    // the crop of its front matter's pages. Read into the page's lines it came
    // back as four lines of bars on a page of their own, and — standing at
    // x=0 — it put every page's left margin against the paper's edge.
    const pdf = onePagePdf(
      '/MediaBox [0 0 400 600]',
      [
        'BT /F1 12 Tf 60 500 Td (The body of the page starts here.) Tj ET',
        'BT /F1 12 Tf 60 486 Td (It runs on for a line or two more.) Tj ET',
        'BT /F1 40 Tf 0 -9 Td (off the sheet) Tj ET',
      ].join('\n'),
    );
    const flowed = reconstructByLayout(PdfFile.parse(pdf));
    const texts = paragraphs(flowed.doc).map((p) => p.paragraph.runs.map((r) => r.text).join(''));
    expect(texts.join(' ')).not.toContain('off the sheet');
    const mark = flowed.doc.body.find((b) => b.kind === 'shape');
    if (mark?.kind !== 'shape') throw new Error('expected the mark placed');
    expect(mark.shape.float?.posV?.relativeFrom).toBe('page');
    expect(flowed.doc.section?.margins?.left).toBeGreaterThan(50);
  });

  it('sets words that run UP the sheet across it, and says so', () => {
    // No section runs its lines up a sheet — Word and LibreOffice set a
    // section down it or across it — so /Rotate 270 over upright words is
    // read across the sheet, with a loss that names what was not kept.
    const pdf = onePagePdf(
      '/MediaBox [0 0 400 800] /Rotate 270',
      'BT /F1 20 Tf 100 600 Td (Up) Tj ET',
    );
    const flowed = reconstructByLayout(PdfFile.parse(pdf));
    expect(flowed.doc.section?.textDirection).toBeUndefined();
    expect(flowed.losses.some((l) => /running up the sheet/u.test(l.detail))).toBe(true);
  });

  it('leaves a page its box describes exactly where it stands', () => {
    // The same file with no turn: portrait, and the words still on their side.
    const placed = reconstructByLayout(PdfFile.parse(turnedPagePdf(0)), 'positional');
    expect(placed.doc.section?.pageSize?.orientation).toBe('portrait');
    const shape = placed.doc.body.find((b) => b.kind === 'shape');
    if (shape?.kind !== 'shape') throw new Error('expected a placed line');
    expect(shape.shape.transform?.rotation60k).toBe(90 * 60000);
  });

  it('measures a placed mark off the page’s own corner, not off the origin', () => {
    // §14.11.2 — a /MediaBox need not start at (0, 0), and 160F-2019.pdf's is
    // [-11.96 11.99 583.24 853.67]. Taking the corner for the origin put every
    // line and every rule twelve points up and to the left of the page's own.
    const placed = reconstructByLayout(PdfFile.parse(offsetBoxPdf()), 'positional');
    const text = placed.doc.body.find((b) => b.kind === 'shape' && b.shape.text !== undefined);
    const rule = placed.doc.body.find((b) => b.kind === 'shape' && b.shape.text === undefined);
    expect(text?.kind).toBe('shape');
    expect(rule?.kind).toBe('shape');
    if (text?.kind !== 'shape' || rule?.kind !== 'shape') return;
    // Text at x 20 on a box whose left edge is −12 stands 32pt in from the page.
    expect(text.shape.float?.posH?.offsetPt).toBeCloseTo(32, 5);
    // Baseline 700, box top 812: 812 − (700 − 2.5) − 12.5.
    expect(text.shape.float?.posV?.offsetPt).toBeCloseTo(102, 5);
    expect(rule.shape.float?.posH?.offsetPt).toBeCloseTo(32, 5);
    expect(rule.shape.float?.posV?.offsetPt).toBeCloseTo(682, 5); // 812 − 130
  });

  it('keeps a turned baseline turned, and out of the flat line it crosses', () => {
    // §9.4.2 — the text matrix turns as well as moves. 160F-2019.pdf sets
    // "Nature" on its side down the middle of a column; read flat, it joined
    // the row it happened to cross and lay across it.
    const placed = reconstructByLayout(PdfFile.parse(turnedLabelPdf()), 'positional');
    const shapes = placed.doc.body.filter((b) => b.kind === 'shape').map((b) => b.shape);
    expect(shapes).toHaveLength(2);
    const words = (s: (typeof shapes)[number]): string => {
      const first = s.text?.content[0];
      return first?.kind === 'paragraph' ? first.paragraph.runs.map((r) => r.text).join('') : '';
    };
    const flat = shapes.find((s) => words(s) === 'Flat');
    const side = shapes.find((s) => words(s) === 'Side');
    expect(flat).toBeDefined();
    expect(side).toBeDefined();
    // §20.1.7.6 — a shape turns clockwise, and this baseline runs up the page.
    expect(flat!.transform?.rotation60k).toBeUndefined();
    expect(side!.transform?.rotation60k).toBe(270 * 60000);
  });

  it('still flows by default, so a PDF reads back as a document', async () => {
    // The placed reading is opt-in: it has no reading order, no paragraphs and
    // no tables — which is exactly what a docx or a markdown conversion needs,
    // so the default must stay the flowed one.
    const docx = buildDocxFromBody('<w:p><w:r><w:t>FirstLine</w:t></w:r></w:p>');
    const pdf = await Ream.parse(docx).convert('pdf', { fonts: FONTS });
    const flowed = reconstructByLayout(PdfFile.parse(pdf));
    expect(flowed.doc.body.some((b) => b.kind === 'paragraph')).toBe(true);
    expect(flowed.doc.body.some((b) => b.kind === 'shape')).toBe(false);
  });
});

describe('heuristic layout reconstruction (E-PDF EP4)', () => {
  it('groups untagged text into paragraphs in reading order', async () => {
    const flow = await layoutFlow(
      spaced('AlphaLine') + spaced('BravoLine') + spaced('CharlieLine'),
    );
    const paras = paragraphs(flow);
    expect(paras.length).toBeGreaterThanOrEqual(2); // the paragraphs separated
    const joined = paras.map((p) => p.paragraph.runs.map((r) => r.text).join('')).join(' | ');
    expect(joined).toContain('AlphaLine');
    expect(joined).toContain('BravoLine');
    expect(joined).toContain('CharlieLine');
    expect(joined.indexOf('Alpha')).toBeLessThan(joined.indexOf('Bravo'));
    expect(joined.indexOf('Bravo')).toBeLessThan(joined.indexOf('Charlie'));
  });

  it('carries the source page size + orientation into the FlowDoc section (F1)', async () => {
    // Force a non-A4 landscape MediaBox on the generated PDF, then confirm the
    // reader reflects it back — so a re-render keeps the size/orientation
    // instead of falling back to the layout engine's A4 default.
    const pdf = await Ream.parse(buildDocxFromBody(spaced('OnlyLine'))).convert('pdf', {
      fonts: FONTS,
      pageWidth: 1000,
      pageHeight: 600,
    });
    const section = reconstructByLayout(PdfFile.parse(pdf)).doc.section;
    expect(section?.pageSize?.width).toBe(1000);
    expect(section?.pageSize?.height).toBe(600);
    expect(section?.pageSize?.orientation).toBe('landscape');
    // A PDF states no margins, but its WORDS say where they were: the leftmost
    // glyph is the left margin. Left at zero — which is what this used to do,
    // on the argument that the page box is the content box — a reflowed
    // document prints its text against the edge of the paper, which is what
    // every converted PDF looked like.
    expect(section?.margins?.left).toBeGreaterThan(0);
    expect(section?.margins?.top).toBeGreaterThan(0);
    // Never more than a third of the sheet: a margin that eats the text area
    // is worse than none.
    expect(section?.margins?.left).toBeLessThanOrEqual(1000 / 3);
    expect(section?.margins?.top).toBeLessThanOrEqual(600 / 3);
  });

  it('keeps a PLACED reading at zero margins, where every mark is anchored', () => {
    // The anchors are measured from the page, so a margin would move them all.
    const placed = reconstructByLayout(PdfFile.parse(twoColumnLinePdf()), 'positional');
    expect(placed.doc.section?.margins?.left).toBe(0);
    expect(placed.doc.section?.margins?.top).toBe(0);
  });

  it('reads the weight the page set, where the descriptor states none', async () => {
    // §9.8.1 — a descriptor that gives no /FontWeight and does not force bold
    // has said NOTHING about weight; reading that silence as "regular" is how
    // TAMReview.pdf's Times-Bold came back light, and every bold word on the
    // page with it — its title, "Abstract", "Keywords:".
    const body =
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>HeavyWord</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>PlainWord</w:t></w:r></w:p>';
    const flow = await layoutFlow(body);
    const runs = flow.body.flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs : []));
    expect(runs.find((r) => r.text.includes('Heavy'))?.properties.bold).toBe(true);
    expect(runs.find((r) => r.text.includes('Plain'))?.properties.bold).toBeFalsy();
  });

  it('ends a paragraph at a line that stops short of the measure', async () => {
    // Leading alone cannot tell a wrapped line from a finished one: two
    // paragraphs set with no extra space between them look exactly like one.
    // But a wrapping engine pulls the next word UP, so a line that stops well
    // short stopped because its author stopped it. alphatrans.pdf stacks five
    // short labels at ordinary leading and they came back as one paragraph,
    // re-wrapped into two lines of run-together text.
    const body =
      '<w:p><w:r><w:t>Short one</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Short two</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>A much longer line that runs the whole width of the text measure here</w:t></w:r></w:p>';
    const flow = await layoutFlow(body);
    const texts = paragraphs(flow).map((p) => p.paragraph.runs.map((r) => r.text).join(''));
    expect(texts.some((t) => t.startsWith('Short one') && !t.includes('Short two'))).toBe(true);
  });

  it('marks a line far larger than the median as a heading', async () => {
    const big = '<w:p><w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>BigTitle</w:t></w:r></w:p>';
    const flow = await layoutFlow(
      big + spaced('body one') + spaced('body two') + spaced('body three'),
    );
    const title = paragraphs(flow).find((p) =>
      p.paragraph.runs
        .map((r) => r.text)
        .join('')
        .includes('BigTitle'),
    );
    expect(title?.paragraph.properties.outlineLevel).toBe(0);
  });

  it('gives a rule to the paragraph it separates, not to the page (§17.3.1.24)', () => {
    // A rule sits in the white between two blocks: under a table's headings,
    // over a total. Anchored to the page at the y it was drawn at it stays
    // there while the words re-set — a receipt came back with a black line
    // struck through "Max plan - 5x" and another through "Payment history".
    const lines = [
      'BT /F1 10 Tf 1 0 0 1 72 700 Tm (Description) Tj ET',
      '0 0 0 RG 0.75 w 72 678 m 520 678 l S',
      'BT /F1 10 Tf 1 0 0 1 72 670 Tm (Max plan) Tj ET',
      'BT /F1 10 Tf 1 0 0 1 72 650 Tm (Aug 11 to Sep 11) Tj ET',
    ].join('\n');
    const doc = reconstructByLayout(
      PdfFile.parse(onePagePdf('/MediaBox [0 0 612 792]', lines)),
    ).doc;
    const paras = paragraphs(doc);
    const under = paras.find((p) =>
      p.paragraph.runs
        .map((r) => r.text)
        .join('')
        .includes('Max'),
    );
    expect(under?.paragraph.properties.borders?.top?.style).toBe('single');
    // …and the rule is not drawn a second time as a shape over the words.
    expect(doc.body.some((b) => b.kind === 'shape')).toBe(false);
  });

  it('reads lines set out on the same stops as a TABLE, and numbers the foot (§17.4.38)', () => {
    // An invoice's item table is a head and a row broken in the same places.
    // Written as tabbed paragraphs the picture is right and the document is
    // not: nothing downstream can read a column out of it.
    const rows = [
      'BT /F1 8 Tf 1 0 0 1 45 700 Tm (Description) Tj ET',
      'BT /F1 8 Tf 1 0 0 1 300 700 Tm (Qty) Tj ET',
      'BT /F1 8 Tf 1 0 0 1 420 700 Tm (Amount) Tj ET',
      'BT /F1 8 Tf 1 0 0 1 45 680 Tm (Max plan) Tj ET',
      'BT /F1 8 Tf 1 0 0 1 300 680 Tm (1) Tj ET',
      'BT /F1 8 Tf 1 0 0 1 424 680 Tm ($100.00) Tj ET',
    ].join('\n');
    const doc = reconstructByLayout(PdfFile.parse(onePagePdf('/MediaBox [0 0 612 792]', rows))).doc;
    const table = doc.body.find((b) => b.kind === 'table');
    expect(table?.kind).toBe('table');
    if (table?.kind !== 'table') return;
    expect(table.table.grid).toHaveLength(3);
    expect(table.table.rows).toHaveLength(2);
    const text = (r: number, c: number): string =>
      table.table.rows[r]!.cells[c]!.content.map((el) =>
        el.kind === 'paragraph' ? el.paragraph.runs.map((x) => x.text).join('') : '',
      ).join('');
    expect(text(0, 0)).toBe('Description');
    expect(text(1, 2)).toBe('$100.00');
  });

  it('keeps a rule typed out of hyphens on a line of its own', () => {
    // A page with no rule to draw types one. Read as prose it joined the
    // sentence over it: an invoice came back "…NOT to our San Francisco office.
    // ----------------------------" on one line, and the address the rule
    // introduces ran on from there.
    const lines = [
      'BT /F1 9 Tf 1 0 0 1 45 700 Tm (any checks must be sent to the address below.) Tj ET',
      'BT /F1 9 Tf 1 0 0 1 45 688 Tm (----------------------------) Tj ET',
      'BT /F1 9 Tf 1 0 0 1 45 676 Tm (PAYMENT ADDRESS) Tj ET',
    ].join('\n');
    const doc = reconstructByLayout(
      PdfFile.parse(onePagePdf('/MediaBox [0 0 612 792]', lines)),
    ).doc;
    const texts = paragraphs(doc).map((p) =>
      p.paragraph.runs
        .map((r) => r.text)
        .join('')
        .trim(),
    );
    expect(texts).toContain('----------------------------');
    expect(texts.some((t) => t.startsWith('any checks') && t.includes('---'))).toBe(false);
    expect(texts).toContain('PAYMENT ADDRESS');
  });

  it('measures the margins to the PICTURES too, not to the words alone', () => {
    // bug1708040.pdf is one short line over a logo 384 points wide. Measured to
    // the line, the right margin came in past the picture's own edge, and the
    // layout — which may not set a block wider than its measure — shrank the
    // logo by an eighth to fit.
    const wide = onePagePdf(
      '/MediaBox [0 0 612 792] /Resources << /XObject << /Im0 5 0 R >> >>',
      'BT /F1 10 Tf 1 0 0 1 72 700 Tm (a short line) Tj ET\nq 384 0 0 400 72 260 cm /Im0 Do Q',
      [
        `<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Length 4 >>\nstream\n\u0000\u0080\u0080\u0000\nendstream`,
      ],
    );
    const section = reconstructByLayout(PdfFile.parse(wide)).doc.section;
    // 612 - (72 + 384) = 156 at the most, and never so wide that the picture
    // no longer fits the measure it is set across.
    expect((section?.margins?.right as number) + (section?.margins?.left as number)).toBeLessThan(
      612 - 384,
    );
  });
});

describe('a flowing reading re-sets the page where the page set it', () => {
  /** A one-page letter-size sheet drawn in Helvetica, whose metrics every reader knows. */
  const helvetica = (content: string): Uint8Array =>
    onePagePdf('/MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >>', content, [
      '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    ]);
  const textOf = (p: { paragraph: { runs: ReadonlyArray<{ text: string }> } }): string =>
    p.paragraph.runs.map((r) => r.text).join('');

  it('stands its lines EXACTLY as far apart as the page stood them (§17.3.1.33)', () => {
    // An invoice sets its 9pt lines 13.5 apart. Left to a reader's single
    // spacing they closed up to the substitute face's own leading, and every
    // block of the page rose a little further than the one above it.
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            'BT /F1 9 Tf 1 0 0 1 45 700 Tm (Date due) Tj 1 0 0 1 108 700 Tm (August 12, 2026) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 45 686.5 Tm (Date paid) Tj 1 0 0 1 108 686.5 Tm (August 12, 2026) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 45 673 Tm (Receipt) Tj 1 0 0 1 108 673 Tm (2411) Tj ET',
          ].join('\n'),
        ),
      ),
    ).doc;
    const paras = paragraphs(doc);
    expect(paras).toHaveLength(3);
    for (const p of paras)
      expect(p.paragraph.properties).toMatchObject({ spacingLineRule: 'exact' });
    // Baseline to baseline: what the box above leaves below its baseline, the
    // white put before the next, and how far down its own box it stands.
    const [a, b] = paras.map((p) => p.paragraph.properties as Record<string, number | undefined>);
    const pitch =
      (1 - BASELINE_AT) * a!.spacingLine! + (b!.spacingBefore ?? 0) + BASELINE_AT * b!.spacingLine!;
    expect(pitch).toBeCloseTo(13.5, 1);
  });

  it('begins the page’s text where the page began it, not at the top margin', () => {
    // A Stripe invoice sets one blank at the very corner of the sheet and its
    // title forty points under it. The margin is measured to the blank, and
    // set against it the whole page rose by the difference.
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            'BT /F1 12 Tf 1 0 0 1 0 779 Tm ( ) Tj ET',
            'BT /F1 18 Tf 1 0 0 1 30 744 Tm (Invoice) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 30 714 Tm (Invoice number) Tj ET',
          ].join('\n'),
        ),
      ),
    ).doc;
    const title = paragraphs(doc).find((p) => textOf(p) === 'Invoice');
    const props = title?.paragraph.properties as Record<string, number | undefined> | undefined;
    const top = doc.section?.margins?.top as number;
    expect(top + (props?.spacingBefore ?? 0) + BASELINE_AT * (props?.spacingLine ?? 0)).toBeCloseTo(
      792 - 744,
      1,
    );
  });

  it('stands a value on the stop its column shares, however short the gap to it (§17.3.1.38)', () => {
    // "Date due" leaves its value a tab's width away; "Invoice number" leaves
    // it less than two ems, which a word space could almost be. Both values
    // stand at 108, and read as a space the second came back a word after its
    // label instead of in the column.
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            'BT /F1 9 Tf 1 0 0 1 30 714 Tm (Invoice number) Tj 1 0 0 1 108 714 Tm (6VOBWUGP) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 30 700.5 Tm (Date due) Tj 1 0 0 1 108 700.5 Tm (August 12, 2026) Tj ET',
          ].join('\n'),
        ),
      ),
    ).doc;
    const texts = paragraphs(doc).map(textOf);
    expect(texts).toContain('Invoice number\t6VOBWUGP');
    expect(texts).toContain('Date due\tAugust 12, 2026');
  });

  it('keeps a short line where its column starts, rather than reading it as centred', () => {
    // "walonade@icloud.com" closes an address block at the x every line of it
    // starts at; short, and near the middle of the sheet, it read as centred.
    const long =
      'This line runs the whole measure of the page from the left margin across to the right one, as prose does.';
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            'BT /F1 9 Tf 1 0 0 1 30 700 Tm (548 Market Street) Tj 1 0 0 1 250 700 Tm (Organization) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 250 686.5 Tm (Kazakhstan) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 250 673 Tm (someone@example.com) Tj ET',
            `BT /F1 9 Tf 1 0 0 1 30 600 Tm (${long}) Tj ET`,
          ].join('\n'),
        ),
      ),
    ).doc;
    const last = paragraphs(doc).find((p) => textOf(p) === 'someone@example.com');
    expect(last?.paragraph.properties).not.toMatchObject({ alignment: 'center' });
  });

  it('sets a column of figures against its right edge, and states every width (§17.4.38)', () => {
    // An invoice's "Qty" stands eight points left of the "1" under it and both
    // END at the same place. Set from the left the "1" stood under the Q; and
    // a table that states no widths is sized by its reader to its contents.
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            'BT /F1 8 Tf 1 0 0 1 45 700 Tm (Description) Tj ET',
            'BT /F1 8 Tf 1 0 0 1 300 700 Tm (Qty) Tj ET',
            'BT /F1 8 Tf 1 0 0 1 420 700 Tm (Amount) Tj ET',
            'BT /F1 8 Tf 1 0 0 1 45 680 Tm (Max plan) Tj ET',
            'BT /F1 8 Tf 1 0 0 1 308 680 Tm (1) Tj ET',
            'BT /F1 8 Tf 1 0 0 1 411.984 680 Tm ($1,100.00) Tj ET',
          ].join('\n'),
        ),
      ),
    ).doc;
    const table = doc.body.find((b) => b.kind === 'table');
    if (table?.kind !== 'table') throw new Error('the rows are a table');
    const alignment = (r: number, c: number): string | undefined => {
      const el = table.table.rows[r]!.cells[c]!.content[0];
      return el?.kind === 'paragraph' ? el.paragraph.properties.alignment : undefined;
    };
    expect(alignment(0, 0)).not.toBe('right');
    expect([alignment(0, 1), alignment(1, 1)]).toEqual(['right', 'right']);
    expect([alignment(0, 2), alignment(1, 2)]).toEqual(['right', 'right']);
    expect(table.table.properties).toMatchObject({ layout: 'fixed', widthType: 'dxa' });
    expect(table.table.rows[0]!.cells.every((c) => c.properties.width !== undefined)).toBe(true);
  });

  it('sets blocks with leading of their own side by side, as a table of one row', () => {
    // An invoice's stack of labels beside the address it bills, each a point or
    // three off the other's lines. Read across, "INVOICE → Jun 3, 2013 → 23
    // Main Street" was one line of three different things.
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            'BT /F1 9 Tf 1 0 0 1 360 669.5 Tm (Invoice Date) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 360 658.8 Tm (Jun 3, 2013) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 360 640.9 Tm (Invoice Number) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 472 668.5 Tm (Orange Demo Inc.) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 472 655.5 Tm (23 Main Street) Tj ET',
            'BT /F1 9 Tf 1 0 0 1 472 642.5 Tm (Central City) Tj ET',
          ].join('\n'),
        ),
      ),
    ).doc;
    const table = doc.body.find((b) => b.kind === 'table');
    if (table?.kind !== 'table') throw new Error('the band is a table');
    expect(table.table.rows).toHaveLength(1);
    const cells = table.table.rows[0]!.cells.map((c) =>
      c.content.map((el) => (el.kind === 'paragraph' ? textOf(el) : '')),
    );
    expect(cells).toEqual([
      ['Invoice Date', 'Jun 3, 2013', 'Invoice Number'],
      ['Orange Demo Inc.', '23 Main Street', 'Central City'],
    ]);
    // …and read once: nothing of it is left in the body's own lines.
    expect(paragraphs(doc).map(textOf).join(' ')).not.toContain('Invoice');
  });

  it('sets a block of PROSE beside another as the paragraphs it is, not a line apiece', () => {
    // comments.pdf's two columns, read as blocks side by side, were set a line
    // to a paragraph, and each line a word wider in a substitute's widths left
    // that word standing alone under it: half again as long.
    const line = 'the words of a column run out to its edge here';
    // Each column at a leading of its own, so their lines never share a baseline.
    const column = (x: number, top: number, leading: number): Array<string> =>
      Array.from(
        { length: 6 },
        (_, i) =>
          `BT /F1 9 Tf 1 0 0 1 ${String(x)} ${String(top - i * leading)} Tm (${line}) Tj ET`,
      );
    const doc = reconstructByLayout(
      PdfFile.parse(helvetica([...column(40, 700, 11), ...column(300, 697, 11.5)].join('\n'))),
    ).doc;
    const table = doc.body.find((b) => b.kind === 'table');
    if (table?.kind !== 'table') throw new Error('the band is a table');
    const cells = table.table.rows[0]!.cells.map((c) =>
      c.content.filter((el) => el.kind === 'paragraph'),
    );
    expect(cells.map((paras) => paras.length)).toEqual([1, 1]);
    expect(textOf(cells[0]![0]!)).toContain(`${line} ${line}`);
  });

  it('takes a rule drawn in pieces as one rule (§17.3.1.24)', () => {
    // An invoice draws the rule under its headings cell by cell. Measured apart
    // only the widest piece was long enough to be a rule: it moved with the
    // row, and the rest stayed where the page drew them, through the figures.
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            'BT /F1 9 Tf 1 0 0 1 30 700 Tm (Description) Tj ET',
            '0 0 0 RG 0.75 w 28 690 m 270 690 l S',
            '0 0 0 RG 0.75 w 270 690 m 350 690 l S',
            '0 0 0 RG 0.75 w 350 690 m 430 690 l S',
            'BT /F1 9 Tf 1 0 0 1 30 677 Tm (Project management) Tj ET',
          ].join('\n'),
        ),
      ),
    ).doc;
    const ruled = paragraphs(doc).find((p) => textOf(p) === 'Project management');
    expect(ruled?.paragraph.properties.borders?.top?.style).toBe('single');
    expect(doc.body.some((b) => b.kind === 'shape')).toBe(false);
  });

  it('rules EVERY member of a set of ruled lines, not only the first (§17.3.1.5)', () => {
    // An invoice rules each line of its totals. Paragraphs with the same
    // borders are one bordered set, ruled on its outside only, and read as
    // five tops the totals came back with one rule over "Subtotal".
    const doc = reconstructByLayout(
      PdfFile.parse(
        helvetica(
          [
            '0.92 0.92 0.92 RG 0.75 w 306 712 m 582 712 l S',
            'BT /F1 9 Tf 1 0 0 1 306 700 Tm (Subtotal) Tj ET',
            '0.92 0.92 0.92 RG 0.75 w 306 697.75 m 582 697.75 l S',
            'BT /F1 9 Tf 1 0 0 1 306 685.75 Tm (Total) Tj ET',
          ].join('\n'),
        ),
      ),
    ).doc;
    const ruled = paragraphs(doc).filter((p) => p.paragraph.properties.borders?.top !== undefined);
    expect(ruled).toHaveLength(2);
    for (const p of ruled) {
      const borders = p.paragraph.properties.borders as Record<string, unknown>;
      expect(borders.insideH).toEqual(borders.top);
    }
  });
});

describe('a right-to-left line ends on its left (§17.3.1.13)', () => {
  it('reads a line short of the measure on the LEFT as the end of a paragraph', () => {
    // ArabicCIDTrueType.pdf sets four lines flush right, each shorter than the
    // one before. Read as left-to-right they all reached the measure — their
    // right ends are where they start — and every pair ran together.
    const measure = { left: 160, right: 510 };
    const full = { x: 160, width: 350, fontSize: 36, text: 'انواع الخطوط العربية' };
    const short = { x: 270, width: 240, fontSize: 36, text: 'انواع الخطوط العربية' };
    expect(endedParagraph(short, { x: 258, width: 252 }, measure)).toBe(true);
    expect(endedParagraph(full, { x: 270, width: 240 }, measure)).toBe(false);
  });
});

describe('what a page draws for want of characters', () => {
  /** A traced glyph: a box at (x, y), `w` wide and 7 tall. */
  const glyph = (x: number, y: number, w = 4) => ({
    orderKey: [x],
    segs: [
      { op: 'move' as const, x, y },
      { op: 'line' as const, x: x + w, y },
      { op: 'line' as const, x: x + w, y: y + 7 },
      { op: 'close' as const },
    ],
    minX: x,
    minY: y,
    maxX: x + w,
    maxY: y + 7,
    fillHex: '231F20',
    glyph: true,
  });

  it('draws a word of traced glyphs as one shape, and the next word as another', () => {
    // TAMReview.pdf sets its body in a subset whose glyphs name nothing, and
    // traced one glyph at a time it came back as forty-two thousand shapes —
    // a package no reader opened in under three minutes.
    const words = drawnWords([
      glyph(10, 100),
      glyph(14.5, 100),
      glyph(19, 100),
      glyph(40, 100),
      glyph(44.5, 100),
      glyph(10, 80),
    ]);
    expect(words).toHaveLength(3);
    expect(words[0]).toMatchObject({ minX: 10, maxX: 23 });
    expect(words[0]!.segs).toHaveLength(12);
  });
});

describe('a blank page is still a page', () => {
  it('keeps every sheet of a document that draws nothing on them', () => {
    // doc_actions.pdf is three blank sheets, and came back as one.
    const kids = [3, 4, 5];
    const objects = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      `<< /Type /Pages /Kids [${kids.map((k) => `${String(k)} 0 R`).join(' ')}] /Count 3 >>`,
      ...kids.map(() => '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>'),
    ];
    let pdf = '%PDF-1.7\n';
    const offsets: Array<number> = [];
    objects.forEach((body, i) => {
      offsets.push(pdf.length);
      pdf += `${String(i + 1)} 0 obj\n${body}\nendobj\n`;
    });
    const xref = pdf.length;
    pdf += `xref\n0 ${String(objects.length + 1)}\n0000000000 65535 f \n`;
    for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`;
    pdf += `trailer\n<< /Size ${String(objects.length + 1)} /Root 1 0 R >>\nstartxref\n${String(xref)}\n%%EOF\n`;
    const doc = reconstructByLayout(PdfFile.parse(new TextEncoder().encode(pdf))).doc;
    const breaks = doc.body.filter(
      (b) => b.kind === 'paragraph' && b.paragraph.properties.pageBreakBefore === true,
    );
    expect(breaks).toHaveLength(2);
    // …and the first sheet holds a line of its own, so the second begins after it.
    expect(
      doc.body[0]?.kind === 'paragraph' && doc.body[0].paragraph.properties.pageBreakBefore,
    ).not.toBe(true);
  });
});

describe('a sheet of a line or two shows no measure (§17.3.1)', () => {
  const COURIER =
    '/MediaBox [0 0 612 792] /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 ' +
    '/BaseFont /Courier >> >> >>';
  const texts = (pdf: Uint8Array): Array<string> =>
    paragraphs(reconstructByLayout(PdfFile.parse(pdf)).doc).map((p) =>
      p.paragraph.runs.map((r) => r.text).join(''),
    );

  it('keeps its lines as the page set them', () => {
    // Its longest line reaches the edge only because it IS the edge.
    // checkbox-bad-appearance.pdf sets "Checkbox 1 - not checked" over
    // "Checkbox 2 - Checked", and run together the two came back side by side.
    const pdf = onePagePdf(
      COURIER,
      'BT /F1 10 Tf 50 700 Td (Checkbox 1 - not checked) Tj 0 -12 Td (Checkbox 2 - Checked) Tj ET',
    );
    expect(texts(pdf)).toEqual(['Checkbox 1 - not checked', 'Checkbox 2 - Checked']);
  });

  it('still runs together the lines of a sheet that shows one', () => {
    // Two lines out of three break at the same edge: that is the measure.
    const pdf = onePagePdf(
      COURIER,
      'BT /F1 10 Tf 50 700 Td (aaaa bbbb cccc dddd) Tj 0 -12 Td (eeee ffff gggg hhhh) Tj ' +
        '0 -12 Td (iiii.) Tj ET',
    );
    expect(texts(pdf)).toEqual(['aaaa bbbb cccc dddd eeee ffff gggg hhhh iiii.']);
  });
});

describe('a word space is as wide as the type it stands in', () => {
  it('sets the space it puts between two runs in the size of the first', () => {
    // issue10665_reduced.pdf sets "78" and "110" twenty points apart in
    // 60-point type, and a space at the document's default size between them
    // closed them up to "78110".
    const pdf = onePagePdf(
      '/MediaBox [0 0 250 100]',
      'BT /F1 60 Tf 20 25 Td (78) Tj 80 0 Td (110) Tj ET',
    );
    const runs = paragraphs(reconstructByLayout(PdfFile.parse(pdf)).doc).flatMap(
      (p) =>
        p.paragraph.runs as ReadonlyArray<{ text: string; properties: { fontSizePt?: number } }>,
    );
    expect(runs.map((r) => r.text).join('')).toBe('78 110');
    // Set in the digits' own size, the space runs on with them as one run.
    const spaced = runs.find((r) => r.text.includes(' '));
    expect(spaced?.properties.fontSizePt).toBeCloseTo(60, 0);
  });
});

describe('a line set where the page set it', () => {
  it('stands its baseline where its box was measured to put it (§17.3.1.33)', () => {
    // The box is a line and a quarter of the size tall with its top an em
    // over the baseline, which is where an EXACT line of that height puts it
    // (`BASELINE_AT`). Left to single spacing the face's own ascent placed
    // it, a tenth of an em too high: bug1724918.pdf's field values rode up
    // against the tops of their fields.
    const el = positionedText(
      [{ text: 'world', sizePt: 12 }],
      { x: 10, y: 10, width: 60, height: 15 },
      { left: 0, top: 100 },
      1,
    );
    if (el.kind !== 'shape') throw new Error('a positioned line is a shape');
    const first = el.shape.text?.content[0];
    if (first?.kind !== 'paragraph') throw new Error('holding a paragraph');
    expect(first.paragraph.properties).toMatchObject({
      spacingLine: 15,
      spacingLineRule: 'exact',
    });
    expect(BASELINE_AT * 15).toBeCloseTo(12, 6);
  });

  it('is carried by a paragraph that takes no room', () => {
    // Floats in a row share one carrier, the first one's; left at single
    // spacing a placed line's was a blank line that moved everything under it.
    const el = positionedText(
      [{ text: 'world', sizePt: 12 }],
      { x: 10, y: 10, width: 60, height: 15 },
      { left: 0, top: 100 },
      1,
    );
    if (el.kind !== 'shape') throw new Error('a positioned line is a shape');
    expect(el.shape.paragraphProperties).toEqual(FLOAT_CARRIER);
  });
});

describe('what an annotation writes stands in its own box (§12.5.5)', () => {
  /**
   * A page of one line of prose, and a push button whose appearance fills
   * its box grey and writes "Execute" in it — evaljs.pdf's button.
   */
  const buttonPdf = (): Uint8Array => {
    // A line under the button too, so the caption is not the page's last word
    // (which is placed where it stands for a reason of its own).
    const content =
      'BT /F1 12 Tf 40 700 Td (A line of the page itself.) Tj 0 -600 Td (And one under it.) Tj ET';
    const ap = '0.75 g 0 0 72 20 re f BT /F1 12 Tf 0 g 13 6 Td (Execute) Tj ET';
    return onePagePdf(
      '/MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> ' +
        '/Annots [<< /Type /Annot /Subtype /Widget /FT /Btn /Rect [265 142 337 162] ' +
        '/AP << /N 6 0 R >> >>]',
      content,
      [
        '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        '<< /Type /XObject /Subtype /Form /BBox [0 0 72 20] ' +
          `/Resources << /Font << /F1 5 0 R >> >> /Length ${String(ap.length)} >>\n` +
          `stream\n${ap}\nendstream`,
      ],
    );
  };

  it('marks the words an appearance writes as the annotation’s', () => {
    const file = PdfFile.parse(buttonPdf());
    const runs = extractPageText(file, file.pages()[0]!);
    expect(runs.find((r) => r.text.includes('Execute'))?.annotation).toBe(true);
    expect(runs.find((r) => r.text.includes('page itself'))?.annotation).toBeUndefined();
  });

  it('places them where the box stands, over the box they are written on', () => {
    // Read into the page's lines, evaljs.pdf's "Execute" stood at the margin,
    // two hundred points from its button; placed but keyed with the page's
    // own words, it went under the grey of the button it names.
    const body = reconstructByLayout(PdfFile.parse(buttonPdf())).doc.body;
    const caption = body.find(
      (b) =>
        b.kind === 'shape' &&
        b.shape.text?.content.some(
          (p) => p.kind === 'paragraph' && p.paragraph.runs.some((r) => r.text === 'Execute'),
        ),
    );
    if (caption?.kind !== 'shape') throw new Error('the caption is placed');
    expect(caption.shape.float?.posH?.offsetPt).toBeCloseTo(278, 0);
    const fills = body.filter(
      (b) => b.kind === 'shape' && b.shape.text === undefined && b.shape.fill.kind === 'solid',
    );
    expect(fills.length).toBeGreaterThan(0);
    for (const fill of fills) {
      if (fill.kind !== 'shape') continue;
      expect(caption.shape.float?.zOrder ?? -1).toBeGreaterThan(fill.shape.float?.zOrder ?? 0);
    }
    // The page's own line is still read as the page's.
    expect(
      paragraphs({ body }).some((p) => p.paragraph.runs.some((r) => r.text.includes('itself'))),
    ).toBe(true);
  });
});
