// E-PDF EP4 — heuristic reconstruction for untagged PDFs. Convert a docx to a
// plain (untagged) PDF, then rebuild a FlowDoc from the positioned text alone:
// lines clustered by baseline, paragraphs by vertical spacing, headings by a
// font size above the median.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { reconstructByLayout } from '@/pdf-reader/layout';

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
      paragraph: { properties: { outlineLevel?: number }; runs: ReadonlyArray<{ text: string }> };
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
    // not two paragraphs.
    const paras = parasOf(
      [
        line(40, 360, 'A line of body text right across the measure'),
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
          ['<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'],
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

describe('a running foot is a foot, not a paragraph (§17.6.13)', () => {
  /** Three pages of body with the same line standing alone at the bottom. */
  const paper = (): Uint8Array => {
    const page = (n: number): string => {
      const ops: Array<string> = [];
      for (let i = 0; i < 8; i++)
        ops.push(
          `BT /F0 10 Tf 1 0 0 1 40 ${String(360 - i * 14)} Tm (body line ${String(i)}) Tj ET`,
        );
      // Alone at the foot, a long way below the text block.
      ops.push(`BT /F0 8 Tf 1 0 0 1 40 20 Tm (The Journal of Things ${String(n)}) Tj ET`);
      return ops.join('\n');
    };
    const contents = [page(1), page(2), page(3)];
    const objects: Array<string> = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R 5 0 R 7 0 R] /Count 3 >>',
    ];
    contents.forEach((content, i) => {
      objects.push(
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Contents ${String(4 + i * 2)} 0 R ` +
          '/Resources << /Font << /F0 9 0 R >> >> >>',
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
    const flowed = reconstructByLayout(PdfFile.parse(twoColumnLinePdf()));
    const paras = paragraphs(flowed.doc);
    expect(paras).toHaveLength(1);
    expect(paras[0]!.paragraph.runs.map((r) => r.text).join('')).toBe('Left Right');
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
});
