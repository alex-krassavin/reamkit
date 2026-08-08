import { describe, expect, it } from 'vitest';

import { buildDocxFromBody } from './fixtures/build-docx';
import { createConverter } from '@/core/converter/facade';
import { Ream } from '@/core/converter/ream';
import { markdownWriter, writeMarkdown } from '@/markdown/markdown-writer';
import { readDocx } from '@/word/docx-reader';

const decode = (b: Uint8Array) => new TextDecoder().decode(b);

/** Read a docx fixture and render it to markdown — the flow medium, no fonts. */
function md(bodyXml: string, options?: Parameters<typeof buildDocxFromBody>[1]): string {
  const { doc } = readDocx(buildDocxFromBody(bodyXml, options));
  return decode(writeMarkdown(doc).bytes);
}

const p = (inner: string) => `<w:p>${inner}</w:p>`;
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

describe('markdown writer (FlowDoc adapter)', () => {
  it('renders paragraphs and outline-level headings', () => {
    const out = md(
      p('<w:pPr><w:outlineLvl w:val="0"/></w:pPr>' + t('Title')) +
        p('<w:pPr><w:outlineLvl w:val="2"/></w:pPr>' + t('Third level')) +
        p(t('Body text.')),
    );
    expect(out).toBe('# Title\n\n### Third level\n\nBody text.\n');
  });

  it('falls back to a "Heading N" style id when no outline level is declared', () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>';
    const out = md(
      p('<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>' + t('Styled heading')) +
        p('<w:pPr><w:pStyle w:val="Title"/></w:pPr>' + t('Doc title')),
      { stylesXml: styles },
    );
    expect(out).toContain('## Styled heading');
    expect(out).toContain('# Doc title');
  });

  it('maps bold, italic and strike to their GFM delimiters', () => {
    const out = md(
      p(
        '<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>' +
          '<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>' +
          '<w:r><w:rPr><w:strike/></w:rPr><w:t>gone</w:t></w:r>' +
          '<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>both</w:t></w:r>',
      ),
    );
    expect(out).toBe('**bold***italic*~~gone~~***both***\n');
  });

  it('says underline and super/subscript with the inline html GFM allows', () => {
    const out = md(
      p(
        '<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>under</w:t></w:r>' +
          '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>up</w:t></w:r>' +
          '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>down</w:t></w:r>',
      ),
    );
    expect(out).toBe('<u>under</u><sup>up</sup><sub>down</sub>\n');
  });

  it('coalesces adjacent runs that carry the same emphasis', () => {
    // Emitted per run these would read `**one****two**` — legal, unreadable.
    const out = md(
      p(
        '<w:r><w:rPr><w:b/></w:rPr><w:t>one </w:t></w:r>' +
          '<w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t></w:r>',
      ),
    );
    expect(out).toBe('**one two**\n');
  });

  it('moves a span’s outer whitespace outside the delimiters', () => {
    // CommonMark §6.2: `** bold **` opens nothing at all.
    const bold = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> b </w:t></w:r>';
    expect(md(p(t('a') + bold + t('c')))).toBe('a **b** c\n');
  });

  it('strips the trailing spaces that would read as a hard break', () => {
    expect(md(p(t('trailing  ')))).toBe('trailing\n');
  });

  it('escapes markdown metacharacters in text', () => {
    const out = md(p(t('a *star*, a [bracket], a `tick` and a | pipe')));
    expect(out).toBe('a \\*star\\*, a \\[bracket\\], a \\`tick\\` and a \\| pipe\n');
  });

  it('escapes an underscore only where it could open emphasis', () => {
    const out = md(p(t('snake_case but _emphasis_ too')));
    expect(out).toBe('snake_case but \\_emphasis\\_ too\n');
  });

  it('guards a line that would otherwise open a block it never asked for', () => {
    const out = md(p(t('1. not a list')) + p(t('# not a heading')) + p(t('- not a bullet')));
    expect(out).toBe('\\1. not a list\n\n\\# not a heading\n\n\\- not a bullet\n');
  });

  it('renders a w:br soft break as a GFM hard break', () => {
    const out = md(p('<w:r><w:t>one</w:t><w:br/><w:t>two</w:t></w:r>'));
    expect(out).toBe('one\\\ntwo\n');
  });

  it('drops an empty paragraph — a blank line is markdown’s own separator', () => {
    const out = md(p(t('before')) + p('') + p(t('after')));
    expect(out).toBe('before\n\nafter\n');
  });

  it('reports each recurring omission once, not once per paragraph', () => {
    const centred = p('<w:pPr><w:jc w:val="center"/></w:pPr>' + t('x'));
    const { doc } = readDocx(buildDocxFromBody(centred.repeat(20)));
    const { losses } = writeMarkdown(doc);
    const alignment = losses.filter((l) => l.detail.includes('alignment'));
    expect(alignment).toHaveLength(1);
    expect(alignment[0]!.severity).toBe('dropped');
  });

  it('exposes the writer adapter', () => {
    expect(markdownWriter.id).toBe('md');
    expect(markdownWriter.consumes).toBe('flow');
    const { doc } = readDocx(buildDocxFromBody(p(t('hi'))));
    expect(decode(markdownWriter.write(doc).bytes)).toBe('hi\n');
  });
});

describe('markdown writer — lists', () => {
  /** An abstract numbering + instance, one level per (ilvl, format, template). */
  const numbering = (
    levels: ReadonlyArray<{ ilvl: number; format: string; text: string; start?: number }>,
    numId = '1',
  ): string =>
    `<w:abstractNum w:abstractNumId="0">${levels
      .map(
        (l) =>
          `<w:lvl w:ilvl="${l.ilvl}"><w:start w:val="${l.start ?? 1}"/>` +
          `<w:numFmt w:val="${l.format}"/><w:lvlText w:val="${l.text}"/></w:lvl>`,
      )
      .join('')}</w:abstractNum>` +
    `<w:num w:numId="${numId}"><w:abstractNumId w:val="0"/></w:num>`;

  const item = (text: string, ilvl = 0, numId = '1') =>
    p(
      `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
        t(text),
    );

  it('re-derives a bullet list from the materialized markers', () => {
    const out = md(item('first') + item('second'), {
      numberingXml: numbering([{ ilvl: 0, format: 'bullet', text: '' }]),
    });
    // One block, plain newline between items: a blank line would make the list
    // loose and wrap every item's text in its own paragraph.
    expect(out).toBe('- first\n- second\n');
  });

  it('re-derives an ordered list, keeping the numbers the source states', () => {
    const out = md(item('one') + item('two'), {
      numberingXml: numbering([{ ilvl: 0, format: 'decimal', text: '%1.', start: 5 }]),
    });
    expect(out).toBe('5. one\n6. two\n');
  });

  it('indents a nested level to its parent’s content column', () => {
    const out = md(item('parent') + item('child', 1) + item('back'), {
      numberingXml: numbering([
        { ilvl: 0, format: 'decimal', text: '%1.' },
        { ilvl: 1, format: 'bullet', text: '' },
      ]),
    });
    expect(out).toBe('1. parent\n   - child\n2. back\n');
  });

  it('carries one list on across two instances of the same abstract definition', () => {
    // §17.9.23 — the counter belongs to the abstract numbering, not the
    // instance, so these are one list continued and not two.
    const two =
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>';
    const out = md(item('a') + item('b') + item('c', 0, '2'), { numberingXml: two });
    expect(out).toBe('1. a\n2. b\n3. c\n');
  });

  it('a plain paragraph closes the list', () => {
    const out = md(item('one') + p(t('prose')) + item('two'), {
      numberingXml: numbering([{ ilvl: 0, format: 'bullet', text: '' }]),
    });
    expect(out).toBe('- one\n\nprose\n\n- two\n');
  });

  it('numbers a numbered heading in its text rather than demoting it to a list', () => {
    const styles =
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
      '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
    const out = md(p('<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' + t('Introduction')), {
      stylesXml: styles,
      numberingXml: numbering([{ ilvl: 0, format: 'decimal', text: '%1.' }]),
    });
    expect(out).toBe('# 1. Introduction\n');
  });

  it('reports a marker alphabet markdown cannot keep', () => {
    const { doc } = readDocx(
      buildDocxFromBody(item('a') + item('b'), {
        numberingXml: numbering([{ ilvl: 0, format: 'lowerLetter', text: '%1.' }]),
      }),
    );
    const { bytes, losses } = writeMarkdown(doc);
    // Still an ordered list — only its alphabet is normalized.
    expect(decode(bytes)).toBe('1. a\n2. b\n');
    expect(losses.some((l) => l.detail.includes('lowerLetter'))).toBe(true);
  });
});

describe('markdown writer — emphasis that has to flank', () => {
  it('falls back to a tag when the delimiter could not open', () => {
    // §6.2: `1**. x**` — the `**` is preceded by a letter and followed by
    // punctuation, so it is not left-flanking and prints as two asterisks.
    const out = md(p(t('1') + '<w:r><w:rPr><w:b/></w:rPr><w:t>. Auksin</w:t></w:r>'));
    expect(out).toBe('1<strong>. Auksin</strong>\n');
  });

  it('keeps the delimiter when it flanks perfectly well', () => {
    const out = md(p(t('a ') + '<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>'));
    expect(out).toBe('a **bold**\n');
  });

  it('moves a non-breaking space out of the span like any other', () => {
    // U+00A0 is Unicode whitespace, and a closing `**` behind one cannot close.
    const nb = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Figure\u00a0</w:t></w:r>';
    expect(md(p(nb + t('x')))).toBe('**Figure**\u00a0x\n');
  });

  it('drops a hard break at the end of a cell’s paragraph', () => {
    const brCell =
      '<w:tc><w:p><w:r><w:t>one</w:t><w:br/></w:r></w:p><w:p>' + t('two') + '</w:p></w:tc>';
    const tbl2 =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr>${brCell}<w:tc><w:p>${t('z')}</w:p></w:tc></w:tr></w:tbl>`;
    expect(md(tbl2)).toBe('| one<br>two | z |\n| --- | --- |\n');
  });
});

describe('markdown writer — what counts as empty', () => {
  const numbering =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
    '<w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const li = (text: string) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${text ? t(text) : ''}</w:p>`;

  it('drops a marked paragraph with nothing in it — a bullet against no words', () => {
    const out = md(li('one') + li('') + li('two'), { numberingXml: numbering });
    expect(out).toBe('- one\n- two\n');
  });

  it('an empty paragraph between items does not split the list', () => {
    // A deck spaces its bullets with blank lines; split, each becomes its own list.
    const out = md(li('one') + p('') + li('two'), { numberingXml: numbering });
    expect(out).toBe('- one\n- two\n');
  });

  it('treats a zero-width-only paragraph as empty', () => {
    // The .pptx reader marks a slide boundary with a U+200B paragraph.
    const out = md(p(t('before')) + p(t('\u200b')) + p(t('after')));
    expect(out).toBe('before\n\nafter\n');
  });

  it('moves a hard break out of an emphasis span that would not close', () => {
    // `**text\` + newline + `**` is not right-flanking: it prints two asterisks.
    const out = md(p('<w:r><w:rPr><w:b/></w:rPr><w:t>Title</w:t><w:br/></w:r>' + t('after')));
    expect(out).toBe('**Title**\\\nafter\n');
  });

  it('drops a hard break with no next line to break to', () => {
    const out = md(p('<w:r><w:rPr><w:b/></w:rPr><w:t>Title</w:t><w:br/></w:r>'));
    expect(out).toBe('**Title**\n');
  });
});

describe('markdown writer — tables', () => {
  const cell = (text: string, tcPr = '') =>
    `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}<w:p>${t(text)}</w:p></w:tc>`;
  const tbl = (rows: string, grid = 2) =>
    '<w:tbl><w:tblGrid>' + '<w:gridCol w:w="2000"/>'.repeat(grid) + `</w:tblGrid>${rows}</w:tbl>`;

  it('renders a GFM pipe table, first row as the header', () => {
    const out = md(
      tbl(`<w:tr>${cell('A')}${cell('B')}</w:tr>` + `<w:tr>${cell('c')}${cell('d')}</w:tr>`),
    );
    expect(out).toBe('| A | B |\n| --- | --- |\n| c | d |\n');
  });

  it('carries each column’s alignment in the delimiter row', () => {
    const out = md(
      tbl(
        '<w:tr>' +
          '<w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
          t('mid') +
          '</w:p></w:tc>' +
          '<w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr>' +
          t('end') +
          '</w:p></w:tc>' +
          '</w:tr>',
      ),
    );
    expect(out).toBe('| mid | end |\n| :---: | ---: |\n');
  });

  it('flattens spans and vertical merges into the plain grid markdown has', () => {
    const rows =
      `<w:tr>${cell('A', '<w:gridSpan w:val="2"/>')}${cell('B', '<w:vMerge w:val="restart"/>')}</w:tr>` +
      `<w:tr>${cell('c')}${cell('d')}${cell('', '<w:vMerge/>')}</w:tr>`;
    const { doc } = readDocx(buildDocxFromBody(tbl(rows, 3)));
    const { bytes, losses } = writeMarkdown(doc);
    expect(decode(bytes)).toBe('| A |  | B |\n| --- | --- | --- |\n| c | d |  |\n');
    expect(losses.some((l) => l.detail.includes('merged cells flattened'))).toBe(true);
  });

  it('keeps run formatting inside a cell and joins its blocks with <br>', () => {
    const rich =
      '<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>' +
      `<w:p>${t('second line')}</w:p></w:tc>`;
    const out = md(tbl(`<w:tr>${rich}${cell('plain')}</w:tr>`));
    expect(out).toBe('| **bold**<br>second line | plain |\n| --- | --- |\n');
  });

  it('trims the padding Word carries as spaces in the cell text', () => {
    const out = md(tbl(`<w:tr>${cell('  Cell 1 ')}${cell('b')}</w:tr>`));
    expect(out).toBe('| Cell 1 | b |\n| --- | --- |\n');
  });

  it('escapes a pipe in cell text so it cannot split the row', () => {
    const out = md(tbl(`<w:tr>${cell('a|b')}${cell('c')}</w:tr>`));
    expect(out).toBe('| a\\|b | c |\n| --- | --- |\n');
  });

  it('flattens a nested table into the cell that holds it', () => {
    const inner = tbl(`<w:tr>${cell('x')}${cell('y')}</w:tr>`);
    const outer = `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>`;
    const { doc } = readDocx(buildDocxFromBody(outer));
    const { bytes, losses } = writeMarkdown(doc);
    expect(decode(bytes)).toContain('x<br>y');
    expect(losses.some((l) => l.feature === 'tables.nested')).toBe(true);
  });

  it('flattens a heading inside a cell — a cell holds inline content only', () => {
    const head =
      '<w:tc><w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr>' + t('Title') + '</w:p></w:tc>';
    const out = md(tbl(`<w:tr>${head}${cell('b')}</w:tr>`));
    expect(out).toBe('| Title | b |\n| --- | --- |\n');
  });

  it('renders a list inside a cell as text — a row is one line', () => {
    const numbering =
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const li = (text: string) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${t(text)}</w:p>`;
    const out = md(tbl(`<w:tr><w:tc>${li('one')}${li('two')}</w:tc>${cell('z')}</w:tr>`), {
      numberingXml: numbering,
    });
    expect(out).toBe('| 1. one<br>2. two | z |\n| --- | --- |\n');
  });
});

describe('markdown writer — links and pictures', () => {
  const PNG = new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52,
  ]);
  const link = (rId: string, text: string) => `<w:hyperlink r:id="${rId}">${t(text)}</w:hyperlink>`;
  const drawing = (rId: string, alt = '') =>
    '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<wp:extent cx="914400" cy="914400"/>' +
    `<wp:docPr id="1" name="p" descr="${alt}"/>` +
    `<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill>` +
    '</pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';

  it('renders an external hyperlink, passing it through the scheme allowlist', () => {
    const out = md(p(link('rId100', 'Ream')), {
      hyperlinks: { rId100: 'https://reamkit.dev/' },
    });
    expect(out).toBe('[Ream](https://reamkit.dev/)\n');
  });

  it('leaves a disallowed scheme as plain text and says so', () => {
    const { doc } = readDocx(
      buildDocxFromBody(p(link('rId100', 'click')), {
        hyperlinks: { rId100: 'javascript:alert(1)' },
      }),
    );
    const { bytes, losses } = writeMarkdown(doc);
    expect(decode(bytes)).toBe('click\n');
    expect(decode(bytes)).not.toContain('javascript');
    expect(losses.some((l) => l.feature === 'hyperlinks')).toBe(true);
  });

  it('wraps a destination that holds whitespace in pointy brackets', () => {
    const out = md(p(link('rId100', 'x')), {
      hyperlinks: { rId100: 'https://example.com/a b' },
    });
    expect(out).toBe('[x](<https://example.com/a b>)\n');
  });

  it('links an internal target to the anchor it plants on the bookmark', () => {
    const body =
      p('<w:hyperlink w:anchor="Chapter One">' + t('see below') + '</w:hyperlink>') +
      '<w:bookmarkStart w:id="1" w:name="Chapter One"/>' +
      p(t('Here')) +
      '<w:bookmarkEnd w:id="1"/>';
    const out = md(body);
    expect(out).toContain('[see below](#chapter-one)');
    expect(out).toContain('<a id="chapter-one"></a>Here');
  });

  it('inlines a picture as a data URI by default', () => {
    const out = md(p(`<w:r>${drawing('rId50')}</w:r>`), {
      images: { rId50: { contentType: 'image/png', extension: 'png', bytes: PNG } },
    });
    expect(out).toContain('![](data:image/png;base64,');
  });

  it('names a picture under ./media when the caller writes the bytes itself', () => {
    const { doc } = readDocx(
      buildDocxFromBody(p(`<w:r>${drawing('rId50')}</w:r>`), {
        images: { rId50: { contentType: 'image/png', extension: 'png', bytes: PNG } },
      }),
    );
    const { bytes, losses } = writeMarkdown(doc, { images: 'link' });
    expect(decode(bytes)).toBe('![](./media/image1.png)\n');
    expect(losses.some((l) => l.detail.includes('./media'))).toBe(true);
  });

  it('omits pictures entirely on request', () => {
    const { doc } = readDocx(
      buildDocxFromBody(p(`<w:r>${drawing('rId50')}</w:r>`), {
        images: { rId50: { contentType: 'image/png', extension: 'png', bytes: PNG } },
      }),
    );
    const { bytes, losses } = writeMarkdown(doc, { images: 'drop' });
    expect(decode(bytes)).toBe('');
    expect(losses.some((l) => l.severity === 'dropped' && l.feature === 'images')).toBe(true);
  });
});

describe('markdown writer — notes, comments and shapes', () => {
  it('renders footnotes and endnotes as GFM footnotes', () => {
    const body =
      p(t('Claim') + '<w:r><w:footnoteReference w:id="2"/></w:r>') +
      p(t('Aside') + '<w:r><w:endnoteReference w:id="2"/></w:r>');
    const out = md(body, {
      footnotesXml:
        '<w:footnote w:id="2"><w:p><w:r><w:t>The source, 1999.</w:t></w:r></w:p></w:footnote>',
      endnotesXml:
        '<w:endnote w:id="2"><w:p><w:r><w:t>A closing remark.</w:t></w:r></w:p></w:endnote>',
    });
    expect(out).toContain('Claim[^fn1]');
    expect(out).toContain('Aside[^en1]');
    expect(out).toContain('[^fn1]: The source, 1999.');
    expect(out).toContain('[^en1]: A closing remark.');
  });

  it('carries a review comment as a footnote attributed to its author', () => {
    const body = p(t('Reviewed text') + '<w:r><w:commentReference w:id="0"/></w:r>');
    const { doc } = readDocx(
      buildDocxFromBody(body, {
        commentsXml:
          '<w:comment w:id="0" w:author="Alice Reviewer" w:initials="AR">' +
          '<w:p><w:r><w:t>Please clarify.</w:t></w:r></w:p></w:comment>',
      }),
    );
    const { bytes, losses } = writeMarkdown(doc);
    expect(decode(bytes)).toContain('Reviewed text[^cm1]');
    expect(decode(bytes)).toContain('[^cm1]: **Alice Reviewer:** Please clarify.');
    expect(losses.some((l) => l.detail.includes('rendered as footnotes'))).toBe(true);
  });

  it('drops a reference whose note is not in the package', () => {
    const out = md(p(t('Claim') + '<w:r><w:footnoteReference w:id="7"/></w:r>'));
    expect(out).toBe('Claim\n');
  });

  it('keeps the words inside a shape and drops only its geometry', () => {
    const shape =
      '<w:p><w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml">' +
      '<v:rect style="width:100pt;height:50pt"><v:textbox><w:txbxContent>' +
      '<w:p><w:r><w:t>Words in a box</w:t></w:r></w:p>' +
      '</w:txbxContent></v:textbox></v:rect></w:pict></w:r></w:p>';
    const { doc } = readDocx(buildDocxFromBody(shape));
    const { bytes, losses } = writeMarkdown(doc);
    expect(decode(bytes)).toContain('Words in a box');
    expect(losses.some((l) => l.feature === 'shapes')).toBe(true);
  });
});

describe('markdown writer — page breaks', () => {
  const broken = (text: string) => p(`<w:pPr><w:pageBreakBefore/></w:pPr>${t(text)}`);

  it('drops a page break and reports it, by default', () => {
    const { doc } = readDocx(buildDocxFromBody(p(t('one')) + broken('two')));
    const { bytes, losses } = writeMarkdown(doc);
    expect(decode(bytes)).toBe('one\n\ntwo\n');
    expect(losses.some((l) => l.detail.includes('page breaks'))).toBe(true);
  });

  it('writes a thematic break instead when asked, and stops reporting it', () => {
    const { doc } = readDocx(buildDocxFromBody(p(t('one')) + broken('two')));
    const { bytes, losses } = writeMarkdown(doc, { pageBreaks: 'rule' });
    expect(decode(bytes)).toBe('one\n\n---\n\ntwo\n');
    expect(losses.some((l) => l.detail.includes('page breaks'))).toBe(false);
  });

  it('never opens the document with a rule', () => {
    const { doc } = readDocx(buildDocxFromBody(broken('first')));
    expect(decode(writeMarkdown(doc, { pageBreaks: 'rule' }).bytes)).toBe('first\n');
  });

  it('never writes two rules in a row', () => {
    // The .pptx reader marks a slide boundary with an empty zero-width
    // paragraph: the rule stands for it, and the paragraph itself is nothing.
    const { doc } = readDocx(buildDocxFromBody(p(t('one')) + broken('\u200b') + broken('two')));
    expect(decode(writeMarkdown(doc, { pageBreaks: 'rule' }).bytes)).toBe('one\n\n---\n\ntwo\n');
  });

  it('writes no rule inside a cell, where three hyphens are three hyphens', () => {
    const inner = `<w:tc><w:p>${t('a')}</w:p>${broken('b')}</w:tc>`;
    const table =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
      `<w:tr>${inner}<w:tc><w:p>${t('z')}</w:p></w:tc></w:tr></w:tbl>`;
    const { doc } = readDocx(buildDocxFromBody(table));
    expect(decode(writeMarkdown(doc, { pageBreaks: 'rule' }).bytes)).toBe(
      '| a<br>b | z |\n| --- | --- |\n',
    );
  });

  it('reaches the writer through Ream.convert', async () => {
    const docx = buildDocxFromBody(p(t('one')) + broken('two'));
    const out = decode(await Ream.parse(docx).convert('md', { pageBreaks: 'rule' }));
    expect(out).toBe('one\n\n---\n\ntwo\n');
  });
});

describe('markdown writer — the md target', () => {
  it('converts through Ream with no fonts and no I/O', async () => {
    const docx = buildDocxFromBody(
      p('<w:pPr><w:outlineLvl w:val="0"/></w:pPr>' + t('Title')) + p(t('Body.')),
    );
    const out = decode(await Ream.parse(docx).convert('md'));
    expect(out).toBe('# Title\n\nBody.\n');
  });

  it('passes the picture mode through to the writer', async () => {
    const docx = buildDocxFromBody(p(t('x')));
    const out = decode(await Ream.parse(docx).convert('md', { images: 'drop' }));
    expect(out).toBe('x\n');
  });

  it('converts through the facade', async () => {
    const docx = buildDocxFromBody(p(t('facade')));
    const { bytes } = await createConverter().convert(docx, { to: 'md' });
    expect(decode(bytes)).toBe('facade\n');
  });

  it('renders a heading style with no outline level as a heading in html too', async () => {
    // core/outline is one rule for every target; html used to miss the fallback.
    const docx = buildDocxFromBody(p('<w:pPr><w:pStyle w:val="Heading3"/></w:pPr>' + t('H')), {
      stylesXml:
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/></w:style>',
    });
    expect(decode(await Ream.parse(docx).convert('html'))).toContain('<h3');
    expect(decode(await Ream.parse(docx).convert('md'))).toContain('### H');
  });
});
