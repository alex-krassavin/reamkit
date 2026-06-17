// E-SHEET W7 — cell comments / notes. Legacy notes (xl/comments) and modern
// threaded comments (xl/threadedComments + xl/persons) are read through the
// worksheet relationships and listed in a "Comments" section after the grid
// (Excel's "print comments at end of sheet"): a heading + one line per comment,
// each "<ref> — <author>: <text>". Render-only — not written back.

import { readFileSync } from 'node:fs';

import { describe, expect, it } from 'vitest';

import { buildXlsx } from './fixtures/build-xlsx';
import type { BodyElement } from '@/core/document-model';
import { parseLegacyComments, parsePersons, parseThreadedComments } from '@/excel/comments-parser';
import { Ream } from '@/core/converter/ream';
import { convertXlsxToPdfSync } from '@/core/converter';

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);
const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const TC = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

// The flattened text of every paragraph in the body (after the grid table).
function paragraphTexts(body: ReadonlyArray<BodyElement>): Array<string> {
  const out: Array<string> = [];
  for (const el of body) {
    if (el.kind === 'paragraph') out.push(el.paragraph.runs.map((r) => r.text).join(''));
  }
  return out;
}

describe('comment parsers (E-SHEET W7)', () => {
  it('parses legacy notes and resolves authorId, stripping the author prefix', () => {
    const xml = `<comments xmlns="${M}"><authors><author>Ada</author><author>Bo</author></authors>
      <commentList>
        <comment ref="B2" authorId="1"><text><r><t>Bo:</t></r><r><t xml:space="preserve">
check this</t></r></text></comment>
        <comment ref="C3" authorId="0"><text><r><t>looks good</t></r></text></comment>
      </commentList></comments>`;
    const out = parseLegacyComments(enc(xml));
    expect(out).toEqual([
      { ref: 'B2', author: 'Bo', text: 'check this', threaded: false },
      { ref: 'C3', author: 'Ada', text: 'looks good', threaded: false },
    ]);
  });

  it('resolves threaded comments through the person directory', () => {
    const persons = parsePersons(
      enc(
        `<personList xmlns="${TC}"><person displayName="Ada" id="p1"/><person displayName="Bo" id="p2"/></personList>`,
      ),
    );
    const out = parseThreadedComments(
      enc(
        `<ThreadedComments xmlns="${TC}"><threadedComment ref="A1" id="{1}" personId="p1"><text>first</text></threadedComment><threadedComment ref="A1" id="{2}" personId="p2"><text>reply</text></threadedComment></ThreadedComments>`,
      ),
      persons,
    );
    expect(out).toEqual([
      { ref: 'A1', author: 'Ada', text: 'first', threaded: true },
      { ref: 'A1', author: 'Bo', text: 'reply', threaded: true },
    ]);
  });
});

describe('cell comments — end to end (E-SHEET W7)', () => {
  it('lists legacy notes in a Comments section after the grid', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['data']],
        comments: [
          { ref: 'A1', author: 'Ada', text: 'the first note' },
          { ref: 'B2', author: 'Bo', text: 'a second note' },
        ],
      }),
    ).flow;
    const texts = paragraphTexts(flow.body);
    expect(texts).toContain('Comments');
    expect(texts).toContain('A1 — Ada: the first note');
    expect(texts).toContain('B2 — Bo: a second note');
  });

  it('lists threaded comments with their resolved authors', () => {
    const flow = Ream.parse(
      buildXlsx({
        rows: [['data']],
        threadedComments: [
          { ref: 'A1', personId: 'p1', text: 'question?' },
          { ref: 'A1', personId: 'p2', text: 'answer.' },
        ],
        persons: [
          { id: 'p1', name: 'Ada' },
          { id: 'p2', name: 'Bo' },
        ],
      }),
    ).flow;
    const texts = paragraphTexts(flow.body);
    expect(texts).toContain('A1 — Ada: question?');
    expect(texts).toContain('A1 — Bo: answer.');
  });

  it('adds no Comments section to a sheet without comments (byte-zero)', () => {
    const flow = Ream.parse(buildXlsx({ rows: [['data']] })).flow;
    expect(paragraphTexts(flow.body)).not.toContain('Comments');
  });

  it('renders a commented sheet to a valid PDF', () => {
    const pdf = convertXlsxToPdfSync(
      buildXlsx({ rows: [['x']], comments: [{ ref: 'A1', author: 'Ada', text: 'note' }] }),
      {
        fonts: {
          regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
          bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
        },
      },
    );
    expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe('%PDF-');
    expect(pdf.length).toBeGreaterThan(1000);
  });
});
