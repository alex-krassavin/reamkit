// E-COMMENTS CM0 — a docx review comment. A body run carries w:commentReference
// @w:id; word/comments.xml holds the comment's block content plus author/date.
// The reader loads it into FlowDoc.comments and tags the anchoring run with
// commentRef. Rendering (inline marker + "Comments" section) lands in CM1.

import { zipSync } from 'fflate';

import { describe, expect, it } from 'vitest';

import { Ream } from '@/core/converter/ream';

const enc = new TextEncoder();
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

function commentDocx(withComments = true): Uint8Array {
  const body =
    `<w:p>` +
    `<w:commentRangeStart w:id="0"/>` +
    `<w:r><w:t>Reviewed text</w:t></w:r>` +
    `<w:commentRangeEnd w:id="0"/>` +
    `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>` +
    `</w:p>`;
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`;
  const comments =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="${W_NS}">` +
    `<w:comment w:id="0" w:author="Alice Reviewer" w:initials="AR" w:date="2026-01-02T10:00:00Z">` +
    `<w:p><w:r><w:t>Please clarify this sentence.</w:t></w:r></w:p>` +
    `</w:comment></w:comments>`;

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT}">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        (withComments
          ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
          : '') +
        `</Types>`,
    ),
    '_rels/.rels': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    'word/document.xml': enc.encode(document),
  };
  if (withComments) {
    files['word/_rels/document.xml.rels'] = enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId10" Type="${R_NS}/comments" Target="comments.xml"/></Relationships>`,
    );
    files['word/comments.xml'] = enc.encode(comments);
  }
  return zipSync(files);
}

function bodyRuns(
  doc: ReturnType<typeof Ream.parse>,
): Array<{ text: string; commentRef?: string }> {
  return doc.flow.body.flatMap((e) => (e.kind === 'paragraph' ? [...e.paragraph.runs] : []));
}

describe('Word comments in docx (E-COMMENTS CM0)', () => {
  it('loads the comment content and attribution into FlowDoc.comments', () => {
    const c = Ream.parse(commentDocx()).flow.comments?.get('0');
    expect(c?.author).toBe('Alice Reviewer');
    expect(c?.initials).toBe('AR');
    expect(c?.date).toBe('2026-01-02T10:00:00Z');
    const text = (c?.content ?? [])
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join('');
    expect(text).toBe('Please clarify this sentence.');
  });

  it('tags the anchoring run with the comment id, keeping the commented text in flow', () => {
    const runs = bodyRuns(Ream.parse(commentDocx()));
    expect(runs.some((r) => r.commentRef === '0')).toBe(true);
    expect(runs.map((r) => r.text).join('')).toContain('Reviewed text');
  });

  it('omits the comments field when the docx ships no comments part', () => {
    // The body's commentReference still tags the run (a dangling id the renderer
    // skips), but with no comments.xml there is no content map to attach.
    const doc = Ream.parse(commentDocx(false));
    expect(doc.flow.comments).toBeUndefined();
    expect(
      bodyRuns(doc)
        .map((r) => r.text)
        .join(''),
    ).toContain('Reviewed text');
  });
});
