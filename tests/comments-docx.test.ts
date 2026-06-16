// E-COMMENTS CM0 — a docx review comment. A body run carries w:commentReference
// @w:id; word/comments.xml holds the comment's block content plus author/date.
// The reader loads it into FlowDoc.comments and tags the anchoring run with
// commentRef. Rendering (inline marker + "Comments" section) lands in CM1.

import { readFileSync } from 'node:fs';

import { zipSync } from 'fflate';

import { describe, expect, it } from 'vitest';

import { Ream } from '@/core/converter/ream';
import { PdfFile } from '@/pdf-reader/document';
import { extractPageText } from '@/pdf-reader/text';

const FONTS = {
  regular: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Regular.ttf')),
  bold: new Uint8Array(readFileSync('tests/fixtures/fonts/Roboto-Bold.ttf')),
};

const enc = new TextEncoder();
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const REL_COMMENTS_EXTENDED =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const COMMENTS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_EX_CT =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';

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

  it('renders an inline marker and a Comments section in HTML (CM1)', async () => {
    const html = new TextDecoder().decode(await Ream.parse(commentDocx()).convert('html'));
    // The in-text marker links to the comment's entry in the end section.
    expect(html).toContain('id="cmref-1"');
    expect(html).toContain('href="#cm-1"');
    expect(html).toContain('[1]');
    // The Comments section carries the author and the comment text.
    expect(html).toContain('<section class="comments">');
    expect(html).toContain('Alice Reviewer');
    expect(html).toContain('Please clarify this sentence.');
  });

  it('renders the marker and a comment entry in the PDF (CM1)', async () => {
    const file = PdfFile.parse(await Ream.parse(commentDocx()).convert('pdf', { fonts: FONTS }));
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('')
      .replace(/\s/g, '');
    expect(text).toContain('[1]'); // the in-text marker
    expect(text).toContain('AliceReviewer'); // the comment entry's attribution
    expect(text).toContain('Pleaseclarifythissentence.'); // the comment body, after the document
  });

  it('makes the comment marker a clickable jump to its entry (CM2)', async () => {
    const pdf = Buffer.from(
      await Ream.parse(commentDocx()).convert('pdf', { fonts: FONTS }),
    ).toString('latin1');
    expect(pdf).toContain('/Subtype /Link'); // the marker is a link annotation
    expect(pdf).toContain('/S /GoTo'); // an internal jump to the comment entry
  });

  it('round-trips comments through the docx writer (CM3)', async () => {
    const out = await Ream.parse(commentDocx()).convert('docx');
    const reread = Ream.parse(out);
    const c = reread.flow.comments?.get('0');
    expect(c?.author).toBe('Alice Reviewer');
    expect(c?.initials).toBe('AR');
    expect(c?.date).toBe('2026-01-02T10:00:00Z');
    const text = (c?.content ?? [])
      .flatMap((b) => (b.kind === 'paragraph' ? b.paragraph.runs.map((r) => r.text) : []))
      .join('');
    expect(text).toBe('Please clarify this sentence.');
    // the body run still anchors the comment by id on the re-read
    expect(bodyRuns(reread).some((r) => r.commentRef === '0')).toBe(true);
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

// E-COMMENTS CM4 — a reply thread. Two comments (a parent and a reply) each
// carry a w14:paraId; word/commentsExtended.xml links the reply to its parent
// (w15:paraIdParent) and marks the parent thread resolved (w15:done="1").
function threadedCommentDocx(): Uint8Array {
  const body =
    `<w:p>` +
    `<w:commentRangeStart w:id="0"/>` +
    `<w:r><w:t>Claim under review</w:t></w:r>` +
    `<w:commentRangeEnd w:id="0"/>` +
    `<w:r><w:commentReference w:id="0"/></w:r>` +
    `<w:commentRangeStart w:id="1"/>` +
    `<w:commentRangeEnd w:id="1"/>` +
    `<w:r><w:commentReference w:id="1"/></w:r>` +
    `</w:p>`;
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`;
  const comments =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">` +
    `<w:comment w:id="0" w:author="Alice Reviewer" w:initials="AR" w:date="2026-01-02T10:00:00Z">` +
    `<w:p w14:paraId="0000AAA1"><w:r><w:t>Needs a citation</w:t></w:r></w:p>` +
    `</w:comment>` +
    `<w:comment w:id="1" w:author="Bob Author" w:initials="BA" w:date="2026-01-03T09:00:00Z">` +
    `<w:p w14:paraId="0000BBB2"><w:r><w:t>Citation added now</w:t></w:r></w:p>` +
    `</w:comment></w:comments>`;
  const commentsExtended =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w15:commentsEx xmlns:w15="${W15_NS}">` +
    `<w15:commentEx w15:paraId="0000AAA1" w15:done="1"/>` +
    `<w15:commentEx w15:paraId="0000BBB2" w15:paraIdParent="0000AAA1" w15:done="0"/>` +
    `</w15:commentsEx>`;

  return zipSync({
    '[Content_Types].xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT}">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_CT}"/>` +
        `<Override PartName="/word/commentsExtended.xml" ContentType="${COMMENTS_EX_CT}"/>` +
        `</Types>`,
    ),
    '_rels/.rels': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    'word/document.xml': enc.encode(document),
    'word/_rels/document.xml.rels': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId10" Type="${R_NS}/comments" Target="comments.xml"/>` +
        `<Relationship Id="rId11" Type="${REL_COMMENTS_EXTENDED}" Target="commentsExtended.xml"/>` +
        `</Relationships>`,
    ),
    'word/comments.xml': enc.encode(comments),
    'word/commentsExtended.xml': enc.encode(commentsExtended),
  });
}

describe('Word comment threads in docx (E-COMMENTS CM4)', () => {
  it('threads replies and resolved flags from commentsExtended', () => {
    const flow = Ream.parse(threadedCommentDocx()).flow;
    const parent = flow.comments?.get('0');
    const reply = flow.comments?.get('1');
    expect(parent?.done).toBe(true);
    expect(parent?.parentId).toBeUndefined();
    expect(reply?.parentId).toBe('0');
    expect(reply?.done).toBeFalsy();
  });

  it('nests a reply under its parent and flags the resolved thread in HTML', async () => {
    const html = new TextDecoder().decode(await Ream.parse(threadedCommentDocx()).convert('html'));
    expect(html).toContain('class="comment resolved"'); // parent thread is resolved
    expect(html).toContain('Resolved');
    expect(html).toContain('<ol class="comment-replies">'); // reply nests under the parent
    expect(html).toContain('Citation added now'); // the reply's content
  });

  it('annotates the reply and the resolved thread in the PDF', async () => {
    const file = PdfFile.parse(
      await Ream.parse(threadedCommentDocx()).convert('pdf', { fonts: FONTS }),
    );
    const text = extractPageText(file, file.pages()[0]!)
      .map((r) => r.text)
      .join('')
      .replace(/\s/g, '');
    expect(text).toContain('(resolved)'); // parent entry marked resolved
    expect(text).toContain('inreplyto[1]'); // reply points at its parent's marker
  });

  it('round-trips the thread (parent + resolved) through the docx writer', async () => {
    const out = await Ream.parse(threadedCommentDocx()).convert('docx');
    const reread = Ream.parse(out);
    // commentsExtended is re-emitted with fresh paraIds; the semantic links survive.
    expect(reread.flow.comments?.get('0')?.done).toBe(true);
    expect(reread.flow.comments?.get('1')?.parentId).toBe('0');
  });
});

// word/people.xml resolves a comment author's presence identity (usually an
// email) — matched on the author display name and attached as Comment.authorId.
function peopleCommentDocx(): Uint8Array {
  const body =
    `<w:p><w:r><w:t>Reviewed text</w:t></w:r>` + `<w:r><w:commentReference w:id="0"/></w:r></w:p>`;
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`;
  const comments =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="${W_NS}">` +
    `<w:comment w:id="0" w:author="Alice Reviewer" w:initials="AR" w:date="2026-01-02T10:00:00Z">` +
    `<w:p><w:r><w:t>Please clarify.</w:t></w:r></w:p>` +
    `</w:comment></w:comments>`;
  const people =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w15:people xmlns:w15="${W15_NS}">` +
    `<w15:person w15:author="Alice Reviewer">` +
    `<w15:presenceInfo w15:providerId="None" w15:userId="alice@example.com"/>` +
    `</w15:person></w15:people>`;
  const peopleCt = 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml';

  return zipSync({
    '[Content_Types].xml': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT}">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_CT}"/>` +
        `<Override PartName="/word/people.xml" ContentType="${peopleCt}"/>` +
        `</Types>`,
    ),
    '_rels/.rels': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    'word/document.xml': enc.encode(document),
    'word/_rels/document.xml.rels': enc.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">` +
        `<Relationship Id="rId10" Type="${R_NS}/comments" Target="comments.xml"/></Relationships>`,
    ),
    'word/comments.xml': enc.encode(comments),
    'word/people.xml': enc.encode(people),
  });
}

describe('Word comment author identity from people.xml (E-COMMENTS)', () => {
  it('resolves the author identity onto the comment', () => {
    const c = Ream.parse(peopleCommentDocx()).flow.comments?.get('0');
    expect(c?.author).toBe('Alice Reviewer');
    expect(c?.authorId).toBe('alice@example.com');
  });

  it('shows the identity in the HTML comment meta', async () => {
    const html = new TextDecoder().decode(await Ream.parse(peopleCommentDocx()).convert('html'));
    expect(html).toContain('Alice Reviewer');
    expect(html).toContain('alice@example.com');
  });
});
