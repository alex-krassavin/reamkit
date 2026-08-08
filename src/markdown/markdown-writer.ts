// Markdown writer (ir-design §7): the sixth adapter, and the second written
// against the FLOW tree. Like the HTML writer it is a flow medium — no
// pagination, no layout engine, no fonts to embed — so
// `Ream.parse(bytes).convert('md')` performs zero I/O.
//
// The dialect is GitHub-Flavored Markdown (GFM): CommonMark plus tables,
// strikethrough and footnotes. GFM also parses markdown inside INLINE html
// spans, which is what lets underline and super/subscript — three things
// CommonMark cannot say at all — survive as `<u>` / `<sup>` / `<sub>` wrappers
// around ordinarily-escaped text.
//
// Stage-6 contract at work: `flow.body` already carries FINAL effective
// properties, so this writer reads values straight off the resolved
// properties. It still routes every read through the cascade resolver over the
// EMPTY sheet — on reader output a memoized identity, on a hand-built raw tree
// an actual resolve, exactly like the HTML writer and the PDF layout.
//
// Markdown is a far narrower medium than HTML, so a great deal is deliberately
// dropped: alignment, indents, colour, font family and size, tab stops, page
// breaks, headers/footers. Every one of those records a Loss — deduplicated by
// (severity, feature, detail), because a document with two hundred centred
// paragraphs must not produce two hundred identical report lines.

import type {
  BodyElement,
  Numbering,
  NumberingLevel,
  Paragraph,
  Run,
  ShapeBlock,
  Table,
  TableCell,
  TableRow,
} from '@/core/document-model';
import type { DocumentWriter, WriteResult } from '@/core/ir/adapters';
import type { FlowDoc } from '@/core/ir/flow';
import type { Loss, LossSeverity, ResourceId, ResourceStore } from '@/core/ir';
import type { ResolvedParagraphProperties, ResolvedRunProperties } from '@/core/style-cascade';

import { FEATURES } from '@/core/ir';
import { headingLevelOf } from '@/core/outline';
import { detectImageFormat } from '@/core/images';
import { effectiveAbstract } from '@/core/numbering/state';
import { sanitizeHref } from '@/core/links';
import { toBase64 } from '@/core/bytes';
import {
  EMPTY_STYLE_SHEET,
  resolveParagraphProperties,
  resolveRunProperties,
} from '@/core/style-cascade';

/** Options for {@link writeMarkdown}. */
export interface MarkdownWriteOptions {
  /**
   * How a picture reaches the output. `'dataUri'` (the default) inlines the
   * bytes as a `data:` URI, keeping the single-blob {@link WriteResult}
   * contract and the zero-I/O promise; `'link'` emits a relative
   * `./media/…` path and records that the bytes were not written anywhere;
   * `'drop'` omits pictures entirely.
   */
  readonly images?: 'dataUri' | 'link' | 'drop';
  /**
   * What a page break becomes. `'drop'` (the default) leaves it out and
   * reports it: a paginated document breaks its pages wherever the layout
   * needs to, and a rule at every one of them would litter the text. `'rule'`
   * writes a `---` thematic break instead, which is what a SLIDE DECK wants —
   * the `.pptx` and `.ppt` readers mark each slide boundary with a page break,
   * and it is the only structure a deck has.
   */
  readonly pageBreaks?: 'rule' | 'drop';
}

/**
 * Render a {@link FlowDoc} to GitHub-Flavored Markdown (ir-design §7).
 *
 * A flow medium: no pagination, no layout engine and no fonts, so this is a
 * pure, zero-I/O transform. Markdown says far less than the document model
 * does — alignment, indents, colour, font metrics, tab stops and page
 * geometry have no expression at all — so those are dropped and reported as
 * {@link Loss} entries, deduplicated so one recurring omission reports once.
 *
 * @param flow    The format-neutral interlayer document tree.
 * @param options Picture handling; see {@link MarkdownWriteOptions}.
 * @returns The encoded Markdown bytes plus the recorded {@link Loss} list.
 */
export function writeMarkdown(flow: FlowDoc, options: MarkdownWriteOptions = {}): WriteResult {
  const ctx: EmitCtx = {
    losses: [],
    seen: new Set<string>(),
    images: options.images ?? 'dataUri',
    pageBreaks: options.pageBreaks ?? 'drop',
    list: [],
    inCell: false,
    resources: flow.resources,
    anchors: referencedAnchors(flow.body),
    notes: noteNumbers(flow),
    mediaNames: new Map<ResourceId, string>(),
    ...(flow.numbering ? { numbering: flow.numbering } : {}),
  };

  if (flow.headersFooters && flow.headersFooters.size > 0) {
    lose(ctx, 'dropped', FEATURES.headersFooters, 'markdown has no pages to band headers onto');
  }

  const blocks: Array<string> = [];
  for (const el of flow.body) emitBlock(blocks, el, ctx);
  emitNoteDefinitions(blocks, flow, ctx);

  return { bytes: new TextEncoder().encode(joinBlocks(blocks)), losses: ctx.losses };
}

/**
 * The flow-medium {@link DocumentWriter} adapter (id `'md'`), wrapping
 * {@link writeMarkdown}, with the set of {@link FEATURES} it renders.
 */
export const markdownWriter: DocumentWriter<FlowDoc> = {
  id: 'md',
  consumes: 'flow',
  supports: new Set([
    FEATURES.text,
    FEATURES.lists,
    FEATURES.tables,
    FEATURES.images,
    FEATURES.hyperlinks,
  ]),
  write: (doc, opts) => writeMarkdown(doc, opts ?? {}),
};

interface EmitCtx {
  readonly losses: Array<Loss>;
  /** Keys of losses already recorded, so a recurring omission reports once. */
  readonly seen: Set<string>;
  readonly images: 'dataUri' | 'link' | 'drop';
  readonly pageBreaks: 'rule' | 'drop';
  /** The list levels currently open, outermost first (see {@link ListLevel}). */
  readonly list: Array<ListLevel>;
  /** The document's raw numbering definitions, for the ordered/bullet decision. */
  readonly numbering?: Numbering;
  /** Inside a table cell, where only inline content exists. */
  inCell: boolean;
  readonly resources: ResourceStore;
  /** Bookmark names some run actually links to — the only ones worth an anchor. */
  readonly anchors: ReadonlySet<string>;
  /** Stable `./media/…` names per resource, for `images: 'link'`. */
  readonly mediaNames: Map<ResourceId, string>;
  /** §17.11 note and §17.13.4 comment ids numbered by order of reference. */
  readonly notes: NoteNumbers;
}

/** Footnote, endnote and comment ids numbered by where they are referenced. */
interface NoteNumbers {
  readonly footnotes: ReadonlyMap<string, number>;
  readonly endnotes: ReadonlyMap<string, number>;
  readonly comments: ReadonlyMap<string, number>;
}

/**
 * One open list level. `indent` is the column its own markers start at and
 * `markerWidth` the width of the last marker plus its space — which is where a
 * nested level starts, since CommonMark measures a sublist against the content
 * column of the item containing it, not against a fixed step.
 */
interface ListLevel {
  readonly ilvl: number;
  readonly indent: number;
  markerWidth: number;
  counter: number;
}

/** Record a loss unless an identical one was recorded already. */
function lose(ctx: EmitCtx, severity: LossSeverity, feature: string, detail: string): void {
  const key = `${severity}|${feature}|${detail}`;
  if (ctx.seen.has(key)) return;
  ctx.seen.add(key);
  ctx.losses.push({ severity, feature, detail });
}

/**
 * Join the emitted blocks: one blank line between them (markdown's block
 * separator), no leading blank, exactly one trailing newline.
 *
 * Trailing spaces are stripped from every line: they are invisible, editors
 * and formatters eat them, and two of them are markdown's own hard break —
 * a meaning no document ever asked for. Ream writes a hard break as the
 * backslash GFM also accepts, which survives all three.
 */
function joinBlocks(blocks: ReadonlyArray<string>): string {
  const body = blocks
    .map(trimHardBreaks)
    .filter((b) => b.length > 0)
    .join('\n\n');
  return body.length > 0 ? `${body.replace(/[ \t]+$/gm, '')}\n` : '';
}

/**
 * Drop a hard break sitting at either end of a block: at the end there is no
 * next line to break to, at the start no previous one, and either way the
 * backslash is left standing in the text as itself. Only a break is taken —
 * an escaped literal backslash is `\\` with no newline behind it.
 */
function trimHardBreaks(block: string): string {
  return block.replace(/^(?:\\\n)+/, '').replace(/\\\n$/, '');
}

// ---------------------------------------------------------------------------
// Blocks
// ---------------------------------------------------------------------------

function emitBlock(out: Array<string>, el: BodyElement, ctx: EmitCtx): void {
  if (el.kind === 'paragraph') {
    emitParagraph(out, el.paragraph, ctx);
    return;
  }
  // Any other block ends whatever list was running.
  ctx.list.length = 0;
  if (el.kind === 'table') {
    emitTable(out, el.table, ctx);
  } else if (el.kind === 'image') {
    const img = pictureMarkdown(el.image.resource, el.image.altText, ctx);
    if (img.length > 0) out.push(img);
  } else if (el.kind === 'chart') {
    lose(ctx, 'dropped', FEATURES.charts, 'markdown cannot draw a chart');
  } else {
    emitShape(out, el.shape, ctx);
  }
}

/**
 * A shape contributes its WORDS. Markdown draws no geometry, no fill and no
 * line, but the text inside a callout or a text box is document content and is
 * emitted as ordinary blocks — including a group's members, however deep.
 */
function emitShape(out: Array<string>, shape: ShapeBlock, ctx: EmitCtx): void {
  lose(ctx, 'dropped', FEATURES.shapes, 'shape geometry dropped; the text inside it is kept');
  for (const el of shape.text?.content ?? []) emitBlock(out, el, ctx);
  for (const child of shape.children ?? []) emitShape(out, child.shape, ctx);
}

function emitParagraph(out: Array<string>, p: Paragraph, ctx: EmitCtx): void {
  const resolved = resolveParagraphProperties(p.properties, EMPTY_STYLE_SHEET);
  // §17.13.6.2 — a bookmark something links to opens here, as the inline
  // anchor markdown has no syntax of its own for.
  //
  // The leading whitespace goes: a paragraph Word indents with spaces or a tab
  // would open an INDENTED CODE BLOCK at four of them (CommonMark §4.4), and
  // the whole of its text would be typeset as source. The indent itself has no
  // markdown expression either way, and is already reported as dropped.
  const anchors = bookmarkAnchors(p, ctx);
  const text = inlineRuns(p.runs, p, ctx).replace(/^[ \t]+/, '');
  const inline = anchors + text;
  const marker = markerText(p, ctx);
  reportParagraphLosses(resolved, ctx);
  if (breaksPage(p, resolved)) emitRule(out, ctx);
  // Nothing to say and nowhere to point: the paragraph contributes no block.
  // An anchor alone is content — it is what an internal link lands on.
  const empty = isBlank(text) && anchors.length === 0;

  // A heading outranks a list: §17.9.24 numbers Word's own Heading styles, and
  // demoting "1 Introduction" to a list item would cost the document its whole
  // outline. The chapter number rides along in the heading's text instead.
  const level = headingLevelOf(resolved);
  if (level !== undefined) {
    if (empty) return;
    ctx.list.length = 0;
    const heading = marker !== undefined ? `${marker} ${inline}` : inline;
    // §4.2 ATX headings are single-line: a soft break inside one would end it.
    const oneLine = heading.replaceAll('\\\n', ' ').replaceAll('\n', ' ');
    // A table cell holds INLINE content only: a `#` written there is a literal
    // hash, not a heading, so inside one the heading keeps its text alone.
    if (ctx.inCell) {
      lose(ctx, 'degraded', FEATURES.tables, 'heading inside a cell flattened to plain text');
      out.push(oneLine);
      return;
    }
    out.push(`${'#'.repeat(level)} ${oneLine}`);
    return;
  }

  if (marker !== undefined) {
    // A marked paragraph with nothing in it is a BLANK LINE inside a text box,
    // which is how a deck spaces its bullets out; it is not an item of the
    // list, and drawn as one it is a bullet standing alone against no words.
    // The list itself carries on around it.
    if (!empty) emitListItem(out, resolved, marker, inline, ctx);
    return;
  }

  // An empty paragraph is a line on a page and nothing at all in a flow of
  // blocks — the blank line between blocks is markdown's own separator. It
  // does not close a list either: contributing nothing, it cannot come between
  // two items, and a deck that spaces its bullets with blank lines would come
  // out as one list per bullet.
  if (empty) return;
  ctx.list.length = 0;
  // The block-start guard is for block context; a cell has none to open.
  out.push(ctx.inCell ? inline : guardBlockStart(inline));
}

/**
 * True when a paragraph's rendered text says nothing at all. Whitespace counts
 * as nothing, and so do the ZERO-WIDTH characters — the `.pptx` reader marks a
 * slide boundary with a U+200B paragraph carrying the page break, and a run of
 * them down the left of a deck is a column of empty lines, not content.
 */
function isBlank(text: string): boolean {
  // U+200B..U+200D zero-width space/non-joiner/joiner, U+FEFF zero-width no-break.
  return /^[\s\u200B-\u200D\uFEFF]*$/u.test(text);
}

/** Whether a page starts at this paragraph — its own break, or one in a run. */
function breaksPage(p: Paragraph, resolved: ResolvedParagraphProperties): boolean {
  return resolved.pageBreakBefore || p.runs.some((r) => r.pageBreak === true);
}

/**
 * A `---` thematic break where a page ends, when the caller asked for one.
 *
 * Never leading and never doubled: a rule before the first block would open
 * the document with a line, and two in a row say nothing the one does not.
 * Never inside a cell or a note either — those hold inline content, where the
 * three hyphens are just three hyphens.
 */
function emitRule(out: Array<string>, ctx: EmitCtx): void {
  if (ctx.pageBreaks !== 'rule' || ctx.inCell) return;
  if (out.length === 0 || out[out.length - 1] === RULE) return;
  out.push(RULE);
}

const RULE = '---';

/**
 * The marker `applyNumbering` materialized as the paragraph's leading runs
 * (`"1."`, `"•"`, or a picture bullet), stripped of the tab that follows it —
 * or `undefined` when the paragraph is not a list item. The text is taken raw:
 * it is re-rendered as markup, never emitted as content.
 */
function markerText(p: Paragraph, ctx: EmitCtx): string | undefined {
  const runs: Array<Run> = [];
  for (const run of p.runs) {
    if (run.listMarker !== true) break;
    runs.push(run);
  }
  if (runs.length === 0) return undefined;
  if (runs.some((r) => r.inlineImage !== undefined)) {
    lose(ctx, 'degraded', FEATURES.lists, 'picture bullet rendered as a plain bullet');
  }
  return runs
    .map((r) => r.text)
    .join('')
    .replaceAll('\t', ' ')
    .trim();
}

function emitListItem(
  out: Array<string>,
  resolved: ResolvedParagraphProperties,
  marker: string,
  inline: string,
  ctx: EmitCtx,
): void {
  const ref = resolved.numbering;
  const numId = ref?.numId ?? '';
  const ilvl = ref?.ilvl ?? 0;
  const stack = ctx.list;
  const wasOpen = stack.length > 0;

  while (stack.length > 0 && stack[stack.length - 1]!.ilvl > ilvl) stack.pop();
  let top = stack[stack.length - 1];
  if (!top || top.ilvl < ilvl) {
    const parent = top;
    top = {
      ilvl,
      indent: parent ? parent.indent + parent.markerWidth : 0,
      markerWidth: 2,
      counter: 0,
    };
    stack.push(top);
  }
  top.counter += 1;

  const bullet = listBullet(marker, numId, ilvl, top.counter, ctx);
  top.markerWidth = bullet.length + 1;
  const pad = ' '.repeat(top.indent);
  const cont = ' '.repeat(top.indent + top.markerWidth);
  // A soft break inside an item continues at the item's content column.
  const body = inline.replaceAll('\\\n', `\\\n${cont}`);
  const line = `${pad}${bullet} ${body}`.replace(/\s+$/, '');

  // Items of one list are a single block joined by plain newlines — a blank
  // line between them would make the list loose, wrapping every item's text in
  // its own paragraph.
  if (wasOpen && out.length > 0) out[out.length - 1] += `\n${line}`;
  else out.push(line);
}

/**
 * The markdown marker for an item: `-` for a bullet, `N.` for an ordered list.
 * Ordered lists keep their real number — the one the source's own marker states,
 * so a list starting at 5 (§17.9.28 `w:startOverride`) still starts at 5 —
 * falling back to this level's running count when the marker states no digits.
 */
function listBullet(
  marker: string,
  numId: string,
  ilvl: number,
  counter: number,
  ctx: EmitCtx,
): string {
  const level = numberingLevel(numId, ilvl, ctx.numbering);
  const ordered = level ? level.format !== 'bullet' && level.format !== 'none' : /\d/.test(marker);
  if (!ordered) return '-';
  if (level && level.format !== 'decimal' && level.format !== 'decimalZero') {
    lose(ctx, 'degraded', FEATURES.lists, `${level.format} list markers render as decimal`);
  }
  // §17.9.11 — a multi-level template ("%1.%2.") ends with THIS level's own
  // number; markdown has one number per item, so the parents' are lost.
  const digits = /(\d+)\D*$/.exec(marker);
  if (marker.replace(/\D+/g, '').length > (digits?.[1]?.length ?? 0)) {
    lose(ctx, 'degraded', FEATURES.lists, 'multi-level list marker flattened to one number');
  }
  return `${digits ? Number(digits[1]) : counter}.`;
}

function numberingLevel(
  numId: string,
  ilvl: number,
  numbering: Numbering | undefined,
): NumberingLevel | undefined {
  if (!numbering) return undefined;
  const instance = numbering.numInstances.get(numId);
  if (!instance) return undefined;
  return effectiveAbstract(numbering, instance)?.levels.get(ilvl);
}

/** Everything a paragraph says that markdown has no way to say back. */
function reportParagraphLosses(r: ResolvedParagraphProperties, ctx: EmitCtx): void {
  if (r.alignment !== 'left') {
    lose(ctx, 'dropped', FEATURES.text, 'paragraph alignment has no markdown expression');
  }
  if (r.indentLeft !== 0 || r.indentRight !== 0 || r.indentFirstLine !== 0) {
    lose(ctx, 'dropped', FEATURES.text, 'paragraph indents have no markdown expression');
  }
  if (r.pageBreakBefore && ctx.pageBreaks === 'drop') {
    lose(ctx, 'dropped', FEATURES.sections, 'page breaks have no markdown expression');
  }
}

/**
 * Backslash-escape a leading character that would otherwise open a block the
 * source never asked for — a heading, a quote, a list item, a setext rule.
 */
function guardBlockStart(text: string): string {
  return text.replace(/^(\s*)([#>+-]|\d+[.)]|={2,}$)/, '$1\\$2');
}

// ---------------------------------------------------------------------------
// Notes and comments
// ---------------------------------------------------------------------------

/**
 * Number the notes by the order their references appear in reading order
 * (§17.11: footnotes and endnotes each keep their own counter), and the review
 * comments alongside them.
 *
 * Only ids the package actually holds content for are numbered: an unmatched
 * `[^fn1]` is not a footnote reference at all in GFM — it renders as those
 * six literal characters — so a dangling reference must leave no mark.
 */
function noteNumbers(flow: FlowDoc): NoteNumbers {
  const footnotes = new Map<string, number>();
  const endnotes = new Map<string, number>();
  const comments = new Map<string, number>();
  const visitShape = (shape: ShapeBlock): void => {
    if (shape.text) visit(shape.text.content);
    for (const child of shape.children ?? []) visitShape(child.shape);
  };
  const visit = (els: ReadonlyArray<BodyElement>): void => {
    for (const el of els) {
      if (el.kind === 'paragraph') {
        for (const r of el.paragraph.runs) {
          const add = (
            id: string | undefined,
            into: Map<string, number>,
            content: ReadonlyMap<string, unknown> | undefined,
          ): void => {
            if (id === undefined || into.has(id) || !content?.has(id)) return;
            into.set(id, into.size + 1);
          };
          add(r.footnoteRef, footnotes, flow.footnotes);
          add(r.endnoteRef, endnotes, flow.endnotes);
          add(r.commentRef, comments, flow.comments);
        }
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
      } else if (el.kind === 'shape') {
        visitShape(el.shape);
      }
    }
  };
  visit(flow.body);
  return { footnotes, endnotes, comments };
}

/**
 * The GFM footnote definitions every reference above points at, in reference
 * order: the notes first, then the review comments — which markdown has no
 * concept of, and which are carried as footnotes attributed to their author.
 */
function emitNoteDefinitions(out: Array<string>, flow: FlowDoc, ctx: EmitCtx): void {
  define(out, ctx.notes.footnotes, flow.footnotes, 'fn', ctx);
  define(out, ctx.notes.endnotes, flow.endnotes, 'en', ctx);
  for (const [id, n] of sorted(ctx.notes.comments)) {
    const comment = flow.comments?.get(id);
    if (!comment) continue;
    const who = comment.author !== undefined ? `**${escapeInline(comment.author)}:** ` : '';
    out.push(`[^cm${String(n)}]: ${who}${flatten(comment.content, ctx)}`);
  }
}

function define(
  out: Array<string>,
  numbers: ReadonlyMap<string, number>,
  content: ReadonlyMap<string, ReadonlyArray<BodyElement>> | undefined,
  prefix: string,
  ctx: EmitCtx,
): void {
  for (const [id, n] of sorted(numbers)) {
    const blocks = content?.get(id);
    if (!blocks) continue;
    out.push(`[^${prefix}${String(n)}]: ${flatten(blocks, ctx)}`);
  }
}

const sorted = (m: ReadonlyMap<string, number>): Array<[string, number]> =>
  [...m.entries()].sort((a, b) => a[1] - b[1]);

/**
 * Blocks flattened to the single line a footnote definition occupies. Note
 * content is short by nature; a note that holds several paragraphs keeps them
 * all, joined by the line break the syntax allows.
 */
function flatten(blocks: ReadonlyArray<BodyElement>, ctx: EmitCtx): string {
  const wasInCell = ctx.inCell;
  const outer = ctx.list.splice(0);
  ctx.inCell = true;
  const rendered: Array<string> = [];
  for (const el of blocks) emitBlock(rendered, el, ctx);
  ctx.inCell = wasInCell;
  ctx.list.length = 0;
  ctx.list.push(...outer);
  return rendered
    .map(trimHardBreaks)
    .filter((b) => b.length > 0)
    .join('<br>')
    .replaceAll('\\\n', '<br>')
    .replaceAll('\n', '<br>');
}

// ---------------------------------------------------------------------------
// Links, bookmarks and pictures
// ---------------------------------------------------------------------------

/** A link target as written in the source: either an external URL or a bookmark. */
interface LinkTarget {
  readonly href?: string;
  readonly anchor?: string;
}

/**
 * Wrap already-rendered inline content in the link it carries. An external
 * target passes the scheme allowlist first (`core/links`) — untrusted input
 * never becomes a clickable link on a scheme markdown viewers will follow.
 */
function linkify(inner: string, target: LinkTarget, ctx: EmitCtx): string {
  if (inner.length === 0) return inner;
  if (target.href !== undefined) {
    const safe = sanitizeHref(target.href);
    if (safe === undefined) {
      lose(
        ctx,
        'degraded',
        FEATURES.hyperlinks,
        'hyperlink target with a disallowed scheme rendered as plain text',
      );
      return inner;
    }
    return `[${inner}](${destination(safe)})`;
  }
  if (target.anchor !== undefined) return `[${inner}](#${bookmarkSlug(target.anchor)})`;
  return inner;
}

/**
 * A link destination: bare when it is plain enough, and in the pointy-bracket
 * form CommonMark §6.3 provides when it holds whitespace or unbalanced
 * parentheses that would otherwise end it early.
 */
function destination(url: string): string {
  if (!/[\s()<>]/.test(url)) return url;
  return `<${url.replaceAll('<', '%3C').replaceAll('>', '%3E')}>`;
}

/**
 * The fragment an internal link points at. Markdown has no bookmark of its
 * own, so the name is slugified and planted as an inline `<a id>` on the
 * paragraph it belongs to — the same shape GFM's own heading anchors take.
 */
function bookmarkSlug(name: string): string {
  const slug = name
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, '-')
    .replace(/^-+|-+$/g, '');
  return slug.length > 0 ? slug : 'bookmark';
}

/** The bookmark names some run actually links to — the only ones worth an anchor. */
function referencedAnchors(body: ReadonlyArray<BodyElement>): ReadonlySet<string> {
  const out = new Set<string>();
  const visitShape = (shape: ShapeBlock): void => {
    if (shape.text) visit(shape.text.content);
    for (const child of shape.children ?? []) visitShape(child.shape);
  };
  const visit = (els: ReadonlyArray<BodyElement>): void => {
    for (const el of els) {
      if (el.kind === 'paragraph') {
        for (const r of el.paragraph.runs) if (r.anchor !== undefined) out.add(r.anchor);
      } else if (el.kind === 'table') {
        for (const row of el.table.rows) for (const cell of row.cells) visit(cell.content);
      } else if (el.kind === 'shape') {
        visitShape(el.shape);
      }
    }
  };
  visit(body);
  return out;
}

/** The `<a id>` markers a paragraph opens, for the bookmarks something links to. */
function bookmarkAnchors(p: Paragraph, ctx: EmitCtx): string {
  const named = (p.bookmarks ?? []).filter((b) => ctx.anchors.has(b));
  return named.map((b) => `<a id="${bookmarkSlug(b)}"></a>`).join('');
}

const MIME_BY_FORMAT = {
  jpeg: 'image/jpeg',
  png: 'image/png',
  jpeg2000: 'image/jp2',
  gif: 'image/gif',
  tiff: 'image/tiff',
  bmp: 'image/bmp',
} as const;

const EXTENSION_BY_FORMAT = {
  jpeg: 'jpg',
  png: 'png',
  jpeg2000: 'jp2',
  gif: 'gif',
  tiff: 'tif',
  bmp: 'bmp',
} as const;

/**
 * A picture as `![alt](…)`, or `''` when there is nothing to point at. The
 * destination follows {@link MarkdownWriteOptions.images}: inlined as a
 * `data:` URI by default, named under `./media/` when the caller means to
 * write the bytes out itself, or dropped.
 */
function pictureMarkdown(
  resource: ResourceId | undefined,
  altText: string | undefined,
  ctx: EmitCtx,
): string {
  if (ctx.images === 'drop') {
    lose(ctx, 'dropped', FEATURES.images, 'pictures omitted at the caller’s request');
    return '';
  }
  if (resource === undefined) return '';
  const bytes = ctx.resources.get(resource);
  if (!bytes) return '';
  const format = detectImageFormat(bytes);
  if (!format) {
    lose(ctx, 'dropped', FEATURES.images, 'picture in a format markdown viewers cannot show');
    return '';
  }
  const alt = escapeInline(altText ?? '').replaceAll('\n', ' ');
  if (ctx.images === 'link') {
    let name = ctx.mediaNames.get(resource);
    if (name === undefined) {
      name = `./media/image${String(ctx.mediaNames.size + 1)}.${EXTENSION_BY_FORMAT[format]}`;
      ctx.mediaNames.set(resource, name);
    }
    lose(
      ctx,
      'degraded',
      FEATURES.images,
      'pictures referenced under ./media — a single-file writer does not write their bytes',
    );
    return `![${alt}](${name})`;
  }
  return `![${alt}](data:${MIME_BY_FORMAT[format]};base64,${toBase64(bytes)})`;
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

/**
 * A GFM pipe table (§4.10). Markdown's table is a plain grid: it has no
 * merged cells and no headerless form, so a `w:gridSpan` fills its first
 * column and pads the rest empty, a `w:vMerge` continuation stays empty, and
 * the first row is promoted to the header the syntax requires. Each of those
 * is reported once.
 */
function emitTable(out: Array<string>, table: Table, ctx: EmitCtx): void {
  const cols = columnCount(table);
  if (cols === 0 || table.rows.length === 0) return;

  const rows = table.rows.map((row) => rowCells(row, cols, ctx));
  const header = rows[0]!;
  if (table.rows[0]!.properties.isHeader !== true) {
    lose(ctx, 'degraded', FEATURES.tables, 'first row promoted to the header GFM tables require');
  }

  const line = (cells: ReadonlyArray<string>): string => `| ${cells.join(' | ')} |`;
  const lines = [line(header), line(delimiters(table, cols))];
  for (const row of rows.slice(1)) lines.push(line(row));
  out.push(lines.join('\n'));
}

/** The widest row wins: a `w:gridSpan` may reach past the declared `w:tblGrid`. */
function columnCount(table: Table): number {
  let widest = table.grid.length;
  for (const row of table.rows) {
    let n = 0;
    for (const cell of row.cells) n += cell.properties.colSpan ?? 1;
    widest = Math.max(widest, n);
  }
  return widest;
}

/** One row's cells, spans expanded to empty columns and padded to `cols`. */
function rowCells(row: TableRow, cols: number, ctx: EmitCtx): Array<string> {
  const cells: Array<string> = [];
  for (const cell of row.cells) {
    const span = cell.properties.colSpan ?? 1;
    if (span > 1 || (cell.properties.merge !== undefined && cell.properties.merge !== 'start')) {
      lose(ctx, 'degraded', FEATURES.tables, 'merged cells flattened — markdown has no spans');
    }
    // A vertical-merge continuation repeats no text: its content belongs to
    // the `start` cell above, which already printed it.
    cells.push(
      cell.properties.merge === 'middle' || cell.properties.merge === 'end'
        ? ''
        : cellInline(cell, ctx),
    );
    for (let k = 1; k < span; k++) cells.push('');
  }
  while (cells.length < cols) cells.push('');
  return cells.slice(0, cols);
}

/**
 * The delimiter row, carrying each column's alignment when the header cell's
 * first paragraph states one — the only paragraph property a markdown table
 * can keep.
 */
function delimiters(table: Table, cols: number): Array<string> {
  const out: Array<string> = [];
  const first = table.rows[0]!;
  let col = 0;
  for (const cell of first.cells) {
    const span = cell.properties.colSpan ?? 1;
    const p = cell.content.find((el) => el.kind === 'paragraph');
    const align = p
      ? resolveParagraphProperties(p.paragraph.properties, EMPTY_STYLE_SHEET).alignment
      : 'left';
    const bar = align === 'center' ? ':---:' : align === 'right' ? '---:' : '---';
    for (let k = 0; k < span && col < cols; k++, col++) out.push(bar);
  }
  while (out.length < cols) out.push('---');
  return out.slice(0, cols);
}

/**
 * A cell's blocks flattened to one line: a table row is a single line, so
 * every break between and inside its blocks becomes a `<br>`. The cell's own
 * lists number independently of whatever list surrounds the table.
 *
 * Each block is trimmed at its ends. Word carries a cell's padding as spaces in
 * the text itself, and a pipe table already sets its own — kept, they only
 * widen the source line, since §4.10 strips a cell's outer whitespace before
 * rendering it anyway.
 */
function cellInline(cell: TableCell, ctx: EmitCtx): string {
  const outer = ctx.list.splice(0);
  const wasInCell = ctx.inCell;
  ctx.inCell = true;
  const blocks: Array<string> = [];
  for (const el of cell.content) {
    if (el.kind === 'table') {
      // A nested table has no second dimension to live in here.
      lose(ctx, 'degraded', FEATURES.tablesNested, 'nested table flattened into its cell');
      for (const row of el.table.rows) {
        for (const inner of row.cells) blocks.push(cellInline(inner, ctx));
      }
      continue;
    }
    emitBlock(blocks, el, ctx);
  }
  ctx.inCell = wasInCell;
  ctx.list.length = 0;
  ctx.list.push(...outer);
  return blocks
    .map((b) => trimHardBreaks(b).trim())
    .filter((b) => b.length > 0)
    .join('<br>')
    .replaceAll('\\\n', '<br>')
    .replaceAll('\n', '<br>');
}

// ---------------------------------------------------------------------------
// Runs
// ---------------------------------------------------------------------------

/** One span of the paragraph: either finished text, or text still to be dressed. */
type Piece = { readonly literal: string } | { readonly marks: Marks; text: string };

/**
 * The paragraph's runs as one inline string.
 *
 * Two passes, because whether a delimiter may open or close depends on the
 * characters on either side of it (§6.2), and those belong to the NEIGHBOURING
 * spans: the first pass coalesces runs into spans, the second dresses each
 * span knowing what stands beside it.
 */
function inlineRuns(runs: ReadonlyArray<Run>, p: Paragraph, ctx: EmitCtx): string {
  const pieces: Array<Piece> = [];
  // Consecutive runs that carry the same emphasis are one span: emitted
  // separately they would read `**a****b**` — legal, and unreadable.
  let pending: { marks: Marks; text: string } | undefined;
  const flush = (): void => {
    if (pending && pending.text.length > 0) pieces.push(pending);
    pending = undefined;
  };

  for (const run of runs) {
    const literal = literalRun(run, ctx);
    if (literal !== undefined) {
      flush();
      if (literal.length > 0) pieces.push({ literal });
      continue;
    }
    if (run.text.length === 0) continue;
    const resolved = resolveRunProperties(run.properties, p.properties, EMPTY_STYLE_SHEET);
    reportRunLosses(resolved, ctx);
    const marks = marksOf(resolved, run);
    if (pending && sameMarks(pending.marks, marks)) pending.text += run.text;
    else {
      flush();
      pending = { marks, text: run.text };
    }
  }
  flush();

  let out = '';
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i]!;
    if ('literal' in piece) {
      out += piece.literal;
      continue;
    }
    out += applyMarks(piece.text, piece.marks, ctx, out.slice(-1), leadOf(pieces, i + 1));
  }
  return out;
}

/**
 * The first character the next span will contribute. A dressed span always
 * begins with its delimiter or its html tag — both punctuation — so a marked
 * neighbour answers without being rendered first.
 */
function leadOf(pieces: ReadonlyArray<Piece>, from: number): string {
  const next = pieces[from];
  if (!next) return '';
  if ('literal' in next) return next.literal.slice(0, 1);
  if (wrapped(next.marks)) return '*';
  return escapeInline(next.text).slice(0, 1);
}

/** Whether a span carries anything that puts a delimiter or a tag in front of it. */
function wrapped(m: Marks): boolean {
  return (
    m.bold ||
    m.italic ||
    m.strike ||
    m.underline ||
    m.vertical !== 'baseline' ||
    m.href !== undefined ||
    m.anchor !== undefined
  );
}

/**
 * The runs that are not a span of styled text: reference markers, inline
 * objects, and the list marker the reader materialized. Returns the literal
 * they contribute (possibly `''`), or `undefined` when the run IS styled text.
 */
function literalRun(run: Run, ctx: EmitCtx): string | undefined {
  if (run.footnoteRef !== undefined || run.endnoteRef !== undefined) {
    const foot = run.footnoteRef !== undefined;
    const id = foot ? run.footnoteRef : run.endnoteRef!;
    const n = (foot ? ctx.notes.footnotes : ctx.notes.endnotes).get(id);
    // A reference whose note is not in the package has nothing to point at.
    return n === undefined ? '' : `[^${foot ? 'fn' : 'en'}${String(n)}]`;
  }
  if (run.commentRef !== undefined) {
    const n = ctx.notes.comments.get(run.commentRef);
    if (n === undefined) return '';
    lose(ctx, 'degraded', FEATURES.trackedChanges, 'review comments rendered as footnotes');
    return `[^cm${String(n)}]`;
  }
  // Inside note content this placeholder marks the note's own number, which
  // the footnote syntax prints by itself.
  if (run.noteNumber) return '';
  // The marker is re-derived from the numbering definition by the list step;
  // carried through as text it would read as literal "1." before the item.
  if (run.listMarker) return '';
  if (run.math !== undefined) {
    lose(ctx, 'dropped', FEATURES.math, 'markdown has no math notation (GFM)');
    return '';
  }
  if (run.inlineImage !== undefined) {
    return linkify(pictureMarkdown(run.inlineImage.resource, undefined, ctx), run, ctx);
  }
  if (run.columnBreak === true || (run.pageBreak === true && ctx.pageBreaks === 'drop')) {
    lose(ctx, 'dropped', FEATURES.sections, 'page breaks have no markdown expression');
  }
  return undefined;
}

/** Everything a run says that markdown has no way to say back. */
function reportRunLosses(r: ResolvedRunProperties, ctx: EmitCtx): void {
  if (r.colorHex !== '000000') {
    lose(ctx, 'dropped', FEATURES.text, 'run colour has no markdown expression');
  }
  if (r.caps || r.smallCaps) {
    lose(ctx, 'dropped', FEATURES.text, 'w:caps / w:smallCaps has no markdown expression');
  }
}

interface Marks {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly strike: boolean;
  readonly underline: boolean;
  readonly vertical: 'baseline' | 'superscript' | 'subscript';
  /** §17.16.22 — the run's external hyperlink target, as written in the rels. */
  readonly href?: string;
  /** §17.16.22 `@w:anchor` — an internal target: a bookmark in this document. */
  readonly anchor?: string;
}

function marksOf(r: ResolvedRunProperties, run: Run): Marks {
  return {
    bold: r.bold,
    italic: r.italic,
    strike: r.strike,
    underline: r.underline !== 'none',
    vertical:
      r.verticalAlign === 'superscript' || r.verticalAlign === 'subscript'
        ? r.verticalAlign
        : 'baseline',
    ...(run.href !== undefined ? { href: run.href } : {}),
    ...(run.anchor !== undefined ? { anchor: run.anchor } : {}),
  };
}

function sameMarks(a: Marks, b: Marks): boolean {
  return (
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.strike === b.strike &&
    a.underline === b.underline &&
    a.vertical === b.vertical &&
    a.href === b.href &&
    a.anchor === b.anchor
  );
}

/**
 * Escape the text, then dress it: markdown delimiters innermost, then the
 * inline-html wrappers for what markdown cannot say. CommonMark §6.2 requires
 * a delimiter to hug non-whitespace, so any leading or trailing space of the
 * span moves outside the marks.
 *
 * A HARD BREAK at either end moves out too, and it matters more than a space
 * does: a closing `**` that a line break stands in front of is not
 * right-flanking, cannot close, and prints as two asterisks. A `w:br` at the
 * end of a bold title — which every deck writes — did exactly that.
 */
const SPAN_EDGE = String.raw`(?:\\\n|\s)`;

function applyMarks(
  text: string,
  marks: Marks,
  ctx: EmitCtx,
  before: string,
  after: string,
): string {
  const escaped = escapeInline(text);
  const lead = new RegExp(`^${SPAN_EDGE}*`).exec(escaped)?.[0] ?? '';
  const trail = new RegExp(`${SPAN_EDGE}*$`).exec(escaped)?.[0] ?? '';
  const core = escaped.slice(lead.length, escaped.length - trail.length);
  if (core.length === 0) return escaped;

  // Whatever whitespace moved out now stands between the delimiter and the
  // neighbour, so it — not the neighbour — is what the delimiter sees.
  const outerBefore = lead.length > 0 ? lead[lead.length - 1]! : before;
  const outerAfter = trail.length > 0 ? trail[0]! : after;

  let out = core;
  // Innermost first, so each level tests against what the level below emitted.
  const levels: ReadonlyArray<readonly [boolean, string, string]> = [
    [marks.italic, '*', 'em'],
    [marks.bold, '**', 'strong'],
    [marks.strike, '~~', 'del'],
  ];
  for (const [on, delimiter, tag] of levels) {
    if (!on) continue;
    const flanks = canOpen(outerBefore, out[0]!) && canClose(out[out.length - 1]!, outerAfter);
    // A delimiter that cannot flank is not a delimiter — it prints as two
    // asterisks and the emphasis is lost. The html tag GFM parses says the
    // same thing and answers to no neighbour.
    out = flanks ? `${delimiter}${out}${delimiter}` : `<${tag}>${out}</${tag}>`;
  }
  // GFM parses markdown inside an inline html span, so these wrap the
  // already-escaped text rather than replacing the escaping regime.
  if (marks.underline) out = `<u>${out}</u>`;
  if (marks.vertical !== 'baseline') {
    const tag = marks.vertical === 'superscript' ? 'sup' : 'sub';
    out = `<${tag}>${out}</${tag}>`;
  }
  return `${lead}${linkify(out, marks, ctx)}${trail}`;
}

// CommonMark §2.1 counts the start and end of a line as whitespace, and its
// "punctuation" spans the Unicode punctuation AND symbol categories.
const isSpace = (c: string): boolean => c === '' || /\s/u.test(c);
const isPunct = (c: string): boolean => /[\p{P}\p{S}]/u.test(c);

/** §6.2 — the delimiter run is left-flanking, so it can open emphasis. */
function canOpen(before: string, after: string): boolean {
  return !isSpace(after) && (!isPunct(after) || isSpace(before) || isPunct(before));
}

/** §6.2 — the delimiter run is right-flanking, so it can close emphasis. */
function canClose(before: string, after: string): boolean {
  return !isSpace(before) && (!isPunct(before) || isSpace(after) || isPunct(after));
}

/**
 * Escape the inline text so nothing in it is read as markup, and map the two
 * control characters the model uses: `\t` (a `w:tab`, which has no medium
 * here) to a space, and `\n` (a `w:br` soft break) to GFM's backslash hard
 * break.
 *
 * `_` is escaped only at a word boundary — GFM does not open intraword
 * emphasis with it, so escaping every one would turn `snake_case` into noise.
 */
function escapeInline(text: string): string {
  return text
    .replace(/[\\`*[\]<>|~]/g, (c) => `\\${c}`)
    .replace(/(^|[\s\p{P}])_|_(?=[\s\p{P}]|$)/gu, (m) => m.replace('_', '\\_'))
    .replaceAll('\t', ' ')
    .replaceAll('\n', '\\\n');
}
