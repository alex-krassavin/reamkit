// Cell comments / notes (E-SHEET W7). Two parts carry them:
//   • legacy  xl/comments#.xml  (§18.7) — <authors> + <commentList>, each
//     <comment ref authorId> with a rich <text>; the author is an index into
//     <authors>. The accompanying VML drawing (the yellow note box) is ignored —
//     only the text + author are surfaced.
//   • modern  xl/threadedComments/threadedComment#.xml (a conversation) — each
//     <threadedComment ref personId> with a plain <text>; the author is a
//     person id resolved through xl/persons/person.xml.
// Both reduce to the same format-neutral SheetComment list (cell ref, author,
// text), which the projection lists in a "Comments" section after the grid.

import { XMLParser } from 'fast-xml-parser';

import type { SheetComment } from '@/core/ir/sheet';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  // Tolerate explicit namespace prefixes (some producers prefix every element).
  removeNSPrefix: true,
});

// §18.7.3/§18.7.6 xl/comments#.xml — legacy notes. authorId indexes <authors>;
// the text is rich (<r><t>) so it flattens like a shared string. Excel stores a
// bold "Author:\n" run at the head of the text — stripped here so the resolved
// author is not shown twice.
export function parseLegacyComments(data: Uint8Array): Array<SheetComment> {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const root = asObject(tree['comments']);
  if (!root) return [];
  const authors = toArray(asObject(root['authors'])?.['author']).map((a) => textOf(a));
  const out: Array<SheetComment> = [];
  for (const c of toArray(asObject(root['commentList'])?.['comment'])) {
    const obj = asObject(c);
    if (!obj) continue;
    const ref = strAttr(obj, 'ref');
    if (!ref) continue;
    const authorIdx = Number(strAttr(obj, 'authorId'));
    const author = Number.isInteger(authorIdx) ? authors[authorIdx] : undefined;
    const text = stripAuthorPrefix(richText(obj['text']), author);
    if (text.length === 0 && !author) continue;
    out.push({ ref, ...(author ? { author } : {}), text, threaded: false });
  }
  return out;
}

// xl/persons/person.xml — id → displayName, for resolving threaded authors.
export function parsePersons(data: Uint8Array): Map<string, string> {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const root = asObject(tree['personList']);
  const map = new Map<string, string>();
  if (!root) return map;
  for (const p of toArray(root['person'])) {
    const obj = asObject(p);
    const id = obj ? strAttr(obj, 'id') : undefined;
    const name = obj ? strAttr(obj, 'displayName') : undefined;
    if (id && name) map.set(id, name);
  }
  return map;
}

// xl/threadedComments/threadedComment#.xml — a conversation. Each entry is one
// message (replies share the ref); personId resolves through `persons`. They are
// returned in document order so a thread reads top to bottom.
export function parseThreadedComments(
  data: Uint8Array,
  persons: ReadonlyMap<string, string>,
): Array<SheetComment> {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const root = asObject(tree['ThreadedComments']);
  if (!root) return [];
  const out: Array<SheetComment> = [];
  for (const c of toArray(root['threadedComment'])) {
    const obj = asObject(c);
    if (!obj) continue;
    const ref = strAttr(obj, 'ref');
    if (!ref) continue;
    const personId = strAttr(obj, 'personId');
    const author = personId ? persons.get(personId) : undefined;
    const text = textOf(obj['text']);
    if (text.length === 0) continue;
    out.push({ ref, ...(author ? { author } : {}), text, threaded: true });
  }
  return out;
}

// A legacy comment's <text> is rich (<r><t>…</t></r>…) or a bare <t>; flatten it.
function richText(node: unknown): string {
  const obj = asObject(node);
  if (!obj) return textOf(node);
  const direct = textOf(obj['t']);
  if (direct) return direct;
  return toArray(obj['r'])
    .map((r) => textOf(asObject(r)?.['t']))
    .join('');
}

// Drop a leading "Author:" / "Author:\n" the legacy producer prepends, so the
// resolved author (shown separately) is not duplicated in the body.
function stripAuthorPrefix(text: string, author: string | undefined): string {
  if (!author) return text;
  const prefix = `${author}:`;
  if (text.startsWith(prefix)) return text.slice(prefix.length).replace(/^\s+/, '');
  return text;
}

function asObject(v: unknown): Record<string, unknown> | undefined {
  return v && typeof v === 'object' && !Array.isArray(v)
    ? (v as Record<string, unknown>)
    : undefined;
}

function toArray(v: unknown): ReadonlyArray<unknown> {
  return Array.isArray(v) ? v : v !== undefined ? [v] : [];
}

function strAttr(obj: Record<string, unknown>, key: string): string | undefined {
  const v = obj[`@_${key}`];
  return typeof v === 'string' ? v : undefined;
}

function textOf(node: unknown): string {
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  const obj = asObject(node);
  if (!obj) return '';
  const inner = obj['#text'];
  if (typeof inner === 'string') return inner;
  if (typeof inner === 'number') return String(inner);
  return '';
}
