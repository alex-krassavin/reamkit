// ECMA-376 Part 1 §18.4.8 — Shared Strings Table.
// xl/sharedStrings.xml stores deduplicated cell strings. Cells with type
// "s" reference a string by index into this table.

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
  // Tolerate an explicit `x:` namespace prefix (<x:sst>, <x:si>, <x:t>) used by
  // some producers — see workbook-parser.ts.
  removeNSPrefix: true,
});

// ECMA-376 Part 1 §3.2.2 / Excel limit: a cell holds at most 32 767 characters.
// Strings longer than this are invalid; capping here is both spec-correct and a
// DoS guard — a crafted ~1 MB shared string referenced by thousands of cells
// would otherwise be shaped/measured per cell and hang the renderer.
const MAX_CELL_CHARS = 32_767;

function capLen(s: string): string {
  return s.length > MAX_CELL_CHARS ? s.slice(0, MAX_CELL_CHARS) : s;
}

export function parseSharedStrings(data: Uint8Array): Array<string> {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const sst = tree['sst'];
  if (!sst || typeof sst !== 'object') return [];
  const siRaw = (sst as Record<string, unknown>)['si'];
  const items = Array.isArray(siRaw) ? siRaw : siRaw !== undefined ? [siRaw] : [];
  return items.map((item) => capLen(extractSiText(item)));
}

function extractSiText(si: unknown): string {
  if (typeof si === 'string') return si;
  if (typeof si === 'number') return String(si);
  if (!si || typeof si !== 'object') return '';
  const obj = si as Record<string, unknown>;
  const t = obj['t'];
  const direct = textOf(t);
  if (direct) return direct;
  // Rich text fallback: <si><r><t>...</t></r><r><t>...</t></r></si>
  const r = obj['r'];
  if (Array.isArray(r)) {
    return r.map((rr) => textOf((rr as Record<string, unknown> | undefined)?.['t'])).join('');
  }
  if (r && typeof r === 'object') {
    return textOf((r as Record<string, unknown>)['t']);
  }
  return '';
}

function textOf(node: unknown): string {
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  if (!node || typeof node !== 'object') return '';
  const obj = node as Record<string, unknown>;
  const inner = obj['#text'];
  if (typeof inner === 'string') return inner;
  if (typeof inner === 'number') return String(inner);
  return '';
}
