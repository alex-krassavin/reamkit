// ECMA-376 Part 1 §18.4.8 — Shared Strings Table.
// xl/sharedStrings.xml stores deduplicated cell strings. Cells with type
// "s" reference a string by index into this table.

import { XMLParser } from 'fast-xml-parser';

import type { SheetRichRun } from '@/core/spreadsheet-model';

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

/**
 * The flattened text of every shared string (a cell with `t="s"` indexes into it),
 * plus — for the strings that carry per-run formatting — the parsed rich runs
 * (E-SHEET W6). `runs[i]` is undefined unless si `i` has ≥ 2 runs OR a single run
 * that actually sets formatting, so the common plain-text case stays cheap.
 */
export interface SharedStrings {
  /** The flattened text of each shared string, parallel to the `<si>` order. */
  readonly texts: ReadonlyArray<string>;
  /** Per-index rich runs, parallel to {@link SharedStrings.texts}; `undefined` for plain strings. */
  readonly runs: ReadonlyArray<ReadonlyArray<SheetRichRun> | undefined>;
}

/**
 * Parse `xl/sharedStrings.xml` (§18.4.8) into the deduplicated string table plus,
 * for rich strings, their per-run formatting. Each string is capped to the Excel
 * cell limit (32 767 chars).
 */
export function parseSharedStrings(data: Uint8Array): SharedStrings {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Record<string, unknown>;
  const sst = tree['sst'];
  if (!sst || typeof sst !== 'object') return { texts: [], runs: [] };
  const siRaw = (sst as Record<string, unknown>)['si'];
  const items = Array.isArray(siRaw) ? siRaw : siRaw !== undefined ? [siRaw] : [];
  const texts: Array<string> = [];
  const runs: Array<ReadonlyArray<SheetRichRun> | undefined> = [];
  for (const item of items) {
    const parsed = extractSi(item);
    texts.push(capLen(parsed.text));
    runs.push(parsed.runs);
  }
  return { texts, runs };
}

interface ParsedSi {
  readonly text: string;
  // Present only when the <si> is rich (≥ 1 <r> with a formatting <rPr>).
  readonly runs: ReadonlyArray<SheetRichRun> | undefined;
}

function extractSi(si: unknown): ParsedSi {
  if (typeof si === 'string') return { text: si, runs: undefined };
  if (typeof si === 'number') return { text: String(si), runs: undefined };
  if (!si || typeof si !== 'object') return { text: '', runs: undefined };
  const obj = si as Record<string, unknown>;
  // A plain <si><t>…</t></si> — no per-run formatting.
  const direct = textOf(obj['t']);
  if (direct) return { text: direct, runs: undefined };
  // Rich text: <si><r><rPr>…</rPr><t>…</t></r>…</si>.
  const r = obj['r'];
  const rs = Array.isArray(r) ? r : r && typeof r === 'object' ? [r] : [];
  if (rs.length === 0) return { text: '', runs: undefined };
  const runs: Array<SheetRichRun> = [];
  let anyFormatting = false;
  for (const rr of rs) {
    const ro = rr as Record<string, unknown> | undefined;
    const text = textOf(ro?.['t']);
    if (text.length === 0) continue;
    const run = richRun(text, ro?.['rPr']);
    if (Object.keys(run).length > 1) anyFormatting = true; // more than just `text`
    runs.push(run);
  }
  const text = runs.map((x) => x.text).join('');
  // Only surface runs when at least one carries formatting — otherwise the plain
  // text path is identical and cheaper (and keeps existing snapshots unchanged).
  return { text, runs: anyFormatting && runs.length > 0 ? runs : undefined };
}

// §18.4.7 <rPr> — a run's own font properties (NOT a cellXf index): bold,
// italic, underline, colour (rgb), size and vertical alignment.
function richRun(text: string, rPr: unknown): SheetRichRun {
  const p = rPr && typeof rPr === 'object' ? (rPr as Record<string, unknown>) : undefined;
  if (!p) return { text };
  const out: { -readonly [K in keyof SheetRichRun]: SheetRichRun[K] } = { text };
  if (has(p, 'b')) out.bold = true;
  if (has(p, 'i')) out.italic = true;
  if (has(p, 'u')) out.underline = true;
  const color = colorRgb(p['color']);
  if (color) out.colorHex = color;
  const sz = attrNum(p['sz'], 'val');
  if (sz !== undefined) out.sizePt = sz;
  const vert = attrStr(p['vertAlign'], 'val');
  if (vert === 'superscript' || vert === 'subscript') out.vertAlign = vert;
  return out;
}

// A boolean rPr toggle (<b/>, <i/>, <u/>) is present unless it carries val="0".
function has(p: Record<string, unknown>, key: string): boolean {
  if (!(key in p)) return false;
  const node = p[key];
  const val = attrStr(node, 'val');
  return val !== '0' && val !== 'false';
}

function attrStr(node: unknown, key: string): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const v = (node as Record<string, unknown>)[`@_${key}`];
  return typeof v === 'string' ? v : undefined;
}

function attrNum(node: unknown, key: string): number | undefined {
  const s = attrStr(node, key);
  if (s === undefined) return undefined;
  const n = Number(s);
  return Number.isFinite(n) ? n : undefined;
}

// <color rgb="FFRRGGBB"/> → "RRGGBB" (ARGB alpha stripped). Theme/indexed colours
// (no rgb) are not resolved here — a run without a usable rgb keeps the cell font.
function colorRgb(node: unknown): string | undefined {
  const rgb = attrStr(node, 'rgb');
  if (!rgb) return undefined;
  const hex = rgb.length === 8 ? rgb.slice(2) : rgb;
  return /^[0-9A-Fa-f]{6}$/.test(hex) ? hex.toUpperCase() : undefined;
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
