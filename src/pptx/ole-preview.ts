// The picture an embedded OLE object shows on the slide.
//
// §19.3.2.4 — a `p:oleObj` is a whole other document sitting on a slide: a
// worksheet, an equation, another deck. What the SLIDE shows of it is a
// snapshot the producer wrote beside it, and the deck names that snapshot in
// one of two ways.
//
// The modern one puts a `p:pic` inside the `p:oleObj`, which is an ordinary
// picture and needs nothing from here. The legacy one — what PowerPoint 2003
// wrote and what half the corpus carries — puts a shape in the slide's VML
// drawing part instead: `<v:shape id="_x0000_s681987"><v:imagedata o:relid=…>`,
// where the id is the `@spid` on the `p:oleObj` and the relationship is the
// VML part's own, not the slide's. The picture behind it is usually a metafile.
//
// 45541_Footer's eighth slide is one object and nothing else — an embedded deck
// covering the whole slide — so unread, the slide is blank paper.

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
});

// VML shape ids are written `_x0000_s681987` in the drawing and referenced the
// same way from `@spid`, but a producer may drop the prefix on one side.
const SPID_PREFIX = /^_x0000_s/u;

/**
 * Shape id → the relationship id of its `v:imagedata`, for every shape in a
 * legacy VML drawing part that has one.
 *
 * @param data The raw `vmlDrawing#.vml` bytes.
 * @returns The map, keyed by shape id with the `_x0000_s` prefix stripped.
 */
export function parseVmlImageRels(data: Uint8Array): Map<string, string> {
  const out = new Map<string, string>();
  const doc = parser.parse(decoder.decode(data)) as unknown;
  walk(doc, out);
  return out;
}

/** Strip the `_x0000_s` prefix a VML shape id carries, so both spellings meet. */
export function normalizeSpid(id: string): string {
  return id.replace(SPID_PREFIX, '');
}

// The part is small and its shapes may sit under a group, so this walks the
// parsed tree rather than assuming a depth.
function walk(node: unknown, out: Map<string, string>): void {
  if (node === null || typeof node !== 'object') return;
  if (Array.isArray(node)) {
    for (const item of node) walk(item, out);
    return;
  }
  const rec = node as Record<string, unknown>;
  const shape = rec['shape'];
  if (shape !== undefined) {
    for (const s of Array.isArray(shape) ? shape : [shape]) collect(s, out);
  }
  for (const value of Object.values(rec)) walk(value, out);
}

function collect(shape: unknown, out: Map<string, string>): void {
  if (shape === null || typeof shape !== 'object') return;
  const rec = shape as Record<string, unknown>;
  const id = rec['@_id'];
  const imagedata = rec['imagedata'];
  if (typeof id !== 'string' || imagedata === undefined) return;
  const first = Array.isArray(imagedata) ? imagedata[0] : imagedata;
  if (first === null || typeof first !== 'object') return;
  const img = first as Record<string, unknown>;
  // `o:relid` is what PowerPoint writes; `r:id` is the spec spelling, and both
  // arrive here without their prefix (removeNSPrefix).
  const rel = img['@_relid'] ?? img['@_id'];
  if (typeof rel === 'string') out.set(normalizeSpid(id), rel);
}
