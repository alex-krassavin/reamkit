// ECMA-376 Part 2 §11.1 — docProps/core.xml uses the Dublin Core / OPC core
// vocabulary to describe package-level metadata. Both docx and xlsx packages
// use the same schema, so we share the parser.

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

export interface CoreProperties {
  readonly title?: string;
  readonly subject?: string;
  readonly creator?: string; // dc:creator → PDF "Author"
  readonly keywords?: string;
  readonly description?: string;
  readonly lastModifiedBy?: string;
  readonly created?: Date; // dcterms:created
  readonly modified?: Date; // dcterms:modified
}

export function parseCoreProperties(data: Uint8Array): CoreProperties {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  // The root is <cp:coreProperties> but fast-xml-parser strips namespaces by
  // default if `removeNSPrefix` is false (it's false here), so we look up the
  // tag in its prefixed form.
  const root = (tree['cp:coreProperties'] ?? tree['coreProperties']) as
    | Record<string, unknown>
    | undefined;
  if (!root) return {};

  const out: { -readonly [K in keyof CoreProperties]: CoreProperties[K] } = {};
  const title = textValue(root['dc:title'] ?? root['title']);
  if (title) out.title = title;
  const subject = textValue(root['dc:subject'] ?? root['subject']);
  if (subject) out.subject = subject;
  const creator = textValue(root['dc:creator'] ?? root['creator']);
  if (creator) out.creator = creator;
  const keywords = textValue(root['cp:keywords'] ?? root['keywords']);
  if (keywords) out.keywords = keywords;
  const description = textValue(root['dc:description'] ?? root['description']);
  if (description) out.description = description;
  const lastModifiedBy = textValue(root['cp:lastModifiedBy'] ?? root['lastModifiedBy']);
  if (lastModifiedBy) out.lastModifiedBy = lastModifiedBy;

  const created = parseIsoDate(textValue(root['dcterms:created'] ?? root['created']));
  if (created) out.created = created;
  const modified = parseIsoDate(textValue(root['dcterms:modified'] ?? root['modified']));
  if (modified) out.modified = modified;

  return out;
}

function textValue(node: unknown): string | undefined {
  if (node === undefined || node === null) return undefined;
  if (typeof node === 'string') return node.length > 0 ? node : undefined;
  if (typeof node === 'number') return String(node);
  if (typeof node === 'object') {
    const inner = (node as Record<string, unknown>)['#text'];
    if (typeof inner === 'string' && inner.length > 0) return inner;
    if (typeof inner === 'number') return String(inner);
  }
  return undefined;
}

function parseIsoDate(value: string | undefined): Date | undefined {
  if (!value) return undefined;
  const d = new Date(value);
  return Number.isFinite(d.getTime()) ? d : undefined;
}
