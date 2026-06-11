// ECMA-376 Part 2 §9.3 — Relationships Part.
// Parses XML of the form:
//   <Relationships xmlns="…/package/2006/relationships">
//     <Relationship Id="…" Type="…" Target="…" TargetMode="External|Internal"/>
//   </Relationships>

import { XMLParser } from 'fast-xml-parser';

export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode: 'Internal' | 'External';
}

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  trimValues: false,
  isArray: (tagName) => tagName === 'Relationship',
});

export function parseRelationships(data: Uint8Array): Array<Relationship> {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as {
    Relationships?: { Relationship?: ReadonlyArray<Record<string, string>> };
  };
  const list = tree.Relationships?.Relationship;
  if (!list) return [];
  return list.map((r) => ({
    id: requireAttr(r, 'Id'),
    type: requireAttr(r, 'Type'),
    target: requireAttr(r, 'Target'),
    targetMode: r['@_TargetMode'] === 'External' ? 'External' : 'Internal',
  }));
}

function requireAttr(r: Record<string, string>, name: string): string {
  const v = r[`@_${name}`];
  if (v === undefined) {
    throw new Error(`Relationship missing required attribute: ${name}`);
  }
  return v;
}
