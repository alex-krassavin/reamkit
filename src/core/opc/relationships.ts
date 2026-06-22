// ECMA-376 Part 2 §9.3 — Relationships Part.
// Parses XML of the form:
//   <Relationships xmlns="…/package/2006/relationships">
//     <Relationship Id="…" Type="…" Target="…" TargetMode="External|Internal"/>
//   </Relationships>
// Some producers put the relationships namespace on a PREFIX
// (<ns0:Relationships><ns0:Relationship/>) instead of the default; removeNSPrefix
// makes the element lookups match regardless (corpus: 58760.xlsx read 0 sheets).

import { XMLParser } from 'fast-xml-parser';

/**
 * One `<Relationship>` (ECMA-376 Part 2 §9.3): its id, Type URI, target, and
 * whether the target is package-`Internal` or `External` (defaults to Internal
 * when the attribute is absent).
 */
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
  // Tolerate a namespace prefix on the elements (<ns0:Relationship>); the Id/
  // Type/Target attributes themselves are always unprefixed.
  removeNSPrefix: true,
  isArray: (tagName) => tagName === 'Relationship',
});

/**
 * Parse a relationships part (`_rels/*.rels`) into {@link Relationship}s.
 * Tolerates a namespace prefix on the elements; returns `[]` when there are no
 * relationships.
 *
 * @throws Error when a `<Relationship>` lacks a required `Id`/`Type`/`Target`.
 */
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
