// ECMA-376 Part 1 §17.15 — word/settings.xml. We only extract the few flags
// the renderer needs; everything else (compat, autoSpaceDE, ...) is ignored.

import { XMLParser } from 'fast-xml-parser';

import type { PoNode } from '@/core/po-helpers';
import { poAttr, poChildren, poFindByPath, poIs } from '@/core/po-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  // §4.1 of XML 1.0: a numeric character reference is not an entity — `&#10;`
  // IS a line feed and every parser must decode it. fast-xml-parser gates that
  // on `htmlEntities`, which defaults to false, so `&#10;` reached the page as
  // five literal characters (formats.xlsx writes "Hello,&#10;Calc!"). Named
  // HTML entities come along with the switch; in XML they are undefined anyway,
  // and reading `&nbsp;` as a space beats drawing it. Nested DOCTYPE entities
  // stay unexpanded either way — the parser never registers them (54764-2.xlsx).
  htmlEntities: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

/** The handful of `word/settings.xml` flags the renderer consumes. */
export interface DocumentSettings {
  /**
   * ECMA-376 §17.15.1.36 — `w:evenAndOddHeaders`. When `true`, even-numbered
   * pages use the `'even'` header/footer references instead of `'default'`.
   */
  readonly evenAndOddHeaders: boolean;
}

/** The all-defaults {@link DocumentSettings}, returned when no `w:settings` root is found. */
export const EMPTY_SETTINGS: DocumentSettings = {
  evenAndOddHeaders: false,
};

/**
 * Parse `word/settings.xml` (ECMA-376 §17.15), extracting only the flags the
 * renderer needs; everything else (compat, autoSpaceDE, …) is ignored.
 *
 * @param data The raw `word/settings.xml` bytes.
 * @returns The extracted {@link DocumentSettings}, or {@link EMPTY_SETTINGS} when the root is absent.
 */
export function parseSettings(data: Uint8Array): DocumentSettings {
  const xml = decoder.decode(data);
  const tree = parser.parse(xml) as Array<PoNode>;
  const settings = poFindByPath(tree, ['w:settings']);
  if (!settings) return EMPTY_SETTINGS;

  let evenAndOddHeaders = false;
  for (const child of poChildren(settings)) {
    if (poIs(child, 'w:evenAndOddHeaders')) {
      const val = poAttr(child, 'val');
      evenAndOddHeaders = val === undefined || val === '' || (val !== '0' && val !== 'false');
    }
  }
  return { evenAndOddHeaders };
}
