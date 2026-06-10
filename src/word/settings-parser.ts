// ECMA-376 Part 1 §17.15 — word/settings.xml. We only extract the few flags
// the renderer needs; everything else (compat, autoSpaceDE, ...) is ignored.

import { XMLParser } from 'fast-xml-parser';

import type { PoNode } from '@/word/po-helpers';
import { poAttr, poChildren, poFindByPath, poIs } from '@/word/po-helpers';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  preserveOrder: true,
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: false,
});

export interface DocumentSettings {
  // ECMA-376 §17.15.1.36 — w:evenAndOddHeaders. When true, even-numbered
  // pages use the 'even' header/footer references instead of 'default'.
  readonly evenAndOddHeaders: boolean;
}

export const EMPTY_SETTINGS: DocumentSettings = {
  evenAndOddHeaders: false,
};

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
