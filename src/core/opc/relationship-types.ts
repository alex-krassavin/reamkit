// ECMA-376 Part 2 — well-known Relationship Type URIs.

/** Relationship Type URI of the package's main document part (officeDocument). */
export const REL_OFFICE_DOCUMENT =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

/** Relationship Type URI of the core-properties part (`docProps/core.xml`). */
export const REL_CORE_PROPERTIES =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

/** Relationship Type URI of the extended-properties part (`docProps/app.xml`). */
export const REL_EXTENDED_PROPERTIES =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';

// ISO 29500 Strict uses purl.oclc.org relationship URIs for the same
// relationship names Transitional spells under schemas.openxmlformats.org.
// Compare by name against both bases.
const OOXML_REL_BASES = [
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/',
  'http://purl.oclc.org/ooxml/officeDocument/relationships/',
] as const;

/**
 * Whether a relationship `type` URI names the OOXML relationship `name` (e.g.
 * `'officeDocument'`), matching against both the Transitional
 * (`schemas.openxmlformats.org`) and Strict (`purl.oclc.org`) base URIs.
 */
export function isOoxmlRel(type: string, name: string): boolean {
  return OOXML_REL_BASES.some((base) => type === base + name);
}
