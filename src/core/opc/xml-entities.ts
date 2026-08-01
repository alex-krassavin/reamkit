// XML 1.0 §4.2.2 / ISO/IEC 29500-2 §10.1.1 — the internal DTD subset.
//
// OPC forbids a DTD outright, so a part carrying one is already malformed and
// Excel refuses the whole workbook. We would rather draw what can be read, but
// the reason the DTD is banned is real: `<!ENTITY a5 "&a4;&a4;&a4;&a4;">` five
// levels deep is the billion-laughs bomb, and a parser that expands it eagerly
// dies. fast-xml-parser never registers these declarations at all, which is
// safe — and leaves every `&a5;` sitting in the text as five literal characters
// the document does not contain (poc-xmlbomb-empty.xlsx drew
// "test123&a5;&a5;&a5;…" across the page where Excel shows "test123").
//
// So resolve them here, under two budgets. Each declaration is expanded once
// and capped, and the part as a whole may only grow so far; whatever exceeds
// either resolves to nothing. A bomb costs a bounded amount of work and prints
// as the empty text it would have been, and no `&name;` survives into the page.

/** Longest text any one declared entity may expand to. */
const MAX_EXPANSION = 65_536;

/** Longest total the substitutions may add to one part. */
const MAX_TOTAL = 1_048_576;

const DOCTYPE = /<!DOCTYPE\s[^[>]*(?:\[[\s\S]*?\]\s*)?>/;
const ENTITY_DECL = /<!ENTITY\s+([^\s%<>&;]+)\s+(?:"([^"]*)"|'([^']*)')\s*>/g;
const REFERENCE = /&([^\s;&<>]+);/g;

/** The five entities XML predefines; a declaration never overrides them. */
const PREDEFINED = new Set(['lt', 'gt', 'amp', 'apos', 'quot']);

/** Escape resolved text so it re-enters the document as characters, not markup. */
function escape(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Substitute every entity reference in `text` using `valueOf`, leaving the
 * predefined and numeric references for the XML parser. Returns undefined once
 * the result passes `cap` — the caller drops what it was building.
 */
function substitute(
  text: string,
  cap: number,
  valueOf: (name: string) => string,
): string | undefined {
  if (!text.includes('&')) return text.length > cap ? undefined : text;
  let out = '';
  let last = 0;
  for (const m of text.matchAll(REFERENCE)) {
    const name = m[1] ?? '';
    out += text.slice(last, m.index);
    last = m.index + m[0].length;
    // A predefined or numeric reference is the parser's business, not ours.
    out += PREDEFINED.has(name) || name.startsWith('#') ? m[0] : valueOf(name);
    if (out.length > cap) return undefined;
  }
  out += text.slice(last);
  return out.length > cap ? undefined : out;
}

/**
 * Expand each declaration once, so a name referenced ten thousand times costs
 * one expansion rather than ten thousand. A cycle, an undeclared name and a
 * value that outgrows MAX_EXPANSION all resolve to the empty string: XML says
 * such a document is in error, and drawing the reference's own spelling is the
 * one answer that is certainly wrong.
 */
function expander(declared: ReadonlyMap<string, string>): (name: string) => string {
  const done = new Map<string, string>();
  const active = new Set<string>();
  const valueOf = (name: string): string => {
    const memo = done.get(name);
    if (memo !== undefined) return memo;
    // A name still being expanded refers to itself, directly or through others.
    if (active.has(name)) return '';
    const raw = declared.get(name);
    if (raw === undefined) return '';
    active.add(name);
    const out = substitute(raw, MAX_EXPANSION, valueOf) ?? '';
    active.delete(name);
    done.set(name, out);
    return out;
  };
  return valueOf;
}

/**
 * Resolve the entities an internal DTD subset declares and remove the subset,
 * so what reaches the XML parser is a well-formed part with no `&name;` left
 * to leak. Markup inside an entity value is escaped to text rather than
 * spliced in — an entity that opens a tag is a document we will not build.
 *
 * Text with no `<!DOCTYPE` is returned unchanged, which is every real file.
 *
 * @param xml The decoded part, as read from the package.
 * @returns The part with its internal subset resolved and stripped.
 */
export function resolveInternalEntities(xml: string): string {
  if (!xml.includes('<!DOCTYPE')) return xml;
  const doctype = DOCTYPE.exec(xml);
  if (!doctype) return xml;

  const declared = new Map<string, string>();
  for (const decl of doctype[0].matchAll(ENTITY_DECL)) {
    const name = decl[1];
    // A parameter entity (`<!ENTITY % name …>`) declares markup, not text.
    if (name === undefined || name.startsWith('%') || PREDEFINED.has(name)) continue;
    // First declaration wins (XML 1.0 §4.2).
    if (!declared.has(name)) declared.set(name, decl[2] ?? decl[3] ?? '');
  }

  const body = xml.slice(0, doctype.index) + xml.slice(doctype.index + doctype[0].length);
  if (!body.includes('&')) return body;
  const valueOf = expander(declared);
  // Past MAX_TOTAL the part is no longer the document anyone wrote; keep the
  // text and drop every remaining reference rather than build it.
  return (
    substitute(body, MAX_TOTAL, (name) => escape(valueOf(name))) ??
    substitute(body, MAX_TOTAL, () => '') ??
    body
  );
}
