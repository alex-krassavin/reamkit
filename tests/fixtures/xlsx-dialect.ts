// Rewrite a .xlsx into an equivalent but differently-spelled dialect.
//
// build-xlsx.ts emits one spelling of SpreadsheetML: the default namespace,
// `rIdN` relationship ids, no markup-compatibility attributes. That is what our
// own writer produces, so a test built on it can only ever prove the parser
// reads our writer. Real producers spell the same document differently, and
// every difference is somewhere a parser can be accidentally literal.
//
// This transform re-spells a package without changing what it means. A parse of
// the original and a parse of the transform must agree exactly — which is the
// property worth asserting, and one that holds for every fixture we have rather
// than the handful of real documents we managed to adopt.

import { unzipSync, zipSync } from 'fflate';

export interface DialectOptions {
  /**
   * Namespace prefix for SpreadsheetML elements: `<worksheet>` → `<x:worksheet>`
   * with the default namespace moved to `xmlns:x`. ECMA-376 permits any prefix;
   * Haansoft HCell and the producer behind tdf122336.xlsx both use `x:`.
   */
  readonly nsPrefix?: string;
  /**
   * Re-spell relationship ids in the GUID shape Open XML SDK writers emit
   * (`R30efda98d0b449c4`) instead of the sequential `rId1`. Nothing in the spec
   * makes `rId` meaningful — it is an opaque token — so a parser that pattern-
   * matches on it is wrong in a way only a real file will reveal.
   */
  readonly guidRelIds?: boolean;
  /**
   * Add `mc:Ignorable` plus the extension namespaces Excel 2010+ stamps on
   * every worksheet it saves.
   */
  readonly mcIgnorable?: boolean;
}

const SML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const decoder = new TextDecoder('utf-8');
const encoder = new TextEncoder();

// An UNPREFIXED element name: `<tag`, `</tag`. The lookahead rejects `<mc:Alt`
// because ':' is not a name character here, so already-prefixed elements and
// XML declarations / comments are left alone.
const ELEMENT = /<(\/?)([A-Za-z_][\w.-]*)(?=[\s/>])/g;

/** Deterministic GUID-shaped id for `rIdN`, stable across runs. */
function guidRelId(n: number): string {
  const seed = (n * 2654435761) >>> 0;
  return 'R' + seed.toString(16).padStart(8, '0') + 'a4b94c' + (n % 10);
}

function isSpreadsheetMlPart(path: string, xml: string): boolean {
  return path.endsWith('.xml') && !path.endsWith('.rels') && xml.includes(SML_NS);
}

function applyPrefix(xml: string, prefix: string): string {
  return xml
    .replace(ELEMENT, (_m, slash: string, name: string) => `<${slash}${prefix}:${name}`)
    .replace(`xmlns="${SML_NS}"`, `xmlns:${prefix}="${SML_NS}"`);
}

function applyMcIgnorable(xml: string, prefix: string): string {
  const root = prefix ? `<${prefix}:worksheet` : '<worksheet';
  const at = xml.indexOf(root);
  if (at < 0) return xml;
  const insertAt = at + root.length;
  return (
    xml.slice(0, insertAt) +
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
    ' mc:Ignorable="x14ac"' +
    ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' +
    xml.slice(insertAt)
  );
}

/**
 * Re-spell an `.xlsx` package per `options`, leaving its meaning untouched.
 *
 * @param xlsx    The package bytes.
 * @param options Which spellings to apply.
 * @returns The re-spelled package.
 */
export function toDialect(xlsx: Uint8Array, options: DialectOptions): Uint8Array {
  const entries = unzipSync(xlsx);
  const out: Record<string, Uint8Array> = {};
  for (const [path, data] of Object.entries(entries)) {
    if (!path.endsWith('.xml') && !path.endsWith('.rels')) {
      out[path] = data;
      continue;
    }
    let xml = decoder.decode(data);
    if (options.nsPrefix && isSpreadsheetMlPart(path, xml)) {
      xml = applyPrefix(xml, options.nsPrefix);
    }
    if (options.mcIgnorable && path.includes('worksheets/')) {
      xml = applyMcIgnorable(xml, options.nsPrefix ?? '');
    }
    if (options.guidRelIds) {
      // Covers both the declaration (Id="rId1" in .rels) and every reference
      // (r:id="rId1"), since they share the token.
      xml = xml.replace(/rId(\d+)/g, (_m, n: string) => guidRelId(Number(n)));
    }
    out[path] = encoder.encode(xml);
  }
  return zipSync(out);
}
