// Rich values (`xl/richData/*`) — the value a cell REALLY holds when `<v>` holds
// only a legacy stand-in.
//
// A cell whose value Excel cannot express in the 2006 schema writes the error a
// pre-feature reader should show — almost always `#VALUE!` — and points at the
// truth with §18.3.1.4's `vm` attribute:
//
//   <c r="D2" t="e" cm="1" vm="1"><f t="array" ref="D2">UNIQUE(…)</f><v>#VALUE!</v></c>
//
// `vm` is a 1-based index into `xl/metadata.xml`'s `<valueMetadata>`. Each entry
// is `<bk><rc t="T" v="V"/></bk>`, where `t` is a 1-based index into
// `<metadataTypes>` and `v` a 0-based index into that type's `<futureMetadata>`
// blocks. For the `XLRICHVALUE` type the block carries `<xlrd:rvb i="N"/>`, and
// N indexes `xl/richData/rdrichvalue.xml`. Each `<rv s="S">` there lists its
// values in the key order that structure S declares in
// `rdrichvaluestructure.xml`.
//
// This module resolves that chain for ONE structure — `_error`, whose whole
// purpose is to say what the legacy `<v>` could not. Everything else (linked
// data types, images in cells) resolves to nothing and the stand-in stands.

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  trimValues: true,
  removeNSPrefix: true,
});

/** The metadata type whose future blocks point into the rich-value table. */
const RICH_VALUE_TYPE = 'XLRICHVALUE';

/**
 * The `_error` structure's `errorType` codes, and only the ones a real document
 * has corroborated. `Spill.xlsx` writes 8 alongside `subType="1"` and
 * `rwOffset="3"`, and the cell three rows down is the one holding the text that
 * blocks the spill — which names the error beyond doubt. The rest of the
 * enumeration is not written down anywhere this code can check, and a wrong
 * error is worse than the stand-in the file already provides, so an unmapped
 * code resolves to nothing.
 */
const ERROR_TYPES = new Map<number, string>([[8, '#SPILL!']]);

/** One entry of `rdrichvaluestructure.xml`: a type name and its key order. */
interface RichValueStructure {
  readonly type: string;
  readonly keys: ReadonlyArray<string>;
}

type XmlNode = Record<string, unknown>;

function nodes(parent: unknown, name: string): Array<XmlNode> {
  if (typeof parent !== 'object' || parent === null) return [];
  const child = (parent as XmlNode)[name];
  if (child === undefined) return [];
  const list = Array.isArray(child) ? child : [child];
  return list.filter((n): n is XmlNode => typeof n === 'object' && n !== null);
}

function attr(node: XmlNode | undefined, name: string): string | undefined {
  const value = node?.[`@_${name}`];
  return typeof value === 'string' ? value : undefined;
}

function intAttr(node: XmlNode | undefined, name: string): number | undefined {
  const raw = attr(node, name);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isInteger(n) ? n : undefined;
}

function textValues(node: XmlNode): Array<string> {
  const out: Array<string> = [];
  const v = node['v'];
  for (const item of Array.isArray(v) ? v : [v]) {
    if (typeof item === 'string') out.push(item);
    else if (typeof item === 'number') out.push(String(item));
    else if (typeof item === 'object' && item !== null) {
      const text = (item as XmlNode)['#text'];
      out.push(typeof text === 'string' ? text : '');
    }
  }
  return out;
}

/**
 * Build the `vm` → display-text map for a workbook's rich values.
 *
 * All three parts are optional and any of them being absent, malformed or of a
 * structure this understands nothing about simply yields no entry for that
 * index — the caller then keeps whatever the cell's `<v>` said.
 *
 * @param metadata   `xl/metadata.xml`, or undefined when the package has none.
 * @param structures `xl/richData/rdrichvaluestructure.xml`.
 * @param values     `xl/richData/rdrichvalue.xml`.
 * @returns A map from a cell's 1-based `vm` index to the text to render.
 */
export function parseRichValueText(
  metadata: Uint8Array | undefined,
  structures: Uint8Array | undefined,
  values: Uint8Array | undefined,
): ReadonlyMap<number, string> {
  const out = new Map<number, string>();
  if (!metadata || !structures || !values) return out;

  const structureList = parseStructures(structures);
  if (structureList.length === 0) return out;
  const richValues = parseValues(values);
  if (richValues.length === 0) return out;

  const meta = parser.parse(decoder.decode(metadata)) as XmlNode;
  const root = nodes(meta, 'metadata')[0];
  if (!root) return out;

  // `rc/@t` is 1-based into <metadataTypes>; only the rich-value type leads
  // anywhere from here.
  const typeNames = nodes(root, 'metadataTypes')
    .flatMap((t) => nodes(t, 'metadataType'))
    .map((t) => attr(t, 'name') ?? '');

  // Each <futureMetadata name="…"> block list, by type name: `rc/@v` indexes it.
  const futureByType = new Map<string, Array<number | undefined>>();
  for (const fm of nodes(root, 'futureMetadata')) {
    const name = attr(fm, 'name');
    if (name === undefined) continue;
    futureByType.set(
      name,
      nodes(fm, 'bk').map((bk) => {
        const rvb = nodes(bk, 'extLst')
          .flatMap((ext) => nodes(ext, 'ext'))
          .flatMap((ext) => nodes(ext, 'rvb'))[0];
        return intAttr(rvb, 'i');
      }),
    );
  }

  const valueMetadata = nodes(root, 'valueMetadata').flatMap((vm) => nodes(vm, 'bk'));
  for (const [index, bk] of valueMetadata.entries()) {
    const rc = nodes(bk, 'rc')[0];
    const typeIndex = intAttr(rc, 't');
    const blockIndex = intAttr(rc, 'v');
    if (typeIndex === undefined || blockIndex === undefined) continue;
    if (typeNames[typeIndex - 1] !== RICH_VALUE_TYPE) continue;
    const richIndex = futureByType.get(RICH_VALUE_TYPE)?.[blockIndex];
    if (richIndex === undefined) continue;
    const text = displayText(richValues[richIndex], structureList);
    // The cell's `vm` is 1-based over this same list.
    if (text !== undefined) out.set(index + 1, text);
  }
  return out;
}

/** What one rich value renders as, or undefined when nothing here understands it. */
function displayText(
  value: { readonly structureIndex: number; readonly values: ReadonlyArray<string> } | undefined,
  structures: ReadonlyArray<RichValueStructure>,
): string | undefined {
  if (!value) return undefined;
  const structure = structures[value.structureIndex];
  if (structure?.type !== '_error') return undefined;
  const at = structure.keys.indexOf('errorType');
  if (at < 0) return undefined;
  const code = Number(value.values[at]);
  return Number.isInteger(code) ? ERROR_TYPES.get(code) : undefined;
}

function parseStructures(part: Uint8Array): Array<RichValueStructure> {
  const doc = parser.parse(decoder.decode(part)) as XmlNode;
  return nodes(doc, 'rvStructures')
    .flatMap((r) => nodes(r, 's'))
    .map((s) => ({
      type: attr(s, 't') ?? '',
      keys: nodes(s, 'k').map((k) => attr(k, 'n') ?? ''),
    }));
}

function parseValues(
  part: Uint8Array,
): Array<{ readonly structureIndex: number; readonly values: ReadonlyArray<string> }> {
  const doc = parser.parse(decoder.decode(part)) as XmlNode;
  return nodes(doc, 'rvData')
    .flatMap((r) => nodes(r, 'rv'))
    .map((rv) => ({ structureIndex: intAttr(rv, 's') ?? 0, values: textValues(rv) }));
}
