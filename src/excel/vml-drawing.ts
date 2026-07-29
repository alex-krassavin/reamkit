// Legacy VML drawing (`xl/drawings/vmlDrawing#.vml`) — the pre-DrawingML shape
// part a worksheet points at with `<legacyDrawing r:id>`.
//
// It matters for form controls. A checkbox, an option button or a group box put
// on a sheet by Excel's Forms toolbar lives ONLY here: the shape carries the
// caption in its `<v:textbox>` and its state in `<x:ClientData>`, and there is
// no `<control>` entry, no ctrlProps part and no DrawingML anchor to find it
// through. tdf111980_radioButtons.xlsx has five such radio buttons and a group
// box beside its five ActiveX ones, and with only the `<control>` list read we
// showed the ActiveX five and lost the other six without a word.
//
// An ActiveX control has a shape here too — the same `o:spid` its `<control
// shapeId>` names — so the caller can tell the two apart by that id.

import { XMLParser } from 'fast-xml-parser';

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
  textNodeName: '#text',
  parseAttributeValue: false,
  parseTagValue: false,
  removeNSPrefix: true,
});

/** A form control declared by a legacy VML shape's `<x:ClientData>`. */
export interface VmlFormControl {
  /** `ST_ObjectType` — `Radio`, `CheckBox`, `GBox`, `Button`, `Drop`, … */
  readonly objectType: string;
  /** The shape id (`o:spid` without its `_x0000_s` prefix), pairing with `<control shapeId>`. */
  readonly shapeId?: string;
  /** The visible label, from the shape's `<v:textbox>`. */
  readonly caption?: string;
  /** `<x:Checked>` — set on a checked check/option button, absent otherwise. */
  readonly checked?: boolean;
  /** `<x:FirstButton/>` — this option button starts a new group. */
  readonly firstButton?: boolean;
}

const SPID_PREFIX = /^_x0000_s/;

/**
 * Read the form controls out of a legacy VML drawing part.
 *
 * Shapes with no `<x:ClientData>`, or whose object type is not a control (a
 * plain drawing, a comment box), are skipped — this is not a general VML shape
 * reader, only the control channel.
 *
 * @param data The raw `vmlDrawing#.vml` bytes.
 * @returns One entry per control shape, in document order.
 */
export function parseVmlFormControls(data: Uint8Array): Array<VmlFormControl> {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const root = asObject(tree['xml']) ?? tree;
  const out: Array<VmlFormControl> = [];
  for (const raw of asArray(root['shape'])) {
    const shape = asObject(raw);
    if (!shape) continue;
    const client = asObject(shape['ClientData']);
    if (!client) continue;
    const objectType = strAttr(client, 'ObjectType');
    if (!objectType) continue;
    const spid = strAttr(shape, 'spid');
    const control: Mutable<VmlFormControl> = { objectType };
    if (spid) control.shapeId = spid.replace(SPID_PREFIX, '');
    const caption = textboxText(shape['textbox']);
    if (caption) control.caption = caption;
    // Present-and-1 means checked; the element is simply absent when it is not.
    const checked = flatText(client['Checked']);
    if (checked === '1') control.checked = true;
    if ('FirstButton' in client) control.firstButton = true;
    out.push(control);
  }
  return out;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };

/** The visible text of a `<v:textbox>`, whose content is an HTML fragment. */
function textboxText(node: unknown): string | undefined {
  const text = flatText(node);
  if (text === undefined) return undefined;
  const collapsed = text.replace(/\s+/g, ' ').trim();
  return collapsed.length > 0 ? collapsed : undefined;
}

/** Every text node under `node`, concatenated — the element tree flattened. */
function flatText(node: unknown): string | undefined {
  if (node === undefined || node === null) return undefined;
  if (typeof node === 'string') return node;
  if (typeof node === 'number' || typeof node === 'boolean') return String(node);
  if (Array.isArray(node)) return node.map((n) => flatText(n) ?? '').join('');
  if (typeof node !== 'object') return undefined;
  let out = '';
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key.startsWith('@_')) continue;
    out += flatText(value) ?? '';
  }
  return out;
}

function asObject(value: unknown): Record<string, unknown> | undefined {
  return value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

function asArray(value: unknown): Array<unknown> {
  return Array.isArray(value) ? value : value === undefined ? [] : [value];
}

function strAttr(obj: Record<string, unknown>, name: string): string | undefined {
  const v = obj[`@_${name}`];
  return typeof v === 'string' ? v : typeof v === 'number' ? String(v) : undefined;
}
