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

/**
 * Where a VML shape sits, in points from the sheet's top-left — read from the
 * CSS `style` attribute (`margin-left`/`margin-top`/`width`/`height`).
 *
 * The same rectangle is also expressed by `<x:ClientData><x:Anchor>` as
 * (column, offset px, row, offset px) pairs, which needs the sheet's column
 * widths and row heights to resolve. The style is already absolute, and the two
 * agree exactly in every file checked, so this reads the simpler one.
 */
export interface VmlShapeBox {
  readonly xPt: number;
  readonly yPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
}

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
  /** Where the control sits, from the shape's `style` (§{@link VmlShapeBox}). */
  readonly box?: VmlShapeBox;
  /** `<v:textbox><font size>` — twentieths of a point, so 160 is 8pt. */
  readonly fontSizePt?: number;
}

/** A parsed legacy VML drawing: its form controls plus every shape's box. */
export interface VmlDrawing {
  readonly controls: ReadonlyArray<VmlFormControl>;
  /**
   * Shape id (`o:spid` without its `_x0000_s` prefix) → box, for EVERY shape in
   * the part, controls or not. An ActiveX control's box lives nowhere else: its
   * `activeX#.xml` carries only a class id and its `<control>` element carries
   * no anchor, so the `ObjectType="Pict"` shape that shares its `shapeId` is
   * the one thing that says where it goes.
   */
  readonly boxes: ReadonlyMap<string, VmlShapeBox>;
}

const SPID_PREFIX = /^_x0000_s/;

/**
 * The `x:ObjectType` values that name a FORM CONTROL.
 *
 * Not everything with an `<x:ClientData>` is one — a cell comment is a VML
 * shape with `ObjectType="Note"`, and so are text boxes, pictures and movies.
 * Taking the element's presence as the test listed every comment in the
 * document as a form control.
 */
const CONTROL_TYPES: ReadonlySet<string> = new Set([
  'Button',
  'Checkbox',
  'CheckBox',
  'Dialog',
  'Drop',
  'Edit',
  'EditBox',
  'GBox',
  'Label',
  'List',
  'Radio',
  'Scroll',
  'Spin',
]);

/**
 * Read the form controls out of a legacy VML drawing part.
 *
 * Shapes with no `<x:ClientData>`, or whose object type is not a control (a
 * comment, a text box, a picture), are skipped — this is not a general VML
 * shape reader, only the control channel. See {@link CONTROL_TYPES}.
 *
 * @param data The raw `vmlDrawing#.vml` bytes.
 * @returns One entry per control shape, in document order.
 */
export function parseVmlFormControls(data: Uint8Array): Array<VmlFormControl> {
  return [...parseVmlDrawing(data).controls];
}

/**
 * Read a legacy VML drawing part: its form controls and every shape's box.
 *
 * @param data The raw `vmlDrawing#.vml` bytes.
 */
export function parseVmlDrawing(data: Uint8Array): VmlDrawing {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const root = asObject(tree['xml']) ?? tree;
  const out: Array<VmlFormControl> = [];
  const boxes = new Map<string, VmlShapeBox>();
  for (const raw of asArray(root['shape'])) {
    const shape = asObject(raw);
    if (!shape) continue;
    const spid = strAttr(shape, 'spid');
    const shapeId = spid?.replace(SPID_PREFIX, '');
    const box = shapeBox(strAttr(shape, 'style'));
    if (shapeId && box) boxes.set(shapeId, box);
    const client = asObject(shape['ClientData']);
    if (!client) continue;
    const objectType = strAttr(client, 'ObjectType');
    if (!objectType || !CONTROL_TYPES.has(objectType)) continue;
    const control: Mutable<VmlFormControl> = { objectType };
    if (shapeId) control.shapeId = shapeId;
    if (box) control.box = box;
    const caption = textboxText(shape['textbox']);
    if (caption) control.caption = caption;
    const fontSizePt = textboxFontSizePt(shape['textbox']);
    if (fontSizePt !== undefined) control.fontSizePt = fontSizePt;
    // Present-and-1 means checked; the element is simply absent when it is not.
    const checked = flatText(client['Checked']);
    if (checked === '1') control.checked = true;
    if ('FirstButton' in client) control.firstButton = true;
    out.push(control);
  }
  return { controls: out, boxes };
}

/** CSS length → points. A bare number is pixels, VML's implicit unit. */
const CSS_UNITS: ReadonlyMap<string, number> = new Map([
  ['pt', 1],
  ['px', 0.75], // 96 dpi
  ['in', 72],
  ['cm', 72 / 2.54],
  ['mm', 7.2 / 2.54],
  ['pc', 12],
]);

function cssLengthPt(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const m = /^\s*(-?[\d.]+)\s*([a-z]*)\s*$/i.exec(value);
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  const unit = m[2]!.toLowerCase();
  const factor = unit === '' ? 0.75 : CSS_UNITS.get(unit);
  return factor === undefined ? undefined : n * factor;
}

/** The shape's rectangle from its `style` declaration; undefined if incomplete. */
function shapeBox(style: string | undefined): VmlShapeBox | undefined {
  if (style === undefined) return undefined;
  const props = new Map<string, string>();
  for (const decl of style.split(';')) {
    const i = decl.indexOf(':');
    if (i < 0) continue;
    props.set(decl.slice(0, i).trim().toLowerCase(), decl.slice(i + 1).trim());
  }
  const widthPt = cssLengthPt(props.get('width'));
  const heightPt = cssLengthPt(props.get('height'));
  if (widthPt === undefined || heightPt === undefined) return undefined;
  return {
    xPt: cssLengthPt(props.get('margin-left')) ?? 0,
    yPt: cssLengthPt(props.get('margin-top')) ?? 0,
    widthPt,
    heightPt,
  };
}

/** `<v:textbox>`'s `<font size>` — twentieths of a point (160 ⇒ 8pt). */
function textboxFontSizePt(node: unknown): number | undefined {
  const box = asObject(node);
  const div = box ? asObject(box['div']) : undefined;
  const font = asObject((div ?? box)?.['font']);
  const size = font ? strAttr(font, 'size') : undefined;
  const twips = size === undefined ? NaN : Number(size);
  return Number.isFinite(twips) && twips > 0 ? twips / 20 : undefined;
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
