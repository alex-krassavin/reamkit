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
 * (column, offset px, row, offset px) pairs (VmlAnchor). The style is in the
 * producer's own measure of the columns, which is not ours wherever the two
 * measure a digit differently; the anchor counts cells, and lands on the grid
 * we draw. A reader takes the anchor where there is one.
 */
export interface VmlShapeBox {
  readonly xPt: number;
  readonly yPt: number;
  readonly widthPt: number;
  readonly heightPt: number;
}

/**
 * `<x:Anchor>` — LeftColumn, LeftOffset, TopRow, TopOffset, RightColumn,
 * RightOffset, BottomRow, BottomOffset: cells, 0-based, and pixels into them.
 */
export type VmlAnchor = readonly [number, number, number, number, number, number, number, number];

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
  /** …and by the cells it covers, `<x:Anchor>` ({@link VmlAnchor}). */
  readonly anchor?: VmlAnchor;
  /** `<v:textbox><font size>` — twentieths of a point, so 160 is 8pt. */
  readonly fontSizePt?: number;
}

/**
 * A cell note's box — the `ObjectType="Note"` shape Excel keeps beside every
 * legacy comment. The comment part holds what the note SAYS; this says whether
 * it is shown and where, and what it is painted with.
 */
export interface VmlNote {
  /** `<x:Row>` / `<x:Column>` — the cell the note belongs to, 0-based. */
  readonly row: number;
  readonly column: number;
  /** `<x:Visible/>` — the note is shown, not merely flagged in its cell. */
  readonly visible: boolean;
  /**
   * `<x:Anchor>` — LeftColumn, LeftOffset, TopRow, TopOffset, RightColumn,
   * RightOffset, BottomRow, BottomOffset: the box in cells, offsets in pixels.
   */
  readonly anchor?: VmlAnchor;
  /** The same box from the shape's `style`, already in points. */
  readonly box?: VmlShapeBox;
  /** `fillcolor` / `strokecolor`, as 6-hex. */
  readonly fillHex: string;
  readonly lineHex: string;
  /** `<v:shadow on="t">` — the hard shadow a note casts down and to the right. */
  readonly shadow: boolean;
  /** `<v:textbox><div style="text-align">` — the side its text is set to. */
  readonly textAlign?: 'left' | 'center' | 'right';
}

/** A parsed legacy VML drawing: its form controls plus every shape's box. */
export interface VmlDrawing {
  readonly controls: ReadonlyArray<VmlFormControl>;
  /** Every cell note's box, shown or not, in document order. */
  readonly notes: ReadonlyArray<VmlNote>;
  /**
   * Shape id (`o:spid` without its `_x0000_s` prefix) → box, for EVERY shape in
   * the part, controls or not. An ActiveX control's box lives nowhere else: its
   * `activeX#.xml` carries only a class id and its `<control>` element carries
   * no anchor, so the `ObjectType="Pict"` shape that shares its `shapeId` is
   * the one thing that says where it goes.
   */
  readonly boxes: ReadonlyMap<string, VmlShapeBox>;
  /** Shape id → `<x:Anchor>`, for every shape that carries one. */
  readonly anchors: ReadonlyMap<string, VmlAnchor>;
  /**
   * Shape ids whose `<x:ClientData>` clears `<x:PrintObject>` — Excel's "Print
   * object" checkbox, the legacy spelling of §18.3.1.20 `<controlPr print>`.
   * Such a shape is on screen only, and the `<control>` that shares its id must
   * not be drawn either.
   */
  readonly nonPrinting: ReadonlySet<string>;
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
  const notes: Array<VmlNote> = [];
  const boxes = new Map<string, VmlShapeBox>();
  const anchors = new Map<string, VmlAnchor>();
  const nonPrinting = new Set<string>();
  for (const { shape, box } of shapesOf(root)) {
    // The shape id is `o:spid` when the producer writes one and the plain `id`
    // otherwise — button-form-control.xlsx spells it only the second way, and
    // reading just `o:spid` left its shape anonymous, so nothing could pair it
    // with the `<control shapeId="1025">` that names it.
    const idAttr = strAttr(shape, 'id');
    const spid =
      strAttr(shape, 'spid') ??
      (idAttr !== undefined && SPID_PREFIX.test(idAttr) ? idAttr : undefined);
    const shapeId = spid?.replace(SPID_PREFIX, '');
    if (shapeId && box) boxes.set(shapeId, box);
    const client = asObject(shape['ClientData']);
    if (!client) continue;
    const anchor = anchorOf(client);
    if (shapeId && anchor) anchors.set(shapeId, anchor);
    // Excel's "Print object", the legacy spelling of §18.3.1.20's `print`.
    // Absent means print — only `False` takes the shape off the page.
    const prints = flatText(client['PrintObject'])?.toLowerCase() !== 'false';
    if (!prints) {
      if (shapeId) nonPrinting.add(shapeId);
      continue;
    }
    const objectType = strAttr(client, 'ObjectType');
    if (objectType === 'Note') {
      const note = noteOf(shape, client, box);
      if (note) notes.push(note);
      continue;
    }
    if (!objectType || !CONTROL_TYPES.has(objectType)) continue;
    const control: Mutable<VmlFormControl> = { objectType };
    if (shapeId) control.shapeId = shapeId;
    if (box) control.box = box;
    if (anchor) control.anchor = anchor;
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
  return { controls: out, notes, boxes, anchors, nonPrinting };
}

/** `<x:Anchor>`'s eight whole numbers, or undefined when it is not that. */
function anchorOf(client: Record<string, unknown>): VmlAnchor | undefined {
  const numbers = (flatText(client['Anchor']) ?? '').split(',').map((n) => Number(n.trim()));
  return numbers.length === 8 && numbers.every((n) => Number.isInteger(n))
    ? (numbers as unknown as VmlAnchor)
    : undefined;
}

// What a note's shape says about its box. Excel's own notes are pale yellow
// (the system's tooltip colour, which older files name `infoBackground`) with
// a black outline and a shadow; a shape that says nothing else gets the same.
function noteOf(
  shape: Record<string, unknown>,
  client: Record<string, unknown>,
  box: VmlShapeBox | undefined,
): VmlNote | undefined {
  const row = Number(flatText(client['Row'])?.trim());
  const column = Number(flatText(client['Column'])?.trim());
  if (!Number.isInteger(row) || !Number.isInteger(column) || row < 0 || column < 0) {
    return undefined;
  }
  const anchor = anchorOf(client);
  const shadowNode = asObject(shape['shadow']);
  const shadowOn = shadowNode ? strAttr(shadowNode, 'on') : undefined;
  const textAlign = /text-align:\s*(left|center|right)/i
    .exec(styleWithin(shape['textbox']) ?? '')?.[1]
    ?.toLowerCase() as VmlNote['textAlign'];
  return {
    row,
    column,
    // A shown note says so; a hidden one says nothing, or `visibility:hidden`.
    visible: 'Visible' in client,
    ...(anchor ? { anchor } : {}),
    ...(box ? { box } : {}),
    fillHex: vmlColorHex(strAttr(shape, 'fillcolor')) ?? 'FFFFE1',
    lineHex: vmlColorHex(strAttr(shape, 'strokecolor')) ?? '000000',
    shadow: shadowOn === 't' || shadowOn === 'true',
    ...(textAlign ? { textAlign } : {}),
  };
}

/** The first `style` declared anywhere under `node` — a textbox's `<div style>`. */
function styleWithin(node: unknown): string | undefined {
  if (Array.isArray(node)) {
    for (const n of node) {
      const found = styleWithin(n);
      if (found !== undefined) return found;
    }
    return undefined;
  }
  const obj = asObject(node);
  if (!obj) return undefined;
  const own = strAttr(obj, 'style');
  if (own !== undefined && /text-align/i.test(own)) return own;
  for (const [key, value] of Object.entries(obj)) {
    if (key.startsWith('@_')) continue;
    const found = styleWithin(value);
    if (found !== undefined) return found;
  }
  return undefined;
}

/** A VML colour — `#rrggbb`, `#rgb`, or one of the few names notes use — as 6-hex. */
function vmlColorHex(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const v = value.trim().toLowerCase();
  const long = /^#([0-9a-f]{6})\b/.exec(v);
  if (long) return long[1]!.toUpperCase();
  const short = /^#([0-9a-f]{3})\b/.exec(v);
  if (short)
    return [...short[1]!]
      .map((c) => c + c)
      .join('')
      .toUpperCase();
  if (v.startsWith('infobackground')) return 'FFFFE1';
  if (v.startsWith('infotext') || v.startsWith('black')) return '000000';
  if (v.startsWith('white')) return 'FFFFFF';
  return undefined;
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
  // Trimmed first: blanks on both sides of an empty unit were split every
  // way between the two `\s*` that surrounded it — quadratic in a long run.
  const m = /^(-?[\d.]+)\s*([a-z]*)$/i.exec(value.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  const unit = m[2]!.toLowerCase();
  const factor = unit === '' ? 0.75 : CSS_UNITS.get(unit);
  return factor === undefined ? undefined : n * factor;
}

/** A `style` attribute's declarations, by lower-cased name. */
function styleProps(style: string | undefined): Map<string, string> {
  const props = new Map<string, string>();
  for (const decl of (style ?? '').split(';')) {
    const i = decl.indexOf(':');
    if (i < 0) continue;
    props.set(decl.slice(0, i).trim().toLowerCase(), decl.slice(i + 1).trim());
  }
  return props;
}

/** The shape's rectangle from its `style` declaration; undefined if incomplete. */
function shapeBox(style: string | undefined): VmlShapeBox | undefined {
  if (style === undefined) return undefined;
  const props = styleProps(style);
  const widthPt = cssLengthPt(props.get('width'));
  const heightPt = cssLengthPt(props.get('height'));
  if (widthPt === undefined || heightPt === undefined) return undefined;
  return {
    xPt: (cssLengthPt(props.get('margin-left')) ?? 0) + (cssLengthPt(props.get('left')) ?? 0),
    yPt: (cssLengthPt(props.get('margin-top')) ?? 0) + (cssLengthPt(props.get('top')) ?? 0),
    widthPt,
    heightPt,
  };
}

/**
 * VML `coordorigin`/`coordsize` — the space a group lays its children out in,
 * mapped onto the group's own box: child units to points on the sheet.
 */
interface GroupSpace {
  readonly x: (u: number) => number;
  readonly y: (u: number) => number;
  /** Points per unit across and down. */
  readonly sx: number;
  readonly sy: number;
}

function groupSpace(group: Record<string, unknown>, box: VmlShapeBox): GroupSpace {
  const pair = (value: string | undefined, fallback: number): [number, number] => {
    const [a, b] = (value ?? '').split(',').map((n) => Number.parseFloat(n));
    return [Number.isFinite(a) ? a! : fallback, Number.isFinite(b) ? b! : fallback];
  };
  const [ox, oy] = pair(strAttr(group, 'coordorigin'), 0);
  // VML's default space is a thousand units square.
  const [cw, ch] = pair(strAttr(group, 'coordsize'), 1000);
  const sx = cw !== 0 ? box.widthPt / cw : 0;
  const sy = ch !== 0 ? box.heightPt / ch : 0;
  return {
    x: (u) => box.xPt + (u - ox) * sx,
    y: (u) => box.yPt + (u - oy) * sy,
    sx,
    sy,
  };
}

/** A child's box in its group's space: bare numbers, the group's units. */
function boxInGroup(style: string | undefined, space: GroupSpace): VmlShapeBox | undefined {
  const props = styleProps(style);
  const units = (name: string): number | undefined => {
    const n = Number.parseFloat(props.get(name) ?? '');
    return Number.isFinite(n) ? n : undefined;
  };
  const width = units('width');
  const height = units('height');
  if (width === undefined || height === undefined) return undefined;
  const left = (units('margin-left') ?? 0) + (units('left') ?? 0);
  const top = (units('margin-top') ?? 0) + (units('top') ?? 0);
  return {
    xPt: space.x(left),
    yPt: space.y(top),
    widthPt: width * space.sx,
    heightPt: height * space.sy,
  };
}

/**
 * Every shape in the drawing with its box on the sheet — those at the top and
 * those inside a `<v:group>`, whose `left`/`top` count in the group's own
 * space. Reading only the top level lost every control a group held:
 * 45540_form_Footer.xlsx groups the 27 check boxes of its "What industry are
 * you in?" block, and they ended up in a list after the form.
 */
function shapesOf(
  root: Record<string, unknown>,
): Array<{ shape: Record<string, unknown>; box: VmlShapeBox | undefined }> {
  const out: Array<{ shape: Record<string, unknown>; box: VmlShapeBox | undefined }> = [];
  const visit = (node: Record<string, unknown>, space: GroupSpace | undefined, depth: number) => {
    const boxOf = (style: string | undefined): VmlShapeBox | undefined =>
      space ? boxInGroup(style, space) : shapeBox(style);
    for (const raw of asArray(node['shape'])) {
      const shape = asObject(raw);
      if (shape) out.push({ shape, box: boxOf(strAttr(shape, 'style')) });
    }
    if (depth >= 8) return;
    for (const raw of asArray(node['group'])) {
      const group = asObject(raw);
      const box = group ? boxOf(strAttr(group, 'style')) : undefined;
      if (group && box) visit(group, groupSpace(group, box), depth + 1);
    }
  };
  visit(root, undefined, 0);
  return out;
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
