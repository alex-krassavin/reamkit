// ActiveX control property bag (E-SHEET W10). An embedded ActiveX control points
// (through a worksheet <oleObject> relationship) at its xl/activeX/activeX#.xml
// part — an <ax:ocx> element whose <ax:ocxPr name value> children persist the
// control's visible state (Caption, Value, GroupName, …) when it is saved as a
// property bag. That is all the print model needs to list the control with a
// type-appropriate affordance and its current value, the way form controls are.
//
// W10 tail — a control saved with `ax:persistence="persistStreamInit"` keeps NO
// <ax:ocxPr>; its state lives in the binary activeX#.bin (MS-OFORMS). For the
// MorphData family (checkbox / option / toggle / text / combo / list — one shared
// MorphDataControl structure) parseActiveXBin walks that stream to recover the
// caption / value / group name. The byte layout is grounded in [MS-OFORMS] §2.2.5
// validated against a real Office-format sample (the LibreOffice activex_checkbox
// fixture). A CommandButton/Label .bin (a different structure, not validated here)
// and a CFB-storage .bin are left to the property bag — a missing caption, never
// a wrong one.

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  removeNSPrefix: true, // ax:ocx → ocx, ax:name → @_name, r:id → @_id
});

/** An ActiveX control's visible state from its property bag / binary stream (E-SHEET W10). */
export interface ActiveXProps {
  /** The control's displayed text (CommandButton/Label/CheckBox/OptionButton). */
  readonly caption?: string;
  /**
   * The control's state/value as persisted (`"1"`/`"0"` for a checkbox/option, the
   * text for a textbox, a number for a spin/scroll).
   */
  readonly value?: string;
  /** OptionButton group membership (mutually-exclusive set). */
  readonly groupName?: string;
}

/**
 * Parse a `xl/activeX/activeX#.xml` part (an `<ax:ocx>` whose `<ax:ocxPr name
 * value>` children persist the visible state) into {@link ActiveXProps}. Reads the
 * caption / value / group name; returns `{}` for a non-property-bag control.
 */
export function parseActiveX(data: Uint8Array): ActiveXProps {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const ocx = asObject(tree['ocx']);
  if (!ocx) return {};
  const props = new Map<string, string>();
  for (const pr of toArray(ocx['ocxPr'])) {
    const o = asObject(pr);
    const name = o && strAttr(o, 'name');
    const value = o && strAttr(o, 'value');
    if (name) props.set(name.toLowerCase(), value ?? '');
  }
  const caption = props.get('caption');
  const value = props.get('value');
  const groupName = props.get('groupname');
  return {
    ...(caption !== undefined ? { caption } : {}),
    ...(value !== undefined ? { value } : {}),
    ...(groupName !== undefined ? { groupName } : {}),
  };
}

/**
 * Map an ActiveX `<oleObject progId>` to the affordance key the projection switches
 * on: `Forms.CheckBox.1 → 'checkbox'`, `Forms.CommandButton.1 → 'button'`, …. An
 * unknown progId falls back to a generic `'control'`.
 */
export function activeXType(progId: string | undefined): string {
  const m = /Forms\.(\w+)\.\d+/i.exec(progId ?? '');
  switch ((m?.[1] ?? '').toLowerCase()) {
    case 'checkbox':
      return 'checkbox';
    case 'optionbutton':
      return 'option';
    case 'commandbutton':
      return 'button';
    case 'togglebutton':
      return 'toggle';
    case 'textbox':
      return 'textbox';
    case 'combobox':
      return 'combo';
    case 'listbox':
      return 'list';
    case 'label':
      return 'label';
    case 'spinbutton':
      return 'spin';
    case 'scrollbar':
      return 'scroll';
    case 'frame':
      return 'frame';
    default:
      return 'control';
  }
}

function asObject(v: unknown): Record<string, unknown> | undefined {
  return v && typeof v === 'object' && !Array.isArray(v)
    ? (v as Record<string, unknown>)
    : undefined;
}

function toArray(v: unknown): Array<unknown> {
  return Array.isArray(v) ? v : v !== undefined && v !== null ? [v] : [];
}

function strAttr(obj: Record<string, unknown>, key: string): string | undefined {
  const v = obj[`@_${key}`];
  return typeof v === 'string' ? v : undefined;
}

/**
 * The `<ax:ocx r:id>` of a control whose state is persisted to a binary stream
 * (`persistStreamInit` / `persistStream` / `persistStorage`) rather than to
 * `<ax:ocxPr>`; the reader resolves it (through the `activeX#.xml` part's own
 * relationships) to the `activeX#.bin`. Returns undefined for a property-bag
 * control (the `.bin` is not needed).
 */
export function activeXBinRelId(xmlData: Uint8Array): string | undefined {
  const tree = parser.parse(decoder.decode(xmlData)) as Record<string, unknown>;
  const ocx = asObject(tree['ocx']);
  if (!ocx) return undefined;
  const persistence = strAttr(ocx, 'persistence');
  if (
    persistence !== 'persistStreamInit' &&
    persistence !== 'persistStream' &&
    persistence !== 'persistStorage'
  ) {
    return undefined;
  }
  return strAttr(ocx, 'id');
}

// --- MS-OFORMS MorphDataControl (§2.2.5) binary parse -----------------------

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
// MorphDataDataBlock fixed properties before the value/caption sizes — each a
// [propmask bit, byte size]; consumed in order when the bit is set (§2.2.5.3).
const MORPH_DATA_BLOCK_1: ReadonlyArray<readonly [number, number]> = [
  [0, 4],
  [1, 4],
  [2, 4],
  [3, 4],
  [4, 1],
  [5, 1],
  [6, 1],
  [7, 1],
  [9, 2],
  [10, 4],
  [11, 2],
  [12, 2],
  [13, 2],
  [14, 2],
  [15, 2],
  [16, 1],
  [17, 1],
  [18, 1],
  [20, 1],
  [21, 1],
];
// …and the fixed properties between the caption size and the group-name size.
const MORPH_DATA_BLOCK_2: ReadonlyArray<readonly [number, number]> = [
  [24, 4],
  [25, 4],
  [26, 4],
  [27, 2],
  [28, 2],
  [29, 2],
];
const BIT_F_SIZE = 8;
const BIT_F_VALUE = 22;
const BIT_F_CAPTION = 23;
const BIT_F_GROUP_NAME = 32;
const MORPH_MAJOR_VERSION = 2;
const ONE_8 = new TextDecoder('windows-1252'); // a compressed (1 byte/char) string
const UTF16 = new TextDecoder('utf-16le'); // an uncompressed string

function u16le(d: Uint8Array, i: number): number {
  return d[i]! | (d[i + 1]! << 8);
}
function u32le(d: Uint8Array, i: number): number {
  return (d[i]! | (d[i + 1]! << 8) | (d[i + 2]! << 16) | (d[i + 3]! << 24)) >>> 0;
}

/**
 * Parse an `activeX#.bin` (a `persistStreamInit` MorphDataControl) for its caption,
 * value and group name, per [MS-OFORMS] §2.2.5. Returns `{}` for any non-MorphData
 * stream (a CFB storage, a CommandButton/Label, a structurally implausible blob) —
 * a graceful miss, never a wrong caption. Bounds-checked throughout; any overflow
 * bails to `{}`.
 */
export function parseActiveXBin(data: Uint8Array): ActiveXProps {
  if (data.length >= 8 && CFB_MAGIC.every((b, i) => data[i] === b)) return {};
  // The stream is [16-byte classid GUID][MorphDataControl]; tolerate a missing
  // GUID. The control opens with a 2-byte version whose major byte is 2.
  let off: number | undefined;
  if (data.length >= 18 && data[16] === 0 && data[17] === MORPH_MAJOR_VERSION) off = 16;
  else if (data.length >= 2 && data[0] === 0 && data[1] === MORPH_MAJOR_VERSION) off = 0;
  if (off === undefined) return {};

  let p = off + 2;
  if (p + 10 > data.length) return {};
  const cb = u16le(data, p);
  p += 2;
  const blockEnd = Math.min(data.length, p + cb);
  const propLo = u32le(data, p);
  const propHi = u32le(data, p + 4);
  p += 8;
  const dataStart = p; // alignment is relative to the start of the DataBlock
  const bit = (b: number): boolean =>
    (b < 32 ? (propLo >>> b) & 1 : (propHi >>> (b - 32)) & 1) === 1;
  const align = (n: number): void => {
    p += (n - ((p - dataStart) % n)) % n;
  };

  // The fixed DataBlock properties only need skipping; the string sizes are read.
  for (const [b, size] of MORPH_DATA_BLOCK_1) {
    if (bit(b)) {
      align(size);
      p += size;
    }
  }
  const readCount = (): number | undefined => {
    align(4);
    if (p + 4 > blockEnd) return undefined;
    const v = u32le(data, p);
    p += 4;
    return v;
  };
  let valueCount: number | undefined;
  let captionCount: number | undefined;
  let groupCount: number | undefined;
  if (bit(BIT_F_VALUE)) {
    valueCount = readCount();
    if (valueCount === undefined) return {};
  }
  if (bit(BIT_F_CAPTION)) {
    captionCount = readCount();
    if (captionCount === undefined) return {};
  }
  for (const [b, size] of MORPH_DATA_BLOCK_2) {
    if (bit(b)) {
      align(size);
      p += size;
    }
  }
  if (bit(BIT_F_GROUP_NAME)) {
    groupCount = readCount();
    if (groupCount === undefined) return {};
  }

  // ExtraDataBlock (§2.2.5.4): the 8-byte DisplayedSize (when fSize), then each
  // present string 4-byte aligned. A count's top bit ⇒ compressed (1 byte/char).
  if (bit(BIT_F_SIZE)) {
    align(4);
    p += 8;
  }
  const readString = (count: number | undefined): string | undefined => {
    if (count === undefined) return undefined;
    align(4);
    const len = count & 0x7fffffff;
    if (len === 0) return '';
    if (p + len > blockEnd) return undefined;
    const slice = data.subarray(p, p + len);
    p += len;
    return (count & 0x80000000) !== 0 ? ONE_8.decode(slice) : UTF16.decode(slice);
  };
  const value = readString(valueCount);
  const caption = readString(captionCount);
  const groupName = readString(groupCount);
  return {
    ...(caption !== undefined && caption !== '' ? { caption } : {}),
    ...(value !== undefined && value !== '' ? { value } : {}),
    ...(groupName !== undefined && groupName !== '' ? { groupName } : {}),
  };
}
