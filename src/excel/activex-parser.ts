// ActiveX control property bag (E-SHEET W10). An embedded ActiveX control points
// (through a worksheet <oleObject> relationship) at its xl/activeX/activeX#.xml
// part — an <ax:ocx> element whose <ax:ocxPr name value> children persist the
// control's visible state (Caption, Value, GroupName, …) when it is saved as a
// property bag. That is all the print model needs to list the control with a
// type-appropriate affordance and its current value, the way form controls are.
// (For a control persisted to its binary .bin — MS-OFORMS — the property bag is
// the OLE/CFB stream; reading that is a documented follow-up. The CFB container
// reader (src/core/ole) is the keystone that makes it reachable.)

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  removeNSPrefix: true, // ax:ocx → ocx, ax:name → @_name, r:id → @_id
});

export interface ActiveXProps {
  // The control's displayed text (CommandButton/Label/CheckBox/OptionButton).
  readonly caption?: string;
  // The control's state/value as persisted ("1"/"0" for a checkbox/option, the
  // text for a textbox, a number for a spin/scroll).
  readonly value?: string;
  // OptionButton group membership (mutually-exclusive set).
  readonly groupName?: string;
}

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

// Forms.CheckBox.1 → 'checkbox', Forms.CommandButton.1 → 'button', … — the
// control class from the <oleObject progId>, normalised to the affordance keys
// the projection switches on. Unknown progIds fall back to a generic control.
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
