// Form control properties (E-SHEET W8). A worksheet form control (checkbox,
// option button, spinner, …) points through a relationship at its ctrlProp part
// (xl/ctrlProps/ctrlProp#.xml), a single <formControlPr> element carrying the
// control's `objectType` plus its state — `checked` for check/option buttons and
// `val` for the spin / scroll / list controls. That is all the print model needs
// to list the control with a type-appropriate affordance and its current value.

import { XMLParser } from 'fast-xml-parser';

const decoder = new TextDecoder('utf-8');

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  removeNSPrefix: true,
});

export interface FormControlProps {
  // §18.18.18-ish ST_ObjectType — CheckBox, Radio, Spin, Scroll, Drop, List,
  // Buttons, Label, GBox, Dialog, EditBox, Note … (the producer's spelling kept).
  readonly objectType?: string;
  readonly checked?: boolean;
  readonly value?: number;
}

export function parseFormControlProps(data: Uint8Array): FormControlProps {
  const tree = parser.parse(decoder.decode(data)) as Record<string, unknown>;
  const pr = tree['formControlPr'];
  const obj = pr && typeof pr === 'object' ? (pr as Record<string, unknown>) : undefined;
  if (!obj) return {};
  const objectType = strAttr(obj, 'objectType');
  const checkedRaw = strAttr(obj, 'checked');
  const valRaw = strAttr(obj, 'val');
  const value = valRaw !== undefined && /^-?\d+$/.test(valRaw) ? Number(valRaw) : undefined;
  return {
    ...(objectType ? { objectType } : {}),
    ...(checkedRaw !== undefined
      ? { checked: checkedRaw === 'Checked' || checkedRaw === '1' }
      : {}),
    ...(value !== undefined ? { value } : {}),
  };
}

function strAttr(obj: Record<string, unknown>, key: string): string | undefined {
  const v = obj[`@_${key}`];
  return typeof v === 'string' ? v : undefined;
}
