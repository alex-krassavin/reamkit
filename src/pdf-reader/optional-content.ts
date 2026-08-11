// §8.11 — optional content: the layers of a PDF, and which of them are ON.
//
// A file may divide its marks into groups and state, in its own default
// configuration, which groups a viewer shows. A CAD drawing keeps its
// dimensions on one layer and its hatching on another; a form keeps each
// language on its own; issue11144_reduced.pdf keeps three versions of the same
// page on three, two of them off.
//
// Reading them all is not a smaller loss than reading none: the hidden layers
// are drawn OVER the visible one, so the page comes back showing text no viewer
// shows and colours no viewer shows either.

import type { PdfDict, PdfStream, PdfValue } from '@/pdf/objects';
import type { PdfFile } from './document';
import { PDF_NULL, PdfName } from '@/pdf/objects';

/** The document's own answer to "is this group shown?", worked out once. */
const cache = new WeakMap<PdfFile, ReadonlySet<PdfValue>>();

/** An `/OCMD` naming more groups than this is not read: it is not a document. */
const MAX_GROUPS = 4096;

/**
 * The optional-content groups the file's DEFAULT configuration turns off
 * (§8.11.4.3).
 *
 * `/BaseState` says what an unlisted group does — `/ON` unless the file says
 * `/OFF` — and the `/ON` and `/OFF` arrays name the exceptions, `/OFF` last.
 *
 * @param file The document.
 * @returns The resolved OCG dictionaries that are hidden.
 */
export function hiddenGroups(file: PdfFile): ReadonlySet<PdfValue> {
  const had = cache.get(file);
  if (had) return had;
  const hidden = new Set<PdfValue>();
  const props = file.get(file.catalog, 'OCProperties');
  const config = props instanceof Map ? file.get(props, 'D') : undefined;
  if (config instanceof Map) {
    const base = file.get(config, 'BaseState');
    if (base instanceof PdfName && base.value === 'OFF') {
      const all = file.get(props instanceof Map ? props : new Map(), 'OCGs');
      for (const g of asArray(file, all)) hidden.add(g);
    }
    for (const g of asArray(file, file.get(config, 'ON'))) hidden.delete(g);
    for (const g of asArray(file, file.get(config, 'OFF'))) hidden.add(g);
  }
  cache.set(file, hidden);
  return hidden;
}

/**
 * Whether an `/OC` entry names something the page does not show.
 *
 * The entry is either a group itself or an `/OCMD` — a membership dictionary
 * naming several groups and a `/P` policy over them (§8.11.2.3). `AnyOn` is the
 * default and the common case: the content shows if any of its groups does.
 *
 * @param file The document.
 * @param oc   The `/OC` value, unresolved.
 * @returns `true` where the content it guards is hidden.
 */
export function hiddenByOc(file: PdfFile, oc: PdfValue | undefined): boolean {
  if (oc === undefined) return false;
  const hidden = hiddenGroups(file);
  const resolved = file.resolve(oc);
  if (!(resolved instanceof Map)) return false;
  const type = file.get(resolved, 'Type');
  if (!(type instanceof PdfName) || type.value !== 'OCMD') return hidden.has(resolved);
  const groups = asArray(file, file.get(resolved, 'OCGs'));
  // A single group may be given bare rather than in an array.
  const single = file.resolve(resolved.get('OCGs') ?? PDF_NULL);
  const list = groups.length > 0 ? groups : single instanceof Map ? [single] : [];
  if (list.length === 0) return false;
  const on = list.filter((g) => !hidden.has(g)).length;
  const policy = file.get(resolved, 'P');
  switch (policy instanceof PdfName ? policy.value : 'AnyOn') {
    case 'AllOn':
      return on < list.length;
    case 'AnyOff':
      return on === list.length;
    case 'AllOff':
      return on > 0;
    default: // AnyOn
      return on === 0;
  }
}

/**
 * The `/Properties` names a page's content may name in `/OC … BDC` that are
 * hidden (§8.11.3.2).
 *
 * @param file      The document.
 * @param resources The resource dictionary in force.
 * @returns Name → hidden, for the names that ARE hidden.
 */
export function hiddenProperties(file: PdfFile, resources: PdfDict | undefined): Set<string> {
  const out = new Set<string>();
  if (!resources) return out;
  const props = file.get(resources, 'Properties');
  if (!(props instanceof Map)) return out;
  for (const [name, value] of props) {
    if (hiddenByOc(file, value)) out.add(name);
  }
  return out;
}

/** Whether an XObject carries an `/OC` that hides it (§8.11.3.1). */
export function hiddenXObject(file: PdfFile, stream: PdfStream): boolean {
  return hiddenByOc(file, stream.dict.get('OC'));
}

/** An array entry, resolved; a missing or malformed one comes back empty. */
function asArray(file: PdfFile, value: PdfValue | undefined): Array<PdfValue> {
  const r = value !== undefined ? file.resolve(value) : undefined;
  if (!Array.isArray(r)) return [];
  return r.slice(0, MAX_GROUPS).map((v) => file.resolve(v));
}
