// ISO 32000-1:2008 §14.7–14.8 — Tagged PDF logical structure.
//
// Builds the document structure tree that accessible / PDF/A-1a readers use to
// recover reading order: a /StructTreeRoot pointing at a tree of /StructElem
// objects (Document → H1…H6 / P / L→LI→LBody / Table→TR→TD → Figure). Each
// element ties back to one or more marked-content sequences in the page content
// streams through marked-content references (MCRs) and the /ParentTree number
// tree (§14.7.4.4).
//
// This module is a generic structure emitter. The docx → structure-type mapping
// policy lives in the renderer; here we only assemble PDF objects. Object ids
// are allocated by a deterministic pre-order DFS so the output stays
// byte-identical for identical input (the writer's /ID hash depends on it).

import type { PdfDict, PdfRef, PdfValue } from '@/pdf/objects';
import type { PdfDocument } from '@/pdf/writer';
import { PDF_NULL, dict, name, unicodeString } from '@/pdf/objects';

// Standard structure types (ISO 32000-1 Table 333/337). Every type here is
// recognised without a /RoleMap, so the builder never emits one.
export type StructType =
  | 'Document'
  | 'Part'
  | 'Sect'
  | 'H1'
  | 'H2'
  | 'H3'
  | 'H4'
  | 'H5'
  | 'H6'
  | 'P'
  | 'L'
  | 'LI'
  | 'Lbl'
  | 'LBody'
  | 'Table'
  | 'TR'
  | 'TH'
  | 'TD'
  | 'Caption'
  | 'Figure'
  | 'Span';

// One marked-content sequence owned by a node: MCID `mcid` on page `pageIndex`.
interface Mcref {
  readonly pageIndex: number;
  readonly mcid: number;
}

export class StructNode {
  readonly children: Array<StructNode> = [];
  readonly mcrefs: Array<Mcref> = [];
  parent: StructNode | null = null;
  ref: PdfRef | null = null;
  // Alternate text (/Alt, §14.9.4) — required on Figure for PDF/A-1a.
  alt: string | null = null;
  // Natural language (/Lang) when it differs from the document default.
  lang: string | null = null;
  // §14.8.5.2 Table attributes (emitted via a /A attribute object on the cell):
  //   scope   — /Scope on a TH (Row/Column) so AT binds headers to data cells.
  //   colSpan — /ColSpan when the cell spans >1 column (gridSpan).
  //   rowSpan — /RowSpan when the cell spans >1 row (vertical merge).
  scope: 'Row' | 'Column' | null = null;
  colSpan: number | null = null;
  rowSpan: number | null = null;

  constructor(
    readonly id: number,
    public type: StructType,
  ) {}
}

export class StructTreeBuilder {
  private readonly nodes: Array<StructNode> = [];
  // The root logical element (the single /Document under StructTreeRoot).
  readonly root: StructNode;

  constructor() {
    this.root = this.create('Document', null);
  }

  // Create a node of `type` as the last child of `parent` (null = a free node;
  // only the root is created that way).
  create(type: StructType, parent: StructNode | null): StructNode {
    const node = new StructNode(this.nodes.length, type);
    node.parent = parent;
    if (parent) parent.children.push(node);
    this.nodes.push(node);
    return node;
  }

  node(id: number): StructNode {
    const n = this.nodes[id];
    if (!n) throw new Error(`Unknown struct node id ${id}`);
    return n;
  }

  // Record that marked content `mcid` on page `pageIndex` is the content of the
  // node `nodeId`. Called from the emit phase as MCIDs are assigned.
  addMcref(nodeId: number, pageIndex: number, mcid: number): void {
    this.node(nodeId).mcrefs.push({ pageIndex, mcid });
  }

  // Emit the StructTreeRoot, every StructElem, and the ParentTree; return the
  // StructTreeRoot ref for the catalog. Must run after all pages are added (so
  // `pageRefs` is complete) and after every addMcref call.
  emit(doc: PdfDocument, pageRefs: ReadonlyArray<PdfRef>): PdfRef {
    // Deterministic pre-order DFS → stable object-id assignment.
    const order: Array<StructNode> = [];
    const walk = (n: StructNode): void => {
      order.push(n);
      for (const c of n.children) walk(c);
    };
    walk(this.root);

    // Pass 1: allocate one StructElem object per node so /P and /K refs resolve.
    const dicts = new Map<number, PdfDict>();
    for (const n of order) {
      const d = dict({ Type: name('StructElem'), S: name(n.type) });
      n.ref = doc.add(d);
      dicts.set(n.id, d);
    }

    // StructTreeRoot needs a ref before we set the root element's /P to it.
    const rootDict = dict({ Type: name('StructTreeRoot') });
    const rootRef = doc.add(rootDict);

    // Pass 2: populate /P, /Pg, /K, /Alt, /Lang now that every ref exists.
    for (const n of order) {
      const d = dicts.get(n.id)!;
      d.set('P', n.parent ? n.parent.ref! : rootRef);
      // /Pg = the page of this element's first content (bare-MCID resolution
      // base); leaf MCRs below always carry their own /Pg so split elements are
      // still correct.
      const firstPage = n.mcrefs.length > 0 ? n.mcrefs[0]!.pageIndex : undefined;
      if (firstPage !== undefined) d.set('Pg', pageRefs[firstPage]!);

      // /K = child elements (logical order) followed by this node's own MCRs.
      // Our construction keeps a node either a container (child elems) or a leaf
      // (MCRs); mixing is permitted by the spec but we never need it.
      const kids: Array<PdfValue> = [];
      for (const c of n.children) kids.push(c.ref!);
      for (const m of n.mcrefs) {
        kids.push(dict({ Type: name('MCR'), Pg: pageRefs[m.pageIndex]!, MCID: m.mcid }));
      }
      if (kids.length === 1) d.set('K', kids[0]!);
      else if (kids.length > 1) d.set('K', kids);

      if (n.alt !== null) d.set('Alt', unicodeString(n.alt));
      if (n.lang !== null) d.set('Lang', n.lang);
      // §14.8.5.2 — table cells carry scope/span via a /Table attribute object.
      if (n.scope !== null || n.colSpan !== null || n.rowSpan !== null) {
        const attrs: Record<string, PdfValue> = { O: name('Table') };
        if (n.scope !== null) attrs['Scope'] = name(n.scope);
        if (n.colSpan !== null) attrs['ColSpan'] = n.colSpan;
        if (n.rowSpan !== null) attrs['RowSpan'] = n.rowSpan;
        d.set('A', dict(attrs));
      }
    }

    // /ParentTree (§14.7.4.4): a number tree keyed by each page's
    // /StructParents value (= pageIndex here). The value is an array indexed by
    // MCID giving the StructElem that owns each marked-content sequence.
    const perPage = new Map<number, Array<PdfRef>>();
    for (const n of order) {
      for (const m of n.mcrefs) {
        let arr = perPage.get(m.pageIndex);
        if (!arr) {
          arr = [];
          perPage.set(m.pageIndex, arr);
        }
        arr[m.mcid] = n.ref!;
      }
    }
    const sortedKeys = [...perPage.keys()].sort((a, b) => a - b);
    const nums: Array<PdfValue> = [];
    for (const k of sortedKeys) {
      const arr = perPage.get(k)!;
      const dense: Array<PdfValue> = [];
      for (let i = 0; i < arr.length; i++) dense.push(arr[i] ?? PDF_NULL);
      nums.push(k, dense);
    }
    const parentTreeRef = doc.add(dict({ Nums: nums }));

    rootDict.set('K', this.root.ref!);
    rootDict.set('ParentTree', parentTreeRef);
    rootDict.set(
      'ParentTreeNextKey',
      sortedKeys.length > 0 ? sortedKeys[sortedKeys.length - 1]! + 1 : 0,
    );
    return rootRef;
  }
}
