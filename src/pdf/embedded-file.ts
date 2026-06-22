// ISO 32000-1 §7.11 (embedded file streams + file specifications) and
// ISO 19005-3 §6.8 (associated files). PDF/A-3 permits embedding any file —
// e.g. the source .docx/.xlsx — alongside the rendered PDF. Each attachment is
// an /EmbeddedFile stream wrapped in a /Filespec that carries an
// /AFRelationship; the catalog lists the filespecs in /AF (associated files)
// and a /Names /EmbeddedFiles name tree (so it appears in the Attachments UI).

import type { PdfRef, PdfValue } from '@/pdf/objects';
import type { PdfDocument } from '@/pdf/writer';
import { dict, name, ref, stream, unicodeString } from '@/pdf/objects';

/** How an embedded file relates to the document (`/AFRelationship`, ISO 32000-2 §14.13). */
export type AFRelationship = 'Source' | 'Data' | 'Alternative' | 'Supplement' | 'Unspecified';

/** One file to embed as a PDF/A-3 associated file (see {@link embedAssociatedFile}). */
export interface AttachedFile {
  /** File name as shown to the user, e.g. `"source.docx"`. ASCII recommended. */
  readonly name: string;
  readonly bytes: Uint8Array;
  /**
   * MIME type for the `/EmbeddedFile` `/Subtype`, e.g.
   * `"application/vnd.openxmlformats-officedocument.wordprocessingml.document"`.
   */
  readonly mimeType: string;
  readonly relationship?: AFRelationship;
  readonly description?: string;
}

/**
 * Embed `file` as an `/EmbeddedFile` stream wrapped in a `/Filespec` (ISO 32000-1
 * §7.11, ISO 19005-3 §6.8). The MIME type becomes a `Name` (the serializer
 * hex-escapes the `/`), which readers decode back.
 *
 * @param file The file bytes plus its name, MIME type and relationship.
 * @returns A reference to the `/Filespec`, to place in the catalog `/AF` array
 *   and the `/Names` `/EmbeddedFiles` name tree.
 */
export function embedAssociatedFile(doc: PdfDocument, file: AttachedFile): PdfRef {
  const efRef = doc.add(
    stream(
      {
        Type: name('EmbeddedFile'),
        Subtype: name(file.mimeType),
        Params: dict({ Size: file.bytes.byteLength }),
      },
      file.bytes,
    ),
  );
  const filespec: Record<string, PdfValue> = {
    Type: name('Filespec'),
    F: file.name,
    UF: unicodeString(file.name),
    EF: dict({ F: ref(efRef.id), UF: ref(efRef.id) }),
    AFRelationship: name(file.relationship ?? 'Unspecified'),
  };
  if (file.description) filespec['Desc'] = unicodeString(file.description);
  return doc.add(dict(filespec));
}
