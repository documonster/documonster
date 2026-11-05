/**
 * The members of the `Io` namespace both platforms publish — `Io.read(buffer)`,
 * `Io.toBuffer(doc)`, `Io.toBase64(doc)`, `Io.toFlatOpc(doc)`. `io.ts` adds the
 * Node-only file-path members to them; `io.browser.ts` publishes them alone.
 *
 * Part-level editing of DOCX bytes lives in `Package`, byte-level template
 * filling and placeholder patching in `Package`, in-memory filling in `Template`, whole-document transforms in `Transform`.
 */
import {
  updateFields as updateFieldsImpl,
  updateTableOfContents as updateTableOfContentsImpl
} from "@word/advanced/field-engine";
import {
  compileTemplate as compileTemplateImpl,
  toFlatOpcFromDoc as toFlatOpcFromDocImpl,
  fillTemplateFromBuffer as fillTemplateFromBufferImpl,
  patchDocument as patchDocumentImpl,
  patchTemplate as patchTemplateImpl,
  toBuffer
} from "@word/document-io";
import {
  editDocxIncremental as editDocxIncrementalImpl,
  listDocxParts as listDocxPartsImpl,
  readDocxPart as readDocxPartImpl
} from "@word/incremental-edit";
import { mergeDocuments } from "@word/query/merge";
import { splitDocument } from "@word/query/split";

export { readDocx as read } from "@word/reader/docx-reader";
export { toBuffer, toBase64, toFlatOpc } from "@word/document-io";

/**
 * @deprecated Use {@link toFlatOpc}, which takes an options object
 * (`{ compressionLevel }`). Removed in 1.0.
 */
export const toFlatOpcFromDoc = toFlatOpcFromDocImpl;

/** @deprecated Use {@link toBuffer}. Removed in 1.0. */
const packageDocument = toBuffer;
// `package` is a reserved word, so the deprecated name can only be published
// through an export specifier.
export { packageDocument as package };

/** @deprecated Use `Package.edit`. Removed in 1.0. */
export const editDocxIncremental = editDocxIncrementalImpl;
/** @deprecated Use `Package.listParts`. Removed in 1.0. */
export const listDocxParts = listDocxPartsImpl;
/** @deprecated Use `Package.readPart`. Removed in 1.0. */
export const readDocxPart = readDocxPartImpl;
/** @deprecated Use `Package.patch`. Removed in 1.0. */
export const patchDocument = patchDocumentImpl;
/** @deprecated Use `Package.compileTemplate`. Removed in 1.0. */
export const compileTemplate = compileTemplateImpl;
/** @deprecated Use `Package.patchTemplate`. Removed in 1.0. */
export const patchTemplate = patchTemplateImpl;
/** @deprecated Use `Package.fillTemplate`. Removed in 1.0. */
export const fillTemplateFromBuffer = fillTemplateFromBufferImpl;
/** @deprecated Use `Transform.updateFields`. Removed in 1.0. */
export const updateFields = updateFieldsImpl;
/** @deprecated Use `Transform.updateTableOfContents`. Removed in 1.0. */
export const updateTableOfContents = updateTableOfContentsImpl;
/** @deprecated Use `Transform.merge`. Removed in 1.0. */
export const merge = mergeDocuments;
/** @deprecated Use `Transform.split`. Removed in 1.0. */
export const split = splitDocument;
