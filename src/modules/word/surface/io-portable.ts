/**
 * The members of the deprecated `Io` namespace both platforms publish. Every one
 * is a forward: reading and writing a document live in `Document`
 * (`Document.read`, `Document.toBuffer`, `Document.toBase64`,
 * `Document.toFlatOpc`). `io.ts` adds the Node-only file-path members to them;
 * `io.browser.ts` publishes them alone.
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
  toBase64 as toBase64Impl,
  toBuffer as toBufferImpl,
  toFlatOpc as toFlatOpcImpl
} from "@word/document-io";
import {
  editDocxIncremental as editDocxIncrementalImpl,
  listDocxParts as listDocxPartsImpl,
  readDocxPart as readDocxPartImpl
} from "@word/incremental-edit";
import { mergeDocuments } from "@word/query/merge";
import { splitDocument } from "@word/query/split";
import { readDocx } from "@word/reader/docx-reader";

/** @deprecated Use `Document.read`. Removed in 1.0. */
export const read = readDocx;
/** @deprecated Use `Document.toBuffer`. Removed in 1.0. */
export const toBuffer = toBufferImpl;
/** @deprecated Use `Document.toBase64`. Removed in 1.0. */
export const toBase64 = toBase64Impl;
/** @deprecated Use `Document.toFlatOpc`. Removed in 1.0. */
export const toFlatOpc = toFlatOpcImpl;

/**
 * @deprecated Use `Document.toFlatOpc`, which takes an options object
 * (`{ compressionLevel }`). Removed in 1.0.
 */
export const toFlatOpcFromDoc = toFlatOpcFromDocImpl;

/** @deprecated Use `Document.toBuffer`. Removed in 1.0. */
const packageDocument = toBufferImpl;
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
