/**
 * `Io` namespace surface — document serialization / deserialization,
 * incremental editing, and field updates.
 *
 * `import { Io } from "documonster/word"` →
 *   `Io.package(doc)`, `Io.read(buffer)`, `Io.toBuffer(doc)`,
 *   `Io.patchDocument(...)`, `Io.updateFields(doc)`, … — tree-shaken via
 *   `export * as Io`.
 */
export { packageDocx as package } from "@word/writer/docx-packager";
export { readDocx as read } from "@word/reader/docx-reader";
export {
  toBuffer,
  toBase64,
  patchDocument,
  compileTemplate,
  patchTemplate,
  fillTemplateFromBuffer,
  toFlatOpcFromDoc
} from "@word/document-io";
export { editDocxIncremental, listDocxParts, readDocxPart } from "@word/incremental-edit";
export { updateFields, updateTableOfContents } from "@word/advanced/field-engine";
export { mergeDocuments as merge } from "@word/query/merge";
export { splitDocument as split } from "@word/query/split";
