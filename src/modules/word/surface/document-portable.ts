/**
 * The members of the `Document` namespace both platforms publish: building and
 * editing a document (`Document.create`, `Document.addParagraph`, …), reading
 * one from bytes (`Document.read`) and writing it (`Document.toBuffer`,
 * `Document.toBase64`, `Document.toFlatOpc`). `document.ts` adds the Node-only
 * file-path members to them; `document.browser.ts` publishes them alone.
 */
export * from "@word/builder/document-handle";
export { readDocx as read } from "@word/reader/docx-reader";
export { toBuffer, toBase64, toFlatOpc } from "@word/document-io";
