/**
 * `Document` namespace surface — create, read, edit and write a document.
 *
 * `import { Document } from "documonster/word"` →
 *   `Document.create()`, `Document.read(bytes)`, `Document.addParagraph(doc, …)`,
 *   `Document.toBuffer(doc)`, and on Node `Document.readFile(path)`,
 *   `Document.writeFile(doc, path)` — tree-shaken via `export * as Document`. The
 *   browser build resolves this module to `document.browser.ts`, which publishes
 *   no file-path member.
 */
export * from "@word/surface/document-portable";
export { readFile, writeFile } from "@word/document-io";
