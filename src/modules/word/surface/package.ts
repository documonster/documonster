/**
 * `Package` namespace surface — operations that take `.docx` bytes and return
 * `.docx` bytes (or a part of them), for callers that do not want to hold a
 * `DocxDocument` themselves.
 *
 * `listParts`, `readPart` and `edit` work on the package's parts and leave the
 * rest of the file as it was. `fillTemplate`, `patch` and `compileTemplate`
 * read the whole document model (as `Document.read` does) and write it out again.
 *
 * Every one of them reads its bytes as `Document.read` does — a ZIP or Flat OPC
 * package, an encrypted one decrypted with `{ password }`, under
 * `{ securityPolicy }` — and the writing ones take `Document.toBuffer`'s
 * `compressionLevel` and `modTime`. What they write is never encrypted.
 *
 * Two template engines, for two kinds of template:
 * - `fillTemplate` is `Template.fill` on bytes: `{{variables}}`, `{{#if}}`
 *   conditionals and `{{#each}}` loops resolved from one data object, comment
 *   bodies included. Use it when the template carries logic.
 * - `patch` / `patchTemplate` replace each named placeholder exactly with
 *   text, a paragraph, a table or an image — no data object, no conditionals
 *   or loops, comments left alone. Use it when you only swap placeholders for
 *   ready-made content.
 *
 * `import { Package } from "documonster/word"` →
 *   `Package.listParts(bytes)`, `Package.readPart(bytes, path)`,
 *   `Package.edit(bytes, edits)`, `Package.fillTemplate(bytes, data)`,
 *   `Package.patch(bytes, patches)`, `Package.compileTemplate(bytes)` /
 *   `Package.patchTemplate(compiled, patches)` — tree-shaken via `export * as Package`.
 */
export {
  editDocxIncremental as edit,
  listDocxParts as listParts,
  readDocxPart as readPart
} from "@word/incremental-edit";
export {
  fillBuffer as fillTemplate,
  compileTemplate,
  patchTemplate,
  patchDocument as patch
} from "@word/document-io";
