/**
 * `Validation` namespace surface — document structure validation.
 *
 * `import { Validation } from "@cj-tech-master/excelts/word"` →
 *   `Validation.document(doc)` — tree-shaken via `export * as Validation`.
 */
export { validateDocument as document } from "@word/advanced/validation";
