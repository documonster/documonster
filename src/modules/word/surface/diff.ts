/**
 * `Diff` namespace surface — document diffing.
 *
 * `import { Diff } from "@cj-tech-master/excelts/word"` →
 *   `Diff.documents(oldDoc, newDoc)` — tree-shaken via `export * as Diff`.
 */
export { diffDocuments as documents } from "@word/advanced/diff";
