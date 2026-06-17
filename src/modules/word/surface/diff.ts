/**
 * `Diff` namespace surface — document diffing.
 *
 * `import { Diff } from "documonster/word"` →
 *   `Diff.documents(oldDoc, newDoc)` — tree-shaken via `export * as Diff`.
 */
export { diffDocuments as documents } from "../advanced/diff";
