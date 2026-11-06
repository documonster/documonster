/**
 * `Diff` namespace surface — document diffing.
 *
 * `import { Diff } from "documonster/word"` →
 *   `Diff.compare(oldDoc, newDoc)` — tree-shaken via `export * as Diff`.
 */
import { diffDocuments } from "@word/advanced/diff";

export { diffDocuments as compare } from "@word/advanced/diff";

/** @deprecated Use {@link compare}. Removed in 1.0. */
export const documents = diffDocuments;
