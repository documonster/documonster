/**
 * `Validation` namespace surface — document structure validation.
 *
 * `import { Validation } from "documonster/word"` →
 *   `Validation.document(doc)` — tree-shaken via `export * as Validation`.
 */
export { validateDocument as document } from "../advanced/validation";
