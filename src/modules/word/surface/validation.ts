/**
 * `Validation` namespace surface — document structure validation.
 *
 * `import { Validation } from "documonster/word"` →
 *   `Validation.validate(doc)` — tree-shaken via `export * as Validation`.
 */
import { validateDocument } from "@word/advanced/validation";

export { validateDocument as validate } from "@word/advanced/validation";

/** @deprecated Use {@link validate}. Removed in 1.0. */
export const document = validateDocument;
