/**
 * `documonster/utils` — helpers for bytes and text that the platforms offer unevenly, with one
 * behaviour on Node, Bun and every browser this package supports.
 *
 * `import { Bytes } from "documonster/utils"` → `Bytes.fromBase64(text)`, `Bytes.concat(parts)`, …
 */

export * as Bytes from "./surface/bytes";

export type {
  UtilsBase64Alphabet,
  UtilsFromBase64Options,
  UtilsToBase64Options,
  UtilsToTextOptions
} from "./surface/types";

export { UtilsError, UtilsArgumentError, isUtilsError } from "./surface/errors";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "./errors";
