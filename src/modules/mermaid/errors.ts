/**
 * Mermaid module error types.
 */

import { BaseError } from "@utils/errors";
import type { BaseErrorOptions } from "@utils/errors";

/**
 * Base class for all Mermaid-related errors.
 */
export class MermaidError extends BaseError {
  override name = "MermaidError";
}

/**
 * Error thrown when a Mermaid function is given an argument or option with no valid reading —
 * source that is not text, an options object of the wrong shape, a theme colour that cannot
 * be read, a value that is not a diagram. Thrown by the call that received it, with a message
 * that starts with the call as the caller wrote it (`Mermaid.toSvg: …`). Mermaid *syntax* is
 * reported by `MermaidParseError` (and unsupported statements by `skipped`), not by this.
 */
export class MermaidArgumentError extends MermaidError {
  override name = "MermaidArgumentError";
  /**
   * The argument or field at fault, as a path from the parameter name in the public
   * signature: `source`, `options.fontSize`, `options.theme.nodeFill`.
   */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

/** Options of {@link MermaidParseError}: where in the source the problem is, when known. */
export interface MermaidParseErrorOptions extends BaseErrorOptions {
  /** 1-based line of the source. */
  line?: number;
  /** 1-based column of `line`. */
  column?: number;
}

/**
 * Thrown when the source names no diagram this module can draw — the one Mermaid syntax
 * problem that is an error; a statement the parser does not understand is reported in the
 * diagram's `skipped` list instead. `line` and `column` (1-based) locate the statement that
 * names the diagram, when there is one: an empty source has none.
 */
export class MermaidParseError extends MermaidError {
  override name = "MermaidParseError";
  /** 1-based line of the source the problem is on, when known. */
  readonly line: number | undefined;
  /** 1-based column of {@link line}, when known. */
  readonly column: number | undefined;

  constructor(message: string, options?: MermaidParseErrorOptions) {
    super(message, options);
    this.line = options?.line;
    this.column = options?.column;
  }
}

/**
 * @deprecated Use {@link MermaidParseError} — the same class, under the name every module's
 * parse error has (`XmlParseError`, `MarkdownParseError`, …), so `instanceof` either answers
 * alike. Removed in 1.0.
 */
export const MermaidSyntaxError = MermaidParseError;
/** @deprecated Use {@link MermaidParseError}. Removed in 1.0. */
export type MermaidSyntaxError = MermaidParseError;

/**
 * Check if an error is a MermaidError.
 */
export function isMermaidError(err: unknown): err is MermaidError {
  return err instanceof MermaidError;
}
