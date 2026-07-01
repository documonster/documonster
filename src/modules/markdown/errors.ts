/**
 * Markdown module error types.
 */

import { BaseError } from "@utils/errors";
import type { BaseErrorOptions } from "@utils/errors";

/**
 * Base class for all Markdown-related errors.
 */
export class MarkdownError extends BaseError {
  override name = "MarkdownError";
}

/**
 * Check if an error is a MarkdownError.
 */
export function isMarkdownError(err: unknown): err is MarkdownError {
  return err instanceof MarkdownError;
}

/**
 * Error thrown when Markdown parsing fails.
 */
export class MarkdownParseError extends MarkdownError {
  override name = "MarkdownParseError";

  /** 1-based line number where the error occurred */
  readonly line: number;

  constructor(message: string, line: number, options?: BaseErrorOptions) {
    super(`Line ${line}: ${message}`, options);
    this.line = line;
  }
}
