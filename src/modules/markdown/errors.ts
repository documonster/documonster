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
 * Error thrown when a Markdown function is given an argument or option it cannot use — a value of the
 * wrong type, outside its range, or with no reading — before anything is read or written.
 * `field` names it, e.g. `options.delimiter`. A subclass of {@link MarkdownError}, so a
 * `catch` or `isMarkdownError` check that matched before still matches. Problems in the input
 * data itself are not argument errors.
 */
export class MarkdownArgumentError extends MarkdownError {
  override name = "MarkdownArgumentError";
  /** The argument or option at fault, e.g. `options.delimiter` or `rows[2]`. */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
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
