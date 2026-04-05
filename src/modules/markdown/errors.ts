/**
 * Markdown module error types.
 */

import { BaseError } from "@utils/errors";

/**
 * Base class for all Markdown-related errors.
 */
export class MarkdownError extends BaseError {
  override name = "MarkdownError";
}

/**
 * Error thrown when Markdown parsing fails.
 */
export class MarkdownParseError extends MarkdownError {
  override name = "MarkdownParseError";

  /** 1-based line number where the error occurred */
  readonly line: number;

  constructor(message: string, line: number, options?: { cause?: unknown }) {
    super(`Line ${line}: ${message}`, options);
    this.line = line;
  }
}
