/**
 * Markdown module error types.
 */

import { BaseError } from "@utils/errors";

/**
 * Base class for all Markdown-related errors.
 */
export class MdError extends BaseError {
  override name = "MdError";
}

/**
 * Error thrown when Markdown parsing fails.
 */
export class MdParseError extends MdError {
  override name = "MdParseError";

  /** 1-based line number where the error occurred */
  readonly line: number;

  constructor(message: string, line: number, options?: { cause?: unknown }) {
    super(`Line ${line}: ${message}`, options);
    this.line = line;
  }
}
