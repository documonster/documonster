/**
 * Formula module error types.
 */

import { BaseError } from "@utils/errors";
import type { BaseErrorOptions } from "@utils/errors";

// =============================================================================
// Base Error
// =============================================================================

/**
 * Base class for all formula-engine errors.
 */
export class FormulaError extends BaseError {
  override name = "FormulaError";
}

/**
 * Check if an error is a FormulaError.
 */
export function isFormulaError(err: unknown): err is FormulaError {
  return err instanceof FormulaError;
}

/**
 * Error thrown when `Formula.tokenize`, `parse` or `format` is given an argument or option it cannot use — a value of the
 * wrong type, outside its range, or with no reading — before anything is read or written.
 * `field` names it, e.g. `options.delimiter`. A subclass of {@link FormulaError}, so a
 * `catch` or `isFormulaError` check that matched before still matches. Problems in the input
 * data itself are not argument errors.
 */
export class FormulaArgumentError extends FormulaError {
  override name = "FormulaArgumentError";
  /** The argument or option at fault, e.g. `options.delimiter` or `rows[2]`. */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

// =============================================================================
// Parse Errors
// =============================================================================

/**
 * Error thrown when tokenizing or parsing a formula fails (syntax error).
 *
 * `position` is the 0-based character offset into the source text where the error was
 * detected. It is absent when there is no text to point into — `Formula.parse` given tokens —
 * and the message names the token instead.
 */
export class FormulaParseError extends FormulaError {
  override name = "FormulaParseError";
  /** 0-based character offset into the source text, when the source was text. */
  readonly position?: number;
  /**
   * The message without its position suffix.
   *
   * @internal
   */
  declare readonly detail: string;

  constructor(message: string, position?: number, options?: BaseErrorOptions) {
    super(position !== undefined ? `${message} (at position ${position})` : message, options);
    this.position = position;
    // Not enumerable, so `toJSON` (which serialises own fields) does not repeat the message.
    Object.defineProperty(this, "detail", { value: message, enumerable: false });
  }
}
