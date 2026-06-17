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

// =============================================================================
// Parse Errors
// =============================================================================

/**
 * Error thrown when tokenizing or parsing a formula fails (syntax error).
 *
 * `position` is the 0-based character offset into the source string where the
 * error was detected, when known.
 */
export class FormulaParseError extends FormulaError {
  override name = "FormulaParseError";
  readonly position?: number;

  constructor(message: string, position?: number, options?: BaseErrorOptions) {
    super(position !== undefined ? `${message} (at position ${position})` : message, options);
    this.position = position;
  }
}
