/**
 * CSV module error types.
 */

import { BaseError } from "@utils/errors";
import type { BaseErrorOptions } from "@utils/errors";

/**
 * Base class for all CSV-related errors.
 */
export class CsvError extends BaseError {
  override name = "CsvError";
}

/**
 * Error thrown when a CSV function is given an argument or option it cannot use — a value of the
 * wrong type, outside its range, or with no reading — before anything is read or written.
 * `field` names it, e.g. `options.delimiter`. A subclass of {@link CsvError}, so a
 * `catch` or `isCsvError` check that matched before still matches. Problems in the input
 * data itself are not argument errors.
 */
export class CsvArgumentError extends CsvError {
  override name = "CsvArgumentError";
  /** The argument or option at fault, e.g. `options.delimiter` or `rows[2]`. */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

/**
 * Check if an error is a CsvError.
 */
export function isCsvError(err: unknown): err is CsvError {
  return err instanceof CsvError;
}

/**
 * Error thrown when CSV worker operations fail.
 */
export class CsvWorkerError extends CsvError {
  override name = "CsvWorkerError";
}
