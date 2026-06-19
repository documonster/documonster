/**
 * CSV module error types.
 */

import { BaseError } from "@utils/errors";

/**
 * Base class for all CSV-related errors.
 */
export class CsvError extends BaseError {
  override name = "CsvError";
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
