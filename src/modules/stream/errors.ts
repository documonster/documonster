/**
 * Stream module error types.
 *
 * All stream-related errors extend StreamError.
 */

import type { BaseErrorOptions } from "@utils/errors";
import { BaseError } from "@utils/errors";

// Re-export abort utility used by browser/readable
export { createAbortError } from "@utils/errors";

/**
 * Base class for all stream-related errors.
 */
export class StreamError extends BaseError {
  constructor(message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.name = "StreamError";
  }
}

/**
 * Check if an error is a StreamError.
 */
export function isStreamError(err: unknown): err is StreamError {
  return err instanceof StreamError;
}

/**
 * Error thrown when a stream operation fails due to invalid state.
 */
export class StreamStateError extends StreamError {
  override name = "StreamStateError";

  constructor(
    public readonly operation: string,
    public readonly state: string,
    options?: BaseErrorOptions
  ) {
    super(`Cannot ${operation}: ${state}`, options);
  }
}

/**
 * Error thrown when data type conversion fails.
 */
export class StreamTypeError extends StreamError {
  override name = "StreamTypeError";

  constructor(
    public readonly expectedType: string,
    public readonly actualType: string,
    options?: BaseErrorOptions
  ) {
    super(`Expected ${expectedType}, got ${actualType}`, options);
  }
}

/**
 * Error thrown when a stream type is not supported.
 */
export class UnsupportedStreamTypeError extends StreamError {
  override name = "UnsupportedStreamTypeError";

  constructor(
    public readonly operation: string,
    public readonly streamType: string,
    options?: BaseErrorOptions
  ) {
    super(`${operation}: unsupported stream type "${streamType}"`, options);
  }
}
