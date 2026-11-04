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

/** Options for {@link StreamTypeError} and {@link UnsupportedStreamTypeError}. */
export interface StreamErrorOptions extends BaseErrorOptions {
  /**
   * The argument or option at fault, when the error is about one: the path from the parameter
   * name in the public signature (`stream`, `options.highWaterMark`).
   */
  field?: string;
  /**
   * The public call that refused the argument (`Streams.fromString`), which a
   * {@link StreamTypeError} message then starts with.
   */
  where?: string;
}

/**
 * Error thrown when data type conversion fails, or when an argument is of the wrong type —
 * then `field` names the argument or option.
 */
export class StreamTypeError extends StreamError {
  override name = "StreamTypeError";
  /** The argument or option at fault; absent when the error is about a stream's contents. */
  readonly field?: string;

  constructor(
    public readonly expectedType: string,
    public readonly actualType: string,
    options?: StreamErrorOptions
  ) {
    super(
      options?.where === undefined
        ? `Expected ${expectedType}, got ${actualType}`
        : `${options.where}: expected ${expectedType}, got ${actualType}`,
      options
    );
    if (options?.field !== undefined) {
      this.field = options.field;
    }
  }
}

/**
 * Error thrown when a stream type is not supported. `operation` is the public call that was
 * given it (`Streams.bytes`), which the message starts with.
 */
export class UnsupportedStreamTypeError extends StreamError {
  override name = "UnsupportedStreamTypeError";

  /** The argument at fault, when the error is about one; absent for a stream's contents. */
  readonly field?: string;

  constructor(
    public readonly operation: string,
    public readonly streamType: string,
    options?: StreamErrorOptions
  ) {
    super(`${operation}: unsupported stream type "${streamType}"`, options);
    if (options?.field !== undefined) {
      this.field = options.field;
    }
  }
}
