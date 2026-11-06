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

/** Options for {@link StreamArgumentError} and {@link StreamUnsupportedTypeError}. */
export interface StreamErrorOptions extends BaseErrorOptions {
  /**
   * The argument or option at fault, when the error is about one: the path from the parameter
   * name in the public signature (`stream`, `options.highWaterMark`).
   */
  field?: string;
  /**
   * The public call that refused the argument (`Streams.fromString`), which a
   * {@link StreamArgumentError} message then starts with.
   */
  where?: string;
}

/**
 * The stream module's argument error: an argument, an option or a written chunk is not what
 * the call accepts. `field` names the argument or option, and `where` (when given) the public
 * call, which the message then starts with.
 */
export class StreamArgumentError extends StreamError {
  override name = "StreamArgumentError";
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

/** @deprecated Use {@link StreamArgumentError}. Removed in 1.0. */
export const StreamTypeError = StreamArgumentError;
/** @deprecated Use {@link StreamArgumentError}. Removed in 1.0. */
export type StreamTypeError = StreamArgumentError;

/**
 * Error thrown when a stream type is not supported. `operation` is the public call that was
 * given it (`Streams.bytes`), which the message starts with.
 */
export class StreamUnsupportedTypeError extends StreamError {
  override name = "StreamUnsupportedTypeError";

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

/** @deprecated Use {@link StreamUnsupportedTypeError}. Removed in 1.0. */
export const UnsupportedStreamTypeError = StreamUnsupportedTypeError;
/** @deprecated Use {@link StreamUnsupportedTypeError}. Removed in 1.0. */
export type UnsupportedStreamTypeError = StreamUnsupportedTypeError;
