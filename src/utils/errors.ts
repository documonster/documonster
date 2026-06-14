/**
 * Base error classes and utility functions for error handling.
 *
 * Module-specific errors should extend BaseError and live in their own modules.
 *
 * @example
 * ```typescript
 * // Catching and identifying errors
 * try {
 *   await Workbook.readXlsxFile(workbook, 'test.xlsx');
 * } catch (e) {
 *   if (isExcelError(e)) {
 *     console.error(`Excel error: ${e.message}`);
 *     if (e.cause) console.error('Caused by:', e.cause);
 *   }
 * }
 *
 * // Creating errors with cause chain
 * throw new ExcelError('Failed to parse', { cause: originalError });
 * ```
 */

// =============================================================================
// Base Error Class
// =============================================================================

/**
 * Options for BaseError constructor.
 */
export interface BaseErrorOptions {
  /** The original error that caused this error (ES2022 Error Cause) */
  cause?: unknown;
}

/**
 * Base class for all library errors.
 * Module-specific errors should extend this class.
 *
 * Features:
 * - Supports ES2022 error cause for error chaining
 * - Properly captures stack trace
 * - Sets correct prototype for instanceof checks
 * - JSON serialization support for logging
 */
export class BaseError extends Error {
  constructor(message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.name = "BaseError";
    // Fix prototype chain for ES5 environments
    Object.setPrototypeOf(this, new.target.prototype);
    // Capture stack trace (V8 engines)
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, this.constructor);
    }
  }

  /**
   * Serialize error for logging/transmission.
   * Includes cause chain for debugging.
   */
  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      message: this.message,
      stack: this.stack,
      cause: this.cause instanceof Error ? errorToJSON(this.cause) : this.cause
    };
  }
}

// =============================================================================
// Common Error Classes
// =============================================================================

/**
 * Error thrown when an operation is aborted.
 */
export class AbortError extends BaseError {
  override name = "AbortError";
  readonly code = "ABORT_ERR";

  constructor(reason?: unknown) {
    super("The operation was aborted", reason !== undefined ? { cause: reason } : undefined);
  }
}

// =============================================================================
// Helper Functions
// =============================================================================

/**
 * Create an abort error from a reason.
 */
export function createAbortError(reason?: unknown): AbortError {
  if (reason instanceof AbortError) {
    return reason;
  }
  return new AbortError(reason);
}

/**
 * Check if an error is an abort error.
 */
export function isAbortError(err: unknown): err is { name: string } {
  return !!err && typeof err === "object" && (err as any).name === "AbortError";
}

/**
 * Throw if the signal is aborted.
 */
export function throwIfAborted(signal?: AbortSignal, reason?: unknown): void {
  if (!signal) {
    return;
  }
  if (!signal.aborted) {
    return;
  }
  const r = reason ?? (signal as any).reason;
  throw createAbortError(r);
}

/**
 * Create a linked AbortController that aborts when the parent signal aborts.
 *
 * @param parentSignal - Optional parent signal to link to
 * @returns Controller and cleanup function
 */
export function createLinkedAbortController(parentSignal?: AbortSignal): {
  controller: AbortController;
  cleanup: () => void;
} {
  const controller = new AbortController();

  if (!parentSignal) {
    return { controller, cleanup: () => {} };
  }

  if (parentSignal.aborted) {
    controller.abort((parentSignal as any).reason);
    return { controller, cleanup: () => {} };
  }

  const onAbort = (): void => {
    try {
      controller.abort((parentSignal as any).reason);
    } catch {
      controller.abort();
    }
  };

  parentSignal.addEventListener("abort", onAbort, { once: true });

  const cleanup = (): void => {
    try {
      parentSignal.removeEventListener("abort", onAbort);
    } catch {
      // ignore
    }
  };

  return { controller, cleanup };
}

// =============================================================================
// Error Normalization
// =============================================================================

/**
 * Convert an unknown value to an Error.
 *
 * If the value is already an Error, it's returned as-is.
 * Otherwise, it's converted to a string and wrapped in an Error.
 *
 */
export function toError(err: unknown): Error {
  return err instanceof Error ? err : new Error(String(err));
}

// =============================================================================
// Promise Utilities
// =============================================================================

/**
 * Suppress unhandled rejection warnings for a promise.
 *
 * Use this when you intentionally want to ignore a promise's rejection,
 * typically for fire-and-forget cleanup operations.
 */
export function suppressUnhandledRejection(promise: Promise<unknown>): void {
  promise.catch(() => {});
}

// =============================================================================
// Error Serialization
// =============================================================================

/**
 * Serialize any Error to a plain object for logging/transmission.
 * Handles both BaseError and native Error instances.
 */
export function errorToJSON(err: Error): Record<string, unknown> {
  if (err instanceof BaseError) {
    return err.toJSON();
  }
  return {
    name: err.name,
    message: err.message,
    stack: err.stack,
    cause: err.cause instanceof Error ? errorToJSON(err.cause) : err.cause
  };
}

/**
 * Get the full error chain as an array.
 * Useful for logging all errors in a cause chain.
 *
 * @example
 * ```typescript
 * const chain = getErrorChain(error);
 * chain.forEach((e, i) => console.log(`${i}: ${e.message}`));
 * ```
 */
export function getErrorChain(err: Error): Error[] {
  const chain: Error[] = [err];
  let current: unknown = err.cause;
  while (current instanceof Error) {
    chain.push(current);
    current = current.cause;
  }
  return chain;
}

/**
 * Get the root cause of an error chain.
 * Returns the deepest error in the cause chain.
 */
export function getRootCause(err: Error): Error {
  let current: Error = err;
  while (current.cause instanceof Error) {
    current = current.cause;
  }
  return current;
}
