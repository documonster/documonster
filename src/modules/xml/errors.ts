/**
 * XML module error types.
 */

import { BaseError } from "@utils/errors";
import type { BaseErrorOptions } from "@utils/errors";

// =============================================================================
// Base Error
// =============================================================================

/**
 * Base class for all XML-related errors.
 */
export class XmlError extends BaseError {
  override name = "XmlError";
}

// =============================================================================
// Parse Errors
// =============================================================================

/**
 * Error thrown during XML parsing (SAX or DOM).
 */
export class XmlParseError extends XmlError {
  override name = "XmlParseError";
  readonly line?: number;
  readonly column?: number;
  readonly fileName?: string;

  constructor(
    message: string,
    context?: { line?: number; column?: number; fileName?: string },
    options?: BaseErrorOptions
  ) {
    const parts: string[] = [];
    if (context?.fileName) {
      parts.push(context.fileName);
    }
    if (context?.line !== undefined) {
      parts.push(`${context.line}:${context.column ?? 0}`);
    }
    const prefix = parts.length > 0 ? `${parts.join(":")}: ` : "";
    super(`${prefix}${message}`, options);
    this.line = context?.line;
    this.column = context?.column;
    this.fileName = context?.fileName;
  }
}

// =============================================================================
// Write Errors
// =============================================================================

/**
 * Error thrown during XML writing when the writer is in an invalid state.
 */
export class XmlWriteError extends XmlError {
  override name = "XmlWriteError";
  readonly operation: string;
  readonly state: string;

  constructor(operation: string, state: string, options?: BaseErrorOptions) {
    super(`Cannot ${operation}: ${state}`, options);
    this.operation = operation;
    this.state = state;
  }
}

// =============================================================================
// Type Guards
// =============================================================================

/** Check if an error is an XmlError. */
export function isXmlError(err: unknown): err is XmlError {
  return err instanceof XmlError;
}

/** Check if an error is an XmlParseError. */
export function isXmlParseError(err: unknown): err is XmlParseError {
  return err instanceof XmlParseError;
}
