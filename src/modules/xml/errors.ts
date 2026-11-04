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

/**
 * Error thrown when an XML function, parser or writer is given an argument or option it cannot use — a value of the
 * wrong type, outside its range, or with no reading — before anything is read or written.
 * `field` names it, e.g. `options.delimiter`. A subclass of {@link XmlError}, so a
 * `catch` or `isXmlError` check that matched before still matches. Problems in the input
 * data itself are not argument errors.
 */
export class XmlArgumentError extends XmlError {
  override name = "XmlArgumentError";
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

/** Resource limit that stopped a parse — see {@link XmlParseError.limit}. */
export type XmlParseLimit = "depth" | "entityExpansions";

/**
 * Error thrown during XML parsing (SAX or DOM).
 */
export class XmlParseError extends XmlError {
  override name = "XmlParseError";
  /** 1-based line of the error, when the parser tracks positions. */
  readonly line?: number;
  /** 0-based column of the error, when the parser tracks positions. */
  readonly column?: number;
  /** The `fileName` the parser was given, if any. */
  readonly fileName?: string;
  /**
   * Set when parsing stopped because a configured resource limit was hit
   * (`maxDepth` or `maxEntityExpansions`) rather than because the input is
   * malformed, so callers can react to it without matching the message.
   */
  readonly limit?: XmlParseLimit;

  constructor(
    message: string,
    context?: { line?: number; column?: number; fileName?: string; limit?: XmlParseLimit },
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
    this.limit = context?.limit;
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
  /** What was attempted, e.g. `"close node"`. */
  readonly operation: string;
  /** Why the writer could not do it, e.g. `"no element is open"`. */
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
