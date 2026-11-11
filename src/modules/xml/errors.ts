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

/**
 * Resource limit that stopped a parse — see {@link XmlParseError.limit}. `"prologContext"` is
 * `maxContextBytes`: the size of the prolog context — the entities and attribute declarations a
 * DOCTYPE internal subset makes, gathered with `internalEntities` — or a markup declaration of the
 * subset too long to hold and check. `"rootTag"` is `Xml.readRoot`'s `maxStartByte` or `maxTagBytes`.
 */
export type XmlParseLimit =
  | "depth"
  | "entityExpansions"
  | "entityExpansionLength"
  | "prologContext"
  | "rootTag";

/** Where a parse error was found: what {@link XmlParseError} carries. */
export interface XmlParseErrorContext {
  /** 1-based line, when the parser tracks positions. */
  line?: number;
  /** 0-based column, when the parser tracks positions. */
  column?: number;
  /** The `fileName` the parser was given. */
  fileName?: string;
  /** Offset of the byte in the file, when the parser is given bytes. */
  byte?: number;
}

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
  /**
   * Offset in the file of the byte where the error was found, when the parser was given bytes
   * (the parser's `byteOffset` counts toward it). Also named at the end of the message.
   */
  readonly byte?: number;

  constructor(
    message: string,
    context?: XmlParseErrorContext & { limit?: XmlParseLimit },
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
    const suffix = context?.byte !== undefined ? ` (at byte ${context.byte})` : "";
    super(`${prefix}${message}${suffix}`, options);
    this.line = context?.line;
    this.column = context?.column;
    this.fileName = context?.fileName;
    this.limit = context?.limit;
    this.byte = context?.byte;
  }
}

/**
 * Error thrown when parsing stopped because a configured resource limit was hit — `maxDepth`
 * (`limit: "depth"`), `maxEntityExpansions` (`limit: "entityExpansions"`) or
 * `maxEntityExpansionLength` (`limit: "entityExpansionLength"`), say — rather than because
 * the input is malformed. `allowed` is the configured value, as on the other modules'
 * `…LimitExceededError` classes. A subclass of {@link XmlParseError}, so a `catch` that matched
 * before still matches, and `.limit` reads as it always did.
 */
export class XmlLimitExceededError extends XmlParseError {
  override name = "XmlLimitExceededError";
  declare readonly limit: XmlParseLimit;
  /** The configured value of the limit that was hit. */
  readonly allowed: number;

  constructor(
    message: string,
    limit: XmlParseLimit,
    allowed: number,
    context?: XmlParseErrorContext,
    options?: BaseErrorOptions
  ) {
    super(message, { ...context, limit }, options);
    this.allowed = allowed;
  }
}

/**
 * Error thrown when the input ends where the document cannot: inside a tag, a comment or another
 * construct, with elements still open, or before the root element. A subclass of
 * {@link XmlParseError}, so a `catch` that matched before still matches.
 *
 * `Xml.readRecords` throws it with `recordStartByte` when the input ends inside a record the
 * call was asked for.
 */
export class XmlTruncatedError extends XmlParseError {
  override name = "XmlTruncatedError";
  /** Offset of the `<` of the record the input ended inside, when there is one. */
  readonly recordStartByte?: number;

  constructor(
    message: string,
    context?: XmlParseErrorContext & { recordStartByte?: number },
    options?: BaseErrorOptions
  ) {
    super(message, context, options);
    this.recordStartByte = context?.recordStartByte;
  }
}

// =============================================================================
// Chunk Errors
// =============================================================================

/**
 * Error thrown by `Xml.readRecords` when a byte range of a file cannot be read on its own with
 * certainty: reading from the range's start found other records, or other boundaries, than
 * reading through it from an earlier point did — a chunk that starts inside a record holding an
 * element of the same name, say, or inside a comment or CDATA section full of tags.
 *
 * It describes the split, not the file: read the range together with the one before it (or the
 * file from its start) and the records come out right. `cause` is the parse error that showed the
 * start was misplaced, when it was one.
 */
export class XmlChunkAmbiguityError extends XmlError {
  override name = "XmlChunkAmbiguityError";
  /** Offset of the chunk start that could not be placed. */
  readonly chunkStart: number;
  /** Offset of the byte where the two readings were found to differ. */
  readonly atByte: number;

  constructor(message: string, chunkStart: number, atByte: number, options?: BaseErrorOptions) {
    super(`${message} (chunk starting at byte ${chunkStart}, at byte ${atByte})`, options);
    this.chunkStart = chunkStart;
    this.atByte = atByte;
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
