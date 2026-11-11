/**
 * XML Module - Public API
 *
 * Standalone XML processing module with zero dependencies on Excel or other modules.
 * Provides complete XML reading and writing capabilities, both streaming and buffered.
 *
 * Design principles:
 * - Zero external dependencies — fully self-contained
 * - Decoupled from the Excel module — can be used independently
 * - Dual-mode: streaming (SAX parser + stream writer) and buffered (DOM parser + writer)
 * - Shared XmlSink interface lets rendering code target both modes transparently
 */

// =============================================================================
// Core Types
// =============================================================================

export type {
  // Attributes
  XmlAttributes,

  // DOM node types
  XmlNodeType,
  XmlElement,
  XmlText,
  XmlCData,
  XmlComment,
  XmlProcessingInstruction,
  XmlNode,
  XmlDocument,

  // Writer sink interface
  XmlSink,

  // SAX types
  XmlSaxTag,
  XmlSaxEvent,
  XmlSaxEventAny,
  XmlSaxHandlers,
  XmlSaxOptions,

  // Invalid character handling
  XmlInvalidCharHandling,

  // Stream writer types
  XmlWritableTarget,

  // Parse options
  XmlParseOptions,

  // Byte readers
  XmlByteSource,
  XmlPrologContext,
  XmlAttributeDeclaration,
  XmlRecord,
  XmlReadRecordsOptions,
  XmlReadRecordsResult,
  XmlPartitionStructure,
  XmlPartitionStep,
  XmlReadRootOptions,
  XmlRootElement,

  // Conversion options
  XmlToPlainObjectOptions,
  XmlParseToObjectOptions,
  ToPlainObjectOptions,
  ParseXmlToObjectOptions,
  InvalidCharHandling,
  WritableTarget,
  SaxTag,
  SaxEvent,
  SaxEventAny,
  SaxHandlers,
  SaxOptions
} from "@xml/types";

// =============================================================================
// `Xml` domain namespace — encode/decode, writers, parsers, query
// (tree-shaken via `export * as`)
// =============================================================================

export * as Xml from "@xml/surface/xml";

// =============================================================================
// Errors
// =============================================================================

export {
  XmlError,
  XmlArgumentError,
  XmlParseError,
  XmlLimitExceededError,
  XmlTruncatedError,
  XmlChunkAmbiguityError,
  XmlWriteError,
  isXmlError,
  isXmlParseError
} from "@xml/errors";
// What `Xml.readRecords` and `Xml.readRoot` reject with when their signal fires, and its guard (also
// true for a native `DOMException` named "AbortError").
export { AbortError, isAbortError } from "@utils/errors";
export type { XmlParseErrorContext, XmlParseLimit } from "@xml/errors";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
