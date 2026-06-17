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
  SaxTag,
  SaxEvent,
  SaxEventAny,
  SaxHandlers,
  SaxOptions,

  // Invalid character handling
  InvalidCharHandling,

  // Stream writer types
  WritableTarget,

  // Parse options
  XmlParseOptions,

  // Conversion options
  ToPlainObjectOptions,
  ParseXmlToObjectOptions
} from "./types";

// =============================================================================
// `Xml` domain namespace — encode/decode, writers, parsers, query
// (tree-shaken via `export * as`)
// =============================================================================

export * as Xml from "./surface/xml";

// =============================================================================
// Errors
// =============================================================================

export { XmlError, XmlParseError, XmlWriteError, isXmlError, isXmlParseError } from "./errors";
