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

  // Stream writer types
  WritableTarget,

  // Parse options
  XmlParseOptions,

  // Conversion options
  ToPlainObjectOptions,
  ParseXmlToObjectOptions
} from "./types";

// =============================================================================
// Encoding / Decoding
// =============================================================================

export {
  xmlEncode,
  xmlDecode,
  xmlEncodeAttr,
  validateXmlName,
  encodeCData,
  validateCommentText
} from "./encode";

// =============================================================================
// Writers
// =============================================================================

export { XmlWriter, StdDocAttributes } from "./writer";
export { XmlStreamWriter } from "./stream-writer";

// =============================================================================
// Parsers
// =============================================================================

export { SaxParser, parseSax, saxStream } from "./sax";
export { parseXml, findChild, findChildren, textContent, attr, walk, toPlainObject } from "./dom";
export { parseXmlToObject } from "./to-object";
export { query, queryAll } from "./query";

// =============================================================================
// Errors
// =============================================================================

export { XmlError, XmlParseError, XmlWriteError, isXmlError, isXmlParseError } from "./errors";
