/**
 * XML Module - Core Types
 *
 * Centralized type definitions for the XML module.
 * All interfaces, types, and type utilities live here.
 */

// =============================================================================
// XML Attributes
// =============================================================================

/** XML attribute map: attribute name to string value. */
export type XmlAttributes = Record<string, string | number | boolean | undefined>;

// =============================================================================
// XML DOM Node Types
// =============================================================================

/** Discriminated union tag for XML node types. */
export type XmlNodeType = "element" | "text" | "cdata" | "comment" | "processing-instruction";

/** Base interface shared by all XML node types. */
interface XmlNodeBase {
  readonly type: XmlNodeType;
}

/** XML element node. */
export interface XmlElement extends XmlNodeBase {
  readonly type: "element";
  /** Tag name (e.g. "worksheet", "row"). Includes prefix if present. */
  name: string;
  /** Attribute map. */
  attributes: Record<string, string>;
  /** Ordered child nodes. */
  children: XmlNode[];
  /** Namespace prefix (e.g. "x" for `<x:row>`). Empty string if unprefixed. */
  prefix?: string;
  /** Local name without prefix (e.g. "row" for `<x:row>`). */
  local?: string;
  /** Namespace URI resolved from `xmlns` declarations. */
  uri?: string;
  /** Namespace declarations on this element (`xmlns:prefix="uri"`). */
  ns?: Record<string, string>;
}

/** XML text node. */
export interface XmlText extends XmlNodeBase {
  readonly type: "text";
  /** Text content (already decoded). */
  value: string;
}

/** XML CDATA section node. */
export interface XmlCData extends XmlNodeBase {
  readonly type: "cdata";
  /** CDATA content (raw, not entity-encoded). */
  value: string;
}

/** XML comment node. */
export interface XmlComment extends XmlNodeBase {
  readonly type: "comment";
  /** Comment content. */
  value: string;
}

/** XML processing instruction node. */
export interface XmlProcessingInstruction extends XmlNodeBase {
  readonly type: "processing-instruction";
  /** Target (e.g. "xml-stylesheet"). */
  target: string;
  /** Body content. */
  body: string;
}

/** Any XML node. */
export type XmlNode = XmlElement | XmlText | XmlCData | XmlComment | XmlProcessingInstruction;

/** Complete XML document with optional declaration and root element. */
export interface XmlDocument {
  /** XML declaration attributes (version, encoding, standalone). */
  declaration?: Record<string, string>;
  /** Root element (first root element in fragment mode). */
  root: XmlElement;
  /**
   * All root-level elements. In non-fragment mode this always has exactly one element.
   * In fragment mode (`{ fragment: true }`), this contains all top-level elements.
   */
  roots: XmlElement[];
  /**
   * Top-level nodes that appear before the root element (comments and processing
   * instructions). Only populated when `comments` or `processingInstructions`
   * options are enabled.
   */
  prologue: Array<XmlComment | XmlProcessingInstruction>;
}

// =============================================================================
// XML Writer Sink Interface
// =============================================================================

/**
 * Output target for XML writing.
 *
 * Both `XmlWriter` (buffered) and `XmlStreamWriter` (streaming) implement
 * this interface so that rendering code can be written against the interface
 * and work with either backend.
 */
export interface XmlSink {
  /** Write the XML declaration. */
  openXml(attributes?: XmlAttributes): void;

  /** Open an element with optional attributes. */
  openNode(name: string, attributes?: XmlAttributes): void;

  /** Add a single attribute to the currently-open element. */
  addAttribute(name: string, value: string | number | boolean): void;

  /** Add multiple attributes to the currently-open element. */
  addAttributes(attributes: XmlAttributes): void;

  /** Write escaped text content inside the current element. */
  writeText(text: string | number): void;

  /** Write raw (pre-escaped) XML inside the current element. */
  writeRaw(xml: string): void;

  /** Write a CDATA section inside the current element. */
  writeCData(text: string): void;

  /** Write an XML comment. */
  writeComment(text: string): void;

  /** Close the most recently opened element. */
  closeNode(): void;

  /**
   * Convenience: write a complete leaf element.
   * Equivalent to openNode + optional writeText + closeNode.
   */
  leafNode(name: string, attributes?: XmlAttributes, text?: string | number): void;
}

// =============================================================================
// SAX Event Types
// =============================================================================

/** SAX open-tag event data. */
export interface SaxTag {
  /** Element name (includes prefix if present, e.g. "x:row"). */
  name: string;
  /** Attribute map. */
  attributes: Record<string, string>;
  /** Whether this is a self-closing tag (e.g. `<br/>`). */
  isSelfClosing: boolean;
  /** Namespace prefix (e.g. "x"). Empty string if unprefixed. Only set when `xmlns` option is true. */
  prefix?: string;
  /** Local name without prefix (e.g. "row"). Only set when `xmlns` option is true. */
  local?: string;
  /** Namespace URI. Only set when `xmlns` option is true. */
  uri?: string;
  /** Namespace declarations on this element. Only set when `xmlns` option is true. */
  ns?: Record<string, string>;
}

/**
 * The subset of a {@link SaxTag} a parser consumer needs at open-tag time:
 * the element name, its attributes, and (optionally) whether it self-closes.
 * A full `SaxTag` is assignable to this, while test fixtures can supply just
 * `{ name, attributes }`. Use this for `parseOpen`-style handler parameters.
 */
export type ParseOpenTag = Pick<SaxTag, "name" | "attributes"> & {
  isSelfClosing?: boolean;
};

/** SAX event discriminated union. */
export type SaxEvent =
  | { eventType: "opentag"; value: SaxTag }
  | { eventType: "text"; value: string }
  | { eventType: "closetag"; value: SaxTag }
  | { eventType: "cdata"; value: string }
  | { eventType: "comment"; value: string }
  | { eventType: "pi"; value: { target: string; body: string } }
  | { eventType: "error"; value: Error };

/**
 * Loose-typed SAX event for legacy consumers that don't narrow on eventType.
 * New code should use {@link SaxEvent} with proper discriminant checks.
 *
 * `value` is the union of every event payload rather than `any`, so a consumer
 * that ignores `eventType` still gets a checked type (and `any` never leaks out
 * of the parser into caller code).
 */
export interface SaxEventAny {
  eventType: "opentag" | "text" | "closetag" | "cdata" | "comment" | "pi" | "error";
  value: SaxTag | string | { target: string; body: string } | Error;
}

/** SAX event handler map. */
export interface SaxHandlers {
  opentag?: (tag: SaxTag) => void;
  text?: (text: string) => void;
  closetag?: (tag: SaxTag) => void;
  cdata?: (text: string) => void;
  comment?: (text: string) => void;
  pi?: (target: string, body: string) => void;
  error?: (err: Error) => void;
}

/**
 * Strategy for handling invalid XML characters (control chars, lone surrogates,
 * non-characters like U+FFFE/U+FFFF).
 *
 * - `"error"` — Report via error handler or throw (XML 1.0 strict). **Default.**
 * - `"skip"` — Silently remove the invalid character from the output.
 * - `"replace"` — Replace the invalid character with U+FFFD (REPLACEMENT CHARACTER).
 */
export type InvalidCharHandling = "error" | "skip" | "replace";

/** SAX parser options. */
export interface SaxOptions {
  /** Track position (line/column) for error messages. Default: true */
  position?: boolean;
  /** File name for error messages. */
  fileName?: string;
  /** Parse as fragment (no root element required). Default: false */
  fragment?: boolean;
  /** Enable namespace processing. Default: false */
  xmlns?: boolean;
  /** Maximum element nesting depth. Default: 256. Set 0 to disable. */
  maxDepth?: number;
  /**
   * Maximum total entity expansions (named entities only).
   * Protects against Billion Laughs and quadratic blowup attacks.
   * Default: 10000. Set 0 to disable.
   */
  maxEntityExpansions?: number;
  /**
   * How to handle invalid XML characters (ASCII control chars, lone surrogates,
   * non-characters U+FFFE/U+FFFF, DEL U+007F, etc.).
   *
   * - `"error"` — Report via error handler or throw. **(Default)**
   * - `"skip"` — Silently discard the character.
   * - `"replace"` — Replace with U+FFFD (REPLACEMENT CHARACTER).
   *
   * @default "error"
   */
  invalidCharHandling?: InvalidCharHandling;
}

// =============================================================================
// Stream Write Target
// =============================================================================

/**
 * Minimal writable interface for XmlStreamWriter.
 *
 * Accepts strings or Uint8Array chunks.
 * Compatible with Node.js Writable, browser WritableStream wrappers,
 * and the project's own StreamBuf.
 *
 * **Backpressure caveat**: `write()` returns `void`. The XmlStreamWriter
 * cannot signal or wait for backpressure on the target — it pushes
 * chunks synchronously as the caller invokes `openNode`/`writeText`/etc.
 *
 * If the target is a slow sink (HTTP response, fs stream, etc), the
 * caller is responsible for ensuring chunks are consumed at a rate
 * matching production. Two safe patterns:
 *
 *   1. Wrap the target so `write()` buffers into a bounded queue and
 *      pause your XML production loop when the queue is full.
 *   2. Use a small in-memory buffer as the `WritableTarget`, then ship
 *      the assembled bytes to the slow sink with proper backpressure
 *      (e.g. via `pipeline()`).
 *
 * `XmlStreamWriter` is used internally by documonster's xlsx writer. The xlsx
 * writer wraps it in a backpressure-aware zip pipeline that awaits drain
 * BETWEEN zip entries — so memory grows at most by one entry's worth of
 * uncompressed XML before the producer is parked. Within a single very
 * large worksheet entry the synchronous push pattern still applies; that
 * is the practical bound on how much a slow user sink can buffer behind
 * the xlsx writer.
 */
export interface WritableTarget {
  write(chunk: string | Uint8Array): void;
}

// =============================================================================
// DOM Conversion Options
// =============================================================================

/**
 * Options for `toPlainObject()`.
 *
 * Controls how an {@link XmlElement} DOM tree is converted to a plain
 * JavaScript object.
 */
export interface ToPlainObjectOptions {
  /**
   * When true, discard all attributes entirely. Takes precedence over
   * `attributePrefix` — when `ignoreAttributes` is true, no attribute keys
   * appear in the output regardless of prefix settings.
   * @default false
   */
  ignoreAttributes?: boolean;
  /**
   * Prefix for attribute keys.
   * Set to `""` to use bare attribute names.
   * Ignored when `ignoreAttributes` is true.
   * @default "@_"
   */
  attributePrefix?: string;
  /**
   * Key used for text content when an element has both text and child elements
   * (mixed content).
   * @default "#text"
   */
  textKey?: string;
  /**
   * When true, always wrap child elements in arrays even if there is only one.
   * When false (default), single children are stored as a plain value.
   * @default false
   */
  alwaysArray?: boolean;
  /**
   * Callback to determine whether a specific tag name should always be wrapped
   * in an array, even if only a single element exists. Takes precedence over
   * `alwaysArray` for matching names. When both `isArray` and `alwaysArray` are
   * set, a tag is wrapped if either returns/is true.
   */
  isArray?: (name: string) => boolean;
  /**
   * When true, include CDATA node values (already merged to text by default
   * in `parseXml`, only relevant with `cdataAsNodes`).
   * @default true
   */
  preserveCData?: boolean;
  /**
   * When true, discard whitespace-only text in elements that also have child
   * elements (typical of pretty-printed XML indentation). Leaf elements that
   * contain only whitespace text are **not** affected — their text is preserved.
   * @default true
   */
  ignoreWhitespaceText?: boolean;
}

/**
 * Options for `parseXmlToObject()`.
 *
 * Extends {@link ToPlainObjectOptions} with SAX parser settings, since
 * `parseXmlToObject` handles both parsing and conversion in a single pass.
 */
export interface ParseXmlToObjectOptions extends ToPlainObjectOptions {
  /** Parse as fragment (no root element required). Default: false */
  fragment?: boolean;
  /** Maximum element nesting depth. Default: 256. */
  maxDepth?: number;
  /** Maximum total entity expansions. Default: 10000. */
  maxEntityExpansions?: number;
  /** How to handle invalid XML characters. Default: "error". */
  invalidCharHandling?: InvalidCharHandling;
}

// =============================================================================
// DOM Parse Options
// =============================================================================

/** Options for `parseXml()`. */
export interface XmlParseOptions {
  /** Include comments in the tree. Default: false */
  comments?: boolean;
  /** Include processing instructions in the tree. Default: false */
  processingInstructions?: boolean;
  /** Include CDATA as explicit nodes (vs merged into text). Default: false */
  cdataAsNodes?: boolean;
  /** Parse as fragment (no root element required). Default: false */
  fragment?: boolean;
  /** Enable namespace processing on elements. Default: false */
  xmlns?: boolean;
  /** Maximum element nesting depth. Default: 256. */
  maxDepth?: number;
  /** Maximum total entity expansions. Default: 10000. */
  maxEntityExpansions?: number;
  /** How to handle invalid XML characters. Default: "error". */
  invalidCharHandling?: InvalidCharHandling;
}
