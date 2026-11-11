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
  /**
   * Top-level comments and processing instructions after the root element (in fragment mode,
   * after the first root element). Like `prologue`, only populated when `comments` or
   * `processingInstructions` are enabled; always present on a parsed document.
   */
  epilogue?: Array<XmlComment | XmlProcessingInstruction>;
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
  writeText(text: string | number | boolean): void;

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
  leafNode(name: string, attributes?: XmlAttributes, text?: string | number | boolean): void;
}

// =============================================================================
// SAX Event Types
// =============================================================================

/** SAX open-tag event data. */
export interface XmlSaxTag {
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
 * The subset of a {@link XmlSaxTag} a parser consumer needs at open-tag time:
 * the element name, its attributes, and (optionally) whether it self-closes.
 * A full `XmlSaxTag` is assignable to this, while test fixtures can supply just
 * `{ name, attributes }`. Use this for `parseOpen`-style handler parameters.
 */
export type ParseOpenTag = Pick<XmlSaxTag, "name" | "attributes"> & {
  isSelfClosing?: boolean;
};

/** SAX event discriminated union. */
export type XmlSaxEvent =
  | { eventType: "opentag"; value: XmlSaxTag }
  | { eventType: "text"; value: string }
  | { eventType: "closetag"; value: XmlSaxTag }
  | { eventType: "cdata"; value: string }
  | { eventType: "comment"; value: string }
  | { eventType: "pi"; value: { target: string; body: string } }
  | { eventType: "error"; value: Error };

/**
 * Loose-typed SAX event for legacy consumers that don't narrow on eventType.
 * New code should use {@link XmlSaxEvent} with proper discriminant checks.
 *
 * `value` is the union of every event payload rather than `any`, so a consumer
 * that ignores `eventType` still gets a checked type (and `any` never leaks out
 * of the parser into caller code).
 */
export interface XmlSaxEventAny {
  eventType: "opentag" | "text" | "closetag" | "cdata" | "comment" | "pi" | "error";
  value: XmlSaxTag | string | { target: string; body: string } | Error;
}

/** SAX event handler map. */
export interface XmlSaxHandlers {
  opentag?: (tag: XmlSaxTag) => void;
  text?: (text: string) => void;
  closetag?: (tag: XmlSaxTag) => void;
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
export type XmlInvalidCharHandling = "error" | "skip" | "replace";

/** SAX parser options. */
export interface XmlSaxOptions {
  /** Track position (line/column) for error messages. Default: true */
  position?: boolean;
  /** File name for error messages. */
  fileName?: string;
  /**
   * Parse as a fragment: any number of root elements, none included, with text between them reported
   * as text. Default: false
   */
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
   * Maximum total length of the text internal entities expand to, in UTF-16 code units — measured
   * before it is built. A few entities used many times stay within `maxEntityExpansions` and can
   * still make a string larger than memory. Default: 10,000,000. Set 0 to disable.
   */
  maxEntityExpansionLength?: number;
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
  invalidCharHandling?: XmlInvalidCharHandling;
  /**
   * Skip whatever comes before the document's first `<` — text no XML document holds there, such as
   * the two lines `unzip -c` writes before a file's content (`Archive:  …` / `  inflating: …`) — as
   * not part of the document: it is neither reported as text nor refused, and the XML declaration may
   * follow it. Positions stay positions in what was given; how much was skipped is
   * {@link SaxParser.skippedLeadingText}. The skipped text must still decode, and what follows must be
   * a whole document: only its start moves. No effect on a parser that resumes, or reads a fragment.
   * @default false
   */
  skipLeadingText?: boolean;
  /**
   * Offset in the file of the first byte the parser is given, so the byte positions it reports
   * (`XmlParseError.byte`, `startByte`, `endByte`) are positions in the file. Only bytes
   * have positions: it has no effect on a parser given text. A byte order mark is stripped only at
   * offset 0.
   * @default 0
   */
  byteOffset?: number;
  /**
   * Start inside the root element, at an unknown depth: the input begins in the content of some
   * element. An end tag of an element opened before the start is passed over, a root element is not
   * required and not counted, and a byte input that begins inside a character drops the rest of it.
   * Elements opened since the start and still open at the end are reported; with `context`, so is
   * an end that comes after a tag but before the root element's end tag. Pair it with `byteOffset`
   * and `context`. `maxDepth` counts from the start.
   * @default false
   */
  resume?: boolean;
  /**
   * Read the general entities the DOCTYPE internal subset declares (`<!ENTITY name "value">`) and
   * expand them where they are used, as XML 1.0 §4.4 describes; `getPrologContext()`
   * then carries them. Replacement text that holds markup (`<`) is refused, not parsed. External
   * entities, an external DTD subset and parameter entities are never read: a reference to an
   * entity declared outside the internal subset, or after a parameter-entity reference, is an
   * undefined entity. Off by default, so such a reference is reported as it always was.
   * @default false
   */
  internalEntities?: boolean;
  /**
   * The largest prolog context `internalEntities` may gather, in bytes of its JSON form, and the
   * longest markup declaration of the DOCTYPE internal subset the parser holds to check — and, in
   * characters with a small share per entry, what it keeps of that subset's declarations (entity
   * bindings, attribute declarations) to read the rest by, with or without `internalEntities`. Beyond
   * either the parse fails with an `XmlLimitExceededError` (`limit: "prologContext"`), rather than
   * keep part of a context or take a declaration unchecked. `Xml.parse` and `Xml.parseToObject`,
   * whose input is all in memory, set no such bound.
   * @default 65536
   */
  maxContextBytes?: number;
  /**
   * The prolog context of the file a `resume` parser starts inside — what another parser's
   * `getPrologContext()` returned — so the entities declared there expand here as
   * they do for a parser that read the file from its start.
   */
  context?: XmlPrologContext;
}

// =============================================================================
// Prolog Context, Records and Byte Sources
// =============================================================================

/**
 * What a reader that starts inside a file's root element needs from the prolog it did not read:
 * the root element's name and the general entities the DOCTYPE internal subset declares. Plain
 * data, so it can be sent with `JSON.stringify` to wherever that reader runs.
 */
export interface XmlPrologContext {
  /** Name of the root element. */
  readonly root: string;
  /** Replacement text of each internal general entity, by name, with character references expanded. */
  readonly entities: Readonly<Record<string, string>>;
  /**
   * `"ascii"` when the XML declaration names ASCII: a byte past `0x7F` later in the file is then no
   * character, as it is for a reader that read the declaration. Absent for UTF-8.
   */
  readonly encoding?: "ascii";
  /**
   * The attributes the internal subset's `<!ATTLIST>` declarations give a default or a type other
   * than CDATA, by element name and then attribute name. Absent when it declares none.
   */
  readonly attributes?: Readonly<Record<string, Readonly<Record<string, XmlAttributeDeclaration>>>>;
}

/**
 * What a parser does with an attribute an `<!ATTLIST>` in the internal subset declares, as XML 1.0
 * §5.1 asks of one that reads the declarations: give an element that does not specify it its
 * `default`, and — for a type other than CDATA (`tokenized`) — read a value with its spaces
 * collapsed and trimmed (§3.3.3).
 */
export interface XmlAttributeDeclaration {
  /** Its type is other than CDATA: a value read has runs of spaces collapsed, and none at either end. */
  readonly tokenized?: true;
  /** Its default value, normalized as a value read is; absent for `#REQUIRED` and `#IMPLIED`. */
  readonly default?: string;
}

/** The bytes of a file, whole or in chunks; a chunk may end anywhere, inside a character too. */
export type XmlByteSource = Uint8Array | Iterable<Uint8Array> | AsyncIterable<Uint8Array>;

/** A record `Xml.readRecords` read: one element, as `Xml.parseToObject` reads its text. */
export interface XmlRecord {
  /** The element's name. */
  readonly name: string;
  /**
   * The element as an object, deep-equal to `Xml.parseToObject` of its text with the same
   * conversion options: `{ [name]: value }`.
   */
  readonly object: Record<string, unknown>;
  /** Offset in the file of the `<` of its start tag. */
  readonly startByte: number;
  /** Offset in the file just past the `>` of its end tag (of its start tag, when it self-closes). */
  readonly endByte: number;
  /** Its text as the file holds it, with `rawXml`. */
  readonly xml?: string;
}

/** Options for `Xml.readRecords`. */
export interface XmlReadRecordsOptions extends XmlToPlainObjectOptions {
  /**
   * Names of the elements to read as records. A record is an element with one of these names
   * that is not inside another record: one inside it, of any of the names, is part of it.
   */
  tags: readonly string[];
  /**
   * Called with each record, in file order. A returned promise is awaited before the next record
   * is read, so a slow consumer slows the read rather than letting records pile up. Records arrive
   * as they are read: a call that fails may already have given some, so stage them until the call —
   * and, for a partition, every range of it — has succeeded.
   */
  onRecord: (record: XmlRecord) => void | Promise<void>;
  /** Offset in the file of the first byte of the source, and the start of the chunk to read. */
  offset?: number;
  /**
   * Skip what comes before the document's first `<` — such as the lines `unzip -c` writes before a
   * file's content — as `Xml.SaxParser`'s `skipLeadingText` does; the result's `skippedLeadingBytes`
   * says how much. Applies to a read from offset 0, where the document starts; byte positions stay
   * the file's. Give it to every range of a partition and to `Xml.readRoot` alike.
   * @default false
   */
  skipLeadingText?: boolean;
  /**
   * Length of the chunk to read, in bytes: the records whose start tag begins in
   * `[offset, offset + size)` are read. Omitted, every record from `offset` on.
   */
  size?: number;
  /**
   * The prolog context of the file — `Xml.readRoot(...).context` — needed when `offset` is
   * not 0, where the prolog is not read.
   */
  context?: XmlPrologContext;
  /** Give each record its text as the file holds it (`xml`). Off by default: it costs a copy. */
  rawXml?: boolean;
  /** Maximum element nesting depth; from the chunk start when `offset` is not 0. Default: 256. */
  maxDepth?: number;
  /** Maximum total entity expansions, counted per call. Default: 10000. */
  maxEntityExpansions?: number;
  /** Maximum total length of entity expansions, in UTF-16 code units, counted per call. Default: 10,000,000. */
  maxEntityExpansionLength?: number;
  /** How to handle invalid XML characters, and bytes that are not UTF-8. Default: "error". */
  invalidCharHandling?: XmlInvalidCharHandling;
  /** The largest prolog context read at offset 0, in bytes of its JSON form. Default: 65536. */
  maxContextBytes?: number;
  /** Cancels the read: the source is closed and the call rejects with an `AbortError` whose `cause` is the signal's reason. */
  signal?: AbortSignal;
}

/** What `Xml.readRecords` read. */
export interface XmlReadRecordsResult {
  /** How many records were given to `onRecord`. */
  readonly records: number;
  /** How many bytes `skipLeadingText` passed over before the document; absent for none. */
  readonly skippedLeadingBytes?: number;
  /**
   * Offset in the file of the first byte not parsed: the read went past the chunk's end until its
   * last record closed and the next chunk's start was placed.
   */
  readonly endByte: number;
  /**
   * The structure of the chunk — the elements it closes that opened before it, the ones it leaves
   * open, and what it holds outside any element — for `Xml.verifyPartition`, which joins the chunks of
   * a partition and checks the file is whole. Plain data: it can be sent back from a worker as JSON.
   */
  readonly structure: XmlPartitionStructure;
}

/**
 * One chunk's share of a file's structure, as `Xml.readRecords` reports it for `Xml.verifyPartition`.
 * Its fields are for that function to read; they are documented so the object can be stored or sent,
 * not to be built by hand.
 */
export interface XmlPartitionStructure {
  /** The chunk: `[start, end)`; `end` is `null` for one that runs to the end of the file. */
  readonly start: number;
  readonly end: number | null;
  /** Whether the read reached the end of the file and found it well-formed there. */
  readonly final: boolean;
  /** Where the file ends — its length — when the read reached it (`final`). */
  readonly eof?: number;
  /**
   * In file order: `["/", byte, name]` — an end tag closing an element opened before the chunk;
   * `["e", byte]` — an element begun and ended in the chunk outside any element the chunk opened;
   * `["t", byte]` — text there that is not white space; `["l", byte]` — such text read from the
   * very start of the chunk, which may be the end of markup begun before it. Between two `/` steps
   * no more is kept than the join needs: the `l`, the first other entry, the first element, and the
   * entry after it.
   */
  readonly steps: readonly XmlPartitionStep[];
  /** The elements the chunk opened and left open at its end, outermost first. */
  readonly open: readonly string[];
  /**
   * What the read found past the chunk's end, before the first markup there: the byte where text
   * that is not white space begins, or `null` for none. Absent when the read did not go that far.
   * It tells `Xml.verifyPartition` what the next chunk's `l` step is.
   */
  readonly follow?: number | null;
  /**
   * Present, `true`, when the read — from a start inside the file — met no markup in the whole chunk:
   * it took all of it for text, so its `follow` is no surer, and the chunk before speaks for both.
   */
  readonly textOnly?: true;
}

/** One entry of {@link XmlPartitionStructure.steps}. */
export type XmlPartitionStep =
  | readonly [kind: "/", byte: number, name: string]
  | readonly [kind: "e" | "t" | "l", byte: number];

/** Options for `Xml.readRoot`. */
export interface XmlReadRootOptions {
  /** The root element's `<` must begin before this offset. Default: 1,000,000. */
  maxStartByte?: number;
  /** The longest the root element's start tag may be, in bytes. Default: 65,536. */
  maxTagBytes?: number;
  /** The largest prolog context, in bytes of its JSON form. Default: 65,536. */
  maxContextBytes?: number;
  /**
   * Maximum total entity expansions in the root's attributes, as `Xml.readRecords` takes it, so a
   * file it can read with a raised limit can give its context too. Default: 10000.
   */
  maxEntityExpansions?: number;
  /** Maximum total length of entity expansions, in UTF-16 code units. Default: 10,000,000. */
  maxEntityExpansionLength?: number;
  /** Cancels the read: the source is closed and the call rejects with an `AbortError` whose `cause` is the signal's reason. */
  signal?: AbortSignal;
  /**
   * Skip what comes before the document's first `<`, as `Xml.readRecords`' `skipLeadingText` does.
   * The skipped bytes count toward `maxStartByte`.
   * @default false
   */
  skipLeadingText?: boolean;
}

/** The root element `Xml.readRoot` read, and the prolog before it. */
export interface XmlRootElement {
  /** The root element's name. */
  readonly name: string;
  /**
   * Its attributes, references expanded. Namespace declarations (`xmlns`, `xmlns:p`) are among
   * them.
   */
  readonly attributes: Readonly<Record<string, string>>;
  /** What a reader that starts later in the file needs: `Xml.readRecords`' `context`. */
  readonly context: XmlPrologContext;
  /** How many bytes `skipLeadingText` passed over before the document; absent for none. */
  readonly skippedLeadingBytes?: number;
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
 *   2. Use a small in-memory buffer as the `XmlWritableTarget`, then ship
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
export interface XmlWritableTarget {
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
export interface XmlToPlainObjectOptions {
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
 * Extends {@link XmlToPlainObjectOptions} with SAX parser settings, since
 * `parseXmlToObject` handles both parsing and conversion in a single pass.
 */
export interface XmlParseToObjectOptions extends XmlToPlainObjectOptions {
  /**
   * Parse as a fragment: any number of root elements, each becoming a key of the result, none
   * included (the result is `{}`). Text outside every element has no key to go under and is not in
   * the result. Default: false
   */
  fragment?: boolean;
  /** Maximum element nesting depth. Default: 256. */
  maxDepth?: number;
  /** Maximum total entity expansions. Default: 10000. */
  maxEntityExpansions?: number;
  /** Maximum total length of entity expansions, in UTF-16 code units. Default: 10,000,000. */
  maxEntityExpansionLength?: number;
  /** How to handle invalid XML characters. Default: "error". */
  invalidCharHandling?: XmlInvalidCharHandling;
  /**
   * Read the general entities the DOCTYPE internal subset declares and expand them where they are
   * used — see `XmlSaxOptions.internalEntities`. External entities are never read.
   * @default false
   */
  internalEntities?: boolean;
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
  /**
   * Accept several root elements: `root` is the first and `roots` holds them all. At least one
   * element is still required — input with none is refused, as it is without this option. Default: false
   */
  fragment?: boolean;
  /** Enable namespace processing on elements. Default: false */
  xmlns?: boolean;
  /** Maximum element nesting depth. Default: 256. */
  maxDepth?: number;
  /** Maximum total entity expansions. Default: 10000. */
  maxEntityExpansions?: number;
  /** Maximum total length of entity expansions, in UTF-16 code units. Default: 10,000,000. */
  maxEntityExpansionLength?: number;
  /** How to handle invalid XML characters. Default: "error". */
  invalidCharHandling?: XmlInvalidCharHandling;
  /**
   * Read the general entities the DOCTYPE internal subset declares and expand them where they are
   * used — see `XmlSaxOptions.internalEntities`. External entities are never read.
   * @default false
   */
  internalEntities?: boolean;
}

/** @deprecated Use {@link XmlToPlainObjectOptions}. Removed in 1.0. */
export type ToPlainObjectOptions = XmlToPlainObjectOptions;
/** @deprecated Use {@link XmlParseToObjectOptions}. Removed in 1.0. */
export type ParseXmlToObjectOptions = XmlParseToObjectOptions;
/** @deprecated Use {@link XmlInvalidCharHandling}. Removed in 1.0. */
export type InvalidCharHandling = XmlInvalidCharHandling;
/** @deprecated Use {@link XmlWritableTarget}. Removed in 1.0. */
export type WritableTarget = XmlWritableTarget;
/** @deprecated Use {@link XmlSaxTag}. Removed in 1.0. */
export type SaxTag = XmlSaxTag;
/** @deprecated Use {@link XmlSaxEvent}. Removed in 1.0. */
export type SaxEvent = XmlSaxEvent;
/** @deprecated Use {@link XmlSaxEventAny}. Removed in 1.0. */
export type SaxEventAny = XmlSaxEventAny;
/** @deprecated Use {@link XmlSaxHandlers}. Removed in 1.0. */
export type SaxHandlers = XmlSaxHandlers;
/** @deprecated Use {@link XmlSaxOptions}. Removed in 1.0. */
export type SaxOptions = XmlSaxOptions;
