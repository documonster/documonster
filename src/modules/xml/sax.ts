/**
 * SAX XML Parser
 *
 * High-performance streaming SAX parser for XML.
 * Zero external dependencies. Optimized for common XML patterns.
 *
 * Migrated from the excel module's parse-sax.ts into the standalone XML module.
 * Enhancements over the original:
 * - Uses @xml/types for event types (XmlSaxTag, XmlSaxEvent, XmlSaxHandlers)
 * - Uses @xml/errors for parse errors (XmlParseError)
 * - Exposes CDATA, comment, and processing-instruction event handlers
 * - Provides both callback API and async-generator API
 *
 * Based on XML 1.0 specification with fast-path optimizations for ASCII.
 */

import { isUint8Array } from "@utils/binary";
import { toError } from "@utils/errors";
import { check, s } from "@utils/schema";
import { checkSaxOptions, snapshotSaxOptions, xmlArgument } from "@xml/arguments";
import {
  charRefValue,
  isNameChar,
  isNameStartChar,
  isS,
  isXmlName,
  NAME_CHAR_ASCII,
  NAME_START_CHAR_ASCII
} from "@xml/chars";
import { ContextSize } from "@xml/context";
import {
  checkDeclarationStart,
  checkDoctypeHead,
  checkXmlDeclaration,
  DeclaredEntities,
  EntityProblem,
  isPredefinedEntity,
  readMarkupDeclaration,
  summarizeEntity
} from "@xml/dtd";
import type { AttlistDeclaration, EntitySummary } from "@xml/dtd";
import {
  XmlArgumentError,
  XmlError,
  XmlLimitExceededError,
  XmlParseError,
  XmlTruncatedError
} from "@xml/errors";
import type { XmlParseLimit } from "@xml/errors";
import type {
  XmlAttributeDeclaration,
  XmlInvalidCharHandling,
  XmlPrologContext,
  XmlSaxEvent,
  XmlSaxHandlers,
  XmlSaxOptions,
  XmlSaxTag
} from "@xml/types";
import { escapedByte, Utf8ChunkDecoder, utf8Length } from "@xml/utf8-decoder";

/** Whether `text` is all XML white space. */
function isWhitespaceText(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    if (!isS(text.charCodeAt(i))) {
      return false;
    }
  }
  return true;
}

/** A value of a type other than CDATA: runs of spaces as one, none at either end (§3.3.3). */
function collapseSpaces(value: string): string {
  return value.includes("  ") || value.charCodeAt(0) === SPACE || value.endsWith(" ")
    ? value
        .split(" ")
        .filter(token => token !== "")
        .join(" ")
    : value;
}

/** What the first declaration of an entity bound: see {@link SaxParser.entityBindings}. */
type EntityBinding =
  | { readonly kind: "internal"; readonly summary: EntitySummary }
  | { readonly kind: "external" | "parameter" };

// =============================================================================
// Character Codes
// =============================================================================

const TAB = 9;
const NL = 0xa;
const CR = 0xd;
const SPACE = 0x20;
const BANG = 0x21; // !
const PERCENT = 0x25; // %
const DQUOTE = 0x22; // "
const HASH = 0x23; // #
const AMP = 0x26; // &
const SQUOTE = 0x27; // '
const MINUS = 0x2d; // -
const FORWARD_SLASH = 0x2f; // /
const SEMICOLON = 0x3b; // ;
const LESS = 0x3c; // <
const EQUAL = 0x3d; // =
const GREATER = 0x3e; // >
const QUESTION = 0x3f; // ?
const OPEN_BRACKET = 0x5b; // [
const CLOSE_BRACKET = 0x5d; // ]
const REPLACEMENT_CHAR = 0xfffd; // U+FFFD REPLACEMENT CHARACTER
const REPLACEMENT_STR = "\uFFFD"; // Pre-allocated string form of U+FFFD

// =============================================================================
// Pre-computed Lookup Tables
// =============================================================================

const ASCII_CHARS: string[] = /* @__PURE__ */ (() => {
  const t = new Array<string>(128);
  for (let i = 0; i < 128; i++) {
    t[i] = String.fromCharCode(i);
  }
  return t;
})();

function charFromCode(c: number): string {
  return c < 128 ? ASCII_CHARS[c] : String.fromCodePoint(c);
}

function isQuote(c: number): boolean {
  return c === DQUOTE || c === SQUOTE;
}

// =============================================================================
// Built-in XML Entities
// =============================================================================

const SAX_EVENTS: ReadonlySet<string> = new Set([
  "opentag",
  "text",
  "closetag",
  "cdata",
  "comment",
  "pi",
  "error"
]);

function checkEventName(method: string, name: unknown): void {
  if (typeof name !== "string" || !SAX_EVENTS.has(name)) {
    throw new XmlArgumentError(
      "name",
      `Xml.SaxParser#${method}: unknown event ${JSON.stringify(name)}; expected one of ${[...SAX_EVENTS].join(", ")}`
    );
  }
}

/** An XML Name with no colon: a namespace prefix or local part (Namespaces in XML 1.0 §3). */
function isNCName(name: string): boolean {
  return name !== "" && !name.includes(":") && isXmlName(name);
}

/** Whether `name` is a QName: an NCName, or two joined by one colon. */
function isQName(name: string): boolean {
  const colon = name.indexOf(":");
  return colon < 0
    ? isNCName(name)
    : isNCName(name.slice(0, colon)) && isNCName(name.slice(colon + 1));
}

/** The share of `keepDeclared`'s count one kept declaration takes beyond its characters. */
const KEPT_DECLARATION_OVERHEAD = 8;

const XML_ENTITIES: Record<string, string> = {
  amp: "&",
  gt: ">",
  lt: "<",
  quot: '"',
  apos: "'"
};

// =============================================================================
// Parser States
// =============================================================================

const S_TEXT = 0;
const S_OPEN_WAKA = 1;
const S_OPEN_WAKA_BANG = 2;
const S_OPEN_TAG = 3;
const S_OPEN_TAG_SLASH = 4;
const S_ATTRIB = 5;
const S_ATTRIB_NAME = 6;
const S_ATTRIB_NAME_SAW_WHITE = 7;
const S_ATTRIB_VALUE = 8;
const S_ATTRIB_VALUE_QUOTED = 9;
const S_ATTRIB_VALUE_CLOSED = 10;
const S_CLOSE_TAG = 11;
const S_CLOSE_TAG_SAW_WHITE = 12;
const S_COMMENT = 13;
const S_COMMENT_ENDING = 14;
const S_COMMENT_ENDED = 15;
const S_CDATA = 16;
const S_CDATA_ENDING = 17;
const S_CDATA_ENDING_2 = 18;
const S_PI = 19;
const S_PI_ENDING = 20;
const S_DOCTYPE = 21;
const S_DOCTYPE_QUOTE = 22;
const S_DOCTYPE_DTD = 23;
const S_DOCTYPE_DTD_QUOTED = 24;
const S_ENTITY = 25;
const S_DOCTYPE_DTD_LESS = 26;
const S_DOCTYPE_DTD_BANG = 27;
const S_DOCTYPE_DTD_BANG_DASH = 28;
const S_DOCTYPE_DTD_DECL = 29;
const S_DOCTYPE_DTD_COMMENT = 30;
const S_DOCTYPE_DTD_COMMENT_ENDING = 31;
const S_DOCTYPE_DTD_COMMENT_ENDED = 32;
const S_DOCTYPE_DTD_PI = 33;
const S_DOCTYPE_DTD_PI_ENDING = 34;
const S_DOCTYPE_DTD_PE_REF = 35;
const S_DOCTYPE_AFTER_SUBSET = 36;

/** What the document was in the middle of, by state, for an input that ends there. */
function unfinishedConstruct(state: number): string {
  switch (state) {
    case S_OPEN_WAKA:
    case S_OPEN_WAKA_BANG:
      return "markup";
    case S_OPEN_TAG:
    case S_OPEN_TAG_SLASH:
    case S_ATTRIB:
    case S_ATTRIB_NAME:
    case S_ATTRIB_NAME_SAW_WHITE:
    case S_ATTRIB_VALUE:
    case S_ATTRIB_VALUE_QUOTED:
    case S_ATTRIB_VALUE_CLOSED:
      return "a start tag";
    case S_CLOSE_TAG:
    case S_CLOSE_TAG_SAW_WHITE:
      return "an end tag";
    case S_COMMENT:
    case S_COMMENT_ENDING:
    case S_COMMENT_ENDED:
      return "a comment";
    case S_CDATA:
    case S_CDATA_ENDING:
    case S_CDATA_ENDING_2:
      return "a CDATA section";
    case S_PI:
    case S_PI_ENDING:
      return "a processing instruction";
    case S_ENTITY:
      return "an entity reference";
    default:
      return "the DOCTYPE declaration";
  }
}

/** Set while {@link SaxParser.withCheckedOptions} builds a parser. */
let optionsChecked = false;

/** Where the parser is in the document, for {@link SaxParser.getLexicalState}. */
const PHASE_PROLOG = "p";
const PHASE_CONTENT = "c";

/** The longest DOCTYPE head (`<!DOCTYPE name SYSTEM "…"`) read before it is refused. */
const MAX_DOCTYPE_HEAD = 65536;

/** Encodings a byte-fed parser reads: UTF-8, and ASCII, which is a subset of it. */
const UTF8_ENCODINGS: ReadonlySet<string> = new Set(["utf-8", "utf8", "us-ascii", "ascii"]);
const ASCII_ENCODINGS: ReadonlySet<string> = new Set(["us-ascii", "ascii"]);

const EMPTY_BYTES = new Uint8Array(0);

/**
 * The attributes of every tag that has none, while attribute names are prefixed: one frozen object
 * rather than a new one per tag.
 *
 * @internal
 */
export const NO_ATTRIBUTES: Record<string, string> = Object.freeze(
  Object.create(null) as Record<string, string>
);

/** How many names the parser keeps for reuse; a power of two. */
const NAME_SLOTS = 512;

/** The slot of the name cache `text[start, end)` belongs in: a hash of its length and its ends. */
function nameSlot(text: string, start: number, end: number): number {
  return (
    (text.charCodeAt(start) * 31 + text.charCodeAt(end - 1) * 7 + (end - start)) & (NAME_SLOTS - 1)
  );
}

// =============================================================================
// SaxParser
// =============================================================================

/**
 * Streaming SAX XML parser.
 *
 * Feed XML text via {@link write}() in chunks; events fire synchronously
 * as elements are encountered. Call {@link close}() when done.
 *
 * @example
 * ```ts
 * const parser = new SaxParser();
 * parser.on("opentag", tag => console.log("open:", tag.name));
 * parser.on("text", text => console.log("text:", text));
 * parser.on("closetag", tag => console.log("close:", tag.name));
 * parser.write('<root><child>hello</child></root>');
 * parser.close();
 * ```
 */
class SaxParser {
  // Configuration
  private trackPosition: boolean;
  private fileName?: string;
  private fragment: boolean;
  /** `skipLeadingText`, for a parser that reads from the document's start. */
  private readonly skipLeading: boolean;
  /** Whether what is read now is still before the document: skipped while `skipLeading`. */
  private beforeDocument = false;
  /** How much was skipped before the document: see {@link skippedLeadingText}. */
  private skippedLeading = 0;
  private xmlns: boolean;
  private maxDepth: number;
  private maxEntityExpansions: number;
  private maxEntityExpansionLength: number;
  /** The length of the text declared entities have expanded to so far. */
  private entityExpansionLength = 0;
  private invalidCharHandling: XmlInvalidCharHandling;
  private readonly byteOffset: number;
  private readonly resume: boolean;
  private readonly internalEntities: boolean;
  private readonly maxContextBytes: number;
  private readonly context: XmlPrologContext | undefined;

  // Security counters
  private _entityExpansionCount: number = 0;

  // Namespace state
  private _nsStack: Array<Record<string, string>> = [];

  // Parser state
  private state: number = S_TEXT;
  private chunk: string = "";
  private i: number = 0;
  private prevI: number = 0;
  private text: string = "";
  private name: string = "";
  private q: number | null = null;
  private tags: XmlSaxTag[] = [];
  private tag: XmlSaxTag | null = null;
  /** Names repeated on the tag being opened, reported when it opens — `null` until the first one. */
  private duplicateAttributes: string[] | null = null;
  private entity: string = "";
  private entityReturnState: number = S_TEXT;
  private openWakaBang: string = "";
  private sawRoot: boolean = false;
  private sawDoctype: boolean = false;
  private closedRoot: boolean = false;
  private carriedFromPrevious?: string;
  /** Where the carried character starts, in the units {@link unitAt} counts. */
  private carriedUnit = 0;
  private _bomStripped: boolean = false;
  /** In `resume` mode: whether any tag has been read, and whether the last end tag read was the root's. */
  private sawTag = false;
  private rootEnded = false;
  /**
   * Whether the comment, CDATA section or processing instruction being read keeps its text: only
   * when a handler will be given it. Otherwise a long one outside anything the consumer reads would
   * be held whole for nothing.
   */
  private keepText = true;
  /** How many `]` the text read so far ends with, up to two: see {@link findCDataEnd}. */
  private textBrackets = 0;
  /** Where the next `]]>` of the chunk ends: see {@link findCDataEnd}. */
  private cdataEndAt = Infinity;
  /** In `resume` mode, whether no markup has been read yet: see {@link checkCDataEnds}. */
  private leadingResume = false;
  /** In a processing instruction: where its target ends in `text` (`-1` until it does), and whether its body is kept. */
  private piTargetEnd = -1;
  private piKeepBody = false;
  private _closed: boolean = false;
  private _stopped = false;
  /** Set while {@link write} runs, so a handler cannot write to the parser that is calling it. */
  private _writing = false;
  /** Set when the XML declaration names ASCII: see {@link checkDeclaredEncoding}. */
  private switchToAscii = false;
  /** Whether the XML declaration named ASCII: a reader that starts later decodes as this one does. */
  private asciiDeclared = false;
  /** Where the processing instruction being read in the internal subset began. */
  private subsetPIStart = 0;
  private _paused = false;
  /** How far into `chunk` the parse goes, and whether the chunk ends the document. */
  private limit = 0;
  private ending = false;
  private reportedTextBeforeRoot: boolean = false;
  private reportedTextAfterRoot: boolean = false;

  // Position tracking
  /** 1-based line of the parser's current position (when `position` is on). */
  line: number = 1;
  /** 0-based column of the parser's current position (when `position` is on). */
  column: number = 0;
  private positionAtNewLine: number = 0;
  /** Characters consumed before the chunk being read, or last read: `position` adds `i` to it. */
  private chunkPosition: number = 0;

  // Positions in the file — bytes when the parser is given bytes, characters when it is given text.
  /** The decoder of a byte-fed document; `undefined` while the parser is given text. */
  private decoder: Utf8ChunkDecoder | undefined;
  /** Whether this document is given bytes (`true`), text (`false`), or nothing yet. */
  private bytes: boolean | undefined;
  /** Where `chunk[0]` starts. */
  private chunkUnit = 0;
  /** Whether every UTF-16 unit of `chunk` is one unit of position, so a position is an addition. */
  private chunkOneUnit = true;
  /**
   * Bytes beyond one per UTF-16 unit taken by the characters of `chunk` read so far: the scanner
   * adds them as it passes a non-ASCII character, so a position at the point being read costs an
   * addition rather than a pass over the text.
   */
  private extra = 0;
  /** Where the markup being read, or last read, began: its `<`. */
  private markupStart = 0;
  /** Where the document's first character is — past a byte order mark: the one place an XML declaration may begin. */
  private documentStart = 0;
  /** Where the text being read began: just after the markup before it. */
  private textStart = 0;
  /** The range of the event being reported. */
  private eventStart = 0;
  private eventEnd = 0;

  // DOCTYPE
  /** The text of the markup declaration being read in the internal subset. */
  private declaration = "";
  /** Whether that declaration outgrew what is gathered of it: see {@link appendDeclaration}. */
  private declarationOverflow = false;
  /** Whether a parameter-entity reference was met: declarations after it are not processed. */
  private sawParameterEntityReference = false;
  /** Whether the DOCTYPE names an external subset, which is never read. */
  private externalSubset = false;
  /**
   * What the first declaration of each entity in the internal subset bound — whether or not entities
   * are expanded — with parameter entities as `%name`: what an attribute default's references are
   * checked against.
   */
  private entityBindings: Map<string, EntityBinding> | undefined;
  /** Entities an attribute default may refer to, already found sound. */
  private soundDefaultEntities: Set<string> | undefined;
  /** Whether the XML declaration says `standalone="yes"`. */
  private standalone = false;
  /**
   * What the internal subset's `<!ATTLIST>` declarations say of each element's attributes — the ones
   * with a default or a type other than CDATA — or what the `context` gave: see {@link applyAttributeRules}.
   */
  private attributeRules: Record<string, Record<string, XmlAttributeDeclaration>> | undefined;
  /** Attributes a declaration has bound, as `element\0attribute`: the first declaration binds (§3.3). */
  private boundAttributes: Set<string> | undefined;
  /**
   * Defaults that refer to an entity this parser cannot expand, by `element\0attribute`: an element
   * that leaves the attribute out is an error, as the reference would be in content.
   */
  private unresolvedDefaults: Map<string, string> | undefined;

  /** General entities the internal subset declared, or the `context` gave. */
  private declared: Record<string, string> = Object.create(null) as Record<string, string>;
  private declaredEntities: DeclaredEntities | undefined;
  /** External general entities the internal subset declared: never read, named when used. */
  private externalEntities: Set<string> | undefined;
  private contextSize = new ContextSize();
  /**
   * Characters of what is kept of the internal subset's declarations to read the rest by — entity
   * bindings, the attributes bound first, rules not carried in a context — bounded by
   * `maxContextBytes` whether or not `internalEntities` gathers a context.
   */
  private keptDeclarations = 0;
  private rootName: string | undefined;

  /**
   * Named entities this parser expands: the five XML predefines them (`lt`, `gt`, `amp`,
   * `quot`, `apos`). Add to it before writing to define more — entities declared in a
   * DOCTYPE internal subset are read only with `internalEntities`. Expansions of entities other
   * than the five count against `maxEntityExpansions`.
   */
  ENTITIES: Record<string, string> = { ...XML_ENTITIES };
  /**
   * Whether the document's DOCTYPE carried an internal subset. Without `internalEntities` its declarations are checked,
   * not read — an entity declared there is not defined here — so an undefined entity says why rather than looking
   * like a typo.
   */
  private sawInternalSubset = false;

  // Event handlers
  private _handlers: XmlSaxHandlers = {};

  /**
   * Whether to keep `prefixedAttributes`. Off unless a consumer asks, so a parse that does not need it pays one
   * field test per attribute.
   *
   * @internal
   */
  countPrefixedAttributes = false;
  /**
   * How many attribute names with a prefix (`a:b`, `xmlns:a`) have been read while `countPrefixedAttributes` is
   * set. A consumer compares it across an `opentag` to learn whether that tag has one without enumerating its
   * attributes — the namespace-prefix normaliser, which would otherwise search the text for them.
   *
   * @internal
   */
  prefixedAttributes = 0;

  /**
   * The text a capture has gathered from earlier chunks, or `undefined` while nothing is captured.
   * Captured text is the document as written — before line ends are normalised or references expanded.
   */
  private captured: string | undefined;
  /** Where in `chunk` the capture continues from. */
  private captureFrom = 0;
  /** Whether to keep the text of markup that spans chunks, so a capture can start at its `<`. */
  private keepMarkup = false;
  /** The names of the elements a capture may begin at: see {@link keepMarkupText}. */
  private captureNames: ReadonlySet<string> | undefined;
  /** The markup being read, from its `<` to the end of the previous chunk, while `keepMarkup`. */
  private markupCarry = "";
  /** Where in `chunk` the markup being read began, or `-1` when it began in an earlier chunk. */
  private markupIndex = -1;
  /** The prefix attribute keys are written with, when {@link prefixAttributeNames} set one. */
  private keyPrefix: string | undefined;
  private prefixedKeys: Map<string, string> | undefined;
  /** Where in `chunk` the name being read began, or `-1` when it began in an earlier chunk. */
  private nameIndex = -1;
  /** Names met lately, by a hash of their length and ends: see {@link nameAt}. */
  private readonly names: Array<string | undefined> = new Array<string | undefined>(NAME_SLOTS);
  /** The bytes beyond one per unit each cached name takes in UTF-8. */
  private readonly nameBytes = new Int32Array(NAME_SLOTS);

  /**
   * A parser for options this module has already checked — `Xml.readRecords` checks its own at its
   * boundary, then builds a parser for each range it reads. Checking a large `context` again for
   * each of them cost more than reading a 64 KB range. The options must be the module's own, not
   * the caller's object: what was checked is then what is used.
   *
   * @internal
   */
  static withCheckedOptions(options: XmlSaxOptions): SaxParser {
    optionsChecked = true;
    try {
      return new SaxParser(options);
    } finally {
      optionsChecked = false;
    }
  }

  constructor(given?: XmlSaxOptions) {
    // Taken by this construction alone: a parser built while this one starts — from a getter on
    // its options, say — is checked as any other.
    const checked = optionsChecked;
    optionsChecked = false;
    let options = given;
    if (!checked) {
      // Read once, so what is checked is what is used.
      options = snapshotSaxOptions(given);
      checkSaxOptions("Xml.SaxParser", options);
    }
    this.trackPosition = options?.position !== false;
    this.fileName = options?.fileName;
    this.fragment = options?.fragment ?? false;
    this.skipLeading = (options?.skipLeadingText ?? false) && !this.fragment && !options?.resume;
    this.xmlns = options?.xmlns ?? false;
    this.maxDepth = options?.maxDepth !== undefined ? options.maxDepth : 256;
    this.maxEntityExpansions =
      options?.maxEntityExpansions !== undefined ? options.maxEntityExpansions : 10000;
    this.maxEntityExpansionLength = options?.maxEntityExpansionLength ?? 10_000_000;
    this.invalidCharHandling = options?.invalidCharHandling ?? "error";
    this.byteOffset = options?.byteOffset ?? 0;
    this.resume = options?.resume ?? false;
    this.internalEntities = options?.internalEntities ?? false;
    this.maxContextBytes = options?.maxContextBytes ?? 65536;
    this.context = options?.context;
    this._init();
  }

  /** Whether {@link close} has ended the document. A later {@link write} starts a new one. */
  get closed(): boolean {
    return this._closed;
  }

  /** Whether {@link stop} has ended the parse: {@link write} does nothing until {@link close}. */
  get stopped(): boolean {
    return this._stopped;
  }

  /** Characters consumed so far, counted across every chunk written. */
  get position(): number {
    return this.chunkPosition + this.i;
  }

  /**
   * How much `skipLeadingText` passed over before the document's first `<` — bytes when given bytes,
   * characters when given text — or 0 when nothing was, or the option is off. It stays readable after
   * `close()`, until the parser is given the next document.
   */
  get skippedLeadingText(): number {
    return this.skippedLeading;
  }

  /**
   * Offset in the file of the first byte of the event being reported, when the parser is given
   * bytes — `byteOffset` counts toward it. Read it inside a handler: for a tag it is the `<`, for
   * text the byte after the markup before it. `undefined` when the parser is given text.
   */
  get startByte(): number | undefined {
    return this.bytes === true ? this.eventStart : undefined;
  }

  /**
   * Offset in the file just past the last byte of the event being reported, when the parser is
   * given bytes: after the `>` of a tag, comment, CDATA section or processing instruction, at the
   * `<` that ends a text. `undefined` when the parser is given text.
   */
  get endByte(): number | undefined {
    return this.bytes === true ? this.eventEnd : undefined;
  }

  /**
   * Offset in the file of the next byte to parse, when the parser is given bytes: everything
   * before it has been read and reported. `undefined` when the parser is given text.
   */
  get byte(): number | undefined {
    return this.bytes === true ? this.unitAt(this.i) : undefined;
  }

  /** How many elements are open. In `resume` mode, counted from where the parser started. */
  get depth(): number {
    return this.tags.length;
  }

  /**
   * Where the markup being read began — the `<` of a tag, comment or other construct the parser is
   * inside — when given bytes; `undefined` between markup.
   *
   * @internal
   */
  get openMarkupByte(): number | undefined {
    return this.bytes === true && !this.inText() ? this.markupStart : undefined;
  }

  /**
   * The name of the start tag being read, once its name is whole and before its `>`: the input may
   * end there, inside a tag that is known to be a record's.
   *
   * @internal
   */
  get openStartTagName(): string | undefined {
    let state = this.state;
    if (state === S_ENTITY) {
      state = this.entityReturnState;
    }
    return state === S_OPEN_TAG_SLASH ||
      state === S_ATTRIB ||
      state === S_ATTRIB_NAME ||
      state === S_ATTRIB_NAME_SAW_WHITE ||
      state === S_ATTRIB_VALUE ||
      state === S_ATTRIB_VALUE_QUOTED ||
      state === S_ATTRIB_VALUE_CLOSED
      ? this.tag?.name
      : undefined;
  }

  /** Whether the parser is reading text — a reference inside text included — rather than markup. */
  private inText(): boolean {
    const state = this.state;
    return state === S_TEXT || (state === S_ENTITY && this.entityReturnState === S_TEXT);
  }

  /**
   * In `resume` mode, whether a tag has been read since the start. A parser that reaches the end of
   * the input without one cannot tell trailing comments and white space after the root element from
   * white space between two elements of a file cut short.
   *
   * @internal
   */
  get sawAnyTag(): boolean {
    return this.sawTag;
  }

  /**
   * In `resume` mode, whether no markup has been read since the start: all read so far was taken for
   * text, though it may have been the end of markup begun before the start.
   *
   * @internal
   */
  get readNoMarkup(): boolean {
    return this.leadingResume;
  }

  /**
   * Whether all the text read and not yet reported is white space — with a `text` handler, which is
   * what holds it — and no reference or character is half read.
   *
   * @internal
   */
  get holdsBlankText(): boolean {
    return (
      this.state === S_TEXT &&
      this.decoder?.holdsPartialCharacter !== true &&
      (this.carriedFromPrevious === undefined || isS(this.carriedFromPrevious.charCodeAt(0))) &&
      isWhitespaceText(this.text)
    );
  }

  /**
   * The names of the elements open, outermost first — in `resume` mode, those opened since the start.
   *
   * @internal
   */
  get openTagNames(): string[] {
    return this.tags.map(tag => tag.name);
  }

  /**
   * Whether the root element's start tag has been read — in `resume` mode, whether the parser was
   * given the context of one.
   *
   * @internal
   */
  get sawRootStart(): boolean {
    return this.rootName !== undefined;
  }

  /**
   * Whether the markup being read is a start tag, or a `<` not yet followed by what says which
   * markup it begins.
   *
   * @internal
   */
  get inStartTag(): boolean {
    const state = this.state;
    return (
      state === S_OPEN_WAKA ||
      (state >= S_OPEN_TAG && state <= S_ATTRIB_VALUE_CLOSED) ||
      // A reference in an attribute value: still inside the start tag.
      (state === S_ENTITY && this.entityReturnState === S_ATTRIB_VALUE_QUOTED)
    );
  }

  private _init(): void {
    this.state = S_TEXT;
    this.text = "";
    this.name = "";
    this.q = null;
    this.tags = [];
    this.tag = null;
    this.duplicateAttributes = null;
    this.entity = "";
    this.openWakaBang = "";
    this.sawRoot = this.fragment || this.resume;
    this.sawDoctype = false;
    this.closedRoot = this.fragment;
    this.reportedTextBeforeRoot = this.fragment || this.resume;
    this.beforeDocument = this.skipLeading;
    this.reportedTextAfterRoot = this.fragment || this.resume;
    this.carriedFromPrevious = undefined;
    this._bomStripped = false;
    this._stopped = false;
    this._paused = false;
    this.switchToAscii = false;
    this.asciiDeclared = false;
    this.sawTag = false;
    this.rootEnded = false;
    this.textBrackets = 0;
    this.leadingResume = this.resume;
    // Note: _closed is NOT reset here — it is managed by end() and write().
    this.line = 1;
    this.column = 0;
    this.positionAtNewLine = 0;
    this.chunkPosition = 0;
    this.limit = 0;
    this.chunk = "";
    this.i = 0;
    this.prevI = 0;
    this._nsStack = this.xmlns
      ? [
          Object.assign(Object.create(null) as Record<string, string>, {
            xml: "http://www.w3.org/XML/1998/namespace",
            xmlns: "http://www.w3.org/2000/xmlns/"
          })
        ]
      : [];
    this._entityExpansionCount = 0;
    this.entityExpansionLength = 0;
    this.sawInternalSubset = false;
    this.decoder = undefined;
    this.bytes = undefined;
    this.chunkUnit = 0;
    this.chunkOneUnit = true;
    this.extra = 0;
    this.markupStart = 0;
    this.documentStart = 0;
    this.textStart = 0;
    this.eventStart = 0;
    this.eventEnd = 0;
    this.declaration = "";
    this.declarationOverflow = false;
    this.sawParameterEntityReference = false;
    this.externalSubset = false;
    this.entityBindings = undefined;
    this.soundDefaultEntities = undefined;
    this.standalone = false;
    this.attributeRules = undefined;
    this.boundAttributes = undefined;
    this.unresolvedDefaults = undefined;
    this.externalEntities = undefined;
    this.rootName = undefined;
    this.contextSize = new ContextSize();
    this.keptDeclarations = 0;
    const context = this.context;
    if (context === undefined) {
      this.declared = Object.create(null) as Record<string, string>;
    } else {
      // The context's entities are shared, not copied: a reading of a range builds a parser at every
      // boundary it checks, and copying a large table for each cost more than reading the range. The
      // table is the parser's own snapshot, with no prototype; what this document declares is its own.
      const entities = context.entities;
      this.declared = Object.create(
        Object.getPrototypeOf(entities) === null
          ? entities
          : Object.assign(Object.create(null) as Record<string, string>, entities)
      ) as Record<string, string>;
      this.rootName = context.root;
      this.attributeRules = context.attributes as typeof this.attributeRules;
    }
    this.declaredEntities = undefined;
    this.captured = undefined;
    this.markupCarry = "";
    this.markupIndex = -1;
  }

  // ===========================================================================
  // Event Registration
  // ===========================================================================

  /**
   * Set the handler for an event. Each event has **one** handler: a second `on` for the same
   * event replaces the first (it does not add a listener, as `EventEmitter.on` would).
   * Without an `error` handler, the first error is thrown out of {@link write} / {@link close}.
   */
  on(name: "opentag", handler: (tag: XmlSaxTag) => void): void;
  on(name: "text", handler: (text: string) => void): void;
  on(name: "closetag", handler: (tag: XmlSaxTag) => void): void;
  on(name: "cdata", handler: (text: string) => void): void;
  on(name: "comment", handler: (text: string) => void): void;
  on(name: "pi", handler: (target: string, body: string) => void): void;
  on(name: "error", handler: (err: Error) => void): void;
  on<K extends keyof XmlSaxHandlers>(name: K, handler: XmlSaxHandlers[K]): void {
    checkEventName("on", name);
    if (typeof handler !== "function") {
      throw new XmlArgumentError(
        "handler",
        `Xml.SaxParser#on: the ${String(name)} handler must be a function`
      );
    }
    this._handlers[name] = handler;
  }

  /** Remove the handler for an event. */
  off(name: keyof XmlSaxHandlers): void {
    checkEventName("off", name);
    delete this._handlers[name];
  }

  /** Return the currently-registered handler for an event, if any. */
  getHandler<K extends keyof XmlSaxHandlers>(name: K): XmlSaxHandlers[K] {
    return this._handlers[name];
  }

  /**
   * Set or clear a handler without checking it, for a consumer inside this module that switches
   * handlers as it goes (the record reader reads text only inside a record): the text of an event
   * nobody handles is never sliced out.
   *
   * @internal
   */
  setHandler<K extends keyof XmlSaxHandlers>(
    name: K,
    handler: XmlSaxHandlers[K] | undefined
  ): void {
    this._handlers[name] = handler;
  }

  // ===========================================================================
  // State Reporting
  // ===========================================================================

  /**
   * A value that names where the parser is in the grammar of the document, comparable with `===`,
   * for telling whether two parsers reading **one input** have reached the same state at the same
   * position — one that started at the beginning, say, and one that started further on (`resume`).
   * When they have, they report the same events from there on and keep reaching the same lexical
   * states; only errors that depend on the elements open around them may differ — an end tag that
   * matches none, or markup after the root element has ended.
   *
   * It holds what decides how the next character is read — the state of the scanner, an open
   * quote, a half-read reference, whether the parser is still in the prolog — and where the markup or
   * text being read began. It does not hold what has been gathered since: that is the input from
   * that point to here, the same for both parsers only because they read the same input. Two
   * parsers given different inputs can have equal lexical states and report different events next.
   * Nor does it hold the elements open around the parser. This is how `Xml.readRecords` learns that
   * a chunk was read from where it was assumed to start.
   *
   * Positions are bytes when the parser is given bytes, characters when it is given text.
   * Constant time: a few fields joined.
   */
  getLexicalState(): string {
    const state = this.state;
    let key = `${state}:${this.inText() ? this.textStart : this.markupStart}:${this.phase()}`;
    if (this.q !== null) {
      key += `:q${this.q}`;
    }
    if (state === S_ENTITY) {
      key += `:&${this.entityReturnState}:${this.entity}`;
    } else if (state === S_OPEN_WAKA_BANG) {
      key += `:!${this.openWakaBang}`;
    }
    return key;
  }

  /**
   * Whether the parser is in the prolog — where a DOCTYPE may appear and an element begins the
   * root — or past it. After the root element ends a parser is past it still: that the root has
   * ended is a fact about the elements open around it, which a lexical state leaves out.
   */
  private phase(): string {
    return this.sawRoot ? PHASE_CONTENT : PHASE_PROLOG;
  }

  /**
   * What a reader started later in this file needs to know of its prolog: the name of the root
   * element and the general entities the DOCTYPE internal subset declared (read with
   * `internalEntities`). Pass it as `context` to a parser that starts inside the root element
   * (`resume`), which then expands those entities as this one does.
   *
   * `undefined` until the root element's start tag has been read. In `resume` mode, the `context`
   * the parser was given. The result can be serialised with `JSON.stringify`.
   */
  getPrologContext(): XmlPrologContext | undefined {
    if (this.context !== undefined) {
      return this.context;
    }
    if (this.rootName === undefined) {
      return undefined;
    }
    const entities = Object.create(null) as Record<string, string>;
    for (const name of Object.keys(this.declared)) {
      entities[name] = this.declared[name];
    }
    // How the bytes after the prolog are decoded is part of what they mean.
    const encoding = this.asciiDeclared ? { encoding: "ascii" as const } : {};
    const rules = this.attributeRules;
    if (rules === undefined) {
      return { root: this.rootName, entities, ...encoding };
    }
    const attributes = Object.create(null) as Record<
      string,
      Record<string, XmlAttributeDeclaration>
    >;
    for (const element of Object.keys(rules)) {
      attributes[element] = Object.assign(Object.create(null), rules[element]) as Record<
        string,
        XmlAttributeDeclaration
      >;
    }
    return { root: this.rootName, entities, ...encoding, attributes };
  }

  // ===========================================================================
  // Error Handling
  // ===========================================================================

  private makeError(message: string, limit?: XmlParseLimit, truncated?: boolean): XmlParseError {
    const context = {
      line: this.trackPosition ? this.line : undefined,
      column: this.trackPosition ? this.column : undefined,
      fileName: this.fileName,
      byte: this.bytes === true ? this.exactUnitAt(Math.min(this.prevI, this.i)) : undefined
    };
    if (limit !== undefined) {
      const allowed =
        limit === "depth"
          ? this.maxDepth
          : limit === "prologContext"
            ? this.maxContextBytes
            : limit === "entityExpansionLength"
              ? this.maxEntityExpansionLength
              : this.maxEntityExpansions;
      return new XmlLimitExceededError(message, limit, allowed, context);
    }
    if (truncated) {
      return new XmlTruncatedError(message, context);
    }
    return new XmlParseError(message, context);
  }

  /**
   * Report a parse error at the current position: to the `error` handler if one is set,
   * otherwise by throwing it. For a consumer that validates content as it is parsed and wants
   * its own errors to carry the parser's position.
   */
  fail(message: string, limit?: XmlParseLimit): this {
    return this.report(this.makeError(message, limit));
  }

  /** Report an input that ends where the document cannot, as an {@link XmlTruncatedError}. */
  private failTruncated(message: string): void {
    this.report(this.makeError(message, undefined, true));
  }

  private report(err: XmlParseError): this {
    if (this._stopped) {
      // A handler stopped the parse: nothing more is reported, an error found on the way out included.
      return this;
    }
    if (this._handlers.error) {
      this._handlers.error(err);
    } else {
      throw err;
    }
    return this;
  }

  // ===========================================================================
  // Main Write Method
  // ===========================================================================

  /**
   * Parse the next chunk of the document; `null` ends it, like {@link close}. Writing after
   * the document has been closed starts a new document.
   *
   * A chunk is text, or the document's bytes as a `Uint8Array`, which are decoded as UTF-8 — a
   * character split between two chunks is read whole — and give events their byte positions
   * ({@link startByte}, {@link endByte}). One document is given one or the other, not both. A
   * byte that is not part of well-formed UTF-8 is an invalid character, handled as
   * `invalidCharHandling` says.
   */
  write(chunk: string | Uint8Array | null): this {
    if (this._writing) {
      // A handler writing to the parser that called it would parse over the state the outer write
      // is in the middle of: refused, as the outer write would otherwise read past its own chunk.
      throw new XmlError(
        "Xml.SaxParser#write: called from a handler while the parser is parsing; write after it returns"
      );
    }
    this._writing = true;
    try {
      return this.writeChunk(chunk);
    } finally {
      this._writing = false;
    }
  }

  private writeChunk(chunk: string | Uint8Array | null): this {
    if (this._paused) {
      // Only this module's record reader pauses a parser; writing past a pause would lose the rest of the chunk.
      throw new XmlError("Xml.SaxParser#write: the parser is paused inside a chunk");
    }
    const isText = typeof chunk === "string";
    if (chunk !== null && !isText && !isUint8Array(chunk)) {
      throw new XmlArgumentError(
        "chunk",
        `Xml.SaxParser#write: chunk must be a string, a Uint8Array or null, got ${typeof chunk}`
      );
    }
    if (this._stopped) {
      // A stopped parse ignores the rest of its input; ending it starts the next document afresh.
      if (chunk === null) {
        this._closed = true;
        this._init();
      }
      return this;
    }
    // Auto-reset for parser reuse: if previously closed, start fresh.
    if (this._closed) {
      if (chunk === null) {
        return this; // no-op: already closed
      }
      this._closed = false;
      this._init();
      // A new document: what was skipped before the last one is kept until here, for after `close()`.
      this.skippedLeading = 0;
    }
    if (chunk === null) {
      if (this.decoder !== undefined) {
        const text = this.decoder.decode(EMPTY_BYTES, true);
        return this.process(
          text,
          true,
          this.decoder.textStart,
          this.decoder.textEnd,
          this.decoder.oneBytePerUnit
        );
      }
      const unit = this.chunkPosition + this.limit;
      return this.process("", true, unit, unit, true);
    }
    if (this.bytes === undefined) {
      this.bytes = !isText;
      if (!isText) {
        this.decoder = new Utf8ChunkDecoder(
          this.byteOffset,
          this.resume,
          this.context?.encoding === "ascii"
        );
        this.markupStart = this.textStart = this.documentStart = this.byteOffset;
        // A byte order mark is the file's first bytes; a parser that starts later has none to strip.
        this._bomStripped = this.byteOffset !== 0;
      }
    } else if (this.bytes === isText) {
      throw new XmlArgumentError(
        "chunk",
        `Xml.SaxParser#write: this document was given ${this.bytes ? "bytes" : "text"}; it cannot be given ${isText ? "text" : "bytes"} too`
      );
    }
    if (isText) {
      const unit = this.chunkPosition + this.limit + (this.carriedFromPrevious?.length ?? 0);
      return this.process(chunk, false, unit, unit + chunk.length, true);
    }
    const decoder = this.decoder!;
    if (decoder.next + chunk.length > Number.MAX_SAFE_INTEGER) {
      // Past it a byte position rounds, and every position reported after it could be wrong.
      throw new XmlArgumentError(
        "chunk",
        "Xml.SaxParser#write: the input would run past byte Number.MAX_SAFE_INTEGER, where byte positions are no longer exact"
      );
    }
    const text = decoder.decode(chunk, false);
    return this.process(text, false, decoder.textStart, decoder.textEnd, decoder.oneBytePerUnit);
  }

  /**
   * Parse `chunk`, which lies at positions `[unit, endUnit)`; `oneUnit` says each of its UTF-16 units
   * is one unit of position.
   */
  private process(
    chunk: string,
    end: boolean,
    unit: number,
    endUnit: number,
    oneUnit: boolean
  ): this {
    if (this.carriedFromPrevious !== undefined) {
      oneUnit = oneUnit && this.carriedFromPrevious.charCodeAt(0) < 0x80;
      chunk = this.carriedFromPrevious + chunk;
      unit = this.carriedUnit;
      this.carriedFromPrevious = undefined;
    }

    // Strip the byte order mark (U+FEFF) from the very beginning of input.
    if (!this._bomStripped && chunk.length > 0) {
      this._bomStripped = true;
      if (chunk.charCodeAt(0) === 0xfeff) {
        chunk = chunk.slice(1);
        if (this.bytes === true) {
          unit += 3;
          oneUnit = oneUnit || chunk.length === 0;
          this.markupStart = this.textStart = this.documentStart = unit;
        }
      }
    }

    // The chunk before is done with: its length joins the count. Added here rather than as it
    // ends, so `position` — `chunkPosition + i` — holds between writes as well as inside one.
    this.chunkPosition += this.limit;
    this.limit = 0;
    this.chunk = chunk;
    this.cdataEndAt = 0;
    this.chunkUnit = unit;
    this.chunkOneUnit = oneUnit;
    this.extra = 0;
    this.i = 0;
    this.markupIndex = -1;
    this.captureFrom = 0;
    this.nameIndex = -1;

    let limit = chunk.length;
    if (!end && limit > 0) {
      const lastCode = chunk.charCodeAt(limit - 1);
      // A CR waits to see whether an LF follows — unless the decoder holds the start of a character
      // that has not arrived whole, which is what follows and is no LF. Holding it as well would leave
      // the CR, and everything after, unread until more input came.
      if (
        (lastCode === CR && this.decoder?.holdsPartialCharacter !== true) ||
        (lastCode >= 0xd800 && lastCode <= 0xdbff)
      ) {
        // The carried character is the last one, and one unit of position (a CR, or half a pair of text).
        this.carriedUnit = endUnit - 1;
        this.carriedFromPrevious = chunk[limit - 1];
        limit--;
        chunk = chunk.slice(0, limit);
        this.chunk = chunk;
      }
    }

    this.limit = limit;
    this.ending = end;
    this.cdataEndAt = this.firstCDataEnd();
    this.run();
    if (this.switchToAscii) {
      // The declaration just read says ASCII: what follows it was decoded as UTF-8, and is decoded
      // again — the CR held back for the next chunk among it.
      // A consumer's pause made in the same event stays: the re-decoded text waits for resumeParsing.
      this.switchToAscii = false;
      const decoder = this.decoder!;
      const text = decoder.switchToAscii(this.unitAt(this.i));
      this.limit = this.i;
      this.carriedFromPrevious = undefined;
      return this.process(text, end, decoder.textStart, decoder.textEnd, true);
    }
    return this;
  }

  /**
   * Where in `chunk` the `>` of the first `]]>` at or after the reading position is — the two `]`
   * before it may be the end of the chunk before — or `Infinity`. Found with one `indexOf` per
   * chunk, and again only past each one met, so text is not scanned for it a character at a time.
   */
  private findCDataEnd(from: number): number {
    const at = this.chunk.indexOf("]]>", from);
    return at < 0 ? Infinity : at + 2;
  }

  /** {@link findCDataEnd} from the start of a chunk, whose `]]>` may have begun in the chunk before. */
  private firstCDataEnd(): number {
    const chunk = this.chunk;
    const brackets = this.textBrackets;
    if (brackets > 0) {
      // A `]]>` split between chunks: its first `]`, or both, ended the text of the chunk before.
      if (chunk.charCodeAt(0) === GREATER && brackets >= 2) {
        return 0;
      }
      if (chunk.charCodeAt(0) === CLOSE_BRACKET && chunk.charCodeAt(1) === GREATER) {
        return 1;
      }
    }
    return this.findCDataEnd(0);
  }

  /**
   * Report each `]]>` whose `>` the text run ending at `end` reads: one in text, not one that ended
   * a CDATA section or sat in an attribute value or a comment, which an earlier state read. A `]` is
   * never the end of markup, so a `>` read as text makes its `]]>` text whole. Read from a start
   * inside the file, before any markup, the text may be the end of a CDATA section the start cut, and
   * is not judged: a reading from earlier reads it, and does.
   */
  private checkCDataEnds(end: number): void {
    let at = this.cdataEndAt;
    while (at < end && !this._stopped) {
      if (at >= this.i && !this.leadingResume) {
        const column = this.column;
        const i = this.i;
        if (this.trackPosition) {
          this.column += at - i;
        }
        // Read ahead of the cursor: the error is where the `]]>` is, not where the text run began.
        this.prevI = this.i = at;
        this.fail("]]> is not allowed in text");
        if (this._stopped) {
          // The error's handler stopped the parse there: the position stays at the `]]>` it reported.
          break;
        }
        this.i = i;
        this.column = column;
      }
      at = this.findCDataEnd(Math.max(at - 1, 0));
    }
    this.cdataEndAt = at;
  }

  /** Parse the chunk set up by {@link process} from where it stands, then finish it. */
  private run(): this {
    const { chunk, limit } = this;
    while (this.i < limit && !this._stopped && !this.halted) {
      this.processState();
    }
    if (this._stopped || this.halted) {
      return this;
    }

    if (this.captured !== undefined) {
      this.captured += chunk.slice(this.captureFrom);
    }
    // The `]` this chunk's text ends with, for a `]]>` split between chunks.
    let brackets = 0;
    if (this.state === S_TEXT) {
      while (
        brackets < 2 &&
        brackets < limit &&
        chunk.charCodeAt(limit - 1 - brackets) === CLOSE_BRACKET
      ) {
        brackets++;
      }
      if (brackets === limit) {
        brackets = Math.min(2, brackets + this.textBrackets);
      }
    }
    this.textBrackets = brackets;
    if (this.keepMarkup && this.state !== S_TEXT) {
      if (this.captured === undefined && this.mayBeCaptured()) {
        this.markupCarry += this.markupIndex >= 0 ? chunk.slice(this.markupIndex) : chunk;
      } else {
        // Nothing will capture from this markup's `<`, or a capture already holds it: a long comment
        // outside every record is not kept.
        this.markupCarry = "";
      }
    }

    return this.ending ? this.end() : this;
  }

  /**
   * Called, in `resume` mode, with the name of an end tag that closes an element opened before the
   * parser started — which reports no `closetag`, having no tag to report. For this module's record
   * reader, which accounts for the structure of what it reads.
   *
   * @internal
   */
  onOuterClose: ((name: string) => void) | undefined = undefined;

  /**
   * Pause after the event being reported, keeping the rest of the chunk: {@link resumeParsing} goes
   * on from there. For a consumer that must wait between two events — the record reader, until a
   * record's callback has settled — where {@link stop} would end the parse.
   *
   * @internal
   */
  pause(): void {
    this._paused = true;
  }

  /**
   * Inside a `pi` handler, whether the instruction is in the DOCTYPE internal subset rather than
   * among the document's own nodes.
   *
   * @internal
   */
  get inInternalSubset(): boolean {
    return this.state === S_DOCTYPE_DTD_PI_ENDING;
  }

  /**
   * Whether reading the chunk must stop where it stands: a consumer's {@link pause}, or a declaration
   * of ASCII, after which the rest is decoded again. Kept apart so the one never clears the other.
   */
  private get halted(): boolean {
    return this._paused || this.switchToAscii;
  }

  /** Whether the parser is paused with part of a chunk unread. @internal */
  get paused(): boolean {
    return this._paused;
  }

  /**
   * Go on with the chunk a {@link pause} left, to its end or the next pause.
   *
   * @internal
   */
  resumeParsing(): this {
    if (this._writing) {
      // As for write: resuming from a handler would parse over the state the outer call is in.
      throw new XmlError(
        "Xml.SaxParser#resumeParsing: called from a handler while the parser is parsing; resume after it returns"
      );
    }
    this._paused = false;
    this._writing = true;
    try {
      return this.run();
    } finally {
      this._writing = false;
    }
  }

  /** End the document: report anything left unclosed and flush the last text. */
  close(): this {
    return this.write(null);
  }

  /**
   * Stop parsing: nothing after the event being reported is read, the rest of the chunk being
   * written included, and later {@link write} calls do nothing. Call it from a handler once the
   * parser has given what was wanted — the root element's start tag, say. {@link close} then ends
   * the document without reporting what was left unread, and a later write starts a new one.
   */
  stop(): this {
    this._stopped = true;
    return this;
  }

  // ===========================================================================
  // Positions
  // ===========================================================================

  /**
   * The position of `chunk[index]` — bytes when given bytes, characters when given text — where
   * `index` is the point being read: every character before it has been scanned, none after.
   */
  private unitAt(index: number): number {
    return this.chunkOneUnit ? this.chunkUnit + index : this.chunkUnit + index + this.extra;
  }

  /** The position of any `chunk[index]`, by counting: for an error, off the fast path. */
  private exactUnitAt(index: number): number {
    return this.chunkOneUnit
      ? this.chunkUnit + index
      : this.chunkUnit + utf8Length(this.chunk, 0, index);
  }

  /** Mark the `<` at `chunk[index]` as the start of the markup being read. */
  private startMarkup(index: number): void {
    this.leadingResume = false;
    this.markupStart = this.unitAt(index);
    if (this.keepMarkup) {
      this.markupIndex = index;
      this.markupCarry = "";
    }
  }

  /** The markup that just ended was the event `[markupStart, here)`; text starts here. */
  private markupEvent(): void {
    this.eventStart = this.markupStart;
    this.eventEnd = this.textStart = this.unitAt(this.i);
  }

  /** Back to text after markup, which ended at the current position. */
  private endMarkup(): void {
    this.state = S_TEXT;
    this.textStart = this.unitAt(this.i);
  }

  // ===========================================================================
  // Capture
  // ===========================================================================

  /**
   * Keep the text of a start tag that spans chunks, so {@link beginCapture} can start at a `<` that
   * arrived in an earlier chunk — of an element named in `names`, the only ones a capture begins at.
   * Off unless a capture is wanted, as it copies what it keeps.
   *
   * @internal
   */
  keepMarkupText(names: ReadonlySet<string>): void {
    this.keepMarkup = true;
    this.captureNames = names;
  }

  /**
   * Whether a capture may begin at the markup being read: only a start tag can begin one, and only of
   * an element named in {@link captureNames} — while its name is not yet whole, any start tag.
   */
  private mayBeCaptured(): boolean {
    const name = this.openStartTagName;
    if (name !== undefined) {
      return this.captureNames!.has(name);
    }
    return this.state === S_OPEN_WAKA || this.state === S_OPEN_TAG;
  }

  /**
   * Start capturing the document's text, from the `<` of the markup just reported. Needs
   * {@link keepMarkupText}.
   *
   * @internal
   */
  beginCapture(): void {
    const from = this.markupIndex >= 0 ? this.markupIndex : 0;
    this.captured = this.markupCarry + this.chunk.slice(from, this.i);
    this.captureFrom = this.i;
  }

  /**
   * End the capture at the current position — just after the markup just reported — and return
   * the text.
   *
   * @internal
   */
  endCapture(): string {
    const text = this.captured + this.chunk.slice(this.captureFrom, this.i);
    this.captured = undefined;
    return text;
  }

  // ===========================================================================
  // Invalid Character Handling
  // ===========================================================================

  /** The message for an invalid character: a byte that is not UTF-8 is named as the byte. */
  private invalidCharLabel(code: number, kind?: string): string {
    const byte = this.bytes === true ? escapedByte(code) : -1;
    if (byte >= 0) {
      return `invalid UTF-8: byte 0x${byte.toString(16)} is not part of a well-formed sequence`;
    }
    return kind
      ? `invalid XML character: ${kind} 0x${code.toString(16)}`
      : `invalid XML character: 0x${code.toString(16)}`;
  }

  /**
   * Handle an invalid XML character according to the configured strategy.
   *
   * Used by `handleTextInRoot()` fast path which manages its own text accumulation
   * and cannot use the `getCode()` loop approach.
   *
   * - `"error"`: call `fail()` and return the original code.
   * - `"skip"`: return `REPLACEMENT_CHAR` as a sentinel (caller handles skip).
   * - `"replace"`: return `REPLACEMENT_CHAR`.
   *
   * Note: For `getCode()`, invalid char handling is inlined to avoid recursion.
   *
   * @param code - The invalid character code point.
   * @param kind - Optional description (e.g. "lone surrogate") for error messages.
   * @returns The code point to use.
   */
  private handleInvalidChar(code: number, kind?: string): number {
    switch (this.invalidCharHandling) {
      case "replace":
        return REPLACEMENT_CHAR;
      case "skip":
        // Caller is responsible for the actual skip logic.
        // We return -2 as a sentinel to tell getCode()'s loop to continue.
        return -2;
      default: {
        // "error" — existing strict behavior
        const label = this.invalidCharLabel(code, kind);
        this.fail(label);
        return code;
      }
    }
  }

  /**
   * Handle an invalid character inside the `handleTextInRoot()` fast loop.
   *
   * Unlike `handleInvalidChar()` (which returns a code point for `getCode()`),
   * this method manages the text accumulation state (`this.text`, `start`) that
   * the fast text loop relies on.
   *
   * - `"error"`: call `fail()`, leave text accumulation unchanged (char stays in output).
   * - `"skip"`: flush text up to the invalid char, skip it, return new `start`.
   * - `"replace"`: flush text up to the invalid char, append U+FFFD, return new `start`.
   *
   * @returns The updated `start` index for the text accumulation loop.
   */
  private handleInvalidCharInText(
    code: number,
    handler: ((text: string) => void) | undefined,
    start: number,
    kind?: string
  ): number {
    switch (this.invalidCharHandling) {
      case "skip":
        // Flush text accumulated before this invalid char, then skip it
        if (handler && start < this.prevI) {
          this.text += this.chunk.slice(start, this.prevI);
        }
        return this.i;
      case "replace":
        // Flush text accumulated before this invalid char, append replacement
        if (handler) {
          if (start < this.prevI) {
            this.text += this.chunk.slice(start, this.prevI);
          }
          this.text += REPLACEMENT_STR;
        }
        return this.i;
      default: {
        // "error" — existing strict behavior, char stays in output
        const label = this.invalidCharLabel(code, kind);
        this.fail(label);
        return start;
      }
    }
  }

  /**
   * Handle an invalid character inside `sAttribValueQuoted()`.
   *
   * Same pattern as `handleInvalidCharInText()` but for attribute value
   * accumulation (always uses `this.text`, no conditional handler check).
   *
   * @returns The updated `start` index.
   */
  private handleInvalidCharInAttr(code: number, start: number, kind?: string): number {
    switch (this.invalidCharHandling) {
      case "skip":
        if (start < this.prevI) {
          this.text += this.chunk.slice(start, this.prevI);
        }
        return this.i;
      case "replace":
        if (start < this.prevI) {
          this.text += this.chunk.slice(start, this.prevI);
        }
        this.text += REPLACEMENT_STR;
        return this.i;
      default: {
        const label = this.invalidCharLabel(code, kind);
        this.fail(label);
        return start;
      }
    }
  }

  // ===========================================================================
  // Character Reading
  // ===========================================================================

  private getCode(): number {
    // Loop to handle skip mode: when an invalid char returns -2, we retry
    // with the next character instead of recursing (avoids stack overflow
    // on long runs of consecutive invalid characters).
    for (;;) {
      const { chunk } = this;
      const i = this.i;
      this.prevI = i;
      if (i >= chunk.length) {
        // Nothing read: the position stays at the end of the chunk, where `byte` reports it.
        return -1;
      }
      this.i = i + 1;

      const code = chunk.charCodeAt(i);

      // Ultra-fast path: printable ASCII (0x20-0x7E) — the vast majority of XML content.
      // No validation needed; these are always valid XML 1.0 characters.
      if (code >= 0x20 && code <= 0x7e) {
        if (this.trackPosition) {
          this.column++;
        }
        return code;
      }

      // Secondary fast path: TAB (0x09) — common in attribute values
      if (code === TAB) {
        if (this.trackPosition) {
          this.column++;
        }
        return code;
      }

      // Handle CR (normalize CR and CR+LF to LF per XML 1.0 §2.11)
      if (code === CR) {
        if (chunk.charCodeAt(i + 1) === NL) {
          this.i = i + 2;
        }
        if (this.trackPosition) {
          this.line++;
          this.column = 0;
          this.positionAtNewLine = this.position;
        }
        return NL;
      }

      // Handle LF
      if (code === NL) {
        if (this.trackPosition) {
          this.line++;
          this.column = 0;
          this.positionAtNewLine = this.position;
        }
        return NL;
      }

      // Handle surrogates
      if (code >= 0xd800 && code <= 0xdbff) {
        const next = chunk.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) {
          this.i = i + 2;
          this.extra += 2;
          // A column counts UTF-16 units, as `position` does — and as the scans that take a name or
          // a run of text whole count it.
          if (this.trackPosition) {
            this.column += 2;
          }
          return 0x10000 + ((code - 0xd800) * 0x400 + (next - 0xdc00));
        }
        // Lone high surrogate — invalid XML character
        const result = this.handleInvalidChar(code, "lone surrogate");
        if (result !== -2) {
          return result;
        }
        continue; // skip: loop to next char
      }

      // Lone low surrogate — invalid XML character
      if (code >= 0xdc00 && code <= 0xdfff) {
        const result = this.handleInvalidChar(code, "lone surrogate");
        if (result !== -2) {
          return result;
        }
        continue;
      }

      // Non-ASCII above surrogate range (0x80-0xD7FF, 0xE000-0xFFFD) — all valid XML
      if (code >= 0x80) {
        this.extra += code < 0x800 ? 1 : 2;
        if (this.trackPosition) {
          this.column++;
        }
        // Reject 0xFFFE and 0xFFFF
        if (code === 0xfffe || code === 0xffff) {
          const result = this.handleInvalidChar(code);
          if (result !== -2) {
            return result;
          }
          continue;
        }
        return code;
      }

      // Remaining: ASCII control characters (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F)
      // All invalid in XML 1.0
      if (this.trackPosition) {
        this.column++;
      }
      const result = this.handleInvalidChar(code);
      if (result !== -2) {
        return result;
      }
      // skip: continue to next char
    }
  }

  /**
   * The name `chunk[start, end)`. Element and attribute names repeat, and a name made by slicing is a
   * new string that every object key it becomes must hash and intern again, so the last name seen
   * with the same length and ends is kept and returned when the text matches it.
   */
  private nameAt(start: number, end: number): string {
    const chunk = this.chunk;
    const length = end - start;
    if (length > 48) {
      return chunk.slice(start, end);
    }
    const slot = nameSlot(chunk, start, end);
    const known = this.names[slot];
    if (known !== undefined && known.length === length && chunk.startsWith(known, start)) {
      return known;
    }
    const name = chunk.slice(start, end);
    this.names[slot] = name;
    this.nameBytes[slot] = utf8Length(name, 0, length) - length;
    return name;
  }

  /** The bytes beyond one per UTF-16 unit `name` takes, from the name cache when it is there. */
  private extraBytesOf(name: string): number {
    const slot = nameSlot(name, 0, name.length);
    return this.names[slot] === name
      ? this.nameBytes[slot]
      : utf8Length(name, 0, name.length) - name.length;
  }

  /**
   * At the `/` of `</`: when the text is `/name>` for the innermost open element — what nearly every
   * end tag is — close it without scanning the name. Whether it did.
   */
  private closeInnermost(): boolean {
    const tags = this.tags;
    if (tags.length === 0) {
      return false;
    }
    const name = tags[tags.length - 1].name;
    const chunk = this.chunk;
    const at = this.i + 1;
    const end = at + name.length;
    if (chunk.charCodeAt(end) !== GREATER || !chunk.startsWith(name, at)) {
      return false;
    }
    if (!this.chunkOneUnit) {
      this.extra += this.extraBytesOf(name);
    }
    if (this.trackPosition) {
      this.column += end + 1 - this.i;
    }
    this.prevI = end;
    this.i = end + 1;
    this.name = name;
    this.closeTag();
    return true;
  }

  /** The name being read ends at `pos`; the scan that found its end began at `scanFrom`. */
  private finishName(scanFrom: number, pos: number): void {
    if (this.nameIndex >= 0) {
      this.name = this.nameAt(this.nameIndex, pos);
      // Whatever is added to this name from here on (in error recovery) does not follow it in the text.
      this.nameIndex = -1;
    } else if (pos > scanFrom) {
      this.name += this.chunk.slice(scanFrom, pos);
    }
  }

  /**
   * Skip the name characters from `chunk[pos]`, returning where they end — a character outside the
   * Basic Multilingual Plane included — and count the bytes the non-ASCII ones take.
   */
  private scanNameChars(pos: number): number {
    const chunk = this.chunk;
    const len = chunk.length;
    while (pos < len) {
      const code = chunk.charCodeAt(pos);
      if (code < 128) {
        if (NAME_CHAR_ASCII[code] !== 1) {
          break;
        }
        pos++;
      } else if (code >= 0xd800 && code <= 0xdbff) {
        const next = chunk.charCodeAt(pos + 1);
        if (
          !(next >= 0xdc00 && next <= 0xdfff) ||
          !isNameChar(0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00))
        ) {
          break;
        }
        pos += 2;
        this.extra += 2;
      } else {
        if (!isNameChar(code)) {
          break;
        }
        pos++;
        this.extra += code < 0x800 ? 1 : 2;
      }
    }
    return pos;
  }

  private unget(): void {
    this.i = this.prevI;
    if (this.trackPosition) {
      this.column--;
    }
  }

  // ===========================================================================
  // State Machine
  // ===========================================================================

  private processState(): void {
    switch (this.state) {
      case S_TEXT:
        if (this.tags.length !== 0 || this.resume) {
          this.handleTextInRoot();
        } else {
          this.handleTextOutsideRoot();
        }
        break;
      case S_OPEN_WAKA:
        this.sOpenWaka();
        break;
      case S_OPEN_WAKA_BANG:
        this.sOpenWakaBang();
        break;
      case S_OPEN_TAG:
        this.sOpenTag();
        break;
      case S_OPEN_TAG_SLASH:
        this.sOpenTagSlash();
        break;
      case S_ATTRIB:
        this.sAttrib();
        break;
      case S_ATTRIB_NAME:
        this.sAttribName();
        break;
      case S_ATTRIB_NAME_SAW_WHITE:
        this.sAttribNameSawWhite();
        break;
      case S_ATTRIB_VALUE:
        this.sAttribValue();
        break;
      case S_ATTRIB_VALUE_QUOTED:
        this.sAttribValueQuoted();
        break;
      case S_ATTRIB_VALUE_CLOSED:
        this.sAttribValueClosed();
        break;
      case S_CLOSE_TAG:
        this.sCloseTag();
        break;
      case S_CLOSE_TAG_SAW_WHITE:
        this.sCloseTagSawWhite();
        break;
      case S_COMMENT:
        this.sComment();
        break;
      case S_COMMENT_ENDING:
        this.sCommentEnding();
        break;
      case S_COMMENT_ENDED:
        this.sCommentEnded();
        break;
      case S_CDATA:
        this.sCData();
        break;
      case S_CDATA_ENDING:
        this.sCDataEnding();
        break;
      case S_CDATA_ENDING_2:
        this.sCDataEnding2();
        break;
      case S_PI:
        this.sPI();
        break;
      case S_PI_ENDING:
        this.sPIEnding();
        break;
      case S_DOCTYPE:
        this.sDoctype();
        break;
      case S_DOCTYPE_QUOTE:
        this.sDoctypeQuote();
        break;
      case S_DOCTYPE_DTD:
        this.sDoctypeDTD();
        break;
      case S_DOCTYPE_DTD_QUOTED:
        this.sDoctypeDTDQuoted();
        break;
      case S_ENTITY:
        this.sEntity();
        break;
      case S_DOCTYPE_DTD_LESS:
        this.sDoctypeDTDLess();
        break;
      case S_DOCTYPE_DTD_BANG:
        this.sDoctypeDTDBang();
        break;
      case S_DOCTYPE_DTD_BANG_DASH:
        this.sDoctypeDTDBangDash();
        break;
      case S_DOCTYPE_DTD_DECL:
        this.sDoctypeDTDDecl();
        break;
      case S_DOCTYPE_DTD_COMMENT:
        this.sDoctypeDTDComment();
        break;
      case S_DOCTYPE_DTD_COMMENT_ENDING:
        this.sDoctypeDTDCommentEnding();
        break;
      case S_DOCTYPE_DTD_COMMENT_ENDED:
        this.sDoctypeDTDCommentEnded();
        break;
      case S_DOCTYPE_DTD_PI:
        this.sDoctypeDTDPI();
        break;
      case S_DOCTYPE_DTD_PI_ENDING:
        this.sDoctypeDTDPIEnding();
        break;
      case S_DOCTYPE_DTD_PE_REF:
        this.sDoctypeDTDPERef();
        break;
      case S_DOCTYPE_AFTER_SUBSET:
        this.sDoctypeAfterSubset();
        break;
    }
  }

  // ===========================================================================
  // State Handlers
  // ===========================================================================

  private handleTextInRoot(): void {
    const { chunk } = this;
    let { i: start } = this;
    let handler = this._handlers.text;
    const len = chunk.length;

    // Fast inner loop: scan for special characters directly without getCode() overhead. A handler
    // called on the way — an error's, say — may stop the parse: nothing after that is read.
    while (this.i < len && !this._stopped && !this.halted) {
      // Ultra-fast path: a run of printable ASCII (but `<` and `&`) or tabs, skipped with a local
      // index — `this.i` is stored once per run, not once per character.
      // A line feed joins the run when no line count needs it.
      // A character of the Basic Multilingual Plane past ASCII joins it too, its UTF-8 bytes past the
      // first counted here and added to `extra` once for the run: text in most languages is mostly
      // such characters, and each one stepping out of the run cost a turn of the loop below.
      let at = this.i;
      let code = chunk.charCodeAt(at);
      let extra = 0;
      for (;;) {
        if (code < 0x80) {
          if (
            !(
              (code >= 0x20 && code <= 0x7e && code !== LESS && code !== AMP) ||
              code === TAB ||
              (code === NL && !this.trackPosition)
            )
          ) {
            break;
          }
        } else if (code < 0x800) {
          extra++;
        } else if (code < 0xd800 || (code >= 0xe000 && code < 0xfffe)) {
          extra += 2;
        } else {
          break;
        }
        code = chunk.charCodeAt(++at);
      }
      if (extra !== 0) {
        this.extra += extra;
      }
      // A `>` in the run that ends a literal `]]>`: character data may not hold one (XML 1.0 §2.4).
      // Where they are was found once for the chunk, so the run is scanned as fast as before.
      if (this.cdataEndAt < at) {
        this.checkCDataEnds(at);
        if (this._stopped) {
          // The error's handler stopped the parse there: the text is not read on.
          break;
        }
      }
      if (at !== this.i) {
        if (this.trackPosition) {
          this.column += at - this.i;
        }
        this.i = at;
        if (at >= len) {
          break;
        }
      }

      // '<' — end of text, transition to tag parsing
      if (code === LESS) {
        this.prevI = this.i;
        this.i++;
        if (this.trackPosition) {
          this.column++;
        }
        this.startMarkup(this.prevI);
        if (handler) {
          this.eventStart = this.textStart;
          this.eventEnd = this.markupStart;
          const slice = chunk.slice(start, this.prevI);
          if (this.text.length > 0) {
            handler(this.text + slice);
            this.text = "";
          } else if (slice.length > 0) {
            handler(slice);
          }
        } else {
          // What error recovery appended (an unmatched end tag's text, say) is nobody's: drop it.
          this.text = "";
        }
        this.state = S_OPEN_WAKA;
        if (this._stopped || this.halted || this.closedRoot) {
          return;
        }
        // In content, a start or an end tag is by far the commonest markup: step into it here rather
        // than through another turn of the state machine.
        const next = chunk.charCodeAt(this.i);
        if (next < 128 && NAME_START_CHAR_ASCII[next] === 1) {
          this.prevI = this.i;
          this.i++;
          if (this.trackPosition) {
            this.column++;
          }
          this.state = S_OPEN_TAG;
          this.name = ASCII_CHARS[next];
          this.nameIndex = this.prevI;
          this.sOpenTag();
        } else if (next === FORWARD_SLASH) {
          if (!this.closeInnermost()) {
            this.prevI = this.i;
            this.i++;
            if (this.trackPosition) {
              this.column++;
            }
            this.state = S_CLOSE_TAG;
            this.name = "";
            this.sCloseTag();
          }
        } else if (next === BANG) {
          this.prevI = this.i;
          this.i++;
          if (this.trackPosition) {
            this.column++;
          }
          this.state = S_OPEN_WAKA_BANG;
          this.openWakaBang = "";
          this.sOpenWakaBang();
          if (this.state === S_CDATA) {
            this.sCData();
          } else if (this.state === S_COMMENT) {
            this.sComment();
          }
        } else {
          return;
        }
        // Back in content after a whole tag: go on reading text here. A handler may have changed
        // which handler reads it.
        if (
          this.state !== S_TEXT ||
          this._stopped ||
          this.halted ||
          (this.tags.length === 0 && !this.resume)
        ) {
          return;
        }
        start = this.i;
        handler = this._handlers.text;
        continue;
      }

      // '&' — entity reference
      if (code === AMP) {
        this.prevI = this.i;
        this.i++;
        if (this.trackPosition) {
          this.column++;
        }
        if (handler) {
          this.text += chunk.slice(start, this.prevI);
        }
        this.state = S_ENTITY;
        this.entityReturnState = S_TEXT;
        this.entity = "";
        return;
      }

      // LF — already what a line end becomes, so the text runs on through it
      if (code === NL) {
        this.prevI = this.i;
        this.i++;
        if (this.trackPosition) {
          this.line++;
          this.column = 0;
          this.positionAtNewLine = this.position;
        }
        continue;
      }

      // CR — normalize CR / CR+LF to LF
      if (code === CR) {
        this.prevI = this.i;
        this.i++;
        if (chunk.charCodeAt(this.i) === NL) {
          this.i++;
        }
        if (this.trackPosition) {
          this.line++;
          this.column = 0;
          this.positionAtNewLine = this.position;
        }
        if (handler) {
          this.text += chunk.slice(start, this.prevI) + "\n";
        }
        start = this.i;
        continue;
      }

      // Non-ASCII (>= 0x80) — valid XML, just advance
      if (code >= 0x80) {
        this.prevI = this.i;
        // Handle surrogates
        if (code >= 0xd800 && code <= 0xdbff) {
          const next = chunk.charCodeAt(this.i + 1);
          if (next >= 0xdc00 && next <= 0xdfff) {
            this.i += 2;
            this.extra += 2;
            // Two UTF-16 units: one more column than the one added below.
            if (this.trackPosition) {
              this.column++;
            }
          } else {
            this.i++;
            start = this.handleInvalidCharInText(code, handler, start, "lone surrogate");
          }
        } else if (code >= 0xdc00 && code <= 0xdfff) {
          this.i++;
          start = this.handleInvalidCharInText(code, handler, start, "lone surrogate");
        } else if (code === 0xfffe || code === 0xffff) {
          this.i++;
          this.extra += 2;
          start = this.handleInvalidCharInText(code, handler, start);
        } else {
          this.i++;
          this.extra += code < 0x800 ? 1 : 2;
        }
        if (this.trackPosition) {
          this.column++;
        }
        continue;
      }

      // Invalid ASCII control character (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F)
      this.prevI = this.i;
      this.i++;
      if (this.trackPosition) {
        this.column++;
      }
      start = this.handleInvalidCharInText(code, handler, start);
    }

    // End of chunk
    if (handler && start < this.i) {
      this.text += chunk.slice(start, this.i);
    }
  }

  private handleTextOutsideRoot(): void {
    if (this.beforeDocument && !this.skipBeforeDocument()) {
      return;
    }
    const { chunk } = this;
    let { i: start } = this;
    const handler = this._handlers.text;
    const isSkip = this.invalidCharHandling === "skip";
    const isReplace = this.invalidCharHandling === "replace";

    while (!this._stopped && !this.halted) {
      const iBeforeGet = this.i;
      const c = this.getCode();

      if (c === -1) {
        if (handler && start < iBeforeGet) {
          this.text += chunk.slice(start, iBeforeGet);
        }
        break;
      }

      // In skip mode, getCode() may have internally looped past invalid chars.
      // Flush valid text before the gap and advance start past it.
      if (isSkip && this.prevI > iBeforeGet) {
        if (handler && start < iBeforeGet) {
          this.text += chunk.slice(start, iBeforeGet);
        }
        start = this.prevI;
      }

      // In replace mode, getCode() returns REPLACEMENT_CHAR for invalid chars
      // but the original byte is still in the chunk.  Detect this by checking
      // whether getCode() returned REPLACEMENT_CHAR while the raw chunk byte
      // at prevI is NOT U+FFFD (i.e., it was substituted by handleInvalidChar).
      if (
        isReplace &&
        c === REPLACEMENT_CHAR &&
        chunk.charCodeAt(this.prevI) !== REPLACEMENT_CHAR
      ) {
        if (handler) {
          if (start < this.prevI) {
            this.text += chunk.slice(start, this.prevI);
          }
          this.text += REPLACEMENT_STR;
        }
        start = this.i;
        this.textOutsideRoot();
        continue;
      }

      if (c === LESS) {
        this.startMarkup(this.prevI);
        if (handler) {
          this.eventStart = this.textStart;
          this.eventEnd = this.markupStart;
          const slice = chunk.slice(start, this.prevI);
          if (this.text.length > 0) {
            handler(this.text + slice);
            this.text = "";
          } else if (slice.length > 0) {
            handler(slice);
          }
        } else {
          this.text = "";
        }
        this.state = S_OPEN_WAKA;
        break;
      }

      if (c === AMP) {
        if (handler) {
          this.text += chunk.slice(start, this.prevI);
        }
        this.state = S_ENTITY;
        this.entityReturnState = S_TEXT;
        this.entity = "";
        this.textOutsideRoot();
        break;
      }

      if (c === NL) {
        if (handler) {
          this.text += chunk.slice(start, this.prevI) + "\n";
        }
        start = this.i;
      } else if (!isS(c)) {
        this.textOutsideRoot();
      }
    }
  }

  /**
   * With `skipLeadingText`, before the document: pass over the chunk up to its first `<`, which begins
   * the document — counting what is passed over in positions, reporting none of it — and say whether
   * the document began in this chunk.
   */
  private skipBeforeDocument(): boolean {
    const chunk = this.chunk;
    const from = this.i;
    const at = chunk.indexOf("<", from);
    const end = at === -1 ? chunk.length : at;
    if (end > from) {
      if (!this.chunkOneUnit) {
        this.extra += utf8Length(chunk, from, end) - (end - from);
      }
      if (this.trackPosition) {
        for (let k = from; k < end; k++) {
          if (chunk.charCodeAt(k) === NL) {
            this.line++;
            this.column = 0;
            this.positionAtNewLine = this.chunkPosition + k + 1;
          } else {
            this.column++;
          }
        }
      }
      this.i = this.prevI = end;
    }
    if (at === -1) {
      return false;
    }
    this.beforeDocument = false;
    const start = this.unitAt(at);
    this.skippedLeading = start - this.documentStart;
    this.markupStart = this.textStart = this.documentStart = start;
    return true;
  }

  /**
   * The character just read is text outside the root element: report it there — once before the
   * root, once after — as it is found, so a later error cannot be reported first.
   */
  private textOutsideRoot(): void {
    if (!this.sawRoot && !this.reportedTextBeforeRoot) {
      this.reportedTextBeforeRoot = true;
      this.fail("text data outside of root node.");
    }
    if (this.closedRoot && !this.reportedTextAfterRoot) {
      this.reportedTextAfterRoot = true;
      this.fail("text data outside of root node.");
    }
  }

  private sOpenWaka(): void {
    const c = this.getCode();

    if (c === -1) {
      return;
    }

    if (isNameStartChar(c)) {
      // Reject a second root element in non-fragment mode
      if (this.closedRoot && !this.fragment) {
        this.fail("documents may contain only one root element");
      }
      this.state = S_OPEN_TAG;
      this.name = charFromCode(c);
      this.nameIndex = this.prevI;
      return;
    }

    switch (c) {
      case FORWARD_SLASH:
        this.state = S_CLOSE_TAG;
        this.name = "";
        break;
      case BANG:
        this.state = S_OPEN_WAKA_BANG;
        this.openWakaBang = "";
        break;
      case QUESTION:
        this.state = S_PI;
        this.beginPI(this._handlers.pi !== undefined);
        break;
      default:
        this.fail("unexpected character in tag");
        this.endMarkup();
    }
  }

  private sOpenWakaBang(): void {
    // `[CDATA[` and `--` whole in this chunk are taken at once rather than a character at a time.
    if (this.openWakaBang === "") {
      const at = this.i;
      const word = this.chunk.startsWith("[CDATA[", at)
        ? "[CDATA["
        : this.chunk.startsWith("--", at)
          ? "--"
          : undefined;
      if (word !== undefined) {
        this.prevI = at + word.length - 1;
        this.i = at + word.length;
        if (this.trackPosition) {
          this.column += word.length;
        }
        this.openWakaBang = word;
        this.enterBang();
        return;
      }
    }
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    this.openWakaBang += charFromCode(c);
    this.enterBang();
  }

  /**
   * Add to `text` the run of printable ASCII and tabs from the current position up to `stop` — the
   * character that may end the construct — in one slice; anything else is left to `getCode`.
   */
  private skipRun(stop: number): void {
    const chunk = this.chunk;
    const from = this.i;
    let at = from;
    let code = chunk.charCodeAt(at);
    while (((code >= 0x20 && code <= 0x7e) || code === TAB) && code !== stop) {
      code = chunk.charCodeAt(++at);
    }
    if (at !== from) {
      if (this.keepText) {
        this.text += chunk.slice(from, at);
      }
      if (this.trackPosition) {
        this.column += at - from;
      }
      this.i = at;
    }
  }

  /** Step over the `length` ASCII characters that end a construct. */
  private skipEnd(length: number): void {
    this.prevI = this.i + length - 1;
    this.i += length;
    if (this.trackPosition) {
      this.column += length;
    }
  }

  /** Act on what follows `<!` so far: a CDATA section, a comment, a DOCTYPE, or more to read. */
  private enterBang(): void {
    switch (this.openWakaBang) {
      case "[CDATA[":
        if (this.tags.length === 0 && !this.fragment && !this.resume) {
          this.fail("CDATA section outside of root node");
        }
        this.state = S_CDATA;
        this.text = "";
        this.keepText = this._handlers.cdata !== undefined || this._handlers.text !== undefined;
        this.openWakaBang = "";
        break;
      case "--":
        this.state = S_COMMENT;
        this.text = "";
        this.keepText = this._handlers.comment !== undefined;
        this.openWakaBang = "";
        break;
      case "DOCTYPE":
        if (this.sawDoctype) {
          this.fail("a document may contain only one DOCTYPE declaration");
        } else if (this.tags.length > 0 || (this.sawRoot && !this.fragment)) {
          this.fail("the DOCTYPE declaration must come before the root element");
        }
        this.sawDoctype = true;
        this.state = S_DOCTYPE;
        this.text = "";
        this.openWakaBang = "";
        break;
      default:
        // Refused as soon as it can no longer become one of the three, so the characters after it — a `<` among
        // them — are read as what they are rather than swallowed into a seven-character guess.
        if (
          !"[CDATA[".startsWith(this.openWakaBang) &&
          !"--".startsWith(this.openWakaBang) &&
          !"DOCTYPE".startsWith(this.openWakaBang)
        ) {
          this.fail("incorrect syntax");
          this.openWakaBang = "";
          this.endMarkup();
        }
    }
  }

  private sOpenTag(): void {
    // Fast path: scan the tag name as a contiguous slice instead of char-by-char concat
    // Note: first char was already consumed by sOpenWaka and stored in this.name
    const { chunk } = this;
    const nameStart = this.i; // start after the first char already in this.name
    let pos = this.i;

    // Fast ASCII scan
    pos = this.scanNameChars(pos);

    if (pos >= chunk.length) {
      // Chunk ended in the middle of tag name
      this.finishName(nameStart, pos);
      this.i = pos;
      if (this.trackPosition) {
        this.column += pos - nameStart;
      }
      return;
    }

    this.finishName(nameStart, pos);
    this.i = pos;
    if (this.trackPosition) {
      this.column += pos - nameStart;
    }

    // Now read the delimiter character: `>`, the common one, without a call to getCode.
    let c: number;
    if (chunk.charCodeAt(pos) === GREATER) {
      this.prevI = pos;
      this.i = pos + 1;
      if (this.trackPosition) {
        this.column++;
      }
      c = GREATER;
    } else {
      c = this.getCode();
      if (c === -1) {
        return;
      }
    }

    this.tag = {
      name: this.name,
      attributes:
        this.keyPrefix === undefined
          ? (Object.create(null) as Record<string, string>)
          : NO_ATTRIBUTES,
      isSelfClosing: false
    };
    this.duplicateAttributes = null;
    this.sawRoot = true;

    if (c === GREATER) {
      this.openTag();
    } else if (c === FORWARD_SLASH) {
      this.state = S_OPEN_TAG_SLASH;
    } else if (isS(c)) {
      this.scanAttributes();
    } else {
      this.fail("unexpected character in tag");
      this.state = S_ATTRIB;
    }
  }

  /**
   * Read a start tag's ordinary attributes, and its `>` or `/>`, in one pass.
   *
   * The state machine takes one dispatch per step of an attribute — name, `=`, value, after-value — which on a
   * worksheet is most of the parse: `<c r="A1" s="1" t="s">` is twelve of them, for every cell. This consumes the
   * common shape directly: a name of ASCII name characters, `=`, a quoted value of printable ASCII with no `&` or `<`,
   * then whitespace or the end of the tag.
   *
   * Anything else is handed back to the state machine **at the start of the attribute in question**, never part-way
   * through it — an entity, a control or non-ASCII character, a missing `=` or quote, a chunk boundary, a newline when
   * positions are tracked. Every attribute before that point is complete and recorded exactly as the state machine
   * would have recorded it, so the result, the errors and their positions are the state machine's own.
   */
  private scanAttributes(): void {
    const { chunk, trackPosition } = this;
    const len = chunk.length;
    const start = this.i;
    let pos = start;

    for (;;) {
      while (pos < len) {
        const c = chunk.charCodeAt(pos);
        // A line break moves the line counter, which is the state machine's job when positions are tracked.
        if (c === SPACE || c === TAB || (!trackPosition && (c === NL || c === CR))) {
          pos++;
        } else {
          break;
        }
      }
      if (pos >= len) {
        break;
      }
      let c = chunk.charCodeAt(pos);
      if (c === GREATER) {
        this.finishScan(start, pos + 1);
        this.openTag();
        return;
      }
      if (c === FORWARD_SLASH) {
        if (pos + 1 < len && chunk.charCodeAt(pos + 1) === GREATER) {
          this.finishScan(start, pos + 2);
          this.openSelfClosingTag();
          return;
        }
        break;
      }
      if (c >= 128 || NAME_START_CHAR_ASCII[c] !== 1) {
        break;
      }

      const nameStart = pos;
      pos++;
      while (pos < len) {
        c = chunk.charCodeAt(pos);
        if (c < 128 && NAME_CHAR_ASCII[c] === 1) {
          pos++;
        } else {
          break;
        }
      }
      const quote =
        pos + 1 < len && chunk.charCodeAt(pos) === EQUAL ? chunk.charCodeAt(pos + 1) : -1;
      if (quote !== DQUOTE && quote !== SQUOTE) {
        pos = nameStart;
        break;
      }
      const valueStart = pos + 2;
      let end = valueStart;
      while (end < len) {
        c = chunk.charCodeAt(end);
        if (c >= 0x20 && c <= 0x7e && c !== quote && c !== AMP && c !== LESS) {
          end++;
        } else {
          break;
        }
      }
      if (end >= len || c !== quote) {
        pos = nameStart;
        break;
      }
      this.addAttribute(this.nameAt(nameStart, pos), chunk.slice(valueStart, end));
      pos = end + 1;

      // After a value only whitespace or the end of the tag may follow; anything else, and the end of the chunk, is
      // the after-value state's to judge — it is what reports "no whitespace between attributes".
      c = pos < len ? chunk.charCodeAt(pos) : -1;
      if (
        c !== SPACE &&
        c !== TAB &&
        c !== NL &&
        c !== CR &&
        c !== GREATER &&
        c !== FORWARD_SLASH
      ) {
        this.finishScan(start, pos);
        this.state = S_ATTRIB_VALUE_CLOSED;
        return;
      }
    }
    this.finishScan(start, pos);
    this.state = S_ATTRIB;
  }

  /** Move past what `scanAttributes` consumed. It never consumes a line break while positions are tracked. */
  private finishScan(start: number, end: number): void {
    this.i = end;
    if (end > start) {
      // The last character read, as the state machine leaves it: an error reported next is placed there.
      this.prevI = end - 1;
    }
    if (this.trackPosition) {
      this.column += end - start;
    }
  }

  private sOpenTagSlash(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      this.openSelfClosingTag();
    } else {
      this.fail("expected >");
      this.state = S_ATTRIB;
    }
  }

  private sAttrib(): void {
    const c = this.skipSpaces();
    if (c === -1) {
      return;
    }

    if (isNameStartChar(c)) {
      this.name = charFromCode(c);
      this.nameIndex = this.prevI;
      this.state = S_ATTRIB_NAME;
    } else if (c === GREATER) {
      this.openTag();
    } else if (c === FORWARD_SLASH) {
      this.state = S_OPEN_TAG_SLASH;
    } else {
      this.fail("unexpected character in attribute");
    }
  }

  private sAttribName(): void {
    // Fast path: scan attribute name as a contiguous slice
    // Note: first char was already consumed by sAttrib and stored in this.name
    const { chunk } = this;
    const nameStart = this.i;
    let pos = this.i;

    pos = this.scanNameChars(pos);

    if (pos >= chunk.length) {
      this.finishName(nameStart, pos);
      this.i = pos;
      if (this.trackPosition) {
        this.column += pos - nameStart;
      }
      return;
    }

    this.finishName(nameStart, pos);
    this.i = pos;
    if (this.trackPosition) {
      this.column += pos - nameStart;
    }

    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === EQUAL) {
      this.state = S_ATTRIB_VALUE;
    } else if (isS(c)) {
      this.state = S_ATTRIB_NAME_SAW_WHITE;
    } else if (c === GREATER) {
      this.fail("attribute without value");
      this.addAttribute(this.name, this.name);
      this.name = "";
      this.openTag();
    } else {
      this.fail("unexpected character in attribute name");
    }
  }

  private sAttribNameSawWhite(): void {
    const c = this.skipSpaces();
    if (c === -1) {
      return;
    }

    if (c === EQUAL) {
      this.state = S_ATTRIB_VALUE;
    } else {
      this.fail("attribute without value");
      this.name = "";
      this.text = "";
      if (c === GREATER) {
        this.openTag();
      } else if (isNameStartChar(c)) {
        this.name = charFromCode(c);
        this.nameIndex = this.prevI;
        this.state = S_ATTRIB_NAME;
      } else {
        this.fail("unexpected character");
        this.state = S_ATTRIB;
      }
    }
  }

  private sAttribValue(): void {
    const c = this.skipSpaces();
    if (c === -1) {
      return;
    }

    if (isQuote(c)) {
      this.q = c;
      this.text = "";
      this.state = S_ATTRIB_VALUE_QUOTED;
    } else {
      this.fail("unquoted attribute value");
      this.endMarkup();
    }
  }

  private sAttribValueQuoted(): void {
    const { q, chunk } = this;
    let { i: start } = this;
    const len = chunk.length;

    // Fast inner loop: most attribute values are short printable ASCII
    while (this.i < len) {
      const code = chunk.charCodeAt(this.i);

      // Ultra-fast path: printable ASCII excluding quote, <, &
      if (code >= 0x20 && code <= 0x7e && code !== q && code !== AMP && code !== LESS) {
        this.i++;
        if (this.trackPosition) {
          this.column++;
        }
        continue;
      }

      // Closing quote
      if (code === q) {
        this.prevI = this.i;
        this.i++;
        if (this.trackPosition) {
          this.column++;
        }
        this.addAttribute(this.name, this.text + chunk.slice(start, this.prevI));
        this.name = "";
        this.text = "";
        this.q = null;
        this.state = S_ATTRIB_VALUE_CLOSED;
        return;
      }

      // Entity reference
      if (code === AMP) {
        this.prevI = this.i;
        this.i++;
        if (this.trackPosition) {
          this.column++;
        }
        this.text += chunk.slice(start, this.prevI);
        this.state = S_ENTITY;
        this.entityReturnState = S_ATTRIB_VALUE_QUOTED;
        this.entity = "";
        return;
      }

      // < not allowed in attribute value
      if (code === LESS) {
        this.prevI = this.i;
        this.i++;
        if (this.trackPosition) {
          this.column++;
        }
        this.text += chunk.slice(start, this.prevI);
        this.fail("< not allowed in attribute value");
        this.addAttribute(this.name, this.text);
        this.name = "";
        this.text = "";
        this.q = null;
        this.unget();
        this.endMarkup();
        return;
      }

      // TAB or LF in attribute value — normalize to space per XML 1.0
      if (code === TAB || code === NL) {
        this.prevI = this.i;
        this.i++;
        if (code === NL && this.trackPosition) {
          this.line++;
          this.column = 0;
          this.positionAtNewLine = this.position;
        } else if (this.trackPosition) {
          this.column++;
        }
        this.text += chunk.slice(start, this.prevI) + " ";
        start = this.i;
        continue;
      }

      // CR — normalize to space (attribute normalization)
      if (code === CR) {
        this.prevI = this.i;
        this.i++;
        if (chunk.charCodeAt(this.i) === NL) {
          this.i++;
        }
        if (this.trackPosition) {
          this.line++;
          this.column = 0;
          this.positionAtNewLine = this.position;
        }
        this.text += chunk.slice(start, this.prevI) + " ";
        start = this.i;
        continue;
      }

      // Non-ASCII (>= 0x80) — mostly valid, handle inline like handleTextInRoot
      if (code >= 0x80) {
        this.prevI = this.i;
        if (code >= 0xd800 && code <= 0xdbff) {
          const next = chunk.charCodeAt(this.i + 1);
          if (next >= 0xdc00 && next <= 0xdfff) {
            this.i += 2; // valid surrogate pair
            this.extra += 2;
            // Two UTF-16 units: one more column than the one added below.
            if (this.trackPosition) {
              this.column++;
            }
          } else {
            this.i++;
            start = this.handleInvalidCharInAttr(code, start, "lone surrogate");
          }
        } else if (code >= 0xdc00 && code <= 0xdfff) {
          this.i++;
          start = this.handleInvalidCharInAttr(code, start, "lone surrogate");
        } else if (code === 0xfffe || code === 0xffff) {
          this.i++;
          this.extra += 2;
          start = this.handleInvalidCharInAttr(code, start);
        } else {
          this.i++; // valid non-ASCII BMP char
          this.extra += code < 0x800 ? 1 : 2;
        }
        if (this.trackPosition) {
          this.column++;
        }
        if (this._stopped) {
          return;
        }
        continue;
      }

      // Remaining: ASCII control characters (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F)
      // All invalid in XML 1.0
      this.prevI = this.i;
      this.i++;
      if (this.trackPosition) {
        this.column++;
      }
      start = this.handleInvalidCharInAttr(code, start);
      if (this._stopped) {
        // The error's handler stopped the parse: nothing more of the value is read.
        return;
      }
    }

    // End of chunk
    this.text += chunk.slice(start, this.i);
  }

  private sAttribValueClosed(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (isS(c)) {
      this.state = S_ATTRIB;
    } else if (c === GREATER) {
      this.openTag();
    } else if (c === FORWARD_SLASH) {
      this.state = S_OPEN_TAG_SLASH;
    } else if (isNameStartChar(c)) {
      this.fail("no whitespace between attributes");
      this.name = charFromCode(c);
      this.nameIndex = this.prevI;
      this.state = S_ATTRIB_NAME;
    } else {
      this.fail("unexpected character after attribute");
    }
  }

  private sCloseTag(): void {
    const { chunk } = this;
    let pos = this.i;

    if (this.name === "") {
      // First character must be NameStartChar — read it
      if (pos >= chunk.length) {
        const c = this.getCode();
        if (c === -1) {
          return;
        }
        if (isNameStartChar(c)) {
          this.name = charFromCode(c);
          this.nameIndex = this.prevI;
        } else {
          this.fail("unexpected character in close tag");
        }
        return;
      }
      const first = chunk.charCodeAt(pos);
      if (first >= 128) {
        // Read through getCode, which counts its bytes and reads a pair as one character.
        const c = this.getCode();
        if (!isNameStartChar(c)) {
          this.fail("unexpected character in close tag");
          return;
        }
        this.name = charFromCode(c);
        this.nameIndex = this.prevI;
        pos = this.i;
      } else if (NAME_START_CHAR_ASCII[first] !== 1) {
        this.i = pos;
        const c = this.getCode();
        if (c !== -1) {
          this.fail("unexpected character in close tag");
        }
        return;
      } else {
        // Consume the first character
        this.name = chunk[pos];
        this.nameIndex = pos;
        pos++;
        if (this.trackPosition) {
          this.column++;
        }
      }
    }

    // Fast scan remaining name chars
    const nameStart = pos;
    pos = this.scanNameChars(pos);

    // An end tag nearly always closes the innermost element: take its name rather than make another.
    const tags = this.tags;
    const top = tags.length > 0 ? tags[tags.length - 1].name : undefined;
    if (
      this.nameIndex >= 0 &&
      top !== undefined &&
      top.length === pos - this.nameIndex &&
      this.chunk.startsWith(top, this.nameIndex)
    ) {
      this.name = top;
      this.nameIndex = -1;
    } else {
      this.finishName(nameStart, pos);
    }
    if (this.trackPosition) {
      this.column += pos - nameStart;
    }

    if (pos >= chunk.length) {
      this.i = pos;
      return;
    }

    if (chunk.charCodeAt(pos) === GREATER) {
      // The common end, read without a call to getCode.
      this.prevI = pos;
      this.i = pos + 1;
      if (this.trackPosition) {
        this.column++;
      }
      this.closeTag();
      return;
    }
    this.i = pos;
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === GREATER) {
      this.closeTag();
    } else if (isS(c)) {
      this.state = S_CLOSE_TAG_SAW_WHITE;
    } else {
      this.fail("unexpected character in close tag");
    }
  }

  private sCloseTagSawWhite(): void {
    const c = this.skipSpaces();
    if (c === -1) {
      return;
    }

    if (c === GREATER) {
      this.closeTag();
    } else {
      this.fail("unexpected character in close tag");
    }
  }

  private sComment(): void {
    this.skipRun(MINUS);
    if (this.i >= this.chunk.length) {
      return;
    }
    if (this.chunk.startsWith("-->", this.i)) {
      this.skipEnd(3);
      this.endComment();
      return;
    }
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === MINUS) {
      this.state = S_COMMENT_ENDING;
    } else {
      if (this.keepText) {
        this.text += charFromCode(c);
      }
    }
  }

  private sCommentEnding(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === MINUS) {
      this.state = S_COMMENT_ENDED;
    } else {
      if (this.keepText) {
        this.text += "-" + charFromCode(c);
      }
      this.state = S_COMMENT;
    }
  }

  /** The `>` of `-->` has been read: report the comment. */
  private endComment(): void {
    this.markupEvent();
    this._handlers.comment?.(this.text);
    this.text = "";
    this.state = S_TEXT;
  }

  private sCommentEnded(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      this.endComment();
    } else if (c === MINUS) {
      // `--->`: XML forbids `--` in a comment, and a comment ending in `-` puts one before the `-->`.
      this.fail("malformed comment");
      if (this.keepText) {
        this.text += "-";
      }
    } else {
      this.fail("malformed comment");
      if (this.keepText) {
        this.text += "--" + charFromCode(c);
      }
      this.state = S_COMMENT;
    }
  }

  private sCData(): void {
    this.skipRun(CLOSE_BRACKET);
    if (this.i >= this.chunk.length) {
      return;
    }
    if (this.chunk.startsWith("]]>", this.i)) {
      this.skipEnd(3);
      this.endCData();
      return;
    }
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === CLOSE_BRACKET) {
      this.state = S_CDATA_ENDING;
    } else {
      if (this.keepText) {
        this.text += charFromCode(c);
      }
    }
  }

  private sCDataEnding(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === CLOSE_BRACKET) {
      this.state = S_CDATA_ENDING_2;
    } else {
      if (this.keepText) {
        this.text += "]" + charFromCode(c);
      }
      this.state = S_CDATA;
    }
  }

  private sCDataEnding2(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      this.endCData();
    } else if (c === CLOSE_BRACKET) {
      if (this.keepText) {
        this.text += "]";
      }
    } else {
      if (this.keepText) {
        this.text += "]]" + charFromCode(c);
      }
      this.state = S_CDATA;
    }
  }

  /** The `>` of `]]>` has been read: report the section. */
  private endCData(): void {
    // Emit CDATA event; also emit as text for backwards compat
    this.markupEvent();
    if (this._handlers.cdata) {
      this._handlers.cdata(this.text);
    } else if (this._handlers.text && this.text.length > 0) {
      this._handlers.text(this.text);
    }
    this.text = "";
    this.state = S_TEXT;
  }

  private sPI(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === QUESTION) {
      this.state = S_PI_ENDING;
    } else {
      this.appendPI(charFromCode(c));
    }
  }

  private sPIEnding(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      // The target runs to the first whitespace of any kind (XML 1.0 §2.6), and the body starts after all of it.
      const piText = this.text;
      const len = piText.length;
      let end = 0;
      while (end < len && !isS(piText.charCodeAt(end))) {
        end++;
      }
      let bodyStart = end;
      while (bodyStart < len && isS(piText.charCodeAt(bodyStart))) {
        bodyStart++;
      }
      const target = end === len ? piText : piText.slice(0, end);
      if (!isXmlName(target)) {
        this.fail(
          target === ""
            ? "processing instruction without a target"
            : `invalid processing instruction target: ${target}`
        );
      }
      const body = bodyStart === len ? "" : piText.slice(bodyStart);
      if (target.length === 3 && target.toLowerCase() === "xml") {
        // `xml` in any case is reserved (§2.6): only the XML declaration, at the very start, uses it.
        if (target !== "xml") {
          this.fail(`processing instruction target ${target} is reserved`);
        } else if (this.resume || this.markupStart !== this.documentStart) {
          this.fail("the XML declaration may only appear at the start of the document");
        } else {
          const problem = checkXmlDeclaration(body);
          if (problem !== undefined) {
            this.fail(problem);
          } else {
            this.standalone = /\bstandalone\s*=\s*(["'])yes\1/.test(body);
            if (this.bytes === true) {
              this.checkDeclaredEncoding(body);
            }
          }
        }
      }
      this.text = "";
      this.state = S_TEXT;
      if (this._stopped) {
        // An error's handler stopped the parse: the instruction is not reported.
        return;
      }
      this.markupEvent();
      this._handlers.pi?.(target, body);
    } else if (c === QUESTION) {
      this.appendPI("?");
    } else {
      this.appendPI("?" + charFromCode(c));
      this.state = S_PI;
    }
  }

  /**
   * Add to the text of the processing instruction being read. Its target is always kept, since it
   * is checked whatever handlers there are; its body only when a handler is given it, or when it is
   * an XML declaration, whose encoding is checked. So whether anyone listens never changes what is
   * accepted, and a long instruction nobody reads is not held.
   */
  private appendPI(text: string): void {
    if (this.piTargetEnd >= 0 && !this.keepText && !this.piKeepBody) {
      return;
    }
    const from = this.text.length;
    this.text += text;
    if (this.piTargetEnd < 0) {
      for (let k = 0; k < text.length; k++) {
        if (isS(text.charCodeAt(k))) {
          this.piTargetEnd = from + k;
          this.piKeepBody = this.text.slice(0, this.piTargetEnd) === "xml";
          break;
        }
      }
    }
  }

  /** Start reading a processing instruction; `keep` says whether its body is given to anyone. */
  private beginPI(keep: boolean): void {
    this.text = "";
    this.keepText = keep;
    this.piTargetEnd = -1;
    this.piKeepBody = false;
  }

  // ---------------------------------------------------------------------------
  // DOCTYPE
  //
  // The head (`<!DOCTYPE name SYSTEM "…"`) is gathered in `text` and checked whole. The internal
  // subset is scanned as markup — a `]`, `>` or quote inside a comment, a processing instruction
  // or a declaration's literal does not end it — and each declaration is gathered in
  // `declaration` and checked whole at its `>`. Nothing is fetched.
  // ---------------------------------------------------------------------------

  private sDoctype(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === GREATER) {
      this.finishDoctypeHead();
      this.text = "";
      this.endMarkup();
    } else if (c === OPEN_BRACKET) {
      this.finishDoctypeHead();
      this.text = "";
      this.sawInternalSubset = true;
      this.state = S_DOCTYPE_DTD;
    } else {
      if (isQuote(c)) {
        this.q = c;
        this.state = S_DOCTYPE_QUOTE;
      }
      this.appendDoctypeHead(c);
    }
  }

  private sDoctypeQuote(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === this.q) {
      this.q = null;
      this.state = S_DOCTYPE;
    }
    this.appendDoctypeHead(c);
  }

  private appendDoctypeHead(c: number): void {
    if (this.text.length < MAX_DOCTYPE_HEAD) {
      this.text += charFromCode(c);
    } else if (this.text.length === MAX_DOCTYPE_HEAD) {
      this.fail("the DOCTYPE declaration's head is too long");
      this.text += " ";
    }
  }

  private finishDoctypeHead(): void {
    const problem = checkDoctypeHead(this.text);
    if (problem !== undefined) {
      this.fail(problem);
    }
    // `S Name S ExternalID`: a keyword after the name, then a literal.
    this.externalSubset = /^[ \t\r\n]+\S+[ \t\r\n]+(?:SYSTEM|PUBLIC)[ \t\r\n]/.test(this.text);
  }

  /** Between declarations: white space, a parameter-entity reference, markup, or the closing `]`. */
  private sDoctypeDTD(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    this.doctypeSubsetChar(c);
  }

  private doctypeSubsetChar(c: number): void {
    if (c === LESS) {
      this.state = S_DOCTYPE_DTD_LESS;
    } else if (c === CLOSE_BRACKET) {
      this.state = S_DOCTYPE_AFTER_SUBSET;
    } else if (c === PERCENT) {
      this.entity = "";
      this.state = S_DOCTYPE_DTD_PE_REF;
    } else if (!isS(c)) {
      this.fail("unexpected character in the DOCTYPE internal subset");
      this.state = S_DOCTYPE_DTD;
    } else {
      this.state = S_DOCTYPE_DTD;
    }
  }

  /**
   * Whether an entity may be undeclared without that being an error: an external subset or a
   * parameter-entity reference, neither read, may declare it — unless the document says it is
   * standalone, when every entity it uses must be declared in what is read (WFC: Entity Declared).
   */
  private mayBeDeclaredElsewhere(): boolean {
    return !this.standalone && (this.externalSubset || this.sawParameterEntityReference);
  }

  /**
   * The entities an attribute list's defaults refer to, and those they refer to in turn: each must
   * be declared before it (WFC: Entity Declared), not external (WFC: No External Entity References),
   * not refer to itself (WFC: No Recursion), and its replacement text may hold no `<` (WFC: No < in
   * Attribute Values) — the rules a default is held to when it is declared, not only when it is used.
   */
  private checkDefaultReferences(references: readonly string[]): void {
    const bindings = this.entityBindings;
    const sound = (this.soundDefaultEntities ??= new Set());
    for (const reference of references) {
      if (sound.has(reference)) {
        continue;
      }
      // Walked with a stack of its own: a long chain of entities is refused or passed, never a
      // RangeError. What was found sound is not walked again, so a default refers in linear time.
      const stack: Array<{ name: string; refs: readonly string[]; next: number }> = [];
      const active = new Set<string>();
      const enter = (name: string): boolean => {
        const binding = bindings?.get(name);
        if (binding === undefined) {
          if (!this.mayBeDeclaredElsewhere()) {
            this.fail(`an attribute default refers to the undeclared entity ${name}`);
          }
          return false;
        }
        if (binding.kind !== "internal") {
          this.fail(`an attribute default refers to the external entity ${name}`);
          return false;
        }
        if (binding.summary.lt) {
          this.fail(`an attribute default refers to the entity ${name}, which puts a < in it`);
          return false;
        }
        if (binding.summary.malformed) {
          this.fail(
            `an attribute default refers to the entity ${name}, which holds a malformed reference`
          );
          return false;
        }
        active.add(name);
        stack.push({ name, refs: binding.summary.refs, next: 0 });
        return true;
      };
      if (!enter(reference)) {
        continue;
      }
      let failed = false;
      while (stack.length > 0 && !failed) {
        const top = stack[stack.length - 1];
        if (top.next === top.refs.length) {
          stack.pop();
          active.delete(top.name);
          sound.add(top.name);
          continue;
        }
        const ref = top.refs[top.next++];
        if (sound.has(ref)) {
          continue;
        }
        if (active.has(ref)) {
          this.fail(`an attribute default refers to the entity ${ref}, which refers to itself`);
          failed = true;
        } else if (!enter(ref)) {
          failed = true;
        }
      }
    }
  }

  /** A parameter-entity reference: `%name;`. */
  private sDoctypeDTDPERef(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === SEMICOLON) {
      if (!isXmlName(this.entity)) {
        this.fail("malformed parameter-entity reference");
      } else if (this.standalone && this.entityBindings?.has("%" + this.entity) !== true) {
        // A standalone document declares every entity it uses before it uses it (WFC: Entity Declared).
        this.fail(`undeclared parameter entity: %${this.entity};`);
      }
      // It is not read, and it may hold declarations that override later ones (XML 1.0 §5.1), so
      // no entity declared after it is processed.
      this.sawParameterEntityReference = true;
      this.entity = "";
      this.state = S_DOCTYPE_DTD;
    } else if (
      this.entity.length < 256 &&
      (isNameChar(c) || (this.entity === "" && isNameStartChar(c)))
    ) {
      this.entity += charFromCode(c);
    } else {
      this.fail("malformed parameter-entity reference");
      this.entity = "";
      this.state = S_DOCTYPE_DTD;
    }
  }

  private sDoctypeDTDLess(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === BANG) {
      this.declaration = "";
      this.state = S_DOCTYPE_DTD_BANG;
    } else if (c === QUESTION) {
      // A processing instruction is the application's wherever it is (XML 1.0 §2.6), the internal
      // subset included: its body is kept only when a handler is given it.
      this.beginPI(this._handlers.pi !== undefined);
      this.subsetPIStart = this.unitAt(this.prevI) - 1;
      this.state = S_DOCTYPE_DTD_PI;
    } else {
      this.fail("unexpected character in the DOCTYPE internal subset");
      this.doctypeSubsetChar(c);
    }
  }

  private sDoctypeDTDBang(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === MINUS) {
      this.state = S_DOCTYPE_DTD_BANG_DASH;
    } else {
      this.doctypeDeclarationChar(c);
    }
  }

  private sDoctypeDTDBangDash(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === MINUS) {
      this.state = S_DOCTYPE_DTD_COMMENT;
    } else {
      this.declaration = "-";
      this.doctypeDeclarationChar(c);
    }
  }

  /** Inside a markup declaration (`<!ENTITY …>`, `<!ATTLIST …>`, …): only a literal can hold its `>`. */
  private sDoctypeDTDDecl(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    this.doctypeDeclarationChar(c);
  }

  private doctypeDeclarationChar(c: number): void {
    if (c === GREATER) {
      this.finishDeclaration();
      this.state = S_DOCTYPE_DTD;
      return;
    }
    if (isQuote(c)) {
      this.q = c;
      this.state = S_DOCTYPE_DTD_QUOTED;
    } else {
      this.state = S_DOCTYPE_DTD_DECL;
    }
    this.appendDeclaration(c);
  }

  private sDoctypeDTDQuoted(): void {
    // A literal is most of a long declaration — an entity's whole value. A run of printable ASCII
    // other than the quote is taken as one slice rather than a character at a time, so long as it
    // fits within the bound; anything else goes through `getCode`, which checks and counts it.
    const { chunk } = this;
    const quote = this.q;
    const room = this.declarationOverflow ? 0 : this.maxContextBytes - this.declaration.length;
    let at = this.i;
    const stop = Math.min(chunk.length, at + room);
    while (at < stop) {
      const code = chunk.charCodeAt(at);
      if (code < 0x20 || code > 0x7e || code === quote) {
        break;
      }
      at++;
    }
    if (at > this.i) {
      this.declaration += chunk.slice(this.i, at);
      if (this.trackPosition) {
        this.column += at - this.i;
      }
      this.i = at;
    }
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === this.q) {
      this.q = null;
      this.state = S_DOCTYPE_DTD_DECL;
    }
    this.appendDeclaration(c);
  }

  /**
   * Gather a declaration's text to check it at its `>`, up to `maxContextBytes` characters. One that
   * is longer is refused as past that limit, whatever it declares: it cannot be checked without being
   * held, and taking it unchecked would accept a malformed declaration because it was long. A parse
   * whose input is all in memory (`Xml.parse`, `Xml.parseToObject`) sets no such bound; a streamed
   * one can raise it with `maxContextBytes`.
   */
  private appendDeclaration(c: number): void {
    if (this.declarationOverflow) {
      return;
    }
    if (this.declaration.length < this.maxContextBytes) {
      this.declaration += charFromCode(c);
      return;
    }
    this.declarationOverflow = true;
    const problem = checkDeclarationStart(this.declaration);
    if (problem !== undefined) {
      this.fail(problem);
    } else {
      this.fail(
        `a declaration in the DOCTYPE internal subset is longer than maxContextBytes (${this.maxContextBytes}), and cannot be checked`,
        "prologContext"
      );
    }
  }

  private finishDeclaration(): void {
    if (this.declarationOverflow) {
      this.declarationOverflow = false;
      this.declaration = "";
      return;
    }
    const read = readMarkupDeclaration(this.declaration);
    this.declaration = "";
    if (typeof read === "string") {
      this.fail(read);
      return;
    }
    // After a parameter-entity reference, which is not read and may declare what follows otherwise,
    // no declaration is processed (XML 1.0 §5.1) — unless the document is standalone. A skipped
    // attribute list's defaults are not looked into either.
    const skipped = this.sawParameterEntityReference && !this.standalone;
    if (read.attlist !== undefined && !skipped) {
      if (read.references !== undefined) {
        this.checkDefaultReferences(read.references);
      }
      this.declareAttributes(read.attlist);
    }
    const entity = read.entity;
    if (entity === undefined) {
      return;
    }
    const key = entity.kind === "parameter" ? "%" + entity.name : entity.name;
    const bindings = (this.entityBindings ??= new Map());
    if (skipped || bindings.has(key)) {
      // The first declaration of a name binds it (XML 1.0 §4.2).
      return;
    }
    if (!this.keepDeclared(key.length)) {
      return;
    }
    bindings.set(
      key,
      entity.kind === "internal"
        ? { kind: "internal", summary: summarizeEntity(entity.value) }
        : { kind: entity.kind }
    );
    if (entity.kind === "parameter" || isPredefinedEntity(entity.name)) {
      return;
    }
    if (entity.kind === "external") {
      (this.externalEntities ??= new Set()).add(entity.name);
      return;
    }
    if (!this.internalEntities) {
      return;
    }
    this.contextSize.add(entity.name, entity.value);
    if (this.contextSize.total() > this.maxContextBytes) {
      this.fail(
        `the entities the DOCTYPE internal subset declares take more than maxContextBytes (${this.maxContextBytes})`,
        "prologContext"
      );
      return;
    }
    this.declared[entity.name] = entity.value;
  }

  /**
   * Count `length` more characters of declarations kept, with a fixed share for the entry holding
   * them; whether that is still within `maxContextBytes`. Past it, the limit is reported.
   */
  private keepDeclared(length: number): boolean {
    this.keptDeclarations += length + KEPT_DECLARATION_OVERHEAD;
    if (this.keptDeclarations <= this.maxContextBytes) {
      return true;
    }
    this.fail(
      `the declarations the DOCTYPE internal subset makes take more than maxContextBytes (${this.maxContextBytes}) to keep`,
      "prologContext"
    );
    return false;
  }

  /**
   * Keep what an `<!ATTLIST>` says of each attribute it is the first to declare: a non-validating
   * parser uses the declarations it reads to give defaults and to normalize values (XML 1.0 §5.1).
   */
  private declareAttributes(attlist: AttlistDeclaration): void {
    const { element } = attlist;
    const bound = (this.boundAttributes ??= new Set());
    for (const definition of attlist.attributes) {
      const key = element + "\0" + definition.name;
      if (bound.has(key)) {
        continue;
      }
      bound.add(key);
      if (!this.keepDeclared(key.length)) {
        return;
      }
      if (!definition.tokenized && definition.value === undefined) {
        continue;
      }
      let value: string | undefined;
      if (definition.value !== undefined) {
        // Expanding the default is charged here, once: an element given it shares the one string.
        const normalized = this.normalizeDefault(definition.value);
        if (this._stopped) {
          return;
        }
        if (typeof normalized === "string") {
          value = definition.tokenized ? collapseSpaces(normalized) : normalized;
        } else if (normalized !== undefined) {
          (this.unresolvedDefaults ??= new Map()).set(key, normalized.unresolved);
        }
      }
      if (!definition.tokenized && value === undefined) {
        continue;
      }
      const declaration: XmlAttributeDeclaration = {
        ...(definition.tokenized && { tokenized: true as const }),
        ...(value !== undefined && { default: value })
      };
      const rules = (this.attributeRules ??= Object.create(null) as Record<
        string,
        Record<string, XmlAttributeDeclaration>
      >);
      (rules[element] ??= Object.create(null) as Record<string, XmlAttributeDeclaration>)[
        definition.name
      ] = declaration;
      if (value !== undefined && !this.keepDeclared(value.length)) {
        return;
      }
      if (this.internalEntities) {
        this.contextSize.addAttribute(element, definition.name, declaration);
        if (this.contextSize.total() > this.maxContextBytes) {
          this.fail(
            `the declarations the DOCTYPE internal subset makes take more than maxContextBytes (${this.maxContextBytes})`,
            "prologContext"
          );
          return;
        }
      }
    }
  }

  /**
   * An attribute default as a value read is (§3.3.3): references replaced, each white-space
   * character a space. Its references were checked as it was declared; one this parser does not
   * expand is named instead, and `undefined` when one is declared nowhere it reads.
   */
  private normalizeDefault(literal: string): string | { unresolved: string } | undefined {
    let out = "";
    let from = 0;
    for (let i = 0; i < literal.length; i++) {
      const ch = literal.charCodeAt(i);
      if (ch === TAB || ch === NL || ch === CR) {
        out += literal.slice(from, i) + " ";
        from = i + 1;
      } else if (ch === AMP) {
        const semi = literal.indexOf(";", i);
        const ref = literal.slice(i + 1, semi);
        out += literal.slice(from, i);
        from = semi + 1;
        i = semi;
        if (ref.charCodeAt(0) === HASH) {
          out += charRefValue(ref) ?? "";
        } else if (isPredefinedEntity(ref)) {
          out += XML_ENTITIES[ref];
        } else if (typeof this.declared[ref] === "string") {
          out += this.expandDeclared(ref, true);
        } else if (this.entityBindings?.get(ref)?.kind === "internal") {
          // Declared, but not read without `internalEntities`: an error where the default is used.
          return { unresolved: ref };
        } else {
          // Not declared in what is read: the unread external subset or a parameter entity may
          // declare it, and its default cannot be known — it is not given (expat skips the entity).
          return undefined;
        }
      }
    }
    return out + literal.slice(from);
  }

  /**
   * Give `tag` what the internal subset declares of its attributes: the default of one it leaves
   * out, and — for a type other than CDATA — a value with its spaces collapsed and trimmed.
   */
  private applyAttributeRules(tag: XmlSaxTag): void {
    const rules = this.attributeRules?.[tag.name];
    if (rules === undefined) {
      if (this.unresolvedDefaults !== undefined) {
        this.reportUnresolvedDefaults(tag);
      }
      return;
    }
    let attributes = tag.attributes;
    for (const name in rules) {
      const rule = rules[name];
      const key = this.keyPrefix === undefined ? name : this.prefixedKey(name);
      if (Object.hasOwn(attributes, key)) {
        if (rule.tokenized === true) {
          attributes[key] = collapseSpaces(attributes[key]);
        }
      } else if (rule.default !== undefined) {
        if (attributes === NO_ATTRIBUTES) {
          attributes = tag.attributes = Object.create(null) as Record<string, string>;
        }
        attributes[key] = rule.default;
        if (this.countPrefixedAttributes && name.length > 2 && name.indexOf(":") !== -1) {
          // A default the DTD supplies is an attribute of the tag as much as one written in it.
          this.prefixedAttributes++;
        }
      }
    }
    if (this.unresolvedDefaults !== undefined) {
      this.reportUnresolvedDefaults(tag);
    }
  }

  /** An element that leaves out an attribute whose default this parser cannot expand. */
  private reportUnresolvedDefaults(tag: XmlSaxTag): void {
    for (const [key, entity] of this.unresolvedDefaults!) {
      const at = key.indexOf("\0");
      if (key.slice(0, at) !== tag.name) {
        continue;
      }
      const name = key.slice(at + 1);
      const attributeKey = this.keyPrefix === undefined ? name : this.prefixedKey(name);
      if (!Object.hasOwn(tag.attributes, attributeKey)) {
        this.fail(this.undefinedEntityMessage(entity));
      }
    }
  }

  private sDoctypeDTDComment(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === MINUS) {
      this.state = S_DOCTYPE_DTD_COMMENT_ENDING;
    }
  }

  private sDoctypeDTDCommentEnding(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    this.state = c === MINUS ? S_DOCTYPE_DTD_COMMENT_ENDED : S_DOCTYPE_DTD_COMMENT;
  }

  private sDoctypeDTDCommentEnded(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      this.state = S_DOCTYPE_DTD;
    } else {
      this.fail("malformed comment");
      if (c !== MINUS) {
        this.state = S_DOCTYPE_DTD_COMMENT;
      }
    }
  }

  private sDoctypeDTDPI(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === QUESTION) {
      this.state = S_DOCTYPE_DTD_PI_ENDING;
    } else {
      this.appendPI(charFromCode(c));
    }
  }

  private sDoctypeDTDPIEnding(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      let end = 0;
      while (end < this.text.length && !isS(this.text.charCodeAt(end))) {
        end++;
      }
      const target = this.text.slice(0, end);
      if (!isXmlName(target) || target.toLowerCase() === "xml") {
        this.fail(
          target === ""
            ? "processing instruction without a target"
            : `invalid processing instruction target: ${target}`
        );
      }
      const handler = this._handlers.pi;
      if (handler !== undefined && !this._stopped) {
        let bodyStart = end;
        while (bodyStart < this.text.length && isS(this.text.charCodeAt(bodyStart))) {
          bodyStart++;
        }
        // Its own range, not the DOCTYPE's; the text after the DOCTYPE starts where that ends.
        this.eventStart = this.subsetPIStart;
        this.eventEnd = this.unitAt(this.i);
        handler(target, this.text.slice(bodyStart));
      }
      this.text = "";
      this.state = S_DOCTYPE_DTD;
    } else if (c !== QUESTION) {
      this.appendPI("?" + charFromCode(c));
      this.state = S_DOCTYPE_DTD_PI;
    } else {
      this.appendPI("?");
    }
  }

  /** After the internal subset's `]`: white space, then the `>` that ends the DOCTYPE. */
  private sDoctypeAfterSubset(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      this.endMarkup();
    } else if (!isS(c)) {
      this.fail("expected > after the DOCTYPE internal subset");
    }
  }

  /** A byte-fed document is read as UTF-8, so a declaration of any other encoding is refused. */
  private checkDeclaredEncoding(body: string): void {
    const match = /\bencoding\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(body);
    const encoding = match === null ? undefined : (match[1] ?? match[2]);
    if (encoding !== undefined && !UTF8_ENCODINGS.has(encoding.toLowerCase())) {
      this.fail(`the document declares encoding ${encoding}; bytes are read as UTF-8 only`);
    } else if (encoding !== undefined && ASCII_ENCODINGS.has(encoding.toLowerCase())) {
      // A byte past 0x7F is no ASCII character (XML 1.0 §4.3.3): the rest is decoded as ASCII, so
      // such a byte is an invalid character where it is, as a byte that is not UTF-8 is otherwise.
      this.switchToAscii = true;
      this.asciiDeclared = true;
      this.contextSize.ascii = true;
    }
  }

  private sEntity(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === SEMICOLON) {
      const entity = this.entity;
      let resolved: string;

      if (entity === "") {
        this.fail("empty entity");
        resolved = "&;";
      } else {
        resolved = this.parseEntity(entity);
      }

      this.appendEntityText(resolved);
      this.state = this.entityReturnState;
      this.entity = "";
    } else if (isNameChar(c) || c === HASH) {
      this.entity += charFromCode(c);
      // Security: reject excessively long entity names (>64 chars)
      if (this.entity.length > 64) {
        this.fail("entity name too long");
        this.appendEntityText("&" + this.entity);
        this.state = this.entityReturnState;
        this.entity = "";
      }
    } else {
      this.fail("invalid entity character");
      this.appendEntityText("&" + this.entity + charFromCode(c));
      this.state = this.entityReturnState;
      this.entity = "";
    }
  }

  /**
   * Add what a reference reads as — its expansion, or its own text where it was malformed — to the
   * value or text it is in. Content text nobody listens for is dropped as it is read, as the text
   * around the reference is: kept, it would grow with every reference until the next markup.
   */
  private appendEntityText(text: string): void {
    if (this.entityReturnState !== S_TEXT || this._handlers.text !== undefined) {
      this.text += text;
    }
  }

  // ===========================================================================
  // Entity Resolution
  // ===========================================================================

  private parseEntity(entity: string): string {
    if (entity[0] !== "#") {
      // Named entity — resolve first, then check expansion limit for non-predefined. `ENTITIES` is an ordinary
      // object a caller may fill: only a Name refers to an entity (§4.1), and only an own string is a definition.
      if (!isXmlName(entity)) {
        this.fail("invalid entity name: &" + entity + ";");
        return "&" + entity + ";";
      }
      const entities = this.ENTITIES;
      const resolved = Object.hasOwn(entities, entity) ? entities[entity] : undefined;
      if (typeof resolved === "string") {
        // Only count non-predefined entities against the expansion limit.
        // The 5 standard XML entities (lt, gt, amp, quot, apos) are always allowed.
        if (
          this.maxEntityExpansions > 0 &&
          entity !== "lt" &&
          entity !== "gt" &&
          entity !== "amp" &&
          entity !== "quot" &&
          entity !== "apos"
        ) {
          this._entityExpansionCount++;
          if (this._entityExpansionCount > this.maxEntityExpansions) {
            this.fail(
              `entity expansion limit (${this.maxEntityExpansions}) exceeded — possible XML bomb`,
              "entityExpansions"
            );
            return "";
          }
        }
        // A caller's own entity — or a predefined one given another value — is text the document
        // multiplies by referring to it, as a declared one is: it counts toward the same length.
        if (
          !(Object.hasOwn(XML_ENTITIES, entity) && XML_ENTITIES[entity] === resolved) &&
          !this.countExpansionLength(resolved.length)
        ) {
          return "";
        }
        return resolved;
      }
      if (typeof this.declared[entity] === "string") {
        return this.expandDeclared(entity);
      }
      this.fail(this.undefinedEntityMessage(entity));
      return "&" + entity + ";";
    }

    // Numeric character reference: whole digits only (`&#65abc;` is not `A`), and a character XML
    // can hold (no NUL, surrogate, U+FFFE or beyond U+10FFFF).
    const char = charRefValue(entity);
    if (char === undefined) {
      this.fail("invalid character entity: &" + entity + ";");
      return "";
    }
    return char;
  }

  /**
   * Expand an entity the internal subset declared (or the `context` gave): its cost — every entity
   * reference it expands, nested ones included — is counted against `maxEntityExpansions`, and the
   * length of its text against `maxEntityExpansionLength`, before the text is built.
   */
  /** Why `name` cannot be expanded, as an error says it. */
  private undefinedEntityMessage(name: string): string {
    return (
      "undefined entity: " +
      name +
      (this.externalEntities?.has(name)
        ? " (external entities are not read)"
        : this.sawInternalSubset && !this.internalEntities
          ? " (entities declared in a DOCTYPE internal subset are not supported)"
          : "")
    );
  }

  private expandDeclared(
    name: string,
    inAttribute = this.entityReturnState === S_ATTRIB_VALUE_QUOTED
  ): string {
    const entities = (this.declaredEntities ??= new DeclaredEntities(this.declared));
    const max = this.maxEntityExpansions;
    try {
      const cost = entities.expansions(name, max > 0 ? max - this._entityExpansionCount : Infinity);
      if (max > 0) {
        this._entityExpansionCount += cost;
      }
      // Measured, not built: a text past the limit is refused before it takes any memory.
      if (!this.countExpansionLength(entities.length(name))) {
        return "";
      }
      return entities.text(name, inAttribute);
    } catch (err) {
      if (!(err instanceof EntityProblem)) {
        throw err;
      }
      if (err.limit) {
        this.fail(
          `entity expansion limit (${max}) exceeded — possible XML bomb`,
          "entityExpansions"
        );
      } else {
        this.fail(err.message);
      }
      return "";
    }
  }

  /**
   * Count `length` more characters of entity text against `maxEntityExpansionLength`; whether it is
   * still within it. Past it, the limit is reported.
   */
  private countExpansionLength(length: number): boolean {
    const max = this.maxEntityExpansionLength;
    if (max <= 0) {
      return true;
    }
    this.entityExpansionLength += length;
    if (this.entityExpansionLength <= max) {
      return true;
    }
    this.fail(
      `entity expansion length limit (${max}) exceeded — possible XML bomb`,
      "entityExpansionLength"
    );
    return false;
  }

  // ===========================================================================
  // Helpers
  // ===========================================================================

  private skipSpaces(): number {
    while (true) {
      const c = this.getCode();
      if (c === -1 || !isS(c)) {
        return c;
      }
    }
  }

  // ===========================================================================
  // Namespace Helpers
  // ===========================================================================

  /** Split "prefix:local" into [prefix, local]. Returns ["", name] if no prefix. */
  private splitQName(qname: string): [string, string] {
    const colon = qname.indexOf(":");
    if (colon < 0) {
      return ["", qname];
    }
    return [qname.slice(0, colon), qname.slice(colon + 1)];
  }

  /** Look up the URI for a namespace prefix by walking the stack top-down. */
  private resolveNs(prefix: string): string {
    for (let i = this._nsStack.length - 1; i >= 0; i--) {
      const uri = this._nsStack[i][prefix];
      if (uri !== undefined) {
        return uri;
      }
    }
    return "";
  }

  /** Extract xmlns declarations from tag attributes and populate tag namespace fields. */
  private applyNamespaces(tag: XmlSaxTag): void {
    const XML_NS = "http://www.w3.org/XML/1998/namespace";
    const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

    // 1. Collect xmlns declarations from already-populated attributes. A scope is a null-prototype
    // record: a prefix is bound only by a declaration, never by what `Object.prototype` holds —
    // `constructor:r` would otherwise resolve to the `Object` function.
    const nsDecls = Object.create(null) as Record<string, string>;
    const attrs = tag.attributes;
    for (const name in attrs) {
      if (!Object.prototype.hasOwnProperty.call(attrs, name)) {
        continue;
      }
      if (name === "xmlns") {
        const uri = attrs[name];
        // Namespaces in XML 1.0 §3: neither reserved name may be the default namespace either.
        if (uri === XML_NS || uri === XMLNS_NS) {
          this.fail(`namespace "${uri}" must not be the default namespace`);
        }
        nsDecls[""] = uri;
      } else if (name.startsWith("xmlns:")) {
        const prefix = name.slice(6);
        const uri = attrs[name];

        // Namespace well-formedness checks (XML Namespaces §3)
        if (!isNCName(prefix)) {
          this.fail(`invalid namespace declaration "${name}": the prefix must be an NCName`);
        } else if (prefix === "xmlns") {
          this.fail('the "xmlns" prefix must not be declared');
        } else if (prefix === "xml" && uri !== XML_NS) {
          this.fail(`the "xml" prefix must be bound to "${XML_NS}"`);
        } else if (uri === XML_NS && prefix !== "xml") {
          this.fail(`namespace "${XML_NS}" must only be bound to the "xml" prefix`);
        } else if (uri === XMLNS_NS) {
          this.fail(`namespace "${XMLNS_NS}" must not be bound to any prefix`);
        } else if (uri === "") {
          // Namespaces in XML 1.0 §2.2: only the default namespace can be undeclared.
          this.fail(`the "${prefix}" prefix must not be bound to an empty namespace name`);
        }

        nsDecls[prefix] = uri;
      }
    }

    // 2. Push namespace scope
    this._nsStack.push(nsDecls);

    // 3. Set namespace fields on the tag element
    if (Object.keys(nsDecls).length > 0) {
      tag.ns = nsDecls;
    }
    const [prefix, local] = this.splitQName(tag.name);

    // A QName is an NCName, or two joined by one colon (Namespaces in XML 1.0 §4).
    if (!isQName(tag.name)) {
      this.fail(`invalid QName "${tag.name}": prefix and local part must be NCNames`);
    }

    tag.prefix = prefix;
    tag.local = local;
    tag.uri = this.resolveNs(prefix);

    // 4. Report unbound element prefix
    if (prefix === "xmlns") {
      this.fail(
        `invalid element name "${tag.name}": the "xmlns" prefix is reserved for declarations`
      );
    } else if (prefix !== "" && tag.uri === "") {
      this.fail(`unbound namespace prefix: "${prefix}"`);
    }

    // 5. Resolve attribute namespaces, check unbound prefixes and expanded-name duplicates
    //    Note: unprefixed attributes do NOT inherit the default namespace (XML Namespaces §6.2)
    const expandedNames = new Set<string>();
    for (const name in attrs) {
      if (!Object.prototype.hasOwnProperty.call(attrs, name)) {
        continue;
      }
      // Skip xmlns declarations — they are not ordinary attributes
      if (name === "xmlns" || name.startsWith("xmlns:")) {
        continue;
      }
      const [aPrefix, aLocal] = this.splitQName(name);
      if (!isQName(name)) {
        this.fail(`invalid attribute QName "${name}": prefix and local part must be NCNames`);
      }
      if (aPrefix !== "") {
        const aUri = this.resolveNs(aPrefix);
        if (aUri === "") {
          this.fail(`unbound namespace prefix on attribute: "${aPrefix}"`);
        }
        // Check expanded-name duplicate: same URI + local name
        const expanded = aUri + "\0" + aLocal;
        if (expandedNames.has(expanded)) {
          this.fail(`duplicate attribute by expanded name: {${aUri}}${aLocal}`);
        }
        expandedNames.add(expanded);
      }
    }
  }

  // ===========================================================================
  // Tag Emission
  // ===========================================================================

  /**
   * Record an attribute on the tag being opened.
   *
   * Written straight onto `tag.attributes` rather than staged in a list and copied at `openTag`: that cost one
   * `{ name, value }` object per attribute and one array per tag — on a worksheet, three of each per cell — for no
   * information the attributes object does not already hold. The later value still wins, and a duplicate is still
   * reported when the tag opens — at the same position, after any error found earlier in the tag — so error order and
   * positions are unchanged. Only a tag that has one pays for the list.
   */
  private addAttribute(name: string, value: string): void {
    // `a:b` is the shortest prefixed name; most of a worksheet's (`r`, `s`, `t`) are not looked into.
    if (this.countPrefixedAttributes && name.length > 2 && name.indexOf(":") !== -1) {
      this.prefixedAttributes++;
    }
    const tag = this.tag!;
    let attributes = tag.attributes;
    let key = name;
    if (this.keyPrefix !== undefined) {
      key = this.prefixedKey(name);
      if (attributes === NO_ATTRIBUTES) {
        attributes = tag.attributes = Object.create(null) as Record<string, string>;
      }
    }
    if (key in attributes) {
      (this.duplicateAttributes ??= []).push(name);
    }
    attributes[key] = value;
  }

  /** `name` with {@link keyPrefix}, made once per name. */
  private prefixedKey(name: string): string {
    const keys = (this.prefixedKeys ??= new Map());
    let key = keys.get(name);
    if (key === undefined) {
      key = this.keyPrefix + name;
      if (keys.size < 4096) {
        keys.set(name, key);
      }
    }
    return key;
  }

  /**
   * Write attribute names with `prefix` — `Xml.parseToObject`'s `attributePrefix` — so a consumer
   * building those objects can take a tag's `attributes` as it is rather than copy it into another
   * object. A tag without attributes then shares {@link NO_ATTRIBUTES}, which is frozen. For this
   * module's record reader only: the keys are no longer the attribute names.
   *
   * @internal
   */
  prefixAttributeNames(prefix: string): void {
    this.keyPrefix = prefix;
  }

  private reportDuplicateAttributes(): void {
    const duplicates = this.duplicateAttributes;
    if (duplicates !== null) {
      this.duplicateAttributes = null;
      for (const name of duplicates) {
        this.fail(`duplicate attribute: ${name}`);
      }
    }
  }

  private openTag(): void {
    const tag = this.tag!;
    tag.isSelfClosing = false;
    this.reportDuplicateAttributes();
    if (this.attributeRules !== undefined || this.unresolvedDefaults !== undefined) {
      this.applyAttributeRules(tag);
    }

    if (this.xmlns) {
      this.applyNamespaces(tag);
    }

    // Security: warn on nesting depth exceeded but continue processing
    // so that close tags still match (prevents infinite error loops)
    if (this.maxDepth > 0 && this.tags.length >= this.maxDepth) {
      this.fail(`maximum element nesting depth (${this.maxDepth}) exceeded`, "depth");
    }
    if (this._stopped) {
      // An error's handler stopped the parse: the tag is not reported.
      return;
    }

    if (this.rootName === undefined && this.tags.length === 0) {
      this.noteRoot(tag.name);
      if (this._stopped) {
        return;
      }
    }
    this.sawTag = true;
    this.rootEnded = false;
    this.markupEvent();
    this._handlers.opentag?.(tag);
    this.tags.push(tag);
    this.name = "";
    this.state = S_TEXT;
  }

  /** The root element's start tag has been read: the prolog's context is complete. */
  private noteRoot(name: string): void {
    this.rootName = name;
    if (this.internalEntities && this.contextSize.total(name) > this.maxContextBytes) {
      this.fail(
        `the prolog context takes more than maxContextBytes (${this.maxContextBytes})`,
        "prologContext"
      );
    }
  }

  private openSelfClosingTag(): void {
    const tag = this.tag!;
    tag.isSelfClosing = true;
    this.reportDuplicateAttributes();
    if (this.attributeRules !== undefined || this.unresolvedDefaults !== undefined) {
      this.applyAttributeRules(tag);
    }

    if (this.xmlns) {
      this.applyNamespaces(tag);
    }

    // Security: warn on nesting depth exceeded but continue processing
    if (this.maxDepth > 0 && this.tags.length >= this.maxDepth) {
      this.fail(`maximum element nesting depth (${this.maxDepth}) exceeded`, "depth");
    }
    if (this._stopped) {
      // An error's handler stopped the parse: the tag is not reported.
      return;
    }

    if (this.rootName === undefined && this.tags.length === 0) {
      this.noteRoot(tag.name);
      if (this._stopped) {
        return;
      }
    }
    this.sawTag = true;
    // A self-closing element with nothing open around it is the root, if it has the root's name.
    this.rootEnded = this.resume && this.tags.length === 0 && tag.name === this.context?.root;
    this.markupEvent();
    this._handlers.opentag?.(tag);
    // `stop()` from the start's handler ends the parse there; a pause does not split one tag's events.
    if (!this._stopped) {
      this._handlers.closetag?.(tag);
    }

    // Pop namespace scope for self-closing tags immediately
    if (this.xmlns) {
      this._nsStack.pop();
    }

    if (this.tags.length === 0 && !this.resume) {
      this.closedRoot = true;
    }
    this.name = "";
    this.state = S_TEXT;
  }

  private closeTag(): void {
    const { tags, name } = this;
    this.state = S_TEXT;
    this.name = "";
    this.markupEvent();

    if (name === "") {
      this.fail("empty close tag");
      this.text += "</>";
      return;
    }

    let found = false;
    for (let i = tags.length - 1; i >= 0; i--) {
      const tag = tags[i];
      if (tag.name === name) {
        while (tags.length > i) {
          const t = tags.pop()!;
          this._handlers.closetag?.(t);
          if (this.xmlns) {
            this._nsStack.pop();
          }
          if (this._stopped) {
            // The handler stopped the parse: the elements it did not close are not reported.
            return;
          }
          if (tags.length > i) {
            this.fail("unclosed tag: " + t.name);
          }
        }
        found = true;
        break;
      }
    }

    if (!found) {
      if (this.resume && tags.length === 0) {
        // The end tag of an element opened before this parser started — the root's, when it has the root's name.
        this.sawTag = true;
        this.rootEnded = name === this.context?.root;
        this.onOuterClose?.(name);
        return;
      }
      this.fail("unmatched close tag: " + name);
      this.text += "</" + name + ">";
    }

    if (tags.length === 0) {
      if (this.resume) {
        // An element opened since the start closed with nothing left open: the root, if it has its name.
        this.rootEnded = name === this.context?.root;
      } else {
        this.closedRoot = true;
      }
    }
  }

  private end(): this {
    if (this.decoder?.endedInsideCharacter) {
      this.failTruncated("unexpected end of document inside a UTF-8 character");
    }
    // Input that stops inside a construct is truncated. What `text` holds then belongs to the construct — a
    // comment's or an attribute value's characters — and is not the document's text, so it is not flushed.
    const unfinished = this.state !== S_TEXT;
    if (unfinished) {
      this.failTruncated(`unexpected end of document inside ${unfinishedConstruct(this.state)}`);
    }

    if (!this.sawRoot) {
      this.failTruncated("document must contain a root element");
    }

    // In `resume` mode an element opened since the start and not closed is cut off like any other —
    // the elements opened before the start are not this parser's to see. But a document ends with its
    // root element's end tag: a parser that knows the root's name and has read a tag must have read
    // that one last, unless it started after it, in the comments and white space that may follow.
    if (
      this.resume &&
      !unfinished &&
      this.context !== undefined &&
      this.tags.length === 0 &&
      this.sawTag &&
      !this.rootEnded
    ) {
      this.failTruncated(`unexpected end of document inside <${this.context.root}>`);
    }
    while (this.tags.length > 0) {
      const tag = this.tags.pop()!;
      this.failTruncated("unclosed tag: " + tag.name);
    }

    if (!unfinished && !this._stopped && this.text.length > 0 && this._handlers.text) {
      this.eventStart = this.textStart;
      this.eventEnd = this.unitAt(this.i);
      this._handlers.text(this.text);
      this.text = "";
    }

    this._closed = true;
    this._init();
    return this;
  }
}

// =============================================================================
// Async Generator - parseSax
// =============================================================================

function checkChunkSource(where: string, iterable: unknown): void {
  if (
    typeof iterable !== "object" ||
    iterable === null ||
    (typeof (iterable as { [Symbol.asyncIterator]?: unknown })[Symbol.asyncIterator] !==
      "function" &&
      typeof (iterable as { [Symbol.iterator]?: unknown })[Symbol.iterator] !== "function")
  ) {
    throw new XmlArgumentError(
      "iterable",
      `${where}: iterable must be an iterable or async iterable of string or byte chunks`
    );
  }
}

/** A chunk of a source the stream adapters read: text, or the document's bytes. */
function checkChunk(where: string, chunk: unknown): asserts chunk is string | Uint8Array {
  if (typeof chunk !== "string" && !isUint8Array(chunk)) {
    throw new XmlArgumentError(
      "iterable",
      `${where}: iterable must yield string or Uint8Array chunks, got ${chunk === null ? "null" : typeof chunk}`
    );
  }
}

/**
 * Parse an async-iterable of XML chunks as a stream of SAX event batches.
 *
 * Yields an array of {@link XmlSaxEvent} per input chunk. This is the
 * primary integration point for streaming XML parsing.
 *
 * @param iterable - Async iterable of string or Uint8Array chunks.
 * @param options - Parser options.
 *
 * @example
 * ```ts
 * for await (const events of Xml.parseSaxStream(stream)) {
 *   for (const event of events) {
 *     if (event.eventType === "opentag") console.log(event.value.name);
 *   }
 * }
 * ```
 */
function parseSax(
  iterable: AsyncIterable<string | Uint8Array> | Iterable<string | Uint8Array>,
  options?: XmlSaxOptions
): AsyncGenerator<XmlSaxEvent[], void, undefined> {
  // Checked when called, not on the first `next()` a generator body would wait for, and under
  // the call the caller made rather than the `Xml.SaxParser` it builds. Read once: what is checked
  // is what the parser is built from, however the caller's object changes before the first `next()`.
  checkChunkSource("Xml.parseSaxStream", iterable);
  const checked = snapshotSaxOptions(options) ?? {};
  checkSaxOptions("Xml.parseSaxStream", checked);
  return saxEvents(iterable, checked);
}

/** {@link parseSax} once its arguments are checked. */
async function* saxEvents(
  iterable: AsyncIterable<string | Uint8Array> | Iterable<string | Uint8Array>,
  options: XmlSaxOptions
): AsyncGenerator<XmlSaxEvent[], void, undefined> {
  const parser = SaxParser.withCheckedOptions(options);
  // The first error ends the stream, as it is found: nothing after it in the chunk is read.
  parser.on("error", err => {
    throw toError(err);
  });

  let events: XmlSaxEvent[] = [];
  parser.on("opentag", value => events.push({ eventType: "opentag", value }));
  parser.on("text", value => events.push({ eventType: "text", value }));
  parser.on("closetag", value => events.push({ eventType: "closetag", value }));
  parser.on("cdata", value => events.push({ eventType: "cdata", value }));
  parser.on("comment", value => events.push({ eventType: "comment", value }));
  parser.on("pi", (target, body) => events.push({ eventType: "pi", value: { target, body } }));

  // Bytes go to the parser as bytes: it decodes them, with `invalidCharHandling`, `byteOffset` and
  // the declared encoding applied, and positions in bytes.
  for await (const chunk of iterable) {
    checkChunk("Xml.parseSaxStream", chunk);
    parser.write(chunk);
    if (events.length > 0) {
      yield events;
      events = [];
    }
  }
  parser.close();
  if (events.length > 0) {
    yield events;
  }
}

// =============================================================================
// Direct SAX Stream - saxStream
// =============================================================================

const saxParserSchema = s.custom(v => v instanceof SaxParser, "an Xml.SaxParser");

/**
 * High-performance direct SAX streaming: feed an async-iterable to a
 * pre-configured SaxParser **without** creating intermediate event objects.
 *
 * The caller registers callbacks on the parser before calling this function.
 * This eliminates per-event `{ eventType, value }` object allocation and the
 * overhead of the async generator protocol, making it significantly faster
 * for hot inner loops like worksheet/shared-string parsing.
 *
 * @param parser - A SaxParser with handlers already registered via `.on()`.
 * @param iterable - Async iterable of string or Uint8Array chunks.
 *
 * @example
 * ```ts
 * const parser = new Xml.SaxParser();
 * parser.on("opentag", tag => handleOpen(tag));
 * parser.on("text", text => handleText(text));
 * parser.on("closetag", tag => handleClose(tag));
 * await Xml.feedSax(parser, stream);
 * ```
 */
async function saxStream(
  parser: SaxParser,
  iterable: AsyncIterable<string | Uint8Array> | Iterable<string | Uint8Array>
): Promise<void> {
  check(saxParserSchema, parser, xmlArgument("Xml.feedSax", "parser"));
  checkChunkSource("Xml.feedSax", iterable);

  // The caller's handler sees each error first; the first one then ends the feed, as it is found.
  // The handler is the caller's again once the feed is over, however it ends.
  const callerErrorHandler = parser.getHandler("error");
  parser.on("error", (err: Error) => {
    callerErrorHandler?.(err);
    throw toError(err);
  });
  try {
    for await (const chunk of iterable) {
      checkChunk("Xml.feedSax", chunk);
      parser.write(chunk);
    }
    parser.close();
  } finally {
    parser.setHandler("error", callerErrorHandler);
  }
}

export { SaxParser, parseSax, saxStream };
