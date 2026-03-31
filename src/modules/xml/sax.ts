/**
 * SAX XML Parser
 *
 * High-performance streaming SAX parser for XML.
 * Zero external dependencies. Optimized for common XML patterns.
 *
 * Migrated from the excel module's parse-sax.ts into the standalone XML module.
 * Enhancements over the original:
 * - Uses @xml/types for event types (SaxTag, SaxEvent, SaxHandlers)
 * - Uses @xml/errors for parse errors (XmlParseError)
 * - Exposes CDATA, comment, and processing-instruction event handlers
 * - Provides both callback API and async-generator API
 *
 * Based on XML 1.0 specification with fast-path optimizations for ASCII.
 */

import { XmlParseError } from "@xml/errors";
import type { SaxEventAny, SaxHandlers, SaxOptions, SaxTag } from "@xml/types";

// =============================================================================
// Character Codes
// =============================================================================

const TAB = 9;
const NL = 0xa;
const CR = 0xd;
const SPACE = 0x20;
const BANG = 0x21; // !
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

const NAME_START_CHAR_ASCII = /* @__PURE__ */ (() => {
  const t = new Uint8Array(128);
  for (let i = 0x61; i <= 0x7a; i++) {
    t[i] = 1;
  } // a-z
  for (let i = 0x41; i <= 0x5a; i++) {
    t[i] = 1;
  } // A-Z
  t[0x5f] = 1; // _
  t[0x3a] = 1; // :
  return t;
})();

const NAME_CHAR_ASCII = /* @__PURE__ */ (() => {
  const t = new Uint8Array(128);
  for (let i = 0x61; i <= 0x7a; i++) {
    t[i] = 1;
  } // a-z
  for (let i = 0x41; i <= 0x5a; i++) {
    t[i] = 1;
  } // A-Z
  for (let i = 0x30; i <= 0x39; i++) {
    t[i] = 1;
  } // 0-9
  t[0x5f] = 1; // _
  t[0x3a] = 1; // :
  t[0x2d] = 1; // -
  t[0x2e] = 1; // .
  return t;
})();

// =============================================================================
// Character Classification
// =============================================================================

/**
 * Lookup table for valid XML 1.0 characters in the ASCII range (0-127).
 * Invalid: 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F
 * Valid:   0x09 (TAB), 0x0A (LF), 0x0D (CR), 0x20-0x7E
 */
const VALID_CHAR_ASCII = /* @__PURE__ */ (() => {
  const t = new Uint8Array(128);
  t[0x09] = 1; // TAB
  t[0x0a] = 1; // LF
  t[0x0d] = 1; // CR
  for (let i = 0x20; i <= 0x7e; i++) {
    t[i] = 1;
  }
  return t;
})();

function isS(c: number): boolean {
  return c === SPACE || c === NL || c === CR || c === TAB;
}

function isQuote(c: number): boolean {
  return c === DQUOTE || c === SQUOTE;
}

function isNameStartChar(c: number): boolean {
  if (c < 128) {
    return NAME_START_CHAR_ASCII[c] === 1;
  }
  return (
    (c >= 0xc0 && c <= 0xd6) ||
    (c >= 0xd8 && c <= 0xf6) ||
    (c >= 0xf8 && c <= 0x2ff) ||
    (c >= 0x370 && c <= 0x37d) ||
    (c >= 0x37f && c <= 0x1fff) ||
    c === 0x200c ||
    c === 0x200d ||
    (c >= 0x2070 && c <= 0x218f) ||
    (c >= 0x2c00 && c <= 0x2fef) ||
    (c >= 0x3001 && c <= 0xd7ff) ||
    (c >= 0xf900 && c <= 0xfdcf) ||
    (c >= 0xfdf0 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0xeffff)
  );
}

function isNameChar(c: number): boolean {
  if (c < 128) {
    return NAME_CHAR_ASCII[c] === 1;
  }
  return (
    c === 0xb7 ||
    (c >= 0xc0 && c <= 0xd6) ||
    (c >= 0xd8 && c <= 0xf6) ||
    (c >= 0xf8 && c <= 0x2ff) ||
    (c >= 0x300 && c <= 0x36f) ||
    (c >= 0x370 && c <= 0x37d) ||
    (c >= 0x37f && c <= 0x1fff) ||
    c === 0x200c ||
    c === 0x200d ||
    (c >= 0x203f && c <= 0x2040) ||
    (c >= 0x2070 && c <= 0x218f) ||
    (c >= 0x2c00 && c <= 0x2fef) ||
    (c >= 0x3001 && c <= 0xd7ff) ||
    (c >= 0xf900 && c <= 0xfdcf) ||
    (c >= 0xfdf0 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0xeffff)
  );
}

function isChar10(c: number): boolean {
  return (
    (c >= SPACE && c <= 0xd7ff) ||
    c === NL ||
    c === CR ||
    c === TAB ||
    (c >= 0xe000 && c <= 0xfffd) ||
    (c >= 0x10000 && c <= 0x10ffff)
  );
}

// =============================================================================
// Built-in XML Entities
// =============================================================================

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
  private xmlns: boolean;
  private maxDepth: number;
  private maxEntityExpansions: number;

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
  private tags: SaxTag[] = [];
  private tag: SaxTag | null = null;
  private attribList: Array<{ name: string; value: string }> = [];
  private entity: string = "";
  private entityReturnState: number = S_TEXT;
  private openWakaBang: string = "";
  private sawRoot: boolean = false;
  private closedRoot: boolean = false;
  private carriedFromPrevious?: string;
  private _closed: boolean = false;
  private reportedTextBeforeRoot: boolean = false;
  private reportedTextAfterRoot: boolean = false;

  // Position tracking
  line: number = 1;
  column: number = 0;
  private positionAtNewLine: number = 0;
  private chunkPosition: number = 0;

  // Entity storage
  ENTITIES: Record<string, string> = { ...XML_ENTITIES };

  // Event handlers
  private _handlers: SaxHandlers = {};

  constructor(options?: SaxOptions) {
    this.trackPosition = options?.position !== false;
    this.fileName = options?.fileName;
    this.fragment = options?.fragment ?? false;
    this.xmlns = options?.xmlns ?? false;
    this.maxDepth = options?.maxDepth !== undefined ? options.maxDepth : 256;
    this.maxEntityExpansions =
      options?.maxEntityExpansions !== undefined ? options.maxEntityExpansions : 10000;
    this._init();
  }

  get closed(): boolean {
    return this._closed;
  }

  get position(): number {
    return this.chunkPosition + this.i;
  }

  private _init(): void {
    this.state = S_TEXT;
    this.text = "";
    this.name = "";
    this.q = null;
    this.tags = [];
    this.tag = null;
    this.attribList = [];
    this.entity = "";
    this.openWakaBang = "";
    this.sawRoot = this.fragment;
    this.closedRoot = this.fragment;
    this.reportedTextBeforeRoot = this.fragment;
    this.reportedTextAfterRoot = this.fragment;
    this.carriedFromPrevious = undefined;
    // Note: _closed is NOT reset here — it is managed by end() and write().
    this.line = 1;
    this.column = 0;
    this.positionAtNewLine = 0;
    this.chunkPosition = 0;
    this.chunk = "";
    this.i = 0;
    this.prevI = 0;
    this._nsStack = [];
    this._entityExpansionCount = 0;
  }

  // ===========================================================================
  // Event Registration
  // ===========================================================================

  on(name: "opentag", handler: (tag: SaxTag) => void): void;
  on(name: "text", handler: (text: string) => void): void;
  on(name: "closetag", handler: (tag: SaxTag) => void): void;
  on(name: "cdata", handler: (text: string) => void): void;
  on(name: "comment", handler: (text: string) => void): void;
  on(name: "pi", handler: (target: string, body: string) => void): void;
  on(name: "error", handler: (err: Error) => void): void;
  on(name: string, handler: any): void {
    (this._handlers as any)[name] = handler;
  }

  off(name: string): void {
    delete (this._handlers as any)[name];
  }

  // ===========================================================================
  // Error Handling
  // ===========================================================================

  private makeError(message: string): XmlParseError {
    return new XmlParseError(message, {
      line: this.trackPosition ? this.line : undefined,
      column: this.trackPosition ? this.column : undefined,
      fileName: this.fileName
    });
  }

  fail(message: string): this {
    const err = this.makeError(message);
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

  write(chunk: string | null): this {
    // Auto-reset for parser reuse: if previously closed, start fresh.
    if (this._closed) {
      if (chunk === null) {
        return this; // no-op: already closed
      }
      this._closed = false;
      this._init();
    }

    let end = false;
    if (chunk === null) {
      end = true;
      chunk = "";
    }

    if (this.carriedFromPrevious !== undefined) {
      chunk = this.carriedFromPrevious + chunk;
      this.carriedFromPrevious = undefined;
    }

    let limit = chunk.length;
    if (!end && limit > 0) {
      const lastCode = chunk.charCodeAt(limit - 1);
      if (lastCode === CR || (lastCode >= 0xd800 && lastCode <= 0xdbff)) {
        this.carriedFromPrevious = chunk[limit - 1];
        limit--;
        chunk = chunk.slice(0, limit);
      }
    }

    this.chunk = chunk;
    this.i = 0;

    while (this.i < limit) {
      this.processState();
    }

    this.chunkPosition += limit;

    return end ? this.end() : this;
  }

  close(): this {
    return this.write(null);
  }

  // ===========================================================================
  // Character Reading
  // ===========================================================================

  private getCode(): number {
    const { chunk, i } = this;
    this.prevI = i;
    this.i = i + 1;

    if (i >= chunk.length) {
      return -1;
    }

    const code = chunk.charCodeAt(i);

    if (code < 0x0a || (code > 0x0d && code < 0xd800)) {
      if (this.trackPosition) {
        this.column++;
      }
      // Reject invalid XML 1.0 characters: 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F
      if (code < 128 && VALID_CHAR_ASCII[code] === 0) {
        this.fail("invalid XML character: 0x" + code.toString(16));
        return code; // return the char but error is reported
      }
      return code;
    }

    // Handle surrogates
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = chunk.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        this.i = i + 2;
        if (this.trackPosition) {
          this.column++;
        }
        return 0x10000 + ((code - 0xd800) * 0x400 + (next - 0xdc00));
      }
      // Lone high surrogate — invalid XML character
      this.fail("invalid XML character: lone surrogate 0x" + code.toString(16));
    }

    // Lone low surrogate — invalid XML character
    if (code >= 0xdc00 && code <= 0xdfff) {
      this.fail("invalid XML character: lone surrogate 0x" + code.toString(16));
    }

    // Handle CR
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
    if (code === NL && this.trackPosition) {
      this.line++;
      this.column = 0;
      this.positionAtNewLine = this.position;
    } else if (this.trackPosition) {
      this.column++;
    }

    // Reject invalid XML 1.0 characters that reach here:
    // 0x0B (VT), 0x0C (FF), 0xFFFE, 0xFFFF
    if (code === 0x0b || code === 0x0c || code === 0xfffe || code === 0xffff) {
      this.fail("invalid XML character: 0x" + code.toString(16));
    }

    return code;
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
        this.sText();
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
    }
  }

  // ===========================================================================
  // State Handlers
  // ===========================================================================

  private sText(): void {
    if (this.tags.length !== 0) {
      this.handleTextInRoot();
    } else {
      this.handleTextOutsideRoot();
    }
  }

  private handleTextInRoot(): void {
    const { chunk } = this;
    let { i: start } = this;
    const handler = this._handlers.text;

    while (true) {
      const c = this.getCode();

      if (c === -1) {
        if (handler && start < this.i) {
          this.text += chunk.slice(start, this.i);
        }
        return;
      }

      if (c === LESS) {
        if (handler) {
          const slice = chunk.slice(start, this.prevI);
          if (this.text.length > 0) {
            handler(this.text + slice);
            this.text = "";
          } else if (slice.length > 0) {
            handler(slice);
          }
        }
        this.state = S_OPEN_WAKA;
        return;
      }

      if (c === AMP) {
        if (handler) {
          this.text += chunk.slice(start, this.prevI);
        }
        this.state = S_ENTITY;
        this.entityReturnState = S_TEXT;
        this.entity = "";
        return;
      }

      if (c === NL) {
        if (handler) {
          this.text += chunk.slice(start, this.prevI) + "\n";
        }
        start = this.i;
      }
    }
  }

  private handleTextOutsideRoot(): void {
    const { chunk } = this;
    let { i: start } = this;
    const handler = this._handlers.text;
    let nonSpace = false;

    while (true) {
      const c = this.getCode();

      if (c === -1) {
        if (handler && start < this.i) {
          this.text += chunk.slice(start, this.i);
        }
        break;
      }

      if (c === LESS) {
        if (handler) {
          const slice = chunk.slice(start, this.prevI);
          if (this.text.length > 0) {
            handler(this.text + slice);
            this.text = "";
          } else if (slice.length > 0) {
            handler(slice);
          }
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
        nonSpace = true;
        break;
      }

      if (c === NL) {
        if (handler) {
          this.text += chunk.slice(start, this.prevI) + "\n";
        }
        start = this.i;
      } else if (!isS(c)) {
        nonSpace = true;
      }
    }

    if (nonSpace) {
      if (!this.sawRoot && !this.reportedTextBeforeRoot) {
        this.fail("text data outside of root node.");
        this.reportedTextBeforeRoot = true;
      }
      if (this.closedRoot && !this.reportedTextAfterRoot) {
        this.fail("text data outside of root node.");
        this.reportedTextAfterRoot = true;
      }
    }
  }

  private sOpenWaka(): void {
    const c = this.getCode();

    if (c === -1) {
      return;
    }

    if (isNameStartChar(c)) {
      this.state = S_OPEN_TAG;
      this.name = charFromCode(c);
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
        this.text = "";
        break;
      default:
        this.fail("unexpected character in tag");
        this.state = S_TEXT;
    }
  }

  private sOpenWakaBang(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    this.openWakaBang += charFromCode(c);

    switch (this.openWakaBang) {
      case "[CDATA[":
        this.state = S_CDATA;
        this.text = "";
        this.openWakaBang = "";
        break;
      case "--":
        this.state = S_COMMENT;
        this.text = "";
        this.openWakaBang = "";
        break;
      case "DOCTYPE":
        this.state = S_DOCTYPE;
        this.text = "";
        this.openWakaBang = "";
        break;
      default:
        if (this.openWakaBang.length >= 7) {
          this.fail("incorrect syntax");
          this.state = S_TEXT;
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
    while (pos < chunk.length) {
      const code = chunk.charCodeAt(pos);
      if (code < 128 ? NAME_CHAR_ASCII[code] === 1 : isNameChar(code)) {
        pos++;
      } else {
        break;
      }
    }

    if (pos >= chunk.length) {
      // Chunk ended in the middle of tag name
      this.name += chunk.slice(nameStart, pos);
      this.i = pos;
      if (this.trackPosition) {
        this.column += pos - nameStart;
      }
      return;
    }

    this.name += chunk.slice(nameStart, pos);
    this.i = pos;
    if (this.trackPosition) {
      this.column += pos - nameStart;
    }

    // Now read the delimiter character
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    this.tag = {
      name: this.name,
      attributes: Object.create(null) as Record<string, string>,
      isSelfClosing: false
    };
    this.attribList = [];
    this.sawRoot = true;

    if (c === GREATER) {
      this.openTag();
    } else if (c === FORWARD_SLASH) {
      this.state = S_OPEN_TAG_SLASH;
    } else if (isS(c)) {
      this.state = S_ATTRIB;
    } else {
      this.fail("unexpected character in tag");
      this.state = S_ATTRIB;
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

    while (pos < chunk.length) {
      const code = chunk.charCodeAt(pos);
      if (code < 128 ? NAME_CHAR_ASCII[code] === 1 : isNameChar(code)) {
        pos++;
      } else {
        break;
      }
    }

    if (pos >= chunk.length) {
      this.name += chunk.slice(nameStart, pos);
      this.i = pos;
      if (this.trackPosition) {
        this.column += pos - nameStart;
      }
      return;
    }

    this.name += chunk.slice(nameStart, pos);
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
      this.attribList.push({ name: this.name, value: this.name });
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
      this.state = S_TEXT;
    }
  }

  private sAttribValueQuoted(): void {
    const { q, chunk } = this;
    let { i: start } = this;

    while (true) {
      const c = this.getCode();

      if (c === -1) {
        this.text += chunk.slice(start, this.i);
        return;
      }

      if (c === q) {
        this.attribList.push({
          name: this.name,
          value: this.text + chunk.slice(start, this.prevI)
        });
        this.name = "";
        this.text = "";
        this.q = null;
        this.state = S_ATTRIB_VALUE_CLOSED;
        return;
      }

      if (c === AMP) {
        this.text += chunk.slice(start, this.prevI);
        this.state = S_ENTITY;
        this.entityReturnState = S_ATTRIB_VALUE_QUOTED;
        this.entity = "";
        return;
      }

      if (c === NL || c === TAB) {
        this.text += chunk.slice(start, this.prevI) + " ";
        start = this.i;
      }

      if (c === LESS) {
        this.text += chunk.slice(start, this.prevI);
        this.fail("< not allowed in attribute value");
        // Recover: save what we have as attribute, re-process '<' as tag start
        this.attribList.push({ name: this.name, value: this.text });
        this.name = "";
        this.text = "";
        this.q = null;
        this.unget();
        this.state = S_TEXT;
        return;
      }
    }
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
        } else {
          this.fail("unexpected character in close tag");
        }
        return;
      }
      const first = chunk.charCodeAt(pos);
      if (first < 128 ? NAME_START_CHAR_ASCII[first] !== 1 : !isNameStartChar(first)) {
        this.i = pos;
        const c = this.getCode();
        if (c !== -1) {
          this.fail("unexpected character in close tag");
        }
        return;
      }
      // Consume the first character
      this.name = chunk[pos];
      pos++;
      if (this.trackPosition) {
        this.column++;
      }
    }

    // Fast scan remaining name chars
    const nameStart = pos;
    while (pos < chunk.length) {
      const code = chunk.charCodeAt(pos);
      if (code < 128 ? NAME_CHAR_ASCII[code] === 1 : isNameChar(code)) {
        pos++;
      } else {
        break;
      }
    }

    if (pos > nameStart) {
      this.name += chunk.slice(nameStart, pos);
    }
    if (this.trackPosition) {
      this.column += pos - nameStart;
    }

    if (pos >= chunk.length) {
      this.i = pos;
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
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === MINUS) {
      this.state = S_COMMENT_ENDING;
    } else {
      this.text += charFromCode(c);
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
      this.text += "-" + charFromCode(c);
      this.state = S_COMMENT;
    }
  }

  private sCommentEnded(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      // Emit comment event
      this._handlers.comment?.(this.text);
      this.text = "";
      this.state = S_TEXT;
    } else if (c === MINUS) {
      this.text += "-";
    } else {
      this.fail("malformed comment");
      this.text += "--" + charFromCode(c);
      this.state = S_COMMENT;
    }
  }

  private sCData(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === CLOSE_BRACKET) {
      this.state = S_CDATA_ENDING;
    } else {
      this.text += charFromCode(c);
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
      this.text += "]" + charFromCode(c);
      this.state = S_CDATA;
    }
  }

  private sCDataEnding2(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      // Emit CDATA event; also emit as text for backwards compat
      if (this._handlers.cdata) {
        this._handlers.cdata(this.text);
      } else if (this._handlers.text && this.text.length > 0) {
        this._handlers.text(this.text);
      }
      this.text = "";
      this.state = S_TEXT;
    } else if (c === CLOSE_BRACKET) {
      this.text += "]";
    } else {
      this.text += "]]" + charFromCode(c);
      this.state = S_CDATA;
    }
  }

  private sPI(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === QUESTION) {
      this.state = S_PI_ENDING;
    } else {
      this.text += charFromCode(c);
    }
  }

  private sPIEnding(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }
    if (c === GREATER) {
      // Emit PI event
      const piText = this.text;
      const spaceIdx = piText.indexOf(" ");
      if (spaceIdx >= 0) {
        this._handlers.pi?.(piText.slice(0, spaceIdx), piText.slice(spaceIdx + 1));
      } else {
        this._handlers.pi?.(piText, "");
      }
      this.text = "";
      this.state = S_TEXT;
    } else if (c === QUESTION) {
      this.text += "?";
    } else {
      this.text += "?" + charFromCode(c);
      this.state = S_PI;
    }
  }

  private sDoctype(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === GREATER) {
      this.text = "";
      this.state = S_TEXT;
    } else if (isQuote(c)) {
      this.q = c;
      this.state = S_DOCTYPE_QUOTE;
    } else if (c === OPEN_BRACKET) {
      this.state = S_DOCTYPE_DTD;
    } else {
      this.text += charFromCode(c);
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
    } else {
      this.text += charFromCode(c);
    }
  }

  private sDoctypeDTD(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === CLOSE_BRACKET) {
      this.state = S_DOCTYPE;
    } else if (isQuote(c)) {
      this.q = c;
      this.state = S_DOCTYPE_DTD_QUOTED;
    }
  }

  private sDoctypeDTDQuoted(): void {
    const c = this.getCode();
    if (c === -1) {
      return;
    }

    if (c === this.q) {
      this.q = null;
      this.state = S_DOCTYPE_DTD;
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

      this.text += resolved;
      this.state = this.entityReturnState;
      this.entity = "";
    } else if (isNameChar(c) || c === HASH) {
      this.entity += charFromCode(c);
      // Security: reject excessively long entity names (>64 chars)
      if (this.entity.length > 64) {
        this.fail("entity name too long");
        this.text += "&" + this.entity;
        this.state = this.entityReturnState;
        this.entity = "";
      }
    } else {
      this.fail("invalid entity character");
      this.text += "&" + this.entity + charFromCode(c);
      this.state = this.entityReturnState;
      this.entity = "";
    }
  }

  // ===========================================================================
  // Entity Resolution
  // ===========================================================================

  private parseEntity(entity: string): string {
    if (entity[0] !== "#") {
      // Named entity — resolve first, then check expansion limit for non-predefined
      const resolved = this.ENTITIES[entity];
      if (resolved !== undefined) {
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
              `entity expansion limit (${this.maxEntityExpansions}) exceeded — possible XML bomb`
            );
            return "";
          }
        }
        return resolved;
      }
      this.fail("undefined entity: " + entity);
      return "&" + entity + ";";
    }

    // Numeric character reference — validate range strictly
    let num: number;
    if (entity[1] === "x" || entity[1] === "X") {
      num = parseInt(entity.slice(2), 16);
    } else {
      num = parseInt(entity.slice(1), 10);
    }

    // Security: reject NaN, null (0x0), surrogates (0xD800-0xDFFF), out of range (>0x10FFFF)
    if (isNaN(num) || num < 1 || (num >= 0xd800 && num <= 0xdfff) || num > 0x10ffff) {
      this.fail("invalid character entity: &#" + entity.slice(1) + ";");
      return "";
    }

    // Also reject XML-invalid characters
    if (!isChar10(num)) {
      this.fail("invalid character entity");
      return "";
    }

    return String.fromCodePoint(num);
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
  private applyNamespaces(tag: SaxTag): void {
    // 1. Collect xmlns declarations from already-populated attributes
    const nsDecls: Record<string, string> = {};
    const attrs = tag.attributes;
    for (const name in attrs) {
      if (!Object.prototype.hasOwnProperty.call(attrs, name)) {
        continue;
      }
      if (name === "xmlns") {
        nsDecls[""] = attrs[name]; // default namespace
      } else if (name.startsWith("xmlns:")) {
        nsDecls[name.slice(6)] = attrs[name];
      }
    }

    // 2. Push namespace scope
    this._nsStack.push(nsDecls);

    // 3. Set namespace fields on the tag
    if (Object.keys(nsDecls).length > 0) {
      tag.ns = nsDecls;
    }
    const [prefix, local] = this.splitQName(tag.name);
    tag.prefix = prefix;
    tag.local = local;
    tag.uri = this.resolveNs(prefix);
  }

  // ===========================================================================
  // Tag Emission
  // ===========================================================================

  private openTag(): void {
    const tag = this.tag!;
    tag.isSelfClosing = false;

    for (const { name, value } of this.attribList) {
      tag.attributes[name] = value;
    }
    this.attribList = [];

    if (this.xmlns) {
      this.applyNamespaces(tag);
    }

    // Security: warn on nesting depth exceeded but continue processing
    // so that close tags still match (prevents infinite error loops)
    if (this.maxDepth > 0 && this.tags.length >= this.maxDepth) {
      this.fail(`maximum element nesting depth (${this.maxDepth}) exceeded`);
    }

    this._handlers.opentag?.(tag);
    this.tags.push(tag);
    this.name = "";
    this.state = S_TEXT;
  }

  private openSelfClosingTag(): void {
    const tag = this.tag!;
    tag.isSelfClosing = true;

    for (const { name, value } of this.attribList) {
      tag.attributes[name] = value;
    }
    this.attribList = [];

    if (this.xmlns) {
      this.applyNamespaces(tag);
    }

    // Security: warn on nesting depth exceeded but continue processing
    if (this.maxDepth > 0 && this.tags.length >= this.maxDepth) {
      this.fail(`maximum element nesting depth (${this.maxDepth}) exceeded`);
    }

    this._handlers.opentag?.(tag);
    this._handlers.closetag?.(tag);

    // Pop namespace scope for self-closing tags immediately
    if (this.xmlns) {
      this._nsStack.pop();
    }

    if (this.tags.length === 0) {
      this.closedRoot = true;
    }
    this.name = "";
    this.state = S_TEXT;
  }

  private closeTag(): void {
    const { tags, name } = this;
    this.state = S_TEXT;
    this.name = "";

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
          if (tags.length > i) {
            this.fail("unclosed tag: " + t.name);
          }
        }
        found = true;
        break;
      }
    }

    if (!found) {
      this.fail("unmatched close tag: " + name);
      this.text += "</" + name + ">";
    }

    if (tags.length === 0) {
      this.closedRoot = true;
    }
  }

  private end(): this {
    if (!this.sawRoot) {
      this.fail("document must contain a root element");
    }

    while (this.tags.length > 0) {
      const tag = this.tags.pop()!;
      this.fail("unclosed tag: " + tag.name);
    }

    if (this.text.length > 0 && this._handlers.text) {
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

/**
 * Parse an async-iterable of XML chunks as a stream of SAX event batches.
 *
 * Yields an array of {@link SaxEvent} per input chunk. This is the
 * primary integration point for streaming XML parsing.
 *
 * @param iterable - Async iterable of string or Uint8Array chunks.
 * @param options - Parser options.
 *
 * @example
 * ```ts
 * for await (const events of parseSax(stream)) {
 *   for (const event of events) {
 *     if (event.eventType === "opentag") console.log(event.value.name);
 *   }
 * }
 * ```
 */
async function* parseSax(
  iterable: AsyncIterable<any> | Iterable<any>,
  options?: SaxOptions
): AsyncGenerator<SaxEventAny[]> {
  const decoder = new TextDecoder();
  const parser = new SaxParser(options);

  let error: Error | undefined;
  parser.on("error", (err: Error) => {
    error = err;
  });

  let events: SaxEventAny[] = [];
  parser.on("opentag", value => events.push({ eventType: "opentag", value }));
  parser.on("text", value => events.push({ eventType: "text", value }));
  parser.on("closetag", value => events.push({ eventType: "closetag", value }));
  parser.on("cdata", value => events.push({ eventType: "cdata", value }));
  parser.on("comment", value => events.push({ eventType: "comment", value }));
  parser.on("pi", (target, body) => events.push({ eventType: "pi", value: { target, body } }));

  for await (const chunk of iterable) {
    const chunkStr = typeof chunk === "string" ? chunk : decoder.decode(chunk, { stream: true });
    parser.write(chunkStr);
    if (error) {
      throw error;
    }
    if (events.length > 0) {
      yield events;
      events = [];
    }
  }

  parser.close();
  if (error) {
    throw error;
  }
  if (events.length > 0) {
    yield events;
  }
}

export { SaxParser, parseSax };
