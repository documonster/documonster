/**
 * PDF tokenizer / lexer.
 *
 * Scans raw PDF bytes and produces a stream of typed tokens.
 * Handles all PDF token types: numbers, strings (literal and hex),
 * names, booleans, null, keywords, and delimiters.
 *
 * @see PDF Reference 1.7, §3.1 - Lexical Conventions
 */

import { PdfStructureError } from "@pdf/errors";

// =============================================================================
// Token Types
// =============================================================================

export const enum TokenType {
  /** Integer or real number */
  Number,
  /** Literal string delimited by parentheses `(...)` */
  LiteralString,
  /** Hex string delimited by angle brackets `<...>` */
  HexString,
  /** Name object starting with `/` */
  Name,
  /** Boolean `true` or `false` */
  Boolean,
  /** The `null` keyword */
  Null,
  /** Keywords: obj, endobj, stream, endstream, xref, trailer, startxref, R */
  Keyword,
  /** `<<` dict begin */
  DictBegin,
  /** `>>` dict end */
  DictEnd,
  /** `[` array begin */
  ArrayBegin,
  /** `]` array end */
  ArrayEnd,
  /** End of input */
  EOF
}

export interface Token {
  type: TokenType;
  /** Numeric value (for Number tokens) */
  numValue?: number;
  /** String value (for String, Name, Keyword, Boolean tokens) */
  strValue?: string;
  /** Raw bytes (for LiteralString and HexString tokens) */
  rawBytes?: Uint8Array;
  /** Boolean value (for Boolean tokens) */
  boolValue?: boolean;
  /** Byte offset where this token starts */
  offset: number;
}

// =============================================================================
// Character Classification
// =============================================================================

/** Whitespace bytes per PDF spec §3.1 */
function isWhitespace(b: number): boolean {
  return b === 0x00 || b === 0x09 || b === 0x0a || b === 0x0d || b === 0x0c || b === 0x20;
}

/** Delimiter bytes per PDF spec §3.1 */
function isDelimiter(b: number): boolean {
  return (
    b === 0x28 || // (
    b === 0x29 || // )
    b === 0x3c || // <
    b === 0x3e || // >
    b === 0x5b || // [
    b === 0x5d || // ]
    b === 0x7b || // {
    b === 0x7d || // }
    b === 0x2f || // /
    b === 0x25 // %
  );
}

function isDigit(b: number): boolean {
  return b >= 0x30 && b <= 0x39;
}

function isHexDigit(b: number): boolean {
  return (b >= 0x30 && b <= 0x39) || (b >= 0x41 && b <= 0x46) || (b >= 0x61 && b <= 0x66);
}

function hexVal(b: number): number {
  if (b >= 0x30 && b <= 0x39) {
    return b - 0x30;
  }
  if (b >= 0x41 && b <= 0x46) {
    return b - 0x41 + 10;
  }
  return b - 0x61 + 10;
}

// =============================================================================
// Cached Constants
// =============================================================================

/** Pre-encoded "endstream" bytes to avoid repeated TextEncoder allocations */
const ENDSTREAM_BYTES = new Uint8Array([101, 110, 100, 115, 116, 114, 101, 97, 109]); // "endstream"

// =============================================================================
// PDF Tokenizer
// =============================================================================

/**
 * Byte-level PDF tokenizer.
 *
 * Provides a `next()` method that returns the next token from the input.
 * The tokenizer maintains a mutable position pointer that advances through
 * the input bytes.
 */
export class PdfTokenizer {
  private data: Uint8Array;
  declare pos: number;

  constructor(data: Uint8Array, offset = 0) {
    this.data = data;
    this.pos = offset;
  }

  /** Get current position */
  get position(): number {
    return this.pos;
  }

  /** Set current position */
  set position(offset: number) {
    this.pos = offset;
  }

  /** Get the underlying data */
  get bytes(): Uint8Array {
    return this.data;
  }

  /** Peek at the byte at the current position without consuming it */
  peek(): number {
    return this.pos < this.data.length ? this.data[this.pos] : -1;
  }

  /** Read the next token */
  next(): Token {
    this.skipWhitespaceAndComments();

    if (this.pos >= this.data.length) {
      return { type: TokenType.EOF, offset: this.pos };
    }

    const startPos = this.pos;
    const b = this.data[this.pos];

    // Literal string
    if (b === 0x28) {
      return this.readLiteralString(startPos);
    }

    // Hex string or dict delimiter
    if (b === 0x3c) {
      if (this.pos + 1 < this.data.length && this.data[this.pos + 1] === 0x3c) {
        this.pos += 2;
        return { type: TokenType.DictBegin, offset: startPos };
      }
      return this.readHexString(startPos);
    }

    // Dict end
    if (b === 0x3e) {
      if (this.pos + 1 < this.data.length && this.data[this.pos + 1] === 0x3e) {
        this.pos += 2;
        return { type: TokenType.DictEnd, offset: startPos };
      }
      // Stray > — treat as keyword
      this.pos++;
      return { type: TokenType.Keyword, strValue: ">", offset: startPos };
    }

    // Array delimiters
    if (b === 0x5b) {
      this.pos++;
      return { type: TokenType.ArrayBegin, offset: startPos };
    }
    if (b === 0x5d) {
      this.pos++;
      return { type: TokenType.ArrayEnd, offset: startPos };
    }

    // Name
    if (b === 0x2f) {
      return this.readName(startPos);
    }

    // Number (digit, sign, or decimal point)
    if (isDigit(b) || b === 0x2d || b === 0x2b || b === 0x2e) {
      return this.readNumber(startPos);
    }

    // Regular character — keyword (obj, endobj, true, false, null, etc.)
    return this.readKeyword(startPos);
  }

  // ===========================================================================
  // Skip Whitespace and Comments
  // ===========================================================================

  skipWhitespaceAndComments(): void {
    while (this.pos < this.data.length) {
      const b = this.data[this.pos];
      if (isWhitespace(b)) {
        this.pos++;
        continue;
      }
      // PDF comment: % ... EOL
      if (b === 0x25) {
        this.pos++;
        while (this.pos < this.data.length) {
          const c = this.data[this.pos];
          if (c === 0x0a || c === 0x0d) {
            break;
          }
          this.pos++;
        }
        continue;
      }
      break;
    }
  }

  // ===========================================================================
  // Literal String (...)
  // ===========================================================================

  private readLiteralString(startPos: number): Token {
    this.pos++; // skip opening '('
    const bytes: number[] = [];
    let depth = 1;

    while (this.pos < this.data.length && depth > 0) {
      const b = this.data[this.pos];

      if (b === 0x5c) {
        // Backslash escape
        this.pos++;
        if (this.pos >= this.data.length) {
          break;
        }
        const esc = this.data[this.pos];
        switch (esc) {
          case 0x6e:
            bytes.push(0x0a);
            break; // \n
          case 0x72:
            bytes.push(0x0d);
            break; // \r
          case 0x74:
            bytes.push(0x09);
            break; // \t
          case 0x62:
            bytes.push(0x08);
            break; // \b
          case 0x66:
            bytes.push(0x0c);
            break; // \f
          case 0x28:
            bytes.push(0x28);
            break; // \(
          case 0x29:
            bytes.push(0x29);
            break; // \)
          case 0x5c:
            bytes.push(0x5c);
            break; // \\
          case 0x0a: // \<LF> — line continuation
            break;
          case 0x0d: // \<CR> or \<CR><LF> — line continuation
            if (this.pos + 1 < this.data.length && this.data[this.pos + 1] === 0x0a) {
              this.pos++;
            }
            break;
          default:
            // Octal escape: up to 3 octal digits
            if (esc >= 0x30 && esc <= 0x37) {
              let octal = esc - 0x30;
              if (
                this.pos + 1 < this.data.length &&
                this.data[this.pos + 1] >= 0x30 &&
                this.data[this.pos + 1] <= 0x37
              ) {
                this.pos++;
                octal = octal * 8 + (this.data[this.pos] - 0x30);
                if (
                  this.pos + 1 < this.data.length &&
                  this.data[this.pos + 1] >= 0x30 &&
                  this.data[this.pos + 1] <= 0x37
                ) {
                  this.pos++;
                  octal = octal * 8 + (this.data[this.pos] - 0x30);
                }
              }
              bytes.push(octal & 0xff);
            } else {
              // Unknown escape — just include the character
              bytes.push(esc);
            }
            break;
        }
        this.pos++;
      } else if (b === 0x28) {
        // Nested (
        depth++;
        bytes.push(b);
        this.pos++;
      } else if (b === 0x29) {
        // Closing )
        depth--;
        if (depth > 0) {
          bytes.push(b);
        }
        this.pos++;
      } else {
        // Normalize line endings: \r\n or \r → \n
        if (b === 0x0d) {
          bytes.push(0x0a);
          this.pos++;
          if (this.pos < this.data.length && this.data[this.pos] === 0x0a) {
            this.pos++;
          }
        } else {
          bytes.push(b);
          this.pos++;
        }
      }
    }

    const rawBytes = new Uint8Array(bytes);
    return { type: TokenType.LiteralString, rawBytes, offset: startPos };
  }

  // ===========================================================================
  // Hex String <...>
  // ===========================================================================

  private readHexString(startPos: number): Token {
    this.pos++; // skip opening '<'
    const hexBytes: number[] = [];
    let highNibble = -1;

    while (this.pos < this.data.length) {
      const b = this.data[this.pos];
      if (b === 0x3e) {
        this.pos++; // skip closing '>'
        break;
      }
      if (isWhitespace(b)) {
        this.pos++;
        continue;
      }
      if (isHexDigit(b)) {
        if (highNibble < 0) {
          highNibble = hexVal(b);
        } else {
          hexBytes.push((highNibble << 4) | hexVal(b));
          highNibble = -1;
        }
      }
      this.pos++;
    }

    // Odd number of hex digits — pad with 0
    if (highNibble >= 0) {
      hexBytes.push(highNibble << 4);
    }

    const rawBytes = new Uint8Array(hexBytes);
    return { type: TokenType.HexString, rawBytes, offset: startPos };
  }

  // ===========================================================================
  // Name /...
  // ===========================================================================

  private readName(startPos: number): Token {
    this.pos++; // skip '/'
    const chars: number[] = [];

    while (this.pos < this.data.length) {
      const b = this.data[this.pos];
      if (isWhitespace(b) || isDelimiter(b)) {
        break;
      }
      if (b === 0x23 && this.pos + 2 < this.data.length) {
        // #XX escape
        const h1 = this.data[this.pos + 1];
        const h2 = this.data[this.pos + 2];
        if (isHexDigit(h1) && isHexDigit(h2)) {
          chars.push((hexVal(h1) << 4) | hexVal(h2));
          this.pos += 3;
          continue;
        }
      }
      chars.push(b);
      this.pos++;
    }

    const name = String.fromCharCode(...chars);
    return { type: TokenType.Name, strValue: name, offset: startPos };
  }

  // ===========================================================================
  // Number
  // ===========================================================================

  private readNumber(startPos: number): Token {
    let numStr = "";
    const first = this.data[this.pos];

    // Sign
    if (first === 0x2d || first === 0x2b) {
      numStr += String.fromCharCode(first);
      this.pos++;
    }

    let hasDecimal = first === 0x2e;
    if (hasDecimal) {
      numStr += ".";
      this.pos++;
    }

    while (this.pos < this.data.length) {
      const b = this.data[this.pos];
      if (isDigit(b)) {
        numStr += String.fromCharCode(b);
        this.pos++;
      } else if (b === 0x2e && !hasDecimal) {
        hasDecimal = true;
        numStr += ".";
        this.pos++;
      } else {
        break;
      }
    }

    // Edge case: just a sign with no digits — treat as keyword
    if (numStr === "+" || numStr === "-" || numStr === "." || numStr === "") {
      return this.readKeyword(startPos);
    }

    const numValue = hasDecimal ? parseFloat(numStr) : parseInt(numStr, 10);
    return { type: TokenType.Number, numValue, offset: startPos };
  }

  // ===========================================================================
  // Keyword
  // ===========================================================================

  private readKeyword(startPos: number): Token {
    const chars: number[] = [];

    while (this.pos < this.data.length) {
      const b = this.data[this.pos];
      if (isWhitespace(b) || isDelimiter(b)) {
        break;
      }
      chars.push(b);
      this.pos++;
    }

    const word = String.fromCharCode(...chars);

    if (word === "true") {
      return { type: TokenType.Boolean, boolValue: true, strValue: "true", offset: startPos };
    }
    if (word === "false") {
      return { type: TokenType.Boolean, boolValue: false, strValue: "false", offset: startPos };
    }
    if (word === "null") {
      return { type: TokenType.Null, offset: startPos };
    }

    return { type: TokenType.Keyword, strValue: word, offset: startPos };
  }

  // ===========================================================================
  // Utility: Find a byte sequence
  // ===========================================================================

  /**
   * Search forward for a byte sequence starting from the current position.
   * Returns the offset where the sequence starts, or -1 if not found.
   * Does NOT advance the position.
   */
  findSequence(seq: Uint8Array, from?: number): number {
    const start = from ?? this.pos;
    const len = seq.length;
    const end = this.data.length - len;

    for (let i = start; i <= end; i++) {
      let match = true;
      for (let j = 0; j < len; j++) {
        if (this.data[i + j] !== seq[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        return i;
      }
    }
    return -1;
  }

  /**
   * Search backward for a byte sequence starting from `from` (or end of data).
   * Returns the offset where the sequence starts, or -1 if not found.
   */
  findSequenceBackward(seq: Uint8Array, from?: number): number {
    const start = from ?? this.data.length - 1;
    const len = seq.length;

    for (let i = start - len + 1; i >= 0; i--) {
      let match = true;
      for (let j = 0; j < len; j++) {
        if (this.data[i + j] !== seq[j]) {
          match = false;
          break;
        }
      }
      if (match) {
        return i;
      }
    }
    return -1;
  }

  /**
   * Read a line of text at the current position. Advances past the line ending.
   */
  readLine(): string {
    const chars: number[] = [];
    while (this.pos < this.data.length) {
      const b = this.data[this.pos];
      this.pos++;
      if (b === 0x0a) {
        break;
      }
      if (b === 0x0d) {
        if (this.pos < this.data.length && this.data[this.pos] === 0x0a) {
          this.pos++;
        }
        break;
      }
      chars.push(b);
    }
    return String.fromCharCode(...chars);
  }

  /**
   * Extract a slice of the underlying data.
   */
  slice(start: number, end: number): Uint8Array {
    return this.data.subarray(start, end);
  }

  /**
   * Read the stream content following a `stream` keyword.
   * The tokenizer should be positioned right after the `stream` keyword.
   * Returns the raw stream bytes (between stream\n and endstream).
   */
  readStreamContent(length: number): Uint8Array {
    // Skip the EOL after "stream" keyword
    if (this.pos < this.data.length && this.data[this.pos] === 0x0d) {
      this.pos++;
    }
    if (this.pos < this.data.length && this.data[this.pos] === 0x0a) {
      this.pos++;
    }

    if (length < 0) {
      // Length unknown — search for endstream
      const endPos = this.findSequence(ENDSTREAM_BYTES, this.pos);
      if (endPos < 0) {
        throw new PdfStructureError("Could not find endstream marker");
      }
      let streamEnd = endPos;
      // Strip trailing EOL before endstream
      if (streamEnd > this.pos && this.data[streamEnd - 1] === 0x0a) {
        streamEnd--;
      }
      if (streamEnd > this.pos && this.data[streamEnd - 1] === 0x0d) {
        streamEnd--;
      }
      const content = this.data.subarray(this.pos, streamEnd);
      this.pos = endPos + ENDSTREAM_BYTES.length;
      return content;
    }

    const content = this.data.subarray(this.pos, this.pos + length);
    this.pos += length;

    // Skip to endstream
    this.skipWhitespaceAndComments();
    // Try to consume "endstream" keyword
    if (this.pos + ENDSTREAM_BYTES.length <= this.data.length) {
      let match = true;
      for (let i = 0; i < ENDSTREAM_BYTES.length; i++) {
        if (this.data[this.pos + i] !== ENDSTREAM_BYTES[i]) {
          match = false;
          break;
        }
      }
      if (match) {
        this.pos += ENDSTREAM_BYTES.length;
      }
    }

    return content;
  }
}
