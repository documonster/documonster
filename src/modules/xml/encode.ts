/**
 * XML Encoding / Decoding Utilities
 *
 * Self-contained XML entity encoding and decoding functions.
 * No dependencies on other modules — safe to import from anywhere.
 */

// =============================================================================
// Constants
// =============================================================================

/** Standard XML entity decode map. */
const DECODE_MAP: Record<string, string> = {
  lt: "<",
  gt: ">",
  amp: "&",
  quot: '"',
  apos: "'"
};

/** Regex for decoding XML entities (named + numeric). */
const DECODE_RE = /&(#\d+|#[xX][0-9A-Fa-f]+|\w+);/g;

/**
 * Regex that detects the first character requiring encoding.
 * Matches: < > & " ' DEL, invalid control characters, and lone surrogates.
 * Uses negative lookahead/lookbehind to avoid matching valid surrogate pairs.
 */
/* oxlint-disable no-control-regex -- Control characters are intentionally matched */
const ENCODE_DETECT_RE =
  /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

const ENCODE_ALL_RE =
  /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
/* oxlint-enable no-control-regex */

// =============================================================================
// Decode
// =============================================================================

/**
 * Decode XML entities in a string.
 *
 * Handles named entities (`&lt;`, `&gt;`, `&amp;`, `&quot;`, `&apos;`)
 * and numeric character references (`&#123;`, `&#x7B;`).
 *
 * Security: validates numeric code points are in range [1, 0x10FFFF]
 * and rejects surrogate halves (0xD800-0xDFFF).
 *
 * Fast-path: returns the original string if no `&` is found.
 */
export function xmlDecode(text: string): string {
  if (text.indexOf("&") === -1) {
    return text;
  }
  return text.replace(DECODE_RE, (match: string, entity: string) => {
    if (entity[0] === "#") {
      const code =
        entity[1] === "x" || entity[1] === "X"
          ? parseInt(entity.slice(2), 16)
          : parseInt(entity.slice(1), 10);
      // Reject NaN, null char, surrogates, and out-of-range code points
      if (Number.isNaN(code) || code < 1 || (code >= 0xd800 && code <= 0xdfff) || code > 0x10ffff) {
        return match;
      }
      return String.fromCodePoint(code);
    }
    return DECODE_MAP[entity] ?? match;
  });
}

// =============================================================================
// Encode
// =============================================================================

/**
 * Encode special characters for XML output.
 *
 * Escapes `<`, `>`, `&`, `"`, `'` to their entity equivalents.
 * Strips invalid XML control characters (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F, 0x7F).
 *
 * Fast-path: returns the original string if no special characters are found.
 */
export function xmlEncode(text: string): string {
  if (!ENCODE_DETECT_RE.test(text)) {
    return text;
  }

  return text.replace(ENCODE_ALL_RE, ch => {
    const code = ch.charCodeAt(0);
    switch (code) {
      case 34:
        return "&quot;";
      case 38:
        return "&amp;";
      case 39:
        return "&apos;";
      case 60:
        return "&lt;";
      case 62:
        return "&gt;";
      default:
        // Strip invalid control characters and DEL
        return "";
    }
  });
}

// =============================================================================
// Attribute Encoding
// =============================================================================

/**
 * Encode a value for use in an XML attribute.
 *
 * Same as {@link xmlEncode} — provided as a semantic alias.
 * In the future this could apply attribute-specific normalisation
 * (e.g. collapsing whitespace per XML 1.0 §3.3.3).
 */
export function xmlEncodeAttr(value: string): string {
  return xmlEncode(value);
}
