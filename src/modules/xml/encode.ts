/**
 * XML Encoding / Decoding Utilities
 *
 * Self-contained XML entity encoding and decoding functions.
 */

import { XmlError } from "@xml/errors";

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
 * Lookup table for characters that need encoding in the ASCII range (0-127).
 * 0 = safe, 1 = encode to entity, 2 = strip (invalid control char)
 */
const ENCODE_ACTION = /* @__PURE__ */ (() => {
  const t = new Uint8Array(128);
  // Invalid control chars → strip
  for (let i = 0; i <= 0x08; i++) {
    t[i] = 2;
  }
  t[0x0b] = 2;
  t[0x0c] = 2;
  for (let i = 0x0e; i <= 0x1f; i++) {
    t[i] = 2;
  }
  t[0x7f] = 2; // DEL
  // Entity-encode chars
  t[0x22] = 1; // "
  t[0x26] = 1; // &
  t[0x27] = 1; // '
  t[0x3c] = 1; // <
  t[0x3e] = 1; // >
  return t;
})();

const ENCODE_ENTITIES: Record<number, string> = {
  0x22: "&quot;",
  0x26: "&amp;",
  0x27: "&apos;",
  0x3c: "&lt;",
  0x3e: "&gt;"
};

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
 * Strips invalid XML control characters (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F, 0x7F)
 * and lone surrogates (0xD800-0xDFFF without a pair).
 *
 * Optimized: uses a lookup table and manual scan instead of regex for
 * maximum throughput on the hot path (called per attribute/text value).
 */
export function xmlEncode(text: string): string {
  const len = text.length;
  // Fast scan: find first character that needs encoding
  let firstBad = -1;
  for (let i = 0; i < len; i++) {
    const code = text.charCodeAt(i);
    if (code < 128) {
      if (ENCODE_ACTION[code] !== 0) {
        firstBad = i;
        break;
      }
    } else if (code >= 0xd800 && code <= 0xdfff) {
      // Check for lone surrogate
      if (code <= 0xdbff) {
        const next = text.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) {
          i++; // valid pair, skip low surrogate
          continue;
        }
      }
      firstBad = i;
      break;
    } else if (code === 0xfffe || code === 0xffff) {
      // XML noncharacters — must strip
      firstBad = i;
      break;
    }
  }

  if (firstBad === -1) {
    return text; // fast path: nothing to encode
  }

  // Slow path: array + join (V8-optimized pattern)
  const parts: string[] = [];
  let lastIndex = 0;

  for (let i = firstBad; i < len; i++) {
    const code = text.charCodeAt(i);

    if (code < 128) {
      const action = ENCODE_ACTION[code];
      if (action === 0) {
        continue;
      }
      // Flush safe segment
      if (lastIndex < i) {
        parts.push(text.substring(lastIndex, i));
      }
      if (action === 1) {
        parts.push(ENCODE_ENTITIES[code]);
      }
      // action === 2: strip
      lastIndex = i + 1;
    } else if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        i++; // valid pair, keep going
        continue;
      }
      // Lone high surrogate — strip
      if (lastIndex < i) {
        parts.push(text.substring(lastIndex, i));
      }
      lastIndex = i + 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      // Lone low surrogate — strip
      if (lastIndex < i) {
        parts.push(text.substring(lastIndex, i));
      }
      lastIndex = i + 1;
    } else if (code === 0xfffe || code === 0xffff) {
      // XML noncharacters — strip
      if (lastIndex < i) {
        parts.push(text.substring(lastIndex, i));
      }
      lastIndex = i + 1;
    }
  }

  if (lastIndex < len) {
    parts.push(text.substring(lastIndex));
  }

  return parts.length === 1 ? parts[0] : parts.join("");
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

// =============================================================================
// Name Validation
// =============================================================================

/**
 * Characters that must NEVER appear in XML element or attribute names.
 * This is a fast security check to prevent markup injection via names,
 * not a full XML NameChar validation (which would require Unicode tables).
 */
const INVALID_NAME_CHARS = /[\s<>"'/=&]/;

/**
 * Validate an XML element or attribute name against injection attacks.
 *
 * Rejects:
 * - Empty names
 * - Names containing whitespace, `<`, `>`, `"`, `'`, `/`, `=`, `&`
 * - Names starting with a digit, `-`, or `.`
 *
 * This is NOT a full XML Name validation (which requires Unicode NameStartChar
 * tables). It is a focused security check to prevent markup injection.
 */
export function validateXmlName(name: string): void {
  if (!name) {
    throw new XmlError("XML name must not be empty");
  }
  if (INVALID_NAME_CHARS.test(name)) {
    throw new XmlError(`Invalid XML name: contains forbidden character in "${name}"`);
  }
  // XML names cannot start with a digit, hyphen, or dot
  const first = name.charCodeAt(0);
  if ((first >= 0x30 && first <= 0x39) || first === 0x2d || first === 0x2e) {
    throw new XmlError(`Invalid XML name: "${name}" starts with forbidden character`);
  }
}

// =============================================================================
// Writer Helpers
// =============================================================================

/**
 * Encode text for a CDATA section, splitting on `]]>` to produce valid output.
 *
 * The sequence `]]>` cannot appear inside CDATA, so each occurrence is split
 * into adjacent CDATA sections: `<![CDATA[...]]]]><![CDATA[>...]]>`.
 */
export function encodeCData(text: string): string {
  return "<![CDATA[" + text.split("]]>").join("]]]]><![CDATA[>") + "]]>";
}

/**
 * Validate that text is legal for an XML comment.
 *
 * XML spec: comments must not contain `--` and must not end with `-`.
 * @throws {XmlError} if the text is invalid.
 */
export function validateCommentText(text: string): void {
  if (text.includes("--") || text.endsWith("-")) {
    throw new XmlError('Invalid comment: must not contain "--" or end with "-"');
  }
}

// =============================================================================
// Standard XML Declaration
// =============================================================================

/** Default XML declaration attributes (`version`, `encoding`, `standalone`). */
export const StdDocAttributes: Readonly<Record<string, string>> = {
  version: "1.0",
  encoding: "UTF-8",
  standalone: "yes"
};
