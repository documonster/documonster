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
 * Is `code` a legal XML 1.0 character per §2.2?
 *
 *   Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD]
 *          | [#x10000-#x10FFFF]
 *
 * Rejects NULs, the forbidden C0 controls (`0x01`-`0x08`, `0x0B`,
 * `0x0C`, `0x0E`-`0x1F`), lone surrogate halves (`0xD800`-`0xDFFF`),
 * the noncharacters `0xFFFE` / `0xFFFF`, and anything above the
 * Unicode ceiling.
 *
 * Used by both {@link xmlDecode} (to refuse malformed numeric
 * character references) and {@link encodeCData} (to strip illegal
 * bytes before wrapping user content in a CDATA section).
 */
function isLegalXmlChar(code: number): boolean {
  if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) {
    return false;
  }
  if (code === 0x9 || code === 0xa || code === 0xd) {
    return true;
  }
  if (code < 0x20) {
    return false;
  }
  if (code >= 0xd800 && code <= 0xdfff) {
    return false;
  }
  if (code === 0xfffe || code === 0xffff) {
    return false;
  }
  return true;
}

/**
 * Decode XML entities in a string.
 *
 * Handles named entities (`&lt;`, `&gt;`, `&amp;`, `&quot;`, `&apos;`)
 * and numeric character references (`&#123;`, `&#x7B;`).
 *
 * Security: validates numeric code points against the XML 1.0 `Char`
 * production (rejects NUL, forbidden C0 controls, surrogates,
 * noncharacters, and out-of-range code points). Malformed numeric
 * refs are left untouched in the output so downstream layers (e.g. a
 * re-encoder that strips them) can distinguish "author meant this" from
 * "we couldn't decode this".
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
      // Refuse everything outside the XML 1.0 Char range — NUL,
      // forbidden controls, surrogates, noncharacters, and values
      // above `0x10FFFF`. Leaving the ref untouched is safer than
      // materialising an invalid character into the decoded string.
      if (!isLegalXmlChar(code)) {
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
 * Encode a value for use in a double-quoted XML attribute.
 *
 * Does everything {@link xmlEncode} does (escape `<>&"'`, strip invalid
 * control chars and lone surrogates) PLUS encodes `\t`, `\n`, `\r` as
 * numeric character references (`&#x9;`, `&#xA;`, `&#xD;`).
 *
 * The extra whitespace handling is required by XML 1.0 §3.3.3
 * (attribute-value normalisation), which replaces every literal
 * whitespace character inside an attribute value with a single space
 * at parse time. Without the numeric-ref encoding, `"foo\nbar"` written
 * into an attribute round-trips as `"foo bar"` — the newline is gone.
 * Parsers / validators never collapse numeric character references, so
 * `&#xA;` survives verbatim.
 *
 * Use this for every attribute value in XML output that needs to
 * round-trip exactly (OOXML chart / sidecar / pivot content, relationship
 * targets, anything later re-parsed by another tool).
 */
export function xmlEncodeAttr(value: string): string {
  // Fast path: common case has none of `\t \n \r` so the expensive
  // replace chain doesn't run. `xmlEncode` handles `<>&"'` and invalid
  // control chars; whitespace characters fall through it unchanged.
  const encoded = xmlEncode(value);
  if (
    encoded.indexOf("\t") === -1 &&
    encoded.indexOf("\n") === -1 &&
    encoded.indexOf("\r") === -1
  ) {
    return encoded;
  }
  return encoded.replace(/\r/g, "&#xD;").replace(/\n/g, "&#xA;").replace(/\t/g, "&#x9;");
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
 * Encode text for a CDATA section, stripping XML 1.0-illegal
 * characters first and then splitting on `]]>` to produce valid
 * output.
 *
 * CDATA is not a magic passthrough: only the five structural entities
 * are skipped, but every other XML Char production rule still applies.
 * A user string that carries `\x08` or a lone surrogate half wraps
 * into a CDATA section that every conformant parser rejects. Strip
 * them first (same sanitisation as {@link xmlEncode} / the chart
 * module's `escapeXml`) so CDATA output stays well-formed.
 *
 * The sequence `]]>` cannot appear inside CDATA, so each occurrence is split
 * into adjacent CDATA sections: `<![CDATA[...]]]]><![CDATA[>...]]>`.
 */
export function encodeCData(text: string): string {
  return "<![CDATA[" + stripIllegalXmlChars(text).split("]]>").join("]]]]><![CDATA[>") + "]]>";
}

/**
 * Strip characters that XML 1.0 forbids from text / attribute /
 * CDATA content, plus DEL (0x7F) as a project-policy extension.
 *
 * The XML 1.0 Char production allows: `#x9 | #xA | #xD |
 * [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]`. So the
 * disallowed ranges are:
 *   - C0 controls other than `\t` `\n` `\r` (`0x00-0x08`, `0x0B`,
 *     `0x0C`, `0x0E-0x1F`),
 *   - lone UTF-16 surrogate halves (`0xD800-0xDFFF` not in a valid
 *     pair),
 *   - the noncharacters `0xFFFE` / `0xFFFF`.
 *
 * `0x7F` (DEL) is TECHNICALLY legal per the XML spec — it falls
 * within the `[#x20-#xD7FF]` range — but some downstream consumers
 * (older Excel versions, strict SAX libraries) choke on it. We strip
 * it as a defence-in-depth measure; the comment that previously
 * claimed "XML 1.0 forbids DEL" was incorrect.
 *
 * Exported so `chart-utils.ts` and any other module that needs the
 * same sanitisation can share a single implementation instead of
 * carrying a local copy that drifts on edge cases.
 */
export function stripXmlIllegalChars(text: string): string {
  // Fast path: scan for any char that would fail the Char production.
  // Most user input is plain ASCII in the legal range, so this
  // returns the original string without allocation.
  let needsStrip = false;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (
      (code < 0x20 && code !== 9 && code !== 10 && code !== 13) ||
      code === 0x7f ||
      (code >= 0xd800 && code <= 0xdfff) ||
      code === 0xfffe ||
      code === 0xffff
    ) {
      needsStrip = true;
      break;
    }
  }
  if (!needsStrip) {
    return text;
  }
  const out: string[] = [];
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code < 0x20) {
      if (code === 9 || code === 10 || code === 13) {
        out.push(text[i]);
      }
      continue;
    }
    if (code === 0x7f) {
      continue;
    }
    if (code === 0xfffe || code === 0xffff) {
      continue;
    }
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out.push(text[i], text[i + 1]);
        i += 1;
      }
      // else: lone high surrogate — strip
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) {
      continue;
    }
    out.push(text[i]);
  }
  return out.join("");
}

function stripIllegalXmlChars(text: string): string {
  return stripXmlIllegalChars(text);
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
