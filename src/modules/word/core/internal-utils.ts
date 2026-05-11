/**
 * DOCX Module - Internal Utilities (unified shared helpers)
 *
 * This is the single source of truth for cross-cutting low-level utilities
 * used throughout the word module. Consolidates previously duplicated
 * implementations of base64, UUID, UTF-16LE encoding, and XML helpers.
 */

// =============================================================================
// Shared TextEncoder / TextDecoder singletons
// =============================================================================

/**
 * Shared `TextEncoder` instance for UTF-8 encoding. Reusing this avoids the
 * cost of constructing a fresh encoder on every call site (each `new
 * TextEncoder()` allocates a small native object).
 */
export const utf8Encoder = /*#__PURE__*/ new TextEncoder();

/**
 * Shared `TextDecoder` instance for UTF-8 decoding. Stateless — safe to share
 * across modules. Use this instead of `new TextDecoder("utf-8")`.
 */
export const utf8Decoder = /*#__PURE__*/ new TextDecoder("utf-8");

// =============================================================================
// Base64 Encoding / Decoding
// =============================================================================

/**
 * Encode a byte array to base64. Uses native Node `Buffer` if available;
 * otherwise falls back to chunked `btoa()` for browser environments.
 *
 * The chunked approach avoids the O(n²) string concatenation cost that naïve
 * `String.fromCharCode(...data)` implementations exhibit on large buffers.
 */
export function bytesToBase64(data: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(data).toString("base64");
  }
  const CHUNK = 0x8000;
  const parts: string[] = [];
  for (let i = 0; i < data.length; i += CHUNK) {
    const slice = data.subarray(i, i + CHUNK);
    // `apply` on Uint8Array works in all modern JS engines
    parts.push(String.fromCharCode.apply(null, slice as unknown as number[]));
  }
  return btoa(parts.join(""));
}

/**
 * Decode a base64 string to a byte array. Uses native Node `Buffer` if available.
 */
export function base64ToBytes(s: string): Uint8Array {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(s, "base64"));
  }
  const binary = atob(s);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// =============================================================================
// UUID v4 Generation
// =============================================================================

/**
 * Generate a random UUID v4 string in the form "XXXXXXXX-XXXX-4XXX-YXXX-XXXXXXXXXXXX".
 * Uses `crypto.getRandomValues` for cryptographic randomness.
 */
export function generateUuid(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);

  // RFC 4122 v4 UUID: set version (4) and variant (10xx) bits
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;

  const hex = Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
  return `${hex.substring(0, 8)}-${hex.substring(8, 12)}-${hex.substring(12, 16)}-${hex.substring(16, 20)}-${hex.substring(20, 32)}`;
}

/**
 * Generate a GUID in braced format: "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}".
 * Used by font obfuscation (w:fontKey) and building blocks.
 */
export function generateGuid(): string {
  return `{${generateUuid()}}`;
}

// =============================================================================
// UTF-16LE Encoding
// =============================================================================

/**
 * Encode a string as UTF-16LE bytes.
 * Required by ECMA-376 password hashing and OLE operations.
 */
export function stringToUtf16LE(str: string): Uint8Array {
  const buf = new Uint8Array(str.length * 2);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    buf[i * 2] = code & 0xff;
    buf[i * 2 + 1] = (code >> 8) & 0xff;
  }
  return buf;
}

/**
 * Decode UTF-16LE bytes to a string.
 */
export function utf16LEToString(bytes: Uint8Array): string {
  let result = "";
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    result += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
  }
  return result;
}

// =============================================================================
// XML Utilities
// =============================================================================

/**
 * Build an XML attributes object from an array of [name, value] entries,
 * omitting entries whose value is `undefined` or `false`. Numeric and boolean
 * values are stringified: `number → String(n)`, `boolean → "1"`.
 *
 * Reduces repetitive `if (x !== undefined) attrs["w:x"] = String(x)` patterns.
 */
export function buildAttrs(
  entries: ReadonlyArray<[string, string | number | boolean | undefined | null]>
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of entries) {
    if (v === undefined || v === null || v === false) {
      continue;
    }
    out[k] = typeof v === "boolean" ? "1" : String(v);
  }
  return out;
}

/**
 * Escape XML special characters in a string.
 * Prevents XML injection when building XML via string concatenation.
 */
export function escapeXml(str: string): string {
  let result = "";
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    switch (ch) {
      case "&":
        result += "&amp;";
        break;
      case "<":
        result += "&lt;";
        break;
      case ">":
        result += "&gt;";
        break;
      case '"':
        result += "&quot;";
        break;
      case "'":
        result += "&apos;";
        break;
      default:
        result += ch;
        break;
    }
  }
  return result;
}

// =============================================================================
// Random Bytes
// =============================================================================

/**
 * Generate cryptographically secure random bytes.
 * Throws if crypto.getRandomValues is unavailable.
 */
export function randomBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  if (typeof globalThis.crypto !== "undefined" && globalThis.crypto.getRandomValues) {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    throw new Error(
      "crypto.getRandomValues is required. " +
        "This environment does not provide a cryptographically secure random number generator."
    );
  }
  return bytes;
}

// =============================================================================
// Type Utilities
// =============================================================================

/**
 * Make all properties of T mutable (remove readonly).
 * Used in mutation APIs (acceptAllRevisions, replaceText, etc.) that operate
 * on document models in-place. Preferred over `as any` casts.
 */
export type Mutable<T> = { -readonly [P in keyof T]: T[P] };

/**
 * Deep mutable version — recursively removes readonly from nested objects.
 */
export type DeepMutable<T> = {
  -readonly [P in keyof T]: T[P] extends ReadonlyArray<infer U>
    ? Array<DeepMutable<U>>
    : T[P] extends ReadonlyMap<infer K, infer V>
      ? Map<K, DeepMutable<V>>
      : T[P] extends object
        ? DeepMutable<T[P]>
        : T[P];
};

// =============================================================================
// URL safety
// =============================================================================

/**
 * Whitelist of URL schemes considered safe to emit verbatim into HTML, DOCX
 * hyperlink relationships, or Markdown link destinations. Anything outside
 * this list is rejected by `sanitizeUrl` to prevent `javascript:`, `vbscript:`,
 * untrusted `data:` URIs, etc. from leaking through user-controlled input.
 */
const SAFE_URL_SCHEMES = new Set(["http", "https", "mailto", "tel", "ftp", "ftps", "sms"]);

/**
 * Validate and normalize a URL for use in a hyperlink. Returns a safe URL
 * string, or `undefined` if the input is unsafe.
 *
 * Rules:
 * - Empty/whitespace input → `undefined`.
 * - Fragment-only (`#anchor`) and protocol-relative (`//host`) URLs → kept.
 * - Relative URLs without a scheme → kept (no host injection possible since
 *   they resolve against the document base).
 * - URLs with a scheme are accepted only if the scheme is in
 *   {@link SAFE_URL_SCHEMES}.
 * - Surrounding whitespace and control characters that browsers historically
 *   strip when parsing URLs are removed before scheme detection so attackers
 *   can't bypass the check with `"  java\tscript:..."`.
 */
export function sanitizeUrl(url: string | undefined | null): string | undefined {
  if (url == null) {
    return undefined;
  }
  // Strip control characters (\x00-\x1F and \x7F) and surrounding whitespace.
  // Browsers historically strip these *before* parsing the scheme, so
  // `"java\nscript:..."` would otherwise pass a naive check.
  // oxlint-disable-next-line no-control-regex -- intentional sanitization
  const cleaned = url.replace(/[\u0000-\u001F\u007F]/g, "").trim();
  if (cleaned.length === 0) {
    return undefined;
  }
  // Fragment-only is always safe.
  if (cleaned.startsWith("#")) {
    return cleaned;
  }
  // Protocol-relative ("//host/...") is fine — same-scheme as the host page.
  if (cleaned.startsWith("//")) {
    return cleaned;
  }
  // Detect scheme. RFC 3986: scheme = ALPHA *( ALPHA / DIGIT / "+" / "-" / "." )
  const schemeMatch = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(cleaned);
  if (!schemeMatch) {
    // Relative URL — safe.
    return cleaned;
  }
  const scheme = schemeMatch[1].toLowerCase();
  if (SAFE_URL_SCHEMES.has(scheme)) {
    return cleaned;
  }
  return undefined;
}
