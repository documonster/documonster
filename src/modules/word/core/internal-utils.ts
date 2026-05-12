/**
 * DOCX Module - Internal Utilities (unified shared helpers)
 *
 * This is the single source of truth for cross-cutting low-level utilities
 * used throughout the word module. Consolidates previously duplicated
 * implementations of base64, UUID, UTF-16LE encoding, and XML helpers.
 */

import { DocxError } from "../errors";

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
function generateUuid(): string {
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

// =============================================================================
// Random Bytes
// =============================================================================

/**
 * Generate cryptographically secure random bytes.
 * Throws `DocxError` if `crypto.getRandomValues` is unavailable on the host.
 */
export function randomBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  if (typeof globalThis.crypto !== "undefined" && globalThis.crypto.getRandomValues) {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    throw new DocxError(
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

/**
 * Strip path-traversal segments and other unsafe characters from a
 * caller-supplied media file name so it can be used as a leaf entry
 * inside an OPC ZIP package without enabling zipslip.
 *
 * The previous behaviour was to forward `image.fileName` /
 * `font.fileName` straight into `archive.add`, which let a hostile DOCX
 * round-tripped through `readDocx` write entries like
 * `word/media/../../etc/passwd.png` into the output package — a real
 * attack vector when the file is later unpacked by a third-party tool.
 *
 * Returns a single safe leaf name. Falls back to `fallback` (default
 * `"file.bin"`) when the cleaned name would be empty.
 */
export function sanitizeMediaFileName(raw: string | undefined, fallback = "file.bin"): string {
  if (!raw) {
    return fallback;
  }
  // Drop directory components — only the leaf name should ever reach
  // the ZIP path layer.
  const lastSep = Math.max(raw.lastIndexOf("/"), raw.lastIndexOf("\\"));
  let leaf = lastSep >= 0 ? raw.substring(lastSep + 1) : raw;
  // Strip leading dots so attribute-bearing names ("..png", ".htaccess",
  // ".." itself) can't smuggle traversal back in via OS filesystem
  // semantics.
  while (leaf.startsWith(".")) {
    leaf = leaf.substring(1);
  }
  // Whitelist alnum, dash, underscore, dot. Replace everything else
  // with underscore. Empty result triggers fallback.
  leaf = leaf.replace(/[^A-Za-z0-9._-]/g, "_");
  // Collapse runs of dots so "foo..bin" can't be reinterpreted as
  // traversal by a permissive consumer.
  leaf = leaf.replace(/\.{2,}/g, ".");
  return leaf || fallback;
}
