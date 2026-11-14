/**
 * Base64 (RFC 4648) with one reading on every platform.
 *
 * The platform decoders disagree. `atob` and `Uint8Array.fromBase64` refuse a character outside the
 * alphabet; Node's `Buffer.from(text, "base64")` skips it, and reads the URL-safe alphabet too, so
 * the same malformed text decoded to bytes under Node and threw in a browser. Here the text is
 * checked first — as WHATWG's forgiving-base64 checks it, which is what `atob` and the loose mode
 * of `fromBase64` both implement — and only valid text reaches whichever decoder is fastest:
 *
 * - ASCII white space is ignored anywhere; padding is optional, but if present it must complete
 *   the last group (`YQ=` is refused, `YQ` and `YQ==` are read).
 * - Anything else outside the alphabet is refused, and so is a lone final character.
 * - Bits left over in the last character are ignored, as every one of those decoders ignores them.
 */

import { isNode } from "./env";
import { BaseError } from "./errors";

/**
 * Which alphabet: standard (`+` `/`) or URL-safe (`-` `_`). `"either"` accepts both, mixed, as
 * Node's decoder does — for reading text the library did not write.
 */
export type Base64Alphabet = "base64" | "base64url";

/** `Uint8Array.fromBase64` and `Uint8Array.prototype.toBase64` (ES2026); not yet in TypeScript's `lib`. */
interface NativeDecoder {
  readonly fromBase64?: (this: typeof Uint8Array, text: string) => Uint8Array<ArrayBuffer>;
}
interface NativeEncoder {
  readonly toBase64?: (
    this: Uint8Array,
    options?: { alphabet?: Base64Alphabet; omitPadding?: boolean }
  ) => string;
}

const WHITE_SPACE = /[\t\n\f\r ]/g;
const STANDARD = /^[A-Za-z0-9+/]*$/;
const URL_SAFE_ONLY = /^[A-Za-z0-9\-_]*$/;
const EITHER = /^[A-Za-z0-9+/\-_]*$/;

/**
 * `text` as unpadded standard-alphabet base64 with no white space, or `undefined` when it is not
 * base64 in `alphabet`.
 */
function normalize(text: string, alphabet: Base64Alphabet | "either"): string | undefined {
  let t = text.replace(WHITE_SPACE, "");
  if (t.length % 4 === 0) {
    t = t.endsWith("==") ? t.slice(0, -2) : t.endsWith("=") ? t.slice(0, -1) : t;
  }
  if (t.length % 4 === 1) {
    return undefined;
  }
  const valid =
    alphabet === "base64" ? STANDARD : alphabet === "base64url" ? URL_SAFE_ONLY : EITHER;
  if (!valid.test(t)) {
    return undefined;
  }
  return alphabet === "base64" ? t : t.replace(/-/g, "+").replace(/_/g, "/");
}

/** Whether `text` is base64 in `alphabet`, by the rules above, without decoding it. */
export function isBase64(text: string, alphabet: Base64Alphabet | "either" = "base64"): boolean {
  return normalize(text, alphabet) !== undefined;
}

/**
 * The bytes `text` encodes, or `undefined` when it is not base64 in `alphabet`. The result owns
 * its memory: it is never a view of a shared pool.
 */
export function decodeBase64(
  text: string,
  alphabet: Base64Alphabet | "either" = "base64"
): Uint8Array<ArrayBuffer> | undefined {
  const t = normalize(text, alphabet);
  if (t === undefined) {
    return undefined;
  }
  const native = (Uint8Array as NativeDecoder).fromBase64;
  if (typeof native === "function") {
    return native.call(Uint8Array, t);
  }
  if (isNode()) {
    const buffer = Buffer.from(t, "base64");
    // A short result is a slice of Buffer's shared allocation pool; copy it out of the pool
    return buffer.byteOffset === 0 && buffer.byteLength === buffer.buffer.byteLength
      ? new Uint8Array(buffer.buffer as ArrayBuffer)
      : new Uint8Array(buffer);
  }
  const binary = atob(t);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Text given as base64 that is not. Internal: a public call reports it as its own module's error
 * before it gets here, or translates it.
 */
export class InvalidBase64Error extends BaseError {
  override name = "InvalidBase64Error";
}

/** {@link decodeBase64}, throwing {@link InvalidBase64Error} rather than answering `undefined`. */
export function decodeBase64OrThrow(
  text: string,
  alphabet: Base64Alphabet | "either" = "base64"
): Uint8Array<ArrayBuffer> {
  const bytes = decodeBase64(text, alphabet);
  if (bytes === undefined) {
    throw new InvalidBase64Error("The text is not base64");
  }
  return bytes;
}

/** Bytes per `String.fromCharCode` call: within every engine's argument-count limit. */
const CHUNK = 0x8000;

/** `bytes` as base64 in `alphabet`, padded unless `omitPadding`. */
export function encodeBase64(
  bytes: Uint8Array,
  alphabet: Base64Alphabet = "base64",
  omitPadding = false
): string {
  const native = (bytes as NativeEncoder).toBase64;
  if (typeof native === "function") {
    return native.call(bytes, { alphabet, omitPadding });
  }
  let text: string;
  if (isNode()) {
    text = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString("base64");
  } else {
    const parts: string[] = [];
    for (let i = 0; i < bytes.length; i += CHUNK) {
      parts.push(String.fromCharCode(...bytes.subarray(i, i + CHUNK)));
    }
    text = btoa(parts.join(""));
  }
  if (alphabet === "base64url") {
    text = text.replace(/\+/g, "-").replace(/\//g, "_");
  }
  return omitPadding ? text.replace(/=+$/, "") : text;
}
