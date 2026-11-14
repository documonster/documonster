/**
 * The option types of `documonster/utils`.
 */

/** A base64 alphabet: standard (`+` `/`, RFC 4648 §4) or URL-safe (`-` `_`, §5). */
export type UtilsBase64Alphabet = "base64" | "base64url";

/** Options of `Bytes.fromBase64`, as `Uint8Array.fromBase64` takes them. */
export interface UtilsFromBase64Options {
  /** The alphabet `text` is in. Default `"base64"`; the other alphabet's two characters are refused. */
  readonly alphabet?: UtilsBase64Alphabet;
}

/** Options of `Bytes.toBase64`, as `Uint8Array.prototype.toBase64` takes them. */
export interface UtilsToBase64Options {
  /** The alphabet to write. Default `"base64"`. */
  readonly alphabet?: UtilsBase64Alphabet;
  /** Leave out the `=` padding. Default `false`. */
  readonly omitPadding?: boolean;
}

/** Options of `Bytes.toText`, as `TextDecoder` takes them. */
export interface UtilsToTextOptions {
  /** A `TextDecoder` encoding label. Default `"utf-8"`. */
  readonly encoding?: string;
  /** Refuse bytes that are not valid in the encoding, rather than writing U+FFFD for them. Default `false`. */
  readonly fatal?: boolean;
}
