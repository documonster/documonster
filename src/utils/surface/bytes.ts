/**
 * `Bytes` — conversions between bytes and text that the platforms offer unevenly.
 *
 * `import { Bytes } from "documonster/utils"` →
 *   `Bytes.fromBase64(text)`, `Bytes.toBase64(bytes)`, `Bytes.concat(parts)`,
 *   `Bytes.fromHex(text)`, `Bytes.toHex(bytes)`, `Bytes.fromText(text)`, `Bytes.toText(bytes)`,
 *   `Bytes.random(length)`.
 *
 * Each one gives the same answer on every platform this package supports, and uses the
 * platform's own implementation where there is one:
 *
 * - base64: `Uint8Array.fromBase64` / `toBase64` are ES2026 — absent from Node 22 and 24 and from
 *   browsers older than Chrome 140, Firefox 133 and Safari 18.2 — and the decoders that stand in
 *   for them disagree about malformed text (see {@link fromBase64}). The names and options are
 *   the native ones, so moving to the native call later is a rename.
 * - hex: `Uint8Array.fromHex` / `toHex` are ES2026 too, with the same gaps.
 * - `concat`: there is no native `Uint8Array` concatenation; `Buffer.concat` is Node's only.
 * - `fromText` / `toText`: `TextEncoder` / `TextDecoder`, with a refusal reported as this entry's
 *   error.
 * - `random`: `crypto.getRandomValues`, which refuses more than 65,536 bytes in one call.
 *
 * Every result is a plain `Uint8Array` over an `ArrayBuffer` of its own — never a Node `Buffer`,
 * and never a view of memory something else holds.
 */

import { decodeBase64, encodeBase64 } from "../base64";
import { decodeHex, encodeHex } from "../hex";
import { randomBytes } from "../random-bytes";
import { check, s } from "../schema";
import type { SchemaCheckOptions, SchemaMatches } from "../schema";
import { UtilsArgumentError } from "./errors";
import type { UtilsFromBase64Options, UtilsToBase64Options, UtilsToTextOptions } from "./types";

function at(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: (f, message) => new UtilsArgumentError(f, message) };
}

const alphabet = () => s.member<"base64" | "base64url">(["base64", "base64url"]).optional();

const fromBase64Shape = s.object({ alphabet: alphabet() }, { readonly: true });
const _fromBase64Match: SchemaMatches<typeof fromBase64Shape, UtilsFromBase64Options> = true;
void _fromBase64Match;
const fromBase64Options = fromBase64Shape.optional();

const toBase64Shape = s.object(
  { alphabet: alphabet(), omitPadding: s.boolean().optional() },
  { readonly: true }
);
const _toBase64Match: SchemaMatches<typeof toBase64Shape, UtilsToBase64Options> = true;
void _toBase64Match;
const toBase64Options = toBase64Shape.optional();

const toTextShape = s.object(
  { encoding: s.string({ nonEmpty: true }).optional(), fatal: s.boolean().optional() },
  { readonly: true }
);
const _toTextMatch: SchemaMatches<typeof toTextShape, UtilsToTextOptions> = true;
void _toTextMatch;
const toTextOptions = toTextShape.optional();

/** Bytes: a `Uint8Array`, from any realm. */
const bytesSchema = s.custom<Uint8Array>(
  v => ArrayBuffer.isView(v) && Object.prototype.toString.call(v) === "[object Uint8Array]",
  "a Uint8Array"
);

const textSchema = s.string();
const partsSchema = s.readonlyArray(bytesSchema);
const lengthSchema = s.number({ integer: true, min: 0, max: 2 ** 31 - 1 });

/**
 * The bytes `text` encodes, as `Uint8Array.fromBase64(text, options)` reads them.
 *
 * ASCII white space is ignored anywhere and padding is optional, but padding that is present must
 * complete the last group. Any other character outside the alphabet is refused — including the
 * other alphabet's `+` `/` or `-` `_` — and so is a lone final character. Node's
 * `Buffer.from(text, "base64")` skips such characters and returns what is left, and `atob` throws
 * a `DOMException`; this refuses them everywhere.
 *
 * @throws {UtilsArgumentError} When `text` is not base64 in `options.alphabet` (`"base64"` by default).
 */
export function fromBase64(
  text: string,
  options?: UtilsFromBase64Options
): Uint8Array<ArrayBuffer> {
  check(textSchema, text, at("Bytes.fromBase64", "text"));
  check(fromBase64Options, options, at("Bytes.fromBase64", "options"));
  const which = options?.alphabet ?? "base64";
  const bytes = decodeBase64(text, which);
  if (bytes === undefined) {
    throw new UtilsArgumentError("text", `Bytes.fromBase64: text is not ${which} text`);
  }
  return bytes;
}

/** `bytes` as base64, as `bytes.toBase64(options)` writes it: standard alphabet and padded by default. */
export function toBase64(bytes: Uint8Array, options?: UtilsToBase64Options): string {
  check(bytesSchema, bytes, at("Bytes.toBase64", "bytes"));
  check(toBase64Options, options, at("Bytes.toBase64", "options"));
  return encodeBase64(bytes, options?.alphabet ?? "base64", options?.omitPadding ?? false);
}

/**
 * The bytes `text` encodes, as `Uint8Array.fromHex(text)` reads them: pairs of `0-9`, `a-f` and
 * `A-F`, of either case — no white space, no `0x` prefix, no odd final digit.
 *
 * @throws {UtilsArgumentError} When `text` is not hex.
 */
export function fromHex(text: string): Uint8Array<ArrayBuffer> {
  check(textSchema, text, at("Bytes.fromHex", "text"));
  const bytes = decodeHex(text);
  if (bytes === undefined) {
    throw new UtilsArgumentError("text", "Bytes.fromHex: text is not hex text");
  }
  return bytes;
}

/** `bytes` as lower-case hex, as `bytes.toHex()` writes it. */
export function toHex(bytes: Uint8Array): string {
  check(bytesSchema, bytes, at("Bytes.toHex", "bytes"));
  return encodeHex(bytes);
}

/** `parts` joined into one new array. The result is always a copy, even of a single part. */
export function concat(parts: readonly Uint8Array[]): Uint8Array<ArrayBuffer> {
  check(partsSchema, parts, at("Bytes.concat", "parts"));
  let length = 0;
  for (const part of parts) {
    length += part.length;
  }
  const out = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

/** `text` as UTF-8. A lone surrogate is written as U+FFFD, as `TextEncoder` writes it. */
export function fromText(text: string): Uint8Array<ArrayBuffer> {
  check(textSchema, text, at("Bytes.fromText", "text"));
  return new TextEncoder().encode(text) as Uint8Array<ArrayBuffer>;
}

/**
 * `bytes` decoded as text, by `TextDecoder`: UTF-8 unless `options.encoding` names another of its
 * labels. A leading byte order mark is dropped. Bytes that are not valid in the encoding become
 * U+FFFD, or are refused with `options.fatal`.
 *
 * @throws {UtilsArgumentError} When `options.encoding` is a label `TextDecoder` does not know, or
 *   with `options.fatal`, when `bytes` are not valid in the encoding.
 */
export function toText(bytes: Uint8Array, options?: UtilsToTextOptions): string {
  check(bytesSchema, bytes, at("Bytes.toText", "bytes"));
  check(toTextOptions, options, at("Bytes.toText", "options"));
  const encoding = options?.encoding ?? "utf-8";
  let decoder: TextDecoder;
  try {
    decoder = new TextDecoder(encoding, { fatal: options?.fatal ?? false });
  } catch (cause) {
    throw new UtilsArgumentError(
      "options.encoding",
      `Bytes.toText: options.encoding ${JSON.stringify(encoding)} is not an encoding TextDecoder knows`,
      { cause }
    );
  }
  try {
    return decoder.decode(bytes);
  } catch (cause) {
    throw new UtilsArgumentError("bytes", `Bytes.toText: bytes are not valid ${decoder.encoding}`, {
      cause
    });
  }
}

/**
 * `length` cryptographically random bytes, from `crypto.getRandomValues` — filled in pieces, so a
 * length above its 65,536-byte quota is not refused.
 */
export function random(length: number): Uint8Array<ArrayBuffer> {
  check(lengthSchema, length, at("Bytes.random", "length"));
  return randomBytes(length) as Uint8Array<ArrayBuffer>;
}
