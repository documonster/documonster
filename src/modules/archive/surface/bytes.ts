/**
 * The byte utilities `Archive` publishes, with their arguments checked: a wrong one is an
 * `ArchiveArgumentError` rather than a raw `TypeError`, a `RangeError` or a silently wrong result.
 * The deprecated flat names forward to the unchecked implementations, as they always did.
 */
import {
  crc32 as crc32Impl,
  crc32Finalize as crc32FinalizeImpl,
  crc32Update as crc32UpdateImpl
} from "@archive/compression/crc32";
import { checkBytes, checkString, checkUint } from "@archive/core/entry-checks";
import { ArchiveArgumentError } from "@archive/core/errors";
import {
  concatUint8Arrays as concatImpl,
  stringToUint8Array as stringToUint8ArrayImpl,
  uint8ArrayToString as uint8ArrayToStringImpl
} from "@utils/binary";
import {
  base64ToUint8Array as base64ToUint8ArrayImpl,
  uint8ArrayToBase64 as uint8ArrayToBase64Impl
} from "@utils/utils.base";

const CRC_MAX = 0xffffffff;

/** CRC-32 (IEEE) of `data`. */
export function crc32(data: Uint8Array): number {
  checkBytes("Archive.crc32", "data", data);
  return crc32Impl(data);
}

/**
 * Continue a running CRC-32 over `data`. Start from `0xffffffff` and finish with
 * {@link crc32Finalize}.
 */
export function crc32Update(crc: number, data: Uint8Array): number {
  checkUint("Archive.crc32Update", "crc", crc, CRC_MAX);
  checkBytes("Archive.crc32Update", "data", data);
  return crc32UpdateImpl(crc, data);
}

/** Finish a running CRC-32 from {@link crc32Update}. */
export function crc32Finalize(crc: number): number {
  checkUint("Archive.crc32Finalize", "crc", crc, CRC_MAX);
  return crc32FinalizeImpl(crc);
}

/** Standard base64 (RFC 4648 §4), padding optional; whitespace is ignored. */
const BASE64 = /^[A-Za-z0-9+/]*={0,2}$/;

/**
 * Decode standard base64. Whitespace is ignored and padding is optional; anything else
 * outside the alphabet — including the URL-safe `-` and `_` — is refused, so the result is
 * the same on every platform.
 */
export function base64ToUint8Array(base64: string): Uint8Array {
  checkString("Archive.base64ToUint8Array", "base64", base64);
  const compact = base64.replace(/[\t\n\f\r ]+/g, "");
  if (!BASE64.test(compact) || compact.replace(/=+$/, "").length % 4 === 1) {
    throw new ArchiveArgumentError(
      "base64",
      "Archive.base64ToUint8Array: base64 is not valid base64"
    );
  }
  return base64ToUint8ArrayImpl(compact);
}

/** Encode `bytes` as standard base64, padded. */
export function uint8ArrayToBase64(bytes: Uint8Array): string {
  checkBytes("Archive.uint8ArrayToBase64", "bytes", bytes);
  return uint8ArrayToBase64Impl(bytes);
}

/**
 * Concatenate `arrays`. `totalLength`, when given, must be their combined length: it saves
 * the sum, it does not truncate or pad.
 */
export function concatUint8Arrays(arrays: readonly Uint8Array[], totalLength?: number): Uint8Array {
  if (!Array.isArray(arrays)) {
    throw new ArchiveArgumentError(
      "arrays",
      "Archive.concatUint8Arrays: arrays must be an array of Uint8Array"
    );
  }
  let sum = 0;
  for (let i = 0; i < arrays.length; i++) {
    checkBytes("Archive.concatUint8Arrays", `arrays[${i}]`, arrays[i]);
    sum += arrays[i]!.length;
  }
  if (totalLength !== undefined && totalLength !== sum) {
    throw new ArchiveArgumentError(
      "totalLength",
      `Archive.concatUint8Arrays: totalLength is ${String(totalLength)}, but the arrays hold ${sum} bytes`
    );
  }
  return concatImpl(arrays, sum);
}

/** Encode `str` as UTF-8. */
export function stringToUint8Array(str: string): Uint8Array {
  checkString("Archive.stringToUint8Array", "str", str);
  return stringToUint8ArrayImpl(str);
}

/** Decode `arr` with `encoding` (a `TextDecoder` label; default `"utf-8"`). */
export function uint8ArrayToString(arr: Uint8Array, encoding?: string): string {
  checkBytes("Archive.uint8ArrayToString", "arr", arr);
  checkString("Archive.uint8ArrayToString", "encoding", encoding, true);
  try {
    return uint8ArrayToStringImpl(arr, encoding);
  } catch (cause) {
    throw new ArchiveArgumentError(
      "encoding",
      `Archive.uint8ArrayToString: ${JSON.stringify(encoding)} is not an encoding TextDecoder knows`,
      { cause }
    );
  }
}
