import { checkUint } from "@archive/core/entry-checks";
import { randomBytes as platformRandomBytes } from "@utils/crypto";

/**
 * `length` cryptographically random bytes, from the platform's generator. Published as
 * `Archive.Crypto.randomBytes`; a length that is not a non-negative integer is an
 * `ArchiveError`.
 */
export function randomBytes(length: number): Uint8Array {
  checkUint("Archive.Crypto.randomBytes", "length", length, 2 ** 31 - 1);
  return platformRandomBytes(length);
}
