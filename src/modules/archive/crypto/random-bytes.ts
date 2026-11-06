import { checkArchiveArgument, uintSchema } from "@archive/core/arguments";
import { randomBytes as platformRandomBytes } from "@utils/random-bytes";

/**
 * `length` cryptographically random bytes, from the platform's generator. Published only as
 * the deprecated `Archive.randomBytes`; a length that is not a non-negative integer — a
 * missing one included — is an `ArchiveArgumentError`.
 */
export function randomBytes(length: number): Uint8Array {
  // Required: `checkUint` takes `undefined` as an omitted option, and no length is not zero bytes.
  checkArchiveArgument(uintSchema(2 ** 31 - 1), length, "Archive.randomBytes", "length");
  return platformRandomBytes(length);
}
