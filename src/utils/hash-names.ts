/**
 * Hash algorithm names, as every API that takes one accepts them — Word's document protection, and the Office
 * file encryption Word and Excel share.
 *
 * `"SHA-256"` — the WebCrypto spelling, and what `Security.getState` reports — is
 * the canonical form. The spelling Office's encryption file format uses
 * (`"SHA256"`, in an `EncryptionInfo` part) is accepted as well, so a setting can
 * be copied between document protection and file encryption. Each API normalises
 * a name here, then writes the spelling its own format needs.
 */

/** A hash algorithm name in its canonical spelling. */
export type CanonicalHashAlgorithm = "SHA-1" | "SHA-256" | "SHA-384" | "SHA-512";

/** The encryption file format's spelling of a hash algorithm (`EncryptionInfo`'s `hashAlgorithm`). */
export type CompactHashAlgorithm = "SHA1" | "SHA256" | "SHA384" | "SHA512";

const CANONICAL: ReadonlyMap<string, CanonicalHashAlgorithm> = new Map([
  ["SHA-1", "SHA-1"],
  ["SHA1", "SHA-1"],
  ["SHA-256", "SHA-256"],
  ["SHA256", "SHA-256"],
  ["SHA-384", "SHA-384"],
  ["SHA384", "SHA-384"],
  ["SHA-512", "SHA-512"],
  ["SHA512", "SHA-512"]
]);

/** Every spelling accepted, canonical ones first — the vocabulary an argument check names. */
export const HASH_ALGORITHM_NAMES: ReadonlySet<string> = new Set([
  "SHA-1",
  "SHA-256",
  "SHA-384",
  "SHA-512",
  "SHA1",
  "SHA256",
  "SHA384",
  "SHA512"
]);

/** The canonical spelling of `name`, or `undefined` for a name that is neither spelling. */
export function canonicalHashAlgorithm(name: unknown): CanonicalHashAlgorithm | undefined {
  return typeof name === "string" ? CANONICAL.get(name) : undefined;
}

/** The encryption file format's spelling of a canonical name (`"SHA-256"` → `"SHA256"`). */
export function compactHashAlgorithm(name: CanonicalHashAlgorithm): CompactHashAlgorithm {
  return name.replace("-", "") as CompactHashAlgorithm;
}
