/**
 * Hash algorithm names for Word: the spellings every Word API accepts (from `@utils/hash-names`, shared with the
 * Office file encryption), and the legacy CryptoAPI ids a `w:documentProtection` names its hash by.
 */
import { canonicalHashAlgorithm } from "@utils/hash-names";

export {
  HASH_ALGORITHM_NAMES,
  canonicalHashAlgorithm,
  compactHashAlgorithm
} from "@utils/hash-names";
export type { CanonicalHashAlgorithm, CompactHashAlgorithm } from "@utils/hash-names";

/**
 * `w:cryptAlgorithmSid` (ECMA-376 Part 4 §14.8.1, the legacy CryptoAPI ids a
 * `w:documentProtection` hash is named by) and the names the model holds for
 * them. One table both ways, so a hash read as `MD5` is written back as `MD5`.
 */
const SID_NAMES: readonly (readonly [string, string])[] = [
  ["1", "MD2"],
  ["2", "MD4"],
  ["3", "MD5"],
  ["4", "SHA-1"],
  ["12", "SHA-256"],
  ["13", "SHA-384"],
  ["14", "SHA-512"]
];

/** The model's name for a `w:cryptAlgorithmSid`; an id the table lacks is `SID-<id>`. */
export function hashAlgorithmOfSid(sid: string | undefined): string | undefined {
  if (!sid) {
    return undefined;
  }
  return SID_NAMES.find(([id]) => id === sid)?.[1] ?? `SID-${sid}`;
}

/**
 * The `w:cryptAlgorithmSid` for a model name — either spelling of a SHA name, a
 * name {@link hashAlgorithmOfSid} gives — or `undefined` for one it cannot name.
 */
export function sidOfHashAlgorithm(name: string): string | undefined {
  const canonical = canonicalHashAlgorithm(name) ?? name;
  const known = SID_NAMES.find(([, n]) => n === canonical)?.[0];
  if (known !== undefined) {
    return known;
  }
  const kept = /^SID-(\d+)$/.exec(name);
  return kept ? kept[1] : undefined;
}
