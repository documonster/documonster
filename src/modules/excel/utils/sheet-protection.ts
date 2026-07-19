import { Encryptor } from "@excel/utils/encryptor";
import { base64ToUint8Array, uint8ArrayToBase64 } from "@utils/utils";

// =============================================================================
// Sheet Protection Helper
// =============================================================================

/**
 * Result of applying sheet protection — contains the hash fields that both
 * Worksheet and WorksheetWriter need to persist.
 */
export interface SheetProtectionHash {
  sheet: boolean;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
  [key: string]: unknown;
}

/**
 * Build a sheet-protection object with optional password hashing.
 *
 * This is the shared implementation used by both `Worksheet.protect()` and
 * `WorksheetWriter.protect()`. The caller is responsible for assigning the
 * result to its own `sheetProtection` field.
 *
 * @param password - Optional password to hash
 * @param options  - Optional protection flags (objects, scenarios, etc.)
 * @returns A fully-populated sheet-protection object
 */
export async function buildSheetProtection<T extends { spinCount?: number }>(
  password?: string,
  options?: Partial<T>
): Promise<SheetProtectionHash> {
  const protection: SheetProtectionHash = { sheet: true };

  if (options && "spinCount" in options) {
    // force spinCount to be integer >= 0
    options.spinCount = Number.isFinite(options.spinCount)
      ? Math.round(Math.max(0, options.spinCount!))
      : 100000;
  }

  if (password !== undefined) {
    protection.algorithmName = "SHA-512";
    protection.saltValue = uint8ArrayToBase64(Encryptor.randomBytes(16));
    protection.spinCount = options && "spinCount" in options ? options.spinCount : 100000;
    protection.hashValue = await Encryptor.convertPasswordToHash(
      password,
      "SHA-512",
      protection.saltValue,
      protection.spinCount!
    );
  }

  if (options) {
    Object.assign(protection, options);
    if (!password && "spinCount" in options) {
      delete protection.spinCount;
    }
  }

  return protection;
}

/**
 * Verify a candidate password against a hash previously produced by
 * {@link buildSheetProtection} (i.e. `ws.sheetProtection` after `protect()`).
 * Recomputes the hash with the stored algorithm / salt / spin count and
 * compares — `protect()` stores only the hash, so this is the only way to
 * check a candidate without reimplementing the whole hashing scheme.
 *
 * Returns `false` (rather than throwing) when the stored protection carries no
 * hash — "protected with no password" — since there is nothing to verify
 * against.
 */
export async function verifySheetPassword(
  protection:
    | Pick<SheetProtectionHash, "hashValue" | "saltValue" | "algorithmName" | "spinCount">
    | null
    | undefined,
  password: string
): Promise<boolean> {
  if (!protection?.hashValue || !protection.saltValue || !protection.algorithmName) {
    return false;
  }
  const candidateHash = await Encryptor.convertPasswordToHash(
    password,
    protection.algorithmName,
    protection.saltValue,
    protection.spinCount ?? 100000
  );
  const candidate = base64ToUint8Array(candidateHash);
  const expected = base64ToUint8Array(protection.hashValue);
  // Portable constant-work comparison: always inspect every byte of the
  // longest input instead of returning at the first mismatch. JavaScript
  // engines cannot promise strict CPU-level constant time, but this avoids the
  // straightforward prefix timing leak of `candidateHash === hashValue` and
  // works in both Node.js and browsers without a runtime dependency.
  let difference = candidate.length ^ expected.length;
  const length = Math.max(candidate.length, expected.length);
  for (let i = 0; i < length; i++) {
    difference |= (candidate[i] ?? 0) ^ (expected[i] ?? 0);
  }
  return difference === 0;
}
