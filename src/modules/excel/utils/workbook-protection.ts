import { Encryptor } from "@excel/utils/encryptor";
import { uint8ArrayToBase64 } from "@utils/utils";

// =============================================================================
// Workbook Protection Helper
// =============================================================================

/**
 * Result of applying workbook protection — contains the fields that both
 * Workbook and WorkbookWriter need to persist.
 */
export interface WorkbookProtectionHash {
  lockStructure?: boolean;
  lockWindows?: boolean;
  lockRevision?: boolean;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

/**
 * Build a workbook-protection object with optional password hashing.
 *
 * This is the shared implementation used by both `Workbook.protect()` and
 * `WorkbookWriter.protect()`. The caller is responsible for assigning the
 * result to its own `protection` field.
 *
 * @param password - Optional password to hash
 * @param options  - Optional protection flags (lockStructure, lockWindows, lockRevision, spinCount)
 * @returns A fully-populated workbook-protection object
 */
export async function buildWorkbookProtection(
  password?: string,
  options?: {
    lockStructure?: boolean;
    lockWindows?: boolean;
    lockRevision?: boolean;
    spinCount?: number;
  }
): Promise<WorkbookProtectionHash> {
  const protection: WorkbookProtectionHash = {
    lockStructure: options?.lockStructure ?? true,
    lockWindows: options?.lockWindows,
    lockRevision: options?.lockRevision
  };

  if (password) {
    protection.algorithmName = "SHA-512";
    protection.saltValue = uint8ArrayToBase64(Encryptor.randomBytes(16));
    protection.spinCount =
      options?.spinCount != null && Number.isFinite(options.spinCount)
        ? Math.round(Math.max(0, options.spinCount))
        : 100000;
    protection.hashValue = await Encryptor.convertPasswordToHash(
      password,
      "SHA-512",
      protection.saltValue,
      protection.spinCount
    );
  }

  return protection;
}
