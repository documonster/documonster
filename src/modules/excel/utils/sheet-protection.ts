import { Encryptor } from "@excel/utils/encryptor";
import { uint8ArrayToBase64 } from "@utils/utils";

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

  if (password) {
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
