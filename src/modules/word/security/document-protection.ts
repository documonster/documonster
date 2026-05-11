/**
 * DOCX Module - Document Protection
 *
 * Provides APIs for setting document protection passwords and restrictions.
 * Supports all protection types defined in OOXML:
 * - Read-only (no modifications allowed)
 * - Comments only (only add comments)
 * - Forms only (only fill form fields)
 * - Tracked changes (all changes are tracked)
 * - Sections (specific sections editable)
 *
 * Password hashing follows the ECMA-376 specification using
 * iterative SHA-based hashing with salt.
 */

import { bytesToBase64, base64ToBytes, stringToUtf16LE, randomBytes } from "../core/internal-utils";
import type { DocxDocument, DocumentSettings } from "../types";

// =============================================================================
// Types
// =============================================================================

/** Document protection type (edit restriction). */
export type ProtectionEditType = "none" | "readOnly" | "comments" | "trackedChanges" | "forms";

/** Hash algorithm for protection password. */
export type ProtectionHashAlgorithm = "SHA-1" | "SHA-256" | "SHA-384" | "SHA-512";

/** Options for setting document protection. */
export interface DocumentProtectionOptions {
  /** Type of edit restriction. */
  readonly edit: ProtectionEditType;
  /** Password to protect the document (will be hashed). */
  readonly password?: string;
  /** Hash algorithm. Default: "SHA-256". */
  readonly hashAlgorithm?: ProtectionHashAlgorithm;
  /** Number of hash iterations (spin count). Default: 100000. */
  readonly spinCount?: number;
  /** Whether protection is enforced. Default: true. */
  readonly enforcement?: boolean;
  /** Formatting restriction (prevent style changes). */
  readonly formatting?: boolean;
}

/** Protection state as stored in settings. */
export interface ProtectionState {
  /** Edit restriction type. */
  readonly edit: ProtectionEditType;
  /** Whether protection is enforced. */
  readonly enforcement: boolean;
  /** Hash algorithm used. */
  readonly hashAlgorithm?: string;
  /** Computed hash value (base64). */
  readonly hashValue?: string;
  /** Salt (base64). */
  readonly saltValue?: string;
  /** Spin count (iterations). */
  readonly spinCount?: number;
  /** Formatting restriction. */
  readonly formatting?: boolean;
}

// =============================================================================
// Protection API
// =============================================================================

/**
 * Set document protection on a DocxDocument.
 * Returns a new document with protection settings applied.
 *
 * @param doc - The document to protect.
 * @param options - Protection options.
 * @returns A new document with protection settings.
 *
 * @example
 * ```ts
 * const protected = await protectDocument(doc, {
 *   edit: "readOnly",
 *   password: "secret123",
 *   hashAlgorithm: "SHA-256"
 * });
 * ```
 */
export async function protectDocument(
  doc: DocxDocument,
  options: DocumentProtectionOptions
): Promise<DocxDocument> {
  const protection = await buildProtectionSettings(options);

  const existingSettings = doc.settings ?? {};
  const newSettings: DocumentSettings = {
    ...existingSettings,
    documentProtection: protection
  };

  return {
    ...doc,
    settings: newSettings
  };
}

/**
 * Remove document protection.
 *
 * @param doc - The protected document.
 * @returns A new document without protection.
 */
export function unprotectDocument(doc: DocxDocument): DocxDocument {
  if (!doc.settings?.documentProtection) {
    return doc;
  }

  const { documentProtection: _, ...rest } = doc.settings;
  return {
    ...doc,
    settings: rest as DocumentSettings
  };
}

/**
 * Check if a document is protected.
 */
export function isDocumentProtected(doc: DocxDocument): boolean {
  const dp = doc.settings?.documentProtection;
  return dp !== undefined && dp.enforcement === true;
}

/**
 * Get the protection state of a document.
 */
export function getProtectionState(doc: DocxDocument): ProtectionState | undefined {
  const dp = doc.settings?.documentProtection;
  if (!dp) {
    return undefined;
  }

  return {
    edit: (dp.edit as ProtectionEditType) ?? "none",
    enforcement: dp.enforcement === true,
    hashAlgorithm: dp.hashAlgorithm,
    hashValue: dp.hashValue,
    saltValue: dp.saltValue,
    spinCount: dp.spinCount,
    formatting: dp.formatting
  };
}

/**
 * Verify a password against the document's protection hash.
 * Returns true if the password matches, false otherwise.
 *
 * @param doc - The protected document.
 * @param password - The password to verify.
 * @returns Whether the password is correct.
 */
export async function verifyProtectionPassword(
  doc: DocxDocument,
  password: string
): Promise<boolean> {
  const dp = doc.settings?.documentProtection;
  if (!dp || !dp.hashValue || !dp.saltValue || !dp.hashAlgorithm) {
    return false;
  }

  const spinCount = dp.spinCount ?? 100000;
  const computedHash = await computePasswordHash(
    password,
    dp.saltValue,
    dp.hashAlgorithm as ProtectionHashAlgorithm,
    spinCount
  );

  return computedHash === dp.hashValue;
}

// =============================================================================
// Internal: Password Hashing
// =============================================================================

async function buildProtectionSettings(
  options: DocumentProtectionOptions
): Promise<Record<string, unknown>> {
  const enforcement = options.enforcement !== false;
  const result: Record<string, unknown> = {
    edit: options.edit,
    enforcement
  };

  if (options.formatting !== undefined) {
    result.formatting = options.formatting;
  }

  if (options.password) {
    const algorithm = options.hashAlgorithm ?? "SHA-256";
    const spinCount = options.spinCount ?? 100000;
    const salt = generateSalt();
    const hash = await computePasswordHash(options.password, salt, algorithm, spinCount);

    result.hashAlgorithm = algorithm;
    result.saltValue = salt;
    result.hashValue = hash;
    result.spinCount = spinCount;
  }

  return result;
}

async function computePasswordHash(
  password: string,
  saltBase64: string,
  algorithm: ProtectionHashAlgorithm,
  spinCount: number
): Promise<string> {
  const algName = algorithmToSubtle(algorithm);
  const saltBytes = base64ToBytes(saltBase64);
  // ECMA-376 §14.2.1: password must be encoded as UTF-16LE
  const passwordBytes = encodeUtf16LE(password);

  // Initial hash: H0 = Hash(salt + password)
  const initial = new Uint8Array(saltBytes.length + passwordBytes.length);
  initial.set(saltBytes, 0);
  initial.set(passwordBytes, saltBytes.length);

  let hash = new Uint8Array(await crypto.subtle.digest(algName, initial));

  // Iterative hashing: Hi = Hash(LE_int32(i) + Hi-1)
  for (let i = 0; i < spinCount; i++) {
    const iterBytes = new Uint8Array(4);
    iterBytes[0] = i & 0xff;
    iterBytes[1] = (i >> 8) & 0xff;
    iterBytes[2] = (i >> 16) & 0xff;
    iterBytes[3] = (i >> 24) & 0xff;

    const combined = new Uint8Array(4 + hash.length);
    combined.set(iterBytes, 0);
    combined.set(hash, 4);

    hash = new Uint8Array(await crypto.subtle.digest(algName, combined));
  }

  return bytesToBase64(hash);
}

/** Alias for UTF-16LE encoding (uses unified implementation). */
const encodeUtf16LE = stringToUtf16LE;

function algorithmToSubtle(alg: ProtectionHashAlgorithm): string {
  switch (alg) {
    case "SHA-1":
      return "SHA-1";
    case "SHA-256":
      return "SHA-256";
    case "SHA-384":
      return "SHA-384";
    case "SHA-512":
      return "SHA-512";
  }
}

function generateSalt(): string {
  return bytesToBase64(randomBytes(16));
}
