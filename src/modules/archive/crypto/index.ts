/**
 * Crypto module for ZIP encryption/decryption.
 *
 * Exports both ZipCrypto (traditional PKWARE) and AES (WinZip) implementations.
 *
 * Public API (re-exported in index.base.ts):
 * - High-level functions: zipCryptoEncrypt, zipCryptoDecrypt, aesEncrypt, aesDecrypt
 * - Constants: ZIP_CRYPTO_HEADER_SIZE, AES_* constants
 * - Types: ZipEncryptionMethod, AesKeyStrength, etc.
 *
 * Internal functions (used by zip/stream.ts, zip/zip-bytes.ts, unzip/zip-parser.ts):
 * - zipCryptoInitKeys, zipCryptoCreateHeader, zipCryptoEncryptByte
 * - aesEncryptedSize, buildAesExtraField, randomBytes
 */

// ZipCrypto (Traditional PKWARE Encryption)
export {
  ZIP_CRYPTO_HEADER_SIZE,
  zipCryptoInitKeys,
  zipCryptoDecrypt,
  zipCryptoEncrypt,
  zipCryptoCreateHeader,
  zipCryptoEncryptByte,
  zipCryptoDecryptByte,
  zipCryptoCheckPassword as zipCryptoVerifyPassword,
  type ZipCryptoState
} from "@archive/crypto/zip-crypto";

// AES (WinZip AE-1 / AE-2)
export {
  AES_VENDOR_ID,
  AES_VERSION_AE1,
  AES_VERSION_AE2,
  AES_EXTRA_FIELD_ID,
  AES_SALT_LENGTH,
  AES_KEY_LENGTH,
  AES_AUTH_CODE_LENGTH,
  AES_PASSWORD_VERIFY_LENGTH,
  COMPRESSION_METHOD_AES,
  AES_STRENGTH_FROM_BYTE,
  aesDecrypt,
  aesEncrypt,
  aesEncryptedSize,
  aesVerifyPassword,
  buildAesExtraField,
  type AesKeyStrength,
  type AesExtraFieldInfo
} from "@archive/crypto/aes";

// Random bytes — from shared @utils/crypto
export { randomBytes } from "@utils/crypto";

// Types
export {
  type ZipEncryptionMethod,
  type ZipEncryptionInfo,
  type ZipPasswordOptions,
  type ZipEncryptionOptions,
  getEncryptionMethodName,
  isAesEncryption,
  getAesKeyStrength,
  encryptionMethodFromAesKeyStrength
} from "@archive/crypto/types";
