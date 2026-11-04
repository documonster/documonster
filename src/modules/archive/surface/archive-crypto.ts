/**
 * `Archive.Crypto` — the low-level ZIP encryption primitives (traditional PKWARE
 * "ZipCrypto" and WinZip AES AE-1/AE-2) and their format constants.
 *
 * Reading and writing encrypted archives does not need any of this: pass `password`
 * (and `encryptionMethod`) to `Archive.createZip` / `Archive.openZip`. These are for
 * callers that handle entry payloads themselves, e.g. through `getRawCompressedData`.
 *
 * Both decryptors throw an `ArchiveError` when the password does not verify.
 *
 * Names are verb first (`encryptAes`, `decryptZipCrypto`, `createAesExtraField`), like the
 * rest of the library. `randomBytes` keeps the name of the `node:crypto` function it stands
 * in for.
 */
export {
  ZIP_CRYPTO_HEADER_SIZE,
  zipCryptoInitKeys as createZipCryptoKeys,
  zipCryptoEncrypt as encryptZipCrypto,
  AES_VENDOR_ID,
  AES_VERSION_AE1,
  AES_VERSION_AE2,
  AES_EXTRA_FIELD_ID,
  AES_SALT_LENGTH,
  AES_KEY_LENGTH,
  AES_AUTH_CODE_LENGTH,
  AES_PASSWORD_VERIFY_LENGTH,
  COMPRESSION_METHOD_AES,
  aesDecrypt as decryptAes,
  aesEncrypt as encryptAes,
  aesEncryptedSize as getAesEncryptedSize,
  buildAesExtraField as createAesExtraField,
  randomBytes,
  getEncryptionMethodName,
  isAesEncryption,
  getAesKeyStrength,
  encryptionMethodFromAesKeyStrength as getEncryptionMethodForAesKeyStrength
} from "@archive/crypto";
export { zipCryptoDecryptOrThrow as decryptZipCrypto } from "@archive/crypto/zip-crypto";
