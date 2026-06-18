/**
 * `Archive` namespace surface — high-level ZIP/TAR creation & extraction,
 * remote/random-access readers, and crypto/CRC primitives.
 *
 * `import { Archive } from "@cj-tech-master/excelts/archive"` →
 *   `Archive.zip(...)`, `Archive.unzip(...)`, `new Archive.ZipArchive()`,
 *   `new Archive.ZipReader(...)`, `Archive.editZip(...)`,
 *   `new Archive.TarArchive()`, `Archive.tar(...)`, `Archive.parseTar(...)`,
 *   `new Archive.RemoteZipReader(...)`, …
 *
 * Single flat namespace (archive is a single-purpose module). Re-exported via
 * `export * as Archive`, tree-shaken per-member on rolldown / rspack.
 *
 * Only platform-agnostic members live here. Platform-specific value exports
 * (compression, crc32, worker pool, fs, tar-gzip) stay on the platform
 * entrypoints (`index.ts` / `index.browser.ts`). Error classes and pure types
 * remain flat at the index top level (matching the other domain modules).
 */

// Unified archive I/O adapters
export { toAsyncIterable, toReadableStream } from "@archive/io/archive-source";

// Random Access / HTTP Range reading
export { HttpRangeReader, BufferReader } from "@archive/io/random-access";

// Remote ZIP reading over HTTP range requests
export { RemoteZipReader } from "@archive/unzip/remote-zip-reader";

// Abort helpers
export { createAbortError, isAbortError, throwIfAborted } from "@archive/shared/errors";

// High-level ZIP API
export { zip } from "@archive/create-archive";
export { unzip } from "@archive/read-archive";
export { ZipArchive, ZipEditor, editZip, editZipUrl, ZipEditPlan } from "@archive/zip";
export { ZipReader, UnzipEntry } from "@archive/unzip";

// TAR archive support (unified API compatible with ZIP)
export {
  TAR_BLOCK_SIZE,
  TAR_TYPE,
  TarArchive,
  TarReader,
  TarReaderEntry,
  createTarArchive,
  createTarReader,
  tar,
  tarSync,
  parseTar,
  parseTarStream,
  untar,
  isTarFile,
  isTarDirectory,
  isTarSymlink,
  isTarHardLink,
  isTarDataEntry
} from "@archive/tar/index.browser";

// Binary / encoding utilities — re-exported so archive can be used standalone
// without reaching into internal @utils modules.
export { base64ToUint8Array, uint8ArrayToBase64 } from "@utils/utils.base";
export { concatUint8Arrays, stringToUint8Array, uint8ArrayToString } from "@utils/binary";

// Encryption — high-level functions and constants
export {
  ZIP_CRYPTO_HEADER_SIZE,
  zipCryptoInitKeys,
  zipCryptoDecrypt,
  zipCryptoEncrypt,
  AES_VENDOR_ID,
  AES_VERSION_AE1,
  AES_VERSION_AE2,
  AES_EXTRA_FIELD_ID,
  AES_SALT_LENGTH,
  AES_KEY_LENGTH,
  AES_AUTH_CODE_LENGTH,
  AES_PASSWORD_VERIFY_LENGTH,
  COMPRESSION_METHOD_AES,
  aesDecrypt,
  aesEncrypt,
  aesEncryptedSize,
  buildAesExtraField,
  randomBytes,
  getEncryptionMethodName,
  isAesEncryption,
  getAesKeyStrength,
  encryptionMethodFromAesKeyStrength
} from "@archive/crypto";
