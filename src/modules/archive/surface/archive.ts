import { zip as zipImpl } from "@archive/create-archive";
/**
 * `Archive` namespace surface — ZIP / TAR / tar.gz creation and extraction,
 * remote/random-access readers, CRC-32 and PNG encoding; the low-level encryption
 * primitives sit under `Archive.Crypto`.
 *
 * `import { Archive } from "documonster/archive"` →
 *   `Archive.createZip()` / `Archive.openZip(bytes)` (builder / reader),
 *   `Archive.packZip(entries)` / `Archive.unpackZip(bytes)` (one-shot bytes / path map),
 *   the same four verbs for `Tar` and `TarGz`, `Archive.parseTar(bytes)`,
 *   `Archive.editZip(...)`, `Archive.crc32(bytes)`, …
 *
 * Every member here exists on both platforms, and compression-backed members resolve to
 * the platform's codec through the build's platform variants. The Node entry adds
 * `ArchiveFile` and `toNodeReadable` (see `archive-node.ts`). Raw DEFLATE / gzip / zlib
 * live on the `Compression` namespace, and the Node file-system helpers on `Fs`.
 */
import {
  ZIP_CRYPTO_HEADER_SIZE as ZIP_CRYPTO_HEADER_SIZEImpl,
  zipCryptoInitKeys as zipCryptoInitKeysImpl,
  zipCryptoDecrypt as zipCryptoDecryptImpl,
  zipCryptoEncrypt as zipCryptoEncryptImpl,
  AES_VENDOR_ID as AES_VENDOR_IDImpl,
  AES_VERSION_AE1 as AES_VERSION_AE1Impl,
  AES_VERSION_AE2 as AES_VERSION_AE2Impl,
  AES_EXTRA_FIELD_ID as AES_EXTRA_FIELD_IDImpl,
  AES_SALT_LENGTH as AES_SALT_LENGTHImpl,
  AES_KEY_LENGTH as AES_KEY_LENGTHImpl,
  AES_AUTH_CODE_LENGTH as AES_AUTH_CODE_LENGTHImpl,
  AES_PASSWORD_VERIFY_LENGTH as AES_PASSWORD_VERIFY_LENGTHImpl,
  COMPRESSION_METHOD_AES as COMPRESSION_METHOD_AESImpl,
  aesDecrypt as aesDecryptImpl,
  aesEncrypt as aesEncryptImpl,
  aesEncryptedSize as aesEncryptedSizeImpl,
  buildAesExtraField as buildAesExtraFieldImpl,
  randomBytes as randomBytesImpl,
  getEncryptionMethodName as getEncryptionMethodNameImpl,
  isAesEncryption as isAesEncryptionImpl,
  getAesKeyStrength as getAesKeyStrengthImpl,
  encryptionMethodFromAesKeyStrength as encryptionMethodFromAesKeyStrengthImpl
} from "@archive/crypto";
import { unzip as unzipImpl } from "@archive/read-archive";
import {
  createTarArchive as createTarArchiveImpl,
  createTarReader as createTarReaderImpl,
  tar as tarImpl,
  tarSync as tarSyncImpl
} from "@archive/tar/tar-archive";
import { untar as untarImpl } from "@archive/tar/tar-parser";

// Unified archive I/O adapters
export { toAsyncIterable, toReadableStream } from "@archive/io/archive-source";

// Random Access / HTTP Range reading
export { HttpRangeReader, createBufferReader } from "@archive/io/random-access";

// Remote ZIP reading over HTTP range requests
export { RemoteZipReader } from "@archive/unzip/remote-zip-reader";

// Abort helpers
export { createAbortError, isAbortError, throwIfAborted } from "@archive/core/errors";

// The format matrix — one verb per operation, the format in the name:
//
//              builder      reader      entries → bytes          bytes → path map              entry list
//   ZIP        createZip    openZip     packZip / packZipSync    unpackZip / unpackZipSync     —
//   TAR        createTar    openTar     packTar / packTarSync    unpackTar / unpackTarSync     parseTar / parseTarStream
//   tar.gz     createTarGz  openTarGz   packTarGz / …Sync        unpackTarGz / …Sync           parseTarGz / parseTarGzStream
//
// Random access: `RemoteZipReader.open(url)` / `RemoteZipReader.fromReader(reader)` (central
// directory only), over a reader from `HttpRangeReader.open(url)` or `createBufferReader(bytes)`;
// `editZip` / `editZipUrl` open a `ZipEditor`. Node adds `ArchiveFile` (see `archive-node.ts`).
export { createZip } from "@archive/create-archive";
export { openZip } from "@archive/read-archive";
export { packZip, packZipSync, unpackZip, unpackZipSync } from "@archive/zip-one-shot";
export {
  createTarArchive as createTar,
  openTarReader as openTar,
  tar as packTar,
  tarSync as packTarSync
} from "@archive/tar/tar-archive";
export {
  untar as unpackTar,
  untarSync as unpackTarSync,
  parseTar,
  parseTarStream
} from "@archive/tar/tar-parser";
export {
  createTarGzArchive as createTarGz,
  openTarGzReader as openTarGz,
  targz as packTarGz,
  targzSync as packTarGzSync,
  untargz as unpackTarGz,
  untargzSync as unpackTarGzSync,
  parseTarGz,
  // `…Stream` reads incrementally for both TAR forms, as `parseTarStream` does.
  parseTarGzIncremental as parseTarGzStream
} from "@archive/tar/tar-gzip";

// ZIP editing
export { ZipEditor, editZip, editZipUrl, ZipEditPlan } from "@archive/zip";
export { UnzipEntry } from "@archive/unzip";

// TAR entries and constants
export { TAR_BLOCK_SIZE, TAR_TYPE } from "@archive/tar/tar-constants";
export { createTarReaderEntry, type TarReaderEntry } from "@archive/tar/tar-archive";
export {
  isFile as isTarFile,
  isDirectory as isTarDirectory,
  isSymlink as isTarSymlink,
  isHardLink as isTarHardLink,
  isDataEntry as isTarDataEntry
} from "@archive/tar/tar-entry-info";

// CRC-32 (checked; see `surface/bytes.ts`)
export { crc32, crc32Update, crc32Finalize } from "@archive/surface/bytes";

// PNG — DEFLATE plus CRC-32-checked chunks, the two primitives this module owns. The
// drawing engine stops at pixels (`rasterizeToRgba`), and encoding needs DEFLATE.
export { encodePng } from "@archive/png";

// Binary / encoding utilities — re-exported so archive can be used standalone
// without reaching into internal @utils modules.
export {
  base64ToUint8Array,
  uint8ArrayToBase64,
  concatUint8Arrays,
  stringToUint8Array,
  uint8ArrayToString
} from "@archive/surface/bytes";

// Encryption primitives and constants — low level; see `archive-crypto.ts`
export * as Crypto from "@archive/surface/archive-crypto";

// =============================================================================
// Classes — public for `instanceof` checks and subclassing; build with the factories
// =============================================================================

/** The ZIP builder class {@link createZip} returns. */
export { ZipArchive } from "@archive/zip";
/** The ZIP reader class {@link openZip} returns. */
export { ZipReader } from "@archive/unzip";
/** The TAR builder class {@link createTar} returns. */
export { TarArchive } from "@archive/tar/tar-archive";
/** The TAR reader class {@link openTar} and {@link openTarGz} return. */
export { TarReader } from "@archive/tar/tar-archive";
/** The tar.gz builder class {@link createTarGz} returns. */
export { TarGzArchive } from "@archive/tar/tar-gzip";

// =============================================================================
// Deprecated
// =============================================================================

/**
 * @deprecated Use {@link createZip}, or {@link createTar} for `{ format: "tar" }`.
 * `zip` named a builder while `tar` named bytes. Removed in 1.0.
 */
export const zip = zipImpl;
/**
 * @deprecated Use {@link openZip}, or {@link openTar} for `{ format: "tar" }`.
 * `unzip` named a reader while `untar` named a map. Removed in 1.0.
 */
export const unzip = unzipImpl;
/** @deprecated Use {@link createTar}. Removed in 1.0. */
export const createTarArchive = createTarArchiveImpl;
/** @deprecated Use {@link openTar}. Removed in 1.0. */
export const createTarReader = createTarReaderImpl;
/** @deprecated Use {@link packTar}. Removed in 1.0. */
export const tar = tarImpl;
/** @deprecated Use {@link packTarSync}. Removed in 1.0. */
export const tarSync = tarSyncImpl;
/** @deprecated Use {@link unpackTar}. Removed in 1.0. */
export const untar = untarImpl;
/** @deprecated Use {@link Crypto.ZIP_CRYPTO_HEADER_SIZE}. Removed in 1.0. */
export const ZIP_CRYPTO_HEADER_SIZE = ZIP_CRYPTO_HEADER_SIZEImpl;
/** @deprecated Use {@link Crypto.createZipCryptoKeys}. Removed in 1.0. */
export const zipCryptoInitKeys = zipCryptoInitKeysImpl;
/**
 * @deprecated Use {@link Crypto.decryptZipCrypto}, which throws an `ArchiveError` where this
 * returns `null`. Removed in 1.0.
 */
export const zipCryptoDecrypt = zipCryptoDecryptImpl;
/** @deprecated Use {@link Crypto.encryptZipCrypto}. Removed in 1.0. */
export const zipCryptoEncrypt = zipCryptoEncryptImpl;
/** @deprecated Use {@link Crypto.AES_VENDOR_ID}. Removed in 1.0. */
export const AES_VENDOR_ID = AES_VENDOR_IDImpl;
/** @deprecated Use {@link Crypto.AES_VERSION_AE1}. Removed in 1.0. */
export const AES_VERSION_AE1 = AES_VERSION_AE1Impl;
/** @deprecated Use {@link Crypto.AES_VERSION_AE2}. Removed in 1.0. */
export const AES_VERSION_AE2 = AES_VERSION_AE2Impl;
/** @deprecated Use {@link Crypto.AES_EXTRA_FIELD_ID}. Removed in 1.0. */
export const AES_EXTRA_FIELD_ID = AES_EXTRA_FIELD_IDImpl;
/** @deprecated Use {@link Crypto.AES_SALT_LENGTH}. Removed in 1.0. */
export const AES_SALT_LENGTH = AES_SALT_LENGTHImpl;
/** @deprecated Use {@link Crypto.AES_KEY_LENGTH}. Removed in 1.0. */
export const AES_KEY_LENGTH = AES_KEY_LENGTHImpl;
/** @deprecated Use {@link Crypto.AES_AUTH_CODE_LENGTH}. Removed in 1.0. */
export const AES_AUTH_CODE_LENGTH = AES_AUTH_CODE_LENGTHImpl;
/** @deprecated Use {@link Crypto.AES_PASSWORD_VERIFY_LENGTH}. Removed in 1.0. */
export const AES_PASSWORD_VERIFY_LENGTH = AES_PASSWORD_VERIFY_LENGTHImpl;
/** @deprecated Use {@link Crypto.COMPRESSION_METHOD_AES}. Removed in 1.0. */
export const COMPRESSION_METHOD_AES = COMPRESSION_METHOD_AESImpl;
/** @deprecated Use {@link Crypto.decryptAes}. Removed in 1.0. */
export const aesDecrypt = aesDecryptImpl;
/** @deprecated Use {@link Crypto.encryptAes}. Removed in 1.0. */
export const aesEncrypt = aesEncryptImpl;
/** @deprecated Use {@link Crypto.getAesEncryptedSize}. Removed in 1.0. */
export const aesEncryptedSize = aesEncryptedSizeImpl;
/** @deprecated Use {@link Crypto.createAesExtraField}. Removed in 1.0. */
export const buildAesExtraField = buildAesExtraFieldImpl;
/** @deprecated Use {@link Crypto.randomBytes}. Removed in 1.0. */
export const randomBytes = randomBytesImpl;
/** @deprecated Use {@link Crypto.getEncryptionMethodName}. Removed in 1.0. */
export const getEncryptionMethodName = getEncryptionMethodNameImpl;
/** @deprecated Use {@link Crypto.isAesEncryption}. Removed in 1.0. */
export const isAesEncryption = isAesEncryptionImpl;
/** @deprecated Use {@link Crypto.getAesKeyStrength}. Removed in 1.0. */
export const getAesKeyStrength = getAesKeyStrengthImpl;
/** @deprecated Use {@link Crypto.getEncryptionMethodForAesKeyStrength}. Removed in 1.0. */
export const encryptionMethodFromAesKeyStrength = encryptionMethodFromAesKeyStrengthImpl;
