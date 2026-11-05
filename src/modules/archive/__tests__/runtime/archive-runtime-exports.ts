/**
 * Members expected on the `Archive` domain namespace
 * (`export * as Archive`). Platform-agnostic high-level API, available in all
 * environments (Node.js + browser).
 */
export const ARCHIVE_NAMESPACE_EXPORTS = [
  // ArchiveSource adapters
  "toAsyncIterable",
  "toReadableStream",

  // Format matrix — builder, reader, one-shot pack/unpack, entry lists
  "createZip",
  "openZip",
  "packZip",
  "packZipSync",
  "unpackZip",
  "unpackZipSync",
  "createTar",
  "openTar",
  "packTar",
  "packTarSync",
  "unpackTar",
  "unpackTarSync",
  "parseTar",
  "parseTarStream",
  "createTarGz",
  "openTarGz",
  "packTarGz",
  "packTarGzSync",
  "unpackTarGz",
  "unpackTarGzSync",
  "parseTarGz",
  "parseTarGzStream",

  // Classes behind the factories
  "ZipArchive",
  "ZipReader",
  "TarArchive",
  "TarReader",
  "TarGzArchive",

  // ZIP editing and entries
  "ZipEditor",
  "ZipEditPlan",
  "UnzipEntry",

  // TAR entries and constants
  "TAR_BLOCK_SIZE",
  "TarType",
  "createTarReaderEntry",
  "isTarFile",
  "isTarDirectory",
  "isTarSymlink",
  "isTarHardLink",
  "isTarDataEntry",

  // Deprecated forwards
  "zip",
  "unzip",
  "editZip",
  "editZipUrl",
  "createTarArchive",
  "createTarReader",
  "tar",
  "tarSync",
  "untar",
  "TAR_TYPE",
  "createAbortError",
  "isAbortError",
  "base64ToUint8Array",
  "uint8ArrayToBase64",
  "concatUint8Arrays",
  "stringToUint8Array",
  "uint8ArrayToString",

  // CRC-32 and PNG encoding
  "crc32",
  "crc32Update",
  "crc32Finalize",
  "encodePng",

  // Random access / HTTP Range reading
  "HttpRangeReader",
  "Crypto",
  "createBufferReader",
  "RemoteZipReader",

  // Abort utility
  "throwIfAborted",

  // Crypto - ZipCrypto
  "zipCryptoInitKeys",
  "zipCryptoDecrypt",
  "zipCryptoEncrypt",
  "ZIP_CRYPTO_HEADER_SIZE",

  // Crypto - AES
  "aesDecrypt",
  "aesEncrypt",
  "aesEncryptedSize",
  "buildAesExtraField",
  "randomBytes",
  "getAesKeyStrength",
  "encryptionMethodFromAesKeyStrength",
  "getEncryptionMethodName",
  "isAesEncryption",

  // AES constants
  "AES_AUTH_CODE_LENGTH",
  "AES_EXTRA_FIELD_ID",
  "AES_KEY_LENGTH",
  "AES_PASSWORD_VERIFY_LENGTH",
  "AES_SALT_LENGTH",
  "AES_VENDOR_ID",
  "AES_VERSION_AE1",
  "AES_VERSION_AE2",
  "COMPRESSION_METHOD_AES"
] as const;

/**
 * Core top-level archive exports available in all environments
 * (Node.js + browser). The high-level API now lives on the `Archive`
 * namespace; the top level keeps the domain namespace, platform-specific
 * compression/CRC primitives, and error classes.
 */
export const ARCHIVE_BROWSER_EXPORTS = [
  // Domain namespaces
  "Archive",
  "Compression",

  // Everything below except constants and errors is a deprecated forward to a
  // namespace member.

  // CRC32 (platform-specific implementation)
  "crc32",
  "crc32Update",
  "crc32Finalize",

  // Compression
  "compress",
  "compressSync",
  "decompress",
  "decompressSync",
  "hasCompressionStream",
  "hasWorkerSupport",

  // Streaming compression
  "createDeflateStream",
  "createInflateStream",
  "createGzipStream",
  "createGunzipStream",
  "createZlibStream",
  "createUnzlibStream",
  "hasDeflateRaw",
  "hasGzipCompressionStream",
  "hasGzipDecompressionStream",
  "isGzipData",

  // GZIP compression (available in both Node.js and browser)
  "gzip",
  "gunzip",
  "gzipSync",
  "gunzipSync",
  "GZIP_ID1",
  "GZIP_ID2",

  // Zlib compression (RFC 1950)
  "zlib",
  "unzlib",
  "zlibSync",
  "unzlibSync",
  "isZlibData",
  "ZLIB_CM_DEFLATE",
  "ZLIB_CINFO_MAX",
  "ZLIB_MIN_SIZE",

  // Auto-detect decompression
  "decompressAuto",
  "decompressAutoSync",
  "detectCompressionFormat",

  // PNG encoding — a DEFLATE stream plus CRC-32-checked chunks, so it lives
  // beside the two primitives it is built from.
  "encodePng",

  // Worker Pool (browser-only functionality, stub in Node)
  "WorkerPool",
  "getDefaultWorkerPool",
  "terminateDefaultWorkerPool",
  "deflateWithPool",
  "inflateWithPool",
  "deflateBatchWithPool",
  "inflateBatchWithPool",

  // Error types
  "ArchiveError",
  "AbortError",
  "isAbortError",
  "ZipParseError",
  "ZipInvalidSignatureError",
  "InvalidZipSignatureError",
  "ZipEocdNotFoundError",
  "ZipCrc32MismatchError",
  "ZipEntrySizeMismatchError",
  "ArchiveLimitError",
  "ArchiveWorkerStreamStartError",
  "ArchiveWorkerUnsupportedError",
  "ArchiveArgumentError",
  "ZipDecryptionError",
  "ZipPasswordRequiredError",
  "ArchiveRangeNotSupportedError",
  "ArchiveHttpRangeError",
  "ArchiveFileTooLargeError",
  "ArchiveFileError",
  "ZipUnsupportedCompressionError",
  "isArchiveError",

  // Deprecated error names
  "EocdNotFoundError",
  "Crc32MismatchError",
  "EntrySizeMismatchError",
  "WorkerStreamStartError",
  "DecryptionError",
  "PasswordRequiredError",
  "RangeNotSupportedError",
  "HttpRangeError",
  "FileTooLargeError",
  "UnsupportedCompressionError"
] as const;

/**
 * File system convenience layer exports (Node.js only)
 */
export const ARCHIVE_NODE_EXPORTS = [
  // File-system namespace
  "Fs",

  // Node stream adapter
  "toNodeReadable",

  "ArchiveFile",
  "traverseDirectory",
  "traverseDirectorySync",
  "glob",
  "globSync",
  "globToRegex",
  "matchGlob",
  "matchGlobAny",
  "ensureDir",
  "ensureDirSync",
  "fileExists",
  "fileExistsSync",
  "readFileBytes",
  "readFileBytesSync",
  "writeFileBytes",
  "writeFileBytesSync",
  "setFileTime",
  "setFileTimeSync",
  "safeStats",
  "safeStatsSync",
  "readFileText",
  "readFileTextSync",
  "writeFileText",
  "writeFileTextSync",
  "remove",
  "removeSync",
  "copyFile",
  "copyFileSync",
  "createReadStream",
  "createWriteStream",
  "createTempDir",
  "createTempDirSync",

  // TAR + Gzip support (Node.js only - requires zlib for streaming)
  "TarGzArchive",
  "targz",
  "parseTarGz",
  "parseTarGzStream",
  "untargz"
] as const;

/**
 * All archive exports (Node.js environment)
 */
export const ARCHIVE_RUNTIME_EXPORTS = [
  ...ARCHIVE_BROWSER_EXPORTS,
  ...ARCHIVE_NODE_EXPORTS
] as const;

/** Members only the Node entry's `Archive` namespace carries. */
export const ARCHIVE_NODE_NAMESPACE_EXPORTS = ["ArchiveFile", "toNodeReadable"] as const;

/** Members of the `Compression` namespace (identical on both entries). */
export const COMPRESSION_NAMESPACE_EXPORTS = [
  "deflate",
  "deflateSync",
  "inflate",
  "inflateSync",
  "gzip",
  "gunzip",
  "gzipSync",
  "gunzipSync",
  "isGzipData",
  "zlib",
  "unzlib",
  "zlibSync",
  "unzlibSync",
  "isZlibData",
  "decompressAuto",
  "decompressAutoSync",
  "detectFormat",
  "hasCompressionStream",
  "hasWorkerSupport",
  "hasGzipCompressionStream",
  "hasGzipDecompressionStream",
  "createDeflateStream",
  "createInflateStream",
  "createGzipStream",
  "createGunzipStream",
  "createZlibStream",
  "createUnzlibStream",
  "hasDeflateRaw",
  "Workers"
] as const;

export type ArchiveNamespaceExport = (typeof ARCHIVE_NAMESPACE_EXPORTS)[number];
export type ArchiveBrowserExport = (typeof ARCHIVE_BROWSER_EXPORTS)[number];
export type ArchiveNodeExport = (typeof ARCHIVE_NODE_EXPORTS)[number];
export type ArchiveRuntimeExport = (typeof ARCHIVE_RUNTIME_EXPORTS)[number];

export function getRuntimeExportKeys(moduleNamespace: object): string[] {
  return Object.keys(moduleNamespace).sort();
}

/** Members of the Node-only `Fs` namespace. */
export const FS_NAMESPACE_EXPORTS = [
  "traverseDirectory",
  "traverseDirectorySync",
  "glob",
  "globSync",
  "globToRegex",
  "matchGlob",
  "matchGlobAny",
  "ensureDir",
  "ensureDirSync",
  "fileExists",
  "fileExistsSync",
  "readFileBytes",
  "readFileBytesSync",
  "writeFileBytes",
  "writeFileBytesSync",
  "setFileTime",
  "setFileTimeSync",
  "safeStats",
  "safeStatsSync",
  "readFileText",
  "readFileTextSync",
  "writeFileText",
  "writeFileTextSync",
  "remove",
  "removeSync",
  "copyFile",
  "copyFileSync",
  "createTempDir",
  "createTempDirSync"
] as const;
