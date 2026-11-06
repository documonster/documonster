/**
 * Unified ArchiveFile class supporting both ZIP and TAR formats.
 *
 * The class provides a common API for creating, reading, and modifying archives,
 * with format-specific features available based on the `format` option.
 *
 * @example Create a ZIP archive (default)
 * ```ts
 * const archive = new ArchiveFile();
 * archive.addFile("./readme.md");
 * archive.addDirectory("./src");
 * await archive.writeFile("./output.zip");
 * ```
 *
 * @example Create a TAR archive
 * ```ts
 * const archive = new ArchiveFile({ format: "tar" });
 * archive.addFile("./readme.md");
 * archive.addDirectory("./src");
 * await archive.writeFile("./output.tar");
 * ```
 *
 * @example Create a gzipped TAR archive
 * ```ts
 * const archive = new ArchiveFile({ format: "tar", gzip: true });
 * archive.addFile("./readme.md");
 * await archive.writeFile("./output.tar.gz");
 * ```
 *
 * @module
 */

import { createReadStream, createWriteStream } from "node:fs";
import * as path from "node:path";

import { EMPTY_UINT8ARRAY } from "@archive/core/bytes";
import { DEFAULT_COMPRESS_LEVEL, REPRODUCIBLE_ZIP_MOD_TIME } from "@archive/core/defaults";
import {
  checkBytes,
  checkEntryName,
  checkEntryPath,
  checkOptionsObject,
  checkReadOptions,
  checkString,
  checkZipEntryOptions,
  checkZipOptions
} from "@archive/core/entry-checks";
import {
  ArchiveArgumentError,
  ArchiveError,
  ArchiveFileError,
  createAbortError
} from "@archive/core/errors";
import type { ZipStringEncoding } from "@archive/core/text";
import { decodeEntryText } from "@archive/core/text";
import type { ArchiveFormat } from "@archive/core/types";
import type { AesKeyStrength } from "@archive/crypto/aes-format";
import { aesEncryptedSize } from "@archive/crypto/aes-format";
import { ZIP_CRYPTO_HEADER_SIZE } from "@archive/crypto/zip-crypto";
import type {
  AddFileOptions,
  AddDirectoryOptions,
  AddGlobOptions,
  AddTarFileOptions,
  AddTarDirectoryOptions,
  AddTarGlobOptions,
  ExtractToOptions,
  ArchiveFileOptions,
  ArchiveFileOptionsZip,
  ArchiveFileOptionsTar,
  ArchiveFileOpenOptions,
  ArchiveFileOpenOptionsZip,
  ArchiveFileOpenOptionsTar,
  ArchiveFileReadOptions,
  ArchiveFileReadTextOptions,
  WriteZipOptions,
  WriteArchiveOptions,
  ArchiveEntryInfo,
  ZipEntryInfo,
  ZipFileOptions,
  OverwriteStrategy,
  ArchiveWarning,
  ArchiveStreamOptions,
  ArchiveStreamProgress,
  ArchiveStreamOperation,
  ArchiveTransformFunction,
  ArchiveTransformEntryData
} from "@archive/fs/types";
import type { ArchiveSink } from "@archive/io/archive-sink";
import { checkSinkArgument, pipeIterableToSink } from "@archive/io/archive-sink";
import type { ArchiveBytes, ArchiveSource } from "@archive/io/archive-source";
import {
  archiveBytesToUint8Array,
  collectUint8ArrayStream,
  isArchiveBytes,
  iterateArchiveSource
} from "@archive/io/archive-source";
import { gzipTarStream, gzipTarSync } from "@archive/tar/gzip-tar-stream";
import type { TarArchiveEntryOptions } from "@archive/tar/tar-archive";
import { TarArchive, TarReader } from "@archive/tar/tar-archive";
import { TAR_TYPE } from "@archive/tar/tar-constants";
import type { TarEntryInfo } from "@archive/tar/tar-entry-info";
import { isDirectory as isTarDirectory } from "@archive/tar/tar-entry-info";
import { gunzipTarSync } from "@archive/tar/tar-gzip";
import { parseTar, untarSync } from "@archive/tar/tar-parser";
import type { ZipEntryRecord } from "@archive/unzip/zip-parser";
import { ZipParser } from "@archive/unzip/zip-parser";
import type { ZipPathOptions } from "@archive/zip-spec/zip-path";
import { joinZipPath, normalizeZipPath } from "@archive/zip-spec/zip-path";
import { COMPRESSION_AES } from "@archive/zip-spec/zip-records";
import { ZipArchive } from "@archive/zip/zip-archive";
import type { ZipEntry, ZipOptions } from "@archive/zip/zip-bytes";
import { createZip, createZipSync } from "@archive/zip/zip-bytes";
import { isAsyncIterable, isReadableStream } from "@stream/core/type-guards";
import {
  concatUint8Arrays,
  textEncoder as utf8Encoder,
  asPlainBytes,
  plainChunks
} from "@utils/binary";
import type { FileEntry, GlobOptions, TraverseOptions } from "@utils/fs";
import {
  traverseDirectory,
  traverseDirectorySync,
  glob as globFiles,
  globSync as globFilesSync,
  ensureDir,
  ensureDirSync,
  fileExists,
  fileExistsSync,
  readFileBytes,
  readFileBytesSync,
  writeFileBytes,
  writeFileBytesSync,
  setFileTime,
  setFileTimeSync,
  safeStats,
  safeStatsSync,
  createSymlink,
  createSymlinkSync,
  chmod,
  chmodSync,
  supportsUnixPermissions
} from "@utils/fs";
import { isRecordArgument } from "@utils/object";

// =============================================================================
// Transform Helpers (internal)
// =============================================================================

function applyTransform(
  entry: FileEntry,
  prefix: string | undefined,
  transform?: ArchiveTransformFunction
): ArchiveTransformEntryData | null {
  const data: ArchiveTransformEntryData = {
    name: entry.relativePath,
    isDirectory: entry.isDirectory,
    size: entry.size,
    mtime: entry.mtime,
    atime: entry.atime,
    ctime: entry.ctime,
    birthTime: entry.birthTime,
    mode: entry.mode,
    prefix
  };
  if (!transform) {
    return data;
  }
  const result = transform(data);
  return result === false ? null : (result ?? data);
}

function mergeToZipOptions(t: ArchiveTransformEntryData, opts: AddFileOptions): AddFileOptions {
  return {
    ...opts,
    modTime: t.mtime,
    atime: t.atime,
    ctime: t.ctime,
    birthTime: t.birthTime,
    mode: t.mode ?? opts.mode
  };
}

function mergeToTarOptions(
  t: ArchiveTransformEntryData,
  opts: AddTarFileOptions
): AddTarFileOptions {
  return { ...opts, modTime: t.mtime, mode: t.mode ?? opts.mode };
}

// =============================================================================
// ZIP Helper Types
// =============================================================================

/*
 * Pending edits form a log replayed, in order, over the entries of the archive that was
 * opened (or last written) to give the archive a write would produce now — the "view". Every
 * read and every write goes through it. An entry added by name carries the time it was added
 * (`stamp`), which a write records when nothing else dates the entry, so a read before the
 * write and the write itself agree. A pending stream can be read only once: whichever read or
 * write reaches it first keeps its bytes in `buffered`, and everything after uses those.
 */

interface ZipPendingFileEntry {
  type: "file";
  localPath: string;
  zipPath: string;
  options: AddFileOptions;
  stamp: Date;
}

interface ZipPendingBufferEntry {
  type: "buffer";
  data: Uint8Array;
  zipPath: string;
  options: AddFileOptions;
  stamp: Date;
}

interface ZipPendingStreamEntry {
  type: "stream";
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>;
  /** The stream's bytes, once something has read it. */
  buffered?: Uint8Array;
  zipPath: string;
  options: AddFileOptions;
  stamp: Date;
}

interface ZipPendingDirectoryEntry {
  type: "directory";
  localPath: string;
  options: AddDirectoryOptions;
}

interface ZipPendingGlobEntry {
  type: "glob";
  pattern: string;
  options: AddGlobOptions;
}

interface ZipPendingSymlinkEntry {
  type: "symlink";
  zipPath: string;
  target: string;
  mode?: number;
  stamp: Date;
}

type ZipNamedPendingEntry =
  | ZipPendingFileEntry
  | ZipPendingBufferEntry
  | ZipPendingStreamEntry
  | ZipPendingSymlinkEntry;

type ZipPendingEntry =
  | ZipNamedPendingEntry
  | ZipPendingDirectoryEntry
  | ZipPendingGlobEntry
  | { type: "delete"; name: string }
  | { type: "rename"; from: string; to: string };

/** One entry of the archive a write would produce now. */
interface ZipViewEntry {
  /** The name the entry is written under. */
  name: string;
  isDirectory: boolean;
  /** An entry of the opened (or last written) archive, written as it is. */
  base?: ZipEntryRecord;
  /** An entry added by name. */
  op?: ZipNamedPendingEntry;
  /** A file (with its path) or a directory (`null`) found by `addDirectory` / `addGlob`. */
  found?: { path: string | null; build: (data: Uint8Array) => ZipEntry };
  /** The record a write stores for the entry, once computed. */
  record?: ZipEntryRecord;
}

// =============================================================================
// TAR Helper Types
// =============================================================================

type TarNamedPendingEntry =
  | { type: "file"; localPath: string; tarPath: string; options: AddTarFileOptions; stamp: Date }
  | { type: "buffer"; data: Uint8Array; tarPath: string; options: AddTarFileOptions; stamp: Date }
  | {
      type: "stream";
      stream: AsyncIterable<Uint8Array>;
      buffered?: Uint8Array;
      tarPath: string;
      options: AddTarFileOptions;
      stamp: Date;
    }
  | { type: "symlink"; tarPath: string; target: string; mode?: number; stamp: Date };

type TarPendingEntry =
  | TarNamedPendingEntry
  | { type: "directory"; localPath: string; options: AddTarDirectoryOptions }
  | { type: "glob"; pattern: string; options: AddTarGlobOptions };

/** One entry of the TAR a write would produce now. */
interface TarViewEntry {
  name: string;
  isDirectory: boolean;
  base?: { info: TarEntryInfo; data: Uint8Array };
  op?: TarNamedPendingEntry;
  found?: { path: string | null; options: TarArchiveEntryOptions };
  /** The header a write stores for the entry, once computed. */
  info?: TarEntryInfo;
}

/** A directory walk or a glob that resolving a view needs answered, synchronously or not. */
type FsQuery =
  | { kind: "directory"; path: string; options: TraverseOptions }
  | { kind: "glob"; pattern: string; options: GlobOptions };

/** Resolving a view: yields the walks it needs and receives what they found. */
type ViewResolution<V> = Generator<FsQuery, V, FileEntry[]>;

function resolveViewSync<V>(steps: ViewResolution<V>): V {
  let step = steps.next();
  while (!step.done) {
    const query = step.value;
    let found: FileEntry[];
    try {
      found =
        query.kind === "directory"
          ? traverseDirectorySync(query.path, query.options)
          : globFilesSync(query.pattern, query.options);
    } catch (error) {
      throw queryError(query, error);
    }
    step = steps.next(found);
  }
  return step.value;
}

async function resolveViewAsync<V>(steps: ViewResolution<V>): Promise<V> {
  let step = steps.next();
  while (!step.done) {
    const query = step.value;
    const found: FileEntry[] = [];
    try {
      for await (const entry of query.kind === "directory"
        ? traverseDirectory(query.path, query.options)
        : globFiles(query.pattern, query.options)) {
        found.push(entry);
      }
    } catch (error) {
      throw queryError(query, error);
    }
    step = steps.next(found);
  }
  return step.value;
}

/** What extracting one ZIP entry needs: its record and a reader for its data. */
interface ZipExtractItem {
  record: ZipEntryRecord;
  read: (password?: string | Uint8Array) => Promise<Uint8Array | null>;
  readSync: (password?: string | Uint8Array) => Uint8Array | null;
}

/** Whether the record of a ZIP view entry is known without reading its data. */
function hasKnownRecord(v: ZipViewEntry): boolean {
  return v.record !== undefined || v.base !== undefined;
}

/** Whether the header of a TAR view entry is known without reading its data. */
function hasKnownTarInfo(v: TarViewEntry): boolean {
  return v.info !== undefined || v.base !== undefined;
}

/** The order a ZIP write stores entries in: by lower-cased name, ties in view order. */
function byZipWriteOrder(a: ZipViewEntry, b: ZipViewEntry): number {
  const x = a.name.toLowerCase();
  const y = b.name.toLowerCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

/** The error for a synchronous read or write that needs a pending stream nothing has read. */
function unreadStreamError(name: string): ArchiveError {
  return new ArchiveError(
    `Entry "${name}" is a pending stream that has not been read yet, which cannot be done ` +
      "synchronously; use the async form (toBuffer(), readEntry(), getEntries(), extractTo()) first"
  );
}

/**
 * A pending stream as a source a streamed write can read: the bytes already buffered, or the
 * stream itself, kept as it passes so a later read or write still has it.
 */
async function* readPendingStreamOnce(op: {
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>;
  buffered?: Uint8Array;
}): AsyncIterable<Uint8Array> {
  if (op.buffered) {
    yield op.buffered;
    return;
  }
  const chunks: Uint8Array[] = [];
  for await (const chunk of iterateArchiveSource(op.stream)) {
    chunks.push(chunk);
    yield chunk;
  }
  op.buffered = concatUint8Arrays(chunks);
}

// =============================================================================
// ZIP Helper Functions
// =============================================================================

/**
 * Resolve effective ZIP path options for an operation.
 */
function resolveZipPathOptions(globalOptions: ZipFileOptions): ZipPathOptions {
  return {
    mode: "legacy",
    ...(globalOptions.path ?? {})
  };
}

/**
 * The name a ZIP entry is written under, checked as `ZipArchive.add` checks it.
 *
 * With no `path` option an absolute name, a drive prefix or a `..` that climbs above the
 * archive root is refused (see `checkEntryName`); an explicit `path` mode decides instead,
 * which is how a caller who means to write such a name says so.
 */
function zipWritePath(where: string, globalOptions: ZipFileOptions, ...parts: string[]): string {
  const pathOptions = resolveZipPathOptions(globalOptions);
  if (globalOptions.path === undefined) {
    checkEntryName(`Archive.ArchiveFile#${where}`, parts.filter(part => part).join("/"));
  }
  return parts.length === 1
    ? normalizeZipPath(parts[0]!, pathOptions)
    : joinZipPath(pathOptions, ...parts);
}

type ZipModeOptions = { mode?: number };

function resolveEntryMode(
  kind: "file" | "directory",
  globalOptions: ZipFileOptions,
  localOptions?: ZipModeOptions,
  fsMode?: number
): number | undefined {
  if (!(globalOptions.writePermissions ?? false)) {
    return undefined;
  }

  if (localOptions?.mode !== undefined) {
    return localOptions.mode;
  }

  if ((globalOptions.preservePermissions ?? false) && fsMode !== undefined) {
    return fsMode;
  }

  if (kind === "directory") {
    return 0o040755;
  }
  return 0o100644;
}

function buildDirectoryEntry(
  zipPath: string,
  fsEntry: FileEntry,
  globalOptions: ZipFileOptions,
  localOptions: AddDirectoryOptions
): ZipEntry {
  return {
    name: zipPath + "/",
    data: EMPTY_UINT8ARRAY,
    level: 0,
    modTime: fsEntry.mtime,
    atime: fsEntry.atime,
    ctime: fsEntry.ctime,
    birthTime: fsEntry.birthTime,
    mode: resolveEntryMode("directory", globalOptions, localOptions, fsEntry.mode),
    msDosAttributes: localOptions.msDosAttributes,
    encoding: localOptions.encoding ?? globalOptions.encoding
  };
}

/**
 * Check for path traversal attack and throw if detected.
 */
function assertNoPathTraversal(targetPath: string, baseDir: string, entryPath: string): void {
  if (!targetPath.startsWith(baseDir + path.sep) && targetPath !== baseDir) {
    throw new ArchiveError(`Path traversal detected: ${entryPath}`);
  }
}

/**
 * Get effective Unix mode for extraction.
 * If the archive entry has no mode info (mode=0), returns sensible defaults.
 *
 * @param mode - The mode from the archive entry
 * @param kind - "file" or "directory"
 * @returns Effective mode with permission bits
 */
function getEffectiveMode(mode: number, kind: "file" | "directory"): number {
  // Extract permission bits (lower 12 bits: rwx for owner/group/other + sticky/setuid/setgid)
  const permBits = mode & 0o7777;

  if (permBits !== 0) {
    // Entry has permission info
    return permBits;
  }

  // No permission info (e.g., Windows-created ZIP) - use defaults
  // Default: directories 0o755 (rwxr-xr-x), files 0o644 (rw-r--r--)
  return kind === "directory" ? 0o755 : 0o644;
}

/**
 * Assert that the archive format is ZIP, throw a descriptive error otherwise.
 */
function assertZipFormat(format: ArchiveFormat, methodName: string): asserts format is "zip" {
  if (format !== "zip") {
    throw new ArchiveError(`${methodName} is only available for ZIP archives`);
  }
}

/** `readEntry`'s options, from either the options object or the positional password. */
function toReadOptions(
  options: ArchiveFileReadOptions | string | Uint8Array | undefined
): ArchiveFileReadOptions {
  // Anything but an options object is the positional password, `null` included, which then falls
  // back to the archive's password exactly as the positional form did.
  return isRecordArgument(options) ? options : { password: options };
}

/** `readAsText`'s options, from either the options object or the positional encoding. */
function toReadTextOptions(
  options: ArchiveFileReadTextOptions | string | undefined
): ArchiveFileReadTextOptions {
  // Anything but an options object is the positional encoding; only `undefined` defaults it.
  return isRecordArgument(options) ? options : { encoding: options };
}

const DEFAULT_IO_CONCURRENCY = 8;

function isIgnorableFsError(err: unknown): boolean {
  const code = (err as { code?: unknown } | null | undefined)?.code;
  return code === "ENOENT" || code === "EACCES" || code === "EPERM";
}

type WarningCallback = ((warning: ArchiveWarning) => void) | undefined;

function emitExtractWarning(
  onWarning: WarningCallback,
  entryPath: string,
  targetPath: string,
  err: unknown
): void {
  if (!onWarning) {
    return;
  }
  const code = (err as { code?: unknown } | null | undefined)?.code;
  const errMessage = err instanceof Error ? err.message : String(err);
  // For filesystem errors with a code, use a generic message; otherwise preserve the original
  const message =
    code != null
      ? `Skipping extraction due to filesystem error (${String(code)})`
      : errMessage || "Skipping extraction due to error";
  onWarning({ operation: "extract", entryPath, targetPath, message, error: err });
}

/**
 * Try a filesystem operation, emit warning and return false if it's an ignorable error.
 * Throws if the error is not ignorable.
 */
async function tryFsOpWithWarning(
  fn: () => Promise<void>,
  onWarning: WarningCallback,
  entryPath: string,
  targetPath: string
): Promise<boolean> {
  try {
    await fn();
    return true;
  } catch (err) {
    if (isIgnorableFsError(err)) {
      emitExtractWarning(onWarning, entryPath, targetPath, err);
      return false;
    }
    throw toFileError(targetPath, "write", err, "extracted");
  }
}

/**
 * Synchronous version of tryFsOpWithWarning.
 */
function tryFsOpWithWarningSync(
  fn: () => void,
  onWarning: WarningCallback,
  entryPath: string,
  targetPath: string
): boolean {
  try {
    fn();
    return true;
  } catch (err) {
    if (isIgnorableFsError(err)) {
      emitExtractWarning(onWarning, entryPath, targetPath, err);
      return false;
    }
    throw toFileError(targetPath, "write", err, "extracted");
  }
}

async function processInOrderWithConcurrency<T>(
  iterable: AsyncIterable<T> | Iterable<T>,
  concurrency: number,
  task: (item: T) => Promise<() => void>
): Promise<void> {
  const inFlight = new Map<number, Promise<() => void>>();
  let index = 0;
  let next = 0;

  for await (const item of iterable) {
    const current = index++;
    inFlight.set(current, task(item));

    while (inFlight.size >= concurrency) {
      const apply = await inFlight.get(next)!;
      inFlight.delete(next);
      next++;
      apply();
    }
  }

  while (next < index) {
    const apply = await inFlight.get(next)!;
    inFlight.delete(next);
    next++;
    apply();
  }
}

/** Default mode for TAR directories */
const TAR_DIR_MODE = 0o755;

/**
 * Core logic for shouldExtract - shared between async and sync versions.
 */
function shouldExtractCore(
  exists: boolean,
  entryMtime: Date,
  strategy: OverwriteStrategy,
  targetPath: string,
  getStats: () => { mtime: Date } | undefined
): boolean {
  if (!exists) {
    return true;
  }

  switch (strategy) {
    case "skip":
      return false;

    case "overwrite":
      return true;

    case "error":
      throw new ArchiveError(`File already exists: ${targetPath}`);

    case "newer": {
      const stats = getStats();
      if (!stats) {
        return true;
      }
      return entryMtime > stats.mtime;
    }

    default:
      throw new ArchiveArgumentError("overwrite", `Unknown overwrite strategy: ${strategy}`);
  }
}

/**
 * Check if extraction should proceed based on overwrite strategy.
 */
async function shouldExtract(
  targetPath: string,
  entryMtime: Date,
  strategy: OverwriteStrategy
): Promise<boolean> {
  const exists = await fileExists(targetPath);
  if (strategy === "newer" && exists) {
    const stats = await safeStats(targetPath);
    return shouldExtractCore(exists, entryMtime, strategy, targetPath, () => stats);
  }
  return shouldExtractCore(exists, entryMtime, strategy, targetPath, () => undefined);
}

/**
 * Synchronous version of shouldExtract.
 */
function shouldExtractSync(
  targetPath: string,
  entryMtime: Date,
  strategy: OverwriteStrategy
): boolean {
  const exists = fileExistsSync(targetPath);
  return shouldExtractCore(exists, entryMtime, strategy, targetPath, () =>
    safeStatsSync(targetPath)
  );
}

/**
 * Collect all chunks from an async iterable or ReadableStream.
 */
async function collectStream(
  stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>
): Promise<Uint8Array> {
  return collectUint8ArrayStream(stream);
}

/**
 * Convert AddDirectoryOptions/AddGlobOptions filter to FileEntry filter.
 */
type TraverseFilter = (entry: {
  relativePath: string;
  isDirectory: boolean;
  size: number;
}) => boolean;

function wrapFilter(
  filter: ((path: string, stats: { isDirectory: boolean; size: number }) => boolean) | undefined
): TraverseFilter | undefined {
  return filter
    ? e => filter(e.relativePath, { isDirectory: e.isDirectory, size: e.size })
    : undefined;
}

/**
 * Map a ZipParser entry to ArchiveEntryInfo format.
 */
function mapZipEntryToInfo(e: ZipEntryRecord): ZipEntryInfo {
  return {
    path: e.path,
    isDirectory: e.type === "directory",
    size: e.uncompressedSize,
    compressedSize: e.compressedSize,
    lastModified: e.lastModified,
    crc32: e.crc32,
    isEncrypted: e.isEncrypted,
    encryptionMethod:
      e.encryptionMethod === "aes"
        ? "aes"
        : e.encryptionMethod === "zipcrypto"
          ? "zipcrypto"
          : undefined,
    aesKeyStrength: e.aesKeyStrength,
    comment: e.comment
  };
}

/**
 * Build a ZipEntry from common parameters.
 */
function buildZipEntry(
  name: string,
  data: Uint8Array,
  entryOptions: AddFileOptions,
  globalOptions: ZipFileOptions,
  globalPassword: string | Uint8Array | undefined,
  /** The date recorded when neither the options nor the file system give one. */
  stamp: Date,
  fsMetadata?: {
    modTime?: Date;
    mode?: number;
    atime?: Date;
    ctime?: Date;
    birthTime?: Date;
  } | null
): ZipEntry {
  const mode = resolveEntryMode("file", globalOptions, entryOptions, fsMetadata?.mode);
  const externalAttributes = entryOptions.externalAttributes;

  return {
    name,
    data,
    level: entryOptions.level ?? globalOptions.level,
    modTime: entryOptions.modTime ?? fsMetadata?.modTime ?? stamp,
    atime: entryOptions.atime ?? fsMetadata?.atime,
    ctime: entryOptions.ctime ?? fsMetadata?.ctime,
    birthTime: entryOptions.birthTime ?? fsMetadata?.birthTime,
    comment: entryOptions.comment,
    encoding: entryOptions.encoding ?? globalOptions.encoding,
    encryptionMethod: entryOptions.encryptionMethod ?? globalOptions.encryptionMethod,
    password: entryOptions.password ?? globalPassword,
    mode,
    msDosAttributes: entryOptions.msDosAttributes,
    externalAttributes
  };
}

/**
 * Build a ZipEntry for preserving an existing entry (no update).
 */
function buildPreservedEntry(
  name: string,
  existingEntry: { lastModified: Date; comment: string; externalAttributes: number },
  data: Uint8Array,
  globalOptions: ZipFileOptions,
  globalPassword: string | Uint8Array | undefined
): ZipEntry {
  return {
    // The entry's current name: a rename moves it in the edit view, not in its record.
    name,
    data,
    level: globalOptions.level,
    modTime: existingEntry.lastModified,
    comment: existingEntry.comment,
    encoding: globalOptions.encoding,
    encryptionMethod: globalOptions.encryptionMethod,
    password: globalPassword,
    externalAttributes: existingEntry.externalAttributes
  };
}

/** One entry's data, decompressed from `parser` when it is first read. */
async function* extractLazily(
  parser: ZipParser,
  entryPath: string,
  password: string | Uint8Array | undefined
): AsyncIterable<Uint8Array> {
  const data = await parser.extract(entryPath, password);
  if (data && data.length > 0) {
    yield data;
  }
}

/**
 * Build a symlink ZipEntry.
 */
function buildSymlinkEntry(
  zipPath: string,
  target: string,
  mode: number | undefined,
  encoding: ZipStringEncoding | undefined,
  stamp: Date
): ZipEntry {
  return {
    name: zipPath,
    data: utf8Encoder.encode(target),
    level: 0,
    modTime: stamp,
    mode: mode ?? 0o120777,
    encoding
  };
}

type FileKind = "archive" | "source" | "extracted";

/** The `ArchiveFileError` for a file the system would not let us read or write, with its error as the cause. */
function fileSystemError(
  filePath: string,
  operation: "read" | "write",
  error: unknown,
  fileKind: FileKind = "archive"
): ArchiveFileError {
  const code = (error as { code?: unknown } | null)?.code;
  const detail =
    code === "ENOENT" ? "File not found" : error instanceof Error ? error.message : String(error);
  return new ArchiveFileError(filePath, operation, detail, { cause: error, fileKind });
}

/**
 * {@link fileSystemError} for a platform refusal, and `error` itself for anything else — a
 * step that both touches the disk and applies this module's own policy (an `overwrite: "error"`
 * refusal is an `ArchiveError` already) must not have its own errors relabelled.
 */
function toFileError(
  filePath: string,
  operation: "read" | "write",
  error: unknown,
  fileKind: FileKind = "archive"
): unknown {
  const code = (error as { code?: unknown } | null)?.code;
  return error instanceof Error && !(error instanceof ArchiveError) && typeof code === "string"
    ? fileSystemError(filePath, operation, error, fileKind)
    : error;
}

/** {@link toFileError} for a directory walk or glob, naming the path the system refused. */
function queryError(query: FsQuery, error: unknown): unknown {
  const refused = (error as { path?: unknown } | null)?.path;
  const queried = query.kind === "directory" ? query.path : query.pattern;
  return toFileError(typeof refused === "string" ? refused : queried, "read", error, "source");
}

/**
 * An archive stream whose local sources are opened as it runs: a source the system will not
 * read surfaces from the stream as an `ArchiveFileError` naming that source.
 */
async function* readingSources(chunks: AsyncIterable<Uint8Array>): AsyncGenerator<Uint8Array> {
  try {
    yield* chunks;
  } catch (error) {
    const refused = (error as { path?: unknown } | null)?.path;
    throw typeof refused === "string" ? toFileError(refused, "read", error, "source") : error;
  }
}

/** A local source's bytes; a file the system will not read is an `ArchiveFileError`. */
async function readSourceBytes(filePath: string): Promise<Uint8Array> {
  try {
    return await readFileBytes(filePath);
  } catch (error) {
    throw toFileError(filePath, "read", error, "source");
  }
}

/** {@link readSourceBytes}, synchronously. */
function readSourceBytesSync(filePath: string): Uint8Array {
  try {
    return readFileBytesSync(filePath);
  } catch (error) {
    throw toFileError(filePath, "read", error, "source");
  }
}

/**
 * Check overwrite strategy and return whether to proceed.
 */
function checkOverwriteStrategy(
  exists: boolean,
  targetPath: string,
  overwrite: OverwriteStrategy
): boolean {
  if (!exists) {
    return true;
  }

  switch (overwrite) {
    case "skip":
      return false;
    case "error":
      throw new ArchiveError(`File already exists: ${targetPath}`);
    case "overwrite":
    case "newer":
      return true;
    default:
      throw new ArchiveArgumentError("overwrite", `Unknown overwrite strategy: ${overwrite}`);
  }
}

// =============================================================================
// TAR Helper Functions
// =============================================================================

function normalizeTarPath(name: string, prefix?: string): string {
  let result = name.replace(/\\/g, "/");
  result = result.replace(/^\/+/, "");
  if (prefix) {
    let normalizedPrefix = prefix.replace(/\\/g, "/").replace(/^\/+/, "");
    // Trim trailing slashes without regex to avoid ReDoS on long '/' runs
    let end = normalizedPrefix.length;
    while (end > 0 && normalizedPrefix[end - 1] === "/") {
      end--;
    }
    normalizedPrefix = normalizedPrefix.slice(0, end);
    result = normalizedPrefix ? `${normalizedPrefix}/${result}` : result;
  }
  return result;
}
/**
 * Build TAR entry add options from pending entry options and optional file stats.
 */
function buildTarAddOptions(
  entryOptions: AddTarFileOptions,
  stats?: { mode?: number; mtime?: Date } | null
): { mode?: number; mtime?: Date; uid?: number; gid?: number; uname?: string; gname?: string } {
  return {
    mode: entryOptions.mode ?? stats?.mode,
    mtime: entryOptions.modTime ?? stats?.mtime,
    uid: entryOptions.uid,
    gid: entryOptions.gid,
    uname: entryOptions.uname,
    gname: entryOptions.gname
  };
}

/**
 * Build TAR symlink entry options.
 * Type "2" is the symlink type per RFC 1062.
 */
function buildTarSymlinkOptions(target: string, mode?: number) {
  return {
    type: TAR_TYPE.SYMLINK, // TAR symlink type (RFC 1062)
    linkname: target,
    mode
  };
}

/**
 * Wrap a streaming operation to conform to ArchiveStreamOperation interface.
 */
function wrapStreamOperation(
  op: {
    iterable: AsyncIterable<Uint8Array>;
    signal: AbortSignal;
    abort: (reason?: unknown) => void;
    pointer: () => number;
    progress: () => ArchiveStreamProgress;
  },
  overrideIterable?: AsyncIterable<Uint8Array>
): ArchiveStreamOperation {
  return {
    iterable: plainChunks(overrideIterable ?? op.iterable),
    signal: op.signal,
    abort: reason => op.abort(reason),
    pointer: () => op.pointer(),
    progress: () => op.progress()
  };
}

// =============================================================================
// View Helpers
// =============================================================================

/** The file-system facts a ZIP write records for a file added by path. */
function zipFsMetadata(stats: ReturnType<typeof safeStatsSync>) {
  return stats
    ? {
        modTime: stats.mtime,
        mode: stats.mode,
        atime: stats.atime,
        ctime: stats.ctime,
        birthTime: stats.birthtime
      }
    : null;
}

/** Apply a pending edit that names its entry to a ZIP view, as a write applies it. */
function applyZipOp(
  view: Map<string, ZipViewEntry>,
  op: Exclude<ZipPendingEntry, ZipPendingDirectoryEntry | ZipPendingGlobEntry>
): void {
  switch (op.type) {
    case "delete":
      view.delete(op.name);
      return;
    case "rename": {
      const entry = view.get(op.from);
      if (!entry) {
        return;
      }
      view.delete(op.to);
      view.delete(op.from);
      view.set(op.to, {
        ...entry,
        name: op.to,
        record: entry.record && { ...entry.record, path: op.to }
      });
      return;
    }
    default:
      view.set(op.zipPath, { name: op.zipPath, isDirectory: false, op });
  }
}

/** The walk an `addDirectory` / `addGlob` needs; a transform supersedes the filter. */
function fsQueryOf(
  op:
    | {
        type: "directory";
        localPath: string;
        options: AddDirectoryOptions | AddTarDirectoryOptions;
      }
    | { type: "glob"; pattern: string; options: AddGlobOptions | AddTarGlobOptions },
  defaultCwd?: string
): FsQuery {
  const { filter, transform, followSymlinks } = op.options;
  const wrapped = transform ? undefined : wrapFilter(filter);
  if (op.type === "directory") {
    return {
      kind: "directory",
      path: op.localPath,
      options: { recursive: op.options.recursive ?? true, followSymlinks, filter: wrapped }
    };
  }
  const { cwd, ignore, dot } = op.options;
  return {
    kind: "glob",
    pattern: op.pattern,
    options: { cwd: cwd ?? defaultCwd, ignore, dot, followSymlinks, filter: wrapped }
  };
}

/** The ZIP view entries an `addDirectory` / `addGlob` contributes, from what its walk found. */
function expandZipFound(
  op: ZipPendingDirectoryEntry | ZipPendingGlobEntry,
  found: readonly FileEntry[],
  globalOptions: ZipFileOptions,
  globalPassword: string | Uint8Array | undefined
): ZipViewEntry[] {
  const { transform } = op.options;
  const isDirectoryOp = op.type === "directory";
  const where = isDirectoryOp ? "addDirectory" : "addGlob";
  const basePrefix = isDirectoryOp
    ? (op.options.prefix ?? ((op.options.includeRoot ?? true) ? path.basename(op.localPath) : ""))
    : (op.options.prefix ?? "");
  const out: ZipViewEntry[] = [];
  for (const file of found) {
    const transformed = applyTransform(
      file,
      isDirectoryOp ? basePrefix : op.options.prefix,
      transform
    );
    if (transformed === null) {
      continue;
    }
    const zipPath = zipWritePath(
      where,
      globalOptions,
      transformed.prefix ?? basePrefix,
      transformed.name
    );
    const metadata = {
      modTime: transformed.mtime,
      mode: transformed.mode ?? file.mode,
      atime: transformed.atime ?? file.atime,
      ctime: transformed.ctime ?? file.ctime,
      birthTime: transformed.birthTime ?? file.birthTime
    };
    if (file.isDirectory) {
      const entry = buildDirectoryEntry(
        zipPath,
        { ...file, mtime: metadata.modTime, ...metadata },
        globalOptions,
        op.options as AddDirectoryOptions
      );
      out.push({ name: entry.name, isDirectory: true, found: { path: null, build: () => entry } });
      continue;
    }
    const options = transform ? mergeToZipOptions(transformed, op.options) : op.options;
    out.push({
      name: zipPath,
      isDirectory: false,
      found: {
        path: file.absolutePath,
        build: data =>
          buildZipEntry(
            zipPath,
            data,
            options,
            globalOptions,
            globalPassword,
            metadata.modTime,
            metadata
          )
      }
    });
  }
  return out;
}

/**
 * The central-directory record a ZIP write stores for `entry`. The entry is encoded on its own
 * with the archive's settings and parsed back, so sizes, dates and attributes are the written
 * ones. Encryption only adds a fixed-size header (and for AES a trailer and a zero CRC), so the
 * entry is encoded plainly — synchronously, even for AES — and that is added. `store` skips
 * compression when the caller needs everything but the compressed size.
 */
function recordOfPendingZipEntry(
  name: string,
  entry: ZipEntry,
  options: ZipOptions,
  store = false
): ZipEntryRecord {
  const method = entry.encryptionMethod ?? "none";
  const encrypted = method !== "none" && !!entry.password && entry.password.length > 0;
  const bytes = createZipSync(
    [
      {
        ...entry,
        level: store ? 0 : entry.level,
        encryptionMethod: "none",
        password: undefined
      }
    ],
    { ...options, comment: undefined, encryptionMethod: "none", password: undefined }
  );
  const parsed = new ZipParser(bytes, { encoding: entry.encoding, checkCrc32: false });
  const record: ZipEntryRecord = { ...parsed.getEntries()[0]!, path: name };
  if (!encrypted) {
    return record;
  }
  if (method === "zipcrypto") {
    return {
      ...record,
      compressedSize: record.compressedSize + ZIP_CRYPTO_HEADER_SIZE,
      isEncrypted: true,
      encryptionMethod: "zipcrypto"
    };
  }
  const strength = Number(method.slice("aes-".length)) as AesKeyStrength;
  return {
    ...record,
    compressedSize: aesEncryptedSize(record.compressedSize, strength),
    crc32: 0,
    compressionMethod: COMPRESSION_AES,
    originalCompressionMethod: record.compressionMethod,
    isEncrypted: true,
    encryptionMethod: "aes",
    aesVersion: 2,
    aesKeyStrength: strength
  };
}

/** The per-entry options a streaming `ZipArchive` takes for what a buffered write records. */
function zipStreamOptions(entry: ZipEntry) {
  return {
    level: entry.level,
    modTime: entry.modTime,
    atime: entry.atime,
    ctime: entry.ctime,
    birthTime: entry.birthTime,
    comment: entry.comment,
    mode: entry.mode,
    msDosAttributes: entry.msDosAttributes,
    externalAttributes: entry.externalAttributes,
    encoding: entry.encoding
  };
}

/** Whether a ZIP write encrypts `entry`. */
function isEncryptedZipEntry(entry: ZipEntry): boolean {
  return (entry.encryptionMethod ?? "none") !== "none" && !!entry.password;
}

/** The TAR view entries an `addDirectory` / `addGlob` contributes, from what its walk found. */
function expandTarFound(
  op: Extract<TarPendingEntry, { type: "directory" | "glob" }>,
  found: readonly FileEntry[]
): TarViewEntry[] {
  const { transform, prefix } = op.options;
  const out: TarViewEntry[] = [];
  for (const file of found) {
    const transformed = applyTransform(file, prefix, transform);
    if (transformed === null) {
      continue;
    }
    const tarPath = normalizeTarPath(transformed.name, transformed.prefix ?? prefix);
    if (file.isDirectory) {
      out.push({
        name: tarPath + "/",
        isDirectory: true,
        found: { path: null, options: { mode: TAR_DIR_MODE, mtime: transformed.mtime } }
      });
      continue;
    }
    const options = transform ? mergeToTarOptions(transformed, op.options) : op.options;
    out.push({
      name: tarPath,
      isDirectory: false,
      found: {
        path: file.absolutePath,
        options: buildTarAddOptions(options, {
          mtime: transformed.mtime,
          mode: transformed.mode ?? file.mode
        })
      }
    });
  }
  return out;
}

/** The header a TAR write stores for an entry, by encoding it on its own and parsing it back. */
function infoOfPendingTarEntry(
  name: string,
  data: Uint8Array,
  options: TarArchiveEntryOptions,
  modTime: Date | undefined
): TarEntryInfo {
  const archive = new TarArchive({ modTime });
  archive.add(name, data, options);
  return parseTar(archive.toBufferSync(), { maxEntrySize: Infinity })[0]!.info;
}

/** A TAR header as `getEntries` reports it. */
function toTarEntryInfo(info: TarEntryInfo, name = info.path): ArchiveEntryInfo {
  return {
    path: name,
    size: info.size,
    mtime: info.mtime,
    mode: info.mode,
    uid: info.uid,
    gid: info.gid,
    uname: info.uname,
    gname: info.gname,
    linkname: info.linkname,
    type: info.type,
    isDirectory: isTarDirectory(info)
  };
}

/** The add options that copy an entry of the opened TAR unchanged. */
function tarCopyOptions(info: TarEntryInfo): TarArchiveEntryOptions {
  return {
    type: info.type,
    mode: info.mode,
    uid: info.uid,
    gid: info.gid,
    uname: info.uname,
    gname: info.gname,
    mtime: info.mtime,
    linkname: info.linkname
  };
}

// =============================================================================
// ArchiveFile Class
// =============================================================================

/**
 * Non-conditional view of the ZIP-mode mutable state fields of {@link ArchiveFile}.
 *
 * The class declares these fields conditionally on the format parameter `F`
 * (`F extends "zip" ? T : null`). This interface mirrors the resolved `"zip"`
 * branch so the internal `_setZipState` setter can type its `value` argument
 * precisely without the format generic getting in the way.
 */
interface ArchiveZipState {
  _zipData: Uint8Array | null;
  _zipParser: ZipParser | null;
  _zipPassword: string | Uint8Array | undefined;
  _zipAbortController: AbortController | null;
  _zipBytesWritten: number;
  _zipPendingEntries: ZipPendingEntry[];
  _zipSourcePath: string | null;
}

/**
 * Unified archive file class supporting both ZIP and TAR formats.
 *
 * This class provides file system integration for creating and reading archives.
 *
 * Edits (`add*`, `set`, `delete`, `rename`, `setZipComment`) are pending until the archive is
 * written (`toBuffer`, `writeFile`, …), and every read (`has`, `entryCount`, `getEntryNames`,
 * `getEntries`, `getEntry`, `readEntry`, `readAsText`, `extractTo`, …) answers for the archive
 * a write would produce now: the opened (or last written) archive's entries with the edits
 * applied, in the order the write stores them, each described as the write records it. A
 * buffered write then becomes what the next one starts from; a streamed write (`stream`,
 * `pipeTo`, `streamToFile`) writes the same entries but leaves the edits pending.
 *
 * A read touches only what it needs. Listing names reads no entry's data (a pending
 * `addDirectory` / `addGlob` is walked to find its names); reading or describing a pending
 * entry reads it from where it is pending — memory, or the file on disk at that moment. A
 * pending stream can be consumed only once, so the first read or write that needs it keeps
 * its bytes in memory for everything after: that costs the stream's full size in memory
 * until the archive is written and the edit is no longer pending. A synchronous read cannot
 * consume a stream and throws until an async one has. Describing a pending ZIP entry
 * (`getEntry`, `getEntries`) compresses it as the write would, to report its sizes.
 *
 * @template F - Archive format ("zip" or "tar")
 */
export class ArchiveFile<F extends ArchiveFormat = "zip"> {
  private readonly _format: F;

  // ===========================================================================
  // ZIP State
  // ===========================================================================
  private _zipOptions: F extends "zip" ? ZipFileOptions : null;
  private _zipPendingEntries: F extends "zip" ? ZipPendingEntry[] : null;
  private _zipData: F extends "zip" ? Uint8Array | null : null;
  private _zipParser: F extends "zip" ? ZipParser | null : null;
  private _zipSourcePath: F extends "zip" ? string | null : null;
  private _zipPassword: F extends "zip" ? string | Uint8Array | undefined : null;
  private _zipAbortController: F extends "zip" ? AbortController | null : null;
  private _zipBytesWritten: F extends "zip" ? number : null;

  // ===========================================================================
  // TAR State
  // ===========================================================================
  /** The uncompressed TAR this archive was opened from or last written as; its entries are written first. */
  private _tarSource: F extends "tar" ? Uint8Array | null : null;
  private _tarReader: F extends "tar" ? TarReader | null : null;
  private readonly _tarPendingEntries: F extends "tar" ? TarPendingEntry[] : null;
  private readonly _tarOptions: F extends "tar" ? ArchiveFileOptionsTar : null;

  // ===========================================================================
  // View State
  // ===========================================================================
  /** The resolved view of a ZIP archive, kept until an edit it cannot absorb. */
  private _zipView: Map<string, ZipViewEntry> | null = null;
  /** The resolved view of a TAR archive, kept until the next edit. */
  private _tarView: Map<string, TarViewEntry> | null = null;
  /** Bumped by every edit, so a view resolved across an edit is not kept. */
  private _viewGeneration = 0;
  /** `setZipComment` changed the comment since the archive was opened or written. */
  private _zipCommentChanged = false;

  // =============================================================================
  // Constructors
  // =============================================================================

  /**
   * Create a new ZIP archive file.
   */
  constructor(options?: ArchiveFileOptionsZip);
  /**
   * Create a new TAR archive file.
   */
  constructor(options: ArchiveFileOptionsTar);
  /**
   * Create a new archive file with the specified format.
   */
  constructor(options?: ArchiveFileOptions);
  constructor(options: ArchiveFileOptions = {}) {
    checkOptionsObject("Archive.ArchiveFile", options);
    const format = (options.format ?? "zip") as F;
    if (format !== "zip" && format !== "tar") {
      throw new ArchiveArgumentError(
        "options.format",
        `Archive.ArchiveFile: options.format must be "zip" or "tar", got ${JSON.stringify(format)}`
      );
    }
    this._format = format;

    if (format === "zip") {
      // ZIP mode
      const zipOptions = options as ArchiveFileOptionsZip;
      checkZipOptions("Archive.ArchiveFile", zipOptions);
      this._zipOptions = zipOptions as F extends "zip" ? ZipFileOptions : null;
      this._zipPendingEntries = [] as unknown as F extends "zip" ? ZipPendingEntry[] : null;
      this._zipData = null as F extends "zip" ? Uint8Array | null : null;
      this._zipParser = null as F extends "zip" ? ZipParser | null : null;
      this._zipSourcePath = null as F extends "zip" ? string | null : null;
      this._zipPassword = zipOptions.password as F extends "zip"
        ? string | Uint8Array | undefined
        : null;
      this._zipAbortController = null as F extends "zip" ? AbortController | null : null;
      this._zipBytesWritten = 0 as F extends "zip" ? number : null;

      // TAR fields null
      this._tarSource = null as F extends "tar" ? Uint8Array | null : null;
      this._tarReader = null as F extends "tar" ? TarReader | null : null;
      this._tarPendingEntries = null as F extends "tar" ? TarPendingEntry[] : null;
      this._tarOptions = null as F extends "tar" ? ArchiveFileOptionsTar : null;
    } else {
      // TAR mode
      const tarOptions = options as ArchiveFileOptionsTar;
      this._tarSource = null as F extends "tar" ? Uint8Array | null : null;
      this._tarReader = null as F extends "tar" ? TarReader | null : null;
      this._tarPendingEntries = [] as unknown as F extends "tar" ? TarPendingEntry[] : null;
      this._tarOptions = tarOptions as F extends "tar" ? ArchiveFileOptionsTar : null;

      // ZIP fields null
      this._zipOptions = null as F extends "zip" ? ZipFileOptions : null;
      this._zipPendingEntries = null as F extends "zip" ? ZipPendingEntry[] : null;
      this._zipData = null as F extends "zip" ? Uint8Array | null : null;
      this._zipParser = null as F extends "zip" ? ZipParser | null : null;
      this._zipSourcePath = null as F extends "zip" ? string | null : null;
      this._zipPassword = null as F extends "zip" ? string | Uint8Array | undefined : null;
      this._zipAbortController = null as F extends "zip" ? AbortController | null : null;
      this._zipBytesWritten = null as F extends "zip" ? number : null;
    }
  }

  // ===========================================================================
  // ZIP Internal Methods
  // ===========================================================================

  /**
   * Get the ZIP creation options for createZip/createZipSync.
   */
  private _getCreateZipOptions() {
    const opts = this._zipOptions as ZipFileOptions;
    return {
      level: opts.level,
      timestamps: opts.timestamps,
      comment: opts.comment,
      zip64: opts.zip64,
      modTime: opts.modTime,
      reproducible: opts.reproducible,
      smartStore: opts.smartStore,
      encryptionMethod: opts.encryptionMethod,
      password: this._zipPassword as string | Uint8Array | undefined
    };
  }

  // ===========================================================================
  // ZIP State Accessors (type-safe helpers to reduce casts)
  // ===========================================================================

  /** Get ZIP options (throws if not ZIP format) */
  private get _zip_options(): ZipFileOptions {
    return this._zipOptions as ZipFileOptions;
  }

  /** Get ZIP pending entries array */
  private get _zip_pending(): ZipPendingEntry[] {
    return this._zipPendingEntries as ZipPendingEntry[];
  }

  /** Get ZIP parser (may be null) */
  private get _zip_parser(): ZipParser | null {
    return this._zipParser as ZipParser | null;
  }

  /** Get ZIP password */
  private get _zip_password(): string | Uint8Array | undefined {
    return this._zipPassword as string | Uint8Array | undefined;
  }

  /** Get TAR pending entries array */
  private get _tar_pending(): TarPendingEntry[] {
    return this._tarPendingEntries as TarPendingEntry[];
  }

  // ===========================================================================
  // The View: the archive a write would produce now
  // ===========================================================================

  /** The date a ZIP write records for an entry added now that nothing else dates. */
  private _zipStamp(): Date {
    const options = this._zip_options;
    return options.modTime ?? (options.reproducible ? REPRODUCIBLE_ZIP_MOD_TIME : new Date());
  }

  /** The date a TAR write records for an entry added now that nothing else dates. */
  private _tarStamp(): Date {
    return this._tar_options.modTime ?? new Date();
  }

  /** Record a ZIP edit; a cached view absorbs an edit by name, anything else drops it. */
  private _pushZipOp(op: ZipPendingEntry): void {
    this._zip_pending.push(op);
    this._viewGeneration++;
    if (this._zipView) {
      if (op.type === "directory" || op.type === "glob") {
        this._zipView = null;
      } else {
        applyZipOp(this._zipView, op);
      }
    }
  }

  /** Record a TAR edit. */
  private _pushTarOp(op: TarPendingEntry): void {
    this._tar_pending.push(op);
    this._viewGeneration++;
    this._tarView = null;
  }

  /** Replay the pending edits over the opened archive's entries. */
  private *_zipViewSteps(): ViewResolution<Map<string, ZipViewEntry>> {
    const pathOptions = resolveZipPathOptions(this._zip_options);
    const view = new Map<string, ZipViewEntry>();
    for (const record of this._zip_parser?.getEntries() ?? []) {
      const name = normalizeZipPath(record.path, pathOptions);
      view.set(name, { name, isDirectory: record.type === "directory", base: record });
    }
    for (const op of this._zip_pending) {
      if (op.type === "directory" || op.type === "glob") {
        const found = yield fsQueryOf(op);
        for (const entry of expandZipFound(op, found, this._zip_options, this._zip_password)) {
          view.set(entry.name, entry);
        }
      } else {
        applyZipOp(view, op);
      }
    }
    return view;
  }

  /** Replay the pending entries over the opened TAR's entries. */
  private *_tarViewSteps(): ViewResolution<Map<string, TarViewEntry>> {
    const view = new Map<string, TarViewEntry>();
    const source = this._tarSource as Uint8Array | null;
    if (source) {
      for (const [name, base] of untarSync(source, { maxEntrySize: Infinity })) {
        view.set(name, { name, isDirectory: isTarDirectory(base.info), base });
      }
    }
    for (const op of this._tar_pending) {
      if (op.type === "directory" || op.type === "glob") {
        const found = yield fsQueryOf(op, process.cwd());
        for (const entry of expandTarFound(op, found)) {
          view.set(entry.name, entry);
        }
      } else {
        view.set(op.tarPath, { name: op.tarPath, isDirectory: false, op });
      }
    }
    return view;
  }

  private _zipViewSync(): Map<string, ZipViewEntry> {
    return (this._zipView ??= resolveViewSync(this._zipViewSteps()));
  }

  private async _zipViewAsync(): Promise<Map<string, ZipViewEntry>> {
    if (this._zipView) {
      return this._zipView;
    }
    const generation = this._viewGeneration;
    const view = await resolveViewAsync(this._zipViewSteps());
    if (generation === this._viewGeneration) {
      this._zipView ??= view;
    }
    return view;
  }

  private _tarViewSync(): Map<string, TarViewEntry> {
    return (this._tarView ??= resolveViewSync(this._tarViewSteps()));
  }

  private async _tarViewAsync(): Promise<Map<string, TarViewEntry>> {
    if (this._tarView) {
      return this._tarView;
    }
    const generation = this._viewGeneration;
    const view = await resolveViewAsync(this._tarViewSteps());
    if (generation === this._viewGeneration) {
      this._tarView ??= view;
    }
    return view;
  }

  /**
   * A ZIP view's entries in the order a write stores them: as opened while nothing is pending
   * (a write then returns the archive unchanged), sorted by name once a write rebuilds it.
   */
  private _zipOrder(view: Map<string, ZipViewEntry>): ZipViewEntry[] {
    const list = Array.from(view.values());
    return this.hasPendingChanges() ? list.sort(byZipWriteOrder) : list;
  }

  /** The ZIP entry a write produces for a view entry, its data read now. */
  private async _zipEntryOf(v: ZipViewEntry): Promise<ZipEntry> {
    const options = this._zip_options;
    const password = this._zip_password;
    if (v.base) {
      const data = (await this._zip_parser!.extract(v.base.path, password)) ?? EMPTY_UINT8ARRAY;
      return buildPreservedEntry(v.name, v.base, data, options, password);
    }
    if (v.found) {
      return v.found.build(
        v.found.path === null ? EMPTY_UINT8ARRAY : await readSourceBytes(v.found.path)
      );
    }
    const op = v.op!;
    if (op.type === "file") {
      const data = await readSourceBytes(op.localPath);
      const stats = await safeStats(op.localPath);
      return buildZipEntry(
        op.zipPath,
        data,
        op.options,
        options,
        password,
        op.stamp,
        zipFsMetadata(stats)
      );
    }
    if (op.type === "stream") {
      op.buffered ??= await collectStream(op.stream);
    }
    return this._zipEntryOfInMemory(op);
  }

  /** {@link _zipEntryOf}, synchronously; a pending stream nothing has read is refused. */
  private _zipEntryOfSync(v: ZipViewEntry): ZipEntry {
    const options = this._zip_options;
    const password = this._zip_password;
    if (v.base) {
      const data = this._zip_parser!.extractSync(v.base.path, password) ?? EMPTY_UINT8ARRAY;
      return buildPreservedEntry(v.name, v.base, data, options, password);
    }
    if (v.found) {
      return v.found.build(
        v.found.path === null ? EMPTY_UINT8ARRAY : readSourceBytesSync(v.found.path)
      );
    }
    const op = v.op!;
    if (op.type === "file") {
      const data = readSourceBytesSync(op.localPath);
      const stats = safeStatsSync(op.localPath);
      return buildZipEntry(
        op.zipPath,
        data,
        op.options,
        options,
        password,
        op.stamp,
        zipFsMetadata(stats)
      );
    }
    if (op.type === "stream" && !op.buffered) {
      throw unreadStreamError(v.name);
    }
    return this._zipEntryOfInMemory(op);
  }

  private _zipEntryOfInMemory(
    op: ZipPendingBufferEntry | ZipPendingStreamEntry | ZipPendingSymlinkEntry
  ): ZipEntry {
    const options = this._zip_options;
    const password = this._zip_password;
    switch (op.type) {
      case "buffer":
        return buildZipEntry(op.zipPath, op.data, op.options, options, password, op.stamp);
      case "stream":
        return buildZipEntry(op.zipPath, op.buffered!, op.options, options, password, op.stamp);
      case "symlink":
        return buildSymlinkEntry(op.zipPath, op.target, op.mode, options.encoding, op.stamp);
    }
  }

  /** A view entry's data: an opened entry is decompressed with `password` if given. */
  private async _zipDataOf(v: ZipViewEntry, password?: string | Uint8Array): Promise<Uint8Array> {
    if (v.base) {
      return (
        (await this._zip_parser!.extract(v.base.path, password ?? this._zip_password)) ??
        EMPTY_UINT8ARRAY
      );
    }
    return (await this._zipEntryOf(v)).data;
  }

  private _zipDataOfSync(v: ZipViewEntry, password?: string | Uint8Array): Uint8Array {
    if (v.base) {
      return (
        this._zip_parser!.extractSync(v.base.path, password ?? this._zip_password) ??
        EMPTY_UINT8ARRAY
      );
    }
    return this._zipEntryOfSync(v).data;
  }

  /** The record a write stores for a view entry: an opened entry's own, or a pending one's. */
  private _zipRecordOf(v: ZipViewEntry, entry: ZipEntry | undefined): ZipEntryRecord {
    if (v.record) {
      return v.record;
    }
    if (v.base) {
      return (v.record = v.base.path === v.name ? v.base : { ...v.base, path: v.name });
    }
    return (v.record = recordOfPendingZipEntry(v.name, entry!, this._getCreateZipOptions()));
  }

  private async _zipRecord(v: ZipViewEntry): Promise<ZipEntryRecord> {
    return this._zipRecordOf(v, hasKnownRecord(v) ? undefined : await this._zipEntryOf(v));
  }

  private _zipRecordSync(v: ZipViewEntry): ZipEntryRecord {
    return this._zipRecordOf(v, hasKnownRecord(v) ? undefined : this._zipEntryOfSync(v));
  }

  /**
   * What extracting a view entry needs: its record and a reader for its data. For a pending
   * entry the record is computed without compressing, since extraction needs every field but
   * the compressed size; the data is read again when it is written out, not held meanwhile.
   */
  private _zipExtractItem(v: ZipViewEntry, entry: ZipEntry | undefined): ZipExtractItem {
    const record = hasKnownRecord(v)
      ? this._zipRecordOf(v, undefined)
      : recordOfPendingZipEntry(v.name, entry!, this._getCreateZipOptions(), true);
    return {
      record,
      read: password => this._zipDataOf(v, password),
      readSync: password => this._zipDataOfSync(v, password)
    };
  }

  private async _zipExtractItemOf(v: ZipViewEntry): Promise<ZipExtractItem> {
    return this._zipExtractItem(v, hasKnownRecord(v) ? undefined : await this._zipEntryOf(v));
  }

  private async _zipExtractItems(): Promise<ZipExtractItem[]> {
    if (this._zip_pending.length === 0) {
      return this._zipParserItems();
    }
    const items: ZipExtractItem[] = [];
    for (const v of this._zipOrder(await this._zipViewAsync())) {
      items.push(await this._zipExtractItemOf(v));
    }
    return items;
  }

  private _zipExtractItemsSync(): ZipExtractItem[] {
    if (this._zip_pending.length === 0) {
      return this._zipParserItems();
    }
    return this._zipOrder(this._zipViewSync()).map(v =>
      this._zipExtractItem(v, hasKnownRecord(v) ? undefined : this._zipEntryOfSync(v))
    );
  }

  private _zipParserItems(): ZipExtractItem[] {
    const parser = this._zip_parser;
    if (!parser) {
      return [];
    }
    return parser.getEntries().map(record => ({
      record,
      read: password => parser.extract(record.path, password ?? this._zip_password),
      readSync: password => parser.extractSync(record.path, password ?? this._zip_password)
    }));
  }

  /** The data and add options a TAR write uses for a view entry, read now. */
  private async _tarEntryOf(
    v: TarViewEntry
  ): Promise<{ data: Uint8Array; options: TarArchiveEntryOptions }> {
    if (v.op?.type === "file") {
      const data = await readSourceBytes(v.op.localPath);
      return { data, options: this._tarFileOptions(v.op, await safeStats(v.op.localPath)) };
    }
    if (v.found?.path) {
      return { data: await readSourceBytes(v.found.path), options: v.found.options };
    }
    if (v.op?.type === "stream") {
      v.op.buffered ??= await collectStream(v.op.stream);
    }
    return this._tarEntryOfSync(v);
  }

  private _tarEntryOfSync(v: TarViewEntry): { data: Uint8Array; options: TarArchiveEntryOptions } {
    if (v.base) {
      return { data: v.base.data, options: tarCopyOptions(v.base.info) };
    }
    if (v.found) {
      return {
        data: v.found.path === null ? EMPTY_UINT8ARRAY : readSourceBytesSync(v.found.path),
        options: v.found.options
      };
    }
    const op = v.op!;
    switch (op.type) {
      case "file":
        return {
          data: readSourceBytesSync(op.localPath),
          options: this._tarFileOptions(op, safeStatsSync(op.localPath))
        };
      case "buffer":
        return { data: op.data, options: buildTarAddOptions(op.options, { mtime: op.stamp }) };
      case "stream":
        if (!op.buffered) {
          throw unreadStreamError(v.name);
        }
        return { data: op.buffered, options: buildTarAddOptions(op.options, { mtime: op.stamp }) };
      case "symlink":
        return {
          data: EMPTY_UINT8ARRAY,
          options: { ...buildTarSymlinkOptions(op.target, op.mode), mtime: op.stamp }
        };
    }
  }

  private _tarFileOptions(
    op: Extract<TarNamedPendingEntry, { type: "file" }>,
    stats: { mode?: number; mtime?: Date } | undefined
  ): TarArchiveEntryOptions {
    return buildTarAddOptions(op.options, { mode: stats?.mode, mtime: stats?.mtime ?? op.stamp });
  }

  /** The header a TAR write stores for a view entry. */
  private _tarInfoOf(
    v: TarViewEntry,
    entry: () => { data: Uint8Array; options: TarArchiveEntryOptions }
  ): TarEntryInfo {
    if (v.info) {
      return v.info;
    }
    if (v.base) {
      return (v.info = v.base.info);
    }
    const { data, options } = entry();
    return (v.info = infoOfPendingTarEntry(v.name, data, options, this._tar_options.modTime));
  }

  /** A fresh TAR builder holding every entry of a view, for a buffered write. */
  private async _tarArchiveOf(view: Map<string, TarViewEntry>): Promise<TarArchive> {
    const archive = new TarArchive({ modTime: this._tar_options.modTime });
    await processInOrderWithConcurrency(
      view.values(),
      this._tar_options.concurrency ?? DEFAULT_IO_CONCURRENCY,
      async v => {
        const { data, options } = await this._tarEntryOf(v);
        return () => {
          archive.add(v.name, data, options);
        };
      }
    );
    return archive;
  }

  /** Make `tar` the archive's content: what reads see and what the next build starts from. */
  private _loadTar(tar: Uint8Array): void {
    this._tarSource = tar as F extends "tar" ? Uint8Array | null : null;
    this._setTarReader(new TarReader(tar));
  }

  /** Get TAR options */
  private get _tar_options(): ArchiveFileOptionsTar {
    return this._tarOptions as ArchiveFileOptionsTar;
  }

  /** Set a ZIP state field */
  /**
   * Set a ZIP state field.
   *
   * The ZIP state fields are conditionally typed on the archive format `F`
   * (`F extends "zip" ? T : null`). Within this generic class `F` is unresolved,
   * so a precise per-key value type cannot be expressed. Callers only invoke this
   * in the ZIP branch with the correct value type; the assignment is funneled
   * through a single localized cast.
   */
  /** See {@link ArchiveFileOpenOptionsZip.checkCrc32}. */
  private _checkCrc32 = true;

  private _setZipState<
    K extends
      | "_zipData"
      | "_zipParser"
      | "_zipPassword"
      | "_zipAbortController"
      | "_zipBytesWritten"
      | "_zipPendingEntries"
      | "_zipSourcePath"
  >(key: K, value: ArchiveZipState[K]): void {
    (this as unknown as ArchiveZipState)[key] = value;
  }

  /** Set the TAR reader for read mode */
  private _setTarReader(reader: TarReader): void {
    this._tarReader = reader as F extends "tar" ? TarReader | null : null;
  }

  /** Check per-entry ZIP options as `ZipArchive.add` does. */
  private _checkZipEntry(where: string, options: AddFileOptions): void {
    checkZipEntryOptions(`Archive.ArchiveFile#${where}`, options);
  }

  /** The view key of a ZIP entry name: normalised the way the archive normalises names. */
  private _zipKey(entryPath: string): string {
    if (!entryPath) {
      throw new ArchiveArgumentError("name", "Entry name is required");
    }
    return normalizeZipPath(entryPath, resolveZipPathOptions(this._zip_options));
  }

  /**
   * Initialize a ZIP archive from existing ZIP data.
   */
  private _initZipFromData(
    data: Uint8Array,
    password?: string | Uint8Array,
    sourcePath?: string,
    encoding?: ZipStringEncoding,
    checkCrc32 = true,
    limits: { maxEntrySize?: number; maxEntries?: number } = {}
  ): void {
    this._checkCrc32 = checkCrc32;
    this._setZipState("_zipData", data);
    this._setZipState(
      "_zipParser",
      new ZipParser(data, {
        password,
        encoding,
        checkCrc32,
        maxEntrySize: limits.maxEntrySize,
        maxEntries: limits.maxEntries
      })
    );
    this._setZipState("_zipPassword", password);
    if (sourcePath) {
      this._setZipState("_zipSourcePath", sourcePath);
    }
    // The opened archive's comment is what a write keeps unless the caller sets another.
    this._zip_options.comment ??= this._zip_parser!.getZipComment();
  }

  /**
   * Initialize a TAR archive from raw data.
   */
  private _initTarFromData(data: Uint8Array, gzip?: boolean): void {
    let tarData = data;
    if (gzip) {
      tarData = gunzipTarSync(data);
    }
    this._loadTar(tarData);
  }

  /**
   * Shared factory logic for creating an archive from data.
   */
  private static _fromData(
    data: Uint8Array,
    options: ArchiveFileOpenOptions,
    sourcePath?: string,
    autoDetectGzip?: boolean
  ): ArchiveFile<ArchiveFormat> {
    const format = options.format ?? "zip";

    if (format === "zip") {
      const archive = new ArchiveFile<"zip">({ format: "zip" });
      const zipOptions = options as ArchiveFileOpenOptionsZip;
      archive._initZipFromData(
        data,
        zipOptions.password,
        sourcePath,
        zipOptions.encoding,
        zipOptions.checkCrc32,
        { maxEntrySize: zipOptions.maxEntrySize, maxEntries: zipOptions.maxEntries }
      );
      return archive;
    } else {
      const tarOptions = options as ArchiveFileOpenOptionsTar;
      const shouldGunzip = tarOptions.gzip || autoDetectGzip;
      // Written back the way it was read: a `.tgz` opened by path stays gzip-compressed.
      const archive = new ArchiveFile<"tar">({
        ...(options as ArchiveFileOptionsTar),
        gzip: shouldGunzip
      });
      archive._initTarFromData(data, shouldGunzip);
      return archive;
    }
  }

  // =============================================================================
  // Static Factory Methods
  // =============================================================================

  /**
   * Create an ArchiveFile from a file on disk.
   */
  static fromFile(
    filePath: string,
    options?: ArchiveFileOpenOptionsZip
  ): Promise<ArchiveFile<"zip">>;
  static fromFile(
    filePath: string,
    options: ArchiveFileOpenOptionsTar
  ): Promise<ArchiveFile<"tar">>;
  static async fromFile(
    filePath: string,
    options: ArchiveFileOpenOptions = {}
  ): Promise<ArchiveFile<ArchiveFormat>> {
    checkOpenOptions("Archive.ArchiveFile.fromFile", options);
    let data: Uint8Array;
    try {
      data = await readFileBytes(filePath);
    } catch (error) {
      throw fileSystemError(filePath, "read", error);
    }
    const autoDetectGzip = filePath.endsWith(".gz") || filePath.endsWith(".tgz");
    return ArchiveFile._fromData(data, options, path.resolve(filePath), autoDetectGzip);
  }

  /**
   * Create an ArchiveFile from a file on disk (sync).
   */
  static fromFileSync(filePath: string, options?: ArchiveFileOpenOptionsZip): ArchiveFile<"zip">;
  static fromFileSync(filePath: string, options: ArchiveFileOpenOptionsTar): ArchiveFile<"tar">;
  static fromFileSync(
    filePath: string,
    options: ArchiveFileOpenOptions = {}
  ): ArchiveFile<ArchiveFormat> {
    checkOpenOptions("Archive.ArchiveFile.fromFileSync", options);
    let data: Uint8Array;
    try {
      data = readFileBytesSync(filePath);
    } catch (error) {
      throw fileSystemError(filePath, "read", error);
    }
    const autoDetectGzip = filePath.endsWith(".gz") || filePath.endsWith(".tgz");
    return ArchiveFile._fromData(data, options, path.resolve(filePath), autoDetectGzip);
  }

  /**
   * Create an ArchiveFile from a buffer.
   */
  static fromBuffer(data: ArchiveBytes, options?: ArchiveFileOpenOptionsZip): ArchiveFile<"zip">;
  static fromBuffer(data: ArchiveBytes, options: ArchiveFileOpenOptionsTar): ArchiveFile<"tar">;
  static fromBuffer(
    data: ArchiveBytes,
    options: ArchiveFileOpenOptions = {}
  ): ArchiveFile<ArchiveFormat> {
    if (!isArchiveBytes(data)) {
      throw new ArchiveArgumentError(
        "data",
        "Archive.ArchiveFile.fromBuffer: data must be bytes (an ArrayBufferView or ArrayBuffer)"
      );
    }
    checkOpenOptions("Archive.ArchiveFile.fromBuffer", options);
    return ArchiveFile._fromData(archiveBytesToUint8Array(data), options);
  }

  // =============================================================================
  // Properties
  // =============================================================================

  /**
   * Get the archive format.
   */
  get format(): F {
    return this._format;
  }

  /**
   * The number of entries the archive would have if written now: the opened (or last
   * written) archive's, with the pending edits applied. A pending directory or glob is
   * walked to count what it adds; no entry's data is read.
   */
  get entryCount(): number {
    if (this._format === "zip") {
      if (this._zip_pending.length === 0) {
        return this._zip_parser?.getEntries().length ?? 0;
      }
      return this._zipViewSync().size;
    }
    if (this._tar_pending.length === 0) {
      return this._tarSource ? this._tarEntryInfos().length : 0;
    }
    return this._tarViewSync().size;
  }

  /**
   * @deprecated Use {@link entryCount}, which now counts the entries of an opened TAR too.
   * Removed in 1.0.
   */
  async getEntryCountAsync(): Promise<number> {
    return this.entryCount;
  }

  /**
   * Get the source file path if the archive was loaded from disk.
   */
  get sourcePath(): string | undefined {
    if (this._format === "zip") {
      return (this._zipSourcePath as string | null) ?? undefined;
    }
    return undefined; // TAR doesn't track source path
  }

  /**
   * Check if the current operation has been aborted (ZIP only).
   */
  get aborted(): boolean {
    if (this._format === "zip") {
      return (this._zipAbortController as AbortController | null)?.signal.aborted ?? false;
    }
    return false;
  }

  // =============================================================================
  // Adding Files (Write Mode)
  // =============================================================================

  /**
   * Add a file from disk to the archive.
   */
  addFile(filePath: string, options?: F extends "tar" ? AddTarFileOptions : AddFileOptions): this {
    if (this._format === "zip") {
      const opts = (options ?? {}) as AddFileOptions;
      this._checkZipEntry("addFile", opts);
      const resolvedPath = path.resolve(filePath);
      const zipPath = zipWritePath(
        "addFile",
        this._zip_options,
        opts.prefix ?? "",
        opts.name ?? path.basename(filePath)
      );

      this._pushZipOp({
        type: "file",
        localPath: resolvedPath,
        zipPath,
        options: opts,
        stamp: this._zipStamp()
      });
    } else {
      const opts = (options ?? {}) as AddTarFileOptions;
      const tarPath = normalizeTarPath(opts.name ?? path.basename(filePath), opts.prefix);
      this._pushTarOp({
        type: "file",
        localPath: path.resolve(filePath),
        tarPath,
        options: opts,
        stamp: this._tarStamp()
      });
    }
    return this;
  }

  /**
   * Add a buffer to the archive.
   */
  addBuffer(
    data: Uint8Array,
    name: string,
    options?: F extends "tar" ? AddTarFileOptions : AddFileOptions
  ): this {
    checkBytes("Archive.ArchiveFile#addBuffer", "data", data);
    if (this._format === "zip") {
      const opts = (options ?? {}) as AddFileOptions;
      this._checkZipEntry("addBuffer", opts);
      this._pushZipOp({
        type: "buffer",
        data,
        zipPath: zipWritePath("addBuffer", this._zip_options, name),
        options: opts,
        stamp: this._zipStamp()
      });
    } else {
      const opts = (options ?? {}) as AddTarFileOptions;
      const tarPath = normalizeTarPath(name, opts.prefix);
      this._pushTarOp({
        type: "buffer",
        data,
        tarPath,
        options: opts,
        stamp: this._tarStamp()
      });
    }
    return this;
  }

  /**
   * Add text content to the archive.
   */
  addText(
    content: string,
    name: string,
    options?: F extends "tar" ? AddTarFileOptions : AddFileOptions
  ): this {
    checkString("Archive.ArchiveFile#addText", "content", content);
    return this.addBuffer(utf8Encoder.encode(content), name, options);
  }

  /**
   * Add a stream to the archive.
   */
  appendStream(
    stream: AsyncIterable<Uint8Array> | ReadableStream<Uint8Array>,
    name: string,
    options?: F extends "tar" ? AddTarFileOptions : AddFileOptions
  ): this {
    if (!isReadableStream(stream) && !isAsyncIterable(stream)) {
      throw new ArchiveArgumentError(
        "stream",
        "Archive.ArchiveFile#appendStream: stream must be a ReadableStream or an async iterable"
      );
    }
    if (this._format === "zip") {
      const opts = (options ?? {}) as AddFileOptions;
      this._checkZipEntry("appendStream", opts);
      this._pushZipOp({
        type: "stream",
        stream,
        zipPath: zipWritePath("appendStream", this._zip_options, name),
        options: opts,
        stamp: this._zipStamp()
      });
    } else {
      const opts = (options ?? {}) as AddTarFileOptions;
      const tarPath = normalizeTarPath(name, opts.prefix);
      this._pushTarOp({
        type: "stream",
        stream: stream as AsyncIterable<Uint8Array>,
        tarPath,
        options: opts,
        stamp: this._tarStamp()
      });
    }
    return this;
  }

  /**
   * Add a symbolic link to the archive (matches `ZipArchive.addSymlink` /
   * `TarArchive.addSymlink`). `options.mode` defaults to `0o120777` for ZIP and `0o777` for TAR.
   */
  addSymlink(name: string, target: string, options: { mode?: number } = {}): this {
    checkString("Archive.ArchiveFile#addSymlink", "target", target);
    return this._addSymlink(name, target, options.mode);
  }

  /**
   * @deprecated Use {@link addSymlink}, which takes the mode as `{ mode }`. Removed in 1.0.
   */
  symlink(name: string, target: string, mode?: number): this {
    return this._addSymlink(name, target, mode);
  }

  private _addSymlink(name: string, target: string, mode: number | undefined): this {
    if (this._format === "zip") {
      this._pushZipOp({
        type: "symlink",
        zipPath: zipWritePath("symlink", this._zip_options, name),
        target,
        mode,
        stamp: this._zipStamp()
      });
    } else {
      const tarPath = normalizeTarPath(name);
      this._pushTarOp({
        type: "symlink",
        tarPath,
        target,
        mode,
        stamp: this._tarStamp()
      });
    }
    return this;
  }

  /**
   * Add a directory recursively to the archive.
   */
  addDirectory(
    dirPath: string,
    options?: F extends "tar" ? AddTarDirectoryOptions : AddDirectoryOptions
  ): this {
    if (this._format === "zip") {
      this._pushZipOp({
        type: "directory",
        localPath: path.resolve(dirPath),
        options: (options ?? {}) as AddDirectoryOptions
      });
    } else {
      this._pushTarOp({
        type: "directory",
        localPath: path.resolve(dirPath),
        options: (options ?? {}) as AddTarDirectoryOptions
      });
    }
    return this;
  }

  /**
   * Add files matching a glob pattern to the archive.
   */
  addGlob(pattern: string, options?: F extends "tar" ? AddTarGlobOptions : AddGlobOptions): this {
    if (this._format === "zip") {
      this._pushZipOp({
        type: "glob",
        pattern,
        options: (options ?? {}) as AddGlobOptions
      });
    } else {
      this._pushTarOp({
        type: "glob",
        pattern,
        options: (options ?? {}) as AddTarGlobOptions
      });
    }
    return this;
  }

  // =============================================================================
  // ZIP Edit API
  // =============================================================================

  /**
   * Whether the archive a write would produce now has the entry. For a TAR, a pending stream
   * nothing has read yet cannot be looked into synchronously (see {@link readEntrySync}).
   */
  has(entryPath: string): boolean {
    checkEntryPath("Archive.ArchiveFile#has", entryPath, "entryPath");
    if (this._format === "zip") {
      return this._zipViewSync().has(this._zipKey(entryPath));
    }
    return this._tarEntryInfo(entryPath) !== undefined;
  }

  /**
   * Delete an entry (ZIP only): an entry of the opened archive, or a pending one, including
   * one a pending `addDirectory` / `addGlob` would add. `false` when there is no such entry.
   */
  delete(entryPath: string): F extends "zip" ? boolean : never {
    assertZipFormat(this._format, "delete()");
    checkEntryPath("Archive.ArchiveFile#delete", entryPath, "entryPath");
    const name = this._zipKey(entryPath);
    if (!this._zipViewSync().has(name)) {
      return false as F extends "zip" ? boolean : never;
    }
    this._pushZipOp({ type: "delete", name });
    return true as F extends "zip" ? boolean : never;
  }

  /**
   * Set/replace an entry (ZIP only).
   */
  set(
    entryPath: string,
    data: Uint8Array | string,
    options?: AddFileOptions
  ): F extends "zip" ? this : never {
    assertZipFormat(this._format, "set()");
    checkEntryPath("Archive.ArchiveFile#set", entryPath, "entryPath");
    if (typeof data !== "string") {
      checkBytes("Archive.ArchiveFile#set", "data", data);
    }

    const bytes = typeof data === "string" ? utf8Encoder.encode(data) : data;
    const opts = options ?? {};
    this._checkZipEntry("set", opts);

    // An existing entry is replaced under its name as it stands; a new one is checked as an add.
    const key = this._zipKey(entryPath);
    const zipPath = this._zipViewSync().has(key)
      ? key
      : zipWritePath("set", this._zip_options, entryPath);
    this._pushZipOp({
      type: "buffer",
      data: bytes,
      zipPath,
      options: opts,
      stamp: this._zipStamp()
    });
    return this as F extends "zip" ? this : never;
  }

  /**
   * Rename an entry (ZIP only); an entry already at `newPath` is replaced. `false` when there
   * is no entry at `oldPath`.
   */
  rename(oldPath: string, newPath: string): F extends "zip" ? boolean : never {
    assertZipFormat(this._format, "rename()");
    checkEntryPath("Archive.ArchiveFile#rename", oldPath, "oldPath");
    checkEntryPath("Archive.ArchiveFile#rename", newPath, "newPath");
    const to = zipWritePath("rename", this._zip_options, newPath);
    const from = this._zipKey(oldPath);
    const view = this._zipViewSync();
    if (!view.has(from)) {
      return false as F extends "zip" ? boolean : never;
    }
    if (from !== to) {
      this._pushZipOp({ type: "rename", from, to });
    }
    return true as F extends "zip" ? boolean : never;
  }

  /**
   * Set a password for encryption (ZIP only).
   */
  setPassword(password: string | Uint8Array | undefined): F extends "zip" ? this : never {
    assertZipFormat(this._format, "setPassword()");
    this._setZipState("_zipPassword", password);
    if (this._zip_parser) {
      this._zip_parser.setPassword(password);
    }
    // What a write encrypts, and so what it records, depends on the password.
    this._zipView = null;
    this._viewGeneration++;
    return this as F extends "zip" ? this : never;
  }

  // =============================================================================
  // Building (Write Mode)
  // =============================================================================

  /**
   * Build the archive and return as a buffer.
   */
  async toBuffer(): Promise<Uint8Array<ArrayBuffer>> {
    if (this._format === "zip") {
      return asPlainBytes(await this._buildZip());
    } else {
      return asPlainBytes(await this._buildTar());
    }
  }

  /**
   * Build the archive and return as a buffer (sync).
   */
  toBufferSync(): Uint8Array<ArrayBuffer> {
    if (this._format === "zip") {
      return asPlainBytes(this._buildZipSync());
    } else {
      return asPlainBytes(this._buildTarSync());
    }
  }

  /**
   * @deprecated Use {@link toBuffer}. Removed in 1.0.
   */
  async bytes(): Promise<Uint8Array<ArrayBuffer>> {
    return this.toBuffer();
  }

  /**
   * @deprecated Use {@link toBufferSync}. Removed in 1.0.
   */
  bytesSync(): Uint8Array<ArrayBuffer> {
    return this.toBufferSync();
  }

  /**
   * Write the archive to a file. An existing file is refused unless `options.overwrite`
   * says otherwise (default `"error"`), for TAR as for ZIP.
   */
  async writeFile(filePath: string, options?: WriteArchiveOptions): Promise<void> {
    const { overwrite = "error" } = (options ?? {}) as WriteZipOptions;
    const targetPath = path.resolve(filePath);

    const exists = await fileExists(targetPath);
    if (!checkOverwriteStrategy(exists, targetPath, overwrite)) {
      return;
    }

    const data = await this.toBuffer();
    try {
      await ensureDir(path.dirname(targetPath));
      await writeFileBytes(targetPath, data);
    } catch (error) {
      throw fileSystemError(targetPath, "write", error);
    }
  }

  /** @deprecated Use {@link writeFile}. Removed in 1.0. */
  async writeToFile(filePath: string, options?: WriteArchiveOptions): Promise<void> {
    return this.writeFile(filePath, options);
  }

  /**
   * Write the archive to a file (sync); see {@link writeFile}.
   */
  writeFileSync(filePath: string, options?: WriteArchiveOptions): void {
    const { overwrite = "error" } = (options ?? {}) as WriteZipOptions;
    const targetPath = path.resolve(filePath);

    const exists = fileExistsSync(targetPath);
    if (!checkOverwriteStrategy(exists, targetPath, overwrite)) {
      return;
    }

    const data = this.toBufferSync();
    try {
      ensureDirSync(path.dirname(targetPath));
      writeFileBytesSync(targetPath, data);
    } catch (error) {
      throw fileSystemError(targetPath, "write", error);
    }
  }
  /** @deprecated Use {@link writeFileSync}. Removed in 1.0. */
  writeToFileSync(filePath: string, options?: WriteArchiveOptions): void {
    this.writeFileSync(filePath, options);
  }

  // =============================================================================
  // Streaming (Write Mode)
  // =============================================================================

  /**
   * Generate archive as an async iterable stream.
   *
   * This is the most memory-efficient way to create archives, as it streams
   * data directly from sources to output without buffering the entire archive.
   *
   * @example
   * ```ts
   * const archive = new ArchiveFile();
   * archive.addFile("large-file.bin");
   * archive.addDirectory("./data");
   *
   * // Recommended: pipeTo() handles backpressure end-to-end.
   * await archive.pipeTo(createWriteStream("output.zip"));
   *
   * // Or stream manually — REMEMBER to honor backpressure on the sink:
   * import { once } from "node:events";
   * const writeStream = createWriteStream("output.zip");
   * for await (const chunk of archive.stream()) {
   *   if (!writeStream.write(chunk)) {
   *     await once(writeStream, "drain");
   *   }
   * }
   * writeStream.end();
   * ```
   */
  stream(options: ArchiveStreamOptions = {}): AsyncIterable<Uint8Array<ArrayBuffer>> {
    return this.operation(options).iterable;
  }

  /**
   * Get streaming operation with abort/progress control.
   *
   * @example
   * ```ts
   * const op = archive.operation({
   *   onProgress: (p) => console.log(`${p.entriesDone}/${p.entriesTotal}`),
   * });
   *
   * // Read chunks
   * for await (const chunk of op.iterable) {
   *   process(chunk);
   * }
   *
   * // Or abort if needed
   * op.abort();
   * ```
   */
  operation(options: ArchiveStreamOptions = {}): ArchiveStreamOperation {
    const op =
      this._format === "zip" ? this._buildZipStream(options) : this._buildTarStream(options);
    return {
      iterable: plainChunks(readingSources(op.iterable)),
      signal: op.signal,
      abort: reason => op.abort(reason),
      pointer: () => op.pointer(),
      progress: () => op.progress()
    };
  }

  /**
   * Pipe archive stream to a sink (WritableStream or Node.js Writable).
   *
   * @example
   * ```ts
   * // Node.js Writable
   * await archive.pipeTo(createWriteStream("output.zip"));
   *
   * // Web WritableStream
   * await archive.pipeTo(writableStream);
   * ```
   */
  async pipeTo(sink: ArchiveSink, options: ArchiveStreamOptions = {}): Promise<void> {
    checkSinkArgument("Archive.ArchiveFile#pipeTo", sink);
    await pipeIterableToSink(this.stream(options), sink);
  }

  /**
   * Stream archive directly to a file.
   *
   * This is the most efficient way to write large archives to disk,
   * as it avoids buffering the entire archive in memory.
   *
   * @example
   * ```ts
   * const archive = new ArchiveFile();
   * archive.addDirectory("./huge-folder");
   * await archive.streamToFile("output.zip", {
   *   onProgress: (p) => console.log(`${p.bytesOut} bytes written`),
   * });
   * ```
   */
  async streamToFile(
    filePath: string,
    options: ArchiveStreamOptions & WriteArchiveOptions = {}
  ): Promise<void> {
    const targetPath = path.resolve(filePath);
    const { overwrite = "error" } = options;

    const exists = await fileExists(targetPath);
    if (!checkOverwriteStrategy(exists, targetPath, overwrite)) {
      return;
    }

    let writeStream: ReturnType<typeof createWriteStream>;
    try {
      await ensureDir(path.dirname(targetPath));
      writeStream = createWriteStream(targetPath);
    } catch (error) {
      throw toFileError(targetPath, "write", error);
    }
    try {
      await this.pipeTo(writeStream, options);
    } catch (err) {
      // Clean up partial file on error
      writeStream.destroy();
      throw toFileError(targetPath, "write", err);
    }
  }

  // =============================================================================
  // Reading (Read Mode)
  // =============================================================================

  /**
   * Every entry of the archive a write would produce now, described as the write records it.
   *
   * A pending entry's description is computed the way the write computes it, so for ZIP its
   * data is read and compressed (once; the result is kept until the next edit). A pending
   * stream is buffered in memory so the write still has it.
   */
  async getEntries(): Promise<ArchiveEntryInfo[]> {
    if (this._format === "zip") {
      if (this._zip_pending.length === 0) {
        return (this._zip_parser?.getEntries() ?? []).map(mapZipEntryToInfo);
      }
      const entries: ArchiveEntryInfo[] = [];
      for (const v of this._zipOrder(await this._zipViewAsync())) {
        entries.push(mapZipEntryToInfo(await this._zipRecord(v)));
      }
      return entries;
    }
    if (this._tar_pending.length === 0) {
      return this._tarSource ? this._tarEntryInfos() : [];
    }
    const entries: ArchiveEntryInfo[] = [];
    for (const v of (await this._tarViewAsync()).values()) {
      const entry = hasKnownTarInfo(v) ? undefined : await this._tarEntryOf(v);
      entries.push(
        toTarEntryInfo(
          this._tarInfoOf(v, () => entry!),
          v.name
        )
      );
    }
    return entries;
  }

  /** The entries of the TAR this archive was opened from or last written as, in order. */
  private _tarEntryInfos(): ArchiveEntryInfo[] {
    return parseTar(this._tarSource as Uint8Array, { maxEntrySize: Infinity }).map(({ info }) =>
      toTarEntryInfo(info)
    );
  }

  /**
   * {@link getEntries}, synchronously. A pending stream nothing has read yet cannot be read
   * synchronously, so describing one throws until an async read has buffered it.
   */
  getEntriesSync(): ArchiveEntryInfo[] {
    if (this._format === "zip") {
      if (this._zip_pending.length === 0) {
        return (this._zip_parser?.getEntries() ?? []).map(mapZipEntryToInfo);
      }
      return this._zipOrder(this._zipViewSync()).map(v =>
        mapZipEntryToInfo(this._zipRecordSync(v))
      );
    }
    if (this._tar_pending.length === 0) {
      return this._tarSource ? this._tarEntryInfos() : [];
    }
    return Array.from(this._tarViewSync().values(), v =>
      toTarEntryInfo(
        this._tarInfoOf(v, () => this._tarEntryOfSync(v)),
        v.name
      )
    );
  }

  /**
   * The names of the archive a write would produce now, in the order it stores them. No
   * entry's data is read; a pending directory or glob is walked.
   */
  getEntryNames(): string[] {
    if (this._format === "zip") {
      if (this._zip_pending.length === 0) {
        return this._zip_parser?.listFiles() ?? [];
      }
      return this._zipOrder(this._zipViewSync()).map(v => v.name);
    }
    if (this._tar_pending.length === 0) {
      return this._tarSource ? this._tarEntryInfos().map(entry => entry.path) : [];
    }
    return Array.from(this._tarViewSync().keys());
  }

  /**
   * One entry's info as a write records it; `undefined` when the archive a write would produce has
   * no such entry. See {@link getEntries}. Of two TAR members with one name the later is
   * returned, as {@link readEntry} reads it.
   */
  getEntry(entryPath: string): (F extends "tar" ? ArchiveEntryInfo : ZipEntryInfo) | undefined {
    checkEntryPath("Archive.ArchiveFile#getEntry", entryPath, "entryPath");
    if (this._format === "tar") {
      return this._tarEntryInfo(entryPath) as
        | (F extends "tar" ? ArchiveEntryInfo : ZipEntryInfo)
        | undefined;
    }
    const v = this._zipViewSync().get(this._zipKey(entryPath));
    return (v ? mapZipEntryToInfo(this._zipRecordSync(v)) : undefined) as
      | (F extends "tar" ? ArchiveEntryInfo : ZipEntryInfo)
      | undefined;
  }

  /** A TAR entry's info, the last of that name, or `undefined`. */
  private _tarEntryInfo(entryPath: string): ArchiveEntryInfo | undefined {
    const normalized = normalizeTarPath(entryPath);
    return this.getEntriesSync().findLast(e => e.path === entryPath || e.path === normalized);
  }

  /** The TAR view entry for a name, or `undefined`. */
  private _tarViewEntry(
    view: Map<string, TarViewEntry>,
    entryPath: string
  ): TarViewEntry | undefined {
    return view.get(entryPath) ?? view.get(normalizeTarPath(entryPath));
  }

  /**
   * Read an entry as bytes; `undefined` when the archive a write would produce has no such entry.
   * A pending entry's data is read from where it is pending: memory, the file on disk, or —
   * buffered in memory the first time — the stream.
   *
   * @param options.password - Overrides the password the archive was opened with (ZIP)
   */
  readEntry(
    entryPath: string,
    options?: ArchiveFileReadOptions
  ): Promise<Uint8Array<ArrayBuffer> | undefined>;
  /** @deprecated Pass the password as `{ password }`. Removed in 1.0. */
  readEntry(
    entryPath: string,
    password: string | Uint8Array
  ): Promise<Uint8Array<ArrayBuffer> | undefined>;
  async readEntry(
    entryPath: string,
    options?: ArchiveFileReadOptions | string | Uint8Array
  ): Promise<Uint8Array<ArrayBuffer> | undefined> {
    checkEntryPath("Archive.ArchiveFile#readEntry", entryPath, "entryPath");
    if (this._format === "zip") {
      const v = (await this._zipViewAsync()).get(this._zipKey(entryPath));
      return v
        ? asPlainBytes(await this._zipDataOf(v, toReadOptions(options).password))
        : undefined;
    }
    if (this._tar_pending.length === 0) {
      return this._tarReader ? (this._tarReader as TarReader).bytes(entryPath) : undefined;
    }
    const v = this._tarViewEntry(await this._tarViewAsync(), entryPath);
    return v ? asPlainBytes((await this._tarEntryOf(v)).data) : undefined;
  }

  /**
   * {@link readEntry}, synchronously. A pending stream nothing has read yet cannot be read
   * synchronously and throws until an async read has buffered it.
   *
   * @param options.password - Overrides the password the archive was opened with (ZIP)
   */
  readEntrySync(
    entryPath: string,
    options?: ArchiveFileReadOptions
  ): Uint8Array<ArrayBuffer> | undefined;
  /** @deprecated Pass the password as `{ password }`. Removed in 1.0. */
  readEntrySync(
    entryPath: string,
    password: string | Uint8Array
  ): Uint8Array<ArrayBuffer> | undefined;
  readEntrySync(
    entryPath: string,
    options?: ArchiveFileReadOptions | string | Uint8Array
  ): Uint8Array<ArrayBuffer> | undefined {
    checkEntryPath("Archive.ArchiveFile#readEntrySync", entryPath, "entryPath");
    if (this._format === "zip") {
      const v = this._zipViewSync().get(this._zipKey(entryPath));
      return v ? asPlainBytes(this._zipDataOfSync(v, toReadOptions(options).password)) : undefined;
    }
    const v = this._tarViewEntry(this._tarViewSync(), entryPath);
    return v ? asPlainBytes(this._tarEntryOfSync(v).data) : undefined;
  }

  /**
   * Read an entry as text; `undefined` when the archive a write would produce has no such entry.
   *
   * @param options.encoding - Text encoding (default `"utf-8"`)
   * @param options.password - Overrides the password the archive was opened with (ZIP)
   */
  readAsText(entryPath: string, options?: ArchiveFileReadTextOptions): Promise<string | undefined>;
  /** @deprecated Pass the encoding as `{ encoding }`. Removed in 1.0. */
  readAsText(entryPath: string, encoding: string): Promise<string | undefined>;
  async readAsText(
    entryPath: string,
    options?: ArchiveFileReadTextOptions | string
  ): Promise<string | undefined> {
    const { encoding = "utf-8", password } = toReadTextOptions(options);
    const data = await this.readEntry(entryPath, password === undefined ? undefined : { password });
    if (!data) {
      return undefined;
    }
    return decodeEntryText(data, encoding, "whatwg");
  }

  /**
   * {@link readAsText}, synchronously; see {@link readEntrySync}.
   *
   * @param options.encoding - Text encoding (default `"utf-8"`)
   * @param options.password - Overrides the password the archive was opened with (ZIP)
   */
  readAsTextSync(entryPath: string, options?: ArchiveFileReadTextOptions): string | undefined;
  /** @deprecated Pass the encoding as `{ encoding }`. Removed in 1.0. */
  readAsTextSync(entryPath: string, encoding: string): string | undefined;
  readAsTextSync(
    entryPath: string,
    options?: ArchiveFileReadTextOptions | string
  ): string | undefined {
    const { encoding = "utf-8", password } = toReadTextOptions(options);
    const data = this.readEntrySync(entryPath, password === undefined ? undefined : { password });
    if (!data) {
      return undefined;
    }
    return decodeEntryText(data, encoding, "whatwg");
  }

  /**
   * Extract the archive a write would produce now to a directory.
   */
  async extractTo(targetDir: string, options: ExtractToOptions = {}): Promise<void> {
    if (this._format === "zip") {
      await this._extractZip(targetDir, options);
    } else {
      await this._extractTar(targetDir, options);
    }
  }

  /**
   * {@link extractTo}, synchronously; a pending stream nothing has read yet is refused.
   */
  extractToSync(targetDir: string, options: ExtractToOptions = {}): void {
    if (this._format === "zip") {
      this._extractZipSync(targetDir, options);
    } else {
      this._extractTarSync(targetDir, options);
    }
  }

  /**
   * Extract a single entry of the archive a write would produce to a file (ZIP only);
   * `false` when there is no such entry or `overwrite` skipped it.
   */
  async extractEntryTo(
    entryPath: string,
    targetPath: string,
    options: ExtractToOptions = {}
  ): Promise<boolean> {
    assertZipFormat(this._format, "extractEntryTo()");
    checkEntryPath("Archive.ArchiveFile#extractEntryTo", entryPath, "entryPath");

    const v = (await this._zipViewAsync()).get(this._zipKey(entryPath));
    if (!v) {
      return false;
    }
    const entry = (await this._zipExtractItemOf(v)).record;

    const { overwrite = "error", preserveTimestamps = true, password } = options;
    const resolvedTarget = path.resolve(targetPath);

    try {
      if (entry.type === "directory") {
        await ensureDir(resolvedTarget);
        return true;
      }

      if (!(await shouldExtract(resolvedTarget, entry.lastModified, overwrite))) {
        return false;
      }

      await ensureDir(path.dirname(resolvedTarget));

      const data = await this._zipDataOf(v, password);
      await writeFileBytes(resolvedTarget, data);
      if (preserveTimestamps) {
        await setFileTime(resolvedTarget, entry.lastModified);
      }

      return true;
    } catch (error) {
      throw toFileError(resolvedTarget, "write", error, "extracted");
    }
  }

  /**
   * {@link extractEntryTo}, synchronously (ZIP only).
   */
  extractEntryToSync(
    entryPath: string,
    targetPath: string,
    options: ExtractToOptions = {}
  ): boolean {
    assertZipFormat(this._format, "extractEntryToSync()");
    checkEntryPath("Archive.ArchiveFile#extractEntryToSync", entryPath, "entryPath");

    const v = this._zipViewSync().get(this._zipKey(entryPath));
    if (!v) {
      return false;
    }
    const entry = this._zipExtractItem(
      v,
      hasKnownRecord(v) ? undefined : this._zipEntryOfSync(v)
    ).record;

    const { overwrite = "error", preserveTimestamps = true, password } = options;
    const resolvedTarget = path.resolve(targetPath);

    try {
      if (entry.type === "directory") {
        ensureDirSync(resolvedTarget);
        return true;
      }

      if (!shouldExtractSync(resolvedTarget, entry.lastModified, overwrite)) {
        return false;
      }

      ensureDirSync(path.dirname(resolvedTarget));

      const data = this._zipDataOfSync(v, password);
      writeFileBytesSync(resolvedTarget, data);
      if (preserveTimestamps) {
        setFileTimeSync(resolvedTarget, entry.lastModified);
      }

      return true;
    } catch (error) {
      throw toFileError(resolvedTarget, "write", error, "extracted");
    }
  }

  // =============================================================================
  // ZIP Utility Methods
  // =============================================================================

  /**
   * Whether the archive a write would produce now has an encrypted entry (ZIP only): one
   * stored encrypted, or a pending one the archive's or the entry's password encrypts.
   */
  hasEncryptedEntries(): boolean {
    if (this._format !== "zip") {
      return false;
    }
    if (this._zip_pending.length === 0) {
      return this._zip_parser?.hasEncryptedEntries() ?? false;
    }
    for (const v of this._zipViewSync().values()) {
      if (v.base ? v.base.isEncrypted : this._wouldEncrypt(v)) {
        return true;
      }
    }
    return false;
  }

  /** Whether a write encrypts a pending view entry; decided by its options, not its data. */
  private _wouldEncrypt(v: ZipViewEntry): boolean {
    if (v.found) {
      return isEncryptedZipEntry(v.found.build(EMPTY_UINT8ARRAY));
    }
    const op = v.op!;
    if (op.type === "symlink") {
      return false;
    }
    return isEncryptedZipEntry(
      buildZipEntry(
        op.zipPath,
        EMPTY_UINT8ARRAY,
        op.options,
        this._zip_options,
        this._zip_password,
        op.stamp
      )
    );
  }

  /**
   * The archive comment a write records (ZIP only): the opened archive's, unless set.
   */
  getZipComment(): string {
    assertZipFormat(this._format, "getZipComment()");
    return this._zip_options.comment ?? "";
  }

  /**
   * Set or update the archive comment (ZIP only).
   */
  setZipComment(comment: string): F extends "zip" ? this : never {
    assertZipFormat(this._format, "setZipComment()");
    checkString("Archive.ArchiveFile#setZipComment", "comment", comment);
    this._zip_options.comment = comment;
    this._zipCommentChanged = true;
    return this as F extends "zip" ? this : never;
  }

  /**
   * @deprecated Use {@link setZipComment}: this sets the comment, it does not add one.
   * Removed in 1.0.
   */
  addZipComment(comment: string): this {
    assertZipFormat(this._format, "addZipComment()");
    checkString("Archive.ArchiveFile#addZipComment", "comment", comment);
    this.setZipComment(comment);
    return this;
  }

  /**
   * Get the comment for a specific entry (ZIP only).
   */
  getZipEntryComment(entryPath: string): string | undefined {
    assertZipFormat(this._format, "getZipEntryComment()");
    checkEntryPath("Archive.ArchiveFile#getZipEntryComment", entryPath, "entryPath");
    const entry = this.getEntry(entryPath) as ZipEntryInfo | undefined;
    return entry?.comment;
  }

  /**
   * Whether anything is pending: an add, a replacement, a delete, a rename or, for ZIP, a
   * new archive comment.
   */
  hasPendingChanges(): boolean {
    if (this._format === "zip") {
      return this._zip_pending.length > 0 || this._zipCommentChanged;
    }
    return this._tar_pending.length > 0;
  }

  /**
   * Abort the current operation (ZIP only).
   */
  abort(): this {
    if (this._format === "zip" && this._zipAbortController) {
      (this._zipAbortController as AbortController).abort();
      this._setZipState("_zipAbortController", null);
    }
    return this;
  }

  /**
   * Get the number of bytes written so far (ZIP only).
   */
  pointer(): number {
    assertZipFormat(this._format, "pointer()");
    if (this._zipData) {
      return (this._zipData as Uint8Array).length;
    }
    return this._zipBytesWritten as number;
  }

  /**
   * Get the AbortSignal for the current operation (ZIP only).
   */
  getAbortSignal(): AbortSignal | undefined {
    if (this._format === "zip") {
      return (this._zipAbortController as AbortController | null)?.signal;
    }
    return undefined;
  }

  // =============================================================================
  // Private Methods - ZIP Building
  // =============================================================================

  /**
   * Commit a built ZIP: it becomes the archive reads start from and the next build starts
   * from, and the pending edits are now part of it. A new archive is committed the same way
   * as an opened one, so adding after a build adds to what was built.
   */
  private _commitZipBuild(zipData: Uint8Array): void {
    this._setZipState("_zipData", zipData);
    this._setZipState("_zipPendingEntries", []);
    this._setZipState(
      "_zipParser",
      new ZipParser(zipData, { password: this._zip_password, checkCrc32: this._checkCrc32 })
    );
    this._zipView = null;
    this._zipCommentChanged = false;
    this._viewGeneration++;
  }

  private async _buildZip(): Promise<Uint8Array> {
    // Create abort controller for this operation
    this._setZipState("_zipAbortController", new AbortController());
    const signal = (this._zipAbortController as AbortController).signal;

    // Reset bytes counter
    this._setZipState("_zipBytesWritten", 0);

    // Check if we can return cached data
    if (this._zipData && !this.hasPendingChanges()) {
      this._setZipState("_zipAbortController", null);
      return this._zipData as Uint8Array;
    }

    const checkAbort = () => {
      if (signal.aborted) {
        throw createAbortError();
      }
    };

    // A write resolves the view afresh: a pending directory is walked as it is now.
    const view = await resolveViewAsync(this._zipViewSteps());
    const entries: ZipEntry[] = [];
    await processInOrderWithConcurrency(
      this._zipOrder(view),
      this._zip_options.concurrency ?? DEFAULT_IO_CONCURRENCY,
      async v => {
        checkAbort();
        const entry = await this._zipEntryOf(v);
        return () => {
          entries.push(entry);
        };
      }
    );
    checkAbort();

    const zipData = await createZip(entries, this._getCreateZipOptions());
    this._setZipState("_zipBytesWritten", zipData.length);
    this._commitZipBuild(zipData);
    this._setZipState("_zipAbortController", null);
    return zipData;
  }

  private _buildZipSync(): Uint8Array {
    if (this._zipData && !this.hasPendingChanges()) {
      return this._zipData as Uint8Array;
    }
    const view = resolveViewSync(this._zipViewSteps());
    const entries = this._zipOrder(view).map(v => this._zipEntryOfSync(v));
    const zipData = createZipSync(entries, this._getCreateZipOptions());
    this._commitZipBuild(zipData);
    return zipData;
  }

  // =============================================================================
  // Private Methods - ZIP Extraction
  // =============================================================================

  private async _extractZip(targetDir: string, options: ExtractToOptions): Promise<void> {
    const {
      overwrite = "error",
      filter,
      preserveTimestamps = true,
      preservePermissions = supportsUnixPermissions(),
      createSymlinks = true,
      password,
      signal,
      onProgress,
      onWarning
    } = options;

    const resolvedTarget = path.resolve(targetDir);
    const items = await this._zipExtractItems();
    const totalEntries = items.length;
    let extractedEntries = 0;
    let bytesWritten = 0;
    const effectivePassword = password ?? this._zip_password;
    const textDecoder = new TextDecoder();

    // Deferred symlinks - process after all files/dirs to ensure targets exist
    const deferredSymlinks: Array<{
      entry: ZipEntryRecord;
      targetPath: string;
      linkTarget: string;
    }> = [];

    for (let i = 0; i < items.length; i++) {
      const { record: entry, read } = items[i]!;
      // Check abort signal
      if (signal?.aborted) {
        throw createAbortError();
      }

      // Apply filter (pass isDirectory for both dirs and symlinks pointing to dirs)
      if (filter && !filter(entry.path, entry.type === "directory")) {
        continue;
      }

      const targetPath = path.join(resolvedTarget, entry.path);
      assertNoPathTraversal(targetPath, resolvedTarget, entry.path);

      if (entry.type === "directory") {
        // --- Directory ---
        if (
          !(await tryFsOpWithWarning(
            () => ensureDir(targetPath),
            onWarning,
            entry.path,
            targetPath
          ))
        ) {
          continue;
        }

        // Set directory permissions
        if (preservePermissions) {
          const dirMode = getEffectiveMode(entry.mode, "directory");
          await tryFsOpWithWarning(
            () => chmod(targetPath, dirMode & 0o7777), // Strip file type bits
            onWarning,
            entry.path,
            targetPath
          );
        }
      } else if (entry.type === "symlink") {
        // --- Symlink ---
        if (!createSymlinks) {
          // Skip symlinks if disabled
          continue;
        }

        // Extract symlink target (content of the entry is the link target path)
        const data = await read(effectivePassword);
        if (!data) {
          continue;
        }

        const linkTarget = textDecoder.decode(data);

        // Validate symlink target doesn't escape the extraction directory
        const resolvedLinkTarget = path.resolve(path.dirname(targetPath), linkTarget);
        if (
          !resolvedLinkTarget.startsWith(resolvedTarget + path.sep) &&
          resolvedLinkTarget !== resolvedTarget
        ) {
          // Symlink points outside extraction directory - emit warning and skip
          emitExtractWarning(
            onWarning,
            entry.path,
            targetPath,
            new Error(`Symlink target "${linkTarget}" points outside extraction directory`)
          );
          continue;
        }

        // Defer symlink creation to ensure target exists
        deferredSymlinks.push({ entry, targetPath, linkTarget });
      } else {
        // --- Regular File ---
        // Check overwrite strategy
        let shouldWrite: boolean;
        try {
          shouldWrite = await shouldExtract(targetPath, entry.lastModified, overwrite);
        } catch (err) {
          if (isIgnorableFsError(err)) {
            emitExtractWarning(onWarning, entry.path, targetPath, err);
            continue;
          }
          throw toFileError(targetPath, "write", err, "extracted");
        }

        if (!shouldWrite) {
          continue;
        }

        // Ensure parent directory exists
        if (
          !(await tryFsOpWithWarning(
            () => ensureDir(path.dirname(targetPath)),
            onWarning,
            entry.path,
            targetPath
          ))
        ) {
          continue;
        }

        // Extract file content
        const data = await read(effectivePassword);
        if (data) {
          let writeSuccess: boolean;
          try {
            await writeFileBytes(targetPath, data);
            bytesWritten += data.length;
            writeSuccess = true;
          } catch (err) {
            if (isIgnorableFsError(err)) {
              emitExtractWarning(onWarning, entry.path, targetPath, err);
              continue;
            }
            throw toFileError(targetPath, "write", err, "extracted");
          }

          if (writeSuccess) {
            // Set file permissions
            if (preservePermissions) {
              const fileMode = getEffectiveMode(entry.mode, "file");
              await tryFsOpWithWarning(
                () => chmod(targetPath, fileMode & 0o7777),
                onWarning,
                entry.path,
                targetPath
              );
            }

            // Set timestamps
            if (preserveTimestamps) {
              await tryFsOpWithWarning(
                () => setFileTime(targetPath, entry.lastModified),
                onWarning,
                entry.path,
                targetPath
              );
            }
          }
        }
      }

      extractedEntries++;

      // Report progress
      if (onProgress) {
        onProgress({
          currentEntry: entry.path,
          totalEntries,
          extractedEntries,
          bytesWritten
        });
      }
    }

    // Process deferred symlinks
    for (let i = 0; i < deferredSymlinks.length; i++) {
      const deferred = deferredSymlinks[i]!;
      const { entry, targetPath, linkTarget } = deferred;
      if (signal?.aborted) {
        throw createAbortError();
      }

      // Check overwrite strategy for symlink
      let shouldWrite: boolean;
      try {
        shouldWrite = await shouldExtract(targetPath, entry.lastModified, overwrite);
      } catch (err) {
        if (isIgnorableFsError(err)) {
          emitExtractWarning(onWarning, entry.path, targetPath, err);
          continue;
        }
        throw toFileError(targetPath, "write", err, "extracted");
      }

      if (!shouldWrite) {
        continue;
      }

      // Ensure parent directory exists
      if (
        !(await tryFsOpWithWarning(
          () => ensureDir(path.dirname(targetPath)),
          onWarning,
          entry.path,
          targetPath
        ))
      ) {
        continue;
      }

      // Create symlink
      const symlinkCreated = await tryFsOpWithWarning(
        () => createSymlink(linkTarget, targetPath),
        onWarning,
        entry.path,
        targetPath
      );

      if (symlinkCreated) {
        extractedEntries++;

        if (onProgress) {
          onProgress({
            currentEntry: entry.path,
            totalEntries,
            extractedEntries,
            bytesWritten
          });
        }
      }
    }
  }

  private _extractZipSync(targetDir: string, options: ExtractToOptions): void {
    const {
      overwrite = "error",
      filter,
      preserveTimestamps = true,
      preservePermissions = supportsUnixPermissions(),
      createSymlinks = true,
      password,
      signal,
      onProgress,
      onWarning
    } = options;

    const resolvedTarget = path.resolve(targetDir);
    const items = this._zipExtractItemsSync();
    const totalEntries = items.length;
    let extractedEntries = 0;
    let bytesWritten = 0;
    const effectivePassword = password ?? this._zip_password;
    const textDecoder = new TextDecoder();

    // Deferred symlinks - process after all files/dirs to ensure targets exist
    const deferredSymlinks: Array<{
      entry: ZipEntryRecord;
      targetPath: string;
      linkTarget: string;
    }> = [];

    for (let i = 0; i < items.length; i++) {
      const { record: entry, readSync } = items[i]!;
      // Check abort signal
      if (signal?.aborted) {
        throw createAbortError();
      }

      // Apply filter
      if (filter && !filter(entry.path, entry.type === "directory")) {
        continue;
      }

      const targetPath = path.join(resolvedTarget, entry.path);
      assertNoPathTraversal(targetPath, resolvedTarget, entry.path);

      if (entry.type === "directory") {
        // --- Directory ---
        if (
          !tryFsOpWithWarningSync(
            () => ensureDirSync(targetPath),
            onWarning,
            entry.path,
            targetPath
          )
        ) {
          continue;
        }

        // Set directory permissions
        if (preservePermissions) {
          const dirMode = getEffectiveMode(entry.mode, "directory");
          tryFsOpWithWarningSync(
            () => chmodSync(targetPath, dirMode & 0o7777),
            onWarning,
            entry.path,
            targetPath
          );
        }
      } else if (entry.type === "symlink") {
        // --- Symlink ---
        if (!createSymlinks) {
          continue;
        }

        // Extract symlink target
        const data = readSync(effectivePassword);
        if (!data) {
          continue;
        }

        const linkTarget = textDecoder.decode(data);

        // Validate symlink target doesn't escape the extraction directory
        const resolvedLinkTarget = path.resolve(path.dirname(targetPath), linkTarget);
        if (
          !resolvedLinkTarget.startsWith(resolvedTarget + path.sep) &&
          resolvedLinkTarget !== resolvedTarget
        ) {
          emitExtractWarning(
            onWarning,
            entry.path,
            targetPath,
            new Error(`Symlink target "${linkTarget}" points outside extraction directory`)
          );
          continue;
        }

        // Defer symlink creation
        deferredSymlinks.push({ entry, targetPath, linkTarget });
      } else {
        // --- Regular File ---
        let shouldWrite: boolean;
        try {
          shouldWrite = shouldExtractSync(targetPath, entry.lastModified, overwrite);
        } catch (err) {
          if (isIgnorableFsError(err)) {
            emitExtractWarning(onWarning, entry.path, targetPath, err);
            continue;
          }
          throw toFileError(targetPath, "write", err, "extracted");
        }

        if (!shouldWrite) {
          continue;
        }

        if (
          !tryFsOpWithWarningSync(
            () => ensureDirSync(path.dirname(targetPath)),
            onWarning,
            entry.path,
            targetPath
          )
        ) {
          continue;
        }

        const data = readSync(effectivePassword);
        if (data) {
          let writeSuccess: boolean;
          try {
            writeFileBytesSync(targetPath, data);
            bytesWritten += data.length;
            writeSuccess = true;
          } catch (err) {
            if (isIgnorableFsError(err)) {
              emitExtractWarning(onWarning, entry.path, targetPath, err);
              continue;
            }
            throw toFileError(targetPath, "write", err, "extracted");
          }

          if (writeSuccess) {
            if (preservePermissions) {
              const fileMode = getEffectiveMode(entry.mode, "file");
              tryFsOpWithWarningSync(
                () => chmodSync(targetPath, fileMode & 0o7777),
                onWarning,
                entry.path,
                targetPath
              );
            }

            if (preserveTimestamps) {
              tryFsOpWithWarningSync(
                () => setFileTimeSync(targetPath, entry.lastModified),
                onWarning,
                entry.path,
                targetPath
              );
            }
          }
        }
      }

      extractedEntries++;

      if (onProgress) {
        onProgress({
          currentEntry: entry.path,
          totalEntries,
          extractedEntries,
          bytesWritten
        });
      }
    }

    // Process deferred symlinks
    for (let i = 0; i < deferredSymlinks.length; i++) {
      const deferred = deferredSymlinks[i]!;
      const { entry, targetPath, linkTarget } = deferred;
      if (signal?.aborted) {
        throw createAbortError();
      }

      let shouldWrite: boolean;
      try {
        shouldWrite = shouldExtractSync(targetPath, entry.lastModified, overwrite);
      } catch (err) {
        if (isIgnorableFsError(err)) {
          emitExtractWarning(onWarning, entry.path, targetPath, err);
          continue;
        }
        throw toFileError(targetPath, "write", err, "extracted");
      }

      if (!shouldWrite) {
        continue;
      }

      if (
        !tryFsOpWithWarningSync(
          () => ensureDirSync(path.dirname(targetPath)),
          onWarning,
          entry.path,
          targetPath
        )
      ) {
        continue;
      }

      const symlinkCreated = tryFsOpWithWarningSync(
        () => createSymlinkSync(linkTarget, targetPath),
        onWarning,
        entry.path,
        targetPath
      );

      if (symlinkCreated) {
        extractedEntries++;

        if (onProgress) {
          onProgress({
            currentEntry: entry.path,
            totalEntries,
            extractedEntries,
            bytesWritten
          });
        }
      }
    }
  }

  // =============================================================================
  // Private Methods - TAR Building
  // =============================================================================

  private async _buildTar(): Promise<Uint8Array> {
    const archive = await this._tarArchiveOf(await resolveViewAsync(this._tarViewSteps()));
    return this._finishTarBuild(await archive.toBuffer());
  }

  private _buildTarSync(): Uint8Array {
    const archive = new TarArchive({ modTime: this._tar_options.modTime });
    for (const v of resolveViewSync(this._tarViewSteps()).values()) {
      const { data, options } = this._tarEntryOfSync(v);
      archive.add(v.name, data, options);
    }
    return this._finishTarBuild(archive.toBufferSync());
  }

  /**
   * Commit a built TAR, as a ZIP build does: it becomes what reads start from and what the
   * next build starts from, so the pending entries are now part of it. Then compress if asked.
   */
  private _finishTarBuild(tar: Uint8Array): Uint8Array {
    this._loadTar(tar);
    this._tar_pending.length = 0;
    this._tarView = null;
    this._viewGeneration++;
    const opts = this._tar_options;
    return opts?.gzip ? gzipTarSync(tar, opts.gzipLevel ?? DEFAULT_COMPRESS_LEVEL) : tar;
  }

  // =============================================================================
  // Private Methods - Streaming Build
  // =============================================================================

  /**
   * Build the ZIP a write would produce as a stream, with the true streaming ZipArchive.
   * Entries are the view's, in write order and with the metadata a buffered write records;
   * files are read and opened entries decompressed only when the stream reaches them.
   */
  private _buildZipStream(options: ArchiveStreamOptions): ArchiveStreamOperation {
    const globalOptions = this._zip_options;
    const zipArchive = new ZipArchive({
      level: globalOptions.level,
      timestamps: globalOptions.timestamps,
      comment: globalOptions.comment,
      zip64: globalOptions.zip64,
      modTime: globalOptions.modTime,
      reproducible: globalOptions.reproducible,
      smartStore: globalOptions.smartStore,
      encoding: globalOptions.encoding,
      signal: options.signal,
      // Names below are already normalised with these options; passing them on keeps the
      // streaming build to the same naming rules as the buffered one.
      path: resolveZipPathOptions(globalOptions)
    });

    const parser = this._zip_parser;
    const password = this._zip_password;
    for (const v of this._zipOrder(resolveViewSync(this._zipViewSteps()))) {
      let entry: ZipEntry;
      let source: ArchiveSource;
      if (v.base) {
        entry = buildPreservedEntry(v.name, v.base, EMPTY_UINT8ARRAY, globalOptions, password);
        source = v.isDirectory ? EMPTY_UINT8ARRAY : extractLazily(parser!, v.base.path, password);
      } else if (v.found) {
        entry = v.found.build(EMPTY_UINT8ARRAY);
        source = v.found.path === null ? EMPTY_UINT8ARRAY : createReadStream(v.found.path);
      } else {
        const op = v.op!;
        if (op.type === "file") {
          const stats = zipFsMetadata(safeStatsSync(op.localPath));
          entry = buildZipEntry(
            op.zipPath,
            EMPTY_UINT8ARRAY,
            op.options,
            globalOptions,
            password,
            op.stamp,
            stats
          );
          source = createReadStream(op.localPath);
        } else if (op.type === "stream" && !op.buffered) {
          entry = buildZipEntry(
            op.zipPath,
            EMPTY_UINT8ARRAY,
            op.options,
            globalOptions,
            password,
            op.stamp
          );
          source = readPendingStreamOnce(op);
        } else {
          entry = this._zipEntryOfInMemory(op);
          source = entry.data;
        }
      }
      if (isEncryptedZipEntry(entry)) {
        throw new ArchiveError(
          `ArchiveFile.stream: entry "${v.name}" is to be encrypted, which a streamed write ` +
            "cannot do; use toBuffer() or writeFile()"
        );
      }
      zipArchive.add(entry.name, source, zipStreamOptions(entry));
    }

    const zipOp = zipArchive.operation({
      signal: options.signal,
      onProgress: options.onProgress,
      progressIntervalMs: options.progressIntervalMs
    });
    return wrapStreamOperation(zipOp);
  }

  /**
   * Build the TAR a write would produce as a stream, with the true streaming TarArchive.
   */
  private _buildTarStream(options: ArchiveStreamOptions): ArchiveStreamOperation {
    const tarArchive = new TarArchive({
      modTime: this._tar_options.modTime,
      signal: options.signal
    });

    for (const v of resolveViewSync(this._tarViewSteps()).values()) {
      const op = v.op;
      if (op?.type === "file") {
        tarArchive.add(
          v.name,
          createReadStream(op.localPath),
          this._tarFileOptions(op, safeStatsSync(op.localPath))
        );
      } else if (v.found?.path) {
        tarArchive.add(v.name, createReadStream(v.found.path), v.found.options);
      } else if (op?.type === "stream" && !op.buffered) {
        tarArchive.add(
          v.name,
          readPendingStreamOnce(op),
          buildTarAddOptions(op.options, { mtime: op.stamp })
        );
      } else {
        const { data, options: entryOptions } = this._tarEntryOfSync(v);
        tarArchive.add(v.name, data, entryOptions);
      }
    }

    const tarOp = tarArchive.operation({
      signal: options.signal,
      onProgress: options.onProgress,
      progressIntervalMs: options.progressIntervalMs
    });

    const opts = this._tar_options;
    const gzippedIterable = opts?.gzip
      ? gzipTarStream(tarOp.iterable, opts.gzipLevel ?? DEFAULT_COMPRESS_LEVEL)
      : undefined;
    return wrapStreamOperation(tarOp, gzippedIterable);
  }

  // =============================================================================
  // Private Methods - TAR Extraction
  // =============================================================================

  private async _extractTar(targetDir: string, options: ExtractToOptions): Promise<void> {
    const resolvedTargetDir = path.resolve(targetDir);
    const overwrite = options.overwrite ?? "overwrite";
    const { filter, signal, onProgress, onWarning } = options;
    const preserveTimestamps = options.preserveTimestamps ?? true;
    const view = await this._tarViewAsync();

    let extractedEntries = 0;
    let bytesWritten = 0;
    try {
      await ensureDir(resolvedTargetDir);
    } catch (error) {
      throw toFileError(resolvedTargetDir, "write", error, "extracted");
    }

    for (const v of view.values()) {
      if (signal?.aborted) {
        throw createAbortError();
      }
      const entryPath = v.name;
      const targetPath = path.resolve(resolvedTargetDir, entryPath);
      assertNoPathTraversal(targetPath, resolvedTargetDir, entryPath);
      if (filter && !filter(entryPath, v.isDirectory)) {
        continue;
      }

      if (v.isDirectory) {
        if (
          !(await tryFsOpWithWarning(() => ensureDir(targetPath), onWarning, entryPath, targetPath))
        ) {
          continue;
        }
      } else {
        const entry = await this._tarEntryOf(v);
        const mtime = this._tarInfoOf(v, () => entry).mtime;
        let shouldWrite: boolean;
        try {
          shouldWrite = await shouldExtract(targetPath, mtime, overwrite);
        } catch (err) {
          if (isIgnorableFsError(err)) {
            emitExtractWarning(onWarning, entryPath, targetPath, err);
            continue;
          }
          throw toFileError(targetPath, "write", err, "extracted");
        }

        if (shouldWrite) {
          try {
            await ensureDir(path.dirname(targetPath));
            await writeFileBytes(targetPath, entry.data);
            bytesWritten += entry.data.length;
          } catch (err) {
            if (isIgnorableFsError(err)) {
              emitExtractWarning(onWarning, entryPath, targetPath, err);
              continue;
            }
            throw toFileError(targetPath, "write", err, "extracted");
          }
          if (preserveTimestamps) {
            await tryFsOpWithWarning(
              () => setFileTime(targetPath, mtime),
              onWarning,
              entryPath,
              targetPath
            );
          }
        }
      }

      extractedEntries++;
      onProgress?.({
        currentEntry: entryPath,
        totalEntries: view.size,
        extractedEntries,
        bytesWritten
      });
    }
  }

  private _extractTarSync(targetDir: string, options: ExtractToOptions): void {
    const resolvedTargetDir = path.resolve(targetDir);
    const overwrite = options.overwrite ?? "overwrite";
    const { filter, signal, onProgress, onWarning } = options;
    const preserveTimestamps = options.preserveTimestamps ?? true;
    const view = this._tarViewSync();

    let extractedEntries = 0;
    let bytesWritten = 0;
    try {
      ensureDirSync(resolvedTargetDir);
    } catch (error) {
      throw toFileError(resolvedTargetDir, "write", error, "extracted");
    }

    for (const v of view.values()) {
      if (signal?.aborted) {
        throw createAbortError();
      }
      const entryPath = v.name;
      const targetPath = path.resolve(resolvedTargetDir, entryPath);
      assertNoPathTraversal(targetPath, resolvedTargetDir, entryPath);
      if (filter && !filter(entryPath, v.isDirectory)) {
        continue;
      }

      if (v.isDirectory) {
        if (
          !tryFsOpWithWarningSync(() => ensureDirSync(targetPath), onWarning, entryPath, targetPath)
        ) {
          continue;
        }
      } else {
        const entry = this._tarEntryOfSync(v);
        const mtime = this._tarInfoOf(v, () => entry).mtime;
        let shouldWrite: boolean;
        try {
          shouldWrite = shouldExtractSync(targetPath, mtime, overwrite);
        } catch (err) {
          if (isIgnorableFsError(err)) {
            emitExtractWarning(onWarning, entryPath, targetPath, err);
            continue;
          }
          throw toFileError(targetPath, "write", err, "extracted");
        }

        if (shouldWrite) {
          try {
            ensureDirSync(path.dirname(targetPath));
            writeFileBytesSync(targetPath, entry.data);
            bytesWritten += entry.data.length;
          } catch (err) {
            if (isIgnorableFsError(err)) {
              emitExtractWarning(onWarning, entryPath, targetPath, err);
              continue;
            }
            throw toFileError(targetPath, "write", err, "extracted");
          }
          if (preserveTimestamps) {
            tryFsOpWithWarningSync(
              () => setFileTimeSync(targetPath, mtime),
              onWarning,
              entryPath,
              targetPath
            );
          }
        }
      }

      extractedEntries++;
      onProgress?.({
        currentEntry: entryPath,
        totalEntries: view.size,
        extractedEntries,
        bytesWritten
      });
    }
  }
}

/** The ZIP limits an opener takes (a TAR opened here holds every entry anyway). */
const OPEN_ZIP_LIMITS = ["maxEntrySize", "maxEntries"] as const;

/** An opener's options, checked before anything is read, as `Archive.openZip` checks its own. */
function checkOpenOptions(where: string, options: ArchiveFileOpenOptions): void {
  checkReadOptions(
    where,
    options,
    (options as { format?: unknown } | undefined)?.format === "tar" ? [] : OPEN_ZIP_LIMITS
  );
}
