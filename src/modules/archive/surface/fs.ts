/**
 * `Fs` namespace surface (Node only) — the file-system helpers the archive layer is
 * built on: traversal, globbing, existence checks, byte/text I/O, temp dirs.
 *
 * `import { Fs } from "documonster/archive"` → `Fs.glob("**\/*.json")`,
 * `Fs.ensureDir(dir)`, `Fs.readFileBytes(path)`, …
 *
 * A namespace rather than flat exports because these names (`glob`, `remove`,
 * `copyFile`) are generic enough to collide with `node:fs` and with a consumer's own
 * helpers. It lives on the archive entry because it is the layer
 * `ArchiveFile` (`addGlob`, `extractTo`, `writeFile`) is built on: the glob semantics
 * and the overwrite/permission behaviour an archive caller sees are these helpers', and
 * publishing them anywhere else would be a second entry for one implementation.
 *
 * Arguments are checked here, at the boundary: a wrong one is an `ArchiveArgumentError`
 * naming the call and the parameter. What the file system then refuses (`ENOENT`,
 * `EACCES`, …) is passed on as the platform's error, unwrapped. File streams are not
 * re-published: `createReadStream` / `createWriteStream` from `node:fs` are the ones to use.
 */

import type { Stats } from "node:fs";

import { OPTIONS, checkArchiveArgument, dateSchema } from "@archive/core/arguments";
import { ArchiveArgumentError } from "@archive/core/errors";
import type { FileEntry, GlobOptions, TraverseOptions } from "@utils/fs";
import {
  copyFile as copyFileImpl,
  copyFileSync as copyFileSyncImpl,
  createTempDir as createTempDirImpl,
  createTempDirSync as createTempDirSyncImpl,
  ensureDir as ensureDirImpl,
  ensureDirSync as ensureDirSyncImpl,
  fileExists as fileExistsImpl,
  fileExistsSync as fileExistsSyncImpl,
  glob as globImpl,
  globSync as globSyncImpl,
  globToRegex as globToRegexImpl,
  matchGlob as matchGlobImpl,
  matchGlobAny as matchGlobAnyImpl,
  readFileBytes as readFileBytesImpl,
  readFileBytesSync as readFileBytesSyncImpl,
  readFileText as readFileTextImpl,
  readFileTextSync as readFileTextSyncImpl,
  remove as removeImpl,
  removeSync as removeSyncImpl,
  safeStats as safeStatsImpl,
  safeStatsSync as safeStatsSyncImpl,
  setFileTime as setFileTimeImpl,
  setFileTimeSync as setFileTimeSyncImpl,
  traverseDirectory as traverseDirectoryImpl,
  traverseDirectorySync as traverseDirectorySyncImpl,
  writeFileBytes as writeFileBytesImpl,
  writeFileBytesSync as writeFileBytesSyncImpl,
  writeFileText as writeFileTextImpl,
  writeFileTextSync as writeFileTextSyncImpl
} from "@utils/fs";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

const optionalBoolean = () => s.boolean().optional();

/** A path string. Node refuses a NUL byte with a `TypeError` of its own, so that is refused here too. */
function checkPath(where: string, field: string, value: unknown): void {
  checkArchiveArgument(s.string({ expected: "a path string" }), value, where, field);
  if (value.includes("\0")) {
    throw new ArchiveArgumentError(field, `${where}: ${field} must not contain a NUL byte`);
  }
}

function checkGlobPattern(where: string, field: string, value: unknown): void {
  checkArchiveArgument(s.string({ expected: "a glob pattern string" }), value, where, field);
}

const fileFilter = () => s.function<(entry: FileEntry) => boolean>().optional();

const matchOptionsSchema = s.object({ dot: optionalBoolean() }, OPTIONS).optional();

const traverseOptionsSchema = s.object(
  { recursive: optionalBoolean(), followSymlinks: optionalBoolean(), filter: fileFilter() },
  OPTIONS
);
const _traverseOptionsMatch: SchemaMatches<typeof traverseOptionsSchema, TraverseOptions> = true;
void _traverseOptionsMatch;

const globOptionsSchema = s.object(
  {
    cwd: s.string({ expected: "a path string" }).optional(),
    ignore: s
      .union(
        [
          s.string({ expected: "a glob pattern string" }),
          s.array(s.string({ expected: "a glob pattern string" }))
        ],
        { expected: "a glob pattern or an array of them" }
      )
      .optional(),
    dot: optionalBoolean(),
    followSymlinks: optionalBoolean(),
    filter: fileFilter()
  },
  OPTIONS
);
const _globOptionsMatch: SchemaMatches<typeof globOptionsSchema, GlobOptions> = true;
void _globOptionsMatch;

function checkMatchOptions(where: string, options: unknown): void {
  checkArchiveArgument(matchOptionsSchema, options, where, "options");
}

function checkTraverseOptions(where: string, options: unknown): void {
  checkArchiveArgument(traverseOptionsSchema.optional(), options, where, "options");
}

function checkGlobOptions(where: string, options: unknown): void {
  checkArchiveArgument(globOptionsSchema.optional(), options, where, "options");
  if (options?.cwd !== undefined) {
    checkPath(where, "options.cwd", options.cwd);
  }
}

function checkData(where: string, value: unknown): void {
  checkArchiveArgument(
    s.custom(v => v instanceof Uint8Array, "a Uint8Array"),
    value,
    where,
    "data"
  );
}

function checkContent(where: string, value: unknown): void {
  checkArchiveArgument(s.string(), value, where, "content");
}

function checkEncoding(where: string, value: unknown): void {
  checkArchiveArgument(
    s
      .custom(v => typeof v === "string" && Buffer.isEncoding(v), "a Buffer encoding name")
      .optional(),
    value,
    where,
    "encoding"
  );
}

function checkMtime(where: string, value: unknown): void {
  checkArchiveArgument(dateSchema, value, where, "mtime");
}

function checkPrefix(where: string, value: unknown): void {
  if (value !== undefined) {
    checkPath(where, "prefix", value);
  }
}

// =============================================================================
// Directory traversal and globbing
// =============================================================================

/**
 * Recursively traverse a directory and yield file entries, in name order. An entry that
 * vanishes or cannot be read while walking is skipped.
 *
 * @param dirPath - Directory to traverse
 * @param options - Traversal options
 */
export function traverseDirectory(
  dirPath: string,
  options?: TraverseOptions
): AsyncGenerator<FileEntry> {
  checkPath("Fs.traverseDirectory", "dirPath", dirPath);
  checkTraverseOptions("Fs.traverseDirectory", options);
  return traverseDirectoryImpl(dirPath, options);
}

/** Synchronously traverse a directory; see {@link traverseDirectory}. */
export function traverseDirectorySync(dirPath: string, options?: TraverseOptions): FileEntry[] {
  checkPath("Fs.traverseDirectorySync", "dirPath", dirPath);
  checkTraverseOptions("Fs.traverseDirectorySync", options);
  return traverseDirectorySyncImpl(dirPath, options);
}

/**
 * Find files matching a glob pattern.
 *
 * @param pattern - Glob pattern to match, relative to `options.cwd`
 * @param options - Glob options
 */
export function glob(pattern: string, options?: GlobOptions): AsyncGenerator<FileEntry> {
  checkGlobPattern("Fs.glob", "pattern", pattern);
  checkGlobOptions("Fs.glob", options);
  return globImpl(pattern, options);
}

/** Synchronously find files matching a glob pattern; see {@link glob}. */
export function globSync(pattern: string, options?: GlobOptions): FileEntry[] {
  checkGlobPattern("Fs.globSync", "pattern", pattern);
  checkGlobOptions("Fs.globSync", options);
  return globSyncImpl(pattern, options);
}

/**
 * Convert a glob pattern to a regular expression (`*`, `**`, `?`, `[…]`, `{a,b}`).
 *
 * @param pattern - Glob pattern
 * @param options - `dot`: let wildcards match a leading `.`
 */
export function globToRegex(pattern: string, options?: { dot?: boolean }): RegExp {
  checkGlobPattern("Fs.globToRegex", "pattern", pattern);
  checkMatchOptions("Fs.globToRegex", options);
  return globToRegexImpl(pattern, options);
}

/**
 * Test whether a path matches a glob pattern.
 *
 * @param filePath - File path to test
 * @param pattern - Glob pattern
 * @param options - Match options
 */
export function matchGlob(filePath: string, pattern: string, options?: { dot?: boolean }): boolean {
  checkPath("Fs.matchGlob", "filePath", filePath);
  checkGlobPattern("Fs.matchGlob", "pattern", pattern);
  checkMatchOptions("Fs.matchGlob", options);
  return matchGlobImpl(filePath, pattern, options);
}

/**
 * Test whether a path matches any of the given glob patterns.
 *
 * @param filePath - File path to test
 * @param patterns - Glob patterns
 * @param options - Match options
 */
export function matchGlobAny(
  filePath: string,
  patterns: readonly string[],
  options?: { dot?: boolean }
): boolean {
  checkPath("Fs.matchGlobAny", "filePath", filePath);
  checkArchiveArgument(
    s.array(s.string({ expected: "a glob pattern string" }), {
      expected: "an array of glob patterns"
    }),
    patterns,
    "Fs.matchGlobAny",
    "patterns"
  );
  checkMatchOptions("Fs.matchGlobAny", options);
  return matchGlobAnyImpl(filePath, patterns, options);
}

// =============================================================================
// File I/O
// =============================================================================

/** Whether a file or directory exists at `filePath`. */
export async function fileExists(filePath: string): Promise<boolean> {
  checkPath("Fs.fileExists", "filePath", filePath);
  return fileExistsImpl(filePath);
}

/** Synchronously check whether a file or directory exists at `filePath`. */
export function fileExistsSync(filePath: string): boolean {
  checkPath("Fs.fileExistsSync", "filePath", filePath);
  return fileExistsSyncImpl(filePath);
}

/** Ensure a directory exists, creating it and its parents if needed. */
export async function ensureDir(dirPath: string): Promise<void> {
  checkPath("Fs.ensureDir", "dirPath", dirPath);
  return ensureDirImpl(dirPath);
}

/** Synchronously ensure a directory exists. */
export function ensureDirSync(dirPath: string): void {
  checkPath("Fs.ensureDirSync", "dirPath", dirPath);
  ensureDirSyncImpl(dirPath);
}

/** The stats of `filePath`, or `undefined` when it cannot be read. */
export async function safeStats(filePath: string): Promise<Stats | undefined> {
  checkPath("Fs.safeStats", "filePath", filePath);
  return safeStatsImpl(filePath);
}

/** Synchronously get the stats of `filePath`, or `undefined` when it cannot be read. */
export function safeStatsSync(filePath: string): Stats | undefined {
  checkPath("Fs.safeStatsSync", "filePath", filePath);
  return safeStatsSyncImpl(filePath);
}

/** Read a file as a `Uint8Array`. */
export async function readFileBytes(filePath: string): Promise<Uint8Array> {
  checkPath("Fs.readFileBytes", "filePath", filePath);
  return readFileBytesImpl(filePath);
}

/** Synchronously read a file as a `Uint8Array`. */
export function readFileBytesSync(filePath: string): Uint8Array {
  checkPath("Fs.readFileBytesSync", "filePath", filePath);
  return readFileBytesSyncImpl(filePath);
}

/** Write bytes to a file, replacing it. */
export async function writeFileBytes(filePath: string, data: Uint8Array): Promise<void> {
  checkPath("Fs.writeFileBytes", "filePath", filePath);
  checkData("Fs.writeFileBytes", data);
  return writeFileBytesImpl(filePath, data);
}

/** Synchronously write bytes to a file, replacing it. */
export function writeFileBytesSync(filePath: string, data: Uint8Array): void {
  checkPath("Fs.writeFileBytesSync", "filePath", filePath);
  checkData("Fs.writeFileBytesSync", data);
  writeFileBytesSyncImpl(filePath, data);
}

/** Set a file's modification (and access) time. */
export async function setFileTime(filePath: string, mtime: Date): Promise<void> {
  checkPath("Fs.setFileTime", "filePath", filePath);
  checkMtime("Fs.setFileTime", mtime);
  return setFileTimeImpl(filePath, mtime);
}

/** Synchronously set a file's modification (and access) time. */
export function setFileTimeSync(filePath: string, mtime: Date): void {
  checkPath("Fs.setFileTimeSync", "filePath", filePath);
  checkMtime("Fs.setFileTimeSync", mtime);
  setFileTimeSyncImpl(filePath, mtime);
}

/** Read a file as text (UTF-8 unless `encoding` says otherwise). */
export async function readFileText(filePath: string, encoding?: BufferEncoding): Promise<string> {
  checkPath("Fs.readFileText", "filePath", filePath);
  checkEncoding("Fs.readFileText", encoding);
  return readFileTextImpl(filePath, encoding);
}

/** Synchronously read a file as text (UTF-8 unless `encoding` says otherwise). */
export function readFileTextSync(filePath: string, encoding?: BufferEncoding): string {
  checkPath("Fs.readFileTextSync", "filePath", filePath);
  checkEncoding("Fs.readFileTextSync", encoding);
  return readFileTextSyncImpl(filePath, encoding);
}

/** Write text to a file, replacing it (UTF-8 unless `encoding` says otherwise). */
export async function writeFileText(
  filePath: string,
  content: string,
  encoding?: BufferEncoding
): Promise<void> {
  checkPath("Fs.writeFileText", "filePath", filePath);
  checkContent("Fs.writeFileText", content);
  checkEncoding("Fs.writeFileText", encoding);
  return writeFileTextImpl(filePath, content, encoding);
}

/** Synchronously write text to a file, replacing it (UTF-8 unless `encoding` says otherwise). */
export function writeFileTextSync(
  filePath: string,
  content: string,
  encoding?: BufferEncoding
): void {
  checkPath("Fs.writeFileTextSync", "filePath", filePath);
  checkContent("Fs.writeFileTextSync", content);
  checkEncoding("Fs.writeFileTextSync", encoding);
  writeFileTextSyncImpl(filePath, content, encoding);
}

/** Remove a file or directory, recursively; a path that does not exist is not an error. */
export async function remove(targetPath: string): Promise<void> {
  checkPath("Fs.remove", "targetPath", targetPath);
  return removeImpl(targetPath);
}

/** Synchronously remove a file or directory, recursively. */
export function removeSync(targetPath: string): void {
  checkPath("Fs.removeSync", "targetPath", targetPath);
  removeSyncImpl(targetPath);
}

/** Copy a file, creating the destination's directory if needed. */
export async function copyFile(src: string, dest: string): Promise<void> {
  checkPath("Fs.copyFile", "src", src);
  checkPath("Fs.copyFile", "dest", dest);
  return copyFileImpl(src, dest);
}

/** Synchronously copy a file, creating the destination's directory if needed. */
export function copyFileSync(src: string, dest: string): void {
  checkPath("Fs.copyFileSync", "src", src);
  checkPath("Fs.copyFileSync", "dest", dest);
  copyFileSyncImpl(src, dest);
}

/**
 * Create a new directory under the system temp directory.
 *
 * @param prefix - Prefix for the directory name (default `"tmp-"`)
 * @returns Path to the created directory
 */
export async function createTempDir(prefix?: string): Promise<string> {
  checkPrefix("Fs.createTempDir", prefix);
  return createTempDirImpl(prefix);
}

/** Synchronously create a new directory under the system temp directory; see {@link createTempDir}. */
export function createTempDirSync(prefix?: string): string {
  checkPrefix("Fs.createTempDirSync", prefix);
  return createTempDirSyncImpl(prefix);
}
