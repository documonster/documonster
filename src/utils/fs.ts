/**
 * File system utilities for Node.js.
 *
 * This module provides common file system operations used across the library,
 * including directory traversal, glob matching, and file I/O helpers.
 *
 * Supports custom file system injection via `useFs()` for Electron or testing.
 *
 * @module
 */

import * as nodeFs from "node:fs";
import * as nodeFsp from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";

// =============================================================================
// File System Injection
// =============================================================================

/** File system module type (sync APIs) */
export type FsModule = typeof nodeFs;

/** File system promises module type (async APIs) */
export type FsPromisesModule = typeof nodeFsp;

// Internal mutable references
let _fs: FsModule = nodeFs;
let _fsp: FsPromisesModule = nodeFsp;

/**
 * Inject a custom file system module.
 *
 * Useful for:
 * - Electron's `original-fs` to bypass ASAR
 * - Virtual file systems like `memfs` for testing
 *
 * Call without arguments to reset to default Node.js fs.
 *
 * @example
 * ```ts
 * import originalFs from "original-fs";
 * import { useFs } from "@utils/fs";
 *
 * // Use Electron's original-fs
 * useFs(originalFs);
 *
 * // Reset to default
 * useFs();
 * ```
 */
export function useFs(syncFs?: FsModule, asyncFs?: FsPromisesModule): void {
  _fs = syncFs ?? nodeFs;
  _fsp = asyncFs ?? (syncFs?.promises as FsPromisesModule) ?? nodeFsp;
}

// Re-export glob utilities from shared module
export {
  globToRegex,
  matchGlob,
  matchGlobAny,
  createGlobMatcher,
  clearGlobCache,
  normalizePath
} from "./glob";

// =============================================================================
// Types
// =============================================================================

/**
 * Information about a file system entry.
 */
export interface FileEntry {
  /** Absolute path on disk */
  absolutePath: string;

  /** Relative path from the base directory */
  relativePath: string;

  /** Whether this is a directory */
  isDirectory: boolean;

  /** File size in bytes (0 for directories) */
  size: number;

  /** Last modified time */
  mtime: Date;

  /** Last access time */
  atime: Date;

  /** Metadata change time */
  ctime: Date;

  /** Creation time (when supported by the platform) */
  birthTime: Date;

  /** Unix mode (includes file type + permissions). */
  mode: number;
}

/**
 * Options for directory traversal.
 */
export interface TraverseOptions {
  /** Recursively traverse subdirectories (default: true) */
  recursive?: boolean;

  /** Follow symbolic links (default: false) */
  followSymlinks?: boolean;

  /** Filter function */
  filter?: (entry: FileEntry) => boolean;
}

/**
 * Options for glob file matching.
 */
export interface GlobOptions {
  /** Current working directory */
  cwd?: string;

  /** Patterns to ignore */
  ignore?: string | string[];

  /** Include dot files (default: false) */
  dot?: boolean;

  /** Follow symbolic links (default: false) */
  followSymlinks?: boolean;

  /** Filter function */
  filter?: (entry: FileEntry) => boolean;
}

// =============================================================================
// Directory Traversal
// =============================================================================

/**
 * Build a FileEntry from stats.
 */
function buildFileEntry(
  absolutePath: string,
  relativePath: string,
  stats: nodeFs.Stats
): FileEntry {
  const isDirectory = stats.isDirectory();
  return {
    absolutePath,
    relativePath,
    isDirectory,
    size: isDirectory ? 0 : stats.size,
    mtime: stats.mtime,
    atime: stats.atime,
    ctime: stats.ctime,
    birthTime: stats.birthtime,
    mode: stats.mode
  };
}

/**
 * Read the `code` property off an unknown thrown value (Node errors carry it).
 */
function errorCode(err: unknown): string | undefined {
  const code = (err as { code?: unknown }).code;
  return typeof code === "string" ? code : undefined;
}

/**
 * Check if an error is ignorable (file not found or permission denied).
 */
function isIgnorableError(err: unknown): boolean {
  const code = errorCode(err);
  return code === "ENOENT" || code === "EACCES";
}

/**
 * Recursively traverse a directory and yield file entries.
 *
 * @param dirPath - Directory to traverse
 * @param options - Traversal options
 * @yields File entries
 */
export async function* traverseDirectory(
  dirPath: string,
  options: TraverseOptions = {}
): AsyncGenerator<FileEntry> {
  const { recursive = true, followSymlinks = false, filter } = options;
  const basePath = path.resolve(dirPath);

  async function* walk(currentPath: string, relativeTo: string): AsyncGenerator<FileEntry> {
    let entries: nodeFs.Dirent[];
    try {
      entries = await _fsp.readdir(currentPath, { withFileTypes: true });
    } catch (err) {
      if (isIgnorableError(err)) {
        return;
      }
      throw err;
    }

    // Sort entries for deterministic order
    entries.sort((a, b) => a.name.localeCompare(b.name));

    for (const dirent of entries) {
      const absolutePath = path.join(currentPath, dirent.name);
      const relativePath = path.relative(relativeTo, absolutePath);

      let stats: nodeFs.Stats;
      try {
        stats = followSymlinks ? await _fsp.stat(absolutePath) : await _fsp.lstat(absolutePath);
      } catch (err) {
        if (isIgnorableError(err)) {
          continue;
        }
        throw err;
      }

      // Skip symbolic links if not following them
      if (stats.isSymbolicLink() && !followSymlinks) {
        continue;
      }

      const entry = buildFileEntry(absolutePath, relativePath, stats);
      if (filter && !filter(entry)) {
        continue;
      }

      yield entry;

      if (entry.isDirectory && recursive) {
        yield* walk(absolutePath, relativeTo);
      }
    }
  }

  yield* walk(basePath, basePath);
}

/**
 * Synchronously traverse a directory.
 */
export function traverseDirectorySync(dirPath: string, options: TraverseOptions = {}): FileEntry[] {
  const { recursive = true, followSymlinks = false, filter } = options;
  const basePath = path.resolve(dirPath);
  const results: FileEntry[] = [];

  function walk(currentPath: string, relativeTo: string): void {
    let entries: nodeFs.Dirent[];
    try {
      entries = _fs.readdirSync(currentPath, { withFileTypes: true });
    } catch (err) {
      if (isIgnorableError(err)) {
        return;
      }
      throw err;
    }

    entries.sort((a, b) => a.name.localeCompare(b.name));

    for (const dirent of entries) {
      const absolutePath = path.join(currentPath, dirent.name);
      const relativePath = path.relative(relativeTo, absolutePath);

      let stats: nodeFs.Stats;
      try {
        stats = followSymlinks ? _fs.statSync(absolutePath) : _fs.lstatSync(absolutePath);
      } catch (err) {
        if (isIgnorableError(err)) {
          continue;
        }
        throw err;
      }

      if (stats.isSymbolicLink() && !followSymlinks) {
        continue;
      }

      const entry = buildFileEntry(absolutePath, relativePath, stats);
      if (filter && !filter(entry)) {
        continue;
      }

      results.push(entry);

      if (entry.isDirectory && recursive) {
        walk(absolutePath, relativeTo);
      }
    }
  }

  walk(basePath, basePath);
  return results;
}

// =============================================================================
// Glob File Search
// =============================================================================

// Import glob utilities from shared module
import { createGlobMatcher, normalizePath } from "./glob";

/**
 * Parsed glob options with pre-compiled matchers.
 */
interface ParsedGlobOptions {
  basePath: string;
  searchBase: string;
  followSymlinks: boolean;
  filter?: (entry: FileEntry) => boolean;
  ignoreMatcher: ((path: string) => boolean) | null;
  patternMatcher: (path: string) => boolean;
}

/**
 * Parse glob options and pre-compile matchers.
 * Shared between glob() and globSync().
 */
function parseGlobOptions(pattern: string, options: GlobOptions): ParsedGlobOptions {
  const { cwd = process.cwd(), ignore, dot = false, followSymlinks = false, filter } = options;
  const ignorePatterns = ignore ? (Array.isArray(ignore) ? ignore : [ignore]) : [];
  const basePath = path.resolve(cwd);

  // Pre-compile matchers
  const ignoreMatcher =
    ignorePatterns.length > 0 ? createGlobMatcher(ignorePatterns, { dot }) : null;
  const patternMatcher = createGlobMatcher([pattern], { dot });

  // Determine the base directory from the pattern (static prefix optimization)
  const patternParts = pattern.split(/[/\\]/);
  let staticPrefix = "";
  for (const part of patternParts) {
    if (part.includes("*") || part.includes("?") || part.includes("[") || part.includes("{")) {
      break;
    }
    staticPrefix = staticPrefix ? path.join(staticPrefix, part) : part;
  }

  const searchBase = staticPrefix ? path.join(basePath, staticPrefix) : basePath;

  return { basePath, searchBase, followSymlinks, filter, ignoreMatcher, patternMatcher };
}

/**
 * Filter a file entry against glob matchers.
 * Returns the entry with normalized relativePath if matched, null otherwise.
 */
function matchGlobEntry(
  entry: FileEntry,
  basePath: string,
  ignoreMatcher: ((path: string) => boolean) | null,
  patternMatcher: (path: string) => boolean,
  filter?: (entry: FileEntry) => boolean
): FileEntry | null {
  const relativeFromCwd = normalizePath(path.relative(basePath, entry.absolutePath));

  // Skip directories
  if (entry.isDirectory) {
    return null;
  }

  // Check ignore patterns
  if (ignoreMatcher && ignoreMatcher(relativeFromCwd)) {
    return null;
  }

  // Check pattern match
  if (!patternMatcher(relativeFromCwd)) {
    return null;
  }

  // Apply custom filter
  if (filter && !filter(entry)) {
    return null;
  }

  return { ...entry, relativePath: relativeFromCwd };
}

/**
 * Find files matching a glob pattern.
 *
 * @param pattern - Glob pattern to match
 * @param options - Glob options
 * @yields Matching file entries
 */
export async function* glob(pattern: string, options: GlobOptions = {}): AsyncGenerator<FileEntry> {
  const { basePath, searchBase, followSymlinks, filter, ignoreMatcher, patternMatcher } =
    parseGlobOptions(pattern, options);

  // Check if search base exists
  try {
    await _fsp.access(searchBase);
  } catch {
    return;
  }

  // Traverse and filter
  for await (const entry of traverseDirectory(searchBase, { followSymlinks })) {
    const matched = matchGlobEntry(entry, basePath, ignoreMatcher, patternMatcher, filter);
    if (matched) {
      yield matched;
    }
  }
}

/**
 * Synchronously find files matching a glob pattern.
 */
export function globSync(pattern: string, options: GlobOptions = {}): FileEntry[] {
  const { basePath, searchBase, followSymlinks, filter, ignoreMatcher, patternMatcher } =
    parseGlobOptions(pattern, options);

  try {
    _fs.accessSync(searchBase);
  } catch {
    return [];
  }

  const results: FileEntry[] = [];
  const entries = traverseDirectorySync(searchBase, { followSymlinks });

  for (const entry of entries) {
    const matched = matchGlobEntry(entry, basePath, ignoreMatcher, patternMatcher, filter);
    if (matched) {
      results.push(matched);
    }
  }

  return results;
}

// =============================================================================
// File I/O Helpers
// =============================================================================

/**
 * Check if a file exists.
 */
export async function fileExists(filePath: string): Promise<boolean> {
  try {
    await _fsp.access(filePath, _fs.constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

/**
 * Synchronously check if a file exists.
 */
export function fileExistsSync(filePath: string): boolean {
  try {
    _fs.accessSync(filePath, _fs.constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

/**
 * Ensure a directory exists, creating it recursively if needed.
 */
export async function ensureDir(dirPath: string): Promise<void> {
  try {
    await _fsp.mkdir(dirPath, { recursive: true });
  } catch (err) {
    if (errorCode(err) !== "EEXIST") {
      throw err;
    }
  }
}

/**
 * Synchronously ensure a directory exists.
 */
export function ensureDirSync(dirPath: string): void {
  try {
    _fs.mkdirSync(dirPath, { recursive: true });
  } catch (err) {
    if (errorCode(err) !== "EEXIST") {
      throw err;
    }
  }
}

/**
 * Get file stats, or null if file doesn't exist.
 */
export async function safeStats(filePath: string): Promise<nodeFs.Stats | null> {
  try {
    return await _fsp.stat(filePath);
  } catch {
    return null;
  }
}

/**
 * Synchronously get file stats, or null if file doesn't exist.
 */
export function safeStatsSync(filePath: string): nodeFs.Stats | null {
  try {
    return _fs.statSync(filePath);
  } catch {
    return null;
  }
}

/**
 * Read a file as Uint8Array.
 */
export async function readFileBytes(filePath: string): Promise<Uint8Array> {
  const buffer = await _fsp.readFile(filePath);
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

/**
 * Synchronously read a file as Uint8Array.
 */
export function readFileBytesSync(filePath: string): Uint8Array {
  const buffer = _fs.readFileSync(filePath);
  return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
}

/**
 * Write bytes to a file.
 */
export async function writeFileBytes(filePath: string, data: Uint8Array): Promise<void> {
  await _fsp.writeFile(filePath, data);
}

/**
 * Synchronously write bytes to a file.
 */
export function writeFileBytesSync(filePath: string, data: Uint8Array): void {
  _fs.writeFileSync(filePath, data);
}

/**
 * Set file modification time.
 */
export async function setFileTime(filePath: string, mtime: Date): Promise<void> {
  await _fsp.utimes(filePath, mtime, mtime);
}

/**
 * Synchronously set file modification time.
 */
export function setFileTimeSync(filePath: string, mtime: Date): void {
  _fs.utimesSync(filePath, mtime, mtime);
}

/**
 * Read file as text.
 */
export async function readFileText(
  filePath: string,
  encoding: BufferEncoding = "utf8"
): Promise<string> {
  return _fsp.readFile(filePath, { encoding });
}

/**
 * Synchronously read file as text.
 */
export function readFileTextSync(filePath: string, encoding: BufferEncoding = "utf8"): string {
  return _fs.readFileSync(filePath, { encoding });
}

/**
 * Write text to a file.
 */
export async function writeFileText(
  filePath: string,
  content: string,
  encoding: BufferEncoding = "utf8"
): Promise<void> {
  await _fsp.writeFile(filePath, content, { encoding });
}

/**
 * Synchronously write text to a file.
 */
export function writeFileTextSync(
  filePath: string,
  content: string,
  encoding: BufferEncoding = "utf8"
): void {
  _fs.writeFileSync(filePath, content, { encoding });
}

/**
 * Remove a file or directory.
 */
export async function remove(targetPath: string): Promise<void> {
  try {
    await _fsp.rm(targetPath, { recursive: true, force: true });
  } catch {
    // Ignore errors (file may not exist)
  }
}

/**
 * Synchronously remove a file or directory.
 */
export function removeSync(targetPath: string): void {
  try {
    _fs.rmSync(targetPath, { recursive: true, force: true });
  } catch {
    // Ignore errors (file may not exist)
  }
}

/**
 * Copy a file.
 */
export async function copyFile(src: string, dest: string): Promise<void> {
  await ensureDir(path.dirname(dest));
  await _fsp.copyFile(src, dest);
}

/**
 * Synchronously copy a file.
 */
export function copyFileSync(src: string, dest: string): void {
  ensureDirSync(path.dirname(dest));
  _fs.copyFileSync(src, dest);
}

// =============================================================================
// Symlinks and Permissions
// =============================================================================

/**
 * Create a symbolic link.
 *
 * @param target - The path the symlink points to
 * @param linkPath - The path where the symlink will be created
 */
export async function createSymlink(target: string, linkPath: string): Promise<void> {
  await ensureDir(path.dirname(linkPath));
  await _fsp.symlink(target, linkPath);
}

/**
 * Synchronously create a symbolic link.
 *
 * @param target - The path the symlink points to
 * @param linkPath - The path where the symlink will be created
 */
export function createSymlinkSync(target: string, linkPath: string): void {
  ensureDirSync(path.dirname(linkPath));
  _fs.symlinkSync(target, linkPath);
}

/**
 * Change file permissions (Unix mode).
 *
 * @param filePath - Path to the file
 * @param mode - Unix permission mode (e.g., 0o755)
 */
export async function chmod(filePath: string, mode: number): Promise<void> {
  await _fsp.chmod(filePath, mode);
}

/**
 * Synchronously change file permissions (Unix mode).
 *
 * @param filePath - Path to the file
 * @param mode - Unix permission mode (e.g., 0o755)
 */
export function chmodSync(filePath: string, mode: number): void {
  _fs.chmodSync(filePath, mode);
}

/**
 * Check if the current platform supports Unix permissions.
 * Returns true on Unix-like systems (Linux, macOS), false on Windows.
 */
export function supportsUnixPermissions(): boolean {
  return process.platform !== "win32";
}

// =============================================================================
// File Streams
// =============================================================================

/**
 * Options for creating a read stream.
 */
export interface ReadStreamOptions {
  /** File encoding (default: none, returns Buffer) */
  encoding?: BufferEncoding | null;
  /** High water mark for internal buffer (default: 64KB) */
  highWaterMark?: number;
  /** Start position in bytes */
  start?: number;
  /** End position in bytes */
  end?: number;
  /** Auto close on end or error (default: true) */
  autoClose?: boolean;
}

/**
 * Options for creating a write stream.
 */
export interface WriteStreamOptions {
  /** File encoding (default: 'utf8') */
  encoding?: BufferEncoding;
  /** High water mark for internal buffer (default: 64KB) */
  highWaterMark?: number;
  /** File flags (default: 'w') */
  flags?: string;
  /** File mode (default: 0o666) */
  mode?: number;
  /** Auto close on end or error (default: true) */
  autoClose?: boolean;
}

/**
 * Create a readable stream from a file.
 *
 * @param filePath - Path to the file
 * @param options - Stream options
 * @returns A readable stream
 */
export function createReadStream(filePath: string, options?: ReadStreamOptions): nodeFs.ReadStream {
  return _fs.createReadStream(filePath, options as nodeFs.ReadStreamOptions);
}

/**
 * Create a writable stream to a file.
 *
 * @param filePath - Path to the file
 * @param options - Stream options
 * @returns A writable stream
 */
export function createWriteStream(
  filePath: string,
  options?: WriteStreamOptions
): nodeFs.WriteStream {
  return _fs.createWriteStream(filePath, options);
}

/**
 * Create a temporary directory.
 *
 * @param prefix - Prefix for the directory name
 * @returns Path to the created directory
 */
export async function createTempDir(prefix: string = "tmp-"): Promise<string> {
  return _fsp.mkdtemp(path.join(os.tmpdir(), prefix));
}

/**
 * Synchronously create a temporary directory.
 *
 * @param prefix - Prefix for the directory name
 * @returns Path to the created directory
 */
export function createTempDirSync(prefix: string = "tmp-"): string {
  return _fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}
