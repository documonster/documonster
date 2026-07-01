/**
 * Shared view tracker for ZIP editing operations.
 *
 * This module provides a unified way to track pending edits (set, delete, rename)
 * for both ZipEditor and ZipFile classes, reducing code duplication.
 */

import { ArchiveError } from "@archive/core/errors";
import type { ArchiveSource } from "@archive/io/archive-source";
import type { ZipPathOptions } from "@archive/zip-spec/zip-path";
import { normalizeZipPath } from "@archive/zip-spec/zip-path";
import type { ZipEntryOptions } from "@archive/zip/index";

// =============================================================================
// Types
// =============================================================================

/**
 * A "base" entry represents an original entry from the source archive
 * that should be preserved (possibly renamed).
 */
export interface BaseViewEntry<TInfo> {
  kind: "base";
  info: TInfo;
}

/**
 * Internal representation of a "set" entry (without name, stored as Map key).
 */
interface InternalSetEntry {
  kind: "set";
  source: ArchiveSource;
  options?: ZipEntryOptions;
}

/**
 * A "set" entry represents a new or updated entry with fresh content.
 * Returned by getSetEntries() with name derived from Map key.
 */
export interface SetViewEntry {
  kind: "set";
  name: string;
  source: ArchiveSource;
  options?: ZipEntryOptions;
}

/**
 * Union type for all view entry kinds (internal storage).
 */
type InternalViewEntry<TInfo> = BaseViewEntry<TInfo> | InternalSetEntry;

/**
 * Options for ZipEditView.
 */
export interface ZipEditViewOptions {
  /**
   * Path normalization mode.
   * - `false`: no normalization, names are used as-is
   * - `ZipPathOptions`: apply normalization rules
   */
  path?: false | ZipPathOptions;
}

// =============================================================================
// ZipEditView Class
// =============================================================================

/**
 * Tracks pending edits to a ZIP archive in a unified manner.
 *
 * This class manages:
 * - Base entries (preserved from original archive)
 * - Set entries (new or updated content)
 * - Deletions and renames
 *
 * @typeParam TInfo - The type of entry info stored for base entries
 */
export class ZipEditView<TInfo> {
  private readonly _view: Map<string, InternalViewEntry<TInfo>> = new Map();
  private readonly _pathOptions: false | ZipPathOptions;
  private _originalSize: number = 0;

  constructor(options: ZipEditViewOptions = {}) {
    this._pathOptions = options.path ?? false;
  }

  /**
   * Normalize an entry name according to path options.
   */
  private _normalize(name: string): string {
    if (!name) {
      throw new ArchiveError("Entry name is required");
    }
    if (this._pathOptions === false) {
      return name;
    }
    return normalizeZipPath(name, this._pathOptions);
  }

  // ===========================================================================
  // Initialization
  // ===========================================================================

  /**
   * Initialize the view with base entries from an existing archive.
   *
   * @param entries - Array of entries with their info
   * @param getPath - Function to extract the path from entry info
   */
  initFromEntries(entries: readonly TInfo[], getPath: (info: TInfo) => string): void {
    this._view.clear();
    for (const info of entries) {
      const path = getPath(info);
      // Normalize the path so lookups via has()/delete() work correctly
      const normalizedPath = this._normalize(path);
      this._view.set(normalizedPath, { kind: "base", info });
    }
    this._originalSize = this._view.size;
  }

  // ===========================================================================
  // Query Operations
  // ===========================================================================

  /**
   * Check if an entry exists (considering pending edits).
   */
  has(name: string): boolean {
    const n = this._normalize(name);
    return this._view.has(n);
  }

  /**
   * Get all output entry names (after applying edits).
   */
  getOutputNames(): string[] {
    return Array.from(this._view.keys());
  }

  /**
   * Get the number of entries in the view.
   */
  get size(): number {
    return this._view.size;
  }

  /**
   * Iterate over all entries in the view.
   */
  entries(): IterableIterator<[string, InternalViewEntry<TInfo>]> {
    return this._view.entries();
  }

  /**
   * Get a specific entry by name.
   */
  get(name: string): InternalViewEntry<TInfo> | undefined {
    const n = this._normalize(name);
    return this._view.get(n);
  }

  // ===========================================================================
  // Edit Operations
  // ===========================================================================

  /**
   * Delete an entry from the view.
   *
   * @returns `true` if the entry existed and was deleted
   */
  delete(name: string): boolean {
    const n = this._normalize(name);
    return this._view.delete(n);
  }

  /**
   * Delete a directory and all its contents recursively.
   *
   * This method deletes the directory entry itself (if it exists) and all entries
   * whose paths start with the directory prefix (e.g., "folder/" will delete
   * "folder/", "folder/file.txt", "folder/sub/file.txt", etc.).
   *
   * @param prefix - The directory path prefix to delete (with or without trailing slash)
   * @returns The number of entries deleted
   *
   * @example
   * ```ts
   * // Delete "assets/" and all files/folders inside it
   * const count = view.deleteDirectory("assets");
   * // Or with trailing slash (same result)
   * const count = view.deleteDirectory("assets/");
   * ```
   */
  deleteDirectory(prefix: string): number {
    const normalizedPrefix = this._normalize(prefix);
    // Ensure prefix ends with "/" for proper matching
    const dirPrefix = normalizedPrefix.endsWith("/") ? normalizedPrefix : normalizedPrefix + "/";

    let deleted = 0;
    const toDelete: string[] = [];
    for (const name of this._view.keys()) {
      // Match: exact directory entry (without slash) OR anything inside the directory
      if (name === normalizedPrefix || name.startsWith(dirPrefix)) {
        toDelete.push(name);
      }
    }

    for (const name of toDelete) {
      if (this._view.delete(name)) {
        deleted++;
      }
    }
    return deleted;
  }

  /**
   * Add or update an entry with new content.
   */
  set(name: string, source: ArchiveSource, options?: ZipEntryOptions): void {
    const n = this._normalize(name);
    this._view.set(n, { kind: "set", source, options });
  }

  /**
   * Rename an entry.
   *
   * **Overwrite behavior**: If an entry with the target name already exists,
   * it will be replaced (similar to `mv -f`).
   *
   * @returns `true` if the rename was successful, `false` if source doesn't exist
   */
  rename(from: string, to: string): boolean {
    const src = this._normalize(from);
    const dst = this._normalize(to);

    if (src === dst) {
      return this.has(src);
    }

    const node = this._view.get(src);
    if (!node) {
      return false;
    }

    // Overwrite semantics: destination is replaced
    this._view.delete(dst);
    this._view.delete(src);

    // Move node to new key (name is derived from key, no need to update)
    this._view.set(dst, node);
    return true;
  }

  // ===========================================================================
  // Categorized Access
  // ===========================================================================

  /**
   * Get all base (preserved) entries.
   */
  getBaseEntries(): Array<{ name: string; info: TInfo }> {
    const result: Array<{ name: string; info: TInfo }> = [];
    for (const [name, entry] of this._view) {
      if (entry.kind === "base") {
        result.push({ name, info: entry.info });
      }
    }
    return result;
  }

  /**
   * Get all set (new/updated) entries.
   */
  getSetEntries(): SetViewEntry[] {
    const result: SetViewEntry[] = [];
    for (const [name, entry] of this._view) {
      if (entry.kind === "set") {
        result.push({ kind: "set", name, source: entry.source, options: entry.options });
      }
    }
    return result;
  }

  // ===========================================================================
  // State Management
  // ===========================================================================

  /**
   * Check if there are any pending changes (additions, updates, or deletions).
   */
  hasChanges(): boolean {
    // Check if any entries were deleted
    if (this._view.size !== this._originalSize) {
      return true;
    }
    // Check if any entries were updated (set)
    for (const entry of this._view.values()) {
      if (entry.kind === "set") {
        return true;
      }
    }
    return false;
  }

  /**
   * Clear all entries.
   */
  clear(): void {
    this._view.clear();
    this._originalSize = 0;
  }
}
