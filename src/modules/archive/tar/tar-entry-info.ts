/**
 * TAR Entry Information
 *
 * Represents metadata about a single entry in a TAR archive.
 */

import type { TarType } from "@archive/tar/tar-constants";
import { TAR_TYPE } from "@archive/tar/tar-constants";

export interface TarEntryInfo {
  /** File path/name */
  path: string;

  /** File type */
  type: TarType;

  /** File size in bytes */
  size: number;

  /** Unix file mode/permissions */
  mode: number;

  /** User ID */
  uid: number;

  /** Group ID */
  gid: number;

  /** User name */
  uname: string;

  /** Group name */
  gname: string;

  /** Modification time */
  mtime: Date;

  /** Link target (for symlinks and hard links) */
  linkname?: string;

  /** Device major number (for device files) */
  devmajor?: number;

  /** Device minor number (for device files) */
  devminor?: number;

  /** PAX extended attributes */
  pax?: Record<string, string>;
}

/**
 * Check if entry is a regular file
 */
export function isFile(entry: TarEntryInfo): boolean {
  return (
    entry.type === TAR_TYPE.FILE ||
    entry.type === TAR_TYPE.FILE_OLD ||
    entry.type === TAR_TYPE.CONTIGUOUS
  );
}

/**
 * Check if entry is a directory
 */
export function isDirectory(entry: TarEntryInfo): boolean {
  return entry.type === TAR_TYPE.DIRECTORY;
}

/**
 * Check if entry is a symbolic link
 */
export function isSymlink(entry: TarEntryInfo): boolean {
  return entry.type === TAR_TYPE.SYMLINK;
}

/**
 * Check if entry is a hard link
 */
export function isHardLink(entry: TarEntryInfo): boolean {
  return entry.type === TAR_TYPE.HARD_LINK;
}

/**
 * Check if entry type has data (files have data, directories/links don't)
 * Used to determine if entry data should be written/read.
 */
export function isDataEntry(type: TarType | undefined): boolean {
  return (
    type === TAR_TYPE.FILE ||
    type === TAR_TYPE.FILE_OLD ||
    type === TAR_TYPE.CONTIGUOUS ||
    type === undefined // Default is file
  );
}
