import type { ZipTimestampMode } from "@archive/zip-spec/timestamps";

export const DEFAULT_DEFLATE_LEVEL = 6;

// Backward-compatible aliases (avoid default drift across modules).
export const DEFAULT_COMPRESS_LEVEL = DEFAULT_DEFLATE_LEVEL;
export const DEFAULT_ZIP_LEVEL = DEFAULT_DEFLATE_LEVEL;

// Prefer reproducible output by default: omit the Info-ZIP UTC mtime extra field.
export const DEFAULT_ZIP_TIMESTAMPS: ZipTimestampMode = "dos";

/**
 * Default modification time for reproducible ZIP archives.
 *
 * Uses 1980-01-01 00:00:00 (local time), the earliest valid DOS timestamp.
 * Shared to ensure consistent reproducible output across all ZIP creation paths.
 */
export const REPRODUCIBLE_ZIP_MOD_TIME = new Date(1980, 0, 1, 0, 0, 0);

/**
 * Default bound on one entry's uncompressed size when a reader is given no `maxEntrySize`: 512 MiB.
 *
 * Enforced twice: against the *declared* size before anything is inflated, and as the inflater's output bound, so an
 * entry that lies about its size is stopped as soon as it overruns rather than after it has been materialised.
 */
export const DEFAULT_MAX_ENTRY_SIZE = 512 * 1024 * 1024;
