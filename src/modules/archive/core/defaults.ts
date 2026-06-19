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
