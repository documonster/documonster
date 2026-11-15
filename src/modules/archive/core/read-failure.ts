/**
 * What a document reader needs to know about a failed read of a ZIP package, to report it as its own error: whether
 * the archive was describing the bytes it was given, and whether those bytes started as a ZIP package does. Shared by
 * the Excel and Word readers, which sit beside each other and cannot import one another, so they cannot disagree
 * about what counts as damage.
 */
import {
  ArchiveDecompressionError,
  ZipCrc32MismatchError,
  ZipEntrySizeMismatchError,
  ZipParseError
} from "@archive/core/errors";
import type { ArchiveError } from "@archive/core/errors";

/**
 * Whether `error` is the archive reporting something wrong with the bytes it read: a ZIP record that is not one
 * (`ZipParseError`, which a truncated package raises too), a part whose bytes do not match its CRC-32 or its declared
 * size, compressed data that does not decompress (`ArchiveDecompressionError`).
 *
 * Decided by what an error is, never by what it is not: anything else — a configured limit, an entry's own ZIP
 * encryption, a compression method the archive does not implement, a mistake in how it was called, a worker that
 * failed, an abort — says nothing about the bytes, and a reader must not tell its caller that the file is damaged.
 */
export function isArchiveDataError(error: unknown): error is ArchiveError {
  return (
    error instanceof ZipParseError ||
    error instanceof ZipCrc32MismatchError ||
    error instanceof ZipEntrySizeMismatchError ||
    error instanceof ArchiveDecompressionError
  );
}

/**
 * Whether `bytes` start with a ZIP local-file header signature (`PK\x03\x04`), as every OOXML package does — what tells
 * a damaged package (a truncated one keeps its first bytes) from input that never was one.
 */
export function startsLikeZip(bytes: Uint8Array | undefined): boolean {
  return (
    bytes !== undefined &&
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}
