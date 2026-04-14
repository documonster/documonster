import type {
  ArchiveProgressPhase,
  ArchiveStreamOptions,
  ArchiveOperationBase
} from "@archive/shared/progress";
import type { ZipEntryType } from "@archive/zip-spec/zip-entry-info";

import type { UnzipEntry } from "./index";

/**
 * Progress phase for unzip operations.
 */
export type UnzipProgressPhase = ArchiveProgressPhase;

/**
 * Progress information for unzip operations.
 */
export type UnzipProgress = {
  type: "unzip";
  phase: UnzipProgressPhase;

  /** Total bytes consumed from the source stream so far. */
  bytesIn: number;

  /** Total decompressed bytes yielded to consumers so far (best-effort). */
  bytesOut: number;

  /** Number of entries emitted by the streaming parser. */
  entriesEmitted: number;

  currentEntry?: {
    path: string;
    entryType: ZipEntryType;
    bytesOut: number;
  };
};

/**
 * Streaming options for unzip operations.
 */
export type UnzipStreamOptions = ArchiveStreamOptions<UnzipProgress>;

/**
 * Operation handle for streaming unzip.
 */
export type UnzipOperation = ArchiveOperationBase<UnzipProgress> & {
  /** Async iterable of unzip entry objects */
  iterable: AsyncIterable<UnzipEntry>;
};
