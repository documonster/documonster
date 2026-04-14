import type {
  ArchiveProgressPhase,
  ArchiveStreamOptions,
  ArchiveOperationBase
} from "@archive/shared/progress";
import type { Zip64Mode } from "@archive/zip-spec/zip-records";

/**
 * Progress phase for ZIP creation operations.
 */
export type ZipProgressPhase = ArchiveProgressPhase;

/**
 * Progress information for ZIP creation.
 */
export type ZipProgress = {
  type: "zip";
  phase: ZipProgressPhase;

  /** Total number of entries known at start. */
  entriesTotal: number;
  /** Entries fully finalized (data descriptor emitted). */
  entriesDone: number;

  currentEntry?: {
    name: string;
    index: number;
    bytesIn: number;
  };

  /** Total uncompressed bytes consumed from all sources. */
  bytesIn: number;
  /** Total ZIP bytes emitted to the consumer. */
  bytesOut: number;

  /** Zip64 mode in effect for the archive. */
  zip64: Zip64Mode;
};

/**
 * Streaming options for ZIP creation.
 */
export type ZipStreamOptions = ArchiveStreamOptions<ZipProgress>;

/**
 * Operation handle for streaming ZIP creation.
 */
export type ZipOperation = ArchiveOperationBase<ZipProgress> & {
  /** Async iterable of ZIP output chunks */
  iterable: AsyncIterable<Uint8Array>;
};
