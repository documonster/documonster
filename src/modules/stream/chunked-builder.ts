/**
 * Chunked Builder (Browser Compatible)
 *
 * Browser-compatible version that uses Uint8Array instead of Buffer.
 * This file provides the same API but works in browser environments.
 */

import { StreamStateError } from "@stream/errors";
import { textEncoder } from "@utils/binary";

/**
 * Options for ChunkedBuilder
 */
export interface ChunkedBuilderOptions {
  /** Number of pieces to accumulate before consolidation */
  chunkSize?: number;
  /** Initial capacity hint */
  initialCapacity?: number;
}

/**
 * Efficient builder for accumulating and consolidating string chunks
 * Browser-compatible version
 */
export class ChunkedBuilder {
  protected _pieces: string[] = [];
  protected _chunks: string[] = [];
  protected readonly _chunkSize: number;
  protected _totalLength: number = 0;

  constructor(options: ChunkedBuilderOptions = {}) {
    this._chunkSize = options.chunkSize ?? 10000;
  }

  /**
   * Add a piece to the builder
   */
  push(piece: string): void {
    this._pieces.push(piece);
    this._totalLength += piece.length;
    if (this._pieces.length >= this._chunkSize) {
      this._consolidate();
    }
  }

  /**
   * Add multiple pieces
   */
  pushAll(pieces: string[]): void {
    const len = pieces.length;
    for (let i = 0; i < len; i++) {
      this._pieces.push(pieces[i]);
      this._totalLength += pieces[i].length;
    }
    if (this._pieces.length >= this._chunkSize) {
      this._consolidate();
    }
  }

  /**
   * Consolidate pieces into chunks.
   * Subclasses may override to guard against consolidation (e.g. during active snapshots).
   */
  protected _consolidate(): void {
    if (this._pieces.length > 0) {
      this._chunks.push(this._pieces.join(""));
      this._pieces.length = 0;
    }
  }

  /**
   * Get current cursor position (useful for tracking changes)
   */
  get cursor(): number {
    return this._chunks.length * this._chunkSize + this._pieces.length;
  }

  /**
   * Get total piece/chunk count
   */
  get length(): number {
    return this._pieces.length + this._chunks.length;
  }

  /**
   * Get total string length (character count)
   */
  get stringLength(): number {
    return this._totalLength;
  }

  /**
   * Check if empty
   */
  get isEmpty(): boolean {
    return this._pieces.length === 0 && this._chunks.length === 0;
  }

  /**
   * Clear all content
   */
  clear(): void {
    this._pieces.length = 0;
    this._chunks.length = 0;
    this._totalLength = 0;
  }

  /**
   * Build final string
   */
  toString(): string {
    const chunksLen = this._chunks.length;
    const piecesLen = this._pieces.length;

    // Fast path: only pieces, no chunks
    if (chunksLen === 0) {
      if (piecesLen === 0) {
        return "";
      }
      if (piecesLen === 1) {
        return this._pieces[0];
      }
      return this._pieces.join("");
    }

    // Has chunks - consolidate and join
    if (piecesLen > 0) {
      this._chunks.push(this._pieces.join(""));
      this._pieces.length = 0;
    }

    if (this._chunks.length === 1) {
      return this._chunks[0];
    }

    return this._chunks.join("");
  }

  /**
   * Convert to Uint8Array (browser-compatible)
   */
  toUint8Array(): Uint8Array {
    return textEncoder.encode(this.toString());
  }
}

/**
 * Snapshot for rollback support
 */
export interface BuilderSnapshot {
  piecesLength: number;
  chunksLength: number;
  totalLength: number;
}

/**
 * Chunked builder with rollback/commit support
 * Browser-compatible version
 */
export class TransactionalChunkedBuilder extends ChunkedBuilder {
  private _snapshots: BuilderSnapshot[] = [];

  /**
   * Skip consolidation while snapshots are active.
   * Consolidation joins pieces into a chunk and clears the pieces array,
   * which makes it impossible to rollback to a previous pieces position.
   */
  protected override _consolidate(): void {
    if (this._snapshots.length > 0) {
      return;
    }
    super._consolidate();
  }

  /**
   * Create a rollback point
   */
  snapshot(): BuilderSnapshot {
    const snap: BuilderSnapshot = {
      piecesLength: this._pieces.length,
      chunksLength: this._chunks.length,
      totalLength: this._totalLength
    };
    this._snapshots.push(snap);
    return snap;
  }

  /**
   * Commit the current snapshot (remove rollback point)
   */
  commit(): void {
    this._snapshots.pop();
  }

  /**
   * Rollback to the last snapshot
   */
  rollback(): void {
    const snap = this._snapshots.pop();
    if (!snap) {
      throw new StreamStateError("rollback", "no snapshot available");
    }

    if (this._pieces.length > snap.piecesLength) {
      this._pieces.length = snap.piecesLength;
    }

    if (this._chunks.length > snap.chunksLength) {
      this._chunks.length = snap.chunksLength;
    }

    this._totalLength = snap.totalLength;
  }

  /**
   * Check if there are active snapshots
   */
  get hasSnapshots(): boolean {
    return this._snapshots.length > 0;
  }
}
