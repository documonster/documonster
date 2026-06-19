/**
 * Unified progress types and utilities for archive operations.
 *
 * This module defines common progress structures used across zip and unzip operations.
 * Having them in a shared location reduces duplication and enables cross-module type reuse.
 *
 * @module
 */

// -----------------------------------------------------------------------------
// Common Types
// -----------------------------------------------------------------------------

/**
 * Common progress phase states for archive operations.
 */
export type ArchiveProgressPhase = "running" | "done" | "aborted" | "error";

// -----------------------------------------------------------------------------
// Base Stream Options
// -----------------------------------------------------------------------------

/**
 * Common streaming options shared between zip and unzip.
 */
export interface ArchiveStreamOptions<P> {
  /** Abort signal for cancellation */
  signal?: AbortSignal;

  /** Progress callback */
  onProgress?: (p: P) => void;

  /** Throttle progress callbacks; 0 emits on the next microtask */
  progressIntervalMs?: number;
}

/**
 * Base operation result type.
 */
export interface ArchiveOperationBase<P> {
  /** Abort signal linked to this operation */
  signal: AbortSignal;

  /** Abort the operation */
  abort(reason?: unknown): void;

  /** Returns bytes processed so far */
  pointer(): number;

  /** Latest progress snapshot */
  progress(): P;
}

// -----------------------------------------------------------------------------
// Progress Emitter
// -----------------------------------------------------------------------------

export type ProgressListener<T> = (snapshot: T) => void;

export type ProgressEmitterOptions = {
  intervalMs?: number;
};

/**
 * Small helper to batch frequent progress updates.
 * - Always sends the latest snapshot
 * - Throttles to at most once per `intervalMs` (default: every microtask)
 */
export class ProgressEmitter<T extends object> {
  private readonly _listener?: ProgressListener<T>;
  private readonly _intervalMs: number;

  private _lastEmitAt = 0;
  private _snapshot: T;
  private _tokenSeq = 0;
  private _scheduledToken = 0;
  private _timeout: ReturnType<typeof setTimeout> | null = null;

  constructor(initial: T, listener?: ProgressListener<T>, options: ProgressEmitterOptions = {}) {
    this._snapshot = initial;
    this._listener = listener;
    this._intervalMs = Math.max(0, Math.floor(options.intervalMs ?? 0));
  }

  get snapshot(): T {
    return this._snapshot;
  }

  snapshotCopy(): T {
    return { ...(this._snapshot as any) };
  }

  update(patch: Partial<T>): void {
    Object.assign(this._snapshot, patch);
    this._schedule();
  }

  set<K extends keyof T>(key: K, value: T[K]): void {
    (this._snapshot as any)[key] = value;
    this._schedule();
  }

  mutate(mutator: (snapshot: T) => void): void {
    mutator(this._snapshot);
    this._schedule();
  }

  emitNow(): void {
    if (!this._listener) {
      return;
    }

    // Cancel any pending scheduled emit (microtask/timeout).
    this._scheduledToken = 0;
    if (this._timeout) {
      clearTimeout(this._timeout);
      this._timeout = null;
    }

    this._lastEmitAt = Date.now();
    this._listener(this.snapshotCopy());
  }

  private _schedule(): void {
    if (!this._listener) {
      return;
    }

    // Already scheduled.
    if (this._scheduledToken !== 0) {
      return;
    }

    if (this._intervalMs === 0) {
      const token = ++this._tokenSeq;
      this._scheduledToken = token;
      queueMicrotask(() => {
        if (this._scheduledToken === token) {
          this.emitNow();
        }
      });
      return;
    }

    const now = Date.now();
    const dueIn = Math.max(0, this._intervalMs - (now - this._lastEmitAt));
    const token = ++this._tokenSeq;
    this._scheduledToken = token;

    if (dueIn === 0) {
      queueMicrotask(() => {
        if (this._scheduledToken === token) {
          this.emitNow();
        }
      });
      return;
    }

    this._timeout = setTimeout(() => {
      if (this._scheduledToken === token) {
        this.emitNow();
      }
      this._timeout = null;
    }, dueIn);
  }
}
