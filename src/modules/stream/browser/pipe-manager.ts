/**
 * Browser Stream - PipeManager
 *
 * Encapsulates pipe/unpipe logic for Readable streams.
 * Owns the destination list and the per-destination listener bookkeeping,
 * so Readable doesn't need to know anything about piping internals.
 */

import { removeEmitterListener } from "@stream/browser/helpers";
import { StreamTypeError } from "@stream/errors";
import type { WritableLike } from "@stream/types";

// =============================================================================
// Types
// =============================================================================

/** Minimal subset of Readable that PipeManager needs to drive. */
export interface PipeSource {
  on(event: string | symbol, listener: (...args: any[]) => void): any;
  once(event: string | symbol, listener: (...args: any[]) => void): any;
  off(event: string | symbol, listener: (...args: any[]) => void): any;
  pause(): any;
  resume(): any;
}

interface PipeListeners<T> {
  data: (chunk: T) => void;
  end?: () => void;
  drain?: () => void;
  destClose?: () => void;
  destFinish?: () => void;
  eventTarget: any;
}

// =============================================================================
// PipeManager
// =============================================================================

export class PipeManager<T> {
  private _destinations: WritableLike[] = [];
  private _listeners: Map<WritableLike, PipeListeners<T>> = new Map();
  private _sourceOverride: PipeSource | null = null;

  constructor(private readonly _source: PipeSource) {}

  /**
   * Override the source identity used when emitting "pipe" and "unpipe" events
   * on destinations. This allows Transform/Duplex wrappers to present
   * themselves (rather than their internal Readable) as the pipe source.
   */
  setSource(source: PipeSource): void {
    this._sourceOverride = source;
  }

  /** The identity to emit in "pipe"/"unpipe" events. */
  private get _emitSource(): PipeSource {
    return this._sourceOverride ?? this._source;
  }

  /** Pipe source data to `destination`. Returns `destination` for chaining. */
  pipe<W extends WritableLike>(destination: W, options?: { end?: boolean }): W {
    // IMPORTANT:
    // Do not rely on `instanceof` here.
    // In bundled/minified builds, multiple copies of this module can exist,
    // causing `instanceof Transform/Writable/Duplex` to fail even when the object
    // is a valid destination.
    const dest = destination;
    const eventTarget: any = dest;

    const hasWrite = typeof dest?.write === "function";
    const hasEnd = typeof dest?.end === "function";
    const hasOn = typeof eventTarget?.on === "function";
    const hasOnce = typeof eventTarget?.once === "function";
    const hasOff = typeof eventTarget?.off === "function";

    if (!hasWrite || !hasEnd || (!hasOnce && !hasOn) || (!hasOff && !eventTarget?.removeListener)) {
      throw new StreamTypeError("Writable", typeof dest);
    }

    this._destinations.push(dest);

    // Create listeners that we can later remove
    let drainListener: (() => void) | undefined;

    const removeDrainListener = (): void => {
      if (!drainListener) {
        return;
      }
      removeEmitterListener(eventTarget, "drain", drainListener);
      drainListener = undefined;
    };

    const dataListener = (chunk: T): void => {
      const canWrite = dest.write(chunk);
      if (!canWrite) {
        this._source.pause();

        if (!drainListener) {
          drainListener = () => {
            removeDrainListener();
            this._source.resume();
          };
          eventTarget.on("drain", drainListener);
          const entry = this._listeners.get(dest);
          if (entry) {
            entry.drain = drainListener;
          }
        }
      }
    };

    const endEnabled = options?.end !== false;

    const endListener = endEnabled
      ? (): void => {
          dest.end();
        }
      : undefined;

    // Auto-unpipe when destination closes or finishes (Node.js compatibility).
    // Node.js internally listens for 'finish' and 'close' on the destination
    // and calls unpipe() so the source stops pushing data.
    const onDestCleanup = (): void => {
      this.unpipe(dest);
    };

    // Use once() if available, otherwise fall back to on() + manual removal
    const onceFn = typeof eventTarget.once === "function" ? "once" : "on";
    eventTarget[onceFn]("close", onDestCleanup);
    eventTarget[onceFn]("finish", onDestCleanup);

    this._listeners.set(dest, {
      data: dataListener,
      end: endListener,
      destClose: onDestCleanup,
      destFinish: onDestCleanup,
      eventTarget
    });

    this._source.on("data", dataListener);
    if (endListener) {
      this._source.once("end", endListener);
    }
    // Node.js pipe() does NOT forward errors from source to destination.
    // Users must handle errors on each stream independently.

    // Emit 'pipe' event on destination (Node.js compatibility)
    eventTarget.emit?.("pipe", this._emitSource);

    this._source.resume();
    return destination;
  }

  /** Unpipe from a specific destination, or all destinations if none given. */
  unpipe(destination?: WritableLike): void {
    if (destination) {
      const idx = this._destinations.indexOf(destination);
      if (idx !== -1) {
        this._destinations.splice(idx, 1);
        this._removeListeners(destination);

        // Pause source when no destinations remain (match Node.js behavior)
        if (this._destinations.length === 0) {
          this._source.pause();
        }
      }
    } else {
      const hadDestinations = this._destinations.length > 0;
      for (const target of this._destinations) {
        this._removeListeners(target);
      }
      this._destinations = [];

      // Only pause if we actually had destinations to remove
      if (hadDestinations) {
        this._source.pause();
      }
    }
  }

  private _removeListeners(destination: WritableLike): void {
    const listeners = this._listeners.get(destination);
    if (!listeners) {
      return;
    }

    this._source.off("data", listeners.data);
    if (listeners.end) {
      this._source.off("end", listeners.end);
    }

    if (listeners.drain) {
      removeEmitterListener(listeners.eventTarget, "drain", listeners.drain);
    }

    // Remove destination close/finish listeners
    if (listeners.destClose) {
      removeEmitterListener(listeners.eventTarget, "close", listeners.destClose);
    }
    if (listeners.destFinish) {
      removeEmitterListener(listeners.eventTarget, "finish", listeners.destFinish);
    }

    // Emit 'unpipe' event on destination (Node.js compatibility)
    listeners.eventTarget.emit?.("unpipe", this._emitSource);

    this._listeners.delete(destination);
  }
}
