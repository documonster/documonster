/**
 * Node.js Stream - Compose
 *
 * Compose multiple transform streams into one.
 */

import { Transform } from "stream";
import type { TransformCallback as NodeTransformCallback } from "stream";

import { getDefaultHighWaterMark } from "@stream/core/utils";
import type { ITransform } from "@stream/types";

// =============================================================================
// Internal chain-node shape
// =============================================================================

/**
 * The subset of stream members `compose` reads/calls on the chained
 * transforms beyond what `ITransform` declares. These reflect Node's
 * `Writable`/`Readable` runtime surface (cork/uncork and the per-side
 * objectMode / corked / needDrain accessors); they are optional because the
 * chain nodes may be browser polyfills that omit some of them.
 */
interface ComposeChainNode extends ITransform<unknown, unknown> {
  readonly writableObjectMode?: boolean;
  readonly writableCorked?: number;
  readonly writableNeedDrain?: boolean;
  cork?(): void;
  uncork?(): void;
}

// =============================================================================
// Compose
// =============================================================================

/**
 * Compose multiple transform streams into one
 */
export function compose<T = unknown, R = unknown>(
  ...transforms: Array<ITransform<unknown, unknown>>
): ITransform<T, R> {
  const len = transforms.length;
  if (len === 0) {
    return new Transform({
      highWaterMark: getDefaultHighWaterMark(true),
      objectMode: true,
      transform(chunk: any, _encoding: BufferEncoding, callback: NodeTransformCallback) {
        callback(null, chunk);
      }
    });
  }

  if (len === 1) {
    // Identity: composing a single transform returns it unchanged. The input
    // element type is erased by the variadic signature, so re-project onto the
    // declared output type.
    return transforms[0] as ITransform<T, R>;
  }

  // Chain all transforms together once.
  for (let i = 0; i < len - 1; i++) {
    transforms[i]!.pipe(transforms[i + 1]!);
  }

  const first: ComposeChainNode = transforms[0]!;
  const last: ComposeChainNode = transforms[len - 1]!;

  // Track whether last is paused due to backpressure from composed.
  let lastPaused = false;

  // Use per-side objectMode matching browser compose behavior.
  // When the property is missing, default to false (same as browser).
  const readableObjMode = last.readableObjectMode ?? false;
  const writableObjMode = first.writableObjectMode ?? false;

  const composed = new Transform({
    readableHighWaterMark: getDefaultHighWaterMark(readableObjMode),
    writableHighWaterMark: getDefaultHighWaterMark(writableObjMode),
    readableObjectMode: readableObjMode,
    writableObjectMode: writableObjMode,
    transform(chunk: any, encoding: BufferEncoding, callback: NodeTransformCallback) {
      try {
        // Forward writes into the head of the chain.
        first.write(chunk, encoding, callback);
      } catch (err) {
        callback(err as Error);
      }
    },
    flush(callback: NodeTransformCallback) {
      flushing = true;
      // End the head of the chain and wait for `last` to finish emitting all
      // data.  We must wait for `last`'s "end" (readable exhaustion) — not
      // `first`'s "finish" (writable flush) — because data may still be
      // flowing through intermediate transforms after `first` finishes.

      // If `last` already ended independently (e.g. a "take N" transform that
      // pushed null on its own), complete immediately — the "end" event has
      // already fired and a new once("end") listener would never trigger.
      if (last.readableEnded) {
        callback();
        return;
      }

      const onEnd = (): void => {
        cleanupFlush();
        callback();
      };
      const onError = (err: Error): void => {
        cleanupFlush();
        callback(err);
      };
      const cleanupFlush = (): void => {
        last.off?.("end", onEnd);
        last.off?.("error", onError);
      };

      last.once?.("end", onEnd);
      last.once?.("error", onError);
      first.end();
    },
    read(this: Transform, size: number) {
      // Resume last if it was paused due to backpressure.
      if (lastPaused) {
        lastPaused = false;
        last.resume?.();
      }
      Transform.prototype._read.call(this, size);
    },
    destroy(this: Transform, err: Error | null, callback: (error: Error | null) => void) {
      try {
        cleanupListeners();
        for (const t of transforms) {
          t.destroy(err ?? undefined);
        }
      } finally {
        callback(err);
      }
    }
  });

  // Forward data from last directly to composed, with backpressure.
  const onLastData = (chunk: any): void => {
    if (!composed.push(chunk)) {
      lastPaused = true;
      last.pause?.();
    }
  };

  // Track whether flush is handling the end sequence.
  let flushing = false;

  const onLastEnd = (): void => {
    cleanupListeners();
    // When flushing, the flush callback handles stream termination.
    // Otherwise (e.g. last ended independently), we must push(null) ourselves.
    if (!flushing) {
      composed.push(null);
    }
  };

  const onAnyError = (err: Error): void => {
    cleanupListeners();
    composed.destroy(err);
  };

  const transformErrorListeners: Array<{ t: ComposeChainNode; fn: (err: Error) => void }> = [];
  const cleanupListeners = (): void => {
    last.off?.("data", onLastData);
    last.off?.("end", onLastEnd);
    last.off?.("error", onAnyError);
    for (const { t, fn } of transformErrorListeners) {
      t.off?.("error", fn);
    }
    transformErrorListeners.length = 0;
  };

  // Drain is handled by composed's own Writable state machine: when the
  // transform callback fires (after first.write's callback), the Writable
  // checks _writableLength < _highWaterMark and emits drain if needed.
  // Forwarding drain from `first` would cause double-drain.

  // Finish is handled by composed's own Writable state machine: when the
  // flush callback fires, the internal Writable emits "finish" natively.
  // Forwarding finish from `last` would cause double-finish.

  // Eagerly attach data/end forwarding from `last` to composed.
  // This ensures data flows into composed's buffer immediately.
  last.on?.("data", onLastData);
  last.once?.("end", onLastEnd);
  last.on?.("error", onAnyError);

  // Forward errors from all transforms (including last, using persistent
  // listeners to match browser compose which uses registry.add).
  for (const t of transforms) {
    if (t === last) {
      continue;
    }
    const tt: ComposeChainNode = t;
    tt.on?.("error", onAnyError);
    transformErrorListeners.push({ t: tt, fn: onAnyError });
  }

  // Delegate cork/uncork to the head of the chain only.
  // The composed Transform's own write() is overridden to forward to `first`,
  // so corking composed itself has no effect — only `first` needs to be corked.
  // writableCorked is already proxied to `first`, keeping the property in sync.
  composed.cork = (): void => {
    first.cork?.();
  };
  composed.uncork = (): void => {
    first.uncork?.();
  };

  composed.once("close", () => {
    cleanupListeners();
  });

  // Proxy readable/writable to reflect underlying chain state (matches browser compose).
  Object.defineProperty(composed, "readable", {
    get: () => last.readable
  });
  Object.defineProperty(composed, "writable", {
    get: () => first.writable
  });

  // Proxy writable-side state to `first` so properties like writableEnded and
  // writableFinished reflect the actual head-of-chain state.
  Object.defineProperty(composed, "writableEnded", {
    get: () => first.writableEnded ?? false
  });
  Object.defineProperty(composed, "writableFinished", {
    get: () => first.writableFinished ?? false
  });
  Object.defineProperty(composed, "writableLength", {
    get: () => first.writableLength ?? 0
  });
  Object.defineProperty(composed, "writableHighWaterMark", {
    get: () => first.writableHighWaterMark ?? getDefaultHighWaterMark(false)
  });
  Object.defineProperty(composed, "writableCorked", {
    get: () => first.writableCorked ?? 0
  });
  Object.defineProperty(composed, "writableNeedDrain", {
    get: () => first.writableNeedDrain ?? false
  });

  // Proxy readable-side state to `last`.
  Object.defineProperty(composed, "readableEnded", {
    get: () => last.readableEnded ?? false
  });
  Object.defineProperty(composed, "readableLength", {
    get: () => last.readableLength ?? 0
  });
  Object.defineProperty(composed, "readableHighWaterMark", {
    get: () => last.readableHighWaterMark ?? getDefaultHighWaterMark(false)
  });
  Object.defineProperty(composed, "readableFlowing", {
    get: () => last.readableFlowing ?? null
  });

  return composed;
}
