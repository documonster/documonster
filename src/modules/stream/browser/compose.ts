/**
 * Browser Stream - Compose
 *
 * Compose multiple transform streams into one.
 * Aligned with Node.js compose semantics:
 * - Backpressure: pauses `last` when composed buffer is full
 * - Flush: waits for `last` to emit "end" before signalling completion
 * - Error: destroys composed stream on any child error (not just emit)
 *
 * Uses Transform constructor options (write, final, destroy) instead of
 * post-construction method overrides, keeping all internal Writable state
 * (buffering, serialization, _writableLength, drain) properly tracked.
 */

import { createListenerRegistry } from "@stream/browser/helpers";
import { Transform } from "@stream/browser/transform";
import { getDefaultHighWaterMark } from "@stream/common/utils";
import type { ITransform } from "@stream/types";

// =============================================================================
// Compose
// =============================================================================

/**
 * Compose multiple transform streams into one
 * Data flows through each transform in sequence
 */
export function compose<T = any, R = any>(
  ...transforms: Array<ITransform<any, any>>
): ITransform<T, R> {
  const len = transforms.length;

  if (len === 0) {
    return new Transform<T, R>({
      objectMode: true,
      transform: chunk => chunk as any as R
    });
  }

  // Preserve identity: compose(single) returns the same transform.
  if (len === 1) {
    return transforms[0] as any as Transform<T, R>;
  }

  // Chain the transforms: first → second → ... → last
  const first = transforms[0]!;
  const last = transforms[len - 1]!;

  // Pipe all transforms together
  for (let i = 0; i < len - 1; i++) {
    transforms[i].pipe(transforms[i + 1]);
  }

  // Track whether last is paused due to backpressure from composed.
  let lastPaused = false;

  // Track whether flush is handling the end sequence.
  let flushing = false;

  // Use per-side objectMode matching Node.js compose behavior.
  // When the property is missing, default to false (same as Node.js Transform).
  const readableObjMode = (last as any).readableObjectMode ?? false;
  const writableObjMode = (first as any).writableObjectMode ?? false;

  const registry = createListenerRegistry();

  const composed = new Transform<T, R>({
    readableHighWaterMark: getDefaultHighWaterMark(readableObjMode),
    writableHighWaterMark: getDefaultHighWaterMark(writableObjMode),
    readableObjectMode: readableObjMode,
    writableObjectMode: writableObjMode,

    // Write path: forward writes into the head of the chain.
    // Using the `write` constructor option ensures all writes go through the
    // Transform's internal Writable state machine (_writableLength tracking,
    // write serialization via _writeQueue, cork buffering, drain signals).
    write(chunk: T, encoding: string, callback: (error?: Error | null) => void) {
      try {
        (first as any).write(chunk, encoding, callback);
      } catch (err) {
        callback(err as Error);
      }
    },

    // Flush path: end the head of the chain, then wait for `last` to emit
    // "end" (readable exhaustion) before completing the composed stream.
    // This is the `final` handler — called after all pending writes complete,
    // before "finish" is emitted. The Transform constructor wrapper will
    // automatically push(null) on the readable side after our callback.
    final(callback: (error?: Error | null) => void) {
      flushing = true;

      // If `last` already ended independently (e.g. a "take N" transform that
      // pushed null on its own), complete immediately — the "end" event has
      // already fired and a new once("end") listener would never trigger.
      if ((last as any).readableEnded) {
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
        (last as any).off?.("end", onEnd);
        (last as any).off?.("error", onError);
      };

      (last as any).once?.("end", onEnd);
      (last as any).once?.("error", onError);
      (first as any).end();
    },

    // Destroy path: clean up all listeners and destroy all child transforms.
    destroy(_error: Error | null, callback: (error?: Error | null) => void) {
      try {
        registry.cleanup();
        for (const t of transforms) {
          t.destroy(_error ?? undefined);
        }
      } finally {
        callback(_error);
      }
    }
  });

  // Hook into the internal _readable's _read method so that when the
  // PipeManager (or any consumer) pulls data, we resume `last` if it was
  // paused due to backpressure. This is the browser equivalent of Node.js
  // compose's `read()` option which is called by the native pipe mechanism.
  //
  // Compose the new behavior with the Transform's existing _read patch
  // (which releases _afterTransformCallback). Although _afterTransformCallback
  // is not used when the `write` option is provided, preserving it is a
  // defensive measure against future changes.
  const composedReadable = (composed as any)._readable;
  const origRead = composedReadable._read;
  composedReadable._read = () => {
    origRead?.call(composedReadable);
    if (lastPaused) {
      lastPaused = false;
      (last as any).resume?.();
    }
  };

  // Forward errors from all transforms — destroy composed on error (matches Node.js).
  for (const t of transforms) {
    registry.add(t as any, "error", (err: Error) => composed.destroy(err));
  }

  // Drain is handled by the internal Writable's own drain → Transform's
  // _setupSideForwarding → composed.emit("drain"). No need to forward from
  // `first` — that would cause double-drain when both buffers cross below HWM.

  // Track whether both sides have completed so we can auto-destroy the composed
  // stream (emitting 'close'), matching Node.js compose behavior.
  let composedEndFired = false;
  let composedFinishFired = false;
  const maybeAutoDestroy = (): void => {
    if (composedEndFired && composedFinishFired && !composed.destroyed) {
      composed.destroy();
    }
  };

  // Finish is handled by the internal Writable's own lifecycle: when the
  // `final` handler calls callback(), the internal Writable emits "finish",
  // which _setupSideForwarding propagates to composed.emit("finish").
  // Forwarding finish from `last` would cause double-finish.
  // Track the composed stream's own finish for auto-destroy.
  composed.once("finish", () => {
    composedFinishFired = true;
    maybeAutoDestroy();
  });

  // Eagerly attach data/end forwarding from `last` to composed (matching Node.js).
  // Node.js compose immediately attaches last.on("data") so data flows into
  // composed's buffer right away, ensuring no data is missed.
  registry.add(last as any, "data", (chunk: R) => {
    if (!(composed as any).push(chunk)) {
      lastPaused = true;
      (last as any).pause?.();
    }
  });

  registry.once(last as any, "end", () => {
    // When flushing, the final handler + Transform wrapper handles stream
    // termination (push(null)). Otherwise (e.g. last ended independently),
    // we must push(null) ourselves.
    if (!flushing) {
      (composed as any).push(null);
    }
  });

  // Track when the composed stream's own 'end' fires (from push(null)).
  composed.once("end", () => {
    composedEndFired = true;
    maybeAutoDestroy();
  });

  // Delegate cork/uncork to the head of the chain.
  (composed as any).cork = (): void => {
    (first as any).cork?.();
  };
  (composed as any).uncork = (): void => {
    (first as any).uncork?.();
  };

  // Safety net: ensure listener cleanup even if close fires through an
  // unexpected path (destroy option already calls registry.cleanup).
  composed.once("close", () => {
    registry.cleanup();
  });

  // Reflect underlying readability/writability like the previous duck-typed wrapper
  Object.defineProperty(composed, "readable", {
    get: () => last.readable
  });
  Object.defineProperty(composed, "writable", {
    get: () => first.writable
  });

  // Proxy writable-side state to `first` so properties like writableEnded and
  // writableFinished reflect the actual head-of-chain state, not the inner
  // Transform wrapper which is never written to directly.
  Object.defineProperty(composed, "writableEnded", {
    get: () => (first as any).writableEnded ?? false
  });
  Object.defineProperty(composed, "writableFinished", {
    get: () => (first as any).writableFinished ?? false
  });
  Object.defineProperty(composed, "writableLength", {
    get: () => (first as any).writableLength ?? 0
  });
  Object.defineProperty(composed, "writableHighWaterMark", {
    get: () => (first as any).writableHighWaterMark ?? getDefaultHighWaterMark(false)
  });
  Object.defineProperty(composed, "writableCorked", {
    get: () => (first as any).writableCorked ?? 0
  });
  Object.defineProperty(composed, "writableNeedDrain", {
    get: () => (first as any).writableNeedDrain ?? false
  });

  // Proxy readable-side state to `last`.
  Object.defineProperty(composed, "readableEnded", {
    get: () => (last as any).readableEnded ?? false
  });
  Object.defineProperty(composed, "readableLength", {
    get: () => (last as any).readableLength ?? 0
  });
  Object.defineProperty(composed, "readableHighWaterMark", {
    get: () => (last as any).readableHighWaterMark ?? getDefaultHighWaterMark(false)
  });
  Object.defineProperty(composed, "readableFlowing", {
    get: () => (last as any).readableFlowing ?? null
  });

  return composed;
}
