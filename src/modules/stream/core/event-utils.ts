/**
 * Small event utilities for Node-style emitters.
 *
 * Prefer keeping this separate from the main stream implementation so other
 * modules (e.g. archive) can reuse it without pulling in the whole stream API.
 */

import type { EventEmitterLike } from "@stream/types";

function off(emitter: EventEmitterLike, event: string, listener: (...args: any[]) => void): void {
  if (typeof emitter.off === "function") {
    emitter.off(event, listener);
  } else if (typeof emitter.removeListener === "function") {
    emitter.removeListener(event, listener);
  }
}

/**
 * Resolve when an emitter fires `event`, reject on `error`.
 */
export function onceEvent(emitter: EventEmitterLike, event: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const onError = (err: unknown): void => {
      cleanup();
      reject(err instanceof Error ? err : new Error(String(err)));
    };

    const onDone = (): void => {
      cleanup();
      resolve();
    };

    const cleanup = (): void => {
      off(emitter, "error", onError);
      off(emitter, event, onDone);
    };

    if (typeof emitter.once === "function") {
      emitter.once("error", onError);
      emitter.once(event, onDone);
      return;
    }

    emitter.on?.("error", onError);
    emitter.on?.(event, onDone);
  });
}
