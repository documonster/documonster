/**
 * Event Emitter
 *
 * Browser-compatible EventEmitter with a Node.js-like API surface.
 * Kept lightweight and allocation-lean for hot paths.
 */

/** Generic event listener signature. */
type EventListener = (...args: any[]) => void;

type ListenerList = EventListener[];
type ListenerValue = EventListener | ListenerList;

function isListenerList(value: ListenerValue): value is ListenerList {
  return Array.isArray(value);
}

export class EventEmitter {
  // Brand for Documonster browser stream objects.
  // Use a string key (not a Symbol) so it still works if the bundle ends up
  // containing multiple copies of this module.
  readonly __documonster_stream: true = true;

  private _listeners: Map<string | symbol, ListenerValue> = new Map();
  private _maxListeners: number = EventEmitter.defaultMaxListeners;

  static defaultMaxListeners: number = 10;

  // addListener is re-assigned to `on` after the class definition to ensure
  // reference identity: `addListener === on` (matches Node.js EventEmitter).
  addListener(event: string | symbol, listener: EventListener): this {
    return this.on(event, listener);
  }

  private _listenerCount(value: ListenerValue | undefined): number {
    if (!value) {
      return 0;
    }
    return isListenerList(value) ? value.length : 1;
  }

  private _hasListeners(event: string | symbol): boolean {
    return this._listenerCount(this._listeners.get(event)) > 0;
  }

  on(event: string | symbol, listener: EventListener): this {
    // Node.js emits 'newListener' BEFORE adding the listener, allowing
    // newListener handlers to insert listeners that fire before this one.
    if (event !== "newListener" && this._hasListeners("newListener")) {
      this.emit("newListener", event, listener);
    }

    const existing = this._listeners.get(event);

    // Warn if exceeding max listeners (skip check if maxListeners is 0 = unlimited)
    if (this._maxListeners > 0) {
      const count = this._listenerCount(existing);
      if (count >= this._maxListeners) {
        // Avoid hard dependency on console for bundle/minified builds
        console?.warn?.(
          `MaxListenersExceededWarning: Possible EventEmitter memory leak detected. ` +
            `${count + 1} ${String(event)} listeners added. ` +
            `Use emitter.setMaxListeners() to increase limit`
        );
      }
    }

    if (!existing) {
      this._listeners.set(event, listener);
    } else if (isListenerList(existing)) {
      existing.push(listener);
    } else {
      this._listeners.set(event, [existing, listener]);
    }

    return this;
  }

  prependListener(event: string | symbol, listener: EventListener): this {
    // Node.js emits 'newListener' BEFORE adding the listener.
    if (event !== "newListener" && this._hasListeners("newListener")) {
      this.emit("newListener", event, listener);
    }

    const existing = this._listeners.get(event);
    if (!existing) {
      this._listeners.set(event, listener);
    } else if (isListenerList(existing)) {
      existing.unshift(listener);
    } else {
      this._listeners.set(event, [listener, existing]);
    }

    return this;
  }

  once(event: string | symbol, listener: EventListener): this {
    const onceWrapper = (...args: any[]): void => {
      this.off(event, onceWrapper);
      listener.apply(this, args);
    };
    (onceWrapper as any).listener = listener;
    return this.on(event, onceWrapper);
  }

  prependOnceListener(event: string | symbol, listener: EventListener): this {
    const onceWrapper = (...args: any[]): void => {
      this.off(event, onceWrapper);
      listener.apply(this, args);
    };
    (onceWrapper as any).listener = listener;
    return this.prependListener(event, onceWrapper);
  }

  // removeListener is re-assigned to `off` after the class definition to ensure
  // reference identity: `removeListener === off` (matches Node.js EventEmitter).
  removeListener(event: string | symbol, listener: EventListener): this {
    return this.off(event, listener);
  }

  off(event: string | symbol, listener: EventListener): this {
    const existing = this._listeners.get(event);
    if (!existing) {
      return this;
    }

    if (!isListenerList(existing)) {
      if (existing === listener || (existing as any).listener === listener) {
        this._listeners.delete(event);
        if (event !== "removeListener" && this._hasListeners("removeListener")) {
          this.emit("removeListener", event, listener);
        }
      }
      return this;
    }

    const listeners = existing;
    if (listeners.length === 0) {
      this._listeners.delete(event);
      return this;
    }

    // Node.js removes the LAST (most recently added) matching listener,
    // not the first. Use lastIndexOf for direct match, reverse search for once-wrapper.
    const directIdx = listeners.lastIndexOf(listener);
    if (directIdx !== -1) {
      listeners.splice(directIdx, 1);
    } else {
      // Slow path: check for once wrapper (search from end for most-recent)
      for (let i = listeners.length - 1; i >= 0; i--) {
        if ((listeners[i] as any).listener === listener) {
          listeners.splice(i, 1);
          break;
        }
      }
    }

    if (listeners.length === 0) {
      this._listeners.delete(event);
    } else if (listeners.length === 1) {
      this._listeners.set(event, listeners[0]);
    }

    if (event !== "removeListener" && this._hasListeners("removeListener")) {
      this.emit("removeListener", event, listener);
    }

    return this;
  }

  emit(event: string | symbol, ...args: any[]): boolean {
    const existing = this._listeners.get(event);
    if (!existing) {
      // Node.js throws when "error" is emitted with no listener
      if (event === "error") {
        const err = args[0];
        if (err instanceof Error) {
          throw err;
        }
        const message = `Unhandled error.${err !== undefined ? ` (${err})` : " (undefined)"}`;
        throw new Error(message);
      }
      return false;
    }

    if (!isListenerList(existing)) {
      // Node.js does NOT catch errors from listeners — they propagate to the caller.
      existing.apply(this, args);
      return true;
    }

    const listeners = existing;
    const len = listeners.length;
    if (len === 0) {
      // Node.js throws when "error" is emitted with no listener
      if (event === "error") {
        const err = args[0];
        if (err instanceof Error) {
          throw err;
        }
        const message = `Unhandled error.${err !== undefined ? ` (${err})` : " (undefined)"}`;
        throw new Error(message);
      }
      return false;
    }

    if (len === 1) {
      listeners[0].apply(this, args);
      return true;
    }

    // Snapshot to allow removal during emit (matches Node.js behavior).
    // This is necessary because a listener may remove other listeners
    // (e.g. `once` removes itself before invoking the wrapped listener).
    const snapshot = listeners.slice();
    for (let i = 0; i < snapshot.length; i++) {
      snapshot[i].apply(this, args);
    }
    return true;
  }

  removeAllListeners(event?: string | symbol): this {
    if (event !== undefined) {
      // Node.js emits 'removeListener' for each removed listener (if 'removeListener' has listeners).
      const hasRemoveListener = event !== "removeListener" && this._hasListeners("removeListener");
      if (hasRemoveListener) {
        const value = this._listeners.get(event);
        if (value) {
          const listeners = isListenerList(value) ? value.slice() : [value];
          this._listeners.delete(event);
          for (const listener of listeners) {
            // Unwrap once-wrappers to emit the original listener
            const original = (listener as any).listener ?? listener;
            this.emit("removeListener", event, original);
          }
          return this;
        }
      }
      this._listeners.delete(event);
    } else {
      // Removing ALL events — Node.js emits 'removeListener' for every listener
      // on every event (except 'removeListener' itself). We process 'removeListener'
      // last to avoid issues while emitting.
      const hasRemoveListener = this._hasListeners("removeListener");
      if (hasRemoveListener) {
        const events = [...this._listeners.keys()];
        for (const evt of events) {
          if (evt !== "removeListener") {
            const value = this._listeners.get(evt);
            if (value) {
              const listeners = isListenerList(value) ? value.slice() : [value];
              this._listeners.delete(evt);
              for (const listener of listeners) {
                const original = (listener as any).listener ?? listener;
                this.emit("removeListener", evt, original);
              }
            }
          }
        }
      }
      this._listeners.clear();
    }
    return this;
  }

  listenerCount(event: string | symbol): number {
    return this._listenerCount(this._listeners.get(event));
  }

  listeners(event: string | symbol): EventListener[] {
    const value = this._listeners.get(event);
    if (!value) {
      return [];
    }
    // Node.js: listeners() returns the ORIGINAL listener functions,
    // unwrapping once-wrappers. rawListeners() returns the wrappers.
    const raw = isListenerList(value) ? value : [value];
    return raw.map(fn => (fn as any).listener ?? fn);
  }

  rawListeners(event: string | symbol): EventListener[] {
    const value = this._listeners.get(event);
    if (!value) {
      return [];
    }
    // rawListeners returns the actual wrapper functions (including once-wrappers)
    return isListenerList(value) ? value.slice() : [value];
  }

  eventNames(): (string | symbol)[] {
    return [...this._listeners.keys()];
  }

  setMaxListeners(n: number): this {
    this._maxListeners = n;
    return this;
  }

  getMaxListeners(): number {
    return this._maxListeners;
  }
}

// Node.js guarantees `addListener === on` and `removeListener === off` on the
// EventEmitter prototype.  We cannot achieve reference identity inside the
// class body (TypeScript generates separate method slots), so we patch the
// prototype immediately after the class is defined.
EventEmitter.prototype.addListener = EventEmitter.prototype.on;
EventEmitter.prototype.removeListener = EventEmitter.prototype.off;
