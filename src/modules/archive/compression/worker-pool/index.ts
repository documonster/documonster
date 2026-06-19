/**
 * Worker Pool Module - Node Entry (Stub)
 *
 * Node.js doesn't need a worker pool for compression since zlib
 * already uses the libuv thread pool internally.
 *
 * This stub exports a no-op implementation for API parity.
 */

export type {
  WorkerPoolOptions,
  WorkerPoolStats,
  TaskOptions,
  TaskResult,
  WorkerTaskType
} from "./types";

export { hasWorkerSupport, createAbortError } from "./pool.base";

import { ArchiveError } from "@archive/core/errors";

/** Shared error for Node.js stub methods */
const NODEJS_STUB_ERROR = "WorkerPool is not available in Node.js";

function throwStubError(): never {
  throw new ArchiveError(NODEJS_STUB_ERROR);
}

/**
 * Node.js stub - throws if called
 */
export class WorkerPool {
  constructor() {
    throw new ArchiveError(
      "WorkerPool is only available in browser environments. " +
        "Node.js uses the native zlib thread pool automatically."
    );
  }

  execute(): never {
    throwStubError();
  }

  getStats(): never {
    throwStubError();
  }

  terminate(): void {
    // No-op
  }

  isTerminated(): boolean {
    return true;
  }
}

export function getDefaultWorkerPool(): never {
  throwStubError();
}

export function terminateDefaultWorkerPool(): void {
  // No-op
}

export async function deflateWithPool(): Promise<never> {
  throwStubError();
}

export async function inflateWithPool(): Promise<never> {
  throwStubError();
}

export async function deflateBatchWithPool(): Promise<never> {
  throwStubError();
}

export async function inflateBatchWithPool(): Promise<never> {
  throwStubError();
}
