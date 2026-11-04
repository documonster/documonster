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
  WorkerPoolStream,
  TaskOptions,
  TaskResult,
  WorkerTaskType
} from "@archive/compression/worker-pool/types";
import type {
  WorkerPoolOptions,
  WorkerPoolStats,
  WorkerPoolStream,
  TaskOptions,
  TaskResult,
  WorkerTaskType
} from "@archive/compression/worker-pool/types";

export { hasWorkerSupport, createAbortError } from "@archive/compression/worker-pool/pool.base";

import { ArchiveError } from "@archive/core/errors";

/** Shared error for Node.js stub methods */
const NODEJS_STUB_ERROR = "WorkerPool is not available in Node.js";

function throwStubError(): never {
  throw new ArchiveError(NODEJS_STUB_ERROR);
}

function rejectStub(): Promise<never> {
  return Promise.reject(new ArchiveError(NODEJS_STUB_ERROR));
}

/**
 * Node.js stub with the browser class's signatures, so code that checks
 * `hasWorkerSupport()` first type-checks on both platforms; every call throws (the
 * constructor included) except `terminate`, a no-op.
 */
export class WorkerPool {
  constructor(options?: WorkerPoolOptions) {
    void options;
    throw new ArchiveError(
      "WorkerPool is only available in browser environments. " +
        "Node.js uses the native zlib thread pool automatically."
    );
  }

  execute(
    taskType: WorkerTaskType,
    data: Uint8Array,
    options?: TaskOptions & { level?: number; maxOutputLength?: number }
  ): Promise<TaskResult> {
    void taskType;
    void data;
    void options;
    return rejectStub();
  }

  getStats(): WorkerPoolStats {
    throwStubError();
  }

  terminate(): void {
    // No-op
  }

  isTerminated(): boolean {
    return true;
  }

  openStream(
    taskType: WorkerTaskType,
    options?: {
      level?: number;
      allowTransfer?: boolean;
      onData: (chunk: Uint8Array) => void;
      onEnd: () => void;
      onError: (err: Error) => void;
    }
  ): WorkerPoolStream {
    void taskType;
    void options;
    throwStubError();
  }

  executeBatch(
    tasks: Array<{
      taskType: WorkerTaskType;
      data: Uint8Array;
      options?: TaskOptions & { level?: number };
    }>
  ): Promise<TaskResult[]> {
    void tasks;
    return rejectStub();
  }
}

export function getDefaultWorkerPool(options?: WorkerPoolOptions): WorkerPool {
  void options;
  throwStubError();
}

export function terminateDefaultWorkerPool(): void {
  // No-op
}

export async function deflateWithPool(
  data: Uint8Array,
  options?: TaskOptions & { level?: number }
): Promise<Uint8Array> {
  void data;
  void options;
  return rejectStub();
}

export async function inflateWithPool(
  data: Uint8Array,
  options?: TaskOptions & { maxOutputLength?: number }
): Promise<Uint8Array> {
  void data;
  void options;
  return rejectStub();
}

export async function deflateBatchWithPool(
  items: Array<{ data: Uint8Array; options?: TaskOptions & { level?: number } }>
): Promise<Uint8Array[]> {
  void items;
  return rejectStub();
}

export async function inflateBatchWithPool(
  items: Array<{ data: Uint8Array; options?: TaskOptions }>
): Promise<Uint8Array[]> {
  void items;
  return rejectStub();
}
