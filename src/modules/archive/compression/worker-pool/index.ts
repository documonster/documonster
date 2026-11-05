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

// The Node answer (always `false`), not the pool's `typeof Worker` probe: Bun has a global
// `Worker`, yet this entry's pool cannot use it.
export { hasWorkerSupport } from "@archive/compression/compress";
export { createAbortError } from "@archive/compression/worker-pool/pool.base";

import { ArchiveWorkerUnsupportedError } from "@archive/core/errors";

function throwStubError(operation: string): never {
  throw new ArchiveWorkerUnsupportedError(operation);
}

function rejectStub(operation: string): Promise<never> {
  return Promise.reject(new ArchiveWorkerUnsupportedError(operation));
}

/**
 * Node.js stub with the browser class's signatures, so code that checks
 * `Compression.hasWorkerSupport()` (always `false` here) first type-checks on both platforms;
 * every call throws {@link ArchiveWorkerUnsupportedError} (the constructor included) except
 * `terminate`, a no-op.
 */
export class WorkerPool {
  constructor(options?: WorkerPoolOptions) {
    void options;
    throw new ArchiveWorkerUnsupportedError("Compression.Workers.WorkerPool");
  }

  execute(
    taskType: WorkerTaskType,
    data: Uint8Array,
    options?: TaskOptions & { level?: number; maxOutputLength?: number }
  ): Promise<TaskResult> {
    void taskType;
    void data;
    void options;
    return rejectStub("Compression.Workers.WorkerPool#execute");
  }

  getStats(): WorkerPoolStats {
    throwStubError("Compression.Workers.WorkerPool#getStats");
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
    throwStubError("Compression.Workers.WorkerPool#openStream");
  }

  executeBatch(
    tasks: Array<{
      taskType: WorkerTaskType;
      data: Uint8Array;
      options?: TaskOptions & { level?: number };
    }>
  ): Promise<TaskResult[]> {
    void tasks;
    return rejectStub("Compression.Workers.WorkerPool#executeBatch");
  }
}

export function getDefaultWorkerPool(options?: WorkerPoolOptions): WorkerPool {
  void options;
  throwStubError("Compression.Workers.getDefaultPool");
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
  return rejectStub("Compression.Workers.deflate");
}

export async function inflateWithPool(
  data: Uint8Array,
  options?: TaskOptions & { maxOutputLength?: number }
): Promise<Uint8Array> {
  void data;
  void options;
  return rejectStub("Compression.Workers.inflate");
}

export async function deflateBatchWithPool(
  items: Array<{ data: Uint8Array; options?: TaskOptions & { level?: number } }>
): Promise<Uint8Array[]> {
  void items;
  return rejectStub("Compression.Workers.deflateBatch");
}

export async function inflateBatchWithPool(
  items: Array<{ data: Uint8Array; options?: TaskOptions }>
): Promise<Uint8Array[]> {
  void items;
  return rejectStub("Compression.Workers.inflateBatch");
}
