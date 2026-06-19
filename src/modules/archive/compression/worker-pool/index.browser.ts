/**
 * Worker Pool Module - Browser Entry
 *
 * Re-exports the browser implementation of the worker pool.
 */

export {
  WorkerPool,
  getDefaultWorkerPool,
  terminateDefaultWorkerPool,
  deflateWithPool,
  inflateWithPool,
  deflateBatchWithPool,
  inflateBatchWithPool,
  hasWorkerSupport,
  type WorkerPoolOptions,
  type WorkerPoolStats,
  type TaskOptions,
  type TaskResult,
  type WorkerTaskType
} from "@archive/compression/worker-pool/pool.browser";
