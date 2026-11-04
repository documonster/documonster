/**
 * `Compression.Workers` — the browser Web Worker pool for DEFLATE.
 *
 * `Compression.Workers.deflate(bytes)` / `inflate(bytes)` and their `…Batch` forms run on
 * the shared default pool; `new Compression.Workers.WorkerPool(options)` builds a dedicated
 * one. In Node.js zlib already runs on libuv's thread pool, so every member here throws
 * except `terminateDefaultPool` (a no-op) — check `Compression.hasWorkerSupport()` first.
 */
export {
  WorkerPool,
  getDefaultWorkerPool as getDefaultPool,
  terminateDefaultWorkerPool as terminateDefaultPool,
  deflateWithPool as deflate,
  inflateWithPool as inflate,
  deflateBatchWithPool as deflateBatch,
  inflateBatchWithPool as inflateBatch
} from "@archive/compression/worker-pool";
