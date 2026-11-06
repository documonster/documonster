/**
 * `Compression.Workers` — the browser Web Worker pool for DEFLATE.
 *
 * `Compression.Workers.deflate(bytes)` / `inflate(bytes)` and their `…Batch` forms run on
 * the shared default pool; `new Compression.Workers.Pool(options)` builds a dedicated
 * one. In Node.js zlib already runs on libuv's thread pool, so every member here throws (or
 * rejects with) an `ArchiveWorkerUnsupportedError` except `terminateDefaultPool`, a no-op. They
 * stay published there with the browser signatures so isomorphic code type-checks: guard with
 * `Compression.hasWorkerSupport()`, which answers `false` on the Node entry.
 */
export {
  WorkerPool as Pool,
  getDefaultPool,
  terminateDefaultWorkerPool as terminateDefaultPool,
  deflateWithPool as deflate,
  inflateWithPool as inflate,
  deflateBatchWithPool as deflateBatch,
  inflateBatchWithPool as inflateBatch
} from "@archive/compression/worker-pool";
