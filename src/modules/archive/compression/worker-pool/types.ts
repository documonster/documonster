/**
 * Worker Pool Types
 *
 * Shared type definitions for the worker pool system.
 * Used by both the main thread and worker scripts.
 */

/**
 * Task types that workers can handle
 */
export type WorkerTaskType = "deflate" | "inflate";

/**
 * Worker message types from main thread to worker
 */
export type WorkerRequestMessage =
  | {
      /** Batch task request */
      type: "task";
      /** Unique task ID for correlation */
      taskId: number;
      /** Task type (deflate or inflate) */
      taskType: WorkerTaskType;
      /** Input data to process */
      data: Uint8Array;
      /** Compression level (for deflate) */
      level?: number;
      /** Output bound (for inflate); exceeding it fails with `limit` set */
      maxOutputLength?: number;
    }
  | {
      /** Streaming task start */
      type: "start";
      taskId: number;
      taskType: WorkerTaskType;
      level?: number;
    }
  | {
      /** Streaming data chunk */
      type: "chunk";
      taskId: number;
      data: Uint8Array;
    }
  | {
      /** Streaming end (flush + close) */
      type: "end";
      taskId: number;
    }
  | {
      /** Streaming abort/cancel */
      type: "abort";
      taskId: number;
      error?: string;
    }
  | {
      /** Terminate worker */
      type: "terminate";
    };

/**
 * Worker message types from worker to main thread
 */
export type WorkerResponseMessage =
  | {
      /** Worker is ready */
      type: "ready";
    }
  | {
      /** Streaming task started successfully (start handshake) */
      type: "started";
      taskId: number;
    }
  | {
      /** Batch task result */
      type: "result";
      taskId: number;
      data: Uint8Array;
      duration?: number;
    }
  | {
      /** Streaming output chunk */
      type: "out";
      taskId: number;
      data: Uint8Array;
    }
  | {
      /** Acknowledge that an input chunk was consumed */
      type: "ack";
      taskId: number;
    }
  | {
      /** Streaming task completed */
      type: "done";
      taskId: number;
      duration?: number;
    }
  | {
      /** Any task error (batch or streaming) */
      type: "error";
      taskId: number;
      error?: string;
      /** Set when the task failed because output exceeded this `maxOutputLength` */
      limit?: number;
      duration?: number;
    };

/**
 * Worker pool configuration options
 */
export interface WorkerPoolOptions {
  /**
   * Maximum number of workers in the pool.
   * Defaults to navigator.hardwareConcurrency or 4.
   */
  maxWorkers?: number;

  /**
   * Minimum number of workers to keep alive (warm pool).
   * Workers beyond this count may be terminated when idle.
   * Defaults to 0.
   */
  minWorkers?: number;

  /**
   * Time in milliseconds before an idle worker is terminated.
   * Only applies to workers beyond minWorkers.
   * Defaults to 30000 (30 seconds).
   */
  idleTimeout?: number;

  /**
   * Custom worker URL. If provided, uses this instead of inline worker.
   */
  workerUrl?: string;

  /**
   * Whether to use transferable objects for zero-copy data transfer.
   * Defaults to true.
   */
  useTransferables?: boolean;
}

/**
 * Worker pool statistics
 */
export interface WorkerPoolStats {
  /** Total number of workers (active + idle) */
  totalWorkers: number;
  /** Number of busy workers */
  activeWorkers: number;
  /** Number of idle workers */
  idleWorkers: number;
  /** Number of pending tasks in queue */
  pendingTasks: number;
  /** Total tasks completed since pool creation */
  completedTasks: number;
  /** Total tasks failed since pool creation */
  failedTasks: number;
}

/**
 * Task priority levels
 */
export type TaskPriority = "high" | "normal" | "low";

/**
 * Task configuration options
 */
export interface TaskOptions {
  /**
   * Task priority. Higher priority tasks are processed first.
   * Defaults to "normal".
   */
  priority?: TaskPriority;

  /**
   * Abort signal for task cancellation.
   */
  signal?: AbortSignal;

  /**
   * Allow transferring the input buffer to the worker (zero-copy).
   * When true, the input Uint8Array's underlying buffer will be transferred
   * and become unusable in the main thread.
   *
   * Use this for better performance when you don't need the input data after compression.
   *
   * Defaults to false (data is copied to preserve the original buffer).
   */
  allowTransfer?: boolean;
}

/**
 * Result of a worker task
 */
export interface TaskResult {
  /** Processed output data */
  data: Uint8Array;
  /** Processing time in milliseconds */
  duration: number;
}
