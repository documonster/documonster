/**
 * CSV Worker Pool - Browser Implementation
 *
 * High-performance worker pool for CSV operations in the browser.
 * Offloads CPU-intensive operations to Web Workers to keep UI responsive.
 *
 * Features:
 * - Parse/Format: Basic CSV operations
 * - Session Management: Keep data in worker memory for repeated operations
 * - Data Operations: sort, filter, search, groupBy, aggregate, pagination
 * - Batch Query: Execute multiple operations in single round-trip
 * - Task prioritization and cancellation
 * - Automatic worker scaling
 *
 * @example
 * ```ts
 * // Simple parsing
 * const result = await parseWithPool(csvString, { headers: true });
 *
 * // Session-based operations for interactive data exploration
 * const session = new CsvWorkerSession();
 * await session.load(csvString, { headers: true });
 *
 * // Batch query - single round-trip for multiple operations
 * const result = await session.query({
 *   sort: { column: 'age', order: 'desc' },
 *   filter: { conditions: [{ column: 'status', operator: 'eq', value: 'active' }] },
 *   page: { page: 1, pageSize: 20 }
 * });
 *
 * session.dispose();
 * ```
 */

import { CsvWorkerError } from "@csv/errors";
import type {
  CsvWorkerPoolOptions,
  CsvWorkerPoolStats,
  CsvTaskOptions,
  CsvTaskResult,
  CsvTaskPriority,
  CsvWorkerRequestMessage,
  CsvWorkerResponseMessage,
  SortConfig,
  FilterConfig,
  SearchConfig,
  GroupByConfig,
  AggregateConfig,
  PageConfig,
  QueryConfig,
  FilterResult,
  PageResult,
  GroupResult,
  AggregateResult,
  QueryResult,
  CsvParseOptions,
  CsvFormatOptions,
  CsvParseResult
} from "@csv/worker/types";
import { getWorkerBlobUrl, releaseWorkerBlobUrl } from "@csv/worker/worker-script.bundle";

// =============================================================================
// Constants
// =============================================================================

const DEFAULT_OPTIONS: Required<Omit<CsvWorkerPoolOptions, "workerUrl">> & { workerUrl?: string } =
  {
    maxWorkers: typeof navigator !== "undefined" ? navigator.hardwareConcurrency || 4 : 4,
    minWorkers: 0,
    idleTimeout: 30_000,
    workerUrl: undefined
  };

/** Check if Web Workers are available */
export function hasWorkerSupport(): boolean {
  return typeof Worker !== "undefined" && typeof Blob !== "undefined";
}

function createAbortError(message = "Operation was aborted"): Error {
  const error = new Error(message);
  error.name = "AbortError";
  return error;
}

// =============================================================================
// Internal Types
// =============================================================================

interface PendingTask<T> {
  taskId: number;
  resolve: (result: T) => void;
  reject: (error: Error) => void;
  signal?: AbortSignal;
  abortHandler?: () => void;
  startTime: number;
}

interface QueuedTask extends PendingTask<any> {
  message: CsvWorkerRequestMessage;
  priority: CsvTaskPriority;
}

interface PoolWorker {
  id: number;
  worker: Worker;
  busy: boolean;
  currentTaskId: number | null;
  idleTimer: ReturnType<typeof setTimeout> | null;
}

// =============================================================================
// CsvWorkerPool Class (Internal)
// =============================================================================

class CsvWorkerPool {
  private readonly _options: typeof DEFAULT_OPTIONS;
  private readonly _workers: Map<number, PoolWorker> = new Map();
  private readonly _highQueue: QueuedTask[] = [];
  private readonly _normalQueue: QueuedTask[] = [];
  private readonly _lowQueue: QueuedTask[] = [];
  private readonly _pendingTasks: Map<number, PendingTask<any>> = new Map();
  private _nextTaskId = 1;
  private _nextWorkerId = 1;
  private _terminated = false;
  private _completedTasks = 0;
  private _failedTasks = 0;
  private _workerUrl: string | null = null;
  private readonly _useCustomUrl: boolean;
  private _initPromise: Promise<void> | null = null;

  constructor(options?: CsvWorkerPoolOptions) {
    this._options = { ...DEFAULT_OPTIONS, ...options };

    if (this._options.workerUrl) {
      this._workerUrl = this._options.workerUrl;
      this._useCustomUrl = true;
    } else {
      this._useCustomUrl = false;
    }
  }

  /**
   * Create and initialize a worker pool.
   * This is the recommended way to create a pool for immediate use.
   */
  static async create(options?: CsvWorkerPoolOptions): Promise<CsvWorkerPool> {
    const pool = new CsvWorkerPool(options);
    await pool._ensureInitialized();
    return pool;
  }

  /**
   * Ensure the pool is initialized (worker URL loaded).
   * Called automatically before first task execution.
   */
  private async _ensureInitialized(): Promise<void> {
    if (this._workerUrl) {
      return;
    }
    if (this._initPromise) {
      return this._initPromise;
    }

    this._initPromise = getWorkerBlobUrl()
      .then(url => {
        this._workerUrl = url;
        // Create min workers after URL is ready
        for (let i = 0; i < this._options.minWorkers; i++) {
          this._createWorker();
        }
      })
      .catch(err => {
        // Clear the cached promise so subsequent calls can retry initialization
        this._initPromise = null;
        throw err;
      });

    return this._initPromise;
  }

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------

  async parse(
    data: string,
    options?: CsvParseOptions & { sessionId?: string },
    taskOptions?: CsvTaskOptions
  ): Promise<CsvTaskResult<string[][] | CsvParseResult<Record<string, string>>>> {
    const message: CsvWorkerRequestMessage = {
      type: "parse",
      taskId: 0,
      data,
      options,
      sessionId: options?.sessionId
    };
    return this._execute(message, taskOptions);
  }

  async format(
    data: any[][],
    options?: CsvFormatOptions,
    taskOptions?: CsvTaskOptions
  ): Promise<CsvTaskResult<string>> {
    const message: CsvWorkerRequestMessage = {
      type: "format",
      taskId: 0,
      data,
      options
    };
    return this._execute(message, taskOptions);
  }

  async load(
    sessionId: string,
    data: any[] | any[][],
    headers?: string[],
    taskOptions?: CsvTaskOptions
  ): Promise<CsvTaskResult<{ rowCount: number; headers: string[] }>> {
    const message: CsvWorkerRequestMessage = {
      type: "load",
      taskId: 0,
      sessionId,
      data,
      headers
    };
    return this._execute(message, taskOptions);
  }

  async getData(
    sessionId: string,
    taskOptions?: CsvTaskOptions
  ): Promise<CsvTaskResult<{ data: Record<string, any>[]; headers: string[]; rowCount: number }>> {
    const message: CsvWorkerRequestMessage = {
      type: "getData",
      taskId: 0,
      sessionId
    };
    return this._execute(message, taskOptions);
  }

  async clear(sessionId?: string, taskOptions?: CsvTaskOptions): Promise<CsvTaskResult<void>> {
    const message: CsvWorkerRequestMessage = {
      type: "clear",
      taskId: 0,
      sessionId
    };
    return this._execute(message, taskOptions);
  }

  async sort(
    sessionId: string,
    config: SortConfig | SortConfig[],
    taskOptions?: CsvTaskOptions
  ): Promise<CsvTaskResult<{ rowCount: number }>> {
    const message: CsvWorkerRequestMessage = {
      type: "sort",
      taskId: 0,
      sessionId,
      config
    };
    return this._execute(message, taskOptions);
  }

  async filter(
    sessionId: string,
    config: FilterConfig,
    taskOptions?: CsvTaskOptions
  ): Promise<FilterResult> {
    return this._unwrap({ type: "filter", taskId: 0, sessionId, config }, taskOptions);
  }

  async search(
    sessionId: string,
    config: SearchConfig,
    taskOptions?: CsvTaskOptions
  ): Promise<FilterResult> {
    return this._unwrap({ type: "search", taskId: 0, sessionId, config }, taskOptions);
  }

  async groupBy(
    sessionId: string,
    config: GroupByConfig,
    taskOptions?: CsvTaskOptions
  ): Promise<GroupResult> {
    return this._unwrap({ type: "groupBy", taskId: 0, sessionId, config }, taskOptions);
  }

  async aggregate(
    sessionId: string,
    config: AggregateConfig[],
    taskOptions?: CsvTaskOptions
  ): Promise<AggregateResult> {
    return this._unwrap({ type: "aggregate", taskId: 0, sessionId, config }, taskOptions);
  }

  async getPage(
    sessionId: string,
    config: PageConfig,
    taskOptions?: CsvTaskOptions
  ): Promise<PageResult> {
    return this._unwrap({ type: "getPage", taskId: 0, sessionId, config }, taskOptions);
  }

  async query(
    sessionId: string,
    config: QueryConfig,
    taskOptions?: CsvTaskOptions
  ): Promise<QueryResult> {
    return this._unwrap({ type: "query", taskId: 0, sessionId, config }, taskOptions);
  }

  getStats(): CsvWorkerPoolStats {
    const busyWorkers = [...this._workers.values()].filter(w => w.busy).length;
    return {
      totalWorkers: this._workers.size,
      busyWorkers,
      pendingTasks: this._pendingQueueSize,
      completedTasks: this._completedTasks,
      failedTasks: this._failedTasks
    };
  }

  terminate(): void {
    if (this._terminated) {
      return;
    }
    this._terminated = true;

    // Reject all pending tasks
    for (const task of this._pendingTasks.values()) {
      task.reject(new Error("Worker pool terminated"));
      this._cleanupTask(task);
    }
    this._pendingTasks.clear();

    for (const queue of [this._highQueue, this._normalQueue, this._lowQueue]) {
      for (const task of queue) {
        task.reject(new Error("Worker pool terminated"));
        this._cleanupTask(task);
      }
      queue.length = 0;
    }

    // Terminate all workers
    for (const poolWorker of this._workers.values()) {
      this._terminateWorker(poolWorker);
    }
    this._workers.clear();

    if (!this._useCustomUrl) {
      releaseWorkerBlobUrl();
    }
  }

  // ---------------------------------------------------------------------------
  // Private Methods
  // ---------------------------------------------------------------------------

  /** Execute and unwrap result - for operations that return data with duration */
  private async _unwrap<T extends { duration: number }>(
    message: CsvWorkerRequestMessage,
    taskOptions?: CsvTaskOptions
  ): Promise<T> {
    const result = await this._execute<T>(message, taskOptions);
    return { ...result.data, duration: result.duration } as T;
  }

  private async _execute<T>(
    message: CsvWorkerRequestMessage,
    taskOptions?: CsvTaskOptions
  ): Promise<CsvTaskResult<T>> {
    if (this._terminated) {
      return Promise.reject(new Error("Worker pool has been terminated"));
    }

    // Ensure pool is initialized (lazy load worker script)
    await this._ensureInitialized();

    if (this._terminated) {
      return Promise.reject(new Error("Worker pool has been terminated"));
    }

    const { priority = "normal", signal } = taskOptions ?? {};

    if (signal?.aborted) {
      return Promise.reject(createAbortError());
    }

    return new Promise((resolve, reject) => {
      const taskId = this._nextTaskId++;
      (message as any).taskId = taskId;

      const task: QueuedTask = {
        taskId,
        message,
        priority,
        resolve,
        reject,
        signal,
        startTime: performance.now()
      };

      if (signal) {
        task.abortHandler = () => this._cancelTask(taskId);
        signal.addEventListener("abort", task.abortHandler, { once: true });
      }

      this._enqueueTask(task);
      this._processQueue();
    });
  }

  private _cleanupTask(task: PendingTask<any>): void {
    if (task.signal && task.abortHandler) {
      task.signal.removeEventListener("abort", task.abortHandler);
    }
  }

  private get _pendingQueueSize(): number {
    return this._highQueue.length + this._normalQueue.length + this._lowQueue.length;
  }

  private _enqueueTask(task: QueuedTask): void {
    switch (task.priority) {
      case "high":
        this._highQueue.push(task);
        break;
      case "low":
        this._lowQueue.push(task);
        break;
      default:
        this._normalQueue.push(task);
        break;
    }
  }

  private _dequeueTask(): QueuedTask | undefined {
    if (this._highQueue.length > 0) {
      return this._highQueue.shift();
    }
    if (this._normalQueue.length > 0) {
      return this._normalQueue.shift();
    }
    return this._lowQueue.shift();
  }

  private _processQueue(): void {
    if (this._terminated || this._pendingQueueSize === 0 || !this._workerUrl) {
      return;
    }

    while (this._pendingQueueSize > 0) {
      let idleWorker: PoolWorker | null = null;
      for (const worker of this._workers.values()) {
        if (!worker.busy) {
          idleWorker = worker;
          break;
        }
      }

      if (!idleWorker && this._workers.size < this._options.maxWorkers) {
        idleWorker = this._createWorker();
      }

      if (!idleWorker) {
        break; // No available workers
      }

      const task = this._dequeueTask()!;
      this._assignTask(idleWorker, task);
    }
  }

  private _createWorker(): PoolWorker {
    if (!this._workerUrl) {
      throw new Error("Worker pool not initialized. Call _ensureInitialized() first.");
    }

    const id = this._nextWorkerId++;
    const worker = new Worker(this._workerUrl);

    const poolWorker: PoolWorker = {
      id,
      worker,
      busy: false,
      currentTaskId: null,
      idleTimer: null
    };

    worker.onmessage = (event: MessageEvent<CsvWorkerResponseMessage>) => {
      this._handleWorkerMessage(poolWorker, event.data);
    };

    worker.onerror = (event: ErrorEvent) => {
      this._handleWorkerError(poolWorker, event);
    };

    this._workers.set(id, poolWorker);
    return poolWorker;
  }

  private _assignTask(poolWorker: PoolWorker, task: QueuedTask): void {
    if (poolWorker.idleTimer) {
      clearTimeout(poolWorker.idleTimer);
      poolWorker.idleTimer = null;
    }

    poolWorker.busy = true;
    poolWorker.currentTaskId = task.taskId;
    this._pendingTasks.set(task.taskId, task);

    poolWorker.worker.postMessage(task.message);
  }

  private _handleWorkerMessage(poolWorker: PoolWorker, msg: CsvWorkerResponseMessage): void {
    if (msg.type === "ready") {
      this._processQueue();
      return;
    }

    const taskId = (msg as any).taskId;
    if (taskId === undefined) {
      return;
    }

    const task = this._pendingTasks.get(taskId);
    if (!task) {
      this._releaseWorker(poolWorker);
      return;
    }

    this._pendingTasks.delete(taskId);
    this._cleanupTask(task);

    if (msg.type === "error") {
      this._failedTasks++;
      task.reject(new Error(msg.error));
    } else {
      // Unified response: { type: "result", taskId, data, duration }
      this._completedTasks++;
      task.resolve({ data: msg.data, duration: msg.duration });
    }

    this._releaseWorker(poolWorker);
  }

  private _handleWorkerError(poolWorker: PoolWorker, event: ErrorEvent): void {
    const taskId = poolWorker.currentTaskId;
    if (taskId !== null) {
      const task = this._pendingTasks.get(taskId);
      if (task) {
        this._pendingTasks.delete(taskId);
        this._failedTasks++;
        this._cleanupTask(task);
        task.reject(new Error(event.message || "Worker error"));
      }
    }

    this._workers.delete(poolWorker.id);
    poolWorker.worker.terminate();
    this._processQueue();
  }

  private _releaseWorker(poolWorker: PoolWorker): void {
    poolWorker.busy = false;
    poolWorker.currentTaskId = null;

    if (this._workers.size > this._options.minWorkers) {
      poolWorker.idleTimer = setTimeout(() => {
        if (!poolWorker.busy && this._workers.size > this._options.minWorkers) {
          this._workers.delete(poolWorker.id);
          this._terminateWorker(poolWorker);
        }
      }, this._options.idleTimeout);
    }

    this._processQueue();
  }

  private _terminateWorker(poolWorker: PoolWorker): void {
    if (poolWorker.idleTimer) {
      clearTimeout(poolWorker.idleTimer);
    }
    try {
      poolWorker.worker.postMessage({ type: "terminate" });
    } catch {
      // Ignore errors
    }
    poolWorker.worker.terminate();
  }

  private _cancelTask(taskId: number): void {
    for (const queue of [this._highQueue, this._normalQueue, this._lowQueue]) {
      const idx = queue.findIndex(t => t.taskId === taskId);
      if (idx !== -1) {
        const task = queue.splice(idx, 1)[0];
        this._cleanupTask(task);
        queueMicrotask(() => task.reject(createAbortError()));
        return;
      }
    }

    const task = this._pendingTasks.get(taskId);
    if (task) {
      this._pendingTasks.delete(taskId);
      this._cleanupTask(task);
      queueMicrotask(() => task.reject(createAbortError()));
    }
  }
}

// =============================================================================
// CsvWorkerSession - High-level API (Public)
// =============================================================================

let sessionIdCounter = 0;

/**
 * High-level API for interactive CSV data exploration.
 *
 * Keeps data in worker memory for efficient repeated operations.
 *
 * @example
 * ```ts
 * const session = await CsvWorkerSession.create();
 *
 * // Load data
 * await session.load(csvString, { headers: true });
 *
 * // Batch query - most efficient for multiple operations
 * const result = await session.query({
 *   sort: { column: 'age', order: 'desc' },
 *   filter: { conditions: [{ column: 'status', operator: 'eq', value: 'active' }] },
 *   page: { page: 1, pageSize: 50 }
 * });
 *
 * // Or use individual operations
 * await session.sort({ column: 'name', order: 'asc' });
 * const filtered = await session.filter({
 *   conditions: [{ column: 'age', operator: 'gt', value: 30 }]
 * });
 *
 * // Cleanup
 * session.dispose();
 * ```
 */
export class CsvWorkerSession {
  private readonly _pool: CsvWorkerPool;
  private readonly _sessionId: string;
  private _disposed = false;
  private _headers: string[] = [];
  private _rowCount = 0;

  /** Use CsvWorkerSession.create() instead for lazy pool initialization */
  constructor(pool: CsvWorkerPool) {
    this._sessionId = `session_${++sessionIdCounter}_${Date.now()}`;
    this._pool = pool;
  }

  /** Create a new session with optional pool */
  static async create(pool?: CsvWorkerPool): Promise<CsvWorkerSession> {
    const resolvedPool = pool ?? (await getDefaultWorkerPool());
    return new CsvWorkerSession(resolvedPool);
  }

  get sessionId(): string {
    return this._sessionId;
  }

  get headers(): string[] {
    return this._headers;
  }

  get rowCount(): number {
    return this._rowCount;
  }

  /**
   * Load CSV string or data into session
   */
  async load(
    csvOrData: string | any[] | any[][],
    options?: CsvParseOptions & { headers?: string[] | boolean }
  ): Promise<{ rowCount: number; headers: string[] }> {
    if (this._disposed) {
      throw new CsvWorkerError("Session has been disposed");
    }

    if (typeof csvOrData === "string") {
      const parseOptions = { ...options, sessionId: this._sessionId };
      const result = await this._pool.parse(csvOrData, parseOptions);
      const data = result.data as any;
      this._headers = data.headers || [];
      this._rowCount = data.rows?.length ?? (Array.isArray(data) ? data.length : 0);
      return { rowCount: this._rowCount, headers: this._headers };
    } else {
      const result = await this._pool.load(
        this._sessionId,
        csvOrData,
        Array.isArray(options?.headers) ? (options.headers as string[]) : undefined
      );
      this._headers = result.data.headers;
      this._rowCount = result.data.rowCount;
      return result.data;
    }
  }

  /** Get all data */
  getData() {
    return this._wrap(() => this._pool.getData(this._sessionId).then(r => r.data))();
  }

  /** Sort data in place */
  sort(config: SortConfig | SortConfig[]) {
    return this._wrap(() => this._pool.sort(this._sessionId, config).then(r => r.data))();
  }

  /** Filter data (resets to original data before filtering) */
  filter(config: FilterConfig) {
    return this._wrap(() => this._pool.filter(this._sessionId, config))();
  }

  /** Search across columns */
  search(config: SearchConfig) {
    return this._wrap(() => this._pool.search(this._sessionId, config))();
  }

  /** Group by and aggregate */
  groupBy(config: GroupByConfig) {
    return this._wrap(() => this._pool.groupBy(this._sessionId, config))();
  }

  /** Aggregate entire dataset */
  aggregate(config: AggregateConfig[]) {
    return this._wrap(() => this._pool.aggregate(this._sessionId, config))();
  }

  /** Get paginated data */
  getPage(config: PageConfig) {
    return this._wrap(() => this._pool.getPage(this._sessionId, config))();
  }

  /**
   * Execute batch query - multiple operations in single round-trip
   * Order of operations: sort -> filter -> search -> groupBy/aggregate -> page
   */
  query(config: QueryConfig) {
    return this._wrap(() => this._pool.query(this._sessionId, config))();
  }

  /** Dispose session and free worker memory */
  async dispose(): Promise<void> {
    if (this._disposed) {
      return;
    }
    this._disposed = true;
    await this._pool.clear(this._sessionId).catch(() => {});
  }

  /** Wrap operation with disposed check */
  private _wrap<T>(fn: () => Promise<T>): () => Promise<T> {
    return () => {
      if (this._disposed) {
        return Promise.reject(new CsvWorkerError("Session has been disposed"));
      }
      return fn();
    };
  }
}

// =============================================================================
// Default Pool & Convenience Functions
// =============================================================================

let defaultPool: CsvWorkerPool | null = null;
let defaultPoolPromise: Promise<CsvWorkerPool> | null = null;

/** Get or create the default worker pool (with lazy initialization) */
export async function getDefaultWorkerPool(): Promise<CsvWorkerPool> {
  if (defaultPool) {
    return defaultPool;
  }
  if (!defaultPoolPromise) {
    defaultPoolPromise = CsvWorkerPool.create()
      .then(pool => {
        defaultPool = pool;
        return pool;
      })
      .catch(err => {
        defaultPoolPromise = null;
        throw err;
      });
  }
  return defaultPoolPromise;
}

export function terminateDefaultWorkerPool(): void {
  if (defaultPool) {
    defaultPool.terminate();
    defaultPool = null;
    defaultPoolPromise = null;
  } else if (defaultPoolPromise) {
    // Handle in-flight pool creation: wait for it to resolve, then terminate
    const pending = defaultPoolPromise;
    defaultPoolPromise = null;
    pending.then(pool => pool.terminate()).catch(() => {});
  }
}

/** Parse CSV using worker pool */
export async function parseWithPool(
  data: string,
  options?: CsvParseOptions,
  taskOptions?: CsvTaskOptions
): Promise<CsvTaskResult<string[][] | CsvParseResult<Record<string, string>>>> {
  const pool = await getDefaultWorkerPool();
  return pool.parse(data, options, taskOptions);
}

/** Format data to CSV using worker pool */
export async function formatWithPool(
  data: any[][],
  options?: CsvFormatOptions,
  taskOptions?: CsvTaskOptions
): Promise<CsvTaskResult<string>> {
  const pool = await getDefaultWorkerPool();
  return pool.format(data, options, taskOptions);
}

// =============================================================================
// Re-exports
// =============================================================================

// Export Pool class for tests and advanced usage
export { CsvWorkerPool };

export type {
  // Config types
  CsvWorkerPoolOptions,
  CsvWorkerPoolStats,
  CsvTaskOptions,
  CsvTaskResult,
  CsvTaskPriority,
  SortConfig,
  SortOrder,
  FilterConfig,
  FilterCondition,
  FilterOperator,
  SearchConfig,
  GroupByConfig,
  AggregateConfig,
  AggregateFunction,
  PageConfig,
  QueryConfig,
  // Result types
  FilterResult,
  PageResult,
  GroupResult,
  AggregateResult,
  QueryResult
} from "@csv/worker/types";
