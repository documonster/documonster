/**
 * CSV Worker Pool - Node.js Stub
 *
 * Provides noop/error implementations for Node.js.
 * Web Workers are browser-only.
 */

const ERROR_MSG = "CsvWorkerPool is only available in browser environments";

function throwNotSupported(): never {
  throw new Error(ERROR_MSG);
}

export function hasWorkerSupport(): boolean {
  return false;
}

/**
 * Node.js stub - all methods throw "not supported" except terminate/dispose.
 * Uses a Proxy to avoid repeating throwNotSupported() for every method.
 */
const stubHandler: ProxyHandler<object> = {
  get(_target, prop) {
    // Allow these methods to be no-ops
    if (prop === "terminate") {
      return () => {};
    }
    if (prop === "dispose") {
      return () => Promise.resolve();
    }
    // Constructor check
    if (prop === "prototype") {
      return {};
    }
    // Prevent being treated as thenable (avoids confusing errors with `await`)
    if (prop === "then" || typeof prop === "symbol") {
      return undefined;
    }
    // Everything else throws
    return throwNotSupported;
  },
  construct() {
    throwNotSupported();
  }
};

// Create stub classes using Proxy
export const CsvWorkerPool = new Proxy(function () {} as any, stubHandler);
export const CsvWorkerSession = new Proxy(function () {} as any, stubHandler);

export function getDefaultWorkerPool(): typeof CsvWorkerPool {
  throwNotSupported();
}

export function terminateDefaultWorkerPool(): void {}

export function parseWithPool(): never {
  throwNotSupported();
}

export function formatWithPool(): never {
  throwNotSupported();
}

export type {
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
  FilterResult,
  PageResult,
  GroupResult,
  AggregateResult,
  QueryResult
} from "@csv/worker/types";
