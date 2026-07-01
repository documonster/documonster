/**
 * CSV Worker Module - Node.js Entry Point
 *
 * Exports stub implementations. Web Workers are browser-only.
 */

export {
  CsvWorkerPool,
  CsvWorkerSession,
  hasWorkerSupport,
  getDefaultWorkerPool,
  terminateDefaultWorkerPool,
  parseWithPool,
  formatWithPool
} from "@csv/worker/pool";

export type {
  // Pool options
  CsvWorkerPoolOptions,
  CsvWorkerPoolStats,
  CsvTaskOptions,
  CsvTaskResult,
  CsvTaskPriority,
  // Operation configs
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
