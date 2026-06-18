/**
 * CSV Worker Module - Browser Entry Point
 */

export {
  CsvWorkerPool,
  CsvWorkerSession,
  hasWorkerSupport,
  getDefaultWorkerPool,
  terminateDefaultWorkerPool,
  parseWithPool,
  formatWithPool
} from "@csv/worker/pool.browser";

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

export { getWorkerBlobUrl, releaseWorkerBlobUrl } from "@csv/worker/worker-script.bundle";
