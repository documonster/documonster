/**
 * CSV Worker Types
 *
 * Simplified type definitions for the CSV worker system.
 * Uses generics and base types to reduce repetition.
 */

import type { CsvParseOptions, CsvFormatOptions, CsvParseResult, Row } from "@csv/types";

// =============================================================================
// Base Result Type (for reuse)
// =============================================================================

/** Base result with duration */
interface BaseResult {
  duration: number;
}

// =============================================================================
// Data Operation Types
// =============================================================================

/** Sort order */
export type SortOrder = "asc" | "desc";

/** Sort configuration */
export interface SortConfig {
  column: string | number;
  order?: SortOrder;
  comparator?: "string" | "number" | "date" | "auto";
}

/** Filter operator */
export type FilterOperator =
  | "eq"
  | "neq"
  | "gt"
  | "gte"
  | "lt"
  | "lte"
  | "contains"
  | "startsWith"
  | "endsWith"
  | "regex"
  | "in"
  | "notIn"
  | "isNull"
  | "notNull";

/** Filter condition */
export interface FilterCondition {
  column: string | number;
  operator: FilterOperator;
  /** Compare value: a scalar for most operators, or an array for `in`/`notIn`. */
  value?: unknown;
  ignoreCase?: boolean;
}

/** Filter configuration */
export interface FilterConfig {
  conditions: FilterCondition[];
  logic?: "and" | "or";
}

/** Aggregation function */
export type AggregateFunction = "count" | "sum" | "avg" | "min" | "max" | "first" | "last";

/** Aggregation configuration */
export interface AggregateConfig {
  column: string | number;
  fn: AggregateFunction;
  alias?: string;
}

/** Group by configuration */
export interface GroupByConfig {
  columns: (string | number)[];
  aggregates: AggregateConfig[];
}

/** Pagination configuration */
export interface PageConfig {
  page: number;
  pageSize: number;
}

/** Search configuration */
export interface SearchConfig {
  query: string;
  columns?: (string | number)[];
  ignoreCase?: boolean;
}

// =============================================================================
// Batch Query API
// =============================================================================

/**
 * Batch query configuration - execute multiple operations in one round-trip
 */
export interface QueryConfig {
  /** Sort configuration (applied first) */
  sort?: SortConfig | SortConfig[];
  /** Filter configuration (applied after sort) */
  filter?: FilterConfig;
  /** Search configuration (applied after filter) */
  search?: SearchConfig;
  /** Group by configuration */
  groupBy?: GroupByConfig;
  /** Aggregation configuration */
  aggregate?: AggregateConfig[];
  /** Pagination (applied last) */
  page?: PageConfig;
}

// =============================================================================
// Worker Message Types (Simplified)
// =============================================================================

/** All supported worker operations */
export type CsvWorkerTaskType =
  | "parse"
  | "format"
  | "load"
  | "sort"
  | "filter"
  | "search"
  | "groupBy"
  | "aggregate"
  | "getPage"
  | "getData"
  | "clear"
  | "query"
  | "terminate";

/** Base request structure */
interface BaseRequest<T extends CsvWorkerTaskType> {
  type: T;
  taskId: number;
  sessionId?: string;
}

/** Worker request messages (main thread -> worker) */
export type CsvWorkerRequestMessage =
  | (BaseRequest<"parse"> & { data: string; options?: CsvParseOptions })
  | (BaseRequest<"format"> & {
      data: Row[] | Record<string, unknown>[];
      options?: CsvFormatOptions;
    })
  | (BaseRequest<"load"> & {
      sessionId: string;
      data: Record<string, unknown>[] | unknown[][];
      headers?: string[];
    })
  | (BaseRequest<"getData"> & { sessionId: string })
  | (BaseRequest<"clear"> & { sessionId?: string })
  | (BaseRequest<"sort"> & { sessionId: string; config: SortConfig | SortConfig[] })
  | (BaseRequest<"filter"> & { sessionId: string; config: FilterConfig })
  | (BaseRequest<"search"> & { sessionId: string; config: SearchConfig })
  | (BaseRequest<"groupBy"> & { sessionId: string; config: GroupByConfig })
  | (BaseRequest<"aggregate"> & { sessionId: string; config: AggregateConfig[] })
  | (BaseRequest<"getPage"> & { sessionId: string; config: PageConfig })
  | (BaseRequest<"query"> & { sessionId: string; config: QueryConfig })
  | BaseRequest<"terminate">;

/** Unified worker response - single format for all results */
export type CsvWorkerResponseMessage =
  | { type: "ready" }
  | { type: "result"; taskId: number; data: unknown; duration: number }
  | { type: "error"; taskId: number; error: string; duration: number };

// =============================================================================
// Pool & Task Types
// =============================================================================

/** Task priority */
export type CsvTaskPriority = "high" | "normal" | "low";

/** Pool options */
export interface CsvWorkerPoolOptions {
  maxWorkers?: number;
  minWorkers?: number;
  idleTimeout?: number;
  workerUrl?: string;
}

/** Task options */
export interface CsvTaskOptions {
  priority?: CsvTaskPriority;
  signal?: AbortSignal;
}

/** Task result */
export interface CsvTaskResult<T> {
  data: T;
  duration: number;
}

/** Pool statistics */
export interface CsvWorkerPoolStats {
  totalWorkers: number;
  busyWorkers: number;
  pendingTasks: number;
  completedTasks: number;
  failedTasks: number;
}

// =============================================================================
// Result Types (using BaseResult)
// =============================================================================

/** Filter/Search result */
export interface FilterResult<T = Record<string, unknown>> extends BaseResult {
  data: T[];
  matchCount: number;
  totalCount: number;
}

/** Pagination result */
export interface PageResult<T = Record<string, unknown>> extends BaseResult {
  data: T[];
  page: number;
  pageSize: number;
  totalRows: number;
  totalPages: number;
}

/** Group by result */
export interface GroupResult<T = Record<string, unknown>> extends BaseResult {
  data: T[];
  groupCount: number;
}

/** Aggregation result */
export interface AggregateResult extends BaseResult {
  data: Record<string, unknown>;
}

/** Batch query result - adapts based on operations */
export interface QueryResult<T = Record<string, unknown>> extends BaseResult {
  data: T[];
  page?: number;
  pageSize?: number;
  totalRows?: number;
  totalPages?: number;
  matchCount?: number;
  groupCount?: number;
  aggregates?: Record<string, unknown>;
}

// Re-export core CSV types for convenience - allows worker module consumers
// to import all needed types from a single location
export type { CsvParseOptions, CsvFormatOptions, CsvParseResult };
