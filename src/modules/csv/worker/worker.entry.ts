/**
 * CSV Worker Entry
 *
 * This file runs inside a Web Worker (classic script after bundling).
 * It implements the worker message protocol used by CsvWorkerPool.
 *
 * IMPORTANT:
 * - Keep this file browser/worker-safe (no Node.js APIs)
 * - Parsing/formatting are delegated to the main CSV implementations to avoid drift.
 */

import { formatCsv } from "@csv/format";
import { parseCsv } from "@csv/parse";
import type {
  CsvWorkerRequestMessage,
  CsvWorkerResponseMessage,
  QueryConfig,
  SortConfig,
  FilterConfig,
  FilterCondition,
  SearchConfig,
  GroupByConfig,
  AggregateConfig,
  PageConfig
} from "@csv/worker/types";
import { isForbiddenKey } from "@utils/object";

/** A CSV cell carries an unknown scalar (string in array mode, parsed value in typed mode). */
type CsvCell = unknown;
/**
 * A worker row, indexable by column key or numeric index. Object rows are
 * `Record<string, unknown>`; array rows are `unknown[]` (also assignable to a
 * numeric-indexed record), so this single shape covers both without `any`.
 */
type CsvWorkerRow = Record<string | number, CsvCell>;

type SessionData = Record<string, unknown>[] | CsvCell[][];

interface WorkerSession {
  data: CsvWorkerRow[];
  headers: string[] | null;
  originalData: CsvWorkerRow[];
}

const sessions = new Map<string, WorkerSession>();

function getSession(sessionId: string): WorkerSession {
  const session = sessions.get(sessionId);
  if (!session) {
    throw new Error(`Session not found: ${sessionId}`);
  }
  return session;
}

function reply(taskId: number, start: number, data: unknown): void {
  (self as unknown as Worker).postMessage({
    type: "result",
    taskId,
    data,
    duration: performance.now() - start
  } satisfies CsvWorkerResponseMessage);
}

function replyError(taskId: number, start: number, error: unknown): void {
  const message = error instanceof Error ? error.message : String(error);
  (self as unknown as Worker).postMessage({
    type: "error",
    taskId,
    error: message,
    duration: performance.now() - start
  } satisfies CsvWorkerResponseMessage);
}

function toObjectRows(
  data: SessionData,
  headers?: string[] | null
): { rows: CsvWorkerRow[]; headers: string[] } {
  if (!Array.isArray(data)) {
    return { rows: [], headers: headers ?? [] };
  }

  if (data.length === 0) {
    return { rows: [], headers: headers ?? [] };
  }

  const first = data[0];

  // Already objects
  if (first && typeof first === "object" && !Array.isArray(first)) {
    const inferredHeaders = headers ?? Object.keys(first);
    return { rows: data as CsvWorkerRow[], headers: inferredHeaders };
  }

  // Array rows
  const arrayRows = data as CsvCell[][];
  let resolvedHeaders: string[];
  let rows: CsvCell[][];

  if (headers && headers.length > 0) {
    resolvedHeaders = headers;
    rows = arrayRows;
  } else {
    resolvedHeaders = (arrayRows[0] ?? []).map(v => String(v));
    rows = arrayRows.slice(1);
  }

  const objects = rows.map(row => {
    const obj: CsvWorkerRow = Object.create(null) as CsvWorkerRow;
    for (let i = 0; i < resolvedHeaders.length; i++) {
      const key = resolvedHeaders[i];
      if (!isForbiddenKey(key)) {
        obj[key] = row[i];
      }
    }
    return obj;
  });

  return { rows: objects, headers: resolvedHeaders };
}

// =============================================================================
// Data operations (sort/filter/search/groupBy/aggregate/page/query)
// =============================================================================

function sortData(data: CsvWorkerRow[], configs: SortConfig | SortConfig[]): void {
  const list = Array.isArray(configs) ? configs : [configs];
  data.sort((a, b) => {
    for (const config of list) {
      const { column, order = "asc", comparator = "auto" } = config;
      const aVal = a[column];
      const bVal = b[column];
      let result: number;
      if (
        comparator === "number" ||
        (comparator === "auto" && !Number.isNaN(Number(aVal)) && !Number.isNaN(Number(bVal)))
      ) {
        result = Number(aVal ?? 0) - Number(bVal ?? 0);
      } else if (comparator === "date") {
        result =
          new Date((aVal ?? 0) as string | number | Date).getTime() -
          new Date((bVal ?? 0) as string | number | Date).getTime();
      } else {
        result = String(aVal ?? "").localeCompare(String(bVal ?? ""));
      }
      if (result !== 0) {
        return order === "desc" ? -result : result;
      }
    }
    return 0;
  });
}

function evaluateCondition(
  row: CsvWorkerRow,
  condition: FilterCondition,
  compiledRegex?: RegExp
): boolean {
  const { column, operator, value, ignoreCase = false } = condition;
  let fieldValue: unknown = row?.[column];
  let compareValue: unknown = value;

  if (ignoreCase && typeof fieldValue === "string" && operator !== "regex") {
    fieldValue = fieldValue.toLowerCase();
    if (typeof compareValue === "string") {
      compareValue = compareValue.toLowerCase();
    } else if (Array.isArray(compareValue)) {
      compareValue = compareValue.map((v: unknown) =>
        typeof v === "string" ? v.toLowerCase() : v
      );
    }
  }

  switch (operator) {
    case "eq":
      return fieldValue === compareValue;
    case "neq":
      return fieldValue !== compareValue;
    case "gt":
      return Number(fieldValue) > Number(compareValue);
    case "gte":
      return Number(fieldValue) >= Number(compareValue);
    case "lt":
      return Number(fieldValue) < Number(compareValue);
    case "lte":
      return Number(fieldValue) <= Number(compareValue);
    case "contains": {
      const fv = ignoreCase ? String(fieldValue).toLowerCase() : String(fieldValue);
      const cv = ignoreCase ? String(compareValue).toLowerCase() : String(compareValue);
      return fv.includes(cv);
    }
    case "startsWith": {
      const fv = ignoreCase ? String(fieldValue).toLowerCase() : String(fieldValue);
      const cv = ignoreCase ? String(compareValue).toLowerCase() : String(compareValue);
      return fv.startsWith(cv);
    }
    case "endsWith": {
      const fv = ignoreCase ? String(fieldValue).toLowerCase() : String(fieldValue);
      const cv = ignoreCase ? String(compareValue).toLowerCase() : String(compareValue);
      return fv.endsWith(cv);
    }
    case "regex": {
      const re = compiledRegex ?? new RegExp(String(compareValue), ignoreCase ? "i" : "");
      return re.test(String(fieldValue));
    }
    case "in":
      return Array.isArray(compareValue) && compareValue.includes(fieldValue);
    case "notIn":
      return !Array.isArray(compareValue) || !compareValue.includes(fieldValue);
    case "isNull":
      return fieldValue === null || fieldValue === undefined || fieldValue === "";
    case "notNull":
      return fieldValue !== null && fieldValue !== undefined && fieldValue !== "";
    default:
      return true;
  }
}

function filterData(data: CsvWorkerRow[], config: FilterConfig): CsvWorkerRow[] {
  const { conditions, logic = "and" } = config;

  // Pre-compile regex patterns to avoid re-creating RegExp per row
  const compiledRegexMap = new Map<FilterCondition, RegExp>();
  for (const cond of conditions) {
    if (cond.operator === "regex") {
      compiledRegexMap.set(cond, new RegExp(String(cond.value), cond.ignoreCase ? "i" : ""));
    }
  }

  const evaluate =
    logic === "and"
      ? (row: CsvWorkerRow) =>
          conditions.every(cond => evaluateCondition(row, cond, compiledRegexMap.get(cond)))
      : (row: CsvWorkerRow) =>
          conditions.some(cond => evaluateCondition(row, cond, compiledRegexMap.get(cond)));
  return data.filter(evaluate);
}

function searchData(data: CsvWorkerRow[], config: SearchConfig): CsvWorkerRow[] {
  const { query, columns, ignoreCase = true } = config;
  const searchQuery = ignoreCase ? query.toLowerCase() : query;

  const resolvedColumns = columns ?? Object.keys(data[0] ?? {});
  return data.filter(row => {
    return resolvedColumns.some(col => {
      let value = String(row[col] ?? "");
      if (ignoreCase) {
        value = value.toLowerCase();
      }
      return value.includes(searchQuery);
    });
  });
}

function computeAggregate(
  rows: CsvWorkerRow[],
  column: string | number,
  fn: AggregateConfig["fn"]
): unknown {
  if (fn === "count") {
    return rows.length;
  }
  if (fn === "first") {
    return rows.length > 0 ? rows[0]?.[column] : null;
  }
  if (fn === "last") {
    return rows.length > 0 ? rows[rows.length - 1]?.[column] : null;
  }

  const nums = rows.map(r => Number(r?.[column])).filter(n => !Number.isNaN(n));

  if (nums.length === 0) {
    return fn === "avg" ? 0 : null;
  }

  if (fn === "sum" || fn === "avg") {
    const sum = nums.reduce((a, b) => a + b, 0);
    return fn === "avg" ? sum / nums.length : sum;
  }

  if (fn === "min") {
    return nums.reduce((a, b) => (a < b ? a : b), nums[0]);
  }
  if (fn === "max") {
    return nums.reduce((a, b) => (a > b ? a : b), nums[0]);
  }

  return null;
}

function groupByData(data: CsvWorkerRow[], config: GroupByConfig): CsvWorkerRow[] {
  const { columns, aggregates } = config;
  const groups = new Map<string, { keyValues: unknown[]; rows: CsvWorkerRow[] }>();

  for (const row of data) {
    const keyValues = columns.map(col => row[col]);
    const key = keyValues.join("\0");
    const existing = groups.get(key);
    if (existing) {
      existing.rows.push(row);
    } else {
      groups.set(key, { keyValues, rows: [row] });
    }
  }

  const result: CsvWorkerRow[] = [];
  for (const group of groups.values()) {
    const obj: CsvWorkerRow = Object.create(null) as CsvWorkerRow;
    columns.forEach((col, idx) => {
      const k = String(col);
      if (!isForbiddenKey(k)) {
        obj[k] = group.keyValues[idx];
      }
    });
    for (const { column, fn, alias } of aggregates) {
      const key = alias || `${column}_${fn}`;
      if (!isForbiddenKey(key)) {
        obj[key] = computeAggregate(group.rows, column, fn);
      }
    }
    result.push(obj);
  }

  return result;
}

function aggregateData(data: CsvWorkerRow[], configs: AggregateConfig[]): Record<string, unknown> {
  const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  for (const config of configs) {
    const { column, fn, alias } = config;
    const key = alias || `${column}_${fn}`;
    if (!isForbiddenKey(key)) {
      result[key] = computeAggregate(data, column, fn);
    }
  }
  return result;
}

function getPageData(
  data: CsvWorkerRow[],
  config: PageConfig
): {
  data: CsvWorkerRow[];
  page: number;
  pageSize: number;
  totalRows: number;
  totalPages: number;
} {
  const page = Math.max(1, config.page);
  let { pageSize } = config;
  if (pageSize <= 0) {
    pageSize = data.length || 1;
  }
  const start = (page - 1) * pageSize;
  return {
    data: data.slice(start, start + pageSize),
    page,
    pageSize,
    totalRows: data.length,
    totalPages: Math.ceil(data.length / pageSize)
  };
}

interface QueryResultData {
  data: CsvWorkerRow[];
  matchCount?: number;
  groupCount?: number;
  aggregates?: Record<string, unknown>;
  page?: number;
  pageSize?: number;
  totalRows?: number;
  totalPages?: number;
}

function executeQuery(session: WorkerSession, config: QueryConfig): QueryResultData {
  let data = config.sort ? [...session.originalData] : session.originalData;
  const result: QueryResultData = { data: [] };

  if (config.sort) {
    sortData(data, config.sort);
  }

  if (config.filter) {
    data = filterData(data, config.filter);
    result.matchCount = data.length;
  }

  if (config.search) {
    data = searchData(data, config.search);
    result.matchCount = data.length;
  }

  if (config.groupBy) {
    data = groupByData(data, config.groupBy);
    result.groupCount = data.length;
  } else if (config.aggregate) {
    result.aggregates = aggregateData(data, config.aggregate);
  }

  if (config.page) {
    const pageResult = getPageData(data, config.page);
    result.data = pageResult.data;
    result.page = pageResult.page;
    result.pageSize = pageResult.pageSize;
    result.totalRows = pageResult.totalRows;
    result.totalPages = pageResult.totalPages;
  } else {
    result.data = data;
  }

  return result;
}

// =============================================================================
// Message handler
// =============================================================================

// Dedicated Web Workers receive messages only from the parent thread —
// cross-origin messages are impossible by spec. The origin guard below is
// a no-op at runtime but satisfies CodeQL's js/missing-origin-check rule.
self.addEventListener("message", (event: MessageEvent<CsvWorkerRequestMessage>) => {
  if (event.origin !== "" && event.origin !== self.location?.origin) {
    return;
  }
  // Validate incoming message structure (defense-in-depth for dedicated worker)
  const msg = event.data;
  if (!msg || typeof msg.type !== "string") {
    return;
  }
  const taskId = msg.taskId ?? 0;
  const start = performance.now();

  try {
    switch (msg.type) {
      case "parse": {
        const result = msg.options ? parseCsv(msg.data, msg.options) : parseCsv(msg.data);

        if (msg.sessionId) {
          const objResult =
            result && typeof result === "object" && "headers" in result ? result : undefined;
          sessions.set(msg.sessionId, {
            data: (objResult ? objResult.rows : result) as CsvWorkerRow[],
            headers: objResult ? (objResult.headers ?? null) : null,
            originalData: (objResult
              ? [...objResult.rows]
              : [...(result as unknown[])]) as CsvWorkerRow[]
          });
        }

        reply(taskId, start, result);
        break;
      }

      case "format": {
        reply(taskId, start, formatCsv(msg.data, msg.options));
        break;
      }

      case "load": {
        const { rows, headers } = toObjectRows(msg.data, msg.headers ?? null);
        sessions.set(msg.sessionId, {
          data: rows,
          headers: headers ?? null,
          originalData: [...rows]
        });
        reply(taskId, start, { rowCount: rows.length, headers });
        break;
      }

      case "getData": {
        const session = getSession(msg.sessionId);
        reply(taskId, start, {
          data: session.data,
          headers: session.headers || [],
          rowCount: session.data.length
        });
        break;
      }

      case "clear": {
        if (msg.sessionId) {
          sessions.delete(msg.sessionId);
        } else {
          sessions.clear();
        }
        reply(taskId, start, undefined);
        break;
      }

      case "sort": {
        const session = getSession(msg.sessionId);
        sortData(session.data, msg.config);
        session.originalData = [...session.data];
        reply(taskId, start, { rowCount: session.data.length });
        break;
      }

      case "filter": {
        const session = getSession(msg.sessionId);
        const totalCount = session.originalData.length;
        session.data = filterData(session.originalData, msg.config);
        reply(taskId, start, {
          data: session.data,
          matchCount: session.data.length,
          totalCount
        });
        break;
      }

      case "search": {
        const session = getSession(msg.sessionId);
        const totalCount = session.originalData.length;
        session.data = searchData(session.originalData, msg.config);
        reply(taskId, start, {
          data: session.data,
          matchCount: session.data.length,
          totalCount
        });
        break;
      }

      case "groupBy": {
        const session = getSession(msg.sessionId);
        const groups = groupByData(session.data, msg.config);
        reply(taskId, start, { data: groups, groupCount: groups.length });
        break;
      }

      case "aggregate": {
        const session = getSession(msg.sessionId);
        reply(taskId, start, { data: aggregateData(session.data, msg.config) });
        break;
      }

      case "getPage": {
        const session = getSession(msg.sessionId);
        reply(taskId, start, getPageData(session.data, msg.config));
        break;
      }

      case "query": {
        const session = getSession(msg.sessionId);
        reply(taskId, start, executeQuery(session, msg.config as QueryConfig));
        break;
      }

      case "terminate": {
        sessions.clear();
        break;
      }

      default:
        throw new Error(`Unknown message type: ${(msg as CsvWorkerRequestMessage).type}`);
    }
  } catch (error) {
    replyError(taskId, start, error);
  }
});

// Signal ready
(self as unknown as Worker).postMessage({ type: "ready" } satisfies CsvWorkerResponseMessage);
