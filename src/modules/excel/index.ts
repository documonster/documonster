/**
 * documonster/excel — Node entry.
 *
 * Re-exports the platform-independent base (domain namespaces + errors) and
 * adds the Node variants of the two platform-specific namespaces:
 * `Workbook` (file-path IO) and `Stream`.
 */
export * from "@excel/index.base";

export * as Workbook from "@excel/surface/workbook";
export * as Stream from "@excel/surface/stream";
