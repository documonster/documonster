/**
 * documonster/excel — browser entry.
 *
 * Same surface as the Node entry via the shared base, but the two
 * platform-specific namespaces resolve to their browser variants:
 * `Workbook` (cross-platform IO only, no Node file-path operations) and
 * `Stream`.
 */
export * from "./index.base";

export * as Workbook from "@excel/surface/workbook.browser";
export * as Stream from "@excel/surface/stream.browser";
