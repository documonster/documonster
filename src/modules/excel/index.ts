/**
 * documonster/excel — Node entry.
 *
 * Re-exports the platform-independent base (domain namespaces, public types and
 * errors) and adds the Node variants of the two platform-specific namespaces:
 * `Workbook` (file-path IO) and `Stream`. Every public type — the IO and
 * streaming option types included — comes from the base, the same on both
 * platforms; the streaming classes' own types are members of `Stream`.
 */
export * from "@excel/index.base";

export * as Workbook from "@excel/surface/workbook";
export * as Stream from "@excel/surface/stream";
