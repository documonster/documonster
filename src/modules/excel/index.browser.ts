/**
 * documonster/excel — browser entry.
 *
 * Same dot-namespaces as the Node entry; `Workbook` resolves to the browser
 * surface (cross-platform IO only, no Node file-path operations).
 */
export * as Workbook from "@excel/surface/workbook.browser";
export * as Worksheet from "@excel/surface/worksheet";
export * as Cell from "@excel/surface/cell";
export * as Row from "@excel/surface/row";
export * as Column from "@excel/surface/column";
export * as Range from "@excel/surface/range";
export * as Chart from "@excel/surface/chart";
export * as Table from "@excel/surface/table";
export * as Image from "@excel/surface/image";
export * as Pivot from "@excel/surface/pivot";
export * as Sparkline from "@excel/surface/sparkline";
export * as Form from "@excel/surface/form";
export * as Chartsheet from "@excel/surface/chartsheet";
export * as DataValidation from "@excel/surface/data-validation";
export * as DefinedNames from "@excel/surface/defined-names";
export * as Note from "@excel/surface/note";
export * as Address from "@excel/surface/address";
export * as Anchor from "@excel/surface/anchor";
export * as Watermark from "@excel/surface/watermark";
export * as Stream from "@excel/surface/stream.browser";
