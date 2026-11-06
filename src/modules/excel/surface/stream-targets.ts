/**
 * The streaming writers as the document-model namespaces accept them.
 *
 * `Worksheet.merge`, `Image.place`, `Watermark.set`, `ConditionalFormatting.add`, `DataValidation.add`,
 * `Worksheet.protect` and the rest take a `Stream.WorksheetWriter` as well as a worksheet, and hand it to the writer's
 * own implementation, so its checks — a committed sheet or row refuses the edit — apply. What a writer cannot do is
 * refused by name with {@link refuseStreamWorksheet}, rather than reaching fields the writer does not keep.
 *
 * Detected by shape, not by `instanceof`: importing the writer classes here would pull the streaming writer into every
 * bundle that uses one of these namespaces.
 */
import { ExcelStreamStateError, ExcelUnsupportedFeatureError } from "@excel/errors";
import type {
  WorkbookWriterBase,
  WorksheetWriterLike
} from "@excel/stream/workbook-writer.browser";
import type { WorksheetWriter } from "@excel/stream/worksheet-writer";

// The public names, for signatures: each platform's `Stream` entry names its own writers.
export type {
  WorkbookWriter as StreamWorkbook,
  WorksheetWriter as StreamWorksheet
} from "@excel/surface/stream";

/** `ws` as a streaming worksheet writer, whatever workbook writer it belongs to, or `undefined` when it is not one. */
export function asStreamWorksheet(ws: unknown): WorksheetWriter<unknown> | undefined {
  // The writer is the only sheet that commits its own rows: the same test core's `streamingRows` makes.
  return typeof (ws as { _commitRow?: unknown } | null | undefined)?._commitRow === "function"
    ? (ws as WorksheetWriter<unknown>)
    : undefined;
}

/** `wb` as a streaming workbook writer, or `undefined` when it is not one. */
export function asStreamWorkbook(wb: unknown): WorkbookWriterBase<WorksheetWriterLike> | undefined {
  return typeof (wb as { _workbookCommitted?: unknown } | null | undefined)?._workbookCommitted ===
    "boolean"
    ? (wb as WorkbookWriterBase<WorksheetWriterLike>)
    : undefined;
}

/**
 * Refuse `operation` on a streaming worksheet writer, saying why. It used to run against the writer's fields as if it
 * were a worksheet, and either changed nothing in the file or failed with a `TypeError`.
 */
export function refuseStreamWorksheet(ws: unknown, operation: string, reason: string): void {
  if (asStreamWorksheet(ws) !== undefined) {
    throw new ExcelUnsupportedFeatureError(operation, `a streaming worksheet writer ${reason}`);
  }
}

/** Refuse `operation` on a streaming workbook writer, saying why — see {@link refuseStreamWorksheet}. */
export function refuseStreamWorkbook(wb: unknown, operation: string, reason: string): void {
  if (asStreamWorkbook(wb) !== undefined) {
    throw new ExcelUnsupportedFeatureError(operation, `a streaming workbook writer ${reason}`);
  }
}

/**
 * Refuse `where` on a streaming worksheet writer for the reason most refusals share: it keeps only the rows it has not
 * written yet, so it cannot `operation` (`"splice columns"`, `"read the cells of a column"`).
 */
export function refuseWrittenRows(ws: unknown, where: string, operation: string): void {
  refuseStreamWorksheet(
    ws,
    where,
    `keeps only the rows it has not written yet, so it cannot ${operation}`
  );
}

/**
 * `fn`, refusing a streaming workbook writer as its first argument with {@link refuseStreamWorkbook} before anything
 * else runs. For a namespace member that only a workbook in memory can serve; its signature, overloads included, is
 * `fn`'s.
 *
 * Call it as `/* @__PURE__ *\/ workbookOnly(…)`. It runs when the module loads, and a call a bundler cannot prove
 * free of effects keeps `fn` — and everything `fn` imports — in every bundle that loads the namespace, used or not:
 * without the annotation `Workbook.setModel` alone dragged the model loaders, the chart builders and the formula
 * parser into a write-only export (`scripts/treeshake-verify.ts`, issue #241).
 */
export function workbookOnly<F extends (wb: never, ...rest: never[]) => unknown>(
  operation: string,
  reason: string,
  fn: F
): F {
  const call = fn as unknown as (...args: unknown[]) => unknown;
  return ((...args: unknown[]) => {
    refuseStreamWorkbook(args[0], operation, reason);
    return call(...args);
  }) as unknown as F;
}

/**
 * Refuse the call `where` on rows `top`… of a streaming worksheet writer once any of them is written — with
 * `commitRow`, a later row's, or the sheet's own commit — as `Stream.setCell*` refuses an edit through a committed
 * row's handle.
 * Rows are written in order, so the first row of the span decides. Nothing is refused for a worksheet.
 */
export function checkStreamRowsOpen(ws: unknown, where: string, top: number): void {
  const writer = asStreamWorksheet(ws);
  if (writer === undefined) {
    return;
  }
  // `where` is the public call (`Cell.setValue`); the error's operation is a phrase ("Cannot use Cell.setValue: …"),
  // as the setters that name an action write theirs ("Cannot set a column width: …").
  const operation = `use ${where}`;
  writer._checkOpen(operation);
  if (writer.isRowCommitted(top)) {
    throw new ExcelStreamStateError(operation, `row ${top} has been committed`);
  }
}

/**
 * Refuse a read of rows `top`… of a streaming worksheet writer once any of them is written: the writer no longer holds
 * them, and answering as for an empty row would be wrong without saying so. Nothing is refused for a worksheet.
 */
export function checkStreamRowsHeld(ws: unknown, where: string, top: number): void {
  const writer = asStreamWorksheet(ws);
  if (writer !== undefined && writer.isRowCommitted(top)) {
    throw new ExcelStreamStateError(
      `use ${where}`,
      `row ${top} has been committed, and the streaming writer no longer holds it`
    );
  }
}
