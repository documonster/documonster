/**
 * `Watermark` namespace surface — worksheet background watermark.
 *
 * `import { Watermark } from "documonster/excel"` →
 *   `Watermark.set(ws, opts)`, `Watermark.get(ws)`, `Watermark.remove(ws)`.
 *
 * This is a single-value convenience API. `mode: "overlay"` is its primary use
 * case; for `mode: "header"` prefer the `HeaderFooterImage` namespace, which
 * addresses Excel's six header/footer sections independently and can manage
 * images loaded from an existing workbook.
 *
 * Each also takes a streaming worksheet writer (`Stream.WorksheetWriter`), which supports `mode: "overlay"` only.
 */
import {
  addWatermark,
  getWatermark,
  removeWatermark as coreRemoveWatermark
} from "@excel/core/worksheet";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkWorksheet } from "@excel/core/worksheet-core";
import { checkImageId, watermarkOptionsSchema } from "@excel/surface/argument-checks";
import type { StreamWorksheet } from "@excel/surface/stream-targets";
import { asStreamWorksheet } from "@excel/surface/stream-targets";
import type { WatermarkOptions } from "@excel/types";
import { checkSchema } from "@excel/utils/arguments";
import { copyPlain } from "@excel/utils/copy-plain";

/**
 * A copy of the watermark {@link set} put on the sheet, or `undefined` when it has none. Editing the copy changes nothing:
 * replace the watermark with {@link set}. (It used to return the sheet's own record, whose edits showed here and never
 * reached the file.)
 */
export function get(ws: WorksheetData | StreamWorksheet): WatermarkOptions | undefined {
  const writer = asStreamWorksheet(ws);
  if (!writer) {
    checkWorksheet(ws);
  }
  const watermark = writer ? writer.getWatermark() : getWatermark(ws as WorksheetData);
  return watermark === null || watermark === undefined ? undefined : copyPlain(watermark);
}

/**
 * Put a watermark on the sheet, replacing the one this API put there before.
 *
 * @throws {ExcelImageError} when the workbook has no image `imageId` — which used to write a watermark with nothing to show.
 * @throws {ExcelArgumentError} when `mode`, `position` or `applyTo` is outside its vocabulary, `opacity` is not from 0
 * to 1, or a size is not a non-negative number.
 * @throws {ExcelImageError} for `mode: "header"` on a streaming worksheet writer, which cannot write header images.
 * @throws {ExcelStreamStateError} when a streamed sheet has been committed.
 */
export function set(ws: WorksheetData | StreamWorksheet, options: WatermarkOptions): void {
  setWatermark("Watermark.set", ws, options);
}

/** Remove the sheet's watermark — a streamed sheet's too, until it is committed. Returns `true` if it had one. */
export function remove(ws: WorksheetData | StreamWorksheet): boolean {
  const writer = asStreamWorksheet(ws);
  if (writer) {
    writer._checkOpen("remove a watermark");
    return writer.removeWatermark();
  }
  return coreRemoveWatermark(ws as WorksheetData);
}

/**
 * @deprecated Use {@link set}: a sheet has one watermark, and this replaces it rather than adding a second.
 * Removed in 1.0.
 */
export function add(ws: WorksheetData | StreamWorksheet, options: WatermarkOptions): void {
  setWatermark("Watermark.add", ws, options);
}

function setWatermark(
  where: string,
  ws: WorksheetData | StreamWorksheet,
  options: WatermarkOptions
): void {
  const writer = asStreamWorksheet(ws);
  checkSchema(where, "options", watermarkOptionsSchema, options);
  if (!writer) {
    checkImageId(where, ws as WorksheetData, options.imageId);
  }
  if (writer) {
    writer._setWatermark(where, options);
    return;
  }
  addWatermark(ws as WorksheetData, options);
}
