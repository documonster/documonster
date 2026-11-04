/**
 * `Watermark` namespace surface — worksheet background watermark.
 *
 * `import { Watermark } from "documonster/excel"` →
 *   `Watermark.add(ws, opts)`, `Watermark.get(ws)`, `Watermark.remove(ws)`.
 *
 * This is a single-value convenience API. `mode: "overlay"` is its primary use
 * case; for `mode: "header"` prefer the `HeaderFooterImage` namespace, which
 * addresses Excel's six header/footer sections independently and can manage
 * images loaded from an existing workbook.
 */
import { addWatermark, getWatermark } from "@excel/core/worksheet";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkWorksheet } from "@excel/core/worksheet-core";
import { checkHeaderImage, checkImageId } from "@excel/surface/argument-checks";
import type { WatermarkOptions } from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";
import { copyPlain } from "@excel/utils/copy-plain";

export { removeWatermark as remove } from "@excel/core/worksheet";

/**
 * A copy of the watermark {@link add} put on the sheet, or `null` when it has none. Editing the copy changes nothing:
 * replace the watermark with {@link add}. (It used to return the sheet's own record, whose edits showed here and never
 * reached the file.)
 */
export function get(ws: WorksheetData): WatermarkOptions | null {
  checkWorksheet(ws);
  const watermark = getWatermark(ws);
  return watermark === null || watermark === undefined ? null : copyPlain(watermark);
}

/**
 * Put a watermark on the sheet, replacing the one this API put there before.
 *
 * @throws {ImageError} when the workbook has no image `imageId` — which used to write a watermark with nothing to show.
 * @throws {ExcelArgumentError} when `mode`, `position` or `applyTo` is outside its vocabulary, `opacity` is not from 0
 * to 1, or a size is not a non-negative number.
 */
export function add(ws: WorksheetData, options: WatermarkOptions): void {
  const check = argumentChecks("Watermark.add");
  check.required("options", options);
  if (check.record("options", options)) {
    checkImageId("Watermark.add", ws, options.imageId);
    check.oneOf("options.mode", options.mode, ["overlay", "header"]);
    check.number("options.opacity", options.opacity, { min: 0, max: 1 });
    checkHeaderImage(
      check,
      "options",
      { position: "position", width: "headerWidth", height: "headerHeight", applyTo: "applyTo" },
      options
    );
  }
  addWatermark(ws, options);
}
