import type { HeaderFooterImageEntry, HeaderFooterImageOptions } from "@excel/core/worksheet";
import {
  getHeaderFooterImages,
  removeHeaderFooterImage,
  setHeaderFooterImage
} from "@excel/core/worksheet";
/**
 * `HeaderFooterImage` namespace surface — Excel's six printable
 * header/footer image sections.
 *
 * `import { HeaderFooterImage } from "documonster/excel"` →
 *   `HeaderFooterImage.set(ws, opts)`, `HeaderFooterImage.list(ws)`,
 *   `HeaderFooterImage.remove(ws, "CH")`.
 *
 * Use this namespace when a worksheet needs per-section control (including
 * images loaded from an existing workbook). `Watermark` remains the
 * single-value convenience API for the common "one centred header image"
 * case.
 */
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import {
  checkHeaderImage,
  checkHeaderPosition,
  checkImageId
} from "@excel/surface/argument-checks";
import type { HeaderFooterImagePosition } from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";

/** The images in the sheet's header and footer sections. */
export function list(ws: WorksheetData): HeaderFooterImageEntry[] {
  checkWorksheet(ws);
  return getHeaderFooterImages(ws);
}

/** Remove the image in one header/footer section. Returns `true` if there was one. */
export function remove(ws: WorksheetData, position: HeaderFooterImagePosition): boolean {
  checkWorksheet(ws);
  // A section that is not one (`"bogus"`) answered `false`, as if that section simply had no picture.
  checkHeaderPosition("HeaderFooterImage.remove", "position", position);
  return removeHeaderFooterImage(ws, position);
}

/**
 * Put an image in one of the six header/footer sections, replacing what was there.
 *
 * @throws {ExcelImageError} when the workbook has no image `imageId`.
 * @throws {ExcelArgumentError} when `position` or `applyTo` is outside its vocabulary, or a size is not a
 * non-negative number of points.
 */
export function set(ws: WorksheetData, options: HeaderFooterImageOptions): void {
  checkWorksheet(ws);
  const check = argumentChecks("HeaderFooterImage.set");
  check.required("options", options);
  if (check.record("options", options)) {
    checkImageId("HeaderFooterImage.set", ws, options.imageId);
    check.required("options.position", options.position);
    checkHeaderImage(
      check,
      "options",
      { position: "position", width: "width", height: "height", applyTo: "applyTo" },
      options
    );
  }
  setHeaderFooterImage(ws, options);
}

// The flat names are the public ones; these namespace aliases repeated them.
/** @deprecated Use the flat `HeaderFooterImageEntry` type. Removed in 1.0. */
export type Entry = HeaderFooterImageEntry;
/** @deprecated Use the flat `HeaderFooterImageOptions` type. Removed in 1.0. */
export type Options = HeaderFooterImageOptions;
/** @deprecated Use the flat `HeaderFooterImagePosition` type. Removed in 1.0. */
export type Position = HeaderFooterImagePosition;
