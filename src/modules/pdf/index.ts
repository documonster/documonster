/**
 * PDF module for excelts.
 *
 * A full-featured, zero-dependency PDF engine.
 *
 * @example Standalone:
 * ```typescript
 * import { pdf } from "excelts/pdf";
 *
 * const bytes = pdf([
 *   ["Product", "Revenue"],
 *   ["Widget", 1000],
 *   ["Gadget", 2500]
 * ]);
 * ```
 *
 * @example From Excel Workbook:
 * ```typescript
 * import { Workbook } from "excelts";
 * import { excelToPdf } from "excelts/pdf";
 *
 * const workbook = new Workbook();
 * const sheet = workbook.addWorksheet("Sales");
 * sheet.addRow(["Product", "Revenue"]);
 * const bytes = excelToPdf(workbook);
 * ```
 *
 * @module pdf
 */

// =============================================================================
// Public API
// =============================================================================

/** Standalone PDF generation — accepts plain arrays, sheet objects, or workbooks. */
export { pdf } from "./pdf";

/** Excel-to-PDF conversion — accepts an Excel Workbook instance. */
export { excelToPdf } from "./excel-bridge";

// =============================================================================
// Types
// =============================================================================

export type { PdfCell, PdfRow, PdfColumn, PdfSheet, PdfBook, PdfImage } from "./pdf";

export type {
  PdfExportOptions,
  PdfOrientation,
  PdfPageSize,
  PdfMargins,
  PdfColor,
  PageSizeName
} from "./types";

export { PageSizes } from "./types";

// =============================================================================
// Errors
// =============================================================================

export { PdfError, PdfRenderError, PdfFontError, PdfStructureError, isPdfError } from "./errors";
