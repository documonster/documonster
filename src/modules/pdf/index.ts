/**
 * PDF module for excelts.
 *
 * Provides Excel-to-PDF conversion with zero external dependencies.
 * Supports cell values, fonts, colors, borders, fills, alignment,
 * merged cells, pagination, and customizable page layout.
 *
 * @example
 * ```typescript
 * import { Workbook, PdfExporter } from "excelts";
 *
 * const workbook = new Workbook();
 * const sheet = workbook.addWorksheet("Sales");
 * sheet.columns = [
 *   { header: "Product", key: "product", width: 20 },
 *   { header: "Revenue", key: "revenue", width: 15 }
 * ];
 * sheet.addRow({ product: "Widget", revenue: 1000 });
 *
 * const exporter = new PdfExporter(workbook);
 * const pdfBuffer = exporter.export({
 *   pageSize: "A4",
 *   orientation: "portrait",
 *   fitToPage: true,
 *   showGridLines: true,
 *   showPageNumbers: true
 * });
 *
 * // Write to file (Node.js)
 * import { writeFileSync } from "fs";
 * writeFileSync("output.pdf", pdfBuffer);
 * ```
 *
 * @module pdf
 */

// =============================================================================
// Core Export
// =============================================================================

export { PdfExporter, exportPdf } from "./render/pdf-exporter";

// =============================================================================
// Types
// =============================================================================

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
