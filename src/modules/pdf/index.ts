/**
 * PDF module for excelts.
 *
 * A full-featured, zero-dependency PDF engine for both writing and reading.
 *
 * @example Standalone PDF generation:
 * ```typescript
 * import { pdf } from "excelts/pdf";
 *
 * const bytes = await pdf([
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
 * const bytes = await excelToPdf(workbook);
 * ```
 *
 * @example Read PDF — extract text, images, and metadata:
 * ```typescript
 * import { readPdf } from "excelts/pdf";
 *
 * const result = await readPdf(pdfBytes);
 * console.log(result.text);               // All text
 * console.log(result.pages[0].text);      // Page 1 text
 * console.log(result.pages[0].images);    // Page 1 images
 * console.log(result.pages[0].annotations); // Page 1 annotations
 * console.log(result.metadata.title);     // Document title
 * console.log(result.formFields);         // Form fields
 * ```
 *
 * @module pdf
 */

// =============================================================================
// Public API — the `Pdf` domain namespace (tree-shaken via `export * as`)
// =============================================================================

export * as Pdf from "./surface/pdf";

// Conversion option types (the converter functions live on `Pdf.*`).
export type { ChartToPdfOptions } from "./excel-bridge";
export type { DocxToPdfOptions } from "./word-bridge";

// =============================================================================
// Types — Writing
// =============================================================================

export type { PdfCell, PdfRow, PdfColumn, PdfSheet, PdfBook, PdfImage } from "./pdf";

export type {
  PdfExportOptions,
  PdfOrientation,
  PdfPageSize,
  PdfMargins,
  PdfColor,
  PageSizeName,
  PdfWatermark,
  PdfTextWatermark,
  PdfImageWatermark,
  PdfWatermarkFilter
} from "./types";

// =============================================================================
// Types — Reading
// =============================================================================

export type { ReadPdfOptions, ReadPdfResult, ReadPdfPage } from "./reader/pdf-reader";
export type { PdfMetadata } from "./reader/metadata-reader";
export type { ExtractedImage } from "./reader/image-extractor";
export type { TextLine } from "./reader/text-reconstruction";
export type { PdfAnnotation, PdfRect } from "./reader/annotation-extractor";
export type { PdfFormField, PdfFormFieldType } from "./reader/form-extractor";
export type { PdfBookmark } from "./reader/bookmark-extractor";
export type { PdfTable, PdfTableRow, PdfTableCell } from "./reader/table-extractor";

// =============================================================================
// Types — Building
// =============================================================================

export type {
  PageOptions,
  DrawSvgOptions,
  DrawTextOptions,
  DrawRectOptions,
  DrawCircleOptions,
  DrawEllipseOptions,
  DrawLineOptions,
  DrawPathOptions,
  DrawImageOptions,
  DocumentMetadata,
  PathOp,
  TocOptions,
  AnnotationType,
  AnnotationOptions,
  TextMarkupAnnotationOptions,
  TextAnnotationOptions,
  FreeTextAnnotationOptions,
  StampAnnotationOptions,
  FormFieldOptions,
  TextFieldOptions,
  CheckboxOptions,
  DropdownOptions,
  RadioGroupOptions,
  PdfSignatureOptions
} from "./builder/document-builder";
export type { LoadOptions } from "./builder/pdf-editor";

// =============================================================================
// Types — Digital Signatures
// =============================================================================

export type {
  SignatureVerificationResult,
  CmsSignedData,
  SignOptions,
  Asn1Node
} from "./core/digital-signature";

// =============================================================================
// Errors
// =============================================================================

export { PdfError, PdfRenderError, PdfFontError, PdfStructureError, isPdfError } from "./errors";
