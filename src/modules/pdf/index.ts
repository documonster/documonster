/**
 * PDF module for documonster.
 *
 * A full-featured, zero-dependency PDF engine for both writing and reading.
 *
 * @example Standalone PDF generation:
 * ```typescript
 * import { Pdf } from "documonster/pdf";
 *
 * const bytes = await Pdf.fromTable([
 *   ["Product", "Revenue"],
 *   ["Widget", 1000],
 *   ["Gadget", 2500]
 * ]);
 * ```
 *
 * @example From Excel Workbook:
 * ```typescript
 * import { Row, Workbook } from "documonster/excel";
 * import { Pdf } from "documonster/pdf";
 *
 * const workbook = Workbook.create();
 * const sheet = Workbook.addWorksheet(workbook, "Sales");
 * Row.add(sheet, ["Product", "Revenue"]);
 * const bytes = await Pdf.fromExcel(workbook);
 * ```
 *
 * @example Read PDF — extract text, images, and metadata:
 * ```typescript
 * import { Pdf } from "documonster/pdf";
 *
 * const result = await Pdf.extract(pdfBytes);
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

import type { DrawSurface } from "@draw/surface";
import { checkDrawPage } from "@pdf/core/arguments";
import type {
  CmsSignedData as CmsSignedDataModel,
  SignOptions as SignOptionsModel,
  SignatureVerificationResult as SignatureVerificationResultModel
} from "@pdf/core/digital-signature";
import type { DrawSurfaceRect, PdfDrawPage } from "@pdf/render/draw-surface";
import { createScaledPdfDrawSurface } from "@pdf/render/draw-surface";

// =============================================================================
// Public API — the `Pdf` domain namespace (tree-shaken via `export * as`)
// =============================================================================

export * as Pdf from "@pdf/surface/pdf";

// Conversion option types (the converter functions live on `Pdf.*`).
export type { ChartToPdfOptions, ExcelToPdfOptions } from "@pdf/excel-bridge";
export type { DocxToPdfOptions } from "@pdf/word-bridge";

// =============================================================================
// Types — Writing
// =============================================================================

export type {
  PdfCell,
  PdfCellValue,
  PdfRow,
  PdfColumn,
  PdfSheet,
  PdfBook,
  PdfImage,
  PdfInput
} from "@pdf/pdf";

// `textLanguage` appears in five public option types and on the builder, so the
// type naming its values has to be nameable too: a caller could pass the literal
// but not store it in a variable, write a wrapper signature, or build a config
// object.
export type { CjkLanguage } from "@utils/cjk";

export type {
  PdfDocumentOptions,
  PdfExportOptions,
  PdfOrientation,
  PdfPageOrder,
  PdfCellErrorMode,
  PdfPageSize,
  PdfMargins,
  PdfColor,
  PageSizeName,
  PdfWatermark,
  PdfTextWatermark,
  PdfImageWatermark,
  PdfWatermarkFilter,
  PdfHeaderFooterOptions
} from "@pdf/types";
export type {
  PdfFontSource,
  PdfFontFaces,
  PdfNamedFontFamily,
  PdfFontConfig
} from "@pdf/font/font-config";

// =============================================================================
// Types — Reading
// =============================================================================

export type { ReadPdfOptions, ReadPdfResult, ReadPdfPage } from "@pdf/reader/pdf-reader";
export type { PdfMetadata } from "@pdf/reader/metadata-reader";
export type { ExtractedImage } from "@pdf/reader/image-extractor";
export type { TextLine } from "@pdf/reader/text-reconstruction";
export type { TextFragment } from "@pdf/reader/content-interpreter";
export type { PdfAnnotation, PdfRect } from "@pdf/reader/annotation-extractor";
export type { PdfFormField, PdfFormFieldType } from "@pdf/reader/form-extractor";
export type { PdfBookmark } from "@pdf/reader/bookmark-extractor";
export type { PdfTable, PdfTableRow, PdfTableCell } from "@pdf/reader/table-extractor";

// =============================================================================
// Types — Building
// =============================================================================

export type {
  PdfDocumentBuilderOptions,
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
} from "@pdf/builder/document-builder";
export type {
  EditorReadOptions,
  EditorCopyPagesOptions,
  PdfMetadataUpdate,
  EditorSplitOptions,
  EditorWriteOptions,
  EditorToBufferOptions,
  LoadOptions,
  EditorSaveOptions,
  EditorSignOptions
} from "@pdf/builder/pdf-editor";
/** What `page.getContentStream()` returns. */
export type { PdfContentStream } from "@pdf/core/pdf-stream";

// =============================================================================
// Types — Digital Signatures
// =============================================================================

export type { PdfSignatureCheck, SignPdfOptions, Asn1Node } from "@pdf/core/digital-signature";
/**
 * @deprecated What the deprecated `Pdf.verifySignature` returns; `Pdf.Signature.verify`
 * returns {@link PdfSignatureCheck}. Removed in 1.0.
 */
export type SignatureVerificationResult = SignatureVerificationResultModel;
/** @deprecated No public function takes or returns this. Removed in 1.0. */
export type CmsSignedData = CmsSignedDataModel;
/**
 * @deprecated No public function takes this; signing options are {@link SignPdfOptions}
 * (`Pdf.Signature.fillPlaceholder`) and `PdfSignatureOptions` (`sign`). Removed in 1.0.
 */
export type SignOptions = SignOptionsModel;
export type { PdfSignatureReport, VerifySignaturesOptions } from "@pdf/reader/signature-verifier";

// =============================================================================
// Types — Drawing (`documonster/draw` backend)
// =============================================================================

// `Pdf.createDrawSurface` is the PDF backend for `documonster/draw`; these name its arguments.
export type {
  DrawSurfaceRect,
  PdfClipTarget,
  PdfDrawPage,
  PdfDrawSurfaceOptions
} from "@pdf/render/draw-surface";

/**
 * @deprecated Use `Pdf.createDrawSurface(page, rect, { scale })`. Removed in 1.0.
 */
export function createPdfDrawSurface(
  page: PdfDrawPage,
  rect: DrawSurfaceRect,
  scale?: number
): DrawSurface {
  checkDrawPage("createPdfDrawSurface", page);
  return createScaledPdfDrawSurface(page, rect, scale);
}

// =============================================================================
// Errors
// =============================================================================

export {
  PdfError,
  PdfRenderError,
  PdfFontError,
  PdfStructureError,
  PdfPasswordError,
  PdfLimitExceededError,
  PdfSignatureInvalidationError,
  PdfSignaturePlaceholderMissingError,
  PdfIncrementalUpdateError,
  PdfEncryptionError,
  PdfXfaFormError,
  PdfArgumentError,
  isPdfError
} from "@pdf/errors";
export type { PdfLimit } from "@pdf/errors";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
