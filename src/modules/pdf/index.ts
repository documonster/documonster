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
import type {
  PageOptions as PdfPageOptions,
  DrawSvgOptions as PdfDrawSvgOptions,
  DrawTextOptions as PdfDrawTextOptions,
  DrawRectOptions as PdfDrawRectOptions,
  DrawCircleOptions as PdfDrawCircleOptions,
  DrawEllipseOptions as PdfDrawEllipseOptions,
  DrawLineOptions as PdfDrawLineOptions,
  DrawPathOptions as PdfDrawPathOptions,
  DrawImageOptions as PdfDrawImageOptions,
  DocumentMetadata as PdfDocumentMetadata,
  PathOp as PdfPathOp,
  TocOptions as PdfTocOptions,
  AnnotationType as PdfAnnotationType,
  AnnotationOptions as PdfAnnotationOptions,
  TextMarkupAnnotationOptions as PdfTextMarkupAnnotationOptions,
  TextAnnotationOptions as PdfTextAnnotationOptions,
  FreeTextAnnotationOptions as PdfFreeTextAnnotationOptions,
  StampAnnotationOptions as PdfStampAnnotationOptions,
  FormFieldOptions as PdfFormFieldOptions,
  TextFieldOptions as PdfTextFieldOptions,
  CheckboxOptions as PdfCheckboxOptions,
  DropdownOptions as PdfDropdownOptions,
  RadioGroupOptions as PdfRadioGroupOptions
} from "@pdf/builder/document-builder";
import type {
  EditorReadOptions as PdfEditorReadOptions,
  EditorCopyPagesOptions as PdfEditorCopyPagesOptions,
  EditorSplitOptions as PdfEditorSplitOptions,
  EditorWriteOptions as PdfEditorWriteOptions,
  EditorToBufferOptions as PdfEditorToBufferOptions,
  EditorSignOptions as PdfEditorSignOptions
} from "@pdf/builder/pdf-editor";
import { checkDrawPage } from "@pdf/core/arguments";
import type {
  Asn1Node as PdfAsn1Node,
  CmsSignedData as CmsSignedDataModel,
  SignOptions as SignOptionsModel,
  SignPdfOptions as PdfSignOptions,
  SignatureVerificationResult as SignatureVerificationResultModel
} from "@pdf/core/digital-signature";
import type {
  ChartToPdfOptions as PdfFromChartOptions,
  ExcelToPdfOptions as PdfFromExcelOptions
} from "@pdf/excel-bridge";
import type { PdfAnnotationRect as PdfAnnotationRectType } from "@pdf/reader/annotation-extractor";
import type { TextFragment as PdfTextFragment } from "@pdf/reader/content-interpreter";
import type { ExtractedImage as PdfExtractedImage } from "@pdf/reader/image-extractor";
import type {
  ReadPdfOptions as PdfExtractOptions,
  ReadPdfResult as PdfExtractResult,
  ReadPdfPage as PdfExtractedPage
} from "@pdf/reader/pdf-reader";
import type { VerifySignaturesOptions as PdfVerifySignaturesOptions } from "@pdf/reader/signature-verifier";
import type { TextLine as PdfTextLine } from "@pdf/reader/text-reconstruction";
import type { DrawSurfaceRect as PdfDrawSurfaceRect, PdfDrawPage } from "@pdf/render/draw-surface";
import { createScaledPdfDrawSurface } from "@pdf/render/draw-surface";
import type { PageSizeName as PdfPageSizeName } from "@pdf/types";
import type { DocxToPdfOptions as PdfFromDocxOptions } from "@pdf/word-bridge";
import type { CjkLanguage as PdfCjkLanguage } from "@utils/cjk";

// =============================================================================
// Public API — the `Pdf` domain namespace (tree-shaken via `export * as`)
// =============================================================================

export * as Pdf from "@pdf/surface/pdf";

// Conversion option types (the converter functions live on `Pdf.*`).
export type { PdfFromChartOptions, PdfFromExcelOptions, PdfFromDocxOptions };

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
export type { PdfCjkLanguage };

export type {
  PdfDocumentOptions,
  PdfExportOptions,
  PdfOrientation,
  PdfPageOrder,
  PdfCellErrorMode,
  PdfCellCommentMode,
  PdfEncryptionOptions,
  PdfPermissions,
  PdfPageSize,
  PdfMargins,
  PdfColor,
  PdfWatermark,
  PdfTextWatermark,
  PdfImageWatermark,
  PdfWatermarkFilter,
  PdfHeaderFooterOptions
} from "@pdf/types";
export type { PdfPageSizeName };
export type {
  PdfFontSource,
  PdfFontFaces,
  PdfNamedFontFamily,
  PdfFontConfig
} from "@pdf/font/font-config";

// =============================================================================
// Types — Reading
// =============================================================================

export type {
  PdfExtractOptions,
  PdfExtractResult,
  PdfExtractedPage,
  PdfExtractedImage,
  PdfTextLine,
  PdfTextFragment
};
export type { PdfMetadata } from "@pdf/reader/metadata-reader";
export type { PdfAnnotation, PdfAnnotationRect } from "@pdf/reader/annotation-extractor";
/**
 * @deprecated Use {@link PdfAnnotationRect}. A rectangle by its corners, as an annotation's
 * `/Rect` gives it; the general name read as the `{ x, y, width, height }` shape. Removed in 1.0.
 */
export type PdfRect = PdfAnnotationRectType;
export type { PdfFormField, PdfFormFieldType } from "@pdf/reader/form-extractor";
export type { PdfBookmark } from "@pdf/reader/bookmark-extractor";
export type { PdfTable, PdfTableRow, PdfTableCell } from "@pdf/reader/table-extractor";

// =============================================================================
// Types — Building
// =============================================================================

export type { PdfDocumentBuilderOptions, PdfSignatureOptions } from "@pdf/builder/document-builder";
export type {
  PdfPageOptions,
  PdfDrawSvgOptions,
  PdfDrawTextOptions,
  PdfDrawRectOptions,
  PdfDrawCircleOptions,
  PdfDrawEllipseOptions,
  PdfDrawLineOptions,
  PdfDrawPathOptions,
  PdfDrawImageOptions,
  PdfDocumentMetadata,
  PdfPathOp,
  PdfTocOptions,
  PdfAnnotationType,
  PdfAnnotationOptions,
  PdfTextMarkupAnnotationOptions,
  PdfTextAnnotationOptions,
  PdfFreeTextAnnotationOptions,
  PdfStampAnnotationOptions,
  PdfFormFieldOptions,
  PdfTextFieldOptions,
  PdfCheckboxOptions,
  PdfDropdownOptions,
  PdfRadioGroupOptions
};
export type {
  PdfEditorReadOptions,
  PdfEditorCopyPagesOptions,
  PdfEditorSplitOptions,
  PdfEditorWriteOptions,
  PdfEditorToBufferOptions,
  PdfEditorSignOptions
};
export type { PdfMetadataUpdate } from "@pdf/builder/pdf-editor";
/** What `page.getContentStream()` returns. */
export type { PdfContentStream } from "@pdf/core/pdf-stream";

// =============================================================================
// Types — Digital Signatures
// =============================================================================

export type { PdfSignatureCheck } from "@pdf/core/digital-signature";
export type { PdfSignOptions, PdfAsn1Node };
/**
 * @deprecated What the deprecated `Pdf.verifySignature` returns; `Pdf.Signature.verify`
 * returns {@link PdfSignatureCheck}. Removed in 1.0.
 */
export type SignatureVerificationResult = SignatureVerificationResultModel;
/** @deprecated No public function takes or returns this. Removed in 1.0. */
export type CmsSignedData = CmsSignedDataModel;
/**
 * @deprecated No public function takes this; signing options are {@link PdfSignOptions}
 * (`Pdf.Signature.fillPlaceholder`) and `PdfSignatureOptions` (`sign`). Removed in 1.0.
 */
export type SignOptions = SignOptionsModel;
export type { PdfSignatureReport } from "@pdf/reader/signature-verifier";
export type { PdfVerifySignaturesOptions };

// =============================================================================
// Types — Drawing (`documonster/draw` backend)
// =============================================================================

// `Pdf.createDrawSurface` is the PDF backend for `documonster/draw`; these name its arguments.
export type { PdfDrawSurfaceRect };
export type { PdfClipTarget, PdfDrawPage, PdfDrawSurfaceOptions } from "@pdf/render/draw-surface";

/**
 * @deprecated Use `Pdf.createDrawSurface(page, rect, { scale })`. Removed in 1.0.
 */
export function createPdfDrawSurface(
  page: PdfDrawPage,
  rect: PdfDrawSurfaceRect,
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
  PdfUnsupportedFeatureError,
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

// =============================================================================
// Deprecated unprefixed type names — every flat type of this entry carries `Pdf`
// =============================================================================

/** @deprecated Use {@link PdfFromChartOptions}. Removed in 1.0. */
export type ChartToPdfOptions = PdfFromChartOptions;
/** @deprecated Use {@link PdfFromExcelOptions}. Removed in 1.0. */
export type ExcelToPdfOptions = PdfFromExcelOptions;
/** @deprecated Use {@link PdfFromDocxOptions}. Removed in 1.0. */
export type DocxToPdfOptions = PdfFromDocxOptions;
/** @deprecated Use {@link PdfCjkLanguage}. Removed in 1.0. */
export type CjkLanguage = PdfCjkLanguage;
/** @deprecated Use {@link PdfPageSizeName}. Removed in 1.0. */
export type PageSizeName = PdfPageSizeName;
/** @deprecated Use {@link PdfExtractOptions}. Removed in 1.0. */
export type ReadPdfOptions = PdfExtractOptions;
/** @deprecated Use {@link PdfExtractResult}. Removed in 1.0. */
export type ReadPdfResult = PdfExtractResult;
/** @deprecated Use {@link PdfExtractedPage}. Removed in 1.0. */
export type ReadPdfPage = PdfExtractedPage;
/** @deprecated Use {@link PdfExtractedImage}. Removed in 1.0. */
export type ExtractedImage = PdfExtractedImage;
/** @deprecated Use {@link PdfTextLine}. Removed in 1.0. */
export type TextLine = PdfTextLine;
/** @deprecated Use {@link PdfTextFragment}. Removed in 1.0. */
export type TextFragment = PdfTextFragment;
/** @deprecated Use {@link PdfPageOptions}. Removed in 1.0. */
export type PageOptions = PdfPageOptions;
/** @deprecated Use {@link PdfDrawSvgOptions}. Removed in 1.0. */
export type DrawSvgOptions = PdfDrawSvgOptions;
/** @deprecated Use {@link PdfDrawTextOptions}. Removed in 1.0. */
export type DrawTextOptions = PdfDrawTextOptions;
/** @deprecated Use {@link PdfDrawRectOptions}. Removed in 1.0. */
export type DrawRectOptions = PdfDrawRectOptions;
/** @deprecated Use {@link PdfDrawCircleOptions}. Removed in 1.0. */
export type DrawCircleOptions = PdfDrawCircleOptions;
/** @deprecated Use {@link PdfDrawEllipseOptions}. Removed in 1.0. */
export type DrawEllipseOptions = PdfDrawEllipseOptions;
/** @deprecated Use {@link PdfDrawLineOptions}. Removed in 1.0. */
export type DrawLineOptions = PdfDrawLineOptions;
/** @deprecated Use {@link PdfDrawPathOptions}. Removed in 1.0. */
export type DrawPathOptions = PdfDrawPathOptions;
/** @deprecated Use {@link PdfDrawImageOptions}. Removed in 1.0. */
export type DrawImageOptions = PdfDrawImageOptions;
/** @deprecated Use {@link PdfDocumentMetadata}. Removed in 1.0. */
export type DocumentMetadata = PdfDocumentMetadata;
/** @deprecated Use {@link PdfPathOp}. Removed in 1.0. */
export type PathOp = PdfPathOp;
/** @deprecated Use {@link PdfTocOptions}. Removed in 1.0. */
export type TocOptions = PdfTocOptions;
/** @deprecated Use {@link PdfAnnotationType}. Removed in 1.0. */
export type AnnotationType = PdfAnnotationType;
/** @deprecated Use {@link PdfAnnotationOptions}. Removed in 1.0. */
export type AnnotationOptions = PdfAnnotationOptions;
/** @deprecated Use {@link PdfTextMarkupAnnotationOptions}. Removed in 1.0. */
export type TextMarkupAnnotationOptions = PdfTextMarkupAnnotationOptions;
/** @deprecated Use {@link PdfTextAnnotationOptions}. Removed in 1.0. */
export type TextAnnotationOptions = PdfTextAnnotationOptions;
/** @deprecated Use {@link PdfFreeTextAnnotationOptions}. Removed in 1.0. */
export type FreeTextAnnotationOptions = PdfFreeTextAnnotationOptions;
/** @deprecated Use {@link PdfStampAnnotationOptions}. Removed in 1.0. */
export type StampAnnotationOptions = PdfStampAnnotationOptions;
/** @deprecated Use {@link PdfFormFieldOptions}. Removed in 1.0. */
export type FormFieldOptions = PdfFormFieldOptions;
/** @deprecated Use {@link PdfTextFieldOptions}. Removed in 1.0. */
export type TextFieldOptions = PdfTextFieldOptions;
/** @deprecated Use {@link PdfCheckboxOptions}. Removed in 1.0. */
export type CheckboxOptions = PdfCheckboxOptions;
/** @deprecated Use {@link PdfDropdownOptions}. Removed in 1.0. */
export type DropdownOptions = PdfDropdownOptions;
/** @deprecated Use {@link PdfRadioGroupOptions}. Removed in 1.0. */
export type RadioGroupOptions = PdfRadioGroupOptions;
/** @deprecated Use {@link PdfEditorReadOptions}. Removed in 1.0. */
export type EditorReadOptions = PdfEditorReadOptions;
/** @deprecated Use {@link PdfEditorCopyPagesOptions}. Removed in 1.0. */
export type EditorCopyPagesOptions = PdfEditorCopyPagesOptions;
/** @deprecated Use {@link PdfEditorSplitOptions}. Removed in 1.0. */
export type EditorSplitOptions = PdfEditorSplitOptions;
/** @deprecated Use {@link PdfEditorWriteOptions}. Removed in 1.0. */
export type EditorWriteOptions = PdfEditorWriteOptions;
/** @deprecated Use {@link PdfEditorToBufferOptions}. Removed in 1.0. */
export type EditorToBufferOptions = PdfEditorToBufferOptions;
/** @deprecated Use {@link PdfEditorSignOptions}. Removed in 1.0. */
export type EditorSignOptions = PdfEditorSignOptions;
/** @deprecated Use {@link PdfEditorReadOptions}. Removed in 1.0. */
export type LoadOptions = PdfEditorReadOptions;
/**
 * @deprecated Use {@link PdfEditorToBufferOptions} with `toBuffer()`, or
 * {@link PdfEditorWriteOptions} for what every write shares. Removed in 1.0.
 */
export type EditorSaveOptions = PdfEditorWriteOptions;
/** @deprecated Use {@link PdfSignOptions}. Removed in 1.0. */
export type SignPdfOptions = PdfSignOptions;
/** @deprecated Use {@link PdfAsn1Node}. Removed in 1.0. */
export type Asn1Node = PdfAsn1Node;
/** @deprecated Use {@link PdfVerifySignaturesOptions}. Removed in 1.0. */
export type VerifySignaturesOptions = PdfVerifySignaturesOptions;
/** @deprecated Use {@link PdfDrawSurfaceRect}. Removed in 1.0. */
export type DrawSurfaceRect = PdfDrawSurfaceRect;

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
