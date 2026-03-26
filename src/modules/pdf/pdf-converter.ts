/**
 * PDF Converter for ExcelTS workbooks.
 *
 * Converts Excel Workbook/Worksheet data to PDF format preserving:
 * - Cell values and formulas (results)
 * - Font styles (family, size, bold, italic, underline, strikethrough, color)
 * - Cell fills (solid pattern fills)
 * - Borders (all styles, colors)
 * - Alignment (horizontal, vertical, text wrap)
 * - Merged cells
 * - Images (JPEG, PNG)
 * - Page layout (orientation, margins, paper size)
 *
 * Zero external dependencies – generates PDF 1.4 binary directly.
 */

import { PdfWriter, type PdfPageDef, type PdfImageData } from "@pdf/pdf-writer";
import { layoutWorksheet, type PdfLayoutOptions, type LayoutResult } from "@pdf/layout-engine";
import { PdfConversionError } from "@excel/errors";
import type { Worksheet } from "@excel/worksheet";
import type { WorkbookMedia } from "@excel/workbook.browser";
import { PaperSize } from "@excel/types";

// =============================================================================
// Public Types
// =============================================================================

export interface ToPdfOptions {
  /** Page orientation. Defaults to worksheet pageSetup or 'portrait'. */
  orientation?: "portrait" | "landscape";

  /** Paper size. Defaults to A4. */
  paperSize?: "a4" | "letter" | "legal" | "a3" | "a5";

  /** Custom page dimensions in points (overrides paperSize). */
  pageWidth?: number;
  pageHeight?: number;

  /** Page margins in points. */
  margins?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
  };

  /** Whether to display grid lines on the PDF. Default: true. */
  gridLines?: boolean;

  /** Scale factor (0.1 to 3.0). Default: 1.0. */
  scale?: number;

  /** Fit all columns to page width. Default: false. */
  fitToWidth?: boolean;

  /** Specific worksheet indices to include (0-based). Omit for all worksheets. */
  worksheets?: number[];

  /** Custom column widths override (1-based col index → width in Excel units). */
  columnWidths?: Record<number, number>;

  /** Whether to include images. Default: true. */
  includeImages?: boolean;

  /** Creator metadata string for PDF. */
  creator?: string;
}

// =============================================================================
// Paper size dimensions (in points: 1 inch = 72 points)
// =============================================================================

const PAPER_SIZES: Record<string, { width: number; height: number }> = {
  a3: { width: 841.89, height: 1190.55 },
  a4: { width: 595.28, height: 841.89 },
  a5: { width: 419.53, height: 595.28 },
  letter: { width: 612, height: 792 },
  legal: { width: 612, height: 1008 }
};

/**
 * Map the PaperSize enum (numeric) to paper size dimensions in points.
 * The PaperSize enum values come from the OOXML spec.
 */
const PAPER_SIZE_ENUM_MAP: Record<number, { width: number; height: number }> = {
  [PaperSize.Legal]: PAPER_SIZES.legal,
  [PaperSize.Executive]: { width: 522, height: 756 },
  [PaperSize.A4]: PAPER_SIZES.a4,
  [PaperSize.A5]: PAPER_SIZES.a5,
  [PaperSize.B5]: { width: 498.9, height: 708.66 }
};

// =============================================================================
// Main conversion function
// =============================================================================

/**
 * Convert an array of worksheets to a PDF Uint8Array.
 *
 * @param worksheets - The worksheets to convert
 * @param media - The workbook's media array (images)
 * @param options - PDF generation options
 * @returns PDF file as Uint8Array
 */
export function worksheetsToPdf(
  worksheets: Worksheet[],
  media: WorkbookMedia[],
  options: ToPdfOptions = {}
): Uint8Array {
  if (!worksheets || worksheets.length === 0) {
    throw new PdfConversionError("No worksheets to convert");
  }

  // Determine page size
  let pageWidth: number;
  let pageHeight: number;

  if (options.pageWidth && options.pageHeight) {
    pageWidth = options.pageWidth;
    pageHeight = options.pageHeight;
  } else {
    const paper = PAPER_SIZES[options.paperSize ?? "a4"] ?? PAPER_SIZES.a4;
    pageWidth = paper.width;
    pageHeight = paper.height;
  }

  const allPages: PdfPageDef[] = [];
  const allImages: PdfImageData[] = [];
  const seenImageKeys = new Set<string>();

  // Filter worksheets if specified
  let sheetsToConvert = worksheets;
  if (options.worksheets) {
    sheetsToConvert = options.worksheets
      .filter(i => i >= 0 && i < worksheets.length)
      .map(i => worksheets[i]);
    if (sheetsToConvert.length === 0) {
      throw new PdfConversionError("No valid worksheet indices specified");
    }
  }

  for (const ws of sheetsToConvert) {
    // Per-worksheet: resolve page size from worksheet's pageSetup if not overridden
    let wsPageWidth = pageWidth;
    let wsPageHeight = pageHeight;

    if (!options.pageWidth && !options.pageHeight && !options.paperSize) {
      const wsPaperSize = ws.pageSetup?.paperSize;
      if (wsPaperSize !== undefined) {
        const mapped = PAPER_SIZE_ENUM_MAP[wsPaperSize];
        if (mapped) {
          wsPageWidth = mapped.width;
          wsPageHeight = mapped.height;
        }
      }
    }

    // Resolve orientation: explicit option > worksheet pageSetup > portrait
    const wsOrientation =
      options.orientation ?? (ws.pageSetup?.orientation as "portrait" | "landscape") ?? "portrait";

    // Resolve scale:
    // - User-provided scale → apply directly to content dimensions.
    // - Worksheet pageSetup.scale (print zoom %) without a user-specified page
    //   size → enlarge the effective page by 1/printScale so content stays at
    //   natural size (matches Excel's "Print to PDF" behaviour).
    let wsScale: number;
    if (options.scale !== undefined) {
      wsScale = Math.max(0.1, Math.min(3, options.scale));
    } else if (
      !options.pageWidth &&
      !options.pageHeight &&
      !options.paperSize &&
      ws.pageSetup?.scale !== undefined &&
      ws.pageSetup.scale > 0 &&
      ws.pageSetup.scale !== 100
    ) {
      const printScale = Math.max(10, Math.min(400, ws.pageSetup.scale)) / 100;
      wsPageWidth /= printScale;
      wsPageHeight /= printScale;
      wsScale = 1;
    } else {
      wsScale = 1;
    }

    // Resolve margins: explicit option > worksheet pageSetup.margins (inches → points) > defaults
    const wsMargins =
      options.margins ??
      (ws.pageSetup?.margins
        ? {
            top: ws.pageSetup.margins.top * 72,
            right: ws.pageSetup.margins.right * 72,
            bottom: ws.pageSetup.margins.bottom * 72,
            left: ws.pageSetup.margins.left * 72
          }
        : undefined);

    const layoutOptions: PdfLayoutOptions = {
      pageWidth: wsPageWidth,
      pageHeight: wsPageHeight,
      margins: wsMargins,
      gridLines: options.gridLines,
      scale: wsScale,
      orientation: wsOrientation,
      fitToWidth: options.fitToWidth,
      columnWidths: options.columnWidths
    };

    const result: LayoutResult = layoutWorksheet(
      ws,
      options.includeImages === false ? [] : media,
      layoutOptions
    );

    allPages.push(...result.pages);

    // Deduplicate images
    for (const img of result.images) {
      if (!seenImageKeys.has(img.key)) {
        seenImageKeys.add(img.key);
        allImages.push(img);
      }
    }
  }

  // Generate PDF binary
  const writer = new PdfWriter();
  return writer.build(allPages, allImages);
}
