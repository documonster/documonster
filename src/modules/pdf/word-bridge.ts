/**
 * Word-to-PDF Bridge
 *
 * Converts a Word document (`DocxDocument`) to PDF.
 *
 * The bridge is a thin translation layer:
 *
 *   DocxDocument
 *      │
 *      │  layoutDocumentFull()  ← @word/layout
 *      ▼
 *   LayoutDocument (positioned PageContent variants)
 *      │
 *      │  renderLayoutDocumentToPdf()  ← ./render-layout-to-pdf
 *      ▼
 *   PdfDocumentBuilder → bytes
 *
 * Every flow decision (line wrapping, page breaks, table sizing,
 * float positioning) lives in `@word/layout`. This file only handles
 * option mapping, optional chart-renderer auto-detection, and the
 * final `builder.build()` serialization.
 *
 * Like `excel-bridge.ts`, this is the ONLY file in the PDF module that
 * imports from `@word`.
 *
 * @example
 * ```typescript
 * import { readDocx } from "excelts/word";
 * import { docxToPdf } from "excelts/pdf";
 *
 * const doc = await readDocx(docxBytes);
 * const pdfBytes = await docxToPdf(doc);
 * ```
 */

import {
  layoutDocumentFull,
  type FullLayoutOptions,
  type PageGeometryOverride
} from "@word/layout/layout-full";
import type { Chart, ChartContent, ChartExContent, DocxDocument } from "@word/types";

import type { PdfPageBuilder } from "./builder/document-builder";
import { renderLayoutDocumentToPdf, type RenderLayoutOptions } from "./render-layout-to-pdf";

/** Options for DOCX → PDF conversion. */
export interface DocxToPdfOptions {
  /** Page width in points (default: from document sectPr or 612 US Letter). */
  readonly pageWidth?: number;
  /** Page height in points (default: from document sectPr or 792 US Letter). */
  readonly pageHeight?: number;
  /** Top margin in points (default: from sectPr or 72). */
  readonly marginTop?: number;
  /** Bottom margin in points (default: from sectPr or 72). */
  readonly marginBottom?: number;
  /** Left margin in points (default: from sectPr or 72). */
  readonly marginLeft?: number;
  /** Right margin in points (default: from sectPr or 72). */
  readonly marginRight?: number;
  /** Default font family (default: "Helvetica"). */
  readonly defaultFont?: string;
  /** Default font size in points (default: 11). */
  readonly defaultFontSize?: number;
  /**
   * Header margin from top edge in points (default: 36).
   *
   * @deprecated Currently has no effect on the layout-driven render
   * path: header paragraphs are rendered into the page's `marginTop`
   * strip starting at y=0 without an additional offset. Influence
   * header placement by widening `marginTop` instead.
   */
  readonly headerMargin?: number;
  /**
   * Footer margin from bottom edge in points (default: 36).
   *
   * @deprecated See {@link DocxToPdfOptions.headerMargin}. Footer
   * paragraphs are rendered starting at `pageHeight - marginBottom`.
   */
  readonly footerMargin?: number;
  /**
   * Optional high-quality chart renderer callback.
   *
   * When provided, Word charts are rendered using the injected renderer
   * instead of the built-in simplified renderer. This allows consumers
   * to plug in the Excel chart renderer for publication-quality output:
   *
   * ```typescript
   * import { installChartSupport } from "excelts/chart";
   * import { createWordChartPdfRenderer } from "excelts/pdf";
   * installChartSupport();
   * const pdfBytes = await docxToPdf(doc, {
   *   chartRenderer: createWordChartPdfRenderer()
   * });
   * ```
   *
   * The callback receives the original Word `Chart` definition (taken
   * from the source `ChartContent`), a `PdfPageBuilder`, and the
   * destination rectangle in PDF coordinates. The implementation
   * should draw the chart into the rectangle.
   *
   * Return `false` to decline a chart (the translator then falls back
   * to its built-in placeholder rendering: an outlined rectangle with
   * the chart title centred). Return `void` or `true` to indicate the
   * chart was handled.
   *
   * Note: `chartEx` charts never reach this callback because there is
   * no `Chart` instance to pass; they always render through the
   * translator's built-in path.
   */
  readonly chartRenderer?: (
    chart: Chart,
    page: PdfPageBuilder,
    rect: { x: number; y: number; width: number; height: number }
  ) => boolean | void;
}

/**
 * Convert a `DocxDocument` to PDF bytes.
 *
 * @param doc - The DOCX document model (from `readDocx` or `Document.build()`).
 * @param options - Page geometry, fonts, optional chart renderer, …
 * @returns PDF bytes ready to write to disk or stream over HTTP.
 */
export async function docxToPdf(
  doc: DocxDocument,
  options?: DocxToPdfOptions
): Promise<Uint8Array> {
  // 1. Resolve effective page geometry. Section properties win unless the
  //    caller explicitly overrode an axis. Margins are independent: the
  //    section's margins are applied unless the caller overrode them.
  const layoutOptions = mapToLayoutOptions(doc, options);

  // 2. Auto-detect chart support: if no explicit chartRenderer is
  //    provided, try to import the high-quality Excel-based renderer.
  let chartRendererForChart = options?.chartRenderer;
  if (!chartRendererForChart) {
    try {
      const mod = await import("./excel-bridge");
      if (typeof mod.createWordChartPdfRenderer === "function") {
        chartRendererForChart = mod.createWordChartPdfRenderer();
      }
    } catch {
      // Chart support not available — placeholder rendering takes over.
    }
  }

  // 3. Run the layout engine. Everything from line wrapping to page
  //    breaks happens here; word-bridge no longer carries any of that.
  const layout = layoutDocumentFull(doc, layoutOptions);

  // 4. Build a render-options object for the PDF translator. The
  //    chartRenderer adaptation: layout produces `LayoutChart` (which
  //    contains the original `ChartContent` in `source`); the public
  //    chartRenderer API takes the inner `Chart` model. We unwrap so
  //    existing callers keep working unchanged.
  const renderOptions: RenderLayoutOptions = {
    title: doc.coreProperties?.title,
    author: doc.coreProperties?.creator,
    subject: doc.coreProperties?.subject,
    defaultFont: options?.defaultFont ?? "Helvetica",
    defaultFontSize: options?.defaultFontSize ?? 11,
    chartRenderer: chartRendererForChart
      ? (layoutChart, page, rect): boolean | void => {
          const src = layoutChart.source as ChartContent | ChartExContent | undefined;
          if (src && src.type === "chart") {
            // Forward the user's return value so a renderer that knows
            // how to draw classic charts but declines a particular
            // family (e.g. unsupported axis combination) gets the
            // translator's placeholder fallback.
            return chartRendererForChart!(src.chart, page, rect);
          }
          // ChartEx (or absent source) cannot be passed to the user's
          // `Chart`-typed callback. Decline so the translator's
          // placeholder runs instead of leaving a blank slot.
          return false;
        }
      : undefined
  };

  const builder = renderLayoutDocumentToPdf(layout, renderOptions);
  return builder.build();
}

// =============================================================================
// Internal: option mapping
// =============================================================================

/**
 * Resolve the effective page geometry from `DocxToPdfOptions`. Caller-
 * supplied overrides win; otherwise the document's section properties
 * provide the value (with `pageSize.orientation === "landscape"`
 * triggering the conventional width/height swap when sectPr supplies
 * portrait-oriented numbers); otherwise the layout engine's defaults
 * (US Letter, 1-inch margins) take over.
 *
 * `headerMargin` / `footerMargin` are accepted for API compatibility
 * with the previous flow renderer but currently have no effect on the
 * layout-driven path: the layout engine does not yet position header /
 * footer paragraphs against custom inset values, and forwarding the
 * field would silently mislead callers. Pre-existing documents that
 * rely on those parameters still get the section's `headerMargin` /
 * `footerMargin` from sectionProperties when they exist.
 */
function mapToLayoutOptions(
  doc: DocxDocument,
  options: DocxToPdfOptions | undefined
): FullLayoutOptions {
  const sectProps = doc.sectionProperties;

  // Section page size, applying the orientation swap so a landscape
  // sectPr written with portrait numerics still ends up wide.
  let sectionPageWidthPt: number | undefined;
  let sectionPageHeightPt: number | undefined;
  if (sectProps?.pageSize) {
    sectionPageWidthPt = twipsToPt(sectProps.pageSize.width);
    sectionPageHeightPt = twipsToPt(sectProps.pageSize.height);
    if (sectProps.pageSize.orientation === "landscape") {
      [sectionPageWidthPt, sectionPageHeightPt] = [sectionPageHeightPt, sectionPageWidthPt];
    }
  }

  const sectionMarginTopPt =
    sectProps?.margins?.top != null ? twipsToPt(sectProps.margins.top) : undefined;
  const sectionMarginBottomPt =
    sectProps?.margins?.bottom != null ? twipsToPt(sectProps.margins.bottom) : undefined;
  const sectionMarginLeftPt =
    sectProps?.margins?.left != null ? twipsToPt(sectProps.margins.left) : undefined;
  const sectionMarginRightPt =
    sectProps?.margins?.right != null ? twipsToPt(sectProps.margins.right) : undefined;

  const pageGeometry: PageGeometryOverride = {
    pageWidth: options?.pageWidth ?? sectionPageWidthPt,
    pageHeight: options?.pageHeight ?? sectionPageHeightPt,
    marginTop: options?.marginTop ?? sectionMarginTopPt,
    marginBottom: options?.marginBottom ?? sectionMarginBottomPt,
    marginLeft: options?.marginLeft ?? sectionMarginLeftPt,
    marginRight: options?.marginRight ?? sectionMarginRightPt
  };

  const layoutOpts: Mutable<FullLayoutOptions> = {};
  // Only attach pageGeometry when at least one axis is actually
  // overridden; otherwise the layout engine should apply its
  // section-property fallbacks unchanged.
  if (Object.values(pageGeometry).some(v => v !== undefined)) {
    layoutOpts.pageGeometry = pageGeometry;
  }
  return layoutOpts;
}

function twipsToPt(twips: number): number {
  return twips / 20;
}

type Mutable<T> = { -readonly [K in keyof T]: T[K] };
