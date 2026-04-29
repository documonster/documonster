/**
 * Excel-to-PDF Bridge
 *
 * Converts Excel Workbook data into the PDF module's independent data model.
 * This is the ONLY file in the PDF module that imports from @excel.
 *
 * @example
 * ```typescript
 * import { Workbook } from "excelts";
 * import { excelToPdf } from "excelts/pdf";
 *
 * const workbook = new Workbook();
 * // ... build workbook ...
 * const pdf = await excelToPdf(workbook);
 * ```
 */

import type { Chart, RegionMapDataOptions } from "@excel/chart";
import {
  canRenderChartExAsVectorPdf,
  drawChartExPdf,
  drawChartPdf,
  renderChartExPng,
  renderChartPng
} from "@excel/chart";
import { ValueType } from "@excel/enums";
import { formatCellValue } from "@excel/utils/cell-format";
import type { Workbook } from "@excel/workbook";
import type { Worksheet } from "@excel/worksheet";
import { tryInvokeFormulaEngine } from "@formula/host-registry";
import { base64ToUint8Array } from "@utils/utils.base";

import { PdfDocumentBuilder } from "./builder/document-builder";
import { exportPdf } from "./render/pdf-exporter";
import {
  PdfCellType,
  type PdfWorkbook,
  type PdfSheetData,
  type PdfRowData,
  type PdfCellData,
  type PdfColumnData,
  type PdfCellStyle,
  type PdfFillData,
  type PdfColorData,
  type PdfFontStyle,
  type PdfBordersData,
  type PdfBorderSideData,
  type PdfAlignmentData,
  type PdfPageSetupData,
  type PdfSheetImage,
  type PdfExportOptions,
  type PdfCellTypeValue
} from "./types";

// =============================================================================
// Public API
// =============================================================================

/**
 * Export an Excel Workbook directly to PDF.
 *
 * This is a convenience function that converts the Workbook to the PDF module's
 * data model and then generates the PDF.
 * Yields to the event loop between each output page during layout and rendering.
 *
 * @param workbook - An Excel Workbook instance
 * @param options  - PDF export options
 * @returns Promise of PDF file as a Uint8Array
 */
export async function excelToPdf(
  workbook: Workbook,
  options?: PdfExportOptions
): Promise<Uint8Array> {
  // Recalculate all formulas before conversion so that formula results
  // reflect the latest cell values (fixes stale cached results from XLSX).
  //
  // The formula engine is opt-in: callers who import
  // `@cj-tech-master/excelts/formula` get automatic recalculation here; callers
  // who don't import it fall back to whatever cached results the XLSX
  // shipped with (safe for workbooks last saved by Excel itself).
  tryInvokeFormulaEngine(workbook);

  const pdfWorkbook = excelWorkbookToPdf(workbook);
  return exportPdf(pdfWorkbook, options);
}

/**
 * Options for {@link chartToPdf}.
 */
export interface ChartToPdfOptions {
  /** PDF page width in points. Default: max(chart width + 72, 400). */
  pageWidth?: number;
  /** PDF page height in points. Default: max(chart height + 72, 300). */
  pageHeight?: number;
  /** Chart render width in points. Default: 520. */
  width?: number;
  /** Chart render height in points. Default: 360. */
  height?: number;
  /**
   * Left margin in points between the chart and the page edge. Default: 36.
   * Used as top margin too so the chart sits in a 36-pt gutter.
   */
  margin?: number;
  /**
   * Force rasterisation even for classic charts. Default: `false`
   * (classic charts render as vector PDF content; ChartEx always
   * rasterises because there is no vector renderer — see the
   * "ChartEx PDF" note in `src/modules/excel/README.md`).
   */
  forceRaster?: boolean;
  /** PNG raster scale multiplier when rasterising. Default: 2 (for crisp text). */
  rasterScale?: number;
  /** Document metadata forwarded to the resulting PDF. */
  title?: string;
  author?: string;
  /**
   * ChartEx `regionMap` data. When supplied, the vector PDF path
   * uses the TopoJSON polygons (matched via `match` rules) instead
   * of the centroid preview. Ignored for non-regionMap layouts and
   * when the chart rasterises. Mirrors `renderChartExSvg`'s
   * `regionMap` option so a single caller-side object works for
   * both backends.
   */
  regionMap?: RegionMapDataOptions;
}

/**
 * Render a single {@link Chart} to a standalone one-page PDF.
 *
 * The output is a **zero-dependency deterministic preview**, not an
 * Excel-pixel-perfect rendering. Use this for server-side reports,
 * thumbnails, and CI artefacts where the goal is a recognisable chart
 * without a headless Office dependency. When pixel-identical output
 * matters (publication-grade reports, Excel/LibreOffice-compatible
 * formatting), round-trip the `.xlsx` through
 * `soffice --convert-to pdf` — the byte-preserving round-trip in this
 * library makes that a safe handoff. See `src/modules/excel/README.md`
 * → "Rendering scope" for the complete boundary list.
 *
 * Classic charts take the **vector** path: the chart is drawn directly
 * onto the page via `drawChartPdf`, so text stays selectable and shapes
 * remain resolution-independent. ChartEx charts take the **raster**
 * path: they are rendered to PNG through the SVG pipeline and embedded
 * as an image XObject, because the library intentionally does not ship
 * a vector `drawChartExPdf` (see README "ChartEx PDF note"). The raster
 * path is also available for classic charts via `forceRaster: true` for
 * cases where pixel-level fidelity with the SVG/PNG preview matters
 * more than selectable text.
 *
 * Lives in `excel-bridge.ts` because invoking the PDF builder from the
 * chart module would cross the Layer 4 → Layer 5 import boundary
 * documented in `AGENTS.md`. Consumers import it from
 * `@cj-tech-master/excelts/pdf` alongside `excelToPdf`.
 */
export async function chartToPdf(
  chart: Chart,
  options: ChartToPdfOptions = {}
): Promise<Uint8Array> {
  const width = options.width ?? 520;
  const height = options.height ?? 360;
  const margin = options.margin ?? 36;
  const pageWidth = options.pageWidth ?? Math.max(width + margin * 2, 400);
  const pageHeight = options.pageHeight ?? Math.max(height + margin * 2, 300);

  const doc = new PdfDocumentBuilder();
  if (options.title || options.author) {
    doc.setMetadata({
      title: options.title,
      author: options.author
    });
  }
  const page = doc.addPage({ width: pageWidth, height: pageHeight });

  const isChartEx = chart.chartExModel !== undefined;
  // ChartEx charts whose every series has a layoutId the renderer
  // can express as PDF geometry (currently sunburst + treemap) take
  // the vector route alongside classic charts. Anything else — or
  // any chart the caller explicitly asks to rasterise via
  // `forceRaster` — falls through to the SVG → PNG → image-XObject
  // pipeline.
  const chartExModel = chart.chartExModel;
  const chartExVectorable =
    isChartEx && chartExModel !== undefined && canRenderChartExAsVectorPdf(chartExModel);
  const useRaster = options.forceRaster === true || (isChartEx && !chartExVectorable);

  if (!useRaster) {
    if (isChartEx && chartExModel !== undefined) {
      drawChartExPdf(
        page,
        chartExModel,
        {
          x: margin,
          y: pageHeight - margin - height,
          width,
          height
        },
        { title: options.title, regionMap: options.regionMap }
      );
      return doc.build();
    }
    // Vector path for classic charts.
    const model = chart.chartModel;
    if (!model) {
      throw new Error(
        "chartToPdf: Chart has neither a classic model nor a ChartEx model to render"
      );
    }
    drawChartPdf(page, model, {
      x: margin,
      y: pageHeight - margin - height,
      width,
      height
    });
    return doc.build();
  }

  // Raster path: produce a PNG, then embed it on the page. Uses scale
  // 2× by default so the PDF viewer shows crisp text even when zoomed
  // into a 150 % magnification. Callers who need larger prints can
  // bump `rasterScale`; anything above 4 rapidly grows the PDF size.
  const scale = options.rasterScale ?? 2;
  const pngBytes = isChartEx
    ? await renderChartExPng(chart.chartExModel!, { width, height, scale })
    : await renderChartPng(chart.chartModel!, { width, height, scale });
  page.drawImage({
    data: pngBytes,
    format: "png",
    x: margin,
    y: pageHeight - margin - height,
    width,
    height
  });
  return doc.build();
}

/**
 * Convert an Excel Workbook to the internal PdfWorkbook data structure.
 */
function excelWorkbookToPdf(workbook: Workbook): PdfWorkbook {
  return {
    title: workbook.title || undefined,
    creator: workbook.creator || undefined,
    subject: workbook.subject || undefined,
    sheets: workbook.worksheets.map(ws => convertSheet(ws, workbook))
  };
}

// =============================================================================
// Sheet Conversion
// =============================================================================

function convertSheet(ws: Worksheet, workbook: Workbook): PdfSheetData {
  const dimensions = ws.dimensions;
  const hasData = dimensions && dimensions.model.top > 0 && dimensions.model.left > 0;

  const bounds = hasData
    ? {
        top: dimensions.model.top,
        left: dimensions.model.left,
        bottom: dimensions.model.bottom,
        right: dimensions.model.right
      }
    : { top: 0, left: 0, bottom: 0, right: 0 };

  // Convert columns
  const columns = new Map<number, PdfColumnData>();
  if (hasData) {
    for (let c = bounds.left; c <= bounds.right; c++) {
      const col = ws.getColumn(c);
      columns.set(c, {
        hidden: col.hidden || undefined,
        width: col.width ?? undefined
      });
    }
  }

  // Convert rows
  const rows = new Map<number, PdfRowData>();
  if (hasData) {
    for (let r = bounds.top; r <= bounds.bottom; r++) {
      const row = ws.findRow(r);
      if (!row) {
        continue;
      }

      const cells = new Map<number, PdfCellData>();
      row.eachCell({ includeEmpty: true }, cell => {
        const hasValue = cell.type !== ValueType.Null && cell.type !== ValueType.Merge;
        const hasStyle =
          cell.style &&
          ((cell.style.border &&
            (cell.style.border.top ||
              cell.style.border.right ||
              cell.style.border.bottom ||
              cell.style.border.left)) ||
            cell.style.fill ||
            cell.style.font);
        if (hasValue || hasStyle) {
          cells.set(cell.col, convertCell(cell));
        }
      });

      rows.set(r, {
        hidden: row.hidden || undefined,
        height: row.height ?? undefined,
        customHeight: row.customHeight || undefined,
        cells
      });
    }
  }

  // Convert merges
  const merges = ws.hasMerges && ws.model.mergeCells ? [...ws.model.mergeCells] : undefined;

  // Convert pageSetup
  const ps = ws.pageSetup;
  const pageSetup: PdfPageSetupData | undefined = ps
    ? {
        orientation: ps.orientation,
        paperSize: ps.paperSize,
        margins: ps.margins
          ? {
              left: ps.margins.left,
              right: ps.margins.right,
              top: ps.margins.top,
              bottom: ps.margins.bottom
            }
          : undefined,
        scale: ps.scale,
        printTitlesRow: ps.printTitlesRow,
        showGridLines: ps.showGridLines,
        printArea: (ps as any).printArea
      }
    : undefined;

  // Convert row/col breaks
  const rowBreaks: number[] | undefined = (ws as any).rowBreaks?.map((b: { id: number }) => b.id);
  const colBreaks: number[] | undefined = (ws as any).colBreaks?.map((b: { id: number }) => b.id);

  // Convert images
  const images = collectImages(ws, workbook);

  // Extend bounds to cover image anchors (so layout engine includes them)
  if (images) {
    for (const img of images) {
      const tl = img.range.tl;
      const tlCol = (tl.nativeCol ?? tl.col ?? 0) + 1; // 0-indexed → 1-indexed
      const tlRow = (tl.nativeRow ?? tl.row ?? 0) + 1;
      if (bounds.top === 0 && bounds.left === 0) {
        bounds.top = 1;
        bounds.left = 1;
      }
      if (tlCol > bounds.right) {
        bounds.right = tlCol;
      }
      if (tlRow > bounds.bottom) {
        bounds.bottom = tlRow;
      }

      // Also extend to bottom-right anchor if present
      if (img.range.br) {
        const br = img.range.br;
        const brCol = (br.nativeCol ?? br.col ?? 0) + 1;
        const brRow = (br.nativeRow ?? br.row ?? 0) + 1;
        if (brCol > bounds.right) {
          bounds.right = brCol;
        }
        if (brRow > bounds.bottom) {
          bounds.bottom = brRow;
        }
      }
    }

    // Ensure columns/rows exist for extended bounds
    for (let c = bounds.left; c <= bounds.right; c++) {
      if (!columns.has(c)) {
        const col = ws.getColumn(c);
        columns.set(c, {
          hidden: col.hidden || undefined,
          width: col.width ?? undefined
        });
      }
    }
    for (let r = bounds.top; r <= bounds.bottom; r++) {
      if (!rows.has(r)) {
        rows.set(r, { cells: new Map() });
      }
    }
  }

  return {
    name: ws.name,
    state: (ws as any).state ?? "visible",
    bounds,
    columns,
    rows,
    merges,
    pageSetup,
    rowBreaks,
    colBreaks,
    images
  };
}

// =============================================================================
// Cell Conversion
// =============================================================================

// Use any-typed cell to avoid importing the Cell class directly
// (Worksheet.eachCell provides it)
function convertCell(cell: any): PdfCellData {
  const type = mapValueType(cell.type);
  const text = getCellDisplayText(cell);
  const style = convertCellStyle(cell.style);

  return {
    type,
    value: convertCellValue(cell),
    text,
    style,
    hyperlink: cell.hyperlink || undefined,
    result: cell.result ?? undefined,
    col: cell.col
  };
}

function mapValueType(vt: number): PdfCellTypeValue {
  switch (vt) {
    case ValueType.Null:
      return PdfCellType.Empty;
    case ValueType.Merge:
      return PdfCellType.Merge;
    case ValueType.Number:
      return PdfCellType.Number;
    case ValueType.String:
    case ValueType.SharedString:
      return PdfCellType.String;
    case ValueType.Date:
      return PdfCellType.Date;
    case ValueType.Hyperlink:
      return PdfCellType.Hyperlink;
    case ValueType.Formula:
      return PdfCellType.Formula;
    case ValueType.RichText:
      return PdfCellType.RichText;
    case ValueType.Boolean:
      return PdfCellType.Boolean;
    case ValueType.Error:
      return PdfCellType.Error;
    default:
      return PdfCellType.String;
  }
}

/**
 * Get display text for a cell, applying numFmt formatting.
 */
function getCellDisplayText(cell: any): string {
  if (!cell) {
    return "";
  }

  switch (cell.type) {
    case ValueType.Null:
    case ValueType.Merge:
      return "";
    case ValueType.RichText:
    case ValueType.Hyperlink:
      return cell.text ?? "";
    case ValueType.Error: {
      const errValue = cell.value;
      return errValue?.error ?? cell.text ?? "";
    }
    case ValueType.Formula: {
      const result = cell.result;
      if (result !== undefined && result !== null) {
        if (typeof result === "object" && "error" in result) {
          return result.error;
        }
        return formatCellValueSafe(result, cell.style?.numFmt);
      }
      return cell.text ?? "";
    }
    default: {
      const value = cell.value;
      if (value === null || value === undefined) {
        return "";
      }
      return formatCellValueSafe(value, cell.style?.numFmt);
    }
  }
}

function formatCellValueSafe(
  value: unknown,
  numFmt: string | { formatCode: string } | undefined
): string {
  const fmt = typeof numFmt === "string" ? numFmt : numFmt?.formatCode;
  if (fmt && (typeof value === "number" || value instanceof Date || typeof value === "boolean")) {
    try {
      return formatCellValue(value, fmt);
    } catch {
      // Fall through to default
    }
  }
  if (value instanceof Date) {
    return value.toLocaleDateString();
  }
  return String(value);
}

function convertCellValue(cell: any): unknown {
  if (cell.type === ValueType.RichText) {
    // Preserve richText structure for the PDF engine
    const rtValue = cell.value;
    if (rtValue?.richText) {
      return {
        richText: rtValue.richText.map((run: any) => ({
          text: run.text,
          font: run.font ? convertFontStyle(run.font) : undefined
        }))
      };
    }
  }
  return cell.value;
}

// =============================================================================
// Style Conversion
// =============================================================================

function convertCellStyle(style: any): Partial<PdfCellStyle> | undefined {
  if (!style) {
    return undefined;
  }

  return {
    font: style.font ? convertFontStyle(style.font) : undefined,
    numFmt: style.numFmt,
    fill: style.fill ? convertFill(style.fill) : undefined,
    border: style.border ? convertBorders(style.border) : undefined,
    alignment: style.alignment ? convertAlignment(style.alignment) : undefined
  };
}

function convertFontStyle(font: any): Partial<PdfFontStyle> {
  return {
    name: font.name,
    size: font.size,
    bold: font.bold,
    italic: font.italic,
    strike: font.strike,
    underline: font.underline,
    color: font.color ? convertColor(font.color) : undefined
  };
}

function convertColor(color: any): PdfColorData {
  return {
    argb: color.argb,
    theme: color.theme,
    tint: color.tint
  };
}

function convertFill(fill: any): PdfFillData {
  const result: PdfFillData = {
    type: fill.type ?? "pattern",
    pattern: fill.pattern,
    fgColor: fill.fgColor ? convertColor(fill.fgColor) : undefined
  };

  if (fill.stops) {
    result.stops = fill.stops.map((s: any) => ({
      color: convertColor(s.color)
    }));
  }

  return result;
}

function convertBorderSide(border: any): Partial<PdfBorderSideData> {
  return {
    style: border.style,
    color: border.color ? convertColor(border.color) : undefined
  };
}

function convertBorders(borders: any): Partial<PdfBordersData> {
  return {
    top: borders.top ? convertBorderSide(borders.top) : undefined,
    right: borders.right ? convertBorderSide(borders.right) : undefined,
    bottom: borders.bottom ? convertBorderSide(borders.bottom) : undefined,
    left: borders.left ? convertBorderSide(borders.left) : undefined
  };
}

function convertAlignment(alignment: any): Partial<PdfAlignmentData> {
  return {
    horizontal: alignment.horizontal,
    vertical: alignment.vertical,
    wrapText: alignment.wrapText,
    indent: alignment.indent,
    textRotation: alignment.textRotation
  };
}

// =============================================================================
// Image Collection
// =============================================================================

function collectImages(ws: Worksheet, workbook: Workbook): PdfSheetImage[] | undefined {
  const wsImages = (ws as any).getImages?.();
  if (!wsImages || !Array.isArray(wsImages) || wsImages.length === 0) {
    return undefined;
  }

  const images: PdfSheetImage[] = [];

  for (const wsImage of wsImages) {
    if (!wsImage.range?.tl) {
      continue;
    }

    const imageId = wsImage.imageId;
    const mediaItem = (workbook as any).getImage?.(Number(imageId));
    if (!mediaItem) {
      continue;
    }

    // Get image data
    let data: Uint8Array | undefined;
    if (mediaItem.buffer instanceof Uint8Array) {
      data = mediaItem.buffer;
    } else if (mediaItem.base64) {
      data = base64ToUint8Array(mediaItem.base64);
    }
    if (!data || data.length === 0) {
      continue;
    }

    const format = mediaItem.extension as string;
    if (format !== "jpeg" && format !== "png") {
      continue;
    }

    images.push({
      data,
      format: format as "jpeg" | "png",
      range: {
        tl: {
          col: wsImage.range.tl.col ?? 0,
          row: wsImage.range.tl.row ?? 0,
          nativeCol: wsImage.range.tl.nativeCol,
          nativeRow: wsImage.range.tl.nativeRow,
          nativeColOff: wsImage.range.tl.nativeColOff,
          nativeRowOff: wsImage.range.tl.nativeRowOff
        },
        br: wsImage.range.br
          ? {
              col: wsImage.range.br.col ?? 0,
              row: wsImage.range.br.row ?? 0,
              nativeCol: wsImage.range.br.nativeCol,
              nativeRow: wsImage.range.br.nativeRow,
              nativeColOff: wsImage.range.br.nativeColOff,
              nativeRowOff: wsImage.range.br.nativeRowOff
            }
          : undefined,
        ext: wsImage.range.ext
          ? { width: wsImage.range.ext.width, height: wsImage.range.ext.height }
          : undefined
      }
    });
  }

  return images.length > 0 ? images : undefined;
}
