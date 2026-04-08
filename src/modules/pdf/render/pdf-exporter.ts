/**
 * PDF Exporter - Main orchestrator for PDF document generation.
 *
 * Coordinates the layout engine, page renderer, font manager, and PDF writer
 * to produce a complete PDF document from a PdfWorkbook data structure.
 *
 * This module is fully independent of the Excel module.
 * It is used internally by the public `pdf()` and `excelToPdf()` APIs.
 */

import { PdfWriter } from "../core/pdf-writer";
import { PdfDict, pdfRef, pdfNumber, pdfString as pdfStr } from "../core/pdf-object";
import { FontManager, resolvePdfFontName } from "../font/font-manager";
import { parseTtf } from "../font/ttf-parser";
import { initEncryption } from "../core/encryption";
import { layoutSheet } from "./layout-engine";
import { renderPage, alphaGsName } from "./page-renderer";
import { decodePng } from "./png-decoder";
import { PdfError, PdfRenderError } from "../errors";
import {
  PageSizes,
  type PdfWorkbook,
  type PdfSheetData,
  type PdfExportOptions,
  type ResolvedPdfOptions,
  type PdfPageSize,
  type PdfMargins,
  type PdfColor,
  type PdfOrientation,
  type LayoutPage
} from "../types";
import { argbToPdfColor } from "./style-converter";
import { yieldToEventLoop } from "@utils/utils.base";

// =============================================================================
// Public API
// =============================================================================

/**
 * Export a PdfWorkbook to PDF format.
 * Yields to the event loop between each output page during layout and rendering.
 *
 * @param workbook - The workbook data to export
 * @param options - Export options controlling layout, pagination, and appearance
 * @returns Promise of PDF file as a Uint8Array
 * @throws {PdfError} If the workbook has no sheets or export fails
 */
export async function exportPdf(
  workbook: PdfWorkbook,
  options?: PdfExportOptions
): Promise<Uint8Array> {
  const ctx = prepareExport(workbook, options);

  for (const sheet of ctx.sheets) {
    await layoutSheetInto(ctx, sheet, options);
  }

  return finishExport(ctx, workbook, options);
}

// =============================================================================
// Internal — Shared Pipeline
// =============================================================================

/** Shared state for the export pipeline. */
interface ExportContext {
  sheets: PdfSheetData[];
  fontManager: FontManager;
  writer: PdfWriter;
  allPages: LayoutPage[];
}

/**
 * Shared setup: validate sheets, create font manager and writer,
 * register embedded font.
 */
function prepareExport(workbook: PdfWorkbook, options?: PdfExportOptions): ExportContext {
  const sheets = selectSheets(workbook, options?.sheets);

  if (sheets.length === 0) {
    throw new PdfError("No sheets to export. The workbook is empty or no sheets matched.");
  }

  const fontManager = new FontManager();
  const writer = new PdfWriter();

  if (options?.font) {
    try {
      const ttf = parseTtf(options.font);
      fontManager.registerEmbeddedFont(ttf);
    } catch (err) {
      throw new PdfRenderError("Failed to parse TrueType font", { cause: err });
    }
  }

  return { sheets, fontManager, writer, allPages: [] };
}

/**
 * Layout a single sheet and append its pages to the context.
 */
async function layoutSheetInto(
  ctx: ExportContext,
  sheet: PdfSheetData,
  options?: PdfExportOptions
): Promise<void> {
  try {
    const resolved = resolveOptions(options, sheet);
    const pages = await layoutSheet(sheet, resolved, ctx.fontManager);
    ctx.allPages.push(...pages);
  } catch (err) {
    throw new PdfRenderError(`Failed to layout sheet "${sheet.name}"`, { cause: err });
  }
}

/**
 * After layout: fix page numbers, track fonts, write resources,
 * render pages, and build the final PDF binary.
 */
async function finishExport(
  ctx: ExportContext,
  workbook: PdfWorkbook,
  options?: PdfExportOptions
): Promise<Uint8Array> {
  const { allPages, fontManager, writer, sheets } = ctx;
  const documentOptions = resolveOptions(options, sheets[0]);

  ensureAtLeastOnePage(allPages, documentOptions, sheets);
  fixPageNumbers(allPages);
  trackFontsForHeaders(allPages, fontManager);

  const fontObjectMap = fontManager.writeFontResources(writer);
  const { pageObjNums, sheetFirstPage, pagesTreeObjNum } = await renderAllPages(
    allPages,
    fontManager,
    writer,
    fontObjectMap
  );

  return buildFinalPdf(
    writer,
    pageObjNums,
    pagesTreeObjNum,
    sheetFirstPage,
    documentOptions,
    workbook,
    options
  );
}

function ensureAtLeastOnePage(
  allPages: LayoutPage[],
  documentOptions: ResolvedPdfOptions,
  sheets: PdfSheetData[]
): void {
  if (allPages.length === 0) {
    allPages.push({
      pageNumber: 1,
      options: documentOptions,
      cells: [],
      width: documentOptions.pageSize.width,
      height: documentOptions.pageSize.height,
      sheetName: sheets[0]?.name ?? "Sheet1",
      sheetCols: [],
      columnOffsets: [],
      columnWidths: [],
      sheetRows: [],
      rowYPositions: [],
      rowHeights: [],
      images: []
    });
  }
}

function fixPageNumbers(allPages: LayoutPage[]): void {
  for (let i = 0; i < allPages.length; i++) {
    allPages[i].pageNumber = i + 1;
  }
}

function trackFontsForHeaders(allPages: LayoutPage[], fontManager: FontManager): void {
  if (fontManager.hasEmbeddedFont()) {
    for (const page of allPages) {
      if (page.options.showSheetNames) {
        fontManager.trackText(page.sheetName);
      }
    }
  }

  for (const page of allPages) {
    if (page.options.showPageNumbers) {
      fontManager.ensureFont(resolvePdfFontName(page.options.defaultFontFamily, false, false));
    }
  }

  if (!fontManager.hasEmbeddedFont()) {
    for (const page of allPages) {
      if (page.options.showSheetNames) {
        fontManager.ensureFont(resolvePdfFontName(page.options.defaultFontFamily, true, false));
      }
    }
  }
}

interface RenderResult {
  pageObjNums: number[];
  sheetFirstPage: Map<string, number>;
  pagesTreeObjNum: number;
}

async function renderAllPages(
  allPages: LayoutPage[],
  fontManager: FontManager,
  writer: PdfWriter,
  fontObjectMap: Map<string, number>
): Promise<RenderResult> {
  const pageObjNums: number[] = [];
  const pagesTreeObjNum = writer.allocObject();
  const sheetFirstPage = new Map<string, number>();
  const totalPages = allPages.length;

  for (let i = 0; i < allPages.length; i++) {
    renderSinglePage(
      allPages[i],
      fontManager,
      writer,
      fontObjectMap,
      totalPages,
      pageObjNums,
      pagesTreeObjNum,
      sheetFirstPage
    );
    if (i < allPages.length - 1) {
      await yieldToEventLoop();
    }
  }

  return { pageObjNums, sheetFirstPage, pagesTreeObjNum };
}

function renderSinglePage(
  page: LayoutPage,
  fontManager: FontManager,
  writer: PdfWriter,
  fontObjectMap: Map<string, number>,
  totalPages: number,
  pageObjNums: number[],
  pagesTreeObjNum: number,
  sheetFirstPage: Map<string, number>
): void {
  try {
    const { stream: contentStream, alphaValues } = renderPage(
      page,
      page.options,
      fontManager,
      totalPages
    );

    // Handle images: create XObject Image entries and draw them
    const imageXObjects = new Map<string, number>();
    if (page.images.length > 0) {
      for (let imgIdx = 0; imgIdx < page.images.length; imgIdx++) {
        const img = page.images[imgIdx];
        const imgName = `Im${imgIdx + 1}`;
        const imgObjNum = writeImageXObject(writer, img.data, img.format);
        imageXObjects.set(imgName, imgObjNum);
        contentStream.drawImage(imgName, img.rect.x, img.rect.y, img.rect.width, img.rect.height);
      }
    }

    // Add content stream object
    const contentObjNum = writer.allocObject();
    const contentDict = new PdfDict();
    writer.addStreamObject(contentObjNum, contentDict, contentStream);

    // Add resources dictionary object
    const resourcesObjNum = writer.allocObject();
    const fontDictStr = fontManager.buildFontDictString(fontObjectMap);
    const resourcesDict = new PdfDict().set("Font", fontDictStr);
    if (imageXObjects.size > 0) {
      const xobjParts = ["<<"];
      for (const [name, objNum] of imageXObjects) {
        xobjParts.push(`/${name} ${pdfRef(objNum)}`);
      }
      xobjParts.push(">>");
      resourcesDict.set("XObject", xobjParts.join("\n"));
    }
    if (alphaValues.size > 0) {
      const gsParts = ["<<"];
      for (const alpha of alphaValues) {
        const gsObjNum = writer.allocObject();
        const gsDict = new PdfDict()
          .set("Type", "/ExtGState")
          .set("ca", pdfNumber(alpha))
          .set("CA", pdfNumber(alpha));
        writer.addObject(gsObjNum, gsDict);
        gsParts.push(`/${alphaGsName(alpha)} ${pdfRef(gsObjNum)}`);
      }
      gsParts.push(">>");
      resourcesDict.set("ExtGState", gsParts.join("\n"));
    }
    writer.addObject(resourcesObjNum, resourcesDict);

    // Create link annotations for hyperlinks
    const annotRefs: number[] = [];
    for (const cell of page.cells) {
      if (cell.hyperlink) {
        const annotObjNum = writer.allocObject();
        const rect = `[${pdfNumber(cell.rect.x)} ${pdfNumber(cell.rect.y)} ${pdfNumber(cell.rect.x + cell.rect.width)} ${pdfNumber(cell.rect.y + cell.rect.height)}]`;
        const annotDict = new PdfDict()
          .set("Type", "/Annot")
          .set("Subtype", "/Link")
          .set("Rect", rect)
          .set("Border", "[0 0 0]")
          .set(
            "A",
            `<< /Type /Action /S /URI /URI (${cell.hyperlink.replace(/[()\\]/g, "\\$&")}) >>`
          );
        writer.addObject(annotObjNum, annotDict);
        annotRefs.push(annotObjNum);
      }
    }

    // Add page object
    const pageObjNum = writer.addPage({
      parentRef: pagesTreeObjNum,
      width: page.width,
      height: page.height,
      contentsRef: contentObjNum,
      resourcesRef: resourcesObjNum,
      annotRefs: annotRefs.length > 0 ? annotRefs : undefined
    });

    pageObjNums.push(pageObjNum);

    if (!sheetFirstPage.has(page.sheetName)) {
      sheetFirstPage.set(page.sheetName, pageObjNums.length - 1);
    }
  } catch (err) {
    throw new PdfRenderError(`Failed to render page ${page.pageNumber} of "${page.sheetName}"`, {
      cause: err
    });
  }
}

function buildFinalPdf(
  writer: PdfWriter,
  pageObjNums: number[],
  pagesTreeObjNum: number,
  sheetFirstPage: Map<string, number>,
  documentOptions: ResolvedPdfOptions,
  workbook: PdfWorkbook,
  options?: PdfExportOptions
): Uint8Array {
  // --- Step 4: Build page tree ---
  const pagesKids = "[" + pageObjNums.map(n => pdfRef(n)).join(" ") + "]";
  const pagesDict = new PdfDict()
    .set("Type", "/Pages")
    .set("Kids", pagesKids)
    .set("Count", String(pageObjNums.length));
  writer.addObject(pagesTreeObjNum, pagesDict);

  // --- Step 5: Build outlines (bookmarks) for sheet navigation ---
  let outlinesRef: number | undefined;
  if (sheetFirstPage.size > 1) {
    outlinesRef = buildOutlines(writer, sheetFirstPage, pageObjNums);
  }

  // --- Step 6: Build catalog ---
  writer.addCatalog(pagesTreeObjNum, outlinesRef);

  // --- Step 7: Add document info ---
  writer.addInfoDict({
    title: documentOptions.title || workbook.title || undefined,
    author: documentOptions.author || workbook.creator || undefined,
    subject: documentOptions.subject || workbook.subject || undefined,
    creator: documentOptions.creator
  });

  // --- Step 8: Enable encryption if requested ---
  if (options?.encryption) {
    const encState = initEncryption(options.encryption);
    writer.setEncryption(encState);
  }

  // --- Step 9: Build the PDF ---
  return writer.build();
}

// =============================================================================
// Sheet Selection
// =============================================================================

/**
 * Select which sheets to export based on the options.
 */
function selectSheets(workbook: PdfWorkbook, sheets?: (string | number)[]): PdfSheetData[] {
  const allSheets = workbook.sheets;

  if (!sheets || sheets.length === 0) {
    // Export all visible sheets
    return allSheets.filter(ws => ws.state !== "hidden" && ws.state !== "veryHidden");
  }

  const result: PdfSheetData[] = [];
  for (const selector of sheets) {
    if (typeof selector === "string") {
      const ws = allSheets.find(s => s.name.toLowerCase() === selector.toLowerCase());
      if (ws) {
        result.push(ws);
      }
    } else if (typeof selector === "number") {
      // 1-based position in the sheets array
      const ws = allSheets[selector - 1];
      if (ws) {
        result.push(ws);
      }
    }
  }

  return result;
}

// =============================================================================
// Options Resolution
// =============================================================================

/**
 * Resolve user options with defaults.
 */
function resolveOptions(
  options: PdfExportOptions | undefined,
  sheet?: PdfSheetData
): ResolvedPdfOptions {
  // Use sheet's pageSetup as fallback for unspecified options
  const ps = sheet?.pageSetup;

  const pageSize = resolvePageSize(options?.pageSize, ps?.paperSize);
  const orientation: PdfOrientation =
    options?.orientation ?? (ps?.orientation === "landscape" ? "landscape" : "portrait");
  const margins = resolveMargins(options?.margins, ps?.margins);

  const gridLineColorStr = options?.gridLineColor ?? "FFD0D0D0";
  const gridLineColor: PdfColor = argbToPdfColor(gridLineColorStr) ?? {
    r: 0.816,
    g: 0.816,
    b: 0.816
  };

  // Use sheet's printTitlesRow as fallback for repeatRows
  let repeatRows: number | false = options?.repeatRows ?? false;
  if (repeatRows === false && ps?.printTitlesRow) {
    // printTitlesRow format: "1:3" (repeat rows 1-3) or "1" (repeat row 1)
    const match = ps.printTitlesRow.match(/^(\d+)(?::(\d+))?$/);
    if (match) {
      repeatRows = parseInt(match[2] ?? match[1], 10);
    }
  }

  return {
    pageSize,
    orientation,
    margins,
    fitToPage: options?.fitToPage !== undefined ? options.fitToPage : true,
    scale: Math.max(0.1, Math.min(3.0, options?.scale ?? (ps?.scale ? ps.scale / 100 : 1.0))),
    showGridLines: options?.showGridLines ?? ps?.showGridLines ?? false,
    gridLineColor,
    repeatRows,
    defaultFontFamily: options?.defaultFontFamily ?? "Helvetica",
    defaultFontSize: options?.defaultFontSize ?? 11,
    showSheetNames: options?.showSheetNames ?? false,
    showPageNumbers: options?.showPageNumbers ?? false,
    title: options?.title ?? "",
    author: options?.author ?? "",
    subject: options?.subject ?? "",
    creator: options?.creator ?? "excelts"
  };
}

/** Map PaperSize enum values to PDF page sizes. */
const PAPER_SIZE_MAP: Record<number, PdfPageSize> = {
  1: PageSizes.LETTER,
  5: PageSizes.LEGAL,
  9: PageSizes.A4,
  8: PageSizes.A3,
  11: PageSizes.A5,
  17: PageSizes.TABLOID
};

function resolvePageSize(
  size: PdfExportOptions["pageSize"] | undefined,
  paperSize?: number
): PdfPageSize {
  if (size) {
    if (typeof size === "string") {
      return PageSizes[size] ?? PageSizes.A4;
    }
    return size;
  }
  // Fallback to sheet paperSize
  if (paperSize !== undefined) {
    return PAPER_SIZE_MAP[paperSize] ?? PageSizes.A4;
  }
  return PageSizes.A4;
}

/**
 * Resolve margins with defaults. Sheet margins are in inches, convert to points (×72).
 * When partial PDF margins are specified, unset sides fall back to sheet margins,
 * then to the default 72pt (1 inch).
 */
function resolveMargins(
  margins?: Partial<PdfMargins>,
  wsMargins?: { left: number; right: number; top: number; bottom: number }
): PdfMargins {
  // Build a base from sheet pageSetup margins (inches → points), or default 72pt
  const base: PdfMargins = wsMargins
    ? {
        top: wsMargins.top * 72,
        right: wsMargins.right * 72,
        bottom: wsMargins.bottom * 72,
        left: wsMargins.left * 72
      }
    : { top: 72, right: 72, bottom: 72, left: 72 };

  if (!margins) {
    return base;
  }

  return {
    top: margins.top ?? base.top,
    right: margins.right ?? base.right,
    bottom: margins.bottom ?? base.bottom,
    left: margins.left ?? base.left
  };
}

// =============================================================================
// PDF Outlines (Bookmarks)
// =============================================================================

/**
 * Build a PDF outlines tree for sheet-level navigation.
 * Creates one bookmark entry per sheet, pointing to the first page.
 */
function buildOutlines(
  writer: PdfWriter,
  sheetFirstPage: Map<string, number>,
  pageObjNums: number[]
): number {
  const outlinesObjNum = writer.allocObject();
  const entries = Array.from(sheetFirstPage.entries());

  // Allocate outline item object numbers
  const itemObjNums: number[] = [];
  for (let i = 0; i < entries.length; i++) {
    itemObjNums.push(writer.allocObject());
  }

  // Write outline items
  for (let i = 0; i < entries.length; i++) {
    const [sheetName, pageIndex] = entries[i];
    const pageObjNum = pageObjNums[pageIndex];
    const itemDict = new PdfDict()
      .set("Title", pdfStr(sheetName))
      .set("Parent", pdfRef(outlinesObjNum))
      .set("Dest", `[${pdfRef(pageObjNum)} /Fit]`);

    if (i > 0) {
      itemDict.set("Prev", pdfRef(itemObjNums[i - 1]));
    }
    if (i < entries.length - 1) {
      itemDict.set("Next", pdfRef(itemObjNums[i + 1]));
    }

    writer.addObject(itemObjNums[i], itemDict);
  }

  // Write outlines root
  const outlinesDict = new PdfDict()
    .set("Type", "/Outlines")
    .set("First", pdfRef(itemObjNums[0]))
    .set("Last", pdfRef(itemObjNums[itemObjNums.length - 1]))
    .set("Count", String(entries.length));
  writer.addObject(outlinesObjNum, outlinesDict);

  return outlinesObjNum;
}

// =============================================================================
// Image XObject
// =============================================================================

/**
 * Write a JPEG or PNG image as a PDF XObject Image.
 */
function writeImageXObject(writer: PdfWriter, data: Uint8Array, format: "jpeg" | "png"): number {
  if (format === "png") {
    return writePngImageXObject(writer, data);
  }
  return writeJpegImageXObject(writer, data);
}

/**
 * Write a JPEG image using DCTDecode (raw JPEG data embedded directly).
 */
function writeJpegImageXObject(writer: PdfWriter, data: Uint8Array): number {
  const objNum = writer.allocObject();
  const dims = getJpegDimensions(data);
  const dict = new PdfDict()
    .set("Type", "/XObject")
    .set("Subtype", "/Image")
    .set("Width", pdfNumber(dims.width))
    .set("Height", pdfNumber(dims.height))
    .set("ColorSpace", "/DeviceRGB")
    .set("BitsPerComponent", "8")
    .set("Filter", "/DCTDecode");
  writer.addStreamObject(objNum, dict, data);
  return objNum;
}

/**
 * Write a PNG image: decode to raw RGB, create SMask for alpha if needed.
 */
function writePngImageXObject(writer: PdfWriter, data: Uint8Array): number {
  const png = decodePng(data);
  const objNum = writer.allocObject();

  const dict = new PdfDict()
    .set("Type", "/XObject")
    .set("Subtype", "/Image")
    .set("Width", pdfNumber(png.width))
    .set("Height", pdfNumber(png.height))
    .set("ColorSpace", "/DeviceRGB")
    .set("BitsPerComponent", pdfNumber(png.bitsPerComponent));

  // If the image has an alpha channel, create a soft mask XObject
  if (png.alpha) {
    const smaskObjNum = writer.allocObject();
    const smaskDict = new PdfDict()
      .set("Type", "/XObject")
      .set("Subtype", "/Image")
      .set("Width", pdfNumber(png.width))
      .set("Height", pdfNumber(png.height))
      .set("ColorSpace", "/DeviceGray")
      .set("BitsPerComponent", "8");
    writer.addStreamObject(smaskObjNum, smaskDict, png.alpha);
    dict.set("SMask", pdfRef(smaskObjNum));
  }

  // RGB pixel data — compression handled by addStreamObject
  writer.addStreamObject(objNum, dict, png.pixels);
  return objNum;
}

/**
 * Extract width and height from a JPEG file header.
 * Scans for SOF markers (SOF0-SOF3, SOF5-SOF7, SOF9-SOF11, SOF13-SOF15)
 * which all share the same frame header layout.
 * Handles 0xFF padding bytes per JPEG spec.
 */
function getJpegDimensions(data: Uint8Array): { width: number; height: number } {
  let offset = 2; // skip SOI marker (0xFFD8)
  while (offset < data.length - 1) {
    // Skip any 0xFF fill bytes
    while (offset < data.length && data[offset] === 0xff && data[offset + 1] === 0xff) {
      offset++;
    }
    if (offset >= data.length - 1 || data[offset] !== 0xff) {
      break;
    }
    const marker = data[offset + 1];

    // SOF markers: C0-C3, C5-C7, C9-CB, CD-CF (excluding C4=DHT, C8=JPG, CC=DAC)
    const isSof =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      if (offset + 8 < data.length) {
        const height = (data[offset + 5] << 8) | data[offset + 6];
        const width = (data[offset + 7] << 8) | data[offset + 8];
        return { width, height };
      }
      break;
    }

    // Skip segment: 2 byte marker + segment length (includes the 2 length bytes)
    if (offset + 3 >= data.length) {
      break;
    }
    const segLen = (data[offset + 2] << 8) | data[offset + 3];
    offset += 2 + segLen;
  }
  return { width: 1, height: 1 };
}
