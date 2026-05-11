/**
 * Word-to-PDF Bridge
 *
 * Converts a Word document (DocxDocument) to PDF.
 * Like excel-bridge.ts, this is the ONLY file in the PDF module that imports
 * from @word.
 *
 * @example
 * ```typescript
 * import { readDocx } from "excelts/word";
 * import { docxToPdf } from "excelts/pdf";
 *
 * const doc = await readDocx(docxBytes);
 * const pdfBytes = await docxToPdf(doc);
 * ```
 *
 * Note: This is a flow-based renderer with run-level formatting, images,
 * lists, tables, headers/footers, and watermarks. For pixel-perfect output,
 * consider Microsoft Word or LibreOffice.
 */

import type {
  DocxDocument,
  Paragraph,
  ParagraphChild,
  Run,
  RunProperties,
  RunContent,
  Table,
  TableCell,
  TableRow,
  BodyContent,
  ImageDef,
  Hyperlink,
  FloatingImage,
  TextBox,
  StructuredDocumentTag,
  MathBlock,
  MathContent,
  NumberingLevel,
  InlineImageContent,
  CheckBox,
  ChartContent,
  ChartExContent,
  DrawingShape,
  Chart,
  ChartSeries
} from "@word/types";

import { PdfDocumentBuilder, type PdfPageBuilder, type PathOp } from "./builder/document-builder";

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
  /** Header margin from top edge in points (default: 36). */
  readonly headerMargin?: number;
  /** Footer margin from bottom edge in points (default: 36). */
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
   * The callback receives the Word Chart definition and a PdfPageBuilder,
   * and should draw the chart into the specified rectangle.
   */
  readonly chartRenderer?: (
    chart: Chart,
    page: PdfPageBuilder,
    rect: { x: number; y: number; width: number; height: number }
  ) => void;
}

/**
 * Convert a DocxDocument to PDF bytes.
 *
 * @param doc - The DOCX document model (from readDocx or DocumentBuilder.build()).
 * @param options - PDF rendering options.
 * @returns Promise of PDF bytes.
 */
export async function docxToPdf(
  doc: DocxDocument,
  options?: DocxToPdfOptions
): Promise<Uint8Array> {
  const opts: ResolvedOptions = {
    pageWidth: options?.pageWidth ?? 612,
    pageHeight: options?.pageHeight ?? 792,
    marginTop: options?.marginTop ?? 72,
    marginBottom: options?.marginBottom ?? 72,
    marginLeft: options?.marginLeft ?? 72,
    marginRight: options?.marginRight ?? 72,
    defaultFont: options?.defaultFont ?? "Helvetica",
    defaultFontSize: options?.defaultFontSize ?? 11,
    headerMargin: options?.headerMargin ?? 36,
    footerMargin: options?.footerMargin ?? 36
  };

  // Use section properties if available and no explicit overrides
  const sectProps = doc.sectionProperties;
  if (sectProps?.pageSize && !options?.pageWidth) {
    opts.pageWidth = twipsToPt(sectProps.pageSize.width);
    opts.pageHeight = twipsToPt(sectProps.pageSize.height);
    if (sectProps.pageSize.orientation === "landscape") {
      [opts.pageWidth, opts.pageHeight] = [opts.pageHeight, opts.pageWidth];
    }
  }
  if (sectProps?.margins && !options?.marginTop) {
    opts.marginTop = twipsToPt(sectProps.margins.top);
    opts.marginBottom = twipsToPt(sectProps.margins.bottom);
    opts.marginLeft = twipsToPt(sectProps.margins.left);
    opts.marginRight = twipsToPt(sectProps.margins.right);
  }

  // Auto-detect chart support: if no explicit chartRenderer is provided,
  // try to import the high-quality Excel-based renderer.
  let chartRenderer = options?.chartRenderer;
  if (!chartRenderer) {
    try {
      const mod = await import("./excel-bridge");
      if (typeof mod.createWordChartPdfRenderer === "function") {
        chartRenderer = mod.createWordChartPdfRenderer();
      }
    } catch {
      // Chart support not available — fall through to simplified renderer
    }
  }

  const builder = new PdfDocumentBuilder();
  if (doc.coreProperties?.title || doc.coreProperties?.creator) {
    builder.setMetadata({
      title: doc.coreProperties?.title,
      author: doc.coreProperties?.creator,
      subject: doc.coreProperties?.subject
    });
  }

  const state: RenderState = {
    builder,
    opts,
    doc,
    rIdToImage: buildImageMap(doc),
    currentPage: null,
    cursorY: 0,
    currentPageIndex: 0,
    numberingCounters: new Map(),
    availableWidth: opts.pageWidth - opts.marginLeft - opts.marginRight,
    chartRenderer
  };

  newPage(state);

  // Render body content
  for (const item of doc.body) {
    renderBodyContent(state, item);
  }

  return builder.build();
}

// =============================================================================
// Internal Types
// =============================================================================

interface ResolvedOptions {
  pageWidth: number;
  pageHeight: number;
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  defaultFont: string;
  defaultFontSize: number;
  headerMargin: number;
  footerMargin: number;
}

interface RenderState {
  builder: PdfDocumentBuilder;
  opts: ResolvedOptions;
  doc: DocxDocument;
  rIdToImage: Map<string, ImageDef>;
  currentPage: PdfPageBuilder | null;
  cursorY: number;
  currentPageIndex: number;
  numberingCounters: Map<string, number>;
  availableWidth: number;
  chartRenderer?: (
    chart: Chart,
    page: PdfPageBuilder,
    rect: { x: number; y: number; width: number; height: number }
  ) => void;
}

/** Inline fragment with formatting for run-level rendering. */
interface InlineFragment {
  text: string;
  bold: boolean;
  italic: boolean;
  fontSize: number;
  fontFamily: string;
  color?: { r: number; g: number; b: number };
  underline: boolean;
  strike: boolean;
  superscript: boolean;
  subscript: boolean;
}

/** Image fragment within a paragraph. */
interface ImageFragment {
  type: "image";
  data: Uint8Array;
  format: "jpeg" | "png";
  width: number;
  height: number;
}

type ParagraphFragment = InlineFragment | ImageFragment;

// =============================================================================
// Page Management
// =============================================================================

function newPage(state: RenderState): void {
  state.currentPage = state.builder.addPage({
    width: state.opts.pageWidth,
    height: state.opts.pageHeight
  });
  state.cursorY = state.opts.pageHeight - state.opts.marginTop;
  state.currentPageIndex++;

  // Render header
  renderHeaderOnPage(state);
  // Render footer
  renderFooterOnPage(state);
  // Render watermark
  renderWatermarkOnPage(state);
}

function ensureSpace(state: RenderState, needed: number): void {
  if (state.cursorY - needed < state.opts.marginBottom) {
    newPage(state);
  }
}

function forceNewPage(state: RenderState): void {
  newPage(state);
}

// =============================================================================
// Header / Footer / Watermark
// =============================================================================

function renderHeaderOnPage(state: RenderState): void {
  const doc = state.doc;
  if (!doc.headers || doc.headers.size === 0) {
    return;
  }

  const headerDef = doc.headers.get("default") ?? doc.headers.values().next().value;
  if (!headerDef) {
    return;
  }

  const content = headerDef.content;
  const headerY = state.opts.pageHeight - state.opts.headerMargin;

  for (const child of content.children) {
    if (!("type" in child) || child.type !== "paragraph") {
      continue;
    }
    const para = child as Paragraph;
    const fragments = collectFragments(state, para);
    const textFragments = fragments.filter(
      (f): f is InlineFragment => !("type" in f && f.type === "image")
    );
    const text = textFragments.map(f => f.text).join("");
    if (!text.trim()) {
      continue;
    }

    const fontSize = textFragments[0]?.fontSize ?? state.opts.defaultFontSize;
    state.currentPage!.drawText(text, {
      x: state.opts.marginLeft,
      y: headerY - fontSize,
      fontSize,
      fontFamily: textFragments[0]?.fontFamily ?? state.opts.defaultFont,
      bold: textFragments[0]?.bold ?? false,
      italic: textFragments[0]?.italic ?? false,
      maxWidth: state.availableWidth
    });
    break; // Only render first paragraph of header
  }
}

function renderFooterOnPage(state: RenderState): void {
  const doc = state.doc;
  if (!doc.footers || doc.footers.size === 0) {
    return;
  }

  const footerDef = doc.footers.get("default") ?? doc.footers.values().next().value;
  if (!footerDef) {
    return;
  }

  const content = footerDef.content;
  const footerY = state.opts.footerMargin;

  for (const child of content.children) {
    if (!("type" in child) || child.type !== "paragraph") {
      continue;
    }
    const para = child as Paragraph;
    const fragments = collectFragments(state, para);
    const textFragments = fragments.filter(
      (f): f is InlineFragment => !("type" in f && f.type === "image")
    );
    const text = textFragments.map(f => f.text).join("");
    if (!text.trim()) {
      continue;
    }

    const fontSize = textFragments[0]?.fontSize ?? state.opts.defaultFontSize;
    const alignment = para.properties?.alignment;
    const anchor =
      alignment === "center"
        ? ("middle" as const)
        : alignment === "right"
          ? ("end" as const)
          : ("start" as const);
    const x =
      alignment === "center"
        ? state.opts.marginLeft + state.availableWidth / 2
        : alignment === "right"
          ? state.opts.pageWidth - state.opts.marginRight
          : state.opts.marginLeft;

    state.currentPage!.drawText(text, {
      x,
      y: footerY,
      fontSize,
      fontFamily: textFragments[0]?.fontFamily ?? state.opts.defaultFont,
      bold: textFragments[0]?.bold ?? false,
      italic: textFragments[0]?.italic ?? false,
      anchor
    });
    break; // Only render first paragraph of footer
  }
}

function renderWatermarkOnPage(state: RenderState): void {
  const doc = state.doc;
  if (!doc.watermark) {
    return;
  }

  const wm = doc.watermark;
  if (wm.type === "text") {
    const fontSize = wm.fontSize ? wm.fontSize / 2 : 72;
    const color = parseHexColor(wm.color ?? "C0C0C0");
    state.currentPage!.drawText(wm.text, {
      x: state.opts.pageWidth / 2,
      y: state.opts.pageHeight / 2,
      fontSize,
      color: { r: color.r, g: color.g, b: color.b, a: wm.semiTransparent ? 0.3 : 0.6 },
      rotation: wm.rotation ?? -45,
      anchor: "middle",
      fontFamily: wm.font ?? state.opts.defaultFont
    });
  }
}

// =============================================================================
// Body Content Dispatch
// =============================================================================

function renderBodyContent(state: RenderState, item: BodyContent): void {
  if (!("type" in item)) {
    return;
  }

  switch (item.type) {
    case "paragraph":
      renderParagraph(state, item);
      break;
    case "table":
      renderTable(state, item);
      break;
    case "math":
      renderMathBlock(state, item);
      break;
    case "chart":
      renderChart(state, item);
      break;
    case "chartEx":
      renderChartExPlaceholder(state, item);
      break;
    case "textBox":
      renderTextBox(state, item);
      break;
    case "sdt":
      renderSdt(state, item);
      break;
    case "floatingImage":
      renderFloatingImage(state, item);
      break;
    case "tableOfContents":
      renderToc(state, item);
      break;
    case "checkBox":
      renderCheckBox(state, item);
      break;
    case "drawingShape":
      renderDrawingShape(state, item);
      break;
    default:
      break;
  }
}

// =============================================================================
// Paragraph Rendering (run-level formatting)
// =============================================================================

function renderParagraph(state: RenderState, para: Paragraph): void {
  const props = para.properties;

  // Page break before
  if (props?.pageBreakBefore) {
    forceNewPage(state);
  }

  // Check for page break in run content
  if (hasPageBreak(para)) {
    forceNewPage(state);
  }

  // Spacing before
  const spacingBefore = props?.spacing?.before ? twipsToPt(props.spacing.before) : 0;
  if (spacingBefore > 0) {
    state.cursorY -= spacingBefore;
  }

  // Determine paragraph style
  const styleId = props?.style?.toLowerCase() ?? "";
  const headingLevel = getHeadingLevel(styleId, props?.outlineLevel);

  // Collect inline fragments
  const fragments = collectFragments(state, para);
  if (fragments.length === 0) {
    // Empty paragraph — just add line spacing
    state.cursorY -= state.opts.defaultFontSize * 1.2;
    applySpacingAfter(state, props);
    return;
  }

  // Apply heading font size overrides
  if (headingLevel > 0) {
    const headingSize = getHeadingFontSize(headingLevel);
    for (const f of fragments) {
      if (!("type" in f && f.type === "image")) {
        (f as InlineFragment).fontSize = Math.max((f as InlineFragment).fontSize, headingSize);
        (f as InlineFragment).bold = true;
      }
    }
  }

  // Calculate indentation
  const leftIndent = props?.indent?.left ? twipsToPt(props.indent.left) : 0;
  const hangingIndent = props?.indent?.hanging ? twipsToPt(props.indent.hanging) : 0;
  const firstLineIndent = props?.indent?.firstLine ? twipsToPt(props.indent.firstLine) : 0;

  // Numbering / list bullet
  let listPrefix = "";
  let listIndent = 0;
  if (props?.numbering) {
    const { numId, level } = props.numbering;
    const levelDef = resolveNumberingLevel(state.doc, numId, level);
    listIndent = (level + 1) * 18; // 18pt per level

    if (levelDef) {
      const format = levelDef.format;
      if (format === "bullet") {
        listPrefix = "\u2022 ";
      } else {
        const counterKey = `${numId}-${level}`;
        const current = (state.numberingCounters.get(counterKey) ?? 0) + 1;
        state.numberingCounters.set(counterKey, current);
        listPrefix = formatNumber(current, format) + ". ";
      }
    } else {
      listPrefix = "\u2022 ";
    }
  }

  const baseX = state.opts.marginLeft + leftIndent + listIndent;
  const effectiveWidth = state.availableWidth - leftIndent - listIndent;

  // Render list prefix
  if (listPrefix) {
    const prefixX = baseX - hangingIndent;
    const fontSize = getMaxFontSize(fragments, state.opts.defaultFontSize);
    ensureSpace(state, fontSize * 1.2);
    state.currentPage!.drawText(listPrefix, {
      x: Math.max(prefixX, state.opts.marginLeft),
      y: state.cursorY - fontSize,
      fontSize,
      fontFamily: state.opts.defaultFont,
      bold: false
    });
  }

  // Render inline fragments with line wrapping
  const lineHeight = getMaxFontSize(fragments, state.opts.defaultFontSize) * 1.2;
  const firstLineX = baseX + (listPrefix ? 0 : firstLineIndent);
  const alignment = props?.alignment ?? "left";

  renderInlineFragments(state, fragments, firstLineX, baseX, effectiveWidth, lineHeight, alignment);

  // Spacing after
  applySpacingAfter(state, props);
}

function applySpacingAfter(state: RenderState, props: Paragraph["properties"]): void {
  const spacingAfter = props?.spacing?.after ? twipsToPt(props.spacing.after) : 6;
  state.cursorY -= spacingAfter;
}

function renderInlineFragments(
  state: RenderState,
  fragments: ParagraphFragment[],
  firstLineX: number,
  subsequentX: number,
  maxWidth: number,
  lineHeight: number,
  alignment: string
): void {
  // Simple approach: concatenate text fragments and render per-run
  // For proper word-wrapping we need to measure and break
  let x = firstLineX;
  let isFirstLine = true;

  for (const fragment of fragments) {
    if ("type" in fragment && fragment.type === "image") {
      // Render inline image
      const imgFrag = fragment as ImageFragment;
      const imgWidthPt = imgFrag.width;
      const imgHeightPt = imgFrag.height;

      ensureSpace(state, imgHeightPt);
      state.currentPage!.drawImage({
        data: imgFrag.data,
        format: imgFrag.format,
        x,
        y: state.cursorY - imgHeightPt,
        width: imgWidthPt,
        height: imgHeightPt
      });
      x += imgWidthPt;
      continue;
    }

    const f = fragment as InlineFragment;
    if (!f.text) {
      continue;
    }

    const fontSize = f.fontSize;
    const measure = (s: string) =>
      state.currentPage!.measureText(s, {
        fontSize,
        fontFamily: f.fontFamily,
        bold: f.bold,
        italic: f.italic
      });

    // Word-wrap this fragment
    const words = f.text.split(/(?<=\s)/);
    for (const word of words) {
      const wordWidth = measure(word);

      // Check if word fits on current line
      if (
        x + wordWidth > state.opts.marginLeft + state.availableWidth &&
        x > (isFirstLine ? firstLineX : subsequentX)
      ) {
        // Move to next line
        state.cursorY -= lineHeight;
        x = subsequentX;
        isFirstLine = false;
        ensureSpace(state, lineHeight);
      }

      if (!word.trim() && x === (isFirstLine ? firstLineX : subsequentX)) {
        // Skip leading whitespace on new line
        continue;
      }

      ensureSpace(state, lineHeight);

      const drawY = state.cursorY - fontSize;

      state.currentPage!.drawText(word, {
        x,
        y: drawY,
        fontSize,
        fontFamily: f.fontFamily,
        bold: f.bold,
        italic: f.italic,
        color: f.color ? { r: f.color.r, g: f.color.g, b: f.color.b } : undefined
      });

      // Draw underline
      if (f.underline) {
        const ulY = drawY - 1.5;
        state.currentPage!.drawLine({
          x1: x,
          y1: ulY,
          x2: x + wordWidth,
          y2: ulY,
          color: f.color ?? { r: 0, g: 0, b: 0 },
          lineWidth: 0.5
        });
      }

      // Draw strikethrough
      if (f.strike) {
        const stY = drawY + fontSize * 0.35;
        state.currentPage!.drawLine({
          x1: x,
          y1: stY,
          x2: x + wordWidth,
          y2: stY,
          color: f.color ?? { r: 0, g: 0, b: 0 },
          lineWidth: 0.5
        });
      }

      x += wordWidth;
    }
  }

  // Move cursor past the rendered line
  state.cursorY -= lineHeight;
}

// =============================================================================
// Table Rendering
// =============================================================================

function renderTable(state: RenderState, table: Table): void {
  if (table.rows.length === 0) {
    return;
  }

  const tableWidth = state.availableWidth;
  const numCols = Math.max(...table.rows.map(r => r.cells.length));

  // Calculate column widths
  const colWidths = calculateColumnWidths(table, numCols, tableWidth);

  const cellPadding = 4;

  // Render each row
  for (const row of table.rows) {
    const rowHeight = calculateRowHeight(state, row, colWidths);
    ensureSpace(state, rowHeight);

    let x = state.opts.marginLeft;
    const rowTop = state.cursorY;
    const rowBottom = state.cursorY - rowHeight;

    for (let colIdx = 0; colIdx < row.cells.length; colIdx++) {
      const cell = row.cells[colIdx];
      const span = cell.properties?.gridSpan ?? 1;
      let cellWidth = 0;
      for (let s = 0; s < span && colIdx + s < colWidths.length; s++) {
        cellWidth += colWidths[colIdx + s];
      }

      // Cell background
      const shading = cell.properties?.shading;
      if (shading?.fill && shading.fill !== "auto") {
        const bgColor = parseHexColor(shading.fill);
        state.currentPage!.drawRect({
          x,
          y: rowBottom,
          width: cellWidth,
          height: rowHeight,
          fill: bgColor
        });
      }

      // Cell border
      state.currentPage!.drawRect({
        x,
        y: rowBottom,
        width: cellWidth,
        height: rowHeight,
        stroke: { r: 0, g: 0, b: 0 },
        lineWidth: 0.5
      });

      // Determine content height for vertical alignment
      const contentHeight = measureCellContentHeight(state, cell, cellWidth - cellPadding * 2);
      const vAlign = cell.properties?.verticalAlign ?? "top";
      let contentStartY: number;
      if (vAlign === "center") {
        contentStartY = rowTop - (rowHeight - contentHeight) / 2;
      } else if (vAlign === "bottom") {
        contentStartY = rowBottom + contentHeight;
      } else {
        // top
        contentStartY = rowTop - cellPadding;
      }

      // Render cell content (paragraphs and nested tables)
      renderCellContent(state, cell, x + cellPadding, contentStartY, cellWidth - cellPadding * 2);

      x += cellWidth;
    }

    state.cursorY = rowBottom;
  }

  state.cursorY -= 6;
}

/**
 * Measure the total height that cell content will occupy.
 */
function measureCellContentHeight(state: RenderState, cell: TableCell, availWidth: number): number {
  let totalHeight = 0;

  for (const item of cell.content) {
    if ("type" in item && item.type === "paragraph") {
      const para = item as Paragraph;
      const fragments = collectFragments(state, para);
      if (fragments.length === 0) {
        totalHeight += state.opts.defaultFontSize * 1.2;
        continue;
      }
      const lineHeight = getMaxFontSize(fragments, state.opts.defaultFontSize) * 1.2;
      // Estimate number of lines by measuring text widths
      let lineCount = 1;
      let lineX = 0;
      for (const fragment of fragments) {
        if ("type" in fragment && fragment.type === "image") {
          const imgFrag = fragment as ImageFragment;
          totalHeight += imgFrag.height;
          continue;
        }
        const f = fragment as InlineFragment;
        if (!f.text) {
          continue;
        }
        const words = f.text.split(/(?<=\s)/);
        for (const word of words) {
          const wordWidth = state.currentPage!.measureText(word, {
            fontSize: f.fontSize,
            fontFamily: f.fontFamily,
            bold: f.bold,
            italic: f.italic
          });
          if (lineX + wordWidth > availWidth && lineX > 0) {
            lineCount++;
            lineX = 0;
          }
          lineX += wordWidth;
        }
      }
      totalHeight += lineCount * lineHeight;
      // Add spacing after paragraph
      const spacingAfter = para.properties?.spacing?.after
        ? twipsToPt(para.properties.spacing.after)
        : 4;
      totalHeight += spacingAfter;
    } else if ("type" in item && item.type === "table") {
      // Nested table: estimate height as sum of row min heights
      const nested = item as Table;
      totalHeight += estimateNestedTableHeight(state, nested, availWidth);
    }
  }

  return totalHeight;
}

/**
 * Estimate the height of a nested table for layout purposes.
 */
function estimateNestedTableHeight(state: RenderState, table: Table, availWidth: number): number {
  if (table.rows.length === 0) {
    return 0;
  }
  const numCols = Math.max(...table.rows.map(r => r.cells.length));
  const colWidths = calculateColumnWidths(table, numCols, availWidth);
  let totalHeight = 0;
  for (const row of table.rows) {
    totalHeight += calculateRowHeight(state, row, colWidths);
  }
  return totalHeight;
}

/**
 * Render cell content at a specific position, handling paragraphs and nested tables.
 * This uses renderInlineFragments for run-level formatting support.
 */
function renderCellContent(
  state: RenderState,
  cell: TableCell,
  startX: number,
  startY: number,
  availWidth: number
): void {
  // Save state
  const savedCursorY = state.cursorY;
  const savedMarginLeft = state.opts.marginLeft;
  const savedAvailableWidth = state.availableWidth;

  // Set up cell rendering context
  state.cursorY = startY;
  state.opts.marginLeft = startX;
  state.availableWidth = availWidth;

  for (const item of cell.content) {
    if ("type" in item && item.type === "paragraph") {
      const para = item as Paragraph;
      const fragments = collectFragments(state, para);
      if (fragments.length === 0) {
        state.cursorY -= state.opts.defaultFontSize * 1.2;
        continue;
      }
      const lineHeight = getMaxFontSize(fragments, state.opts.defaultFontSize) * 1.2;
      const alignment = para.properties?.alignment ?? "left";
      renderInlineFragments(state, fragments, startX, startX, availWidth, lineHeight, alignment);
      // Spacing after paragraph
      const spacingAfter = para.properties?.spacing?.after
        ? twipsToPt(para.properties.spacing.after)
        : 4;
      state.cursorY -= spacingAfter;
    } else if ("type" in item && item.type === "table") {
      // Recursively render nested tables
      const nested = item as Table;
      const nestedSavedMargin = state.opts.marginLeft;
      const nestedSavedWidth = state.availableWidth;
      state.opts.marginLeft = startX;
      state.availableWidth = availWidth;
      renderTable(state, nested);
      state.opts.marginLeft = nestedSavedMargin;
      state.availableWidth = nestedSavedWidth;
    }
  }

  // Restore state
  state.cursorY = savedCursorY;
  state.opts.marginLeft = savedMarginLeft;
  state.availableWidth = savedAvailableWidth;
}

function calculateColumnWidths(table: Table, numCols: number, tableWidth: number): number[] {
  // Try to use explicit column widths from table grid
  if (table.columnWidths && table.columnWidths.length > 0) {
    const totalTwips = table.columnWidths.reduce((s, w) => s + w, 0);
    if (totalTwips > 0) {
      return table.columnWidths.map(w => (w / totalTwips) * tableWidth);
    }
  }

  // Try to use cell widths from first row
  if (table.rows.length > 0) {
    const firstRow = table.rows[0];
    const cellWidths: number[] = [];
    let totalWidth = 0;
    let hasCellWidths = false;

    for (const cell of firstRow.cells) {
      if (cell.properties?.width?.value) {
        const w =
          cell.properties.width.type === "dxa"
            ? twipsToPt(cell.properties.width.value)
            : cell.properties.width.type === "pct"
              ? (cell.properties.width.value / 5000) * tableWidth
              : tableWidth / numCols;
        cellWidths.push(w);
        totalWidth += w;
        hasCellWidths = true;
      } else {
        cellWidths.push(tableWidth / numCols);
        totalWidth += tableWidth / numCols;
      }
    }

    if (hasCellWidths && totalWidth > 0) {
      // Normalize to fit table width
      const scale = tableWidth / totalWidth;
      return cellWidths.map(w => w * scale);
    }
  }

  // Default: equal widths
  return new Array(numCols).fill(tableWidth / numCols);
}

function calculateRowHeight(state: RenderState, row: TableRow, colWidths: number[]): number {
  let maxHeight = 20; // minimum row height
  const cellPadding = 4;

  for (let i = 0; i < row.cells.length; i++) {
    const cell = row.cells[i];
    const span = cell.properties?.gridSpan ?? 1;
    let cellWidth = 0;
    for (let s = 0; s < span && i + s < colWidths.length; s++) {
      cellWidth += colWidths[i + s];
    }
    const availWidth = cellWidth - cellPadding * 2;
    const contentHeight = measureCellContentHeight(state, cell, availWidth);
    maxHeight = Math.max(maxHeight, contentHeight + cellPadding * 2);
  }

  // Respect explicit row height
  if (row.properties?.height) {
    const explicitH = twipsToPt(row.properties.height.value);
    if (row.properties.height.rule === "exact") {
      return explicitH;
    }
    maxHeight = Math.max(maxHeight, explicitH);
  }

  return maxHeight;
}

// =============================================================================
// Special Content Rendering
// =============================================================================

function renderTextBox(state: RenderState, textBox: TextBox): void {
  // Render text box content as indented block
  const indent = 12;
  const originalMarginLeft = state.opts.marginLeft;
  const originalWidth = state.availableWidth;

  state.opts.marginLeft += indent;
  state.availableWidth -= indent * 2;

  // Draw left border
  const startY = state.cursorY;

  try {
    for (const p of textBox.content) {
      renderParagraph(state, p);
    }
  } finally {
    state.opts.marginLeft = originalMarginLeft;
    state.availableWidth = originalWidth;
  }

  // Draw border line
  state.currentPage!.drawLine({
    x1: originalMarginLeft + indent - 4,
    y1: startY,
    x2: originalMarginLeft + indent - 4,
    y2: state.cursorY,
    color: { r: 0.6, g: 0.6, b: 0.6 },
    lineWidth: 1.5
  });
  state.cursorY -= 6;
}

function renderSdt(state: RenderState, sdt: StructuredDocumentTag): void {
  for (const child of sdt.content) {
    if ("type" in child) {
      if (child.type === "paragraph" || child.type === "table") {
        renderBodyContent(state, child as BodyContent);
      }
    }
  }
}

function renderFloatingImage(state: RenderState, fi: FloatingImage): void {
  if (!fi.rId) {
    return;
  }
  const img = state.rIdToImage.get(fi.rId);
  if (!img) {
    return;
  }

  const format = imageMediaToFormat(img.mediaType);
  if (!format) {
    return;
  }

  const imgData = img.mediaType === "svg" && img.fallbackData ? img.fallbackData : img.data;
  const widthPt = fi.width ? emuToPt(fi.width) : 200;
  const heightPt = fi.height ? emuToPt(fi.height) : 150;

  ensureSpace(state, heightPt + 6);

  state.currentPage!.drawImage({
    data: imgData,
    format,
    x: state.opts.marginLeft,
    y: state.cursorY - heightPt,
    width: widthPt,
    height: heightPt
  });

  state.cursorY -= heightPt + 6;
}

function renderToc(state: RenderState, toc: any): void {
  if (toc.cachedParagraphs) {
    for (const p of toc.cachedParagraphs) {
      renderBodyContent(state, p);
    }
  }
}

function renderCheckBox(state: RenderState, cb: CheckBox): void {
  const fontSize = state.opts.defaultFontSize;
  ensureSpace(state, fontSize * 1.2);
  const symbol = cb.checked ? "\u2611" : "\u2610";
  state.currentPage!.drawText(symbol, {
    x: state.opts.marginLeft,
    y: state.cursorY - fontSize,
    fontSize,
    fontFamily: state.opts.defaultFont
  });
  state.cursorY -= fontSize * 1.2;
}

function renderMathBlock(state: RenderState, block: MathBlock): void {
  const text = extractMathText(block.content);
  if (text.trim()) {
    renderPlaceholder(state, text);
  }
}

function renderPlaceholder(state: RenderState, text: string): void {
  const fontSize = state.opts.defaultFontSize;
  ensureSpace(state, fontSize * 1.2);
  state.currentPage!.drawText(text, {
    x: state.opts.marginLeft,
    y: state.cursorY - fontSize,
    fontSize,
    fontFamily: state.opts.defaultFont,
    italic: true,
    color: { r: 0.4, g: 0.4, b: 0.4 },
    maxWidth: state.availableWidth
  });
  state.cursorY -= fontSize * 1.2 + 6;
}

// =============================================================================
// Fragment Collection (run-level formatting)
// =============================================================================

function collectFragments(state: RenderState, para: Paragraph): ParagraphFragment[] {
  const fragments: ParagraphFragment[] = [];

  for (const child of para.children) {
    collectChildFragments(state, child, fragments);
  }

  return fragments;
}

function collectChildFragments(
  state: RenderState,
  child: ParagraphChild,
  fragments: ParagraphFragment[]
): void {
  if ("type" in child) {
    switch (child.type) {
      case "hyperlink": {
        const link = child as Hyperlink;
        for (const hChild of link.children) {
          collectChildFragments(state, hChild, fragments);
        }
        break;
      }
      case "insertedRun": {
        const ins = child as any;
        if (ins.run) {
          collectRunFragments(state, ins.run, fragments);
        }
        break;
      }
      case "deletedRun":
        // Skip deleted content
        break;
      case "movedToRun": {
        const moved = child as any;
        if (moved.run) {
          collectRunFragments(state, moved.run, fragments);
        }
        break;
      }
      case "movedFromRun":
        break;
      default:
        break;
    }
  } else if ("content" in child) {
    collectRunFragments(state, child as Run, fragments);
  }
}

function collectRunFragments(state: RenderState, run: Run, fragments: ParagraphFragment[]): void {
  const rPr = run.properties;
  const fontSize = resolveRunFontSize(rPr, state.opts.defaultFontSize);
  const fontFamily = resolveRunFont(rPr, state.opts.defaultFont);
  const bold = rPr?.bold ?? false;
  const italic = rPr?.italic ?? false;
  const color = resolveRunColor(rPr);
  const underline = !!(rPr?.underline && rPr.underline !== "none");
  const strike = rPr?.strike ?? false;
  const superscript = rPr?.vertAlign === "superscript";
  const subscript = rPr?.vertAlign === "subscript";

  for (const content of run.content) {
    const result = renderRunContent(state, content);
    if (result === null) {
      continue;
    }

    if (typeof result === "string") {
      if (result) {
        fragments.push({
          text: result,
          bold,
          italic,
          fontSize: superscript || subscript ? fontSize * 0.7 : fontSize,
          fontFamily,
          color,
          underline,
          strike,
          superscript,
          subscript
        });
      }
    } else {
      // Image fragment
      fragments.push(result);
    }
  }
}

function renderRunContent(state: RenderState, content: RunContent): string | ImageFragment | null {
  switch (content.type) {
    case "text":
      return content.text;
    case "tab":
      return "    ";
    case "ptab":
      return "    ";
    case "break":
      if (content.breakType === "page") {
        // Will be handled by hasPageBreak check
        return "\n";
      }
      return "\n";
    case "carriageReturn":
      return "\n";
    case "noBreakHyphen":
      return "\u2011";
    case "softHyphen":
      return "";
    case "symbol":
      try {
        return String.fromCodePoint(parseInt(content.char, 16));
      } catch {
        return content.char;
      }
    case "field":
      return content.cachedValue ?? "";
    case "footnoteRef":
      return `[${content.id}]`;
    case "endnoteRef":
      return `[${content.id}]`;
    case "image":
      return renderInlineImage(state, content as InlineImageContent);
    case "ruby":
      // Render base text only
      return content.baseText
        .map(r => {
          let text = "";
          for (const c of r.content) {
            if (c.type === "text") {
              text += c.text;
            }
          }
          return text;
        })
        .join("");
    case "lastRenderedPageBreak":
    case "annotationReference":
      return null;
    default:
      return null;
  }
}

function renderInlineImage(state: RenderState, img: InlineImageContent): ImageFragment | null {
  if (!img.rId) {
    return null;
  }
  const imgDef = state.rIdToImage.get(img.rId);
  if (!imgDef) {
    return null;
  }

  const format = imageMediaToFormat(imgDef.mediaType);
  if (!format) {
    return null;
  }

  const imgData =
    imgDef.mediaType === "svg" && imgDef.fallbackData ? imgDef.fallbackData : imgDef.data;
  const widthPt = img.width ? emuToPt(img.width) : 100;
  const heightPt = img.height ? emuToPt(img.height) : 75;

  return {
    type: "image",
    data: imgData,
    format,
    width: Math.min(widthPt, state.availableWidth),
    height: heightPt * (Math.min(widthPt, state.availableWidth) / widthPt)
  };
}

// =============================================================================
// Numbering Resolution
// =============================================================================

function resolveNumberingLevel(
  doc: DocxDocument,
  numId: number,
  level: number
): NumberingLevel | undefined {
  if (!doc.numberingInstances || !doc.abstractNumberings) {
    return undefined;
  }

  const instance = doc.numberingInstances.find(n => n.numId === numId);
  if (!instance) {
    return undefined;
  }

  // Check level override
  if (instance.overrides) {
    const override = instance.overrides.find(o => o.level === level);
    if (override?.levelDef) {
      return override.levelDef;
    }
  }

  const abstract = doc.abstractNumberings.find(a => a.abstractNumId === instance.abstractNumId);
  if (!abstract) {
    return undefined;
  }

  return abstract.levels.find(l => l.level === level);
}

function formatNumber(num: number, format: string): string {
  switch (format) {
    case "decimal":
      return String(num);
    case "upperRoman":
      return toRoman(num).toUpperCase();
    case "lowerRoman":
      return toRoman(num).toLowerCase();
    case "upperLetter":
      return String.fromCharCode(64 + ((num - 1) % 26) + 1);
    case "lowerLetter":
      return String.fromCharCode(96 + ((num - 1) % 26) + 1);
    default:
      return String(num);
  }
}

function toRoman(n: number): string {
  const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
  const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
  let result = "";
  let remaining = n;
  for (let i = 0; i < vals.length; i++) {
    while (remaining >= vals[i]) {
      result += syms[i];
      remaining -= vals[i];
    }
  }
  return result;
}

// =============================================================================
// Style Resolution
// =============================================================================

function resolveRunFontSize(rPr: RunProperties | undefined, defaultSize: number): number {
  if (rPr?.size) {
    // size in half-points
    return rPr.size / 2;
  }
  return defaultSize;
}

function resolveRunFont(rPr: RunProperties | undefined, defaultFont: string): string {
  if (!rPr?.font) {
    return defaultFont;
  }
  if (typeof rPr.font === "string") {
    return rPr.font;
  }
  return rPr.font.ascii ?? rPr.font.hAnsi ?? rPr.font.cs ?? defaultFont;
}

function resolveRunColor(
  rPr: RunProperties | undefined
): { r: number; g: number; b: number } | undefined {
  if (!rPr?.color) {
    return undefined;
  }
  if (typeof rPr.color === "string") {
    if (rPr.color === "auto" || rPr.color === "000000") {
      return undefined;
    }
    return parseHexColor(rPr.color);
  }
  if (rPr.color.val && rPr.color.val !== "auto") {
    return parseHexColor(rPr.color.val);
  }
  return undefined;
}

// =============================================================================
// Heading Helpers
// =============================================================================

function getHeadingLevel(styleId: string, outlineLevel?: number): number {
  if (outlineLevel !== undefined && outlineLevel >= 0) {
    return outlineLevel + 1;
  }
  if (styleId === "heading1" || styleId === "heading 1" || styleId === "title") {
    return 1;
  }
  if (styleId === "heading2" || styleId === "heading 2") {
    return 2;
  }
  if (styleId === "heading3" || styleId === "heading 3") {
    return 3;
  }
  if (styleId === "heading4" || styleId === "heading 4") {
    return 4;
  }
  if (styleId === "heading5" || styleId === "heading 5") {
    return 5;
  }
  if (styleId === "heading6" || styleId === "heading 6") {
    return 6;
  }
  const m = /^heading\s*(\d)$/i.exec(styleId);
  if (m) {
    return parseInt(m[1]);
  }
  return 0;
}

function getHeadingFontSize(level: number): number {
  switch (level) {
    case 1:
      return 24;
    case 2:
      return 20;
    case 3:
      return 16;
    case 4:
      return 14;
    case 5:
      return 12;
    case 6:
      return 11;
    default:
      return 11;
  }
}

// =============================================================================
// Utility Functions
// =============================================================================

function twipsToPt(twips: number): number {
  return twips / 20;
}

function emuToPt(emu: number): number {
  return (emu / 914400) * 72;
}

function parseHexColor(hex: string): { r: number; g: number; b: number } {
  const clean = hex.replace("#", "");
  if (clean.length < 6) {
    return { r: 0, g: 0, b: 0 };
  }
  const r = parseInt(clean.slice(0, 2), 16) / 255;
  const g = parseInt(clean.slice(2, 4), 16) / 255;
  const b = parseInt(clean.slice(4, 6), 16) / 255;
  return {
    r: Number.isFinite(r) ? r : 0,
    g: Number.isFinite(g) ? g : 0,
    b: Number.isFinite(b) ? b : 0
  };
}

function imageMediaToFormat(mediaType: string): "jpeg" | "png" | null {
  if (mediaType === "png") {
    return "png";
  }
  if (mediaType === "jpeg") {
    return "jpeg";
  }
  // SVG uses fallback PNG; other formats not directly supported
  if (mediaType === "svg") {
    return "png";
  }
  return null;
}

function getMaxFontSize(fragments: ParagraphFragment[], defaultSize: number): number {
  let max = defaultSize;
  for (const f of fragments) {
    if (!("type" in f && f.type === "image")) {
      max = Math.max(max, (f as InlineFragment).fontSize);
    }
  }
  return max;
}

function hasPageBreak(para: Paragraph): boolean {
  for (const child of para.children) {
    if (!("content" in child)) {
      continue;
    }
    const run = child as Run;
    for (const c of run.content) {
      if (c.type === "break" && c.breakType === "page") {
        return true;
      }
    }
  }
  return false;
}

function buildImageMap(doc: DocxDocument): Map<string, ImageDef> {
  const map = new Map<string, ImageDef>();
  if (doc.images) {
    for (const img of doc.images) {
      if (img.rId) {
        map.set(img.rId, img);
      }
    }
  }
  return map;
}

function extractMathText(contents: readonly MathContent[]): string {
  let result = "";
  for (const item of contents) {
    switch (item.type) {
      case "mathRun":
        result += item.text;
        break;
      case "mathFraction":
        result += extractMathText(item.numerator) + "/" + extractMathText(item.denominator);
        break;
      case "mathSuperScript":
        result += extractMathText(item.base) + "^" + extractMathText(item.superScript);
        break;
      case "mathSubScript":
        result += extractMathText(item.base) + "_" + extractMathText(item.subScript);
        break;
      case "mathRadical":
        result += "\u221A(" + extractMathText(item.content) + ")";
        break;
      case "mathDelimiter":
        result += item.beginChar ?? "(";
        for (let i = 0; i < item.content.length; i++) {
          if (i > 0) {
            result += item.separatorChar ?? ",";
          }
          result += extractMathText(item.content[i]);
        }
        result += item.endChar ?? ")";
        break;
      case "mathNary":
        result += item.char ?? "\u2211";
        result += extractMathText(item.content);
        break;
      case "mathFunction":
        result += extractMathText(item.name) + "(" + extractMathText(item.content) + ")";
        break;
      case "mathMatrix":
        for (const row of item.rows) {
          for (const cell of row) {
            result += extractMathText(cell) + " ";
          }
        }
        break;
      default:
        // Other math structures: try to extract recursively
        if ("content" in item && Array.isArray(item.content)) {
          result += extractMathText(item.content);
        }
        if ("base" in item && Array.isArray(item.base)) {
          result += extractMathText(item.base);
        }
        break;
    }
  }
  return result;
}

// =============================================================================
// Chart Rendering
// =============================================================================

/** Default series colors when no color is specified. */
const DEFAULT_CHART_COLORS: readonly string[] = [
  "4472C4",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "5B9BD5",
  "70AD47",
  "264478",
  "9E480E",
  "636363",
  "997300"
];

function renderChart(state: RenderState, content: ChartContent): void {
  const chart = content.chart;
  const defaultWidthEmu = 5486400; // 6 inches
  const defaultHeightEmu = 3657600; // 4 inches

  let widthPt = emuToPt(chart.width ?? defaultWidthEmu);
  let heightPt = emuToPt(chart.height ?? defaultHeightEmu);

  // Scale down if wider than available space
  if (widthPt > state.availableWidth) {
    const scale = state.availableWidth / widthPt;
    widthPt = state.availableWidth;
    heightPt *= scale;
  }

  ensureSpace(state, heightPt + 12);

  const page = state.currentPage!;
  const chartX = state.opts.marginLeft;
  const chartTop = state.cursorY;
  const chartBottom = chartTop - heightPt;

  // If an injected high-quality chart renderer is available, use it
  if (state.chartRenderer) {
    state.chartRenderer(chart, page, {
      x: chartX,
      y: chartBottom,
      width: widthPt,
      height: heightPt
    });
    state.cursorY = chartBottom - 12;
    return;
  }

  // Fall through to built-in simplified renderer
  // Draw chart area background
  const chartAreaColor = chart.chartAreaColor
    ? parseHexColor(chart.chartAreaColor)
    : { r: 1, g: 1, b: 1 };
  page.drawRect({
    x: chartX,
    y: chartBottom,
    width: widthPt,
    height: heightPt,
    fill: chartAreaColor,
    stroke: { r: 0.8, g: 0.8, b: 0.8 },
    lineWidth: 0.5
  });

  // Layout: margins within the chart area
  const titleHeight = chart.title ? 20 : 0;
  const legendHeight = chart.legend && chart.legend !== "none" ? 18 : 0;
  const axisPadding = 40; // space for axis labels
  const topPad = 10 + titleHeight;
  const bottomPad = 10 + axisPadding + (chart.legend === "b" ? legendHeight : 0);
  const leftPad = axisPadding + 10;
  const rightPad = 10 + (chart.legend === "r" ? 60 : 0);

  const plotX = chartX + leftPad;
  const plotY = chartBottom + bottomPad;
  const plotW = widthPt - leftPad - rightPad;
  const plotH = heightPt - topPad - bottomPad;

  if (plotW <= 0 || plotH <= 0) {
    // Too small to render meaningfully
    state.cursorY = chartBottom - 6;
    return;
  }

  // Draw plot area background
  if (chart.plotAreaColor) {
    const plotBg = parseHexColor(chart.plotAreaColor);
    page.drawRect({ x: plotX, y: plotY, width: plotW, height: plotH, fill: plotBg });
  }

  // Draw title
  if (chart.title) {
    page.drawText(chart.title, {
      x: chartX + widthPt / 2,
      y: chartTop - 16,
      fontSize: 11,
      bold: true,
      fontFamily: state.opts.defaultFont,
      anchor: "middle"
    });
  }

  // Dispatch to type-specific renderer
  const chartType = chart.type;
  if (chartType === "pie" || chartType === "pie3D" || chartType === "doughnut") {
    renderPieChart(page, state, chart, plotX, plotY, plotW, plotH);
  } else if (chartType === "scatter" || chartType === "scatterSmooth" || chartType === "bubble") {
    renderScatterChart(page, state, chart, plotX, plotY, plotW, plotH);
  } else if (chartType === "line" || chartType === "lineStacked" || chartType === "lineMarked") {
    renderAxesAndGrid(page, state, chart, plotX, plotY, plotW, plotH);
    renderLineChart(page, chart, plotX, plotY, plotW, plotH);
  } else if (chartType === "area" || chartType === "areaStacked") {
    renderAxesAndGrid(page, state, chart, plotX, plotY, plotW, plotH);
    renderAreaChart(page, chart, plotX, plotY, plotW, plotH);
  } else {
    // bar, column, barStacked, columnStacked, etc.
    renderAxesAndGrid(page, state, chart, plotX, plotY, plotW, plotH);
    renderBarColumnChart(page, chart, plotX, plotY, plotW, plotH);
  }

  // Draw legend
  if (chart.legend && chart.legend !== "none" && chart.series.length > 0) {
    renderChartLegend(page, state, chart, chartX, chartBottom, widthPt, heightPt);
  }

  state.cursorY = chartBottom - 6;
}

function renderAxesAndGrid(
  page: PdfPageBuilder,
  state: RenderState,
  chart: Chart,
  plotX: number,
  plotY: number,
  plotW: number,
  plotH: number
): void {
  const gridColor = { r: 0.9, g: 0.9, b: 0.9 };
  const axisColor = { r: 0.3, g: 0.3, b: 0.3 };
  const labelFontSize = 7;
  const gridLines = 5;

  // Draw horizontal grid lines + Y-axis labels
  const { min: yMin, max: yMax } = getValueRange(chart);
  for (let i = 0; i <= gridLines; i++) {
    const ratio = i / gridLines;
    const lineY = plotY + ratio * plotH;

    // Grid line
    page.drawLine({
      x1: plotX,
      y1: lineY,
      x2: plotX + plotW,
      y2: lineY,
      color: gridColor,
      lineWidth: 0.5
    });

    // Y-axis label
    const value = yMin + ratio * (yMax - yMin);
    const label = formatAxisValue(value);
    page.drawText(label, {
      x: plotX - 4,
      y: lineY - 3,
      fontSize: labelFontSize,
      fontFamily: state.opts.defaultFont,
      anchor: "end",
      color: axisColor
    });
  }

  // Draw X-axis and Y-axis lines
  page.drawLine({
    x1: plotX,
    y1: plotY,
    x2: plotX + plotW,
    y2: plotY,
    color: axisColor,
    lineWidth: 0.75
  });
  page.drawLine({
    x1: plotX,
    y1: plotY,
    x2: plotX,
    y2: plotY + plotH,
    color: axisColor,
    lineWidth: 0.75
  });

  // Draw X-axis labels (categories)
  const categories = chart.series.length > 0 ? chart.series[0].categories : [];
  if (categories.length > 0) {
    const maxLabels = Math.min(categories.length, Math.floor(plotW / 30));
    const step = Math.max(1, Math.ceil(categories.length / maxLabels));
    for (let i = 0; i < categories.length; i += step) {
      const ratio = (i + 0.5) / categories.length;
      const labelX = plotX + ratio * plotW;
      page.drawText(truncateLabel(categories[i], 10), {
        x: labelX,
        y: plotY - 10,
        fontSize: labelFontSize,
        fontFamily: state.opts.defaultFont,
        anchor: "middle",
        color: axisColor
      });
    }
  }

  // Axis titles
  if (chart.categoryAxis?.title) {
    page.drawText(chart.categoryAxis.title, {
      x: plotX + plotW / 2,
      y: plotY - 25,
      fontSize: 8,
      fontFamily: state.opts.defaultFont,
      anchor: "middle",
      color: axisColor
    });
  }
  if (chart.valueAxis?.title) {
    page.drawText(chart.valueAxis.title, {
      x: plotX - 30,
      y: plotY + plotH / 2,
      fontSize: 8,
      fontFamily: state.opts.defaultFont,
      anchor: "middle",
      color: axisColor,
      rotation: 90
    });
  }
}

function renderBarColumnChart(
  page: PdfPageBuilder,
  chart: Chart,
  plotX: number,
  plotY: number,
  plotW: number,
  plotH: number
): void {
  const series = chart.series;
  if (series.length === 0) {
    return;
  }

  const isHorizontal =
    chart.type === "bar" || chart.type === "barStacked" || chart.type === "barPercentStacked";
  const isStacked =
    chart.type === "barStacked" ||
    chart.type === "columnStacked" ||
    chart.type === "barPercentStacked" ||
    chart.type === "columnPercentStacked";

  const { min: yMin, max: yMax } = getValueRange(chart);
  const valueRange = yMax - yMin || 1;
  const numCategories = series[0].categories.length || 1;

  if (isHorizontal) {
    // Horizontal bars
    const barGroupHeight = plotH / numCategories;
    const barGap = barGroupHeight * 0.2;
    const barHeight = isStacked
      ? barGroupHeight - barGap
      : (barGroupHeight - barGap) / series.length;

    for (let catIdx = 0; catIdx < numCategories; catIdx++) {
      if (isStacked) {
        let positiveOffset = 0;
        let negativeOffset = 0;
        for (let sIdx = 0; sIdx < series.length; sIdx++) {
          const value = series[sIdx].values[catIdx] ?? 0;
          const barW = (Math.abs(value) / valueRange) * plotW;
          const color = getSeriesColor(series[sIdx], sIdx);
          const baseX = plotX + (Math.max(0, -yMin) / valueRange) * plotW;
          const offset = value >= 0 ? positiveOffset : -(negativeOffset + barW);
          const y = plotY + plotH - (catIdx + 1) * barGroupHeight + barGap / 2;
          page.drawRect({
            x: baseX + offset,
            y,
            width: barW,
            height: barHeight,
            fill: color
          });
          if (value >= 0) {
            positiveOffset += barW;
          } else {
            negativeOffset += barW;
          }
        }
      } else {
        for (let sIdx = 0; sIdx < series.length; sIdx++) {
          const value = series[sIdx].values[catIdx] ?? 0;
          const barW = ((value - yMin) / valueRange) * plotW;
          const color = getSeriesColor(series[sIdx], sIdx);
          const y = plotY + plotH - (catIdx + 1) * barGroupHeight + barGap / 2 + sIdx * barHeight;
          page.drawRect({ x: plotX, y, width: Math.max(0, barW), height: barHeight, fill: color });
        }
      }
    }
  } else {
    // Vertical columns
    const barGroupWidth = plotW / numCategories;
    const barGap = barGroupWidth * 0.2;
    const barWidth = isStacked ? barGroupWidth - barGap : (barGroupWidth - barGap) / series.length;

    for (let catIdx = 0; catIdx < numCategories; catIdx++) {
      if (isStacked) {
        let positiveOffset = 0;
        let negativeOffset = 0;
        for (let sIdx = 0; sIdx < series.length; sIdx++) {
          const value = series[sIdx].values[catIdx] ?? 0;
          const barH = (Math.abs(value) / valueRange) * plotH;
          const color = getSeriesColor(series[sIdx], sIdx);
          const baseY = plotY + (-yMin / valueRange) * plotH;
          const offset = value >= 0 ? positiveOffset : -(negativeOffset + barH);
          const x = plotX + catIdx * barGroupWidth + barGap / 2;
          page.drawRect({
            x,
            y: baseY + offset,
            width: barWidth,
            height: barH,
            fill: color
          });
          if (value >= 0) {
            positiveOffset += barH;
          } else {
            negativeOffset += barH;
          }
        }
      } else {
        for (let sIdx = 0; sIdx < series.length; sIdx++) {
          const value = series[sIdx].values[catIdx] ?? 0;
          const barH = ((value - yMin) / valueRange) * plotH;
          const color = getSeriesColor(series[sIdx], sIdx);
          const x = plotX + catIdx * barGroupWidth + barGap / 2 + sIdx * barWidth;
          page.drawRect({ x, y: plotY, width: barWidth, height: Math.max(0, barH), fill: color });
        }
      }
    }
  }
}

function renderLineChart(
  page: PdfPageBuilder,
  chart: Chart,
  plotX: number,
  plotY: number,
  plotW: number,
  plotH: number
): void {
  const series = chart.series;
  if (series.length === 0) {
    return;
  }

  const { min: yMin, max: yMax } = getValueRange(chart);
  const valueRange = yMax - yMin || 1;
  const showMarkers = chart.type === "lineMarked";

  for (let sIdx = 0; sIdx < series.length; sIdx++) {
    const s = series[sIdx];
    const color = getSeriesColor(s, sIdx);
    const numPoints = s.values.length;
    if (numPoints === 0) {
      continue;
    }

    // Draw line segments
    for (let i = 0; i < numPoints - 1; i++) {
      const x1 = plotX + ((i + 0.5) / numPoints) * plotW;
      const y1 = plotY + ((s.values[i] - yMin) / valueRange) * plotH;
      const x2 = plotX + ((i + 1.5) / numPoints) * plotW;
      const y2 = plotY + ((s.values[i + 1] - yMin) / valueRange) * plotH;
      page.drawLine({ x1, y1, x2, y2, color, lineWidth: 1.5 });
    }

    // Draw markers
    if (showMarkers) {
      for (let i = 0; i < numPoints; i++) {
        const x = plotX + ((i + 0.5) / numPoints) * plotW;
        const y = plotY + ((s.values[i] - yMin) / valueRange) * plotH;
        page.drawCircle({
          cx: x,
          cy: y,
          r: 3,
          fill: color,
          stroke: { r: 1, g: 1, b: 1 },
          lineWidth: 1
        });
      }
    }
  }
}

function renderAreaChart(
  page: PdfPageBuilder,
  chart: Chart,
  plotX: number,
  plotY: number,
  plotW: number,
  plotH: number
): void {
  const series = chart.series;
  if (series.length === 0) {
    return;
  }

  const { min: yMin, max: yMax } = getValueRange(chart);
  const valueRange = yMax - yMin || 1;

  for (let sIdx = series.length - 1; sIdx >= 0; sIdx--) {
    const s = series[sIdx];
    const color = getSeriesColor(s, sIdx);
    const numPoints = s.values.length;
    if (numPoints === 0) {
      continue;
    }

    // Build the area path
    const ops: PathOp[] = [];
    const startX = plotX + (0.5 / numPoints) * plotW;
    const baselineY = plotY; // baseline at bottom of plot

    ops.push({ op: "move", x: startX, y: baselineY });

    for (let i = 0; i < numPoints; i++) {
      const x = plotX + ((i + 0.5) / numPoints) * plotW;
      const y = plotY + ((s.values[i] - yMin) / valueRange) * plotH;
      ops.push({ op: "line", x, y });
    }

    // Close back to baseline
    const endX = plotX + ((numPoints - 0.5) / numPoints) * plotW;
    ops.push({ op: "line", x: endX, y: baselineY });
    ops.push({ op: "close" });

    page.drawPath(ops, {
      fill: { ...color, a: 0.5 },
      stroke: color,
      lineWidth: 1,
      closePath: true
    });
  }
}

function renderPieChart(
  page: PdfPageBuilder,
  state: RenderState,
  chart: Chart,
  plotX: number,
  plotY: number,
  plotW: number,
  plotH: number
): void {
  const series = chart.series;
  if (series.length === 0) {
    return;
  }

  const s = series[0]; // Pie uses first series only
  const values = s.values;
  const total = values.reduce((sum, v) => sum + Math.abs(v), 0);
  if (total === 0) {
    return;
  }

  const cx = plotX + plotW / 2;
  const cy = plotY + plotH / 2;
  const radius = Math.min(plotW, plotH) * 0.4;
  const isDoughnut = chart.type === "doughnut";
  const innerRadius = isDoughnut ? radius * 0.5 : 0;

  let startAngle = -Math.PI / 2; // Start from top

  for (let i = 0; i < values.length; i++) {
    const sliceAngle = (Math.abs(values[i]) / total) * Math.PI * 2;
    const endAngle = startAngle + sliceAngle;
    const color = s.pointColors?.[i] ? parseHexColor(s.pointColors[i]) : getSeriesColor(s, i);

    // Draw pie slice using path with bezier approximation
    const ops = buildArcPath(cx, cy, radius, innerRadius, startAngle, endAngle);
    page.drawPath(ops, {
      fill: color,
      stroke: { r: 1, g: 1, b: 1 },
      lineWidth: 1,
      closePath: true
    });

    // Draw label
    if (s.categories[i]) {
      const midAngle = startAngle + sliceAngle / 2;
      const labelRadius = radius * 1.15;
      const labelX = cx + Math.cos(midAngle) * labelRadius;
      const labelY = cy + Math.sin(midAngle) * labelRadius;
      page.drawText(truncateLabel(s.categories[i], 8), {
        x: labelX,
        y: labelY - 3,
        fontSize: 6,
        fontFamily: state.opts.defaultFont,
        anchor: "middle",
        color: { r: 0.3, g: 0.3, b: 0.3 }
      });
    }

    startAngle = endAngle;
  }
}

function renderScatterChart(
  page: PdfPageBuilder,
  state: RenderState,
  chart: Chart,
  plotX: number,
  plotY: number,
  plotW: number,
  plotH: number
): void {
  const series = chart.series;
  if (series.length === 0) {
    return;
  }

  // For scatter, categories are X values (parsed as numbers), values are Y
  let xMin = Infinity;
  let xMax = -Infinity;
  let yMinV = Infinity;
  let yMaxV = -Infinity;

  for (const s of series) {
    for (let i = 0; i < s.values.length; i++) {
      const xVal = parseFloat(s.categories[i]) || i;
      const yVal = s.values[i];
      if (xVal < xMin) {
        xMin = xVal;
      }
      if (xVal > xMax) {
        xMax = xVal;
      }
      if (yVal < yMinV) {
        yMinV = yVal;
      }
      if (yVal > yMaxV) {
        yMaxV = yVal;
      }
    }
  }

  if (!Number.isFinite(xMin)) {
    xMin = 0;
  }
  if (!Number.isFinite(xMax)) {
    xMax = 1;
  }
  if (!Number.isFinite(yMinV)) {
    yMinV = 0;
  }
  if (!Number.isFinite(yMaxV)) {
    yMaxV = 1;
  }

  // Add 10% padding
  const xRange = xMax - xMin || 1;
  const yRange = yMaxV - yMinV || 1;
  xMin -= xRange * 0.05;
  xMax += xRange * 0.05;
  yMinV -= yRange * 0.05;
  yMaxV += yRange * 0.05;
  const xRangeFull = xMax - xMin;
  const yRangeFull = yMaxV - yMinV;

  // Draw grid and axes
  const gridColor = { r: 0.9, g: 0.9, b: 0.9 };
  const axisColor = { r: 0.3, g: 0.3, b: 0.3 };
  const gridLines = 5;

  for (let i = 0; i <= gridLines; i++) {
    const ratio = i / gridLines;
    // Horizontal
    const hy = plotY + ratio * plotH;
    page.drawLine({
      x1: plotX,
      y1: hy,
      x2: plotX + plotW,
      y2: hy,
      color: gridColor,
      lineWidth: 0.5
    });
    const yLabel = formatAxisValue(yMinV + ratio * yRangeFull);
    page.drawText(yLabel, {
      x: plotX - 4,
      y: hy - 3,
      fontSize: 7,
      fontFamily: state.opts.defaultFont,
      anchor: "end",
      color: axisColor
    });
    // Vertical
    const vx = plotX + ratio * plotW;
    page.drawLine({
      x1: vx,
      y1: plotY,
      x2: vx,
      y2: plotY + plotH,
      color: gridColor,
      lineWidth: 0.5
    });
    const xLabel = formatAxisValue(xMin + ratio * xRangeFull);
    page.drawText(xLabel, {
      x: vx,
      y: plotY - 10,
      fontSize: 7,
      fontFamily: state.opts.defaultFont,
      anchor: "middle",
      color: axisColor
    });
  }

  // Axes
  page.drawLine({
    x1: plotX,
    y1: plotY,
    x2: plotX + plotW,
    y2: plotY,
    color: axisColor,
    lineWidth: 0.75
  });
  page.drawLine({
    x1: plotX,
    y1: plotY,
    x2: plotX,
    y2: plotY + plotH,
    color: axisColor,
    lineWidth: 0.75
  });

  // Draw data points
  for (let sIdx = 0; sIdx < series.length; sIdx++) {
    const s = series[sIdx];
    const color = getSeriesColor(s, sIdx);

    for (let i = 0; i < s.values.length; i++) {
      const xVal = parseFloat(s.categories[i]) || i;
      const yVal = s.values[i];
      const px = plotX + ((xVal - xMin) / xRangeFull) * plotW;
      const py = plotY + ((yVal - yMinV) / yRangeFull) * plotH;
      page.drawCircle({
        cx: px,
        cy: py,
        r: 3,
        fill: color,
        stroke: { r: 1, g: 1, b: 1 },
        lineWidth: 0.5
      });
    }

    // Connect with line for scatterSmooth
    if (chart.type === "scatterSmooth" && s.values.length > 1) {
      const points: Array<{ x: number; y: number }> = [];
      for (let i = 0; i < s.values.length; i++) {
        const xVal = parseFloat(s.categories[i]) || i;
        const yVal = s.values[i];
        points.push({
          x: plotX + ((xVal - xMin) / xRangeFull) * plotW,
          y: plotY + ((yVal - yMinV) / yRangeFull) * plotH
        });
      }
      for (let i = 0; i < points.length - 1; i++) {
        page.drawLine({
          x1: points[i].x,
          y1: points[i].y,
          x2: points[i + 1].x,
          y2: points[i + 1].y,
          color,
          lineWidth: 1
        });
      }
    }
  }
}

function renderChartLegend(
  page: PdfPageBuilder,
  state: RenderState,
  chart: Chart,
  chartX: number,
  chartBottom: number,
  widthPt: number,
  heightPt: number
): void {
  const series = chart.series;
  const legendPos = chart.legend ?? "b";
  const fontSize = 7;
  const swatchSize = 8;
  const gap = 4;

  if (legendPos === "b") {
    // Bottom legend
    let x = chartX + 20;
    const y = chartBottom + 6;
    for (let i = 0; i < series.length; i++) {
      const color = getSeriesColor(series[i], i);
      page.drawRect({ x, y, width: swatchSize, height: swatchSize, fill: color });
      x += swatchSize + gap;
      page.drawText(truncateLabel(series[i].name, 15), {
        x,
        y: y + 1,
        fontSize,
        fontFamily: state.opts.defaultFont,
        color: { r: 0.2, g: 0.2, b: 0.2 }
      });
      x += series[i].name.length * fontSize * 0.5 + gap * 3;
    }
  } else if (legendPos === "r" || legendPos === "tr") {
    // Right legend
    const x = chartX + widthPt - 55;
    let y = chartBottom + heightPt - 30;
    for (let i = 0; i < series.length; i++) {
      const color = getSeriesColor(series[i], i);
      page.drawRect({ x, y, width: swatchSize, height: swatchSize, fill: color });
      page.drawText(truncateLabel(series[i].name, 10), {
        x: x + swatchSize + gap,
        y: y + 1,
        fontSize,
        fontFamily: state.opts.defaultFont,
        color: { r: 0.2, g: 0.2, b: 0.2 }
      });
      y -= 14;
    }
  } else if (legendPos === "t") {
    // Top legend
    let x = chartX + 20;
    const y = chartBottom + heightPt - 14;
    for (let i = 0; i < series.length; i++) {
      const color = getSeriesColor(series[i], i);
      page.drawRect({ x, y, width: swatchSize, height: swatchSize, fill: color });
      x += swatchSize + gap;
      page.drawText(truncateLabel(series[i].name, 15), {
        x,
        y: y + 1,
        fontSize,
        fontFamily: state.opts.defaultFont,
        color: { r: 0.2, g: 0.2, b: 0.2 }
      });
      x += series[i].name.length * fontSize * 0.5 + gap * 3;
    }
  } else if (legendPos === "l") {
    // Left legend
    const x = chartX + 5;
    let y = chartBottom + heightPt - 30;
    for (let i = 0; i < series.length; i++) {
      const color = getSeriesColor(series[i], i);
      page.drawRect({ x, y, width: swatchSize, height: swatchSize, fill: color });
      page.drawText(truncateLabel(series[i].name, 10), {
        x: x + swatchSize + gap,
        y: y + 1,
        fontSize,
        fontFamily: state.opts.defaultFont,
        color: { r: 0.2, g: 0.2, b: 0.2 }
      });
      y -= 14;
    }
  }
}

/** Build an arc path for pie/doughnut slices. */
function buildArcPath(
  cx: number,
  cy: number,
  outerR: number,
  innerR: number,
  startAngle: number,
  endAngle: number
): PathOp[] {
  const ops: PathOp[] = [];

  // Approximate arc with cubic bezier segments (max 90 degrees per segment)
  const outerArc = approximateArc(cx, cy, outerR, startAngle, endAngle);
  const innerArc = innerR > 0 ? approximateArc(cx, cy, innerR, endAngle, startAngle) : null;

  // Start at outer arc beginning
  ops.push({
    op: "move",
    x: cx + Math.cos(startAngle) * outerR,
    y: cy + Math.sin(startAngle) * outerR
  });

  // Outer arc
  for (const seg of outerArc) {
    ops.push({
      op: "curve",
      x1: seg.x1,
      y1: seg.y1,
      x2: seg.x2,
      y2: seg.y2,
      x3: seg.x3,
      y3: seg.y3
    });
  }

  if (innerR > 0) {
    // Line to inner arc start
    ops.push({
      op: "line",
      x: cx + Math.cos(endAngle) * innerR,
      y: cy + Math.sin(endAngle) * innerR
    });
    // Inner arc (reverse direction)
    for (const seg of innerArc!) {
      ops.push({
        op: "curve",
        x1: seg.x1,
        y1: seg.y1,
        x2: seg.x2,
        y2: seg.y2,
        x3: seg.x3,
        y3: seg.y3
      });
    }
  } else {
    // Line to center
    ops.push({ op: "line", x: cx, y: cy });
  }

  ops.push({ op: "close" });
  return ops;
}

/** Approximate a circular arc with cubic bezier curves. */
function approximateArc(
  cx: number,
  cy: number,
  r: number,
  startAngle: number,
  endAngle: number
): Array<{ x1: number; y1: number; x2: number; y2: number; x3: number; y3: number }> {
  const segments: Array<{
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    x3: number;
    y3: number;
  }> = [];
  let remaining = endAngle - startAngle;
  let current = startAngle;

  // Handle negative arcs (for inner ring going reverse)
  const sign = remaining < 0 ? -1 : 1;
  remaining = Math.abs(remaining);

  while (remaining > 1e-6) {
    const segAngle = Math.min(remaining, Math.PI / 2) * sign;
    const halfAngle = segAngle / 2;
    const kappa = (4 / 3) * Math.tan(halfAngle / 2);

    const cos0 = Math.cos(current);
    const sin0 = Math.sin(current);
    const cos1 = Math.cos(current + segAngle);
    const sin1 = Math.sin(current + segAngle);

    segments.push({
      x1: cx + r * (cos0 - kappa * sin0),
      y1: cy + r * (sin0 + kappa * cos0),
      x2: cx + r * (cos1 + kappa * sin1),
      y2: cy + r * (sin1 - kappa * cos1),
      x3: cx + r * cos1,
      y3: cy + r * sin1
    });

    current += segAngle;
    remaining -= Math.abs(segAngle);
  }

  return segments;
}

/** Get the value range for a chart (min/max across all series). */
function getValueRange(chart: Chart): { min: number; max: number } {
  let min = chart.valueAxis?.min ?? Infinity;
  let max = chart.valueAxis?.max ?? -Infinity;

  if (!Number.isFinite(min) || !Number.isFinite(max)) {
    for (const s of chart.series) {
      for (const v of s.values) {
        if (v < min) {
          min = v;
        }
        if (v > max) {
          max = v;
        }
      }
    }
  }

  if (!Number.isFinite(min)) {
    min = 0;
  }
  if (!Number.isFinite(max)) {
    max = 1;
  }

  // Ensure min <= 0 for bar charts if data includes 0
  if (min > 0) {
    min = 0;
  }
  // Add a little headroom
  if (max === min) {
    max = min + 1;
  }
  const range = max - min;
  max += range * 0.05;

  return { min, max };
}

/** Get color for a series, falling back to default palette. */
function getSeriesColor(series: ChartSeries, index: number): { r: number; g: number; b: number } {
  if (series.color) {
    return parseHexColor(series.color);
  }
  return parseHexColor(DEFAULT_CHART_COLORS[index % DEFAULT_CHART_COLORS.length]);
}

/** Format an axis value for display. */
function formatAxisValue(value: number): string {
  if (Math.abs(value) >= 1000000) {
    return (value / 1000000).toFixed(1) + "M";
  }
  if (Math.abs(value) >= 1000) {
    return (value / 1000).toFixed(1) + "K";
  }
  if (Number.isInteger(value)) {
    return String(value);
  }
  return value.toFixed(1);
}

/** Truncate a label to fit in limited space. */
function truncateLabel(label: string, maxLen: number): string {
  if (label.length <= maxLen) {
    return label;
  }
  return label.slice(0, maxLen - 1) + "\u2026";
}

// =============================================================================
// ChartEx Placeholder Rendering
// =============================================================================

function renderChartExPlaceholder(state: RenderState, content: ChartExContent): void {
  const defaultWidthEmu = 5486400;
  const defaultHeightEmu = 3657600;

  let widthPt = emuToPt(content.width ?? defaultWidthEmu);
  let heightPt = emuToPt(content.height ?? defaultHeightEmu);

  // Scale to fit
  if (widthPt > state.availableWidth) {
    const scale = state.availableWidth / widthPt;
    widthPt = state.availableWidth;
    heightPt *= scale;
  }

  // Use a smaller fixed height for placeholder
  const boxHeight = Math.min(heightPt, 60);
  ensureSpace(state, boxHeight + 6);

  const page = state.currentPage!;
  const x = state.opts.marginLeft;
  const y = state.cursorY - boxHeight;

  // Draw bordered box
  page.drawRect({
    x,
    y,
    width: widthPt,
    height: boxHeight,
    fill: { r: 0.97, g: 0.97, b: 0.97 },
    stroke: { r: 0.7, g: 0.7, b: 0.7 },
    lineWidth: 1
  });

  // Draw label
  const label = `ChartEx: ${content.altText ?? content.name ?? "advanced chart"}`;
  page.drawText(label, {
    x: x + widthPt / 2,
    y: y + boxHeight / 2 - 5,
    fontSize: 10,
    fontFamily: state.opts.defaultFont,
    italic: true,
    color: { r: 0.4, g: 0.4, b: 0.4 },
    anchor: "middle"
  });

  state.cursorY = y - 6;
}

// =============================================================================
// Drawing Shape Rendering
// =============================================================================

function renderDrawingShape(state: RenderState, shape: DrawingShape): void {
  const widthPt = emuToPt(shape.width);
  const heightPt = emuToPt(shape.height);

  // If shape is too large, scale to fit
  const maxW = state.availableWidth;
  const scale = widthPt > maxW ? maxW / widthPt : 1;
  const w = widthPt * scale;
  const h = heightPt * scale;

  ensureSpace(state, h + 6);

  const page = state.currentPage!;
  const x = state.opts.marginLeft;
  const y = state.cursorY - h;

  // Resolve colors
  const fillColor = shape.noFill
    ? undefined
    : shape.fillColor
      ? parseHexColor(shape.fillColor)
      : undefined;
  const outlineColor = shape.noOutline
    ? undefined
    : shape.outlineColor
      ? parseHexColor(shape.outlineColor)
      : { r: 0, g: 0, b: 0 };
  const outlineWidth = shape.outlineWidth ? emuToPt(shape.outlineWidth) * scale : 1;

  switch (shape.shapeType) {
    case "rect":
    case "flowChartProcess":
    case "flowChartTerminator":
      page.drawRect({
        x,
        y,
        width: w,
        height: h,
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth
      });
      break;

    case "roundRect":
    case "snip1Rect":
    case "snip2SameRect": {
      const radius = Math.min(w, h) * 0.1;
      page.drawRect({
        x,
        y,
        width: w,
        height: h,
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth,
        borderRadius: radius
      });
      break;
    }

    case "ellipse":
    case "circle":
      page.drawEllipse({
        cx: x + w / 2,
        cy: y + h / 2,
        rx: w / 2,
        ry: h / 2,
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth
      });
      break;

    case "line":
    case "straightConnector1":
    case "bentConnector3":
      page.drawLine({
        x1: x,
        y1: y + h / 2,
        x2: x + w,
        y2: y + h / 2,
        color: outlineColor ?? { r: 0, g: 0, b: 0 },
        lineWidth: outlineWidth
      });
      break;

    case "triangle":
    case "rtTriangle": {
      const ops: PathOp[] = [
        { op: "move", x: x + w / 2, y: y + h }, // top center (or top for right triangle)
        { op: "line", x: x + w, y }, // bottom right
        { op: "line", x, y }, // bottom left
        { op: "close" }
      ];
      if (shape.shapeType === "rtTriangle") {
        ops[0] = { op: "move", x, y: y + h }; // top left
        ops[1] = { op: "line", x, y }; // bottom left
        ops[2] = { op: "line", x: x + w, y }; // bottom right
      }
      page.drawPath(ops, {
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth,
        closePath: true
      });
      break;
    }

    case "diamond": {
      const ops: PathOp[] = [
        { op: "move", x: x + w / 2, y: y + h }, // top
        { op: "line", x: x + w, y: y + h / 2 }, // right
        { op: "line", x: x + w / 2, y }, // bottom
        { op: "line", x, y: y + h / 2 }, // left
        { op: "close" }
      ];
      page.drawPath(ops, {
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth,
        closePath: true
      });
      break;
    }

    case "pentagon": {
      const ops = buildRegularPolygonPath(x, y, w, h, 5);
      page.drawPath(ops, {
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth,
        closePath: true
      });
      break;
    }

    case "hexagon": {
      const ops = buildRegularPolygonPath(x, y, w, h, 6);
      page.drawPath(ops, {
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth,
        closePath: true
      });
      break;
    }

    case "star5":
    case "star4": {
      const ops = buildStarPath(x, y, w, h, shape.shapeType === "star5" ? 5 : 4);
      page.drawPath(ops, {
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth,
        closePath: true
      });
      break;
    }

    case "arrow":
    case "rightArrow":
    case "leftArrow": {
      // Simple arrow shape
      const arrowOps: PathOp[] = [
        { op: "move", x, y: y + h * 0.3 },
        { op: "line", x: x + w * 0.6, y: y + h * 0.3 },
        { op: "line", x: x + w * 0.6, y },
        { op: "line", x: x + w, y: y + h / 2 },
        { op: "line", x: x + w * 0.6, y: y + h },
        { op: "line", x: x + w * 0.6, y: y + h * 0.7 },
        { op: "line", x, y: y + h * 0.7 },
        { op: "close" }
      ];
      page.drawPath(arrowOps, {
        fill: fillColor,
        stroke: outlineColor,
        lineWidth: outlineWidth,
        closePath: true
      });
      break;
    }

    default:
      // Fallback: draw a rectangle with shape type annotation
      page.drawRect({
        x,
        y,
        width: w,
        height: h,
        fill: fillColor ?? { r: 0.95, g: 0.95, b: 0.95 },
        stroke: outlineColor ?? { r: 0.7, g: 0.7, b: 0.7 },
        lineWidth: outlineWidth
      });
      page.drawText(shape.shapeType, {
        x: x + w / 2,
        y: y + h / 2 - 4,
        fontSize: 7,
        fontFamily: state.opts.defaultFont,
        italic: true,
        color: { r: 0.5, g: 0.5, b: 0.5 },
        anchor: "middle"
      });
      break;
  }

  // Draw text content inside the shape
  if (shape.textContent && shape.textContent.length > 0) {
    renderShapeTextContent(state, shape.textContent, x, y, w, h);
  }

  state.cursorY = y - 6;
}

/** Render paragraph content inside a shape, centered. */
function renderShapeTextContent(
  state: RenderState,
  paragraphs: readonly Paragraph[],
  x: number,
  y: number,
  w: number,
  h: number
): void {
  const page = state.currentPage!;
  const padding = 6;
  const maxWidth = w - padding * 2;
  if (maxWidth <= 0) {
    return;
  }

  // Collect all text from paragraphs
  const lines: Array<{
    text: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
    color?: { r: number; g: number; b: number };
  }> = [];

  for (const para of paragraphs) {
    const fragments = collectFragments(state, para);
    let lineText = "";
    let lineFontSize = state.opts.defaultFontSize;
    let lineBold = false;
    let lineItalic = false;
    let lineColor: { r: number; g: number; b: number } | undefined;

    for (const f of fragments) {
      if ("type" in f && f.type === "image") {
        continue;
      }
      const inf = f as InlineFragment;
      lineText += inf.text;
      lineFontSize = Math.max(lineFontSize, inf.fontSize);
      if (inf.bold) {
        lineBold = true;
      }
      if (inf.italic) {
        lineItalic = true;
      }
      if (inf.color) {
        lineColor = inf.color;
      }
    }

    if (lineText.trim()) {
      lines.push({
        text: lineText.trim(),
        fontSize: lineFontSize,
        bold: lineBold,
        italic: lineItalic,
        color: lineColor
      });
    }
  }

  if (lines.length === 0) {
    return;
  }

  // Calculate total text height
  const totalHeight = lines.reduce((sum, l) => sum + l.fontSize * 1.3, 0);
  let textY = y + h / 2 + totalHeight / 2 - lines[0].fontSize;

  for (const line of lines) {
    if (textY < y + padding) {
      break;
    }
    page.drawText(line.text, {
      x: x + w / 2,
      y: textY,
      fontSize: line.fontSize,
      fontFamily: state.opts.defaultFont,
      bold: line.bold,
      italic: line.italic,
      color: line.color,
      anchor: "middle",
      maxWidth
    });
    textY -= line.fontSize * 1.3;
  }
}

/** Build path for a regular polygon inscribed in a bounding box. */
function buildRegularPolygonPath(
  x: number,
  y: number,
  w: number,
  h: number,
  sides: number
): PathOp[] {
  const ops: PathOp[] = [];
  const cx = x + w / 2;
  const cy = y + h / 2;
  const rx = w / 2;
  const ry = h / 2;

  for (let i = 0; i < sides; i++) {
    const angle = -Math.PI / 2 + (i * 2 * Math.PI) / sides;
    const px = cx + Math.cos(angle) * rx;
    const py = cy + Math.sin(angle) * ry;
    ops.push(i === 0 ? { op: "move", x: px, y: py } : { op: "line", x: px, y: py });
  }
  ops.push({ op: "close" });
  return ops;
}

/** Build path for a star shape. */
function buildStarPath(x: number, y: number, w: number, h: number, points: number): PathOp[] {
  const ops: PathOp[] = [];
  const cx = x + w / 2;
  const cy = y + h / 2;
  const outerRx = w / 2;
  const outerRy = h / 2;
  const innerRx = outerRx * 0.4;
  const innerRy = outerRy * 0.4;
  const totalPoints = points * 2;

  for (let i = 0; i < totalPoints; i++) {
    const angle = -Math.PI / 2 + (i * Math.PI) / points;
    const isOuter = i % 2 === 0;
    const rx = isOuter ? outerRx : innerRx;
    const ry = isOuter ? outerRy : innerRy;
    const px = cx + Math.cos(angle) * rx;
    const py = cy + Math.sin(angle) * ry;
    ops.push(i === 0 ? { op: "move", x: px, y: py } : { op: "line", x: px, y: py });
  }
  ops.push({ op: "close" });
  return ops;
}
