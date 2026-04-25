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
 * Note: This is a simple flow-based renderer. For high-fidelity output,
 * consider using Microsoft Word or LibreOffice.
 */

import type {
  DocxDocument,
  Paragraph,
  Run,
  Table,
  TableCell,
  BodyContent,
  ParagraphChild,
  ImageDef
} from "@word/types";

import { PdfDocumentBuilder, type PdfPageBuilder } from "./builder/document-builder";

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
  /** Default font family (default: "Times-Roman"). */
  readonly defaultFont?: string;
  /** Default font size in points (default: 11). */
  readonly defaultFontSize?: number;
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
  const opts = {
    pageWidth: options?.pageWidth ?? 612,
    pageHeight: options?.pageHeight ?? 792,
    marginTop: options?.marginTop ?? 72,
    marginBottom: options?.marginBottom ?? 72,
    marginLeft: options?.marginLeft ?? 72,
    marginRight: options?.marginRight ?? 72,
    defaultFont: options?.defaultFont ?? "Times-Roman",
    defaultFontSize: options?.defaultFontSize ?? 11
  };

  // Use section properties if available
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
    currentPageIndex: 0
  };

  newPage(state);

  for (const item of doc.body) {
    renderBodyContent(state, item);
  }

  return builder.build();
}

interface RenderState {
  builder: PdfDocumentBuilder;
  opts: {
    pageWidth: number;
    pageHeight: number;
    marginTop: number;
    marginBottom: number;
    marginLeft: number;
    marginRight: number;
    defaultFont: string;
    defaultFontSize: number;
  };
  doc: DocxDocument;
  rIdToImage: Map<string, ImageDef>;
  currentPage: PdfPageBuilder | null;
  cursorY: number;
  currentPageIndex: number;
}

function newPage(state: RenderState): void {
  state.currentPage = state.builder.addPage({
    width: state.opts.pageWidth,
    height: state.opts.pageHeight
  });
  state.cursorY = state.opts.pageHeight - state.opts.marginTop;
  state.currentPageIndex++;
}

function ensureSpace(state: RenderState, needed: number): void {
  if (state.cursorY - needed < state.opts.marginBottom) {
    newPage(state);
  }
}

function renderBodyContent(state: RenderState, item: BodyContent): void {
  if (!("type" in item)) {
    return;
  }
  switch (item.type) {
    case "paragraph":
      renderParagraphPdf(state, item);
      break;
    case "table":
      renderTablePdf(state, item);
      break;
    case "math":
      renderTextAtCursor(state, "[equation]", { italic: true });
      break;
    case "chart":
      renderTextAtCursor(state, `[${item.chart.type} chart]`, { italic: true });
      break;
    default:
      break;
  }
}

function renderParagraphPdf(state: RenderState, para: Paragraph): void {
  const props = para.properties;
  const spacing = props?.spacing;

  if (spacing?.before) {
    state.cursorY -= twipsToPt(spacing.before);
  }

  const text = collectText(para);
  if (!text.trim()) {
    state.cursorY -= state.opts.defaultFontSize * 1.2;
    return;
  }

  // Determine font size and weight based on style
  let fontSize = state.opts.defaultFontSize;
  let bold = false;
  const styleId = props?.style?.toLowerCase() ?? "";

  if (styleId === "heading1" || styleId === "title") {
    fontSize = 24;
    bold = true;
  } else if (styleId === "heading2") {
    fontSize = 18;
    bold = true;
  } else if (styleId === "heading3") {
    fontSize = 14;
    bold = true;
  } else if (styleId === "heading4") {
    fontSize = 12;
    bold = true;
  }

  let leftIndent = 0;
  if (props?.indent?.left !== undefined) {
    leftIndent = twipsToPt(props.indent.left);
  }

  const x = state.opts.marginLeft + leftIndent;
  const availableWidth =
    state.opts.pageWidth - state.opts.marginLeft - state.opts.marginRight - leftIndent;
  const lineHeight = fontSize * 1.2;

  const estimatedLines = Math.max(1, Math.ceil((text.length * fontSize * 0.5) / availableWidth));
  ensureSpace(state, estimatedLines * lineHeight);

  // Simple approach: render entire paragraph as one text block
  state.currentPage!.drawText(text, {
    x,
    y: state.cursorY - fontSize,
    fontSize,
    bold,
    maxWidth: availableWidth,
    lineHeight: 1.2
  });

  state.cursorY -= estimatedLines * lineHeight;

  if (spacing?.after) {
    state.cursorY -= twipsToPt(spacing.after);
  } else {
    state.cursorY -= 6;
  }
}

function collectText(para: Paragraph): string {
  let text = "";
  for (const child of para.children) {
    const run = childToRun(child);
    if (run) {
      text += runToText(run);
    }
  }
  return text;
}

function childToRun(child: ParagraphChild): Run | undefined {
  if ("content" in child && !("type" in child)) {
    return child as Run;
  }
  if ("type" in child) {
    const typed = child as { type: string; children?: ParagraphChild[]; run?: Run };
    if (typed.type === "hyperlink" && typed.children) {
      for (const c of typed.children) {
        if ("content" in c && !("type" in c)) {
          return c as Run;
        }
      }
    } else if (
      typed.type === "insertedRun" ||
      typed.type === "deletedRun" ||
      typed.type === "movedFromRun" ||
      typed.type === "movedToRun"
    ) {
      return typed.run;
    }
  }
  return undefined;
}

function runToText(run: Run): string {
  let text = "";
  for (const c of run.content) {
    switch (c.type) {
      case "text":
        text += c.text;
        break;
      case "tab":
        text += "\t";
        break;
      case "break":
      case "carriageReturn":
        text += "\n";
        break;
      case "noBreakHyphen":
        text += "\u2011";
        break;
      case "symbol":
        try {
          text += String.fromCodePoint(parseInt(c.char, 16));
        } catch {
          // Invalid symbol code - skip
        }
        break;
      case "field":
        if (c.cachedValue) {
          text += c.cachedValue;
        }
        break;
    }
  }
  return text;
}

function renderTablePdf(state: RenderState, table: Table): void {
  const pageWidth = state.opts.pageWidth - state.opts.marginLeft - state.opts.marginRight;
  const numCols = Math.max(...table.rows.map(r => r.cells.length));
  const defaultColWidth = pageWidth / numCols;
  const rowHeight = 20;

  for (const row of table.rows) {
    ensureSpace(state, rowHeight);

    let x = state.opts.marginLeft;
    const rowTop = state.cursorY;
    const rowBottom = state.cursorY - rowHeight;

    for (const cell of row.cells) {
      const span = cell.properties?.gridSpan ?? 1;
      const cellWidth = defaultColWidth * span;

      state.currentPage!.drawRect({
        x,
        y: rowBottom,
        width: cellWidth,
        height: rowHeight,
        stroke: { r: 0, g: 0, b: 0 },
        lineWidth: 0.5
      });

      const cellText = getCellText(cell);
      if (cellText) {
        state.currentPage!.drawText(cellText, {
          x: x + 4,
          y: rowTop - 14,
          fontSize: 10,
          maxWidth: cellWidth - 8
        });
      }

      x += cellWidth;
    }

    state.cursorY = rowBottom;
  }

  state.cursorY -= 6;
}

function getCellText(cell: TableCell): string {
  let text = "";
  for (const item of cell.content) {
    if ("type" in item && item.type === "paragraph") {
      text += collectText(item as Paragraph) + " ";
    }
  }
  return text.trim();
}

function twipsToPt(twips: number): number {
  return twips / 20;
}

function renderTextAtCursor(
  state: RenderState,
  text: string,
  opts?: { bold?: boolean; italic?: boolean; fontSize?: number }
): void {
  const fontSize = opts?.fontSize ?? state.opts.defaultFontSize;
  ensureSpace(state, fontSize * 1.2);
  state.currentPage!.drawText(text, {
    x: state.opts.marginLeft,
    y: state.cursorY - fontSize,
    fontSize,
    bold: opts?.bold,
    italic: opts?.italic
  });
  state.cursorY -= fontSize * 1.2;
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
