/**
 * Full Layout Engine — produces a complete LayoutDocument with positioned elements.
 *
 * Uses the pagination result from layoutDocument() for page assignments,
 * then computes precise positions (x, y, width, height) for paragraphs,
 * lines, runs, and tables on each page.
 *
 * This is the bridge between the page-number-only LayoutResult and the
 * fully positioned LayoutDocument that renderers (SVG, PDF, Canvas) can consume.
 */

import { measureTextWidth, mapToStandardFont } from "@utils/font-metrics";

import { isHyperlink, isRun } from "../core/text-utils";
import type { DocxDocument, Paragraph, ParagraphProperties, Run, Table } from "../types";
import { layoutDocument } from "./layout";
import type { LayoutOptions, LayoutResult } from "./layout";
import type {
  LayoutDocument,
  LayoutPage,
  LayoutParagraph,
  LayoutTable,
  LayoutTableCell,
  LineBox,
  PageContent,
  PageGeometry,
  PositionedRun
} from "./layout-model";

// =============================================================================
// Public API
// =============================================================================

/** Options for the full layout engine. */
export interface FullLayoutOptions extends LayoutOptions {
  /** Font map for font-family resolution (name → actual font). */
  readonly fonts?: ReadonlyMap<string, string>;
}

/**
 * Perform full document layout, producing a LayoutDocument with precise positions.
 *
 * @param doc - The parsed DOCX document.
 * @param options - Layout and font options.
 * @returns A fully positioned LayoutDocument.
 */
export function layoutDocumentFull(doc: DocxDocument, options?: FullLayoutOptions): LayoutDocument {
  // First pass: get page assignments via the existing lightweight layout
  const layoutResult = layoutDocument(doc, options);

  // Second pass: compute precise positions for each page
  const pages: LayoutPage[] = [];
  const totalPages = layoutResult.pageCount;

  for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
    const page = buildPage(doc, pageNum, layoutResult, options);
    pages.push(page);
  }

  return {
    pages,
    totalPages,
    bookmarkPages: layoutResult.bookmarkPages,
    sectionBreaks: computeSectionBreaks(layoutResult)
  };
}

// =============================================================================
// Internal: Page Building
// =============================================================================

const DEFAULT_PAGE_WIDTH_TWIPS = 12240;
const DEFAULT_PAGE_HEIGHT_TWIPS = 15840;
const DEFAULT_MARGIN_TWIPS = 1440;
const DEFAULT_FONT_SIZE_PT = 12;

function twipsToPt(twips: number): number {
  return twips / 20;
}

function buildPage(
  doc: DocxDocument,
  pageNumber: number,
  layout: LayoutResult,
  options?: FullLayoutOptions
): LayoutPage {
  const sectionProps = doc.sectionProperties;
  const geometry = computePageGeometry(sectionProps);
  const content: PageContent[] = [];

  let cursorY = 0; // relative to content area top

  for (let i = 0; i < doc.body.length; i++) {
    if (layout.contentPages[i] !== pageNumber) {
      continue;
    }

    const item = doc.body[i];
    switch (item.type) {
      case "paragraph": {
        const laid = layoutParagraph(item, cursorY, geometry.contentWidth, options);
        content.push({ ...laid, sourceIndex: i });
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      case "table": {
        const laid = layoutTable(item, cursorY, geometry.contentWidth, i, options);
        content.push(laid);
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      default:
        // Skip unsupported types
        cursorY += DEFAULT_FONT_SIZE_PT * 1.2;
        break;
    }
  }

  return { pageNumber, geometry, content };
}

function computePageGeometry(sectionProps: DocxDocument["sectionProperties"]): PageGeometry {
  const widthTwips = sectionProps?.pageSize?.width ?? DEFAULT_PAGE_WIDTH_TWIPS;
  const heightTwips = sectionProps?.pageSize?.height ?? DEFAULT_PAGE_HEIGHT_TWIPS;
  const marginTop = twipsToPt(sectionProps?.margins?.top ?? DEFAULT_MARGIN_TWIPS);
  const marginBottom = twipsToPt(sectionProps?.margins?.bottom ?? DEFAULT_MARGIN_TWIPS);
  const marginLeft = twipsToPt(sectionProps?.margins?.left ?? DEFAULT_MARGIN_TWIPS);
  const marginRight = twipsToPt(sectionProps?.margins?.right ?? DEFAULT_MARGIN_TWIPS);
  const width = twipsToPt(widthTwips);
  const height = twipsToPt(heightTwips);

  return {
    width,
    height,
    marginTop,
    marginBottom,
    marginLeft,
    marginRight,
    contentWidth: width - marginLeft - marginRight,
    contentHeight: height - marginTop - marginBottom
  };
}

function computeSectionBreaks(layout: LayoutResult): number[] {
  const breaks: number[] = [0]; // First section starts at page 0
  let prevSection = 0;
  for (let i = 0; i < layout.contentPages.length; i++) {
    const section = layout.contentSections[i];
    if (section > prevSection) {
      breaks.push(layout.contentPages[i] - 1);
      prevSection = section;
    }
  }
  return breaks;
}

// =============================================================================
// Internal: Paragraph Layout
// =============================================================================

function layoutParagraph(
  para: Paragraph,
  startY: number,
  contentWidth: number,
  options?: FullLayoutOptions
): LayoutParagraph {
  const props = para.properties;
  const spacing = props?.spacing;
  const headingScale = getHeadingFontScale(getHeadingLevel(props));

  // Space before
  let spaceBefore = 0;
  if (spacing?.beforeAutoSpacing) {
    spaceBefore = 5;
  } else if (spacing?.before != null) {
    spaceBefore = twipsToPt(spacing.before);
  }

  const indent = props?.indent;
  const leftIndentPt = indent?.left ? twipsToPt(indent.left) : 0;
  const firstLineIndentPt = indent?.firstLine ? twipsToPt(indent.firstLine) : 0;
  const alignment = props?.alignment ?? "left";

  // Line height
  let lineHeightPt = DEFAULT_FONT_SIZE_PT * 1.2;
  if (spacing?.line) {
    const rule = spacing.lineRule ?? "auto";
    switch (rule) {
      case "exact":
        lineHeightPt = twipsToPt(spacing.line);
        break;
      case "atLeast":
        lineHeightPt = Math.max(twipsToPt(spacing.line), lineHeightPt);
        break;
      case "auto":
        lineHeightPt = DEFAULT_FONT_SIZE_PT * 1.2 * (spacing.line / 240);
        break;
    }
  }
  lineHeightPt *= headingScale;

  // Collect runs
  const segments = collectParagraphSegments(para);
  const availableWidth = contentWidth - leftIndentPt;
  const lines = wrapSegmentsToLines(segments, availableWidth, firstLineIndentPt, headingScale);

  // Build line boxes
  const lineBoxes: LineBox[] = [];
  let yOffset = spaceBefore;

  for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
    const lineSegments = lines[lineIdx];
    const runs: PositionedRun[] = [];
    let xPos = lineIdx === 0 ? firstLineIndentPt : 0;

    // Calculate line width for alignment
    let lineWidth = 0;
    for (const seg of lineSegments) {
      const fontSize = getRunFontSizePt(seg.properties) * headingScale;
      const fontName = mapToStandardFont(resolveRunFontName(seg.properties));
      lineWidth += measureTextWidth(seg.text, fontName, fontSize);
    }

    // Apply alignment
    if (alignment === "center") {
      xPos = (availableWidth - lineWidth) / 2;
    } else if (alignment === "right" || alignment === "end") {
      xPos = availableWidth - lineWidth;
    }

    xPos += leftIndentPt;

    for (const seg of lineSegments) {
      const fontSize = getRunFontSizePt(seg.properties) * headingScale;
      const fontName = resolveRunFontName(seg.properties);
      const measuredFont = mapToStandardFont(fontName);
      const segWidth = measureTextWidth(seg.text, measuredFont, fontSize);

      runs.push({
        text: seg.text,
        x: xPos,
        width: segWidth,
        font: fontName,
        fontSize,
        bold: seg.properties?.bold || undefined,
        italic: seg.properties?.italic || undefined,
        color: resolveColorHex(seg.properties?.color),
        underline: seg.properties?.underline !== undefined ? true : undefined,
        strikethrough: seg.properties?.strike || undefined
      });

      xPos += segWidth;
    }

    const mappedAlignment =
      alignment === "both"
        ? "justify"
        : alignment === "end"
          ? "right"
          : alignment === "start"
            ? "left"
            : (alignment as "left" | "center" | "right" | "justify");

    lineBoxes.push({
      y: yOffset,
      height: lineHeightPt,
      baseline: lineHeightPt * 0.8,
      runs,
      alignment: mappedAlignment
    });

    yOffset += lineHeightPt;
  }

  // If empty paragraph, still advance by one line
  if (lineBoxes.length === 0) {
    yOffset += lineHeightPt;
  }

  // Space after
  let spaceAfter = 0;
  if (spacing?.afterAutoSpacing) {
    spaceAfter = 5;
  } else if (spacing?.after != null) {
    spaceAfter = twipsToPt(spacing.after);
  }

  const totalHeight = yOffset + spaceAfter;

  return {
    type: "paragraph",
    rect: { x: 0, y: startY, width: contentWidth, height: totalHeight },
    lines: lineBoxes,
    sourceIndex: 0 // overwritten by caller
  };
}

// =============================================================================
// Internal: Table Layout
// =============================================================================

function layoutTable(
  table: Table,
  startY: number,
  contentWidth: number,
  sourceIndex: number,
  options?: FullLayoutOptions
): LayoutTable {
  const numCols = table.rows.length > 0 ? table.rows[0].cells.length : 0;
  const colWidth = numCols > 0 ? contentWidth / numCols : contentWidth;

  const cells: LayoutTableCell[] = [];
  let cursorY = 0;

  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    let maxRowHeight = DEFAULT_FONT_SIZE_PT * 1.5; // minimum row height

    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const cellX = ci * colWidth;
      const cellContent: (LayoutParagraph | LayoutTable)[] = [];
      let cellCursorY = 2; // cell padding top

      for (const block of cell.content) {
        if (block.type === "paragraph") {
          const laid = layoutParagraph(block, cellCursorY, colWidth - 4, options);
          cellContent.push({ ...laid, sourceIndex: -1 });
          cellCursorY = laid.rect.y + laid.rect.height;
        }
        // nested tables skipped for simplicity
      }

      const cellHeight = cellCursorY + 2; // cell padding bottom
      if (cellHeight > maxRowHeight) {
        maxRowHeight = cellHeight;
      }

      cells.push({
        rect: { x: cellX, y: startY + cursorY, width: colWidth, height: cellHeight },
        row: ri,
        col: ci,
        content: cellContent
      });
    }

    // Normalize cell heights to row max
    for (const c of cells) {
      if (c.row === ri) {
        (c as { rect: { height: number } }).rect.height = maxRowHeight;
      }
    }

    cursorY += maxRowHeight;
  }

  return {
    type: "table",
    rect: { x: 0, y: startY, width: contentWidth, height: cursorY },
    cells,
    sourceIndex
  };
}

// =============================================================================
// Internal: Text Helpers
// =============================================================================

interface TextSegment {
  readonly text: string;
  readonly properties: Run["properties"];
}

function collectParagraphSegments(para: Paragraph): TextSegment[] {
  const segments: TextSegment[] = [];
  for (const child of para.children) {
    if (isRun(child)) {
      const text = layoutRunText(child);
      if (text.length > 0) {
        segments.push({ text, properties: child.properties });
      }
    } else if (isHyperlink(child)) {
      for (const run of child.children) {
        const text = layoutRunText(run);
        if (text.length > 0) {
          segments.push({ text, properties: run.properties });
        }
      }
    }
  }
  return segments;
}

/**
 * Extract run text for layout measurement.
 *
 * Differs from `extractRunText` in `core/text-utils.ts`:
 *  - `tab` → 4 spaces (matches greedy-fit measurement; the caller does not
 *    yet honor real tab stops here)
 *  - non-text content (hyphens, field values) is ignored — the layout pass
 *    does not have access to field-engine output, and ignoring them gives
 *    a conservative (slightly under-) estimate of line width.
 */
function layoutRunText(run: Run): string {
  let text = "";
  for (const item of run.content) {
    if (item.type === "text") {
      text += item.text;
    } else if (item.type === "tab") {
      text += "    ";
    } else if (item.type === "break") {
      text += "\n";
    }
  }
  return text;
}

function wrapSegmentsToLines(
  segments: TextSegment[],
  availableWidth: number,
  firstLineIndent: number,
  headingScale: number
): TextSegment[][] {
  const lines: TextSegment[][] = [];
  let currentLine: TextSegment[] = [];
  let currentLineWidth = 0;
  let isFirstLine = true;
  let effectiveWidth = availableWidth - firstLineIndent;

  for (const segment of segments) {
    const fontSize = getRunFontSizePt(segment.properties) * headingScale;
    const fontName = mapToStandardFont(resolveRunFontName(segment.properties));
    const segmentWidth = measureTextWidth(segment.text, fontName, fontSize);

    if (currentLineWidth + segmentWidth <= effectiveWidth || currentLine.length === 0) {
      currentLine.push(segment);
      currentLineWidth += segmentWidth;
    } else {
      // Word-level splitting
      const words = segment.text.split(/(\s+)/);
      let bufferedText = "";
      let bufferedWidth = 0;

      for (const word of words) {
        const wordWidth = measureTextWidth(word, fontName, fontSize);
        if (
          currentLineWidth + bufferedWidth + wordWidth <= effectiveWidth ||
          (currentLine.length === 0 && bufferedText.length === 0)
        ) {
          bufferedText += word;
          bufferedWidth += wordWidth;
        } else {
          if (bufferedText.length > 0) {
            currentLine.push({ text: bufferedText, properties: segment.properties });
          }
          lines.push(currentLine);
          currentLine = [];
          currentLineWidth = 0;
          if (isFirstLine) {
            isFirstLine = false;
            effectiveWidth = availableWidth;
          }
          bufferedText = word;
          bufferedWidth = wordWidth;
        }
      }
      if (bufferedText.length > 0) {
        currentLine.push({ text: bufferedText, properties: segment.properties });
        currentLineWidth += bufferedWidth;
      }
    }
  }

  if (currentLine.length > 0) {
    lines.push(currentLine);
  }

  if (lines.length === 0 && segments.length > 0) {
    lines.push(segments);
  }

  return lines;
}

function getHeadingLevel(props: ParagraphProperties | undefined): number {
  if (!props) {
    return 0;
  }
  if (props.outlineLevel !== undefined && props.outlineLevel >= 0 && props.outlineLevel <= 5) {
    return props.outlineLevel + 1;
  }
  if (props.style) {
    const match = /^[Hh]eading\s*(\d)$/i.exec(props.style);
    if (match) {
      return parseInt(match[1], 10);
    }
  }
  return 0;
}

function getHeadingFontScale(level: number): number {
  switch (level) {
    case 1:
      return 2.0;
    case 2:
      return 1.5;
    case 3:
      return 1.17;
    case 4:
      return 1.0;
    case 5:
      return 0.83;
    case 6:
      return 0.67;
    default:
      return 1.0;
  }
}

function getRunFontSizePt(props: Run["properties"]): number {
  if (props?.size) {
    return props.size / 2;
  }
  return DEFAULT_FONT_SIZE_PT;
}

function resolveRunFontName(props: Run["properties"]): string {
  if (!props?.font) {
    return "Calibri";
  }
  if (typeof props.font === "string") {
    return props.font;
  }
  return (props.font as { ascii?: string }).ascii ?? "Calibri";
}

function resolveColorHex(
  color: Run["properties"] extends { color?: infer C } ? C : unknown
): string | undefined {
  if (!color) {
    return undefined;
  }
  if (typeof color === "string") {
    return color;
  }
  if (typeof color === "object" && color !== null && "value" in (color as object)) {
    return (color as { value: string }).value;
  }
  return undefined;
}
