/**
 * Word Document Page Renderer — SVG Output
 *
 * Renders DOCX document pages to SVG strings for visual preview.
 * Uses the layout engine for pagination and produces approximate
 * visual representations of document content.
 *
 * @stability experimental
 */

import { measureTextWidth, mapToStandardFont } from "@utils/font-metrics";
import { xmlEncode, xmlEncodeAttr } from "@xml/encode";

import { isHyperlink, isRun } from "../core/text-utils";
import type {
  BodyContent,
  DocxDocument,
  FontSpec,
  Paragraph,
  ParagraphProperties,
  Run,
  RunProperties,
  SectionProperties,
  Table,
  TableRow
} from "../types";
import { EMU_PER_POINT } from "../units";
import { layoutDocument } from "./layout";
import type { LayoutResult } from "./layout";
import {
  DEFAULT_PAGE_HEIGHT_TWIPS,
  DEFAULT_PAGE_MARGIN_TWIPS,
  DEFAULT_PAGE_WIDTH_TWIPS
} from "./layout-constants";
import type {
  LayoutChart,
  LayoutCheckBox,
  LayoutDocument,
  LayoutFloat,
  LayoutImage,
  LayoutMath,
  LayoutParagraph,
  LayoutRect,
  LayoutSdt,
  LayoutShape,
  LayoutTable,
  LayoutTableOfContents,
  LayoutTextBox,
  PageContent,
  PageGeometry
} from "./layout-model";

// =============================================================================
// Public API Types
// =============================================================================

/** Options for rendering document pages to SVG. */
export interface RenderOptions {
  /** Output SVG width in pixels. If not set, derived from page dimensions. */
  readonly width?: number;
  /** Output SVG height in pixels. If not set, derived from page dimensions. */
  readonly height?: number;
  /** Font family mapping: document font name → SVG font-family value. */
  readonly fonts?: ReadonlyMap<string, string>;
  /** Background color (CSS color string). Default: "white". */
  readonly backgroundColor?: string;
  /** Scale factor for the output. Default: 1.0. */
  readonly scale?: number;
}

// =============================================================================
// Internal Constants
// =============================================================================

/** Conversion: twips to points. */
const TWIPS_TO_PT = 1 / 20;
/** Conversion: EMU to points. */
const EMU_TO_PT = 1 / EMU_PER_POINT;
/** Default font size in points. */
const DEFAULT_FONT_SIZE_PT = 12;
/** Default font name. */
const DEFAULT_FONT = "Calibri";

// =============================================================================
// Internal Helpers
// =============================================================================

/** Convert twips to points. */
function twipsToPt(twips: number): number {
  return twips * TWIPS_TO_PT;
}

/** Convert EMU to points. */
function emuToPt(emu: number): number {
  return emu * EMU_TO_PT;
}

/** Extract font name from a Run's properties. */
function getRunFontName(run: Run): string {
  const font = run.properties?.font;
  if (!font) {
    return DEFAULT_FONT;
  }
  if (typeof font === "string") {
    return font;
  }
  return (font as FontSpec).ascii ?? (font as FontSpec).hAnsi ?? DEFAULT_FONT;
}

/** Get font size in points from RunProperties. */
function getRunFontSizePt(props: RunProperties | undefined): number {
  if (props?.size) {
    return props.size / 2;
  }
  return DEFAULT_FONT_SIZE_PT;
}

/** Resolve font family for SVG (uses the fonts map if provided). */
function resolveFontFamily(
  fontName: string,
  fontsMap: ReadonlyMap<string, string> | undefined
): string {
  if (fontsMap) {
    const mapped = fontsMap.get(fontName);
    if (mapped) {
      return mapped;
    }
  }
  // Use generic sans-serif fallback for common fonts
  const lower = fontName.toLowerCase();
  if (lower.includes("times") || lower.includes("roman") || lower.includes("serif")) {
    return `"${fontName}", serif`;
  }
  if (lower.includes("courier") || lower.includes("mono") || lower.includes("consolas")) {
    return `"${fontName}", monospace`;
  }
  return `"${fontName}", sans-serif`;
}

/** Sanitize a string into a valid 6-digit hex color, or undefined. */
function sanitizeHexColor(raw: string | undefined): string | undefined {
  if (!raw) {
    return undefined;
  }
  const stripped = raw.replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(stripped) || /^[0-9a-fA-F]{3}$/.test(stripped)) {
    return stripped;
  }
  return undefined;
}

/** Get CSS color from a hex color string or "auto". */
function resolveColor(color: string | { val: string } | undefined): string | undefined {
  if (!color) {
    return undefined;
  }
  if (typeof color === "string") {
    if (color === "auto") {
      return undefined;
    }
    const safe = sanitizeHexColor(color);
    return safe ? `#${safe}` : undefined;
  }
  if (color.val === "auto") {
    return undefined;
  }
  const safe = sanitizeHexColor(color.val);
  return safe ? `#${safe}` : undefined;
}

/** Determine heading level from paragraph style name. Returns 0 for non-headings. */
function getHeadingLevel(props: ParagraphProperties | undefined): number {
  const style = props?.style;
  if (!style) {
    return 0;
  }
  const match = /^[Hh]eading(\d)$/.exec(style);
  if (match) {
    return parseInt(match[1], 10);
  }
  return 0;
}

/** Get font size multiplier based on heading level. */
function getHeadingFontScale(level: number): number {
  switch (level) {
    case 1:
      return 2.0;
    case 2:
      return 1.6;
    case 3:
      return 1.4;
    case 4:
      return 1.2;
    case 5:
      return 1.1;
    case 6:
      return 1.0;
    default:
      return 1.0;
  }
}

/** Extract plain text from a Run. */
function getRunText(run: Run): string {
  let text = "";
  for (const item of run.content) {
    switch (item.type) {
      case "text":
        text += item.text;
        break;
      case "tab":
        text += "    ";
        break;
      case "symbol":
        text += " ";
        break;
      case "noBreakHyphen":
      case "softHyphen":
        text += "-";
        break;
      default:
        break;
    }
  }
  return text;
}

/** Get section properties for a given section index. */
function getSectionProps(
  doc: DocxDocument,
  layout: LayoutResult,
  pageNumber: number
): SectionProperties | undefined {
  // Find which section this page belongs to
  let sectionIdx = 0;
  let accumulatedPages = 0;
  for (let s = 0; s < layout.sectionPageCounts.length; s++) {
    accumulatedPages += layout.sectionPageCounts[s];
    if (pageNumber <= accumulatedPages) {
      sectionIdx = s;
      break;
    }
  }

  // Find the section properties for this section
  // Walk through body to find the nth section break
  let currentSec = 0;
  for (const item of doc.body) {
    if (item.type === "paragraph" && item.properties?.sectionProperties) {
      if (currentSec === sectionIdx) {
        return item.properties.sectionProperties;
      }
      currentSec++;
    }
  }
  // Last section uses doc.sectionProperties
  return doc.sectionProperties;
}

// =============================================================================
// Rendering State
// =============================================================================

interface RenderState {
  readonly pageWidthPt: number;
  readonly pageHeightPt: number;
  readonly marginTopPt: number;
  readonly marginBottomPt: number;
  readonly marginLeftPt: number;
  readonly marginRightPt: number;
  readonly contentWidthPt: number;
  readonly contentHeightPt: number;
  readonly fontsMap: ReadonlyMap<string, string> | undefined;
  readonly doc: DocxDocument;
  cursorY: number;
  elements: string[];
}

// =============================================================================
// Paragraph Rendering
// =============================================================================

/** Render a paragraph and return the Y advance. */
function renderParagraph(para: Paragraph, state: RenderState): void {
  const props = para.properties;
  const spacing = props?.spacing;
  const headingLevel = getHeadingLevel(props);
  const headingScale = getHeadingFontScale(headingLevel);

  // Space before
  let spaceBefore = 0;
  if (spacing?.beforeAutoSpacing) {
    spaceBefore = 5; // ~5pt auto spacing
  } else if (spacing?.before != null) {
    spaceBefore = twipsToPt(spacing.before);
  }
  state.cursorY += spaceBefore;

  // Thematic break (horizontal rule)
  if (props?.thematicBreak) {
    const lineY = state.cursorY;
    state.elements.push(
      `<line x1="${state.marginLeftPt}" y1="${lineY}" x2="${state.marginLeftPt + state.contentWidthPt}" y2="${lineY}" stroke="#999999" stroke-width="1"/>`
    );
    state.cursorY += 6;
    return;
  }

  // Determine paragraph indentation
  const indent = props?.indent;
  const leftIndentPt = indent?.left ? twipsToPt(indent.left) : 0;
  const firstLineIndentPt = indent?.firstLine ? twipsToPt(indent.firstLine) : 0;

  // Determine alignment
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

  // Apply heading scale to line height
  lineHeightPt *= headingScale;

  // Collect runs and render as text spans on lines
  const runs = collectParagraphRuns(para);

  if (runs.length === 0) {
    // Empty paragraph — advance by line height
    state.cursorY += lineHeightPt;
  } else {
    // Simple line wrapping: measure text and wrap at content width
    const availableWidth = state.contentWidthPt - leftIndentPt;
    const lines = wrapRunsToLines(runs, availableWidth, firstLineIndentPt, headingScale);

    for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
      const line = lines[lineIdx];
      state.cursorY += lineHeightPt;

      // Calculate x position based on alignment
      let xOffset = state.marginLeftPt + leftIndentPt;
      if (lineIdx === 0) {
        xOffset += firstLineIndentPt;
      }

      if (alignment === "center") {
        const lineWidth = measureLineWidth(line, headingScale);
        xOffset = state.marginLeftPt + (state.contentWidthPt - lineWidth) / 2;
      } else if (alignment === "right" || alignment === "end") {
        const lineWidth = measureLineWidth(line, headingScale);
        xOffset = state.marginLeftPt + state.contentWidthPt - lineWidth;
      }

      // Render each run segment on this line
      let currentX = xOffset;
      for (const segment of line) {
        const fontSize = getRunFontSizePt(segment.properties) * headingScale;
        const fontFamily = resolveFontFamily(
          getRunFontName({ properties: segment.properties, content: [] }),
          state.fontsMap
        );
        const isBold = segment.properties?.bold || headingLevel > 0;
        const isItalic = segment.properties?.italic;
        const color = resolveColor(segment.properties?.color);
        const underline = segment.properties?.underline;
        const strike = segment.properties?.strike;

        // Build style attributes
        let attrs = `x="${currentX.toFixed(2)}" y="${state.cursorY.toFixed(2)}"`;
        attrs += ` font-family="${xmlEncodeAttr(fontFamily)}"`;
        attrs += ` font-size="${fontSize.toFixed(1)}"`;
        if (isBold) {
          attrs += ` font-weight="bold"`;
        }
        if (isItalic) {
          attrs += ` font-style="italic"`;
        }
        if (color) {
          attrs += ` fill="${color}"`;
        }
        if (underline && underline !== "none") {
          // underline can be true, a style string, or UnderlineSpec
          const isNone =
            typeof underline === "object" && "style" in underline && underline.style === "none";
          if (!isNone) {
            attrs += ` text-decoration="underline"`;
          }
        } else if (strike) {
          attrs += ` text-decoration="line-through"`;
        }

        const escapedText = xmlEncode(segment.text);
        if (escapedText.length > 0) {
          state.elements.push(`<text ${attrs}>${escapedText}</text>`);
        }

        // Advance X position using text measurement
        const measuredFont = mapToStandardFont(
          getRunFontName({ properties: segment.properties, content: [] })
        );
        const textWidth = measureTextWidth(segment.text, measuredFont, fontSize);
        currentX += textWidth;
      }
    }
  }

  // Handle inline images within the paragraph
  renderParagraphImages(para, state);

  // Space after
  let spaceAfter = 0;
  if (spacing?.afterAutoSpacing) {
    spaceAfter = 5;
  } else if (spacing?.after != null) {
    spaceAfter = twipsToPt(spacing.after);
  }
  state.cursorY += spaceAfter;
}

/** A text segment with its formatting. */
interface TextSegment {
  readonly text: string;
  readonly properties: RunProperties | undefined;
}

/** Collect all text segments from a paragraph's children. */
function collectParagraphRuns(para: Paragraph): TextSegment[] {
  const segments: TextSegment[] = [];
  for (const child of para.children) {
    if (isRun(child)) {
      const text = getRunText(child);
      if (text.length > 0) {
        segments.push({ text, properties: child.properties });
      }
    } else if (isHyperlink(child)) {
      for (const run of child.children) {
        const text = getRunText(run);
        if (text.length > 0) {
          segments.push({ text, properties: run.properties });
        }
      }
    }
  }
  return segments;
}

/** Wrap text segments into lines based on available width. */
function wrapRunsToLines(
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
    const fontName = mapToStandardFont(
      getRunFontName({ properties: segment.properties, content: [] })
    );
    const segmentWidth = measureTextWidth(segment.text, fontName, fontSize);

    if (currentLineWidth + segmentWidth <= effectiveWidth || currentLine.length === 0) {
      // Fits on current line
      currentLine.push(segment);
      currentLineWidth += segmentWidth;
    } else {
      // Need to wrap — try word-level splitting
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
          // Flush buffered text to current line
          if (bufferedText.length > 0) {
            currentLine.push({ text: bufferedText, properties: segment.properties });
          }
          // Start new line
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
      // Flush remaining buffered text
      if (bufferedText.length > 0) {
        currentLine.push({ text: bufferedText, properties: segment.properties });
        currentLineWidth += bufferedWidth;
      }
    }
  }

  if (currentLine.length > 0) {
    lines.push(currentLine);
  }

  // Ensure at least one empty line for non-empty segments
  if (lines.length === 0 && segments.length > 0) {
    lines.push(segments);
  }

  return lines;
}

/** Measure the total width of a line of segments in points. */
function measureLineWidth(segments: TextSegment[], headingScale: number): number {
  let total = 0;
  for (const segment of segments) {
    const fontSize = getRunFontSizePt(segment.properties) * headingScale;
    const fontName = mapToStandardFont(
      getRunFontName({ properties: segment.properties, content: [] })
    );
    total += measureTextWidth(segment.text, fontName, fontSize);
  }
  return total;
}

/** Render inline images found in paragraph runs. */
function renderParagraphImages(para: Paragraph, state: RenderState): void {
  for (const child of para.children) {
    if (!isRun(child)) {
      continue;
    }
    for (const item of child.content) {
      if (item.type === "image") {
        const img = item as {
          type: "image";
          rId: string;
          width: number;
          height: number;
          altText?: string;
        };
        const widthPt = emuToPt(img.width);
        const heightPt = emuToPt(img.height);

        // Find image data from document
        const imgData = findImageData(state.doc, img.rId);
        if (imgData) {
          const dataUri = `data:${imageMediaTypeToMime(imgData.mediaType)};base64,${uint8ToBase64(imgData.data)}`;
          state.elements.push(
            `<image x="${state.marginLeftPt}" y="${state.cursorY}" width="${widthPt.toFixed(1)}" height="${heightPt.toFixed(1)}" href="${xmlEncodeAttr(dataUri)}"/>`
          );
        } else {
          // Render placeholder rectangle
          state.elements.push(
            `<rect x="${state.marginLeftPt}" y="${state.cursorY}" width="${widthPt.toFixed(1)}" height="${heightPt.toFixed(1)}" fill="#f0f0f0" stroke="#cccccc" stroke-width="0.5"/>` +
              `<text x="${(state.marginLeftPt + widthPt / 2).toFixed(1)}" y="${(state.cursorY + heightPt / 2).toFixed(1)}" text-anchor="middle" font-size="8" fill="#999999">[Image]</text>`
          );
        }
        state.cursorY += heightPt;
      }
    }
  }
}

/** Find image data by relationship ID. */
function findImageData(
  doc: DocxDocument,
  rId: string
): { data: Uint8Array; mediaType: string } | undefined {
  if (!doc.images) {
    return undefined;
  }
  for (const img of doc.images) {
    if (img.rId === rId) {
      return { data: img.data, mediaType: img.mediaType };
    }
  }
  return undefined;
}

/**
 * Map an `ImageMediaType` value (e.g. "png", "svg") to the MIME type required
 * by `data:` URIs. The model uses short tokens, but data URIs need full
 * `image/<subtype>` form — without this, browsers fail to decode the SVG
 * preview's embedded images.
 */
function imageMediaTypeToMime(mediaType: string): string {
  switch (mediaType) {
    case "svg":
      return "image/svg+xml";
    case "jpeg":
    case "png":
    case "gif":
    case "bmp":
    case "tiff":
    case "webp":
      return `image/${mediaType}`;
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    default:
      return `image/${mediaType}`;
  }
}

/** Convert Uint8Array to base64 string. */
function uint8ToBase64(data: Uint8Array): string {
  // Process in chunks to avoid call stack overflow with large arrays
  const CHUNK_SIZE = 8192;
  let binary = "";
  for (let i = 0; i < data.length; i += CHUNK_SIZE) {
    const chunk = data.subarray(i, Math.min(i + CHUNK_SIZE, data.length));
    for (let j = 0; j < chunk.length; j++) {
      binary += String.fromCharCode(chunk[j]);
    }
  }
  return btoa(binary);
}

// =============================================================================
// Table Rendering
// =============================================================================

/** Render a table. */
function renderTable(table: Table, state: RenderState): void {
  const startY = state.cursorY;
  const tableX = state.marginLeftPt;

  // Calculate column widths
  const colWidths = computeTableColumnWidths(table, state.contentWidthPt);
  const totalWidth = colWidths.reduce((sum, w) => sum + w, 0);

  // Track row Y positions for border drawing
  const rowYPositions: number[] = [startY];

  for (const row of table.rows) {
    const rowStartY = state.cursorY;
    const rowHeight = estimateRowHeightPt(row, colWidths, state);

    // Render cell backgrounds and content
    let cellX = tableX;
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      const gridSpan = cell.properties?.gridSpan ?? 1;
      let cellWidth = 0;
      for (let g = 0; g < gridSpan && c + g < colWidths.length; g++) {
        cellWidth += colWidths[c + g];
      }

      // Cell background
      const shading = cell.properties?.shading;
      if (shading?.fill && shading.fill !== "auto") {
        const safeFill = sanitizeHexColor(shading.fill);
        if (safeFill) {
          state.elements.push(
            `<rect x="${cellX.toFixed(2)}" y="${rowStartY.toFixed(2)}" width="${cellWidth.toFixed(2)}" height="${rowHeight.toFixed(2)}" fill="#${safeFill}" stroke="none"/>`
          );
        }
      }

      // Render cell content (simplified — just text)
      const cellPadding = 4; // ~4pt padding
      let cellCursorY = rowStartY + cellPadding;
      for (const content of cell.content) {
        if (content.type === "paragraph") {
          // Simplified: render first run of each paragraph
          const runs = collectParagraphRuns(content);
          if (runs.length > 0) {
            const allText = runs.map(r => r.text).join("");
            const firstRun = runs[0];
            const fontSize = getRunFontSizePt(firstRun.properties);
            const fontFamily = resolveFontFamily(
              getRunFontName({ properties: firstRun.properties, content: [] }),
              state.fontsMap
            );
            const isBold = firstRun.properties?.bold;
            cellCursorY += fontSize * 1.2;

            let textAttrs = `x="${(cellX + cellPadding).toFixed(2)}" y="${cellCursorY.toFixed(2)}"`;
            textAttrs += ` font-family="${xmlEncodeAttr(fontFamily)}"`;
            textAttrs += ` font-size="${fontSize.toFixed(1)}"`;
            if (isBold) {
              textAttrs += ` font-weight="bold"`;
            }

            // Clip text to cell width
            const maxChars = Math.floor((cellWidth - cellPadding * 2) / (fontSize * 0.5));
            const displayText =
              allText.length > maxChars ? allText.slice(0, maxChars) + "…" : allText;

            state.elements.push(`<text ${textAttrs}>${xmlEncode(displayText)}</text>`);
          } else {
            cellCursorY += DEFAULT_FONT_SIZE_PT * 1.2;
          }
        }
      }

      cellX += cellWidth;
    }

    state.cursorY += rowHeight;
    rowYPositions.push(state.cursorY);
  }

  // Draw table borders (outer rect + row/col lines)
  const tableEndY = state.cursorY;
  const borderColor = getTableBorderColor(table);

  // Outer border
  state.elements.push(
    `<rect x="${tableX.toFixed(2)}" y="${startY.toFixed(2)}" width="${totalWidth.toFixed(2)}" height="${(tableEndY - startY).toFixed(2)}" fill="none" stroke="${borderColor}" stroke-width="0.75"/>`
  );

  // Horizontal row lines
  for (let r = 1; r < rowYPositions.length - 1; r++) {
    const y = rowYPositions[r];
    state.elements.push(
      `<line x1="${tableX.toFixed(2)}" y1="${y.toFixed(2)}" x2="${(tableX + totalWidth).toFixed(2)}" y2="${y.toFixed(2)}" stroke="${borderColor}" stroke-width="0.5"/>`
    );
  }

  // Vertical column lines
  let colX = tableX;
  for (let c = 0; c < colWidths.length - 1; c++) {
    colX += colWidths[c];
    state.elements.push(
      `<line x1="${colX.toFixed(2)}" y1="${startY.toFixed(2)}" x2="${colX.toFixed(2)}" y2="${tableEndY.toFixed(2)}" stroke="${borderColor}" stroke-width="0.5"/>`
    );
  }

  // Add some spacing after the table
  state.cursorY += 6;
}

/** Compute column widths in points for a table. */
function computeTableColumnWidths(table: Table, availableWidthPt: number): number[] {
  if (table.columnWidths && table.columnWidths.length > 0) {
    return table.columnWidths.map(w => twipsToPt(w));
  }
  // Fall back to equal distribution based on max cell count
  let maxCols = 0;
  for (const row of table.rows) {
    if (row.cells.length > maxCols) {
      maxCols = row.cells.length;
    }
  }
  if (maxCols === 0) {
    return [availableWidthPt];
  }
  const colWidth = availableWidthPt / maxCols;
  return Array.from({ length: maxCols }, () => colWidth);
}

/** Estimate row height in points. */
function estimateRowHeightPt(row: TableRow, colWidths: number[], state: RenderState): number {
  // Check explicit height
  if (row.properties?.height?.value) {
    const rule = row.properties.height.rule;
    if (rule === "exact") {
      return twipsToPt(row.properties.height.value);
    }
  }

  // Estimate from content
  let maxHeight = DEFAULT_FONT_SIZE_PT * 1.2 + 8; // minimum: one line + padding
  for (let c = 0; c < row.cells.length; c++) {
    const cell = row.cells[c];
    let cellHeight = 8; // padding
    for (const content of cell.content) {
      if (content.type === "paragraph") {
        const runs = collectParagraphRuns(content);
        const lineHeight = DEFAULT_FONT_SIZE_PT * 1.2;
        if (runs.length === 0) {
          cellHeight += lineHeight;
        } else {
          // Estimate line count
          const allText = runs.map(r => r.text).join("");
          const gridSpan = cell.properties?.gridSpan ?? 1;
          let cellWidth = 0;
          for (let g = 0; g < gridSpan && c + g < colWidths.length; g++) {
            cellWidth += colWidths[c + g];
          }
          const charsPerLine = Math.max(
            1,
            Math.floor((cellWidth - 8) / (DEFAULT_FONT_SIZE_PT * 0.5))
          );
          const lineCount = Math.max(1, Math.ceil(allText.length / charsPerLine));
          cellHeight += lineCount * lineHeight;
        }
      }
    }
    if (cellHeight > maxHeight) {
      maxHeight = cellHeight;
    }
  }

  // Respect atLeast constraint
  if (row.properties?.height?.value) {
    return Math.max(twipsToPt(row.properties.height.value), maxHeight);
  }

  return maxHeight;
}

/** Get the border color for a table. */
function getTableBorderColor(table: Table): string {
  const borders = table.properties?.borders;
  if (borders?.top?.color && borders.top.color !== "auto") {
    const safe = sanitizeHexColor(borders.top.color);
    if (safe) {
      return `#${safe}`;
    }
  }
  return "#000000";
}

// =============================================================================
// Floating Image Rendering
// =============================================================================

/** Render a floating image. */
function renderFloatingImage(item: BodyContent, state: RenderState): void {
  const img = item as {
    type: "floatingImage";
    rId: string;
    width: number;
    height: number;
    altText?: string;
    wrap?: { style: string };
  };

  const widthPt = emuToPt(img.width);
  const heightPt = emuToPt(img.height);

  const imgData = findImageData(state.doc, img.rId);
  if (imgData) {
    const dataUri = `data:${imageMediaTypeToMime(imgData.mediaType)};base64,${uint8ToBase64(imgData.data)}`;
    state.elements.push(
      `<image x="${state.marginLeftPt}" y="${state.cursorY.toFixed(2)}" width="${widthPt.toFixed(1)}" height="${heightPt.toFixed(1)}" href="${xmlEncodeAttr(dataUri)}"/>`
    );
  } else {
    state.elements.push(
      `<rect x="${state.marginLeftPt}" y="${state.cursorY.toFixed(2)}" width="${widthPt.toFixed(1)}" height="${heightPt.toFixed(1)}" fill="#f0f0f0" stroke="#cccccc" stroke-width="0.5"/>` +
        `<text x="${(state.marginLeftPt + widthPt / 2).toFixed(1)}" y="${(state.cursorY + heightPt / 2).toFixed(1)}" text-anchor="middle" font-size="8" fill="#999999">[Image]</text>`
    );
  }

  // topAndBottom wrap consumes space
  const wrapStyle = img.wrap?.style;
  if (wrapStyle === "topAndBottom" || !wrapStyle) {
    state.cursorY += heightPt;
  }
}

// =============================================================================
// Page Rendering Core
// =============================================================================

/** Render a single page's content to SVG elements. */
function renderPageContent(
  doc: DocxDocument,
  layout: LayoutResult,
  pageNumber: number,
  sectionProps: SectionProperties | undefined,
  options: RenderOptions | undefined
): { elements: string[]; pageWidthPt: number; pageHeightPt: number } {
  const pageWidthPt = twipsToPt(sectionProps?.pageSize?.width ?? DEFAULT_PAGE_WIDTH_TWIPS);
  const pageHeightPt = twipsToPt(sectionProps?.pageSize?.height ?? DEFAULT_PAGE_HEIGHT_TWIPS);
  const marginTopPt = twipsToPt(sectionProps?.margins?.top ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const marginBottomPt = twipsToPt(sectionProps?.margins?.bottom ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const marginLeftPt = twipsToPt(sectionProps?.margins?.left ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const marginRightPt = twipsToPt(sectionProps?.margins?.right ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const contentWidthPt = pageWidthPt - marginLeftPt - marginRightPt;
  const contentHeightPt = pageHeightPt - marginTopPt - marginBottomPt;

  const state: RenderState = {
    pageWidthPt,
    pageHeightPt,
    marginTopPt,
    marginBottomPt,
    marginLeftPt,
    marginRightPt,
    contentWidthPt,
    contentHeightPt,
    fontsMap: options?.fonts,
    doc,
    cursorY: marginTopPt,
    elements: []
  };

  // Render body content that belongs to this page
  for (let i = 0; i < doc.body.length; i++) {
    const contentPage = layout.contentPages[i];
    if (contentPage !== pageNumber) {
      continue;
    }

    // Stop rendering if we've exceeded the content area
    if (state.cursorY > pageHeightPt - marginBottomPt) {
      break;
    }

    const item = doc.body[i];
    switch (item.type) {
      case "paragraph":
        renderParagraph(item, state);
        break;
      case "table":
        renderTable(item, state);
        break;
      case "floatingImage":
        renderFloatingImage(item, state);
        break;
      default:
        // Skip unsupported content types (TOC, shapes, etc.)
        // Advance cursor slightly to maintain spacing
        state.cursorY += DEFAULT_FONT_SIZE_PT * 1.2;
        break;
    }
  }

  return { elements: state.elements, pageWidthPt, pageHeightPt };
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Render a specific page of a DOCX document to an SVG string.
 *
 * @param doc - The parsed DOCX document
 * @param pageNumber - 1-based page number to render
 * @param options - Rendering options
 * @returns SVG string for the specified page
 *
 * @stability experimental
 */
export function renderPageToSvg(
  doc: DocxDocument,
  pageNumber: number,
  options?: RenderOptions
): string {
  const layout = layoutDocument(doc, {
    measureText: (text, fontName, fontSize) =>
      measureTextWidth(text, mapToStandardFont(fontName), fontSize)
  });

  if (pageNumber < 1 || pageNumber > layout.pageCount) {
    throw new RangeError(
      `Page number ${pageNumber} out of range. Document has ${layout.pageCount} page(s).`
    );
  }

  const sectionProps = getSectionProps(doc, layout, pageNumber);
  const { elements, pageWidthPt, pageHeightPt } = renderPageContent(
    doc,
    layout,
    pageNumber,
    sectionProps,
    options
  );

  return buildSvgDocument(elements, pageWidthPt, pageHeightPt, options);
}

/**
 * Render all pages of a DOCX document to SVG strings.
 *
 * @param doc - The parsed DOCX document
 * @param options - Rendering options
 * @returns Array of SVG strings, one per page
 *
 * @stability experimental
 */
export function renderDocumentToSvg(doc: DocxDocument, options?: RenderOptions): string[] {
  const layout = layoutDocument(doc, {
    measureText: (text, fontName, fontSize) =>
      measureTextWidth(text, mapToStandardFont(fontName), fontSize)
  });

  const svgPages: string[] = [];

  for (let page = 1; page <= layout.pageCount; page++) {
    const sectionProps = getSectionProps(doc, layout, page);
    const { elements, pageWidthPt, pageHeightPt } = renderPageContent(
      doc,
      layout,
      page,
      sectionProps,
      options
    );
    svgPages.push(buildSvgDocument(elements, pageWidthPt, pageHeightPt, options));
  }

  return svgPages;
}

/** Build the final SVG document string from rendered elements. */
function buildSvgDocument(
  elements: string[],
  pageWidthPt: number,
  pageHeightPt: number,
  options: RenderOptions | undefined
): string {
  const scale = options?.scale ?? 1.0;
  const bgColor = options?.backgroundColor ?? "white";

  // Determine output dimensions
  let outputWidth: number;
  let outputHeight: number;

  if (options?.width && options?.height) {
    outputWidth = options.width;
    outputHeight = options.height;
  } else if (options?.width) {
    outputWidth = options.width;
    outputHeight = (pageHeightPt / pageWidthPt) * options.width;
  } else if (options?.height) {
    outputHeight = options.height;
    outputWidth = (pageWidthPt / pageHeightPt) * options.height;
  } else {
    outputWidth = pageWidthPt * scale;
    outputHeight = pageHeightPt * scale;
  }

  const viewBox = `0 0 ${pageWidthPt.toFixed(2)} ${pageHeightPt.toFixed(2)}`;

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${outputWidth.toFixed(2)}" height="${outputHeight.toFixed(2)}" viewBox="${viewBox}">\n`;
  svg += `  <rect width="100%" height="100%" fill="${xmlEncodeAttr(bgColor)}"/>\n`;

  for (const element of elements) {
    svg += `  ${element}\n`;
  }

  svg += `</svg>`;
  return svg;
}

// =============================================================================
// New API: Render from pre-computed LayoutDocument (no re-layout)
// =============================================================================

/**
 * Render a page from a pre-computed LayoutDocument to SVG.
 * This avoids re-computing layout — just serializes positioned elements to SVG.
 *
 * @param layout - A LayoutDocument produced by layoutDocumentFull().
 * @param pageNumber - 1-based page number to render.
 * @param options - Rendering options (scale, dimensions, background color).
 * @returns SVG string.
 *
 * @stability experimental
 */
export function renderPageFromLayout(
  layout: LayoutDocument,
  pageNumber: number,
  options?: RenderOptions
): string {
  if (pageNumber < 1 || pageNumber > layout.totalPages) {
    throw new RangeError(
      `Page number ${pageNumber} out of range. Document has ${layout.totalPages} page(s).`
    );
  }

  const page = layout.pages[pageNumber - 1];
  const { geometry } = page;
  const elements: string[] = [];

  // Header / footer paragraphs and tables come with layout-y already
  // expressed as a page-absolute offset (the layout engine adds the
  // section's `pgMar.header` to header content and starts footer
  // content at `pageHeight - pgMar.footer`). Use a geometry with
  // `marginTop: 0` so SVG y-coordinates resolve straight from
  // layout-y. Tables in header / footer are uncommon but legal —
  // dispatch through the same renderers used for body content.
  const bandGeometry: PageGeometry = { ...geometry, marginTop: 0 };
  if (page.header) {
    for (const item of page.header) {
      if (item.type === "paragraph") {
        renderLayoutParagraphToSvg(item, bandGeometry, elements);
      } else {
        renderLayoutTableToSvg(item, bandGeometry, elements);
      }
    }
  }

  // Render each positioned content element. Every PageContent variant is
  // handled — most non-paragraph/table types degrade to a placeholder rect
  // (charts, shapes, opaque drawings), an inline glyph (check-boxes, math),
  // or a recursive descent (text-boxes, SDTs). Adding a new PageContent
  // variant is a build error here until a case is added.
  for (const item of page.content) {
    switch (item.type) {
      case "paragraph":
        renderLayoutParagraphToSvg(item, geometry, elements);
        break;
      case "table":
        renderLayoutTableToSvg(item, geometry, elements);
        break;
      case "image":
        renderLayoutImageToSvg(item, geometry, elements);
        break;
      case "float":
        renderLayoutFloatToSvg(item, geometry, elements);
        break;
      case "textBox":
        renderLayoutTextBoxToSvg(item, geometry, elements);
        break;
      case "shape":
        renderLayoutShapeToSvg(item, geometry, elements);
        break;
      case "chart":
        renderLayoutChartToSvg(item, geometry, elements);
        break;
      case "sdt":
        renderLayoutSdtToSvg(item, geometry, elements);
        break;
      case "math":
        renderLayoutMathToSvg(item, geometry, elements);
        break;
      case "checkBox":
        renderLayoutCheckBoxToSvg(item, geometry, elements);
        break;
      case "tableOfContents":
        renderLayoutTocToSvg(item, geometry, elements);
        break;
      case "altChunk":
        renderLayoutPlaceholderToSvg(item.rect, geometry, `[${item.contentType}]`, elements);
        break;
      case "opaqueDrawing":
        renderLayoutPlaceholderToSvg(item.rect, geometry, "[drawing]", elements);
        break;
      default: {
        const _exhaustive: never = item;
        throw new Error(
          `renderPageFromLayout: unhandled PageContent ${(_exhaustive as { type: string }).type}`
        );
      }
    }
  }

  // Footer — `bandGeometry` (declared near the header above) shares the
  // same "layout-y is page-absolute" rule for both bands.
  if (page.footer) {
    for (const item of page.footer) {
      if (item.type === "paragraph") {
        renderLayoutParagraphToSvg(item, bandGeometry, elements);
      } else {
        renderLayoutTableToSvg(item, bandGeometry, elements);
      }
    }
  }

  // Footnote separator (drawn above the footnote area; see ECMA-376
  // §17.11.10). Same coordinate convention as bands: the layout y
  // value is page-absolute.
  if (page.footnoteSeparator) {
    const sep = page.footnoteSeparator;
    const ruleWidth = sep.kind === "separator" ? geometry.contentWidth / 3 : geometry.contentWidth;
    const x1 = geometry.marginLeft;
    const x2 = x1 + ruleWidth;
    elements.push(
      `<line x1="${x1.toFixed(2)}" y1="${sep.y.toFixed(2)}" x2="${x2.toFixed(2)}" y2="${sep.y.toFixed(2)}" stroke="black" stroke-width="0.5"/>`
    );
  }

  // Footnote area paragraphs (page-absolute y, like header/footer).
  if (page.footnoteArea) {
    for (const para of page.footnoteArea) {
      renderLayoutParagraphToSvg(para, bandGeometry, elements);
    }
  }

  return buildSvgDocument(elements, geometry.width, geometry.height, options);
}

/** Render a LayoutParagraph to SVG text elements. */
function renderLayoutParagraphToSvg(
  para: LayoutParagraph,
  geometry: PageGeometry,
  elements: string[]
): void {
  for (const line of para.lines) {
    const lineY = geometry.marginTop + para.rect.y + line.y + line.baseline;
    const lineTopY = geometry.marginTop + para.rect.y + line.y;
    for (const item of line.runs) {
      if (item.type === "image") {
        // Inline image: bottom-aligned within the line, matching
        // Word's default for in-line images. Empty data → emit
        // nothing rather than an invalid <image href> with empty src.
        if (item.data.length === 0) {
          continue;
        }
        const x = geometry.marginLeft + item.x;
        const yBottom = lineTopY + Math.min(line.height, item.height);
        const yTop = yBottom - item.height;
        const dataUri = `data:${item.mimeType};base64,${bytesToBase64(item.data)}`;
        elements.push(
          `<image x="${x.toFixed(2)}" y="${yTop.toFixed(2)}" width="${item.width.toFixed(2)}" height="${item.height.toFixed(2)}" href="${xmlEncodeAttr(dataUri)}"/>`
        );
        continue;
      }
      const run = item;
      const x = geometry.marginLeft + run.x;
      // Sub/superscript: shift the SVG baseline. SVG y-axis points
      // downward (opposite of PDF), so superscript moves UP (smaller
      // y) and subscript moves DOWN (larger y) — opposite signs from
      // the PDF code.
      let runY = lineY;
      if (run.verticalAlign === "superscript") {
        runY = lineY - run.fontSize * 0.33;
      } else if (run.verticalAlign === "subscript") {
        runY = lineY + run.fontSize * 0.33;
      }
      let attrs = `x="${x.toFixed(2)}" y="${runY.toFixed(2)}"`;
      attrs += ` font-family="${xmlEncodeAttr(run.font)}"`;
      attrs += ` font-size="${run.fontSize.toFixed(1)}"`;
      if (run.bold) {
        attrs += ` font-weight="bold"`;
      }
      if (run.italic) {
        attrs += ` font-style="italic"`;
      }
      if (run.color) {
        const safe = sanitizeHexColor(run.color);
        if (safe) {
          attrs += ` fill="#${safe}"`;
        }
      }
      if (run.underline) {
        attrs += ` text-decoration="underline"`;
      } else if (run.strikethrough) {
        attrs += ` text-decoration="line-through"`;
      }
      const escapedText = xmlEncode(run.text);
      if (escapedText.length > 0) {
        elements.push(`<text ${attrs}>${escapedText}</text>`);
      }
    }
  }
}

/** Render a LayoutTable to SVG (borders + cell content). */
function renderLayoutTableToSvg(
  table: LayoutTable,
  geometry: PageGeometry,
  elements: string[]
): void {
  // Draw cell borders and content
  for (const cell of table.cells) {
    const cellX = geometry.marginLeft + table.rect.x + cell.rect.x;
    const cellY = geometry.marginTop + table.rect.y + cell.rect.y;
    const w = cell.rect.width;
    const h = cell.rect.height;

    // Background
    if (cell.backgroundColor) {
      const safeBg = sanitizeHexColor(cell.backgroundColor);
      if (safeBg) {
        elements.push(
          `<rect x="${cellX.toFixed(2)}" y="${cellY.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="#${safeBg}" stroke="none"/>`
        );
      }
    }

    // Border
    elements.push(
      `<rect x="${cellX.toFixed(2)}" y="${cellY.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" fill="none" stroke="#cccccc" stroke-width="0.5"/>`
    );

    // Cell paragraph content
    for (const content of cell.content) {
      if (content.type === "paragraph") {
        const offsetPara: LayoutParagraph = {
          ...content,
          rect: {
            ...content.rect,
            y: table.rect.y + cell.rect.y + content.rect.y
          }
        };
        // Offset line runs by cell x
        const offsetLines = content.lines.map(line => ({
          ...line,
          runs: line.runs.map(run => ({
            ...run,
            x: run.x + table.rect.x + cell.rect.x
          }))
        }));
        renderLayoutParagraphToSvg({ ...offsetPara, lines: offsetLines }, geometry, elements);
      }
    }
  }
}

// =============================================================================
// SVG renderers — extended PageContent variants
// =============================================================================

function absRect(rect: LayoutRect, geometry: PageGeometry): LayoutRect {
  return {
    x: geometry.marginLeft + rect.x,
    y: geometry.marginTop + rect.y,
    width: rect.width,
    height: rect.height
  };
}

function pushPlaceholder(
  abs: LayoutRect,
  fillStroke: { fill?: string; stroke?: string; strokeWidth?: number },
  elements: string[]
): void {
  const fill = fillStroke.fill ?? "none";
  const stroke = fillStroke.stroke ?? "#bbbbbb";
  const sw = (fillStroke.strokeWidth ?? 0.5).toFixed(2);
  elements.push(
    `<rect x="${abs.x.toFixed(2)}" y="${abs.y.toFixed(2)}" width="${abs.width.toFixed(2)}" height="${abs.height.toFixed(2)}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`
  );
}

function renderLayoutImageToSvg(
  img: LayoutImage,
  geometry: PageGeometry,
  elements: string[]
): void {
  const abs = absRect(img.rect, geometry);
  if (img.data.length === 0) {
    pushPlaceholder(abs, { stroke: "#888888" }, elements);
    return;
  }
  const dataUri = `data:${img.mimeType};base64,${bytesToBase64(img.data)}`;
  elements.push(
    `<image x="${abs.x.toFixed(2)}" y="${abs.y.toFixed(2)}" width="${abs.width.toFixed(2)}" height="${abs.height.toFixed(2)}" href="${xmlEncodeAttr(dataUri)}"/>`
  );
}

function renderLayoutFloatToSvg(
  float: LayoutFloat,
  geometry: PageGeometry,
  elements: string[]
): void {
  // Floats currently always wrap a LayoutImage. Behind-text floats render
  // before main content in document order; SVG is painters-algorithm so a
  // dedicated z-ordering pass would belong upstream — for now rendering
  // order matches `page.content` order which is good enough.
  if (float.content.type === "image") {
    renderLayoutImageToSvg(float.content, geometry, elements);
  } else {
    renderLayoutParagraphToSvg(float.content, geometry, elements);
  }
}

function renderLayoutTextBoxToSvg(
  tb: LayoutTextBox,
  geometry: PageGeometry,
  elements: string[]
): void {
  const abs = absRect(tb.rect, geometry);
  const safeStroke = tb.border ? sanitizeHexColor(tb.border.color) : undefined;
  const safeFill = tb.background ? sanitizeHexColor(tb.background) : undefined;
  pushPlaceholder(
    abs,
    {
      fill: safeFill ? `#${safeFill}` : undefined,
      stroke: safeStroke ? `#${safeStroke}` : undefined,
      strokeWidth: tb.border?.width
    },
    elements
  );
  // Translate inner content by the text-box origin and recurse through the
  // generic SVG dispatcher so nested content (paragraphs, tables, even
  // shapes) renders correctly.
  const innerGeometry: PageGeometry = {
    ...geometry,
    marginLeft: geometry.marginLeft + tb.rect.x,
    marginTop: geometry.marginTop + tb.rect.y
  };
  renderPageContentList(tb.content, innerGeometry, elements);
}

function renderLayoutShapeToSvg(
  shape: LayoutShape,
  geometry: PageGeometry,
  elements: string[]
): void {
  const abs = absRect(shape.rect, geometry);
  const fill = shape.fillColor ? `#${sanitizeHexColor(shape.fillColor)}` : "none";
  const stroke = shape.strokeColor ? `#${sanitizeHexColor(shape.strokeColor)}` : "#888888";
  const sw = (shape.strokeWidth ?? 0.75).toFixed(2);

  // Map a few common preset shapes; everything else falls back to a rect.
  if (shape.preset === "ellipse" || shape.preset === "oval") {
    const cx = (abs.x + abs.width / 2).toFixed(2);
    const cy = (abs.y + abs.height / 2).toFixed(2);
    const rx = (abs.width / 2).toFixed(2);
    const ry = (abs.height / 2).toFixed(2);
    elements.push(
      `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`
    );
  } else if (shape.preset === "line") {
    elements.push(
      `<line x1="${abs.x.toFixed(2)}" y1="${abs.y.toFixed(2)}" x2="${(abs.x + abs.width).toFixed(2)}" y2="${(abs.y + abs.height).toFixed(2)}" stroke="${stroke}" stroke-width="${sw}"/>`
    );
  } else {
    elements.push(
      `<rect x="${abs.x.toFixed(2)}" y="${abs.y.toFixed(2)}" width="${abs.width.toFixed(2)}" height="${abs.height.toFixed(2)}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}"/>`
    );
  }

  if (shape.textContent && shape.textContent.length > 0) {
    const innerGeometry: PageGeometry = {
      ...geometry,
      marginLeft: geometry.marginLeft + shape.rect.x,
      marginTop: geometry.marginTop + shape.rect.y
    };
    renderPageContentList(shape.textContent, innerGeometry, elements);
  }
}

function renderLayoutChartToSvg(
  chart: LayoutChart,
  geometry: PageGeometry,
  elements: string[]
): void {
  const abs = absRect(chart.rect, geometry);
  if (chart.svg) {
    // Inline a pre-rendered SVG fragment inside a <g> with an absolute
    // translate so it ends up at the right page coordinates.
    elements.push(
      `<g transform="translate(${abs.x.toFixed(2)} ${abs.y.toFixed(2)})">${chart.svg}</g>`
    );
    return;
  }
  pushPlaceholder(abs, { stroke: "#666666" }, elements);
  if (chart.title) {
    const cx = abs.x + abs.width / 2;
    const cy = abs.y + abs.height / 2;
    elements.push(
      `<text x="${cx.toFixed(2)}" y="${cy.toFixed(2)}" text-anchor="middle" dominant-baseline="central" font-family="Helvetica" font-size="10" fill="#444444">${xmlEncode(chart.title)}</text>`
    );
  }
}

function renderLayoutSdtToSvg(sdt: LayoutSdt, geometry: PageGeometry, elements: string[]): void {
  // SDT is transparent visually; recurse into its children using the
  // page-relative geometry but offset by the SDT's own rect.
  const innerGeometry: PageGeometry = {
    ...geometry,
    marginLeft: geometry.marginLeft + sdt.rect.x,
    marginTop: geometry.marginTop + sdt.rect.y
  };
  renderPageContentList(sdt.content, innerGeometry, elements);
}

function renderLayoutMathToSvg(math: LayoutMath, geometry: PageGeometry, elements: string[]): void {
  const abs = absRect(math.rect, geometry);
  // Render the plain-text fallback. Renderers that want true math display
  // can read `math.mathML` from the layout document directly.
  elements.push(
    `<text x="${abs.x.toFixed(2)}" y="${(abs.y + abs.height * 0.8).toFixed(2)}" font-family="serif" font-style="italic" font-size="${(abs.height * 0.7).toFixed(1)}" fill="#000">${xmlEncode(math.text)}</text>`
  );
}

function renderLayoutCheckBoxToSvg(
  cb: LayoutCheckBox,
  geometry: PageGeometry,
  elements: string[]
): void {
  const abs = absRect(cb.rect, geometry);
  // Draw the actual square so the rendered output is independent of font
  // availability.
  elements.push(
    `<rect x="${abs.x.toFixed(2)}" y="${abs.y.toFixed(2)}" width="${abs.height.toFixed(2)}" height="${abs.height.toFixed(2)}" fill="white" stroke="#000" stroke-width="0.75"/>`
  );
  if (cb.checked) {
    const x1 = abs.x + abs.height * 0.2;
    const y1 = abs.y + abs.height * 0.55;
    const x2 = abs.x + abs.height * 0.45;
    const y2 = abs.y + abs.height * 0.8;
    const x3 = abs.x + abs.height * 0.85;
    const y3 = abs.y + abs.height * 0.2;
    elements.push(
      `<polyline points="${x1.toFixed(2)},${y1.toFixed(2)} ${x2.toFixed(2)},${y2.toFixed(2)} ${x3.toFixed(2)},${y3.toFixed(2)}" fill="none" stroke="#000" stroke-width="1"/>`
    );
  }
}

function renderLayoutTocToSvg(
  toc: LayoutTableOfContents,
  geometry: PageGeometry,
  elements: string[]
): void {
  // TOC is a list of LayoutParagraphs; render them with a y-offset by the
  // TOC's own rect.
  const innerGeometry: PageGeometry = {
    ...geometry,
    marginLeft: geometry.marginLeft + toc.rect.x,
    marginTop: geometry.marginTop + toc.rect.y
  };
  for (const p of toc.entries) {
    renderLayoutParagraphToSvg(p, innerGeometry, elements);
  }
}

function renderLayoutPlaceholderToSvg(
  rect: LayoutRect,
  geometry: PageGeometry,
  label: string,
  elements: string[]
): void {
  const abs = absRect(rect, geometry);
  pushPlaceholder(abs, { stroke: "#888888" }, elements);
  const cx = abs.x + abs.width / 2;
  const cy = abs.y + abs.height / 2;
  elements.push(
    `<text x="${cx.toFixed(2)}" y="${cy.toFixed(2)}" text-anchor="middle" dominant-baseline="central" font-family="Helvetica" font-size="9" fill="#666">${xmlEncode(label)}</text>`
  );
}

/** Recursive dispatch helper used by container variants (SDT, TextBox, Shape). */
function renderPageContentList(
  items: readonly PageContent[],
  geometry: PageGeometry,
  elements: string[]
): void {
  for (const item of items) {
    switch (item.type) {
      case "paragraph":
        renderLayoutParagraphToSvg(item, geometry, elements);
        break;
      case "table":
        renderLayoutTableToSvg(item, geometry, elements);
        break;
      case "image":
        renderLayoutImageToSvg(item, geometry, elements);
        break;
      case "float":
        renderLayoutFloatToSvg(item, geometry, elements);
        break;
      case "textBox":
        renderLayoutTextBoxToSvg(item, geometry, elements);
        break;
      case "shape":
        renderLayoutShapeToSvg(item, geometry, elements);
        break;
      case "chart":
        renderLayoutChartToSvg(item, geometry, elements);
        break;
      case "sdt":
        renderLayoutSdtToSvg(item, geometry, elements);
        break;
      case "math":
        renderLayoutMathToSvg(item, geometry, elements);
        break;
      case "checkBox":
        renderLayoutCheckBoxToSvg(item, geometry, elements);
        break;
      case "tableOfContents":
        renderLayoutTocToSvg(item, geometry, elements);
        break;
      case "altChunk":
        renderLayoutPlaceholderToSvg(item.rect, geometry, `[${item.contentType}]`, elements);
        break;
      case "opaqueDrawing":
        renderLayoutPlaceholderToSvg(item.rect, geometry, "[drawing]", elements);
        break;
      default: {
        const _exhaustive: never = item;
        throw new Error(
          `renderPageContentList: unhandled PageContent ${(_exhaustive as { type: string }).type}`
        );
      }
    }
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  // Same approach as core/internal-utils.ts to stay browser-friendly.
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  if (typeof globalThis.btoa === "function") {
    return globalThis.btoa(binary);
  }
  // Node fallback
  const buf = (
    globalThis as {
      Buffer?: { from(data: string, enc: string): { toString(enc: string): string } };
    }
  ).Buffer;
  if (buf) {
    return buf.from(binary, "binary").toString("base64");
  }
  throw new Error("btoa / Buffer unavailable; cannot encode image data");
}
