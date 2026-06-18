/**
 * Table extraction from PDF pages using text fragment positioning.
 *
 * Detects tabular structures by analyzing the spatial layout of text fragments.
 * Since PDF content streams typically render tables as positioned text (with or
 * without drawn grid lines), this module uses a text-only heuristic:
 *
 * 1. Group fragments into lines by Y proximity
 * 2. Detect column boundaries from consistent X-position clusters
 * 3. Identify contiguous blocks of multi-column lines as tables
 * 4. Map fragments to cells based on column/line membership
 *
 * @see content-interpreter.ts for TextFragment extraction
 * @see text-reconstruction.ts for line grouping logic
 */

import type { TextFragment } from "@pdf/reader/content-interpreter";

// =============================================================================
// Public Types
// =============================================================================

/**
 * A single cell in a PDF table.
 */
export interface PdfTableCell {
  /** Text content of the cell */
  text: string;
  /** X position in page coordinates (points) */
  x: number;
  /** Y position in page coordinates (points) */
  y: number;
  /** Width of the cell in points */
  width: number;
  /** Height of the cell in points */
  height: number;
  /** Number of rows this cell spans (default 1) */
  rowSpan?: number;
  /** Number of columns this cell spans (default 1) */
  colSpan?: number;
}

/**
 * A single row in a PDF table.
 */
export interface PdfTableRow {
  /** Cells in this row, ordered left-to-right */
  cells: PdfTableCell[];
}

/**
 * A table extracted from a PDF page.
 */
export interface PdfTable {
  /** Rows in this table, ordered top-to-bottom */
  rows: PdfTableRow[];
  /** X position of the table (left edge) in page coordinates */
  x: number;
  /** Y position of the table (top edge) in page coordinates */
  y: number;
  /** Width of the table in points */
  width: number;
  /** Height of the table in points */
  height: number;
}

// =============================================================================
// Internal Types
// =============================================================================

/** A line of text fragments grouped by Y proximity */
interface FragmentLine {
  /** Representative Y position for this line */
  y: number;
  /** Fragments on this line, sorted left-to-right by X */
  fragments: TextFragment[];
  /** Font size of the first fragment (used for spacing thresholds) */
  fontSize: number;
}

/** A detected column boundary */
interface ColumnBoundary {
  /** Left edge of the column */
  left: number;
  /** Right edge of the column */
  right: number;
}

// =============================================================================
// Constants
// =============================================================================

/**
 * Minimum number of columns required to consider a block of lines as a table.
 */
const MIN_TABLE_COLUMNS = 2;

/**
 * Minimum number of consecutive multi-column lines to form a table.
 */
const MIN_TABLE_ROWS = 2;

// =============================================================================
// Public API
// =============================================================================

/**
 * Extract tables from a page's text fragments.
 *
 * Uses text positioning heuristics to detect tabular structures without
 * relying on drawn lines or grid paths.
 *
 * @param fragments - Text fragments from `extractTextFromPage`
 * @param pageWidth - Page width in points
 * @param pageHeight - Page height in points
 * @returns Array of detected tables
 */
export function extractTables(
  fragments: TextFragment[],
  pageWidth: number,
  pageHeight: number
): PdfTable[] {
  if (fragments.length === 0) {
    return [];
  }

  // Filter to horizontal text only (vertical CJK tables are not handled here)
  const horizontal = fragments.filter(f => !f.isVertical && f.text.trim().length > 0);
  if (horizontal.length < MIN_TABLE_COLUMNS * MIN_TABLE_ROWS) {
    return [];
  }

  // Step 1: Group fragments into lines by Y proximity
  const lines = groupFragmentsIntoLines(horizontal);
  if (lines.length < MIN_TABLE_ROWS) {
    return [];
  }

  // Step 2: Detect column boundaries across lines
  const columns = detectColumnBoundaries(lines, pageWidth);
  if (columns.length < MIN_TABLE_COLUMNS) {
    return [];
  }

  // Step 3: Identify contiguous runs of lines that form tables
  const tableRanges = findTableRanges(lines, columns);

  // Step 4: Build table structures
  const tables: PdfTable[] = [];
  for (const range of tableRanges) {
    const table = buildTable(lines, columns, range.start, range.end, pageHeight);
    if (table) {
      tables.push(table);
    }
  }

  return tables;
}

// =============================================================================
// Step 1: Group Fragments into Lines
// =============================================================================

/**
 * Group text fragments into horizontal lines based on Y proximity.
 * Returns lines sorted top-to-bottom (descending Y in PDF coordinates).
 */
function groupFragmentsIntoLines(fragments: TextFragment[]): FragmentLine[] {
  // Sort fragments top-to-bottom, then left-to-right
  const sorted = [...fragments].sort((a, b) => {
    const dy = b.y - a.y;
    if (Math.abs(dy) > 1) {
      return dy;
    }
    return a.x - b.x;
  });

  const lines: FragmentLine[] = [];
  let currentFragments: TextFragment[] = [sorted[0]];
  let currentY = sorted[0].y;

  for (let i = 1; i < sorted.length; i++) {
    const f = sorted[i];
    const avgFontSize = (currentFragments[0].fontSize + f.fontSize) / 2;
    const threshold = Math.max(avgFontSize * 0.4, 2);

    if (Math.abs(f.y - currentY) <= threshold) {
      currentFragments.push(f);
    } else {
      // Finalize previous line
      currentFragments.sort((a, b) => a.x - b.x);
      lines.push({
        y: currentY,
        fragments: currentFragments,
        fontSize: currentFragments[0].fontSize
      });
      currentFragments = [f];
      currentY = f.y;
    }
  }

  // Finalize last line
  if (currentFragments.length > 0) {
    currentFragments.sort((a, b) => a.x - b.x);
    lines.push({
      y: currentY,
      fragments: currentFragments,
      fontSize: currentFragments[0].fontSize
    });
  }

  return lines;
}

// =============================================================================
// Step 2: Detect Column Boundaries
// =============================================================================

/**
 * Detect column boundaries by analyzing fragment positions across lines.
 *
 * Algorithm:
 * 1. For each line, compute gaps between consecutive fragments
 * 2. Find vertical divider positions that consistently fall in gaps across lines
 * 3. When gaps don't perfectly align (e.g. right-aligned numbers), use
 *    fragment start positions to infer column boundaries
 *
 * This approach handles mixed alignment (left-aligned text headers with
 * right-aligned numeric data) by looking at both gaps and start positions.
 */
function detectColumnBoundaries(lines: FragmentLine[], _pageWidth: number): ColumnBoundary[] {
  // Strategy: for each pair of lines, find the set of divider positions
  // that would produce the same column count. Then pick the most common
  // column count and find divider positions that work best.

  // Step A: Determine the most common fragment count per line
  const fragCounts = lines.map(l => l.fragments.length);
  const countFreq = new Map<number, number>();
  for (const c of fragCounts) {
    if (c >= MIN_TABLE_COLUMNS) {
      countFreq.set(c, (countFreq.get(c) ?? 0) + 1);
    }
  }

  if (countFreq.size === 0) {
    return [];
  }

  // Find the most common fragment count (the "expected" number of columns)
  let bestCount = 0;
  let bestFreq = 0;
  for (const [count, freq] of countFreq) {
    if (freq > bestFreq || (freq === bestFreq && count > bestCount)) {
      bestCount = count;
      bestFreq = freq;
    }
  }

  if (bestCount < MIN_TABLE_COLUMNS || bestFreq < MIN_TABLE_ROWS) {
    return [];
  }

  // Step B: From lines with the expected fragment count, extract divider positions.
  // For each such line, dividers are placed between consecutive fragments.
  // Divider position = midpoint between fragment[i].rightEdge and fragment[i+1].x
  // (or just the gap midpoint if there's a real gap; if they overlap, use the start
  // of the next fragment).
  const linesWithExpectedCount = lines.filter(l => l.fragments.length === bestCount);

  // Collect divider positions for each gap index (0..bestCount-2)
  // For each gap between column i and column i+1, we need a divider that:
  // - Is to the right of all fragment[i] right-edges (across all lines)
  // - Is to the left of all fragment[i+1] left-edges (across all lines)
  // We compute the max right-edge of fragment[i] and min left-edge of fragment[i+1]
  // across all matching lines, then place the divider at the midpoint.
  const maxRightByIndex: number[] = Array.from({ length: bestCount - 1 }, () => -Infinity);
  const minLeftByIndex: number[] = Array.from({ length: bestCount - 1 }, () => Infinity);

  for (const line of linesWithExpectedCount) {
    const frags = line.fragments;
    for (let i = 0; i + 1 < frags.length; i++) {
      const rightEdge = frags[i].x + frags[i].width;
      const nextStart = frags[i + 1].x;
      maxRightByIndex[i] = Math.max(maxRightByIndex[i], rightEdge);
      minLeftByIndex[i] = Math.min(minLeftByIndex[i], nextStart);
    }
  }

  // Place each divider between the max right of column i and min left of column i+1
  const medianDividers: number[] = [];
  for (let i = 0; i < bestCount - 1; i++) {
    const maxRight = maxRightByIndex[i];
    const minLeft = minLeftByIndex[i];
    if (minLeft > maxRight) {
      // Clean gap — place divider at midpoint
      medianDividers.push((maxRight + minLeft) / 2);
    } else {
      // Overlap — place divider at the left-edge of the next column's fragment
      // (this handles right-aligned numbers that extend into the next column's space)
      medianDividers.push(minLeft);
    }
  }

  if (medianDividers.length < 1) {
    return [];
  }

  // Step C: Build column boundaries from dividers
  let globalLeft = Infinity;
  let globalRight = -Infinity;
  for (const line of lines) {
    for (const f of line.fragments) {
      globalLeft = Math.min(globalLeft, f.x);
      globalRight = Math.max(globalRight, f.x + f.width);
    }
  }

  const columns: ColumnBoundary[] = [];
  let prevRight = globalLeft;
  for (const divider of medianDividers) {
    columns.push({ left: prevRight, right: divider });
    prevRight = divider;
  }
  columns.push({ left: prevRight, right: globalRight });

  return columns.length >= MIN_TABLE_COLUMNS ? columns : [];
}

// =============================================================================
// Step 3: Find Contiguous Table Ranges
// =============================================================================

/**
 * Identify contiguous runs of lines where most columns have content.
 * Returns ranges of line indices that form table blocks.
 */
function findTableRanges(
  lines: FragmentLine[],
  columns: ColumnBoundary[]
): Array<{ start: number; end: number }> {
  // For each line, count how many columns contain at least one fragment
  const lineColumnCounts: number[] = [];
  for (const line of lines) {
    const occupiedColumns = new Set<number>();
    for (const f of line.fragments) {
      const colIdx = findColumnIndex(f.x, columns);
      if (colIdx >= 0) {
        occupiedColumns.add(colIdx);
      }
    }
    lineColumnCounts.push(occupiedColumns.size);
  }

  // A line is "tabular" if it has fragments in at least 2 columns
  const ranges: Array<{ start: number; end: number }> = [];
  let rangeStart = -1;

  for (let i = 0; i < lineColumnCounts.length; i++) {
    const isTabular = lineColumnCounts[i] >= MIN_TABLE_COLUMNS;
    if (isTabular && rangeStart === -1) {
      rangeStart = i;
    } else if (!isTabular && rangeStart !== -1) {
      if (i - rangeStart >= MIN_TABLE_ROWS) {
        ranges.push({ start: rangeStart, end: i - 1 });
      }
      rangeStart = -1;
    }
  }

  // Close any open range
  if (rangeStart !== -1 && lines.length - rangeStart >= MIN_TABLE_ROWS) {
    ranges.push({ start: rangeStart, end: lines.length - 1 });
  }

  return ranges;
}

/**
 * Find which column a given X position belongs to.
 * Returns -1 if the position doesn't fall within any column.
 */
function findColumnIndex(x: number, columns: ColumnBoundary[]): number {
  for (let i = columns.length - 1; i >= 0; i--) {
    if (x >= columns[i].left - 1) {
      return i;
    }
  }
  return -1;
}

// =============================================================================
// Step 4: Build Table Structure
// =============================================================================

/**
 * Build a PdfTable from a range of lines and column boundaries.
 */
function buildTable(
  lines: FragmentLine[],
  columns: ColumnBoundary[],
  startLine: number,
  endLine: number,
  _pageHeight: number
): PdfTable | null {
  const rows: PdfTableRow[] = [];

  for (let li = startLine; li <= endLine; li++) {
    const line = lines[li];
    const row = buildRow(line, columns, lines, li, startLine, endLine);
    rows.push(row);
  }

  if (rows.length === 0) {
    return null;
  }

  // Calculate table bounding box
  const tableLines = lines.slice(startLine, endLine + 1);
  const topY = tableLines[0].y;
  const bottomY = tableLines[tableLines.length - 1].y;
  const bottomFontSize = tableLines[tableLines.length - 1].fontSize;

  const allX: number[] = [];
  const allRightEdges: number[] = [];
  for (const line of tableLines) {
    for (const f of line.fragments) {
      allX.push(f.x);
      allRightEdges.push(f.x + f.width);
    }
  }

  const tableX = allX.length > 0 ? Math.min(...allX) : 0;
  const tableRight = allRightEdges.length > 0 ? Math.max(...allRightEdges) : 0;
  const tableWidth = tableRight - tableX;

  // Height: from top of first line to bottom of last line (including font height)
  const tableHeight = topY - bottomY + bottomFontSize;

  return {
    rows,
    x: tableX,
    y: topY,
    width: tableWidth,
    height: tableHeight
  };
}

/**
 * Build a single table row by mapping fragments to columns.
 */
function buildRow(
  line: FragmentLine,
  columns: ColumnBoundary[],
  _allLines: FragmentLine[],
  _lineIdx: number,
  _startLine: number,
  _endLine: number
): PdfTableRow {
  // Group fragments by column
  const columnFragments: Map<number, TextFragment[]> = new Map();
  for (const f of line.fragments) {
    const colIdx = findColumnIndex(f.x, columns);
    if (colIdx >= 0) {
      const existing = columnFragments.get(colIdx) ?? [];
      existing.push(f);
      columnFragments.set(colIdx, existing);
    }
  }

  // Build cells for each column
  const cells: PdfTableCell[] = [];
  for (let colIdx = 0; colIdx < columns.length; colIdx++) {
    const col = columns[colIdx];
    const frags = columnFragments.get(colIdx);

    if (frags && frags.length > 0) {
      // Concatenate text from all fragments in this cell
      const text = frags.map(f => f.text).join(" ");
      const cellX = frags[0].x;
      const cellY = line.y;
      const lastFrag = frags[frags.length - 1];
      const cellRight = lastFrag.x + lastFrag.width;
      const cellWidth = cellRight - cellX;
      const cellHeight = line.fontSize;

      cells.push({
        text: text.trim(),
        x: cellX,
        y: cellY,
        width: cellWidth,
        height: cellHeight
      });
    } else {
      // Empty cell
      cells.push({
        text: "",
        x: col.left,
        y: line.y,
        width: col.right - col.left,
        height: line.fontSize
      });
    }
  }

  return { cells };
}
