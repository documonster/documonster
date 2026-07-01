import type { Border, BorderDiagonal, Borders, Color, Style } from "@excel/types";
import { copyStyle } from "@excel/utils/copy-style";

type BorderEdge = Partial<Border> | undefined;

/**
 * Borders collected from the perimeter of a merge range.
 * Stored as flat arrays indexed by row/column offset for O(1) lookup.
 */
export interface CollectedBorders {
  /** top edge border for each column (index = col - left) */
  topEdges: BorderEdge[];
  /** bottom edge border for each column (index = col - left) */
  bottomEdges: BorderEdge[];
  /** left edge border for each row (index = row - top) */
  leftEdges: BorderEdge[];
  /** right edge border for each row (index = row - top) */
  rightEdges: BorderEdge[];
  diagonal?: Partial<BorderDiagonal>;
  color?: Partial<Color>;
}

/**
 * Collect perimeter borders from cells before a merge is applied.
 * Must be called BEFORE cell.merge() overwrites slave styles.
 *
 * Only iterates the four edges of the range, not the full rectangle.
 * For perimeter edges where the cell has no border, falls back to the master's border.
 */
export function collectMergeBorders(
  top: number,
  left: number,
  bottom: number,
  right: number,
  findCell: (r: number, c: number) => { style: Partial<Style> } | undefined
): CollectedBorders | undefined {
  const masterBorder: (Partial<Borders> & { color?: Partial<Color> }) | undefined = findCell(
    top,
    left
  )?.style?.border;

  const width = right - left + 1;
  const height = bottom - top + 1;
  const topEdges: BorderEdge[] = new Array(width);
  const bottomEdges: BorderEdge[] = new Array(width);
  const leftEdges: BorderEdge[] = new Array(height);
  const rightEdges: BorderEdge[] = new Array(height);

  let hasAny = false;

  // Top & bottom rows
  for (let j = left; j <= right; j++) {
    const idx = j - left;
    const topBorder = findCell(top, j)?.style?.border;
    topEdges[idx] = topBorder?.top || masterBorder?.top;

    if (bottom !== top) {
      const botBorder = findCell(bottom, j)?.style?.border;
      bottomEdges[idx] = botBorder?.bottom || masterBorder?.bottom;
    } else {
      bottomEdges[idx] = topBorder?.bottom || masterBorder?.bottom;
    }

    if (topEdges[idx] || bottomEdges[idx]) {
      hasAny = true;
    }
  }

  // Left & right columns
  for (let i = top; i <= bottom; i++) {
    const idx = i - top;
    const leftBorder = findCell(i, left)?.style?.border;
    leftEdges[idx] = leftBorder?.left || masterBorder?.left;

    if (right !== left) {
      const rightBorder = findCell(i, right)?.style?.border;
      rightEdges[idx] = rightBorder?.right || masterBorder?.right;
    } else {
      rightEdges[idx] = leftBorder?.right || masterBorder?.right;
    }

    if (leftEdges[idx] || rightEdges[idx]) {
      hasAny = true;
    }
  }

  const diagonal = masterBorder?.diagonal;
  const color = masterBorder?.color;

  if (!hasAny && !diagonal) {
    return undefined;
  }

  return { topEdges, bottomEdges, leftEdges, rightEdges, diagonal, color };
}

/**
 * Apply position-aware borders to a merged cell range.
 * Must be called AFTER cell.merge() so that the master style is available.
 *
 * Each cell receives a deep-copied style from the master so that
 * later mutations to one cell do not leak to others.
 */
export function applyMergeBorders(
  top: number,
  left: number,
  bottom: number,
  right: number,
  collected: CollectedBorders,
  getCell: (r: number, c: number) => { style: Partial<Style> }
): void {
  const { topEdges, bottomEdges, leftEdges, rightEdges, diagonal, color } = collected;
  const masterStyle = getCell(top, left).style;

  for (let i = top; i <= bottom; i++) {
    for (let j = left; j <= right; j++) {
      const cell = getCell(i, j);
      const style: Partial<Style> = copyStyle(masterStyle) || {};

      // `Borders` has per-edge color, but the merge logic also carries a
      // single perimeter color applied alongside the chosen edges.
      const newBorder: Partial<Borders> & { color?: Partial<Color> } = {};
      let hasBorder = false;

      if (i === top && topEdges[j - left]) {
        newBorder.top = topEdges[j - left];
        hasBorder = true;
      }
      if (i === bottom && bottomEdges[j - left]) {
        newBorder.bottom = bottomEdges[j - left];
        hasBorder = true;
      }
      if (j === left && leftEdges[i - top]) {
        newBorder.left = leftEdges[i - top];
        hasBorder = true;
      }
      if (j === right && rightEdges[i - top]) {
        newBorder.right = rightEdges[i - top];
        hasBorder = true;
      }
      if (diagonal) {
        newBorder.diagonal = diagonal;
        hasBorder = true;
      }

      if (hasBorder) {
        if (color) {
          // The perimeter color applies to whichever edges were set above.
          newBorder.color = color;
        }
        style.border = newBorder;
      } else {
        delete style.border;
      }

      cell.style = style;
    }
  }
}
