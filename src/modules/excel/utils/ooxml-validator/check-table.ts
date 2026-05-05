/**
 * Table part check.
 *
 * Three real-world "Removed Records: Table from /xl/tables/tableN.xml"
 * patterns that cause Excel to drop the entire table on open:
 *
 *   1. **Redundant per-column `<filterColumn>`.** The writer used to
 *      emit a bare `<filterColumn colId="N" hiddenButton="1"/>` for
 *      every column in the table. The schema accepts it, but Excel's
 *      loader interprets a fully-hidden autoFilter on a live table as
 *      inconsistent and drops the table to recover.
 *
 *   2. **`<totalsRowFormula>` paired with a built-in
 *      `totalsRowFunction`.** The schema allows the child only when
 *      the function is `"custom"` (or absent). For any of `sum`,
 *      `average`, `count`, `countNums`, `max`, `min`, `stdDev`,
 *      `var` Excel generates the SUBTOTAL formula itself and rejects
 *      the table if we pre-emit one.
 *
 *   3. **autoFilter range covering the totals row.** When
 *      `totalsRowCount="1"` is set on the `<table>`, the autoFilter
 *      range must stop one row above the totals row. Filtering the
 *      totals row itself is semantically invalid and Excel rejects
 *      the table.
 *
 * These are some of the most common "Excel cannot open" symptoms in
 * production xlsx from other tooling too — catching them up front is a
 * high-value check.
 */

import type { XmlElement } from "@xml/types";

import type { ValidationContext } from "./context";
import { attrByLocalName, findChildLocal, findChildrenLocal } from "./xml-utils";

const TABLE_PATH_RE = /^xl\/tables\/table\d+\.xml$/;

/** Built-in totals-row functions that Excel synthesises its own formula for. */
const BUILTIN_TOTALS_FUNCTIONS = new Set([
  "sum",
  "average",
  "count",
  "countNums",
  "max",
  "min",
  "stdDev",
  "var"
]);

export function checkTables(ctx: ValidationContext): void {
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (entry.type === "directory" || !TABLE_PATH_RE.test(path)) {
      continue;
    }
    checkSingleTable(ctx, path);
  }
}

function checkSingleTable(ctx: ValidationContext, path: string): void {
  const dom = ctx.readDom(path);
  if (!dom) {
    return;
  }
  const root = dom.root;

  checkRedundantFilterColumns(ctx, path, root);
  checkTotalsRowFormula(ctx, path, root);
  checkAutoFilterRange(ctx, path, root);
}

// -----------------------------------------------------------------------------
// 1. Redundant per-column <filterColumn>
// -----------------------------------------------------------------------------

function checkRedundantFilterColumns(ctx: ValidationContext, path: string, root: XmlElement): void {
  const autoFilter = findChildLocal(root, "autoFilter");
  if (!autoFilter) {
    return;
  }
  const tableColumnsEl = findChildLocal(root, "tableColumns");
  if (!tableColumnsEl) {
    return;
  }
  const tableColumnCount = findChildrenLocal(tableColumnsEl, "tableColumn").length;
  if (tableColumnCount === 0) {
    return;
  }
  const filterColumns = findChildrenLocal(autoFilter, "filterColumn");
  if (filterColumns.length !== tableColumnCount) {
    return; // only the "one per column" writer bug pattern is flagged
  }

  // All filterColumn entries must be bare `hiddenButton="1"` with no
  // filter-state children. Authored filterColumns that legitimately
  // hide the button for a single column will have fewer entries (not
  // one per column) so the guard above already lets them through.
  const allEmpty = filterColumns.every(fc => {
    const hidden = attrByLocalName(fc, "hiddenButton");
    const hasChildren = fc.children.some(c => c.type === "element");
    return hidden === "1" && !hasChildren;
  });
  if (allEmpty) {
    ctx.reporter.error(
      "table-filterColumn-redundant-per-column",
      `${path}: every column has a bare <filterColumn hiddenButton="1"/> ` +
        `(total ${filterColumns.length}). Excel drops tables with a ` +
        `fully-hidden autoFilter on load. Only emit <filterColumn> for ` +
        `columns that have active filter state or an explicit ` +
        `filterButton=false.`,
      path
    );
  }
}

// -----------------------------------------------------------------------------
// 2. totalsRowFormula paired with built-in totalsRowFunction
// -----------------------------------------------------------------------------

function checkTotalsRowFormula(ctx: ValidationContext, path: string, root: XmlElement): void {
  const tableColumnsEl = findChildLocal(root, "tableColumns");
  if (!tableColumnsEl) {
    return;
  }
  for (const col of findChildrenLocal(tableColumnsEl, "tableColumn")) {
    if (ctx.reporter.capped) {
      return;
    }
    const totalsFn = attrByLocalName(col, "totalsRowFunction");
    if (!totalsFn || !BUILTIN_TOTALS_FUNCTIONS.has(totalsFn)) {
      continue;
    }
    const formulaEl = findChildLocal(col, "totalsRowFormula");
    if (formulaEl) {
      const name = attrByLocalName(col, "name") ?? attrByLocalName(col, "id") ?? "?";
      ctx.reporter.error(
        "table-totalsRowFormula-with-builtin-function",
        `${path}: tableColumn "${name}" has totalsRowFunction="${totalsFn}" ` +
          `alongside a <totalsRowFormula>. The schema accepts the formula ` +
          `child only when totalsRowFunction is "custom" or absent.`,
        path
      );
    }
  }
}

// -----------------------------------------------------------------------------
// 3. autoFilter range covering the totals row
// -----------------------------------------------------------------------------

/**
 * Parse an A1-style range ref into `{ top, bottom }` (1-based, inclusive)
 * for the row dimension only. Returns `undefined` when the ref is not a
 * legal rectangle.
 */
function parseRangeRowRange(ref: string): { top: number; bottom: number } | undefined {
  // Match "A1:C10" or "A1:A1" (degenerate).
  const m = /^[A-Za-z]+(\d+)(?::[A-Za-z]+(\d+))?$/.exec(ref);
  if (!m) {
    return undefined;
  }
  const a = parseInt(m[1], 10);
  const b = m[2] !== undefined ? parseInt(m[2], 10) : a;
  if (!Number.isFinite(a) || !Number.isFinite(b)) {
    return undefined;
  }
  return { top: Math.min(a, b), bottom: Math.max(a, b) };
}

function checkAutoFilterRange(ctx: ValidationContext, path: string, root: XmlElement): void {
  const tableRef = attrByLocalName(root, "ref");
  const totalsRowCountAttr = attrByLocalName(root, "totalsRowCount");
  if (!tableRef || !totalsRowCountAttr) {
    return;
  }
  const totalsRowCount = parseInt(totalsRowCountAttr, 10);
  if (!Number.isFinite(totalsRowCount) || totalsRowCount <= 0) {
    return;
  }
  const autoFilterEl = findChildLocal(root, "autoFilter");
  if (!autoFilterEl) {
    return;
  }
  const autoRef = attrByLocalName(autoFilterEl, "ref");
  if (!autoRef) {
    return;
  }
  const tableRange = parseRangeRowRange(tableRef);
  const autoRange = parseRangeRowRange(autoRef);
  if (!tableRange || !autoRange) {
    return;
  }

  const tableBottom = tableRange.bottom;
  const expectedAutoBottom = tableBottom - totalsRowCount;
  if (autoRange.bottom !== expectedAutoBottom) {
    ctx.reporter.error(
      "table-autoFilter-covers-totalsRow",
      `${path}: autoFilter ref ends at row ${autoRange.bottom} but table ref ends ` +
        `at row ${tableBottom} with totalsRowCount=${totalsRowCount}; ` +
        `autoFilter must end at row ${expectedAutoBottom} (one row above the totals row).`,
      path
    );
  }
}
