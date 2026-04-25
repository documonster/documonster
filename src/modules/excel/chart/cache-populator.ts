/**
 * Chart cache populator.
 *
 * Walks a ChartModel, finds every `numRef`/`strRef` formula, resolves it against
 * the actual worksheet cell values in the workbook, and populates the cache
 * points so that headless consumers (PDF export, image preview, other readers
 * that don't recalculate formulas) see non-empty charts.
 *
 * Excel itself recomputes caches on open, so this is a best-effort enrichment:
 *   - empty/missing cells produce null values at the correct index
 *   - formula-bearing cells use their cached result when available
 *   - non-existent sheets or malformed formulas are skipped silently
 */
import type {
  AxisDataSource,
  ChartModel,
  ChartTypeGroup,
  NumberCache,
  NumberReference,
  SeriesBase,
  StringCache,
  StringReference
} from "@excel/chart/types";
import { colCache } from "@excel/utils/col-cache";
import type { Workbook } from "@excel/workbook";
import type { Worksheet } from "@excel/worksheet";

/**
 * Populate all number/string caches in a ChartModel from the given workbook.
 * Mutates the model in place. Safe to call multiple times (idempotent).
 *
 * @param model - The chart model to enrich
 * @param workbook - The workbook used to resolve sheet/cell references
 */
export function fillChartCaches(model: ChartModel, workbook: Workbook): void {
  const plotArea = model.chart?.plotArea;
  if (!plotArea) {
    return;
  }
  const date1904 = workbook.properties?.date1904;
  for (const group of plotArea.chartTypes) {
    fillGroupCaches(group, workbook, date1904);
  }
}

function fillGroupCaches(group: ChartTypeGroup, workbook: Workbook, date1904?: boolean): void {
  const series = (group as { series?: SeriesBase[] }).series;
  if (!series) {
    return;
  }
  for (const s of series) {
    fillSeriesCaches(s, workbook, date1904);
  }
}

function fillSeriesCaches(series: SeriesBase, workbook: Workbook, date1904?: boolean): void {
  // Series name (tx): may be strRef
  const tx = (series as { tx?: { strRef?: StringReference } }).tx;
  if (tx?.strRef) {
    fillStrRef(tx.strRef, workbook);
  }

  // Category / X axis data source
  const cat = (series as { cat?: AxisDataSource }).cat;
  if (cat) {
    fillAxisDataSource(cat, workbook, date1904);
  }
  const xVal = (series as { xVal?: AxisDataSource }).xVal;
  if (xVal) {
    fillAxisDataSource(xVal, workbook, date1904);
  }

  // Numeric values — val, yVal, bubbleSize
  const val = (series as { val?: { numRef?: NumberReference } }).val;
  if (val?.numRef) {
    fillNumRef(val.numRef, workbook, date1904);
  }
  const yVal = (series as { yVal?: { numRef?: NumberReference } }).yVal;
  if (yVal?.numRef) {
    fillNumRef(yVal.numRef, workbook, date1904);
  }
  const bubbleSize = (series as { bubbleSize?: { numRef?: NumberReference } }).bubbleSize;
  if (bubbleSize?.numRef) {
    fillNumRef(bubbleSize.numRef, workbook, date1904);
  }

  // Error bars may have custom plus/minus references
  const errorBars = (series as { errorBars?: unknown }).errorBars;
  if (errorBars) {
    const list = Array.isArray(errorBars) ? errorBars : [errorBars];
    for (const eb of list as Array<{
      plus?: { numRef?: NumberReference };
      minus?: { numRef?: NumberReference };
    }>) {
      if (eb.plus?.numRef) {
        fillNumRef(eb.plus.numRef, workbook, date1904);
      }
      if (eb.minus?.numRef) {
        fillNumRef(eb.minus.numRef, workbook, date1904);
      }
    }
  }
}

function fillAxisDataSource(src: AxisDataSource, workbook: Workbook, date1904?: boolean): void {
  if (src.strRef) {
    fillStrRef(src.strRef, workbook);
  }
  if (src.numRef) {
    fillNumRef(src.numRef, workbook, date1904);
  }
  // multiLvlStrRef — rare, skip
}

/**
 * Populate a NumberReference cache from the workbook.
 * Only fills if `cache.points` is currently empty.
 */
export function fillNumRef(ref: NumberReference, workbook: Workbook, date1904?: boolean): void {
  if (!ref.formula) {
    return;
  }
  if (ref.cache?.points && ref.cache.points.length > 0) {
    return; // already populated
  }
  const resolved = resolveReference(ref.formula, workbook);
  if (!resolved) {
    return;
  }
  const points: NumberCache["points"] = [];
  let idx = 0;
  for (const cell of resolved.cells) {
    const n = toNumber(cell.value, date1904);
    if (n !== undefined) {
      points.push({ index: idx, value: n });
    }
    idx++;
  }
  if (!ref.cache) {
    ref.cache = { points: [] };
  }
  ref.cache.points = points;
  ref.cache.pointCount = idx; // total slots (including empty)
}

/**
 * Populate a StringReference cache from the workbook.
 * Only fills if `cache.points` is currently empty.
 */
export function fillStrRef(ref: StringReference, workbook: Workbook): void {
  if (!ref.formula) {
    return;
  }
  if (ref.cache?.points && ref.cache.points.length > 0) {
    return;
  }
  const resolved = resolveReference(ref.formula, workbook);
  if (!resolved) {
    return;
  }
  const points: StringCache["points"] = [];
  let idx = 0;
  for (const cell of resolved.cells) {
    const s = toString(cell.value);
    if (s !== undefined && s !== "") {
      points.push({ index: idx, value: s });
    }
    idx++;
  }
  if (!ref.cache) {
    ref.cache = { points: [] };
  }
  ref.cache.points = points;
  ref.cache.pointCount = idx;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface ResolvedReference {
  worksheet: Worksheet;
  cells: Array<{ value: unknown }>;
}

/**
 * Resolve a chart formula like `Sheet1!$A$1:$A$4` or `'My Sheet'!$B$2:$B$5`
 * into a sequence of cell values in row-major order.
 *
 * Multi-range unions (`Sheet1!A1:A3,Sheet1!A5:A7`) are flattened preserving order.
 * Returns undefined if the reference cannot be resolved.
 */
function resolveReference(formula: string, workbook: Workbook): ResolvedReference | undefined {
  // Multi-range support: Excel uses commas or semicolons between sub-refs.
  // Inside quoted sheet names, commas don't count — do a safe split.
  const parts = splitFormulaRanges(formula);
  let worksheet: Worksheet | undefined;
  const cells: Array<{ value: unknown }> = [];

  for (const part of parts) {
    const trimmed = part.trim();
    if (!trimmed) {
      continue;
    }
    let decoded;
    try {
      decoded = colCache.decodeEx(trimmed);
    } catch {
      return undefined;
    }
    const sheetName = decoded.sheetName;
    if (!sheetName) {
      return undefined;
    }
    const ws = workbook.getWorksheet(sheetName);
    if (!ws) {
      return undefined;
    }
    // Track first resolved worksheet for the return value
    if (!worksheet) {
      worksheet = ws;
    }

    if ("top" in decoded && "left" in decoded) {
      // Range: iterate in row-major order matching Excel's cache layout.
      // For vertical ranges (single column), this is top→bottom.
      // For horizontal ranges (single row), this is left→right.
      // For 2D ranges, row-major (left→right, top→bottom).
      const top = decoded.top as number;
      const left = decoded.left as number;
      const bottom = decoded.bottom as number;
      const right = decoded.right as number;
      for (let r = top; r <= bottom; r++) {
        for (let c = left; c <= right; c++) {
          cells.push({ value: extractCellValue(ws, r, c) });
        }
      }
    } else if ("row" in decoded && "col" in decoded) {
      const r = decoded.row as number;
      const c = decoded.col as number;
      cells.push({ value: extractCellValue(ws, r, c) });
    }
  }

  if (!worksheet) {
    return undefined;
  }
  return { worksheet, cells };
}

/**
 * Split a chart formula that may contain multiple ranges (union) separated
 * by commas or semicolons, respecting quoted sheet names.
 */
function splitFormulaRanges(formula: string): string[] {
  const parts: string[] = [];
  let current = "";
  let inQuote = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === "'") {
      // Excel escapes a single quote inside a sheet name as ''
      if (inQuote && formula[i + 1] === "'") {
        current += "''";
        i++;
        continue;
      }
      inQuote = !inQuote;
      current += ch;
    } else if (!inQuote && (ch === "," || ch === ";")) {
      parts.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  if (current) {
    parts.push(current);
  }
  return parts;
}

/**
 * Extract a raw value from a worksheet cell.
 * Avoids forcing creation of empty cells by using the row-level accessor.
 */
function extractCellValue(ws: Worksheet, row: number, col: number): unknown {
  // Worksheet.getCell creates sparse cell slots but does not mutate existing
  // values — safe to call.
  const cell = ws.getCell(row, col);
  const v = cell.value;
  if (v === null || v === undefined) {
    return undefined;
  }
  // Formula values are wrapped: { formula, result }
  if (typeof v === "object" && v !== null && "result" in v) {
    const result = (v as { result?: unknown }).result;
    return result;
  }
  // Rich text
  if (typeof v === "object" && v !== null && "richText" in v) {
    const rt = (v as { richText?: Array<{ text?: string }> }).richText;
    if (rt) {
      return rt.map(r => r.text ?? "").join("");
    }
  }
  // Hyperlink / error
  if (typeof v === "object" && v !== null) {
    if ("error" in v) {
      return undefined;
    }
    if ("text" in v && "hyperlink" in v) {
      return (v as { text?: unknown }).text;
    }
  }
  return v;
}

function toNumber(v: unknown, date1904?: boolean): number | undefined {
  if (v === null || v === undefined) {
    return undefined;
  }
  if (typeof v === "number") {
    return Number.isFinite(v) ? v : undefined;
  }
  if (typeof v === "boolean") {
    return v ? 1 : 0;
  }
  if (v instanceof Date) {
    return dateToSerial(v, date1904);
  }
  if (typeof v === "string") {
    // Try to coerce numeric string
    const n = Number(v);
    if (Number.isFinite(n) && v.trim() !== "") {
      return n;
    }
  }
  return undefined;
}

function toString(v: unknown): string | undefined {
  if (v === null || v === undefined) {
    return undefined;
  }
  if (typeof v === "string") {
    return v;
  }
  if (typeof v === "number" || typeof v === "boolean") {
    return String(v);
  }
  if (v instanceof Date) {
    return v.toISOString();
  }
  return undefined;
}

/**
 * Convert a JS Date to an Excel serial number.
 * When date1904 is true, adjusts by 1462 days (the difference between
 * the 1900 and 1904 date systems).
 */
function dateToSerial(d: Date, date1904?: boolean): number {
  const epoch = Date.UTC(1899, 11, 30);
  let serial = (d.getTime() - epoch) / 86400000;
  if (!date1904 && serial >= 60) {
    serial += 1;
  }
  return date1904 ? serial - 1462 : serial;
}
