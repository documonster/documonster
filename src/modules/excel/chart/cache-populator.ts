import type { ChartExModel } from "@excel/chart/chart-ex-types";
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
  DataLabelsRange,
  NumberCache,
  NumberReference,
  SeriesBase,
  StringCache,
  StringReference,
  MultiLevelStringReference
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
 * @param contextWorksheet - Optional worksheet providing the default scope for
 *   defined-name resolution. When a defined name has a sheet-scoped entry
 *   matching this worksheet's workbook index (`localSheetId`), that entry
 *   wins over the workbook-scoped entry. Supplying this argument is required
 *   to correctly resolve sheet-scoped names; omitting it falls back to
 *   workbook-scoped names only.
 */
export function fillChartCaches(
  model: ChartModel,
  workbook: Workbook,
  contextWorksheet?: Worksheet
): void {
  const plotArea = model.chart?.plotArea;
  if (!plotArea) {
    return;
  }
  const date1904 = workbook.properties?.date1904;
  const ctx = buildResolverContext(workbook, contextWorksheet);
  for (const group of plotArea.chartTypes) {
    fillGroupCaches(group, ctx, date1904);
  }
}

export function fillChartExCaches(
  model: ChartExModel,
  workbook: Workbook,
  contextWorksheet?: Worksheet
): void {
  const ctx = buildResolverContext(workbook, contextWorksheet);
  for (const entry of model.chartSpace.chartData.data) {
    if (entry.strDim?.formula && !hasChartExStringPoints(entry.strDim)) {
      const resolved = resolveReference(entry.strDim.formula, ctx);
      if (resolved) {
        const points: Array<{ index: number; value: string }> = [];
        let idx = 0;
        for (const cell of resolved.cells) {
          const value = toString(cell.value);
          if (value !== undefined && value !== "") {
            points.push({ index: idx, value });
          }
          idx++;
        }
        entry.strDim.levels = [{ ptCount: idx, points }];
      }
    }
    if (entry.numDim?.formula && !hasChartExNumberPoints(entry.numDim)) {
      const resolved = resolveReference(entry.numDim.formula, ctx);
      if (resolved) {
        const points: Array<{ index: number; value: number }> = [];
        let idx = 0;
        for (const cell of resolved.cells) {
          const value = toNumber(cell.value, workbook.properties?.date1904);
          if (value !== undefined) {
            points.push({ index: idx, value });
          }
          idx++;
        }
        entry.numDim.levels = [{ ptCount: idx, points }];
      }
    }
  }
}

function fillGroupCaches(group: ChartTypeGroup, ctx: ResolverContext, date1904?: boolean): void {
  const series = (group as { series?: SeriesBase[] }).series;
  if (!series) {
    return;
  }
  for (const s of series) {
    fillSeriesCaches(s, ctx, date1904);
  }
}

function hasChartExStringPoints(
  dim: NonNullable<ChartExModel["chartSpace"]["chartData"]["data"][number]["strDim"]>
): boolean {
  return dim.levels?.some(level => level.points.length > 0) ?? false;
}

function hasChartExNumberPoints(
  dim: NonNullable<ChartExModel["chartSpace"]["chartData"]["data"][number]["numDim"]>
): boolean {
  return dim.levels?.some(level => level.points.length > 0) ?? false;
}

function fillSeriesCaches(series: SeriesBase, ctx: ResolverContext, date1904?: boolean): void {
  // Series name (tx): may be strRef
  const tx = (series as { tx?: { strRef?: StringReference } }).tx;
  if (tx?.strRef) {
    fillStrRefInternal(tx.strRef, ctx);
  }

  // Category / X axis data source
  const cat = (series as { cat?: AxisDataSource }).cat;
  if (cat) {
    fillAxisDataSource(cat, ctx, date1904);
  }
  const xVal = (series as { xVal?: AxisDataSource }).xVal;
  if (xVal) {
    fillAxisDataSource(xVal, ctx, date1904);
  }

  // Numeric values — val, yVal, bubbleSize
  const val = (series as { val?: { numRef?: NumberReference } }).val;
  if (val?.numRef) {
    fillNumRefInternal(val.numRef, ctx, date1904);
  }
  const yVal = (series as { yVal?: { numRef?: NumberReference } }).yVal;
  if (yVal?.numRef) {
    fillNumRefInternal(yVal.numRef, ctx, date1904);
  }
  const bubbleSize = (series as { bubbleSize?: { numRef?: NumberReference } }).bubbleSize;
  if (bubbleSize?.numRef) {
    fillNumRefInternal(bubbleSize.numRef, ctx, date1904);
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
        fillNumRefInternal(eb.plus.numRef, ctx, date1904);
      }
      if (eb.minus?.numRef) {
        fillNumRefInternal(eb.minus.numRef, ctx, date1904);
      }
    }
  }

  // "Value From Cells" data labels (Excel 2013+) — populate the
  // `c15:datalabelsRange` cache so readers without a formula engine see
  // the right labels and so the writer can emit `<c15:dlblRangeCache>`.
  const dataLabels = (series as { dataLabels?: { dataLabelsRange?: DataLabelsRange } }).dataLabels;
  if (dataLabels?.dataLabelsRange?.formula) {
    fillDataLabelsRange(dataLabels.dataLabelsRange, ctx);
  }
}

function fillDataLabelsRange(range: DataLabelsRange, ctx: ResolverContext): void {
  if (range.cache?.points && range.cache.points.length > 0) {
    return;
  }
  const resolved = resolveReference(range.formula, ctx);
  if (!resolved) {
    return;
  }
  const points: Array<{ index: number; value: string }> = [];
  let idx = 0;
  for (const cell of resolved.cells) {
    const s = toString(cell.value);
    if (s !== undefined && s !== "") {
      points.push({ index: idx, value: s });
    }
    idx++;
  }
  range.cache = { pointCount: idx, points };
}

function fillAxisDataSource(src: AxisDataSource, ctx: ResolverContext, date1904?: boolean): void {
  if (src.strRef) {
    fillStrRefInternal(src.strRef, ctx);
  }
  if (src.numRef) {
    fillNumRefInternal(src.numRef, ctx, date1904);
  }
  if (src.multiLvlStrRef) {
    fillMultiLvlStrRefInternal(src.multiLvlStrRef, ctx);
  }
}

/**
 * Populate a NumberReference cache from the workbook.
 * Only fills if `cache.points` is currently empty.
 *
 * @param contextWorksheet - Optional worksheet whose sheet-scoped defined
 *   names take precedence over workbook-scoped ones.
 */
export function fillNumRef(
  ref: NumberReference,
  workbook: Workbook,
  date1904?: boolean,
  contextWorksheet?: Worksheet
): void {
  fillNumRefInternal(ref, buildResolverContext(workbook, contextWorksheet), date1904);
}

function fillNumRefInternal(ref: NumberReference, ctx: ResolverContext, date1904?: boolean): void {
  if (!ref.formula) {
    return;
  }
  if (ref.cache?.points && ref.cache.points.length > 0) {
    return; // already populated
  }
  const resolved = resolveReference(ref.formula, ctx);
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
 *
 * @param contextWorksheet - Optional worksheet whose sheet-scoped defined
 *   names take precedence over workbook-scoped ones.
 */
export function fillStrRef(
  ref: StringReference,
  workbook: Workbook,
  contextWorksheet?: Worksheet
): void {
  fillStrRefInternal(ref, buildResolverContext(workbook, contextWorksheet));
}

function fillStrRefInternal(ref: StringReference, ctx: ResolverContext): void {
  if (!ref.formula) {
    return;
  }
  if (ref.cache?.points && ref.cache.points.length > 0) {
    return;
  }
  const resolved = resolveReference(ref.formula, ctx);
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

export function fillMultiLvlStrRef(
  ref: MultiLevelStringReference,
  workbook: Workbook,
  contextWorksheet?: Worksheet
): void {
  fillMultiLvlStrRefInternal(ref, buildResolverContext(workbook, contextWorksheet));
}

function fillMultiLvlStrRefInternal(ref: MultiLevelStringReference, ctx: ResolverContext): void {
  if (!ref.formula) {
    return;
  }
  if (ref.cache?.levels?.some(level => level.points.length > 0)) {
    return;
  }
  const resolved = resolveReferenceMatrix(ref.formula, ctx);
  if (!resolved) {
    return;
  }

  const levels: StringCache[] = [];
  for (let col = 0; col < resolved.columnCount; col++) {
    const points: StringCache["points"] = [];
    for (let row = 0; row < resolved.rowCount; row++) {
      const value = toString(resolved.values[row]?.[col]);
      if (value !== undefined && value !== "") {
        points.push({ index: row, value });
      }
    }
    levels.push({ pointCount: resolved.rowCount, points });
  }
  ref.cache = { pointCount: resolved.rowCount, levels };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

interface ResolvedReference {
  worksheet: Worksheet;
  cells: Array<{ value: unknown }>;
}

interface ResolvedReferenceMatrix {
  worksheet: Worksheet;
  values: unknown[][];
  rowCount: number;
  columnCount: number;
}

/**
 * Internal context carried through recursive reference resolution.
 *
 * `localSheetId` is the 0-based workbook position of {@link contextWorksheet}
 * in `workbook.worksheets`, matching the semantics of OOXML's
 * `definedName/@localSheetId`. It is used to give sheet-scoped defined names
 * precedence over workbook-scoped names when both exist with the same bare
 * name.
 *
 * `visitedDefinedNames` tracks the set of defined names currently being
 * expanded in the call stack so that `A -> B -> A` cycles terminate.
 * The set is mutated (added/deleted) rather than copied on each call; this
 * is safe because resolution is single-threaded and each recursive call
 * properly cleans up before returning.
 */
interface ResolverContext {
  workbook: Workbook;
  contextWorksheet?: Worksheet;
  localSheetId?: number;
  visitedDefinedNames: Set<string>;
}

function buildResolverContext(workbook: Workbook, contextWorksheet?: Worksheet): ResolverContext {
  let localSheetId: number | undefined;
  if (contextWorksheet) {
    // workbook.worksheets is ordered/filtered the same way Excel lists sheets,
    // which matches the `localSheetId` assigned at save time by
    // workbook-xform.ts:63-108. Any other indexing (e.g. internal sparse _id)
    // would not match defined names parsed from disk.
    const idx = workbook.worksheets.indexOf(contextWorksheet);
    if (idx >= 0) {
      localSheetId = idx;
    }
  }
  return { workbook, contextWorksheet, localSheetId, visitedDefinedNames: new Set() };
}

/**
 * Resolve a chart formula like `Sheet1!$A$1:$A$4` or `'My Sheet'!$B$2:$B$5`
 * into a sequence of cell values in row-major order.
 *
 * Multi-range unions (`Sheet1!A1:A3,Sheet1!A5:A7`) are flattened preserving order.
 * Returns undefined if the reference cannot be resolved.
 *
 * Resolution strategy (first match wins):
 *   1. Structured table reference (`Table1[Column]`)
 *   2. Defined name (workbook- or sheet-scoped), recursively expanded
 *   3. Direct A1 cell/range references (possibly comma-separated union)
 */
function resolveReference(formula: string, ctx: ResolverContext): ResolvedReference | undefined {
  const structured = resolveStructuredReference(formula, ctx.workbook);
  if (structured) {
    return structured;
  }

  const named = resolveDefinedNameReference(formula, ctx);
  if (named) {
    return named;
  }

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
    const ws = ctx.workbook.getWorksheet(sheetName);
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
 * Regex matching a bare Excel defined name (no sheet prefix, no `$`, no `:`,
 * no `!`, no `,`). Excel's grammar requires the first character to be a
 * letter, underscore, or backslash, and disallows names that look like cell
 * references; we don't try to validate name legality here — the final
 * authority is whether `workbook.definedNames` has a matching entry.
 *
 * Allowing Unicode letters makes this work for CJK defined names, which are
 * common in Chinese Excel files.
 */
const BARE_DEFINED_NAME_RE = /^[A-Za-z_\\\u00A0-\uFFFF][\w.?\\\u00A0-\uFFFF]*$/;

/**
 * Attempt to resolve {@link formula} as a defined name and expand to the
 * underlying A1 ranges. Supports both bare names (`MyRange`) and qualified
 * names (`Sheet1!MyRange`, `'My Sheet'!MyRange`).
 *
 * Name-scope resolution matches Excel semantics:
 *   - Qualified `Sheet!Name` → sheet-scoped entry on `Sheet`, else
 *     workbook-scoped
 *   - Bare `Name` → sheet-scoped entry on the context worksheet (when
 *     provided), else workbook-scoped
 *
 * The result of expanding the name is resolved recursively, so chart
 * formulas that target a named formula (e.g. `OFFSET(...)` stored as a
 * defined name) do not silently fall through — when the name resolves to
 * another reference-like expression, we follow it.
 */
function resolveDefinedNameReference(
  formula: string,
  ctx: ResolverContext
): ResolvedReference | undefined {
  const resolution = findDefinedName(formula, ctx);
  if (!resolution) {
    return undefined;
  }

  // Cycle guard: A -> B -> A must terminate.
  const visitKey = storageKey(resolution.name, resolution.localSheetId);
  if (ctx.visitedDefinedNames.has(visitKey)) {
    return undefined;
  }
  ctx.visitedDefinedNames.add(visitKey);

  try {
    // A defined name may expand to multiple ranges (comma-separated union).
    // We recursively resolve each part through the full reference pipeline
    // so that nested names, structured refs, and A1 refs all work.
    const aggregated: Array<{ value: unknown }> = [];
    let firstWorksheet: Worksheet | undefined;
    for (const rangeStr of resolution.ranges) {
      if (!rangeStr) {
        continue;
      }
      const inner = resolveReference(rangeStr, ctx);
      if (!inner) {
        continue;
      }
      if (!firstWorksheet) {
        firstWorksheet = inner.worksheet;
      }
      for (const cell of inner.cells) {
        aggregated.push(cell);
      }
    }
    if (!firstWorksheet) {
      return undefined;
    }
    return { worksheet: firstWorksheet, cells: aggregated };
  } finally {
    ctx.visitedDefinedNames.delete(visitKey);
  }
}

/**
 * Look up a defined name in the workbook, honouring sheet-scoped vs
 * workbook-scoped precedence. Returns `undefined` when the formula does not
 * look like a defined-name reference at all, or when no entry matches.
 */
function findDefinedName(
  formula: string,
  ctx: ResolverContext
): { name: string; localSheetId?: number; ranges: string[] } | undefined {
  const trimmed = formula.trim();
  if (!trimmed) {
    return undefined;
  }

  // Qualified form: Sheet!Name or 'Sheet Name'!Name
  const qualified = splitQualifiedName(trimmed);
  if (qualified) {
    // Sheet-scoped entry on the qualifying sheet wins; fall back to
    // workbook-scoped if the sheet has no matching entry.
    const qualifyingSheet = ctx.workbook.getWorksheet(qualified.sheetName);
    if (qualifyingSheet) {
      const sheetIdx = ctx.workbook.worksheets.indexOf(qualifyingSheet);
      if (sheetIdx >= 0) {
        const scoped = getDefinedNameRanges(ctx.workbook, qualified.name, sheetIdx);
        if (scoped) {
          return { name: qualified.name, localSheetId: sheetIdx, ranges: scoped };
        }
      }
    }
    const global = getDefinedNameRanges(ctx.workbook, qualified.name, undefined);
    if (global) {
      return { name: qualified.name, ranges: global };
    }
    return undefined;
  }

  // Bare form: must pass the name shape gate to avoid calling into the
  // defined-names API for every failed cell reference.
  if (!BARE_DEFINED_NAME_RE.test(trimmed)) {
    return undefined;
  }
  // Sheet-scoped on the context worksheet wins, then workbook-scoped.
  if (ctx.localSheetId !== undefined) {
    const scoped = getDefinedNameRanges(ctx.workbook, trimmed, ctx.localSheetId);
    if (scoped) {
      return { name: trimmed, localSheetId: ctx.localSheetId, ranges: scoped };
    }
  }
  const global = getDefinedNameRanges(ctx.workbook, trimmed, undefined);
  if (global) {
    return { name: trimmed, ranges: global };
  }
  return undefined;
}

/**
 * Read the ranges (or formula expression) associated with a defined name at
 * a specific scope. Returns `undefined` when no entry exists at that scope.
 */
function getDefinedNameRanges(
  workbook: Workbook,
  name: string,
  localSheetId: number | undefined
): string[] | undefined {
  const definedNames = workbook.definedNames;
  if (!definedNames) {
    return undefined;
  }
  const model = definedNames.getRangesScoped(name, localSheetId);
  if (!model.ranges || model.ranges.length === 0) {
    return undefined;
  }
  // Verify the entry actually exists at the requested scope — getRangesScoped
  // falls back to the bare name when the scoped key is missing, and we must
  // not return workbook-scoped results from a sheet-scoped lookup.
  const matrixMap = definedNames.matrixMap;
  const formulaMap = definedNames.formulaMap;
  const sKey = localSheetId !== undefined ? `${name}\0${localSheetId}` : name;
  if (!matrixMap[sKey] && formulaMap[sKey] === undefined) {
    return undefined;
  }
  return model.ranges;
}

function storageKey(name: string, localSheetId: number | undefined): string {
  return localSheetId !== undefined ? `${name}\0${localSheetId}` : name;
}

/**
 * Split `Sheet!Name` or `'Quoted Sheet'!Name` into its components. Returns
 * `undefined` when the input is not of that shape or the right-hand side
 * contains range punctuation (in which case it is a cell reference, not a
 * defined-name reference).
 */
function splitQualifiedName(formula: string): { sheetName: string; name: string } | undefined {
  let i = 0;
  let sheetName = "";
  if (formula[0] === "'") {
    // Walk a quoted sheet name, treating '' as an escaped single quote.
    i = 1;
    while (i < formula.length) {
      if (formula[i] === "'" && formula[i + 1] === "'") {
        sheetName += "'";
        i += 2;
      } else if (formula[i] === "'") {
        i++;
        break;
      } else {
        sheetName += formula[i];
        i++;
      }
    }
  } else {
    // Unquoted sheet name — ends at the first '!'.
    const bang = formula.indexOf("!");
    if (bang <= 0) {
      return undefined;
    }
    sheetName = formula.slice(0, bang);
    i = bang;
  }
  if (formula[i] !== "!") {
    return undefined;
  }
  const name = formula.slice(i + 1);
  if (!BARE_DEFINED_NAME_RE.test(name)) {
    return undefined;
  }
  return { sheetName, name };
}

function resolveStructuredReference(
  formula: string,
  workbook: Workbook
): ResolvedReference | undefined {
  const parsed = parseStructuredReference(formula.trim());
  if (!parsed) {
    return undefined;
  }
  for (const worksheet of workbook.worksheets) {
    const table = worksheet
      .getTables()
      .find(t => t.name === parsed.tableName || t.displayName === parsed.tableName);
    if (!table) {
      continue;
    }
    const model = table.model;
    const columnIndex = model.columns.findIndex(column => column.name === parsed.columnName);
    if (columnIndex < 0) {
      return undefined;
    }
    const tableRef = colCache.decode(model.tableRef ?? model.ref);
    if (!("top" in tableRef)) {
      return undefined;
    }
    const dataStartRow = tableRef.top + (model.headerRow === false ? 0 : 1);
    const dataEndRow = tableRef.bottom - (model.totalsRow ? 1 : 0);
    const col = tableRef.left + columnIndex;
    const cells: Array<{ value: unknown }> = [];
    for (let row = dataStartRow; row <= dataEndRow; row++) {
      cells.push({ value: extractCellValue(worksheet, row, col) });
    }
    return { worksheet, cells };
  }
  return undefined;
}

function parseStructuredReference(
  formula: string
): { tableName: string; columnName: string } | undefined {
  const table = readStructuredTableName(formula);
  if (!table || formula[table.end] !== "[" || !formula.endsWith("]")) {
    return undefined;
  }
  const body = formula.slice(table.end + 1, -1);
  if (body.length === 0) {
    return undefined;
  }
  const columnName = extractStructuredReferenceColumn(body);
  return columnName ? { tableName: table.name, columnName } : undefined;
}

function readStructuredTableName(formula: string): { name: string; end: number } | undefined {
  if (formula.startsWith("'")) {
    let name = "";
    for (let i = 1; i < formula.length; i++) {
      if (formula[i] === "'" && formula[i + 1] === "'") {
        name += "'";
        i++;
      } else if (formula[i] === "'") {
        return { name, end: i + 1 };
      } else {
        name += formula[i];
      }
    }
    return undefined;
  }
  const bracket = formula.indexOf("[");
  if (bracket <= 0) {
    return undefined;
  }
  return { name: formula.slice(0, bracket), end: bracket };
}

function extractStructuredReferenceColumn(body: string): string | undefined {
  if (body.startsWith("@")) {
    const inner = body.slice(1).replace(/^\[(.*)\]$/, "$1");
    const item = readStructuredReferenceItem(`${inner}]`, 0);
    return item.value;
  }
  if (!body.startsWith("[")) {
    if (body.startsWith("#")) {
      return undefined;
    }
    const item = readStructuredReferenceItem(`${body}]`, 0);
    return item.value;
  }
  const items: string[] = [];
  let i = 0;
  while (i < body.length) {
    if (body[i] !== "[") {
      i++;
      continue;
    }
    const item = readStructuredReferenceItem(body, i + 1);
    items.push(item.value);
    i = item.end;
  }
  return items.filter(item => item && !item.startsWith("#")).pop();
}

function readStructuredReferenceItem(value: string, start: number): { value: string; end: number } {
  let i = start;
  let result = "";
  while (i < value.length) {
    if (value[i] === "'" && i + 1 < value.length) {
      result += value[i + 1];
      i += 2;
    } else if (value[i] === "]") {
      i++;
      break;
    } else {
      result += value[i];
      i++;
    }
  }
  return { value: result.trim(), end: i };
}

function resolveReferenceMatrix(
  formula: string,
  ctx: ResolverContext
): ResolvedReferenceMatrix | undefined {
  // Expand defined-name references up front so a named multi-area range can
  // still drive a multi-level string cache.
  const named = findDefinedName(formula, ctx);
  if (named) {
    const visitKey = storageKey(named.name, named.localSheetId);
    if (ctx.visitedDefinedNames.has(visitKey)) {
      return undefined;
    }
    ctx.visitedDefinedNames.add(visitKey);
    try {
      let firstWorksheet: Worksheet | undefined;
      const mergedValues: unknown[][] = [];
      let mergedColumnCount = 0;
      for (const rangeStr of named.ranges) {
        const inner = resolveReferenceMatrix(rangeStr, ctx);
        if (!inner) {
          continue;
        }
        if (!firstWorksheet) {
          firstWorksheet = inner.worksheet;
        }
        mergedValues.push(...inner.values);
        mergedColumnCount = Math.max(mergedColumnCount, inner.columnCount);
      }
      if (!firstWorksheet) {
        return undefined;
      }
      return {
        worksheet: firstWorksheet,
        values: mergedValues,
        rowCount: mergedValues.length,
        columnCount: mergedColumnCount
      };
    } finally {
      ctx.visitedDefinedNames.delete(visitKey);
    }
  }

  const parts = splitFormulaRanges(formula);
  let worksheet: Worksheet | undefined;
  const rows: unknown[][] = [];
  let columnCount = 0;

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
    const ws = ctx.workbook.getWorksheet(sheetName);
    if (!ws) {
      return undefined;
    }
    if (!worksheet) {
      worksheet = ws;
    }

    if ("top" in decoded && "left" in decoded) {
      const top = decoded.top as number;
      const left = decoded.left as number;
      const bottom = decoded.bottom as number;
      const right = decoded.right as number;
      columnCount = Math.max(columnCount, right - left + 1);
      for (let r = top; r <= bottom; r++) {
        const row: unknown[] = [];
        for (let c = left; c <= right; c++) {
          row.push(extractCellValue(ws, r, c));
        }
        rows.push(row);
      }
    } else if ("row" in decoded && "col" in decoded) {
      columnCount = Math.max(columnCount, 1);
      rows.push([extractCellValue(ws, decoded.row as number, decoded.col as number)]);
    }
  }

  if (!worksheet) {
    return undefined;
  }
  return { worksheet, values: rows, rowCount: rows.length, columnCount };
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
