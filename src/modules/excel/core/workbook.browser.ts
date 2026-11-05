/**
 * Workbook - Cross-platform Excel Workbook (Browser Version)
 *
 * Base implementation with all platform-agnostic functionality.
 * Node.js version (workbook.ts) extends this with file system support for CSV.
 *
 * Features:
 * - xlsx: File/stream/buffer support (file operations Node.js only)
 * - csv: CSV read/write support (file operations Node.js only)
 * - streaming: `Stream.WorkbookWriter` / `Stream.WorkbookReader` for large files
 */

import { nameCall } from "@excel/chart/api/call";
// Chart runtime is imported statically. The chart modules depend only on the
// `*-core` data layer (never on this file), so there is no import cycle, and
// the bundler tree-shakes the whole chart tree out of consumer bundles that
// never create or serialise a chart.
import { fillChartCaches, fillChartExCaches } from "@excel/chart/bridge/excel-chart-host";
import { buildChartModel, buildComboChartModel } from "@excel/chart/build/chart-builder";
import { buildChartExModel } from "@excel/chart/build/chart-ex-builder";
import type {
  AddChartExOptions,
  ChartExEntry,
  ChartExModel
} from "@excel/chart/model/chart-ex-types";
import type { AddComboChartOptions, ChartEntry, ChartModel } from "@excel/chart/model/types";
import { resolvePendingChartImages } from "@excel/chart/serialize/chart-images";
import { buildChartColors, buildChartStyle } from "@excel/chart/serialize/chart-sidecar";
import type {
  ChartsheetData,
  AddChartsheetOptions,
  AddPivotChartsheetOptions
} from "@excel/core/chartsheet";
import { createChartsheet } from "@excel/core/chartsheet";
import { createDefinedNames, definedNamesHasName } from "@excel/core/defined-names";
import { withPivotChartSource } from "@excel/core/pivot-chart";
import type { PivotTable } from "@excel/core/pivot-table";
import { sheetNamesInTabOrder, withReferenceEdit } from "@excel/core/structural-edit";
import { keptSettingsReport } from "@excel/core/unrecognized-values";
import type { WorkbookData } from "@excel/core/workbook-core";
import {
  addChartEntry,
  addChartExStructuredEntry,
  getChartEntry,
  getChartExStructuredEntry,
  getWorksheet,
  getWorksheets,
  copyChartSidecars,
  copyChartExSidecars,
  nextChartExNumber,
  nextChartNumber,
  releaseChartMedia,
  removeChartEntry,
  removeChartExStructuredEntry,
  checkWorkbook,
  refuseLongSheetName,
  setChartColors,
  setChartStyle,
  validateSheetName
} from "@excel/core/workbook-core";
import type { UnrecognizedValue, WorkbookCreateOptions } from "@excel/core/workbook-io-types";
import type { ExternalLinkModel } from "@excel/core/workbook-model";
import type { Worksheet, WorksheetModel } from "@excel/core/worksheet";
import {
  createWorksheet,
  destroy,
  getSheetModel,
  getSheetName,
  getSheetWorkbook
} from "@excel/core/worksheet";
import { setSheetModel } from "@excel/core/worksheet-load";
import { ExcelArgumentError } from "@excel/errors";
import type {
  WorkbookReaderOptions,
  WorkbookReaderInput
} from "@excel/stream/workbook-reader.browser";
import { WorkbookReader } from "@excel/stream/workbook-reader.browser";
import type { WorkbookWriterOptions } from "@excel/stream/workbook-writer.browser";
import { WorkbookWriter } from "@excel/stream/workbook-writer.browser";
import type {
  AddWorksheetOptions,
  Font,
  ThreadedCommentPerson,
  WorkbookProtection
} from "@excel/types";
import { argumentChecks } from "@excel/utils/arguments";
import { synthGuid } from "@excel/utils/guid";
import { renameTablesInFormula } from "@excel/utils/reference-rewrite";
import { splitPasswordOptions, verifySheetPassword } from "@excel/utils/sheet-protection";
import { discardKeptSheetSettings } from "@excel/utils/sheet-settings-checks";
import { buildWorkbookProtection } from "@excel/utils/workbook-protection";
import type { ChartsheetModel } from "@excel/xlsx/xform/sheet/chartsheet-xform";
import type { RuntimeValue } from "@formula/runtime/values";
import { excelNameKey, sheetNameKey } from "@utils/name-key";
import { isRecordArgument } from "@utils/object";

// =============================================================================
// Internal Types
// =============================================================================

// The serialisation model and the workbook-shaped types it is built from live in
// `core/workbook-model.ts`, below both writers. Re-exported here because this is the module
// consumers and the rest of `excel/` name them through.
export type {
  ExternalLinkCachedSheet,
  ExternalLinkModel,
  WorkbookMedia,
  WorkbookModel,
  WorkbookProtectionModel
} from "@excel/core/workbook-model";

function isComboChartOptions(chart: AddChartsheetOptions["chart"]): chart is AddComboChartOptions {
  return !!chart && typeof chart === "object" && "groups" in chart;
}

function isChartExOptions(chart: AddChartsheetOptions["chart"]): chart is AddChartExOptions {
  return !!chart && typeof chart === "object" && "type" in chart && isChartExType(chart.type);
}

function isChartExType(type: string): boolean {
  return (
    type === "sunburst" ||
    type === "treemap" ||
    type === "waterfall" ||
    type === "funnel" ||
    type === "histogram" ||
    type === "pareto" ||
    type === "boxWhisker" ||
    type === "regionMap"
  );
}

// =============================================================================
// Workbook Class
// =============================================================================

/** A new, empty workbook. `options.formulaSyntaxProbe` decides how defined names that may be formulas are classified. */
export function createWorkbook(options?: WorkbookCreateOptions): WorkbookData {
  const check = argumentChecks("Workbook.create");
  if (options !== null && check.record("options", options)) {
    const probe = options.formulaSyntaxProbe;
    if (probe !== undefined && probe !== null && typeof probe !== "function") {
      check.fail("options.formulaSyntaxProbe", "a function", probe);
    }
  }
  const wb = {} as WorkbookData;
  wb._tableNames = new Set<string>();

  wb.category = "";
  wb.company = "";
  wb.created = new Date();
  wb.description = "";
  wb.keywords = "";
  wb.manager = "";
  wb.modified = wb.created;
  wb.properties = {};
  wb.calcProperties = {};
  wb._worksheets = [];
  wb.subject = "";
  wb.title = "";
  wb.views = [];
  wb.media = [];
  wb.pivotTables = [];
  wb.externalLinks = [];
  wb._chartEntries = {};
  wb._chartRels = {};
  wb._chartStyles = {};
  wb._chartColors = {};
  wb._chartExStyles = {};
  wb._chartExColors = {};
  wb._chartExEntries = {};
  wb._chartExRels = {};
  wb._chartExStructuredEntries = {};
  wb._chartsheets = [];
  wb._persons = [];
  wb._slicerParts = {};
  wb._slicerCacheParts = {};
  wb._timelineParts = {};
  wb._timelineCacheParts = {};
  wb._opaqueParts = [];
  wb._opaqueContentTypeDefaults = {};
  wb._opaqueDrops = [];
  wb._writerExternalLinkCache = new Map();
  wb._definedNames = createDefinedNames(options?.formulaSyntaxProbe);
  // Live: the registry refuses a name a table holds — see `DefinedNamesData._tableNames`.
  wb._definedNames._tableNames = wb._tableNames;

  return wb;
}

/**
 * The names the tables of `source` take in `wb`: a name the target already uses — for a table or a defined name — is
 * given the next free one, as Excel's "Move or Copy" does (`Sales` → `Sales2`, `Sales3`, …). Keyed by `excelNameKey`.
 */
function freeTableNames(wb: WorkbookData, source: Worksheet): Map<string, string> {
  const renames = new Map<string, string>();
  const taken = (candidate: string): boolean => {
    const key = excelNameKey(candidate);
    return (
      wb._tableNames.has(key) ||
      definedNamesHasName(wb._definedNames, candidate) ||
      [...renames.values()].some(name => excelNameKey(name) === key)
    );
  };
  for (const tableName of Object.keys(source.tables ?? {})) {
    if (!taken(tableName)) {
      continue;
    }
    let n = 2;
    while (taken(`${tableName}${n}`)) {
      n++;
    }
    renames.set(excelNameKey(tableName), `${tableName}${n}`);
  }
  return renames;
}

/**
 * `model` with its tables renamed by `renames` and every reference the sheet makes to them — cell formulas, table
 * column formulas, conditional formats and validations — restated under the new names. Nothing in `model` itself is
 * changed: it shares objects with the source sheet.
 */
function withRenamedTables(
  model: WorksheetModel,
  renames: ReadonlyMap<string, string>
): WorksheetModel {
  if (renames.size === 0) {
    return model;
  }
  const rename = (formula: unknown): unknown =>
    typeof formula === "string" ? renameTablesInFormula(formula, renames) : formula;
  const renameList = (list: unknown): unknown => (Array.isArray(list) ? list.map(rename) : list);
  const dataValidations: Record<string, unknown> = {};
  for (const [key, rule] of Object.entries(model.dataValidations ?? {})) {
    dataValidations[key] =
      rule && typeof rule === "object"
        ? { ...rule, formulae: renameList((rule as { formulae?: unknown }).formulae) }
        : rule;
  }
  return {
    ...model,
    rows: model.rows?.map(row => ({
      ...row,
      cells: row.cells.map(cell => {
        const formula = (cell as { formula?: unknown }).formula;
        return typeof formula === "string" ? { ...cell, formula: rename(formula) as string } : cell;
      })
    })),
    tables: (model.tables ?? []).map(table => {
      const renamed = renames.get(excelNameKey(table.name));
      return {
        ...table,
        ...(renamed === undefined ? {} : { name: renamed, displayName: renamed }),
        columns: table.columns.map(column => ({
          ...column,
          calculatedColumnFormula: rename(column.calculatedColumnFormula) as string | undefined,
          totalsRowFormula: rename(column.totalsRowFormula) as string | undefined
        }))
      };
    }),
    conditionalFormattings: (model.conditionalFormattings ?? []).map(formatting => ({
      ...formatting,
      rules: formatting.rules.map(rule => {
        const formulae = (rule as { formulae?: unknown }).formulae;
        return formulae === undefined
          ? rule
          : ({ ...rule, formulae: renameList(formulae) } as typeof rule);
      })
    })),
    dataValidations: dataValidations as WorksheetModel["dataValidations"]
  };
}

/**
 * Copy the worksheet `source` — from this workbook or another — into this one under `name` (by default its own) and
 * return the copy. Cells, styles, merges, tables, images and charts come along; media from another workbook are
 * registered here.
 *
 * A copied table whose name the target already uses — for a table or a defined name — is renamed as Excel's "Move or
 * Copy" renames it: `Sales` becomes `Sales2` (or the next free number), and the copied sheet's own references to it —
 * cell formulas, table column formulas, conditional formats and validations — are restated under the new name.
 * References from other sheets keep naming the original.
 *
 * Everything that can refuse the import is checked before the sheet is added, so a refused import adds nothing.
 *
 * @throws {ExcelArgumentError} when `source` is not a worksheet or `name` is not a string.
 * @throws {ExcelWorksheetNameError} when `name` is taken, longer than 31 characters or not a name Excel accepts.
 */
export function importWorksheet(wb: WorkbookData, source: Worksheet, name?: string): Worksheet {
  checkWorkbook(wb);
  // Everything that can refuse the import is checked before the sheet is added, so a refused import leaves nothing
  // behind.
  const check = argumentChecks("Workbook.importWorksheet");
  if (
    typeof source !== "object" ||
    source === null ||
    !Array.isArray((source as { _rows?: unknown })._rows) ||
    typeof (source as { _name?: unknown })._name !== "string"
  ) {
    check.fail("source", "a worksheet", source);
  }
  check.string("name", name);
  refuseLongSheetName(name);
  validateSheetName(wb, name ?? getSheetName(source));

  // Deep copy via model: the getter serializes ALL worksheet properties and the
  // setter deserializes them, so future properties are automatically included.
  const sourceModel = withRenamedTables(getSheetModel(source), freeTableNames(wb, source));
  const newWs = addWorksheet(wb, name ?? getSheetName(source));
  // Remap chart numbers so the source's `chartNumber` / `chartExNumber`
  // references point at entries we actually copy into the target
  // workbook. Build the map here so the rewritten `charts` array and
  // the copied entries use consistent ids.
  const chartMap = new Map<number, number>();
  const chartExMap = new Map<number, number>();
  const sourceWorkbook = getSheetWorkbook(source);
  const differentWorkbook = sourceWorkbook !== wb;
  const sourceCharts = sourceModel.charts ?? [];

  // Worksheet media entries store workbook-level image indexes. A plain model
  // clone is sufficient within one workbook, but across workbooks those indexes
  // point into the SOURCE media registry. Re-register every referenced medium
  // (and an optional SVG companion) in the destination and rewrite the sheet
  // model's imageId; otherwise the writer either drops the image (missing id) or
  // silently renders an unrelated destination image that happens to share it.
  const importedMedia = structuredClone(sourceModel.media ?? []);
  if (differentWorkbook) {
    const mediaMap = new Map<number, number>();
    const copyMedium = (sourceId: number): number | undefined => {
      const existing = mediaMap.get(sourceId);
      if (existing !== undefined) {
        return existing;
      }
      const sourceMedium = sourceWorkbook.media[sourceId];
      if (!sourceMedium) {
        return undefined;
      }
      const cloned = structuredClone(sourceMedium);
      const targetId = wb.media.length;
      // Reserve before following svgMediaId so a malformed self-reference
      // cannot recurse forever.
      mediaMap.set(sourceId, targetId);
      wb.media.push(cloned);
      if (sourceMedium.svgMediaId !== undefined) {
        cloned.svgMediaId = copyMedium(sourceMedium.svgMediaId);
      }
      return targetId;
    };

    for (const medium of importedMedia) {
      const sourceId = Number(medium.imageId);
      if (!Number.isInteger(sourceId) || sourceId < 0) {
        continue;
      }
      const targetId = copyMedium(sourceId);
      if (targetId !== undefined) {
        medium.imageId = String(targetId);
      }
    }
  }
  // `nextChartNumber()` / `nextChartExNumber()` compute `max(existing) + 1`
  // from the entry maps — they do NOT reserve a slot. Calling them in
  // a tight loop without an intervening `addChartEntry` therefore
  // returns the SAME number N times, and the second loop below then
  // overwrites `_chartEntries[dstNum]` repeatedly — only the last
  // cloned entry survives, the others are silently lost. Track the
  // allocator locally so each source chart gets a unique target slot.
  let nextChartAlloc = nextChartNumber(wb);
  let nextChartExAlloc = nextChartExNumber(wb);
  for (const anchor of sourceCharts) {
    if (anchor.chartNumber && anchor.chartNumber > 0 && !chartMap.has(anchor.chartNumber)) {
      chartMap.set(anchor.chartNumber, nextChartAlloc++);
    }
    if (anchor.chartExNumber && anchor.chartExNumber > 0 && !chartExMap.has(anchor.chartExNumber)) {
      chartExMap.set(anchor.chartExNumber, nextChartExAlloc++);
    }
  }
  const remappedCharts = sourceCharts.map(anchor => ({
    ...anchor,
    chartNumber: anchor.chartNumber
      ? (chartMap.get(anchor.chartNumber) ?? anchor.chartNumber)
      : anchor.chartNumber,
    chartExNumber: anchor.chartExNumber
      ? (chartExMap.get(anchor.chartExNumber) ?? anchor.chartExNumber)
      : anchor.chartExNumber
  }));

  // Deep-clone the loaded drawing container and give it a FRESH identity.
  //
  // `getSheetModel`/`setSheetModel` pass `drawing` through by reference, so
  // without this the imported sheet shared the source's `WorksheetModel.drawing`
  // — including its `name` (e.g. "drawing1"). Both sheets then registered a
  // `[Content_Types].xml` Override for the SAME `/xl/drawings/drawingN.xml`
  // part, which is invalid per OPC (duplicate PartName) and trips the built-in
  // validator (issue #189).
  //
  // The writer rebuilds each drawing part's XML from `drawing.anchors` /
  // `drawing.rels` and auto-allocates a fresh, collision-free name for any
  // drawing whose `name` is empty (its allocator is seeded past every name a
  // round-tripped sheet already claims — see WorkSheetXform prepare/prepareModel).
  // So the correct fix is a value copy with the name cleared — not a byte copy
  // of the original part. Chart anchors embedded in the drawing are renumbered
  // with the same `chartMap`/`chartExMap` as the top-level `charts` array so
  // both stay consistent.
  const remapDrawing = (drawing: unknown): unknown => {
    if (!drawing || typeof drawing !== "object") {
      return drawing;
    }
    const cloned = structuredClone(drawing) as {
      name?: string;
      anchors?: Array<{ chartNumber?: number; chartExNumber?: number }>;
    };
    // Clear the name so the writer allocates a fresh, unique one.
    cloned.name = undefined;
    for (const anchor of cloned.anchors ?? []) {
      if (anchor.chartNumber) {
        anchor.chartNumber = chartMap.get(anchor.chartNumber) ?? anchor.chartNumber;
      }
      if (anchor.chartExNumber) {
        anchor.chartExNumber = chartExMap.get(anchor.chartExNumber) ?? anchor.chartExNumber;
      }
    }
    return cloned;
  };

  try {
    applyImportedModel();
  } catch (error) {
    // A model the target refuses part-way leaves no half-imported sheet behind.
    removeWorksheet(wb, newWs);
    throw error;
  }
  function applyImportedModel(): void {
    setSheetModel(newWs, {
      ...sourceModel,
      id: newWs.id,
      name: getSheetName(newWs),
      media: importedMedia,
      charts: remappedCharts,
      drawing: remapDrawing(sourceModel.drawing),
      // getSheetModel()/setSheetModel() pass these through by reference (unlike
      // `charts`, which is rebuilt above), so two sheets would otherwise share
      // the same TableModel / PivotTable / rule / form-control instances —
      // mutating one sheet's copy silently mutated the source. Structured-clone
      // at the import boundary gives the imported sheet independent data. These
      // carry no cross-sheet identity to remap (unlike charts/drawings), so a
      // plain deep copy is the whole fix.
      tables: structuredClone(sourceModel.tables ?? []),
      pivotTables: structuredClone(sourceModel.pivotTables ?? []),
      conditionalFormattings: structuredClone(sourceModel.conditionalFormattings ?? []),
      formControls: structuredClone(sourceModel.formControls ?? [])
    });
  }

  // Copy the actual chart parts + sidecars into the target workbook
  // so the remapped `charts` array references live entries. Without
  // this, `importWorksheet` left the target with chart anchors but no
  // backing chart XML, producing a broken package on save. We copy
  // both the structured model (via `getChartEntry` / `addChartEntry`
  // — the public API) and all sidecars (`copyChartSidecars` /
  // `copyChartExSidecars`).
  if (chartMap.size > 0 || chartExMap.size > 0) {
    for (const [srcNum, dstNum] of chartMap) {
      const entry = getChartEntry(sourceWorkbook, srcNum);
      if (!entry) {
        continue;
      }
      // Deep-clone the entry with every metadata field preserved —
      // rawData / userShapesXml (byte slices), modelSnapshot and the
      // dirty / preferRawPatch / requireRawPatch writer hints, plus
      // per-entry `rels`. Previously only `model`, `rawData`, and
      // `userShapesXml` were copied, so the cross-workbook import
      // path produced charts where the raw-patch fast path couldn't
      // run and the change-detection snapshot didn't reflect the
      // source entry's load-time state.
      addChartEntry(wb, cloneChartEntry(entry, dstNum));
      if (differentWorkbook) {
        copyChartSidecars(sourceWorkbook, srcNum, dstNum, wb);
      } else {
        copyChartSidecars(wb, srcNum, dstNum);
      }
    }
    for (const [srcNum, dstNum] of chartExMap) {
      const exEntry = getChartExStructuredEntry(sourceWorkbook, srcNum);
      if (exEntry) {
        addChartExStructuredEntry(wb, cloneChartExEntry(exEntry, dstNum));
      } else {
        const rawBytes = sourceWorkbook._chartExEntries?.[srcNum];
        if (rawBytes) {
          wb._chartExEntries[dstNum] = rawBytes.slice();
        }
      }
      if (differentWorkbook) {
        copyChartExSidecars(sourceWorkbook, srcNum, dstNum, wb);
      } else {
        copyChartExSidecars(wb, srcNum, dstNum);
      }
    }
  }

  return newWs;
}

/** Options for {@link protectWorkbook}: the protection flags, plus the password that unlocks them. */
export interface ProtectWorkbookOptions extends Partial<WorkbookProtection> {
  /** Password hashed into the workbook; omit to protect without one. */
  password?: string;
}

/** Protect the workbook's structure or windows: `protect(wb, { password: "pw", lockStructure: true })`. */
export function protectWorkbook(wb: WorkbookData, options?: ProtectWorkbookOptions): Promise<void>;
/**
 * @deprecated Pass the password in the options object: `protect(wb, { password, ...options })`.
 * Removed in 1.0.
 */
export function protectWorkbook(
  wb: WorkbookData,
  password: string | undefined,
  options?: Partial<WorkbookProtection>
): Promise<void>;
export async function protectWorkbook(
  wb: WorkbookData,
  passwordOrOptions?: string | ProtectWorkbookOptions,
  legacyOptions?: Partial<WorkbookProtection>
): Promise<void> {
  checkWorkbook(wb);
  const [password, options] = splitPasswordOptions(passwordOrOptions, legacyOptions);
  wb.protection = await buildWorkbookProtection(password, options);
}

/**
 * Whether `password` is the one the workbook's structure or windows were protected with — the hash `protect` stored,
 * or a file's — as `Worksheet.verifyPassword` checks a sheet's. `false` when the workbook is not protected, or was
 * protected without a password.
 *
 * @throws {ExcelArgumentError} when `password` is not a string (the promise rejects).
 */
export async function verifyWorkbookPassword(wb: WorkbookData, password: string): Promise<boolean> {
  checkWorkbook(wb);
  argumentChecks("Workbook.verifyPassword").required("password", password);
  argumentChecks("Workbook.verifyPassword").string("password", password);
  return verifySheetPassword(wb.protection, password);
}

/** Remove the workbook's structure and window protection, as `protect` set it. */
export function unprotectWorkbook(wb: WorkbookData): void {
  checkWorkbook(wb);
  wb.protection = undefined;
}

export function addWorksheet(
  wb: WorkbookData,
  name?: string,
  options?: AddWorksheetOptions
): Worksheet {
  checkWorkbook(wb);
  const id = getNextId(wb);

  // Allocate `orderNo` from the unified worksheet+chartsheet counter.
  // Looking only at `_worksheets` here (the previous implementation)
  // silently collides when a chartsheet has been added in between:
  // e.g. `addWorksheet("A")` → orderNo 0; `addChartsheet(…)` → 1
  // (via `_nextSheetOrderNo()`); `addWorksheet("B")` → 1 again
  // (because `max(worksheets.orderNo) + 1 = 0 + 1 = 1`), so A and
  // B share an ordinal with the chartsheet. The writer's stable
  // sort then interleaves them non-deterministically, scrambling
  // the user's tab order (`[A, CS, B]` could come out as
  // `[A, B, CS]` or `[A, CS, B]` across runs).
  const orderNo = _nextSheetOrderNo(wb);
  const worksheetOptions = {
    ...options,
    id,
    name,
    orderNo,
    workbook: wb
  };

  const worksheet = createWorksheet(worksheetOptions);

  wb._worksheets[id] = worksheet;
  return worksheet;
}

/**
 * The sheet settings a read kept because they are outside their setting's vocabulary, as the workbook keeps them
 * **now** — each with its sheet, its path (`pageSetup.orientation`) and its text, in sheet order: what
 * `Workbook.readWithDiagnostics` reported as `unrecognizedValues`, less what an edit or
 * {@link discardUnrecognizedValues} has since replaced or dropped. The entries are those the discard filter is called
 * with. Copies.
 */
export function listUnrecognizedValues(wb: WorkbookData): UnrecognizedValue[] {
  checkWorkbook(wb);
  return keptSettingsReport(wb);
}

/** Options for {@link discardUnrecognizedValues}. */
export interface DiscardUnrecognizedValuesOptions {
  /**
   * Which kept settings to drop: called with each — its sheet, its path (`pageSetup.orientation`) and its text — and
   * drops those it returns `true` for. Without it, every one is dropped.
   */
  readonly filter?: (setting: UnrecognizedValue) => boolean;
}

/**
 * Drop the sheet settings a read kept because they are outside their setting's vocabulary — what
 * `Workbook.readWithDiagnostics` lists as `unrecognizedValues` — so the workbook is written with each such setting as
 * it now reads. Returns how many there were.
 *
 * Until then a kept value is written back as read while its field is absent; assigning the field takes over, but
 * clearing it again brings the kept text back. This is how to remove one instead — all of them, or with `filter` the
 * ones you choose:
 *
 * ```ts
 * Workbook.discardUnrecognizedValues(wb, { filter: s => s.sheet === "Report" && s.path.startsWith("views") });
 * ```
 */
export function discardUnrecognizedValues(
  wb: WorkbookData,
  options?: DiscardUnrecognizedValuesOptions
): number {
  checkWorkbook(wb);
  const check = argumentChecks("Workbook.discardUnrecognizedValues");
  check.record("options", options);
  const filter = options?.filter;
  if (filter !== undefined && typeof filter !== "function") {
    check.fail("options.filter", "a function", filter);
  }
  let discarded = 0;
  for (const ws of getWorksheets(wb)) {
    const sheet = getSheetName(ws);
    discarded += discardKeptSheetSettings(
      ws,
      filter === undefined ? undefined : setting => filter({ sheet, ...setting })
    );
  }
  return discarded;
}

/**
 * Remove a worksheet — given by id, name or handle — and turn references to it into `#REF!`. Returns `true` if it
 * was in the workbook.
 */
export function removeWorksheet(wb: WorkbookData, sheet: number | string | Worksheet): boolean {
  checkWorkbook(wb);
  // `getWorksheet` reads a missing id as "the first sheet"; removing it for a forgotten argument loses a sheet.
  const check = argumentChecks("Workbook.removeWorksheet");
  check.required("sheet", sheet);
  if (typeof sheet !== "number" && typeof sheet !== "string" && typeof sheet !== "object") {
    check.fail("sheet", "an id, a name or a worksheet", sheet);
  }
  const worksheet =
    typeof sheet === "object" && sheet !== null
      ? wb._worksheets[sheet.id] === sheet
        ? sheet
        : undefined
      : getWorksheet(wb, sheet);
  if (!worksheet) {
    return false;
  }
  destroy(worksheet);
  return true;
}

/**
 * Add a chartsheet — a tab that holds one chart — and return it. `name` defaults to `ChartN`.
 *
 * @throws {ExcelWorksheetNameError} when `name` is taken, longer than 31 characters or not a name Excel accepts.
 */
export function addChartsheet(
  wb: WorkbookData,
  name: string | undefined,
  options: AddChartsheetOptions
): ChartsheetData {
  checkWorkbook(wb);
  const sheetName = _validateChartsheetName(wb, name ?? `Chart${wb._chartsheets.length + 1}`);
  const sheetNo = _nextChartsheetNo(wb);
  const id = _nextSheetId(wb);
  // Assign a unified `orderNo` across worksheets and chartsheets so
  // the writer can preserve the author's interleaved tab layout.
  // Without this, workbook-xform `prepare()` sorted by `sheetNo`
  // (file-path number, independent per family) and reordered
  // `[ws1, cs1, ws2]` into `[ws1, ws2, cs1]`.
  const orderNo = _nextSheetOrderNo(wb);
  const chartsheet: ChartsheetModel = {
    sheetNo,
    id,
    name: sheetName,
    orderNo,
    state: options.state ?? "visible",
    tabSelected: options.tabSelected,
    zoomScale: options.zoomScale,
    workbookViewId: options.workbookViewId,
    zoomToFit: options.zoomToFit,
    pageMargins: options.pageMargins,
    pageSetup: options.pageSetup,
    headerFooter: options.headerFooter,
    drawing: { rId: "rId1" }
  };

  return chartArgument("Workbook.addChartsheet", "options.chart", () =>
    addChartsheetModel(wb, chartsheet, options)
  );
}

/** Build the chartsheet's chart and add both — the second half of {@link addChartsheet}. */
function addChartsheetModel(
  wb: WorkbookData,
  chartsheet: ChartsheetModel,
  options: AddChartsheetOptions
): ChartsheetData {
  if (isChartExOptions(options.chart)) {
    const chartExNumber = nextChartExNumber(wb);
    const model = buildChartExModel(options.chart);
    try {
      fillChartExCaches(model, wb);
    } catch {
      // Cache population is best-effort; never let it break chart creation.
    }
    addChartExStructuredEntry(wb, { chartExNumber, model });
    chartsheet.chartExNumber = chartExNumber;
  } else {
    const chartNumber = nextChartNumber(wb);
    const chartModel = isComboChartOptions(options.chart)
      ? buildComboChartModel(options.chart)
      : buildChartModel(options.chart);
    try {
      fillChartCaches(chartModel, wb);
    } catch {
      // Cache population is best-effort; never let it break chart creation.
    }
    const entry: ChartEntry = { chartNumber, model: chartModel };
    // Resolve programmatic `series.spPr.fill.blip._pendingImage`
    // payloads into workbook media entries and chart rels. The
    // worksheet-embedded `addChart` path does this immediately
    // after `fillChartCaches`; chartsheets ran the same builder
    // output but skipped the image-resolution step entirely, so a
    // picture-fill series authored via `addChartsheet` was
    // registered with its `_pendingImage` stuck on the model and
    // never reached `media/imageN.{ext}` — Excel rendered the
    // series as a transparent fill. Safe to call before
    // `addChartEntry` so the stored entry carries its resolved
    // `entry.rels` from the start.
    try {
      resolvePendingChartImages(entry, wb, chartNumber);
    } catch {
      // Image resolution is best-effort; a broken image payload
      // should never take down chart creation — the series keeps
      // its `pictureOptions`, just without the blipFill.
    }
    addChartEntry(wb, entry);
    _applyChartsheetSidecars(wb, chartNumber, options.chart);
    chartsheet.chartNumber = chartNumber;
  }

  wb._chartsheets.push(chartsheet);
  return createChartsheet(chartsheet, wb);
}

/** {@link addChartsheet} for a chart whose data is the pivot table `pivotTable`. */
export function addPivotChartsheet(
  wb: WorkbookData,
  name: string | undefined,
  pivotTable: PivotTable,
  options: AddPivotChartsheetOptions
): ChartsheetData {
  checkWorkbook(wb);
  const chart = chartArgument("Workbook.addPivotChartsheet", "options.chart", () =>
    withPivotChartSource(pivotTable, options.chart)
  );
  try {
    return addChartsheet(wb, name, { ...options, chart });
  } catch (error) {
    throw renamedCall(error, "Workbook.addChartsheet", "Workbook.addPivotChartsheet");
  }
}

/**
 * Run the chart validator for a chart the caller passed at `root` — `options.chart` for {@link addChartsheet}, `chart`
 * for {@link replaceChartsheetChart} — so an argument error names the caller's path, not the validator's own
 * `options.series[0].values`, and its message starts with `where`.
 */
function chartArgument<T>(where: string, root: string, run: () => T): T {
  try {
    return run();
  } catch (error) {
    throw nameCall(rooted(error, root), where);
  }
}

/** `error` with a field the validator rooted at `options` re-rooted at `root`, in the field and the message. */
function rooted(error: unknown, root: string): unknown {
  if (!(error instanceof ExcelArgumentError)) {
    return error;
  }
  const field = error.field;
  if (field !== "options" && !field.startsWith("options.") && !field.startsWith("options[")) {
    return error;
  }
  const next = `${root}${field.slice("options".length)}`;
  const before = error.message;
  const at = before.indexOf(field);
  error.message =
    at < 0 ? before : `${before.slice(0, at)}${next}${before.slice(at + field.length)}`;
  if (typeof error.stack === "string" && error.stack.includes(before)) {
    error.stack = error.stack.replace(before, error.message);
  }
  (error as { field: string }).field = next;
  return error;
}

/** An argument error that named the call `from`, named for `to` — a public call made through another one. */
function renamedCall(error: unknown, from: string, to: string): unknown {
  if (error instanceof ExcelArgumentError && error.message.startsWith(`${from}: `)) {
    const before = error.message;
    error.message = `${to}: ${before.slice(from.length + 2)}`;
    if (typeof error.stack === "string" && error.stack.includes(before)) {
      error.stack = error.stack.replace(before, error.message);
    }
  }
  return error;
}

/**
 * The chartsheet named `nameOrIndex` (compared as Excel compares sheet names), or at that **0-based position** among
 * the chartsheets — not the sheet id `getWorksheet` takes — or `undefined` when there is none. A position moves when
 * an earlier chartsheet is removed; to keep hold of a chartsheet, keep its handle or its name.
 *
 * @param nameOrIndex - The chartsheet's name, or its 0-based position among the chartsheets (not its sheet id).
 */
export function getChartsheet(
  wb: WorkbookData,
  nameOrIndex: string | number
): ChartsheetData | undefined {
  checkWorkbook(wb);
  const model = _getChartsheetModel(wb, nameOrIndex);
  return model ? createChartsheet(model, wb) : undefined;
}

/**
 * Remove a chartsheet — the handle `addChartsheet` / `getChartsheet` returned, its name, or its 0-based position among
 * the chartsheets. Names scoped to it go and later scopes renumber. Returns `true` if one was removed.
 */
export function removeChartsheet(
  wb: WorkbookData,
  chartsheet: string | number | ChartsheetData
): boolean {
  checkWorkbook(wb);
  const index = chartsheetIndex(wb, chartsheet, "Workbook.removeChartsheet", "chartsheet");
  if (index < 0) {
    return false;
  }
  // A chartsheet holds no cells, but it holds a tab position: names scoped to it go, and later scopes renumber.
  const edit = {
    kind: "removeSheet",
    sheet: wb._chartsheets[index].name,
    sheets: sheetNamesInTabOrder(wb)
  } as const;
  withReferenceEdit(wb, edit, () => {
    const [removed] = wb._chartsheets.splice(index, 1);
    // Its id is not handed out again — see `getNextId`.
    if (removed && typeof removed.id === "number" && Number.isFinite(removed.id)) {
      wb._lastSheetId = Math.max(wb._lastSheetId ?? 0, removed.id);
    }
    if (removed.chartNumber) {
      releaseChartMedia(wb, removed.chartNumber);
      removeChartEntry(wb, removed.chartNumber);
    }
    if (removed.chartExNumber) {
      removeChartExStructuredEntry(wb, removed.chartExNumber);
    }
  });
  return true;
}

/**
 * The position among `wb`'s chartsheets of a chartsheet given by name (compared as Excel compares sheet names),
 * 0-based position or handle — or `-1` when `wb` has no such chartsheet. Anything else is refused, naming `field`.
 */
function chartsheetIndex(wb: WorkbookData, target: unknown, where: string, field: string): number {
  if (typeof target === "number" && Number.isInteger(target)) {
    return target >= 0 && target < wb._chartsheets.length ? target : -1;
  }
  if (typeof target === "string") {
    const key = sheetNameKey(target);
    return wb._chartsheets.findIndex(sheet => sheetNameKey(sheet.name) === key);
  }
  // A fractional position names no chartsheet, as anything else that is not one does.
  const model = (target as { _model?: unknown } | null | undefined)?._model;
  if (typeof model !== "object" || model === null) {
    argumentChecks(where).fail(
      field,
      "a chartsheet, its name or its 0-based position (a whole number)",
      target
    );
  }
  return wb._chartsheets.indexOf(model as ChartsheetModel);
}

function _getChartsheetModel(
  wb: WorkbookData,
  target: string | number | ChartsheetData,
  where = "Workbook.getChartsheet",
  field = "nameOrIndex"
): ChartsheetModel | undefined {
  return wb._chartsheets[chartsheetIndex(wb, target, where, field)];
}

export function renameChartsheet(
  wb: WorkbookData,
  nameOrIndex: string | number,
  name: string
): boolean {
  checkWorkbook(wb);
  const model = _getChartsheetModel(wb, nameOrIndex);
  if (!model) {
    return false;
  }
  const currentName = model.name;
  if (currentName === name) {
    return true;
  }
  model.name = "__documonster_pending_chartsheet_rename__";
  try {
    model.name = _validateChartsheetName(wb, name);
    return true;
  } catch (error) {
    model.name = currentName;
    throw error;
  }
}

/**
 * Copy a chartsheet — by handle, name or 0-based position — chart included, under `name` (by default its name followed
 * by ` Copy`), and return the copy, or `undefined` when there is no such chartsheet.
 *
 * @throws {ExcelWorksheetNameError} when `name` is taken, longer than 31 characters or not a name Excel accepts.
 */
export function copyChartsheet(
  wb: WorkbookData,
  chartsheet: string | number | ChartsheetData,
  name?: string
): ChartsheetData | undefined {
  checkWorkbook(wb);
  const source = _getChartsheetModel(wb, chartsheet, "Workbook.copyChartsheet", "chartsheet");
  if (!source) {
    return undefined;
  }
  // A name the caller gave is refused past 31 characters; the default this makes is shortened instead.
  refuseLongSheetName(name);
  const cloneName = validateSheetName(wb, name ?? `${source.name} Copy`);
  const clone: ChartsheetModel = {
    ...deepClone(source),
    id: _nextSheetId(wb),
    sheetNo: _nextChartsheetNo(wb),
    // New tab position — the clone goes to the tail of the tab
    // bar, matching Excel's "Duplicate" behaviour. Drop the
    // deep-cloned `orderNo` from the source.
    orderNo: _nextSheetOrderNo(wb),
    name: cloneName,
    drawingName: undefined,
    relationships: source.relationships ? deepClone(source.relationships) : undefined
  };
  if (source.chartNumber) {
    const entry = getChartEntry(wb, source.chartNumber);
    if (entry) {
      const chartNumber = nextChartNumber(wb);
      // Clone the entry with ALL metadata: rawData, modelSnapshot,
      // dirty, preferRawPatch, requireRawPatch, rels (per-entry),
      // userShapesXml. A freshly-created entry carrying only `model`
      // would lose Excel-authored user-shape overlays, the raw-patch
      // fast path, and any per-entry rels that aren't in
      // `_chartRels`. Keeping them in lockstep means a clone of a
      // just-loaded chart matches the source byte-for-byte.
      addChartEntry(wb, cloneChartEntry(entry, chartNumber));
      copyChartSidecars(wb, source.chartNumber, chartNumber);
      clone.chartNumber = chartNumber;
      clone.chartExNumber = undefined;
    }
  } else if (source.chartExNumber) {
    const entry = getChartExStructuredEntry(wb, source.chartExNumber);
    const chartExNumber = nextChartExNumber(wb);
    if (entry) {
      // Same rationale as the classic branch — carry dirty /
      // preferRawPatch / requireRawPatch / rawData / modelSnapshot
      // across the clone so the raw-patch path keeps working on
      // the duplicate.
      addChartExStructuredEntry(wb, cloneChartExEntry(entry, chartExNumber));
    } else if (wb._chartExEntries[source.chartExNumber]) {
      wb._chartExEntries[chartExNumber] = wb._chartExEntries[source.chartExNumber].slice();
    }
    // Copy the chartEx sidecars (authored rels) so the cloned
    // chartsheet's XML references stay valid. Previously a chartEx
    // with `cx14:` / media rels on the source lost every relationship
    // on the clone.
    copyChartExSidecars(wb, source.chartExNumber, chartExNumber);
    clone.chartExNumber = chartExNumber;
    clone.chartNumber = undefined;
  }
  wb._chartsheets.push(clone);
  return createChartsheet(clone, wb);
}

/**
 * Replace the chart of a chartsheet — by handle, name or 0-based position — with one built from `chart`. Returns
 * `false` when there is no such chartsheet.
 */
export function replaceChartsheetChart(
  wb: WorkbookData,
  chartsheet: string | number | ChartsheetData,
  chart: AddChartsheetOptions["chart"]
): boolean {
  checkWorkbook(wb);
  const model = _getChartsheetModel(
    wb,
    chartsheet,
    "Workbook.replaceChartsheetChart",
    "chartsheet"
  );
  if (!model) {
    return false;
  }
  // Build the replacement first so a malformed options object throws
  // *before* we remove the existing chart entry. Without this, a
  // failed `buildChartExModel` / `buildChartModel` would leave the
  // chartsheet empty (old chart nuked, new chart never registered).
  let newChartExModel: ChartExModel | undefined;
  let newChartModel: ChartModel | undefined;
  if (isChartExOptions(chart)) {
    newChartExModel = chartArgument("Workbook.replaceChartsheetChart", "chart", () =>
      buildChartExModel(chart)
    );
  } else if (isComboChartOptions(chart)) {
    newChartModel = chartArgument("Workbook.replaceChartsheetChart", "chart", () =>
      buildComboChartModel(chart)
    );
  } else {
    newChartModel = chartArgument("Workbook.replaceChartsheetChart", "chart", () =>
      buildChartModel(chart)
    );
  }
  // Remove existing entries only after the new model builds cleanly.
  if (model.chartNumber) {
    releaseChartMedia(wb, model.chartNumber);
    removeChartEntry(wb, model.chartNumber);
    model.chartNumber = undefined;
  }
  if (model.chartExNumber) {
    removeChartExStructuredEntry(wb, model.chartExNumber);
    model.chartExNumber = undefined;
  }
  if (newChartExModel) {
    const chartExNumber = nextChartExNumber(wb);
    try {
      fillChartExCaches(newChartExModel, wb);
    } catch {
      // Cache population is best-effort; never let it break chart replacement.
    }
    addChartExStructuredEntry(wb, { chartExNumber, model: newChartExModel });
    model.chartExNumber = chartExNumber;
  } else if (newChartModel) {
    const chartNumber = nextChartNumber(wb);
    try {
      fillChartCaches(newChartModel, wb);
    } catch {
      // Cache population is best-effort; never let it break chart replacement.
    }
    const entry: ChartEntry = { chartNumber, model: newChartModel };
    // Resolve programmatic `series.spPr.fill.blip._pendingImage`
    // payloads — matches the classic `addChart` and `addChartsheet`
    // paths. Previously replacement via `replaceChartsheetChart`
    // silently dropped picture-fill payloads on the floor.
    try {
      resolvePendingChartImages(entry, wb, chartNumber);
    } catch {
      // Image resolution is best-effort; a broken image payload
      // should never take down chart replacement.
    }
    addChartEntry(wb, entry);
    _applyChartsheetSidecars(wb, chartNumber, chart);
    model.chartNumber = chartNumber;
  }
  return true;
}

/** Visit every worksheet in tab order, with its sheet id. */
export function eachWorksheet(
  wb: WorkbookData,
  callback: (sheet: Worksheet, id: number) => void
): void {
  checkWorkbook(wb);
  if (typeof callback !== "function") {
    argumentChecks("Workbook.eachWorksheet").fail("callback", "a function", callback);
  }
  getWorksheets(wb).forEach(sheet => {
    callback(sheet, sheet.id);
  });
}

/** The identity a threaded-comment author is registered with — see {@link registerPerson}. */
export interface RegisterPersonOptions {
  /** The author's id with their identity provider (an e-mail address, an AD object id, …). */
  userId?: string;
  /** The identity provider: `"AD"`, `"Windows Live"`, `"PeoplePicker"`, `"None"`, … */
  providerId?: string;
}

/**
 * Register a threaded-comment author and return their person id. Registering the same display name and
 * user id again returns the existing id.
 */
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  options?: RegisterPersonOptions
): string;
/**
 * @deprecated Pass the identity in the options object: `registerPerson(wb, name, { userId, providerId })`.
 * Removed in 1.0.
 */
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  userId?: string,
  providerId?: string
): string;
export function registerPerson(
  wb: WorkbookData,
  displayName: string,
  userIdOrOptions?: string | RegisterPersonOptions,
  legacyProviderId?: string
): string {
  checkWorkbook(wb);
  const { userId, providerId } = isRecordArgument(userIdOrOptions)
    ? userIdOrOptions
    : { userId: userIdOrOptions, providerId: legacyProviderId };
  const existing = wb._persons.find(p => p.displayName === displayName && p.userId === userId);
  if (existing) {
    return existing.id;
  }
  const id = `{${synthGuid()}}`;
  const entry: ThreadedCommentPerson = { id, displayName };
  if (userId !== undefined) {
    entry.userId = userId;
  }
  if (providerId !== undefined) {
    entry.providerId = providerId;
  }
  wb._persons.push(entry);
  return id;
}

/** Options for `Workbook.registerFunction`. */
export interface RegisterFunctionOptions {
  /** The fewest arguments the function takes. Defaults to 0. */
  minArity?: number;
  /** The most arguments the function takes. Defaults to 255, Excel's limit. */
  maxArity?: number;
  /** Recalculate on every calculation, as `NOW()` and `RAND()` are. */
  volatile?: boolean;
  /**
   * The function returns an array, so a formula calling it — `=F()`, `=F()*2` — is a dynamic array: the formula
   * engine spills it and the writers store it with dynamic-array metadata, as for `SEQUENCE`. Without it the function
   * returns one value, even when it shadows a built-in that returns an array: a formula calling it is ordinary, and an
   * array it returns all the same is intersected to its top-left value, as Excel computes an ordinary formula.
   * @default false
   */
  returnsArray?: boolean;
}

/**
 * Register a custom worksheet function the formula engine calls when a formula uses `name` (compared without
 * regard to case). A registered name shadows a built-in one of the same name.
 *
 * @throws {ExcelArgumentError} when `name` is not a non-empty string, `fn` is not a function, or an arity is not a
 * whole number from 0 to 255 with `minArity` at most `maxArity`.
 */
export function registerFunction(
  wb: WorkbookData,
  name: string,
  fn: (args: RuntimeValue[]) => RuntimeValue | number | string | boolean | null | undefined,
  options?: RegisterFunctionOptions
): void {
  checkWorkbook(wb);
  const check = argumentChecks("Workbook.registerFunction");
  check.required("name", name);
  check.string("name", name, { nonEmpty: true });
  if (typeof fn !== "function") {
    check.fail("fn", "a function", fn);
  }
  if (check.record("options", options)) {
    check.number("options.minArity", options.minArity, { min: 0, max: 255, integer: true });
    check.number("options.maxArity", options.maxArity, { min: 0, max: 255, integer: true });
    check.boolean("options.volatile", options.volatile);
    check.boolean("options.returnsArray", options.returnsArray);
    if ((options.minArity ?? 0) > (options.maxArity ?? 255)) {
      check.fail("options.minArity", "at most options.maxArity", options.minArity);
    }
  }
  if (!wb.userFunctions) {
    wb.userFunctions = new Map();
  }
  // Stored as the caller wrote it. Turning a plain `42` into a `RuntimeValue` here would need
  // `rvNumber` and friends, and importing them made every consumer of this module — which is every
  // consumer of the `Workbook` namespace — carry the formula runtime for a member most of them never
  // call. `core/formula-adapter.ts` converts on the way into the engine instead; see
  // {@link WorkbookFunctionDescriptor}.
  wb.userFunctions.set(name.toUpperCase(), {
    minArity: options?.minArity ?? 0,
    maxArity: options?.maxArity ?? 255,
    volatile: options?.volatile,
    ...(options?.returnsArray === undefined ? {} : { returnsArray: options.returnsArray }),
    invoke: fn as (args: never[]) => unknown
  });
}

/** Remove a function {@link registerFunction} registered. Returns `true` if there was one. */
export function unregisterFunction(wb: WorkbookData, name: string): boolean {
  checkWorkbook(wb);
  argumentChecks("Workbook.unregisterFunction").required("name", name);
  argumentChecks("Workbook.unregisterFunction").string("name", name);
  return wb.userFunctions?.delete(name.toUpperCase()) ?? false;
}

export function clearThemes(wb: WorkbookData): void {
  // Note: themes are not an exposed feature, meddle at your peril!
  wb._themes = undefined;
}

function _applyChartsheetSidecars(
  wb: WorkbookData,
  chartNumber: number,
  chartOptions: AddChartsheetOptions["chart"]
): void {
  if (isChartExOptions(chartOptions)) {
    return;
  }
  if (!chartOptions.chartStyle && !chartOptions.chartColors) {
    return;
  }
  if (chartOptions.chartStyle) {
    setChartStyle(
      wb,
      chartNumber,
      new TextEncoder().encode(buildChartStyle(chartOptions.chartStyle))
    );
  }
  if (chartOptions.chartColors) {
    setChartColors(
      wb,
      chartNumber,
      new TextEncoder().encode(buildChartColors(chartOptions.chartColors))
    );
  }
}

function _nextChartsheetNo(wb: WorkbookData): number {
  const existing = wb._chartsheets.map(cs => cs.sheetNo).filter(Number.isFinite);
  return existing.length > 0 ? Math.max(...existing) + 1 : 1;
}

function _nextSheetOrderNo(wb: WorkbookData): number {
  let max = -1;
  for (const ws of wb._worksheets) {
    if (ws && typeof ws.orderNo === "number" && ws.orderNo > max) {
      max = ws.orderNo;
    }
  }
  for (const cs of wb._chartsheets) {
    if (typeof cs.orderNo === "number" && cs.orderNo > max) {
      max = cs.orderNo;
    }
  }
  return max + 1;
}

/** The id for a new chartsheet — the one pool {@link getNextId} allocates from. */
export function _nextSheetId(wb: WorkbookData): number {
  return getNextId(wb);
}

export function _validateChartsheetName(wb: WorkbookData, name: string): string {
  refuseLongSheetName(name);
  return validateSheetName(wb, name);
}

/** What `Workbook.addExternalLink` takes: the linked workbook and, optionally, what this one caches of it. */
export interface AddExternalLinkOptions {
  /** The linked file — a bare name or a path relative to this workbook (`"data/ref.xlsx"`). */
  target: string;
  /** The linked workbook's sheet names, in its order. */
  sheetNames?: readonly string[];
  /** Cached cell values per linked sheet, as Excel stores them for display. */
  cachedValues?: ExternalLinkModel["cachedValues"];
  /** Whether `target` is outside the package (`"External"`, the default) or a part inside it. */
  targetMode?: ExternalLinkModel["targetMode"];
}

export function addExternalLink(
  wb: WorkbookData,
  input: AddExternalLinkOptions
): ExternalLinkModel {
  checkWorkbook(wb);
  const link: ExternalLinkModel = {
    index: wb.externalLinks.length + 1,
    target: input.target,
    targetMode: input.targetMode ?? "External",
    sheetNames: input.sheetNames ? [...input.sheetNames] : [],
    cachedValues: input.cachedValues ? { ...input.cachedValues } : {}
  };
  wb.externalLinks.push(link);
  return link;
}

export function getExternalLink(
  wb: WorkbookData,
  indexOrTarget: number | string
): ExternalLinkModel | undefined {
  checkWorkbook(wb);
  if (typeof indexOrTarget !== "number" && typeof indexOrTarget !== "string") {
    argumentChecks("Workbook.getExternalLink").fail(
      "indexOrTarget",
      "a 1-based index or a target",
      indexOrTarget
    );
  }
  if (typeof indexOrTarget === "number") {
    return wb.externalLinks[indexOrTarget - 1];
  }
  const lower = indexOrTarget.toLowerCase();
  return wb.externalLinks.find(link => link.target.toLowerCase() === lower);
}

export function _collectExternalLinksForWrite(wb: WorkbookData): ExternalLinkModel[] {
  // The links `addExternalLink` returns are the workbook's own, so an edit to one after it was checked is checked
  // here: an empty target is a relationship that points nowhere.
  wb.externalLinks.forEach((link, i) => {
    const check = argumentChecks("Workbook.toBuffer");
    check.required(`externalLinks[${i}].target`, link.target);
    check.string(`externalLinks[${i}].target`, link.target, { nonEmpty: true });
    check.oneOf(`externalLinks[${i}].targetMode`, link.targetMode, ["External", "Internal"]);
  });
  const userLower = new Set(wb.externalLinks.map(l => l.target.toLowerCase()));
  const combined: ExternalLinkModel[] = wb.externalLinks.map((link, i) => ({
    ...link,
    index: i + 1,
    sheetNames: [...(link.sheetNames ?? [])],
    cachedValues: { ...(link.cachedValues ?? {}) },
    targetMode: link.targetMode ?? "External"
  }));
  for (const cached of wb._writerExternalLinkCache.values()) {
    if (userLower.has(cached.target.toLowerCase())) {
      // User explicitly added a link with the same target after an
      // auto-discovery pass — prefer the user's definition, drop the
      // cached one.
      continue;
    }
    combined.push({
      ...cached,
      index: combined.length + 1,
      sheetNames: [...cached.sheetNames],
      cachedValues: { ...cached.cachedValues }
    });
  }
  return combined;
}

export function _recordAutoExternalLink(
  wb: WorkbookData,
  target: string,
  sheetName: string
): number {
  const lower = target.toLowerCase();
  // If the user explicitly declared a link with this target, we respect
  // their definition verbatim: no sheetName upserts, no cache entry.
  // Excel needs the user-declared sheetNames to match the refs, and
  // augmenting them on the user's behalf could silently hide a typo.
  const existingUserIdx = wb.externalLinks.findIndex(l => l.target.toLowerCase() === lower);
  if (existingUserIdx !== -1) {
    return existingUserIdx + 1;
  }
  let cached = wb._writerExternalLinkCache.get(lower);
  if (!cached) {
    cached = {
      // Index is provisional — the real on-disk index is recomputed by
      // `_collectExternalLinksForWrite()` at serialisation time.
      index: 0,
      target,
      targetMode: "External",
      sheetNames: [],
      cachedValues: {}
    };
    wb._writerExternalLinkCache.set(lower, cached);
  }
  if (sheetName && !cached.sheetNames.includes(sheetName)) {
    cached.sheetNames.push(sheetName);
  }
  // Recompute final index: user entries first, then cache entries in
  // insertion order. The caller needs the *on-disk* index so that the
  // formula it's rewriting matches the link that will be serialised.
  const userCount = wb.externalLinks.length;
  let cacheIdx = 0;
  for (const key of wb._writerExternalLinkCache.keys()) {
    cacheIdx++;
    if (key === lower) {
      return userCount + cacheIdx;
    }
  }
  // Unreachable — we just inserted.
  return userCount + wb._writerExternalLinkCache.size;
}

export function getDefaultFont(wb: WorkbookData): Partial<Font> | undefined {
  return wb._defaultFont;
}

/**
 * The id for a new sheet: one past the highest id the workbook holds or has removed. Worksheets and chartsheets
 * share one `sheetId` pool in `workbook.xml` (each must be unique across both), and an id is never handed out twice in
 * a session — a numeric id held across a remove does not come to name another sheet. That is Excel's rule: given
 * sheets 1, 2, 3, Excel 16.114 numbered a sheet added after deleting sheet 3 as 4, and one added after deleting sheet 1
 * as 5 (`tmp/office-tests/r6-sheetid`). Reusing the lowest free id used to hand a worksheet a removed chartsheet's.
 */
export function getNextId(wb: WorkbookData): number {
  let highest = wb._lastSheetId ?? 0;
  for (const ws of wb._worksheets) {
    if (ws && ws.id > highest) {
      highest = ws.id;
    }
  }
  for (const cs of wb._chartsheets) {
    if (cs && typeof cs.id === "number" && Number.isFinite(cs.id) && cs.id > highest) {
      highest = cs.id;
    }
  }
  return highest + 1;
}

export function getChartsheets(wb: WorkbookData): ChartsheetData[] {
  checkWorkbook(wb);
  return wb._chartsheets.map(model => createChartsheet(model, wb));
}

export function getPersons(wb: WorkbookData): ThreadedCommentPerson[] {
  return wb._persons;
}

/**
 * The threaded-comment authors registered on the workbook — by {@link registerPerson} or read from a file — in
 * registration order. They are copies: register an author with `registerPerson`.
 */
export function listPersons(wb: WorkbookData): ThreadedCommentPerson[] {
  checkWorkbook(wb);
  return wb._persons.map(person => ({ ...person }));
}

/**
 * The links to other workbooks this one declares, in order — `getExternalLink(wb, i)` is the `i`-th, counting from 1.
 * Like `getExternalLink` they are the workbook's own: an edit reaches the file, and is checked when it is written.
 */
export function listExternalLinks(wb: WorkbookData): ExternalLinkModel[] {
  checkWorkbook(wb);
  return [...wb.externalLinks];
}

// `getWorkbookModel` lives in `core/workbook-model.ts` — see the note there on why the model
// builder sits below this module rather than in it.
export { getWorkbookModel } from "@excel/core/workbook-model";

export function setDefaultFont(wb: WorkbookData, font: Partial<Font> | undefined): void {
  wb._defaultFont = font;
}

/**
 * Deep-clone workbook-model data.
 *
 * `structuredClone` unguarded: it exists in every supported environment, and
 * the `JSON.parse(JSON.stringify(...))` fallback this used to carry was a
 * lossier second implementation, not a fallback — it strips a typed-array
 * prototype, which the `Uint8Array` values on `ChartExModel.externalParts`
 * need. `chart/chart-handle.ts` had already dropped the same fallback.
 */
function deepClone<T>(value: T): T {
  return structuredClone(value);
}

/**
 * Deep-copy a {@link ChartEntry}, preserving every field that affects
 * write-time behaviour — rawData (for the raw-patch fast path),
 * modelSnapshot (change detection), dirty / preferRawPatch /
 * requireRawPatch (writer hints), style / colors (ancillary parts),
 * rels (per-entry relationship bag), and userShapesXml (annotation
 * overlay). The caller supplies the new `chartNumber`; everything
 * else is a structural clone so later mutations on one entry don't
 * leak into the other.
 *
 * `rawData` and `userShapesXml` are `Uint8Array`s — `.slice()` is
 * used instead of `structuredClone` to keep the fast path cheap.
 */
function cloneChartEntry(entry: ChartEntry, chartNumber: number): ChartEntry {
  return {
    chartNumber,
    model: deepClone(entry.model),
    ...(entry.rawData ? { rawData: entry.rawData.slice() } : {}),
    ...(entry.modelSnapshot !== undefined ? { modelSnapshot: entry.modelSnapshot } : {}),
    ...(entry.dirty !== undefined ? { dirty: entry.dirty } : {}),
    ...(entry.preferRawPatch !== undefined ? { preferRawPatch: entry.preferRawPatch } : {}),
    ...(entry.requireRawPatch !== undefined ? { requireRawPatch: entry.requireRawPatch } : {}),
    ...(entry.style ? { style: deepClone(entry.style) } : {}),
    ...(entry.colors ? { colors: deepClone(entry.colors) } : {}),
    ...(entry.rels ? { rels: entry.rels.map(r => ({ ...r })) } : {}),
    ...(entry.userShapesXml ? { userShapesXml: entry.userShapesXml.slice() } : {})
  };
}

/**
 * Deep-copy a {@link ChartExEntry}, preserving the same write-time
 * fields as {@link cloneChartEntry} but for the ChartEx family
 * (structured model + rawData + dirty / preferRawPatch /
 * requireRawPatch + rels).
 */
function cloneChartExEntry(entry: ChartExEntry, chartExNumber: number): ChartExEntry {
  return {
    chartExNumber,
    model: deepClone(entry.model),
    ...(entry.rawData ? { rawData: entry.rawData.slice() } : {}),
    ...(entry.modelSnapshot !== undefined ? { modelSnapshot: entry.modelSnapshot } : {}),
    ...(entry.dirty !== undefined ? { dirty: entry.dirty } : {}),
    ...(entry.preferRawPatch !== undefined ? { preferRawPatch: entry.preferRawPatch } : {}),
    ...(entry.requireRawPatch !== undefined ? { requireRawPatch: entry.requireRawPatch } : {}),
    ...(entry.rels ? { rels: entry.rels.map(r => ({ ...r })) } : {})
  };
}

export type Workbook = WorkbookData;

// Constructor forwards behind the deprecated `Workbook.createStreamWriter` /
// `createStreamReader`; use `new Stream.WorkbookWriter` / `new Stream.WorkbookReader`.
export function createStreamWriter(options?: WorkbookWriterOptions): WorkbookWriter {
  return new WorkbookWriter(options);
}

export function createStreamReader(
  input: WorkbookReaderInput,
  options?: WorkbookReaderOptions
): WorkbookReader {
  return new WorkbookReader(input, options);
}

// Re-export the workbook-core container layer so `@excel/workbook` stays the
// canonical import path.
export {
  type WorkbookData,
  type NamedStyleEntry,
  getDefinedNames,
  getChartEntry,
  addChartEntry,
  removeChartEntry,
  nextChartNumber,
  nextChartExNumber,
  addChartExStructuredEntry,
  getChartExStructuredEntry,
  removeChartExStructuredEntry,
  getImage,
  addWorkbookImage,
  validateSheetName,
  removeWorksheetEx,
  setChartStyle,
  setChartColors,
  getWorksheet,
  getWorksheets,
  copyChartSidecars,
  copyChartExSidecars,
  addCellStyle,
  getCellStyle,
  listCellStyles,
  removeCellStyle,
  useBuiltInCellStyle
} from "@excel/core/workbook-core";
