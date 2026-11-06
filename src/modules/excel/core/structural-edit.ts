/**
 * Keeping a workbook consistent through a structural edit: rows or columns inserted or deleted, a sheet renamed, a
 * sheet removed.
 *
 * An edit moves cells, and everything that names cells by position has to follow them, on every sheet — or a formula
 * sums the wrong cells, a chart plots the wrong range and a table's part disagrees with its cells, and nothing says
 * so. {@link withReferenceEdit} is the single place that walks all of it; what a reference becomes is decided in
 * `utils/reference-rewrite`, by Excel's rules, so every holder of a reference agrees about it.
 *
 * What follows an edit, and what deliberately does not:
 *
 * - **Cell formulas** on every sheet — plain, shared (the master's text and range, and the address each member
 *   points at) and array or dynamic-array formulas with their ranges. A shared formula stays shared only while every
 *   member would still read as the master slid to it; otherwise each member gets its own text, which is what it
 *   computes.
 * - **Defined names** (reference and formula names), **tables** (their extent, header and totals rows, and the
 *   formulas their columns keep), **conditional formats** and **data validations** (ranges and formulas),
 *   **auto filters**, **print areas and titles**, **sparklines** (data, dates and location), **threaded comments**,
 *   **ignored-error ranges**, **chart series and titles** on any sheet and chartsheet, and **pivot table sources**.
 * - **Hyperlinks** to a place in the workbook are *not* rewritten. Excel stores the location as text and does not
 *   follow edits with it either, so rewriting it would make the two disagree.
 */
import { markChartEntryDirty } from "@excel/chart/chart-handle";
import type { ChartExEntry } from "@excel/chart/model/chart-ex-types";
import type { ChartEntry } from "@excel/chart/model/types";
import type { CellData, FormulaValueData } from "@excel/core/cell";
import { cellFormula, cellGetModel, cellGetValue, cellSetValue, cellType } from "@excel/core/cell";
import { definedNamesApplyEdit } from "@excel/core/defined-names";
import { Enums } from "@excel/core/enums";
import type { PivotTable } from "@excel/core/pivot-table";
import { splitPrintAreaInput } from "@excel/core/print-names";
import { sheetsInTabOrder } from "@excel/core/sheet-order";
import { moveSheetDrawings, rewriteControlLinks } from "@excel/core/structural-edit-drawings";
import {
  markTableRemoved,
  tableApplyLineEdit,
  tableDisplayName,
  tableName,
  tableRewriteFormulas,
  tableSurvivesLineEdit
} from "@excel/core/table";
import type { WorkbookData } from "@excel/core/workbook-core";
import { getWorksheets } from "@excel/core/workbook-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { getCell } from "@excel/core/worksheet-core";
import type { CellValue, ConditionalFormattingOptions, DataValidationRule } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import type { LineEdit, ReferenceEdit } from "@excel/utils/reference-rewrite";
import {
  editLine,
  REF_ERROR,
  rewriteArea,
  rewriteFormula,
  rewriteSqref
} from "@excel/utils/reference-rewrite";
import { slideFormula } from "@excel/utils/shared-formula";
import { excelNameKey, sheetNameKey } from "@utils/name-key";

/** What {@link withReferenceEdit} needs to know about the cells `apply` replaces rather than moves. */
export interface ReferenceEditOptions {
  /**
   * How many lines just before `edit.at` the splice overwrites in place — a splice that both removes and inserts
   * writes its new values over them (see `withLineEdit` in `core/worksheet`), so nothing that sat there is carried.
   */
  readonly overwritten?: number;
}

/**
 * Apply `edit` to the workbook: rewrite every cell formula, call `apply` — which moves the cells, renames the sheet or
 * removes it — then bring everything else into line with the result.
 *
 * Formulas are rewritten *before* the cells move because a shared formula's members are read relative to their own
 * positions, which the move changes; the text a reference becomes does not depend on where its cell sits.
 *
 * Which tables the edit removes is decided first too, so that a structured reference to one becomes `#REF!` in every
 * formula, as in Excel.
 */
export function withReferenceEdit(
  wb: WorkbookData,
  givenEdit: ReferenceEdit,
  apply: () => void,
  options: ReferenceEditOptions = {}
): void {
  const edit = withRemovedTables(wb, givenEdit);
  const { before, after } = planCellWrites(wb, edit, options);
  for (const [cell, value] of before) {
    cellSetValue(cell, value);
  }
  apply();
  for (const [ws, address, value] of after) {
    cellSetValue(getCell(ws, address), value);
  }
  definedNamesApplyEdit(wb._definedNames, edit);
  for (const ws of getWorksheets(wb)) {
    rewriteSheet(ws, edit);
  }
  rewriteCharts(wb, edit);
  rewritePivots(wb, edit);
}

/** `edit` with the tables it deletes: those a line edit leaves nothing of, or every table of a removed sheet. */
function withRemovedTables(wb: WorkbookData, edit: ReferenceEdit): ReferenceEdit {
  if (edit.kind === "rename") {
    return edit;
  }
  const removed = new Set<string>();
  for (const ws of getWorksheets(wb)) {
    if (!sameSheet(ws._name, edit.sheet)) {
      continue;
    }
    for (const table of Object.values(ws.tables ?? {})) {
      if (edit.kind === "removeSheet" || !tableSurvivesLineEdit(table, edit)) {
        removed.add(excelNameKey(tableName(table)));
        removed.add(excelNameKey(tableDisplayName(table)));
      }
    }
  }
  return removed.size === 0 ? edit : { ...edit, removedTables: removed };
}

/** Every sheet's name in tab order, chartsheets included — what a removal resolves 3-D references against. */
export function sheetNamesInTabOrder(wb: WorkbookData): string[] {
  type Named = { name: string; orderNo?: number; sheetNo?: number };
  const worksheets: Named[] = getWorksheets(wb).map(ws => ({
    name: ws._name,
    orderNo: ws.orderNo
  }));
  const chartsheets: Named[] = (wb._chartsheets ?? []).map(cs => ({
    name: cs.name,
    orderNo: cs.orderNo,
    sheetNo: cs.sheetNo
  }));
  return sheetsInTabOrder(worksheets, chartsheets).map(sheet => sheet.name);
}

function sameSheet(a: string, b: string): boolean {
  return a === b || sheetNameKey(a) === sheetNameKey(b);
}

/** The line edit when it is on `ws`, else `undefined`. */
function lineEditOn(ws: WorksheetData, edit: ReferenceEdit): LineEdit | undefined {
  return (edit.kind === "rows" || edit.kind === "columns") && sameSheet(ws._name, edit.sheet)
    ? edit
    : undefined;
}

// ============================================================================
// Cells
// ============================================================================

/** Where the cell at `row`, `col` goes, or `undefined` when the edit deletes it. */
function moveAddress(row: number, col: number, edit: LineEdit | undefined): string | undefined {
  if (edit === undefined) {
    return colCache.encodeAddress(row, col);
  }
  const moved = edit.kind === "rows" ? editLine(row, edit) : editLine(col, edit);
  if (moved === undefined) {
    return undefined;
  }
  return edit.kind === "rows"
    ? colCache.encodeAddress(moved, col)
    : colCache.encodeAddress(row, moved);
}

/** A formula value with some fields replaced and the ones set to `undefined` dropped. */
function formulaValue(value: FormulaValueData, changes: Partial<FormulaValueData>): CellValue {
  const out: Record<string, unknown> = { ...value, ...changes };
  for (const key of Object.keys(out)) {
    if (out[key] === undefined) {
      delete out[key];
    }
  }
  return out as unknown as CellValue;
}

/**
 * The new value of every formula cell the edit changes, decided before anything moves.
 *
 * A shared formula is planned as a group. Its members keep pointing at the master only if each one, rewritten on its
 * own, still equals the rewritten master slid to the member's new place — which holds for an edit that moves the
 * group uniformly or inserts inside it, but not for one that deletes the master or splits what the formula reads.
 */
interface PlannedWrites {
  /** Values written before the cells move, which the move then carries. */
  before: Array<[CellData, CellValue]>;
  /** Values written after the move, at the address the cell moved to. */
  after: Array<[WorksheetData, string, CellValue]>;
}

function planCellWrites(
  wb: WorkbookData,
  edit: ReferenceEdit,
  options: ReferenceEditOptions
): PlannedWrites {
  const writes: Array<[CellData, CellValue]> = [];
  const after: Array<[WorksheetData, string, CellValue]> = [];
  for (const ws of getWorksheets(wb)) {
    if (edit.kind === "removeSheet" && sameSheet(ws._name, edit.sheet)) {
      continue;
    }
    const host = ws._name;
    const local = lineEditOn(ws, edit);
    const formulas: FormulaCell[] = [];
    const members = new Map<string, FormulaCell[]>();
    for (const row of ws._rows) {
      for (const cell of row?.cells ?? []) {
        if (cell && cellType(cell) === Enums.ValueType.Hyperlink) {
          planHyperlinkFormula(ws, cell, edit, local, options, after);
          continue;
        }
        if (!cell || cellType(cell) !== Enums.ValueType.Formula) {
          continue;
        }
        const entry = { cell, value: cellGetValue(cell) as FormulaValueData };
        if (entry.value.sharedFormula !== undefined) {
          const group = members.get(entry.value.sharedFormula) ?? [];
          group.push(entry);
          members.set(entry.value.sharedFormula, group);
        } else if (entry.value.formula !== undefined) {
          formulas.push(entry);
        }
      }
    }
    for (const entry of formulas) {
      // Whatever its share type, a formula other cells point at is planned with them.
      const group = members.get(entry.cell.address);
      if (group !== undefined) {
        planSharedGroup(entry, group, edit, host, local, writes);
        continue;
      }
      const { value } = entry;
      const formula = rewriteFormula(value.formula!, edit, host);
      const ref = value.ref !== undefined && local ? rewriteArea(value.ref, local) : value.ref;
      if (formula !== value.formula || ref !== value.ref) {
        writes.push([entry.cell, formulaValue(value, { formula, ref })]);
      }
    }
  }
  return { before: writes, after };
}

/**
 * A hyperlink cell whose value is a formula (`HYPERLINK(…)` with its target): the formula is rewritten like any other,
 * and written after the move, which carries the formula as it was.
 */
function planHyperlinkFormula(
  ws: WorksheetData,
  cell: CellData,
  edit: ReferenceEdit,
  local: LineEdit | undefined,
  options: ReferenceEditOptions,
  after: Array<[WorksheetData, string, CellValue]>
): void {
  const model = cellGetModel(cell) as {
    formula?: unknown;
    result?: FormulaValueData["result"];
    hyperlink?: string;
    tooltip?: string;
  };
  if (typeof model.formula !== "string") {
    return;
  }
  const line = local?.kind === "rows" ? cell.row.number : cell.column.number;
  const overwritten = options.overwritten ?? 0;
  if (local && overwritten > 0 && line >= local.at - overwritten && line < local.at) {
    return;
  }
  const address = moveAddress(cell.row.number, cell.column.number, local);
  if (address === undefined) {
    return;
  }
  const formula = rewriteFormula(model.formula, edit, ws._name);
  if (address === cell.address && formula === model.formula) {
    return;
  }
  const value: Record<string, unknown> = { formula, hyperlink: model.hyperlink ?? "" };
  if (model.result !== undefined) {
    value.result = model.result;
  }
  if (model.tooltip !== undefined) {
    value.tooltip = model.tooltip;
  }
  after.push([ws, address, value as unknown as CellValue]);
}

interface FormulaCell {
  cell: CellData;
  value: FormulaValueData;
}

function planSharedGroup(
  master: FormulaCell,
  members: FormulaCell[],
  edit: ReferenceEdit,
  host: string,
  local: LineEdit | undefined,
  writes: Array<[CellData, CellValue]>
): void {
  const text = master.value.formula!;
  const formula = rewriteFormula(text, edit, host);
  const masterAddress = moveAddress(master.cell.row.number, master.cell.column.number, local);
  const ref =
    master.value.ref !== undefined && local
      ? rewriteArea(master.value.ref, local)
      : master.value.ref;
  const planned = members.map(member => ({
    member,
    formula: rewriteFormula(cellFormula(member.cell) ?? "", edit, host),
    address: moveAddress(member.cell.row.number, member.cell.column.number, local)
  }));
  const keepShared =
    masterAddress !== undefined &&
    (master.value.ref === undefined || ref !== undefined) &&
    planned.every(
      p => p.address === undefined || slideFormula(formula, masterAddress, p.address) === p.formula
    );
  if (keepShared) {
    if (formula !== text || ref !== master.value.ref) {
      writes.push([master.cell, formulaValue(master.value, { formula, ref })]);
    }
    if (masterAddress !== master.cell.address) {
      for (const p of planned) {
        writes.push([
          p.member.cell,
          formulaValue(p.member.value, { sharedFormula: masterAddress })
        ]);
      }
    }
    return;
  }
  writes.push([
    master.cell,
    formulaValue(
      master.value,
      master.value.shareType === "shared"
        ? { formula, ref: undefined, shareType: undefined }
        : { formula, ref }
    )
  ]);
  for (const p of planned) {
    writes.push([
      p.member.cell,
      formulaValue(p.member.value, { formula: p.formula, sharedFormula: undefined })
    ]);
  }
}

// ============================================================================
// A sheet's own parts
// ============================================================================

function rewriteSheet(ws: WorksheetData, edit: ReferenceEdit): void {
  const host = ws._name;
  const local = lineEditOn(ws, edit);
  const formula = (text: string): string => rewriteFormula(text, edit, host);

  rewriteTables(ws, edit, local);
  ws.conditionalFormattings = rewriteConditionalFormats(ws.conditionalFormattings, formula, local);
  rewriteValidations(ws, formula, local);
  rewriteSparklines(ws, formula, local);
  rewriteControlLinks(ws, formula);

  if (local === undefined) {
    return;
  }
  moveSheetDrawings(ws, local);
  rewriteAutoFilter(ws, local);
  const setup = ws.pageSetup as unknown as Record<string, string | undefined> | undefined;
  if (setup) {
    for (const key of ["printArea", "printTitlesRow", "printTitlesColumn"]) {
      const value = setup[key];
      if (typeof value !== "string" || value === "") {
        continue;
      }
      const kept = splitPrintAreaInput(value)
        .map(part => rewriteFormula(part, local, host))
        .filter(part => part !== REF_ERROR);
      if (kept.length > 0) {
        setup[key] = kept.join("&&");
      } else {
        delete setup[key];
      }
    }
  }
  ws.threadedComments = ws.threadedComments.flatMap(entry => {
    const ref = rewriteArea(entry.ref, local);
    return ref === undefined ? [] : [{ ...entry, ref }];
  });
  ws.ignoredErrors = ws.ignoredErrors.flatMap(entry => {
    const ref = rewriteSqref(entry.ref, local);
    return ref === undefined ? [] : [{ ...entry, ref }];
  });
}

function rewriteTables(ws: WorksheetData, edit: ReferenceEdit, local: LineEdit | undefined): void {
  for (const [name, table] of Object.entries(ws.tables)) {
    tableRewriteFormulas(table, edit, ws._name);
    if (local !== undefined && !tableApplyLineEdit(table, local)) {
      delete ws.tables[name];
      ws._workbook._tableNames.delete(excelNameKey(name));
      markTableRemoved(table);
    }
  }
}

function rewriteConditionalFormats(
  formats: ConditionalFormattingOptions[],
  formula: (text: string) => string,
  local: LineEdit | undefined
): ConditionalFormattingOptions[] {
  const kept: ConditionalFormattingOptions[] = [];
  for (const format of formats) {
    const ref = local === undefined ? format.ref : rewriteSqref(format.ref, local);
    if (ref === undefined) {
      continue;
    }
    format.ref = ref;
    for (const rule of format.rules) {
      const withFormulae = rule as { formulae?: (string | number)[] };
      if (Array.isArray(withFormulae.formulae)) {
        withFormulae.formulae = withFormulae.formulae.map(f =>
          typeof f === "string" ? formula(f) : f
        );
      }
      const withCfvo = rule as { cfvo?: { type: string; value?: number | string }[] };
      for (const cfvo of withCfvo.cfvo ?? []) {
        if (cfvo.type === "formula" && typeof cfvo.value === "string") {
          cfvo.value = formula(cfvo.value);
        }
      }
    }
    kept.push(format);
  }
  return kept;
}

function rewriteValidationRule(
  rule: DataValidationRule,
  formula: (text: string) => string
): DataValidationRule {
  if (rule.type === "any" || !Array.isArray(rule.formulae)) {
    return rule;
  }
  const formulae = rule.formulae.map(f => (typeof f === "string" ? formula(f) : f));
  return formulae.every((f, i) => f === rule.formulae[i]) ? rule : { ...rule, formulae };
}

function rewriteValidations(
  ws: WorksheetData,
  formula: (text: string) => string,
  local: LineEdit | undefined
): void {
  const model = ws.dataValidations.model;
  const next: typeof model = {};
  // One rule object is often stored under many keys; it is rewritten once, so the keys keep sharing it.
  const rewritten = new Map<DataValidationRule, DataValidationRule>();
  for (const [key, rule] of Object.entries(model)) {
    if (rule === undefined) {
      continue;
    }
    let newRule = rewritten.get(rule);
    if (newRule === undefined) {
      newRule = rewriteValidationRule(rule, formula);
      rewritten.set(rule, newRule);
    }
    let newKey = key;
    if (local !== undefined) {
      const isRange = key.startsWith("range:");
      const ref = rewriteArea(isRange ? key.slice(6) : key, local);
      if (ref === undefined) {
        continue;
      }
      newKey = ref.includes(":") ? `range:${ref}` : ref;
    }
    next[newKey] = newRule;
  }
  ws.dataValidations.model = next;
}

function rewriteSparklines(
  ws: WorksheetData,
  formula: (text: string) => string,
  local: LineEdit | undefined
): void {
  ws._sparklineGroups = ws._sparklineGroups.flatMap(group => {
    const sparklines = group.sparklines.flatMap(sparkline => {
      const cellRef =
        local === undefined ? sparkline.cellRef : rewriteArea(sparkline.cellRef, local);
      return cellRef === undefined
        ? []
        : [{ ...sparkline, cellRef, dataRef: formula(sparkline.dataRef) }];
    });
    if (sparklines.length === 0) {
      return [];
    }
    group.sparklines = sparklines;
    if (group.dateAxis !== undefined) {
      group.dateAxis = formula(group.dateAxis);
    }
    return [group];
  });
}

function rewriteAutoFilter(ws: WorksheetData, edit: LineEdit): void {
  const filter = ws.autoFilter;
  if (!filter) {
    return;
  }
  let text: string;
  if (typeof filter === "string") {
    text = filter;
  } else {
    const end = (e: string | { row: number; col: number }): string =>
      typeof e === "string" ? e : colCache.encodeAddress(e.row, e.col);
    text = `${end(filter.from)}:${end(filter.to)}`;
  }
  const ref = rewriteArea(text, edit);
  if (ref === text) {
    return;
  }
  ws._sortStateXml = undefined;
  ws._sortStateAutoFilterRef = undefined;
  if (ref === undefined) {
    ws.autoFilter = null;
    ws._autoFilterCriteria = undefined;
    return;
  }
  if (typeof filter === "string") {
    ws.autoFilter = ref;
  } else {
    const [from, to = from] = ref.split(":");
    ws.autoFilter = { from, to };
  }
}

// ============================================================================
// Charts and pivot tables
// ============================================================================

/**
 * Rewrite every formula a chart model holds — series values, categories, names, titles, labels, error bars — through
 * `edit`. A chart's references are always qualified, so no host sheet is needed. `true` when anything changed.
 */
function rewriteChartModel(node: unknown, edit: ReferenceEdit): boolean {
  if (node === null || typeof node !== "object" || ArrayBuffer.isView(node)) {
    return false;
  }
  let changed = false;
  if (Array.isArray(node)) {
    for (const item of node) {
      changed = rewriteChartModel(item, edit) || changed;
    }
    return changed;
  }
  const record = node as Record<string, unknown>;
  for (const [key, value] of Object.entries(record)) {
    if ((key === "formula" || key === "strRef") && typeof value === "string") {
      const next = rewriteFormula(value, edit);
      if (next !== value) {
        record[key] = next;
        changed = true;
      }
    } else if (key !== "rawXml" && typeof value === "object") {
      changed = rewriteChartModel(value, edit) || changed;
    }
  }
  return changed;
}

function rewriteCharts(wb: WorkbookData, edit: ReferenceEdit): void {
  const entries: Array<ChartEntry | ChartExEntry> = [
    ...Object.values(wb._chartEntries ?? {}),
    ...Object.values(wb._chartExStructuredEntries ?? {})
  ];
  for (const entry of entries) {
    if (entry?.model && rewriteChartModel(entry.model, edit)) {
      // Patch the bytes the chart was read from where the writer can, so what it does not model survives.
      markChartEntryDirty(entry, true);
    }
  }
}

function rewritePivots(wb: WorkbookData, edit: ReferenceEdit): void {
  const pivots = new Set<PivotTable>(wb.pivotTables ?? []);
  for (const ws of getWorksheets(wb)) {
    for (const pivot of ws.pivotTables ?? []) {
      pivots.add(pivot);
    }
  }
  for (const pivot of pivots) {
    const cache = pivot.cacheDefinition;
    if (edit.kind === "rename") {
      if (pivot.worksheetName !== undefined && sameSheet(pivot.worksheetName, edit.from)) {
        pivot.worksheetName = edit.to;
      }
      if (pivot.source !== undefined && sameSheet(pivot.source.name, edit.from)) {
        pivot.source.name = edit.to;
      }
      if (cache?.sourceSheet !== undefined && sameSheet(cache.sourceSheet, edit.from)) {
        cache.sourceSheet = edit.to;
      }
      continue;
    }
    if (edit.kind !== "rows" && edit.kind !== "columns") {
      continue;
    }
    if (
      cache?.sourceSheet !== undefined &&
      cache.sourceRef !== undefined &&
      sameSheet(cache.sourceSheet, edit.sheet)
    ) {
      // A source the edit deletes entirely keeps its last range: the cache still holds its records, and Excel
      // reports the broken source when the table is refreshed.
      cache.sourceRef = rewriteArea(cache.sourceRef, edit) ?? cache.sourceRef;
    }
    if (
      pivot.ref !== undefined &&
      pivot.worksheetName !== undefined &&
      sameSheet(pivot.worksheetName, edit.sheet)
    ) {
      pivot.ref = rewriteArea(pivot.ref, edit) ?? pivot.ref;
    }
    const source = pivot.source;
    if (
      source?.ref !== undefined &&
      source.tableName === undefined &&
      sameSheet(source.name, edit.sheet)
    ) {
      rewritePivotSheetSource(pivot, source as typeof source & { ref: string }, edit);
    }
  }
}

/**
 * Move a worksheet-sourced pivot table's source range with a line edit of its sheet, as any other reference moves,
 * and renumber the fields a column edit shifts. A field whose column the edit deletes leaves the pivot table's axes;
 * a source the edit deletes entirely keeps its last range, as a read cache does.
 */
function rewritePivotSheetSource(pivot: PivotTable, source: { ref: string }, edit: LineEdit): void {
  const before = colCache.decode(source.ref);
  const after = rewriteArea(source.ref, edit);
  if (after === undefined) {
    return;
  }
  source.ref = after;
  if (edit.kind !== "columns" || !("top" in before)) {
    return;
  }
  const moved = colCache.decode(after);
  const newLeft = "top" in moved ? moved.left : moved.col;
  const remap = (index: number): number | undefined => {
    const col = editLine(before.left + index, edit);
    return col === undefined ? undefined : col - newLeft;
  };
  const keep = (indices: number[]): number[] =>
    indices.map(remap).filter((index): index is number => index !== undefined);
  pivot.rows = keep(pivot.rows);
  pivot.columns = keep(pivot.columns);
  if (pivot.pages) {
    pivot.pages = keep(pivot.pages);
  }
  const values: number[] = [];
  const metrics: typeof pivot.valueMetrics = [];
  pivot.values.forEach((index, i) => {
    const next = remap(index);
    if (next !== undefined) {
      values.push(next);
      metrics.push(pivot.valueMetrics[i]);
    }
  });
  pivot.values = values;
  pivot.valueMetrics = metrics;
}
