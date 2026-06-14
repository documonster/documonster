/**
 * `Cell` namespace surface — public `(ws, addr, …)` cell operations.
 *
 * Consumed as `import { Cell } from "documonster/excel"` → `Cell.setValue(ws, "A1", 42)`.
 * Each function resolves the cell handle via `getCell(ws, addr)` and delegates
 * to the internal handle-level helpers. Consumers never hold a `CellData`.
 *
 * This is a flat-named-export module; `excel/index.ts` re-exports it via
 * `export * as Cell`, which tree-shakes per-member on rolldown / rspack.
 */
import {
  type CellData,
  type CellModel,
  type CellValueInputType,
  type CellValueType,
  type FormulaResult,
  type NoteConfig,
  cellAddName,
  cellDataValidation,
  cellDisplayText,
  cellEffectiveType,
  cellFormula,
  cellGetModel,
  cellGetStyle,
  cellGetValue,
  cellHyperlink,
  cellIsMerged,
  cellMaster,
  cellNames,
  cellNote,
  cellRemoveName,
  cellResult,
  cellSetDataValidation,
  cellSetNote,
  cellSetStyle,
  cellSetValue,
  cellText,
  cellType
} from "@excel/cell";
import type { ValueType } from "@excel/enums";
import type { DataValidation, Style } from "@excel/types";
import { getCell } from "@excel/worksheet-core";
import type { WorksheetData } from "@excel/worksheet-core";

/** A worksheet handle (opaque to consumers). */
export type Sheet = WorksheetData;
/** An address: an `"A1"` string, or a 1-based `(row, col)` pair. */
type Addr = string | number;

// --- value / type / text ---

export function getValue(ws: Sheet, addr: Addr, col?: number): CellValueType {
  return cellGetValue(getCell(ws, addr, col));
}
export function setValue(
  ws: Sheet,
  addr: Addr,
  valueOrCol: CellValueInputType,
  value?: CellValueInputType
): void {
  // Supports both `(ws, "A1", value)` and `(ws, row, col, value)`.
  if (arguments.length >= 4) {
    cellSetValue(getCell(ws, addr, valueOrCol as number), value as CellValueInputType);
    return;
  }
  cellSetValue(getCell(ws, addr), valueOrCol);
}
export function getText(ws: Sheet, addr: Addr, col?: number): string {
  return cellText(getCell(ws, addr, col));
}
export function getDisplayText(ws: Sheet, addr: Addr, col?: number): string {
  return cellDisplayText(getCell(ws, addr, col));
}
export function getType(ws: Sheet, addr: Addr, col?: number): ValueType {
  return cellType(getCell(ws, addr, col));
}
export function getEffectiveType(ws: Sheet, addr: Addr, col?: number): ValueType {
  return cellEffectiveType(getCell(ws, addr, col));
}

// --- formula ---

export function getFormula(ws: Sheet, addr: Addr, col?: number): string | undefined {
  return cellFormula(getCell(ws, addr, col));
}
export function getResult(ws: Sheet, addr: Addr, col?: number): FormulaResult | undefined {
  return cellResult(getCell(ws, addr, col));
}

// --- style ---

export function getStyle(ws: Sheet, addr: Addr, col?: number): Partial<Style> {
  return cellGetStyle(getCell(ws, addr, col));
}
export function setStyle(
  ws: Sheet,
  addr: Addr,
  styleOrCol: Partial<Style> | number,
  style?: Partial<Style>
): void {
  // Supports `(ws, "A1", style)` and `(ws, row, col, style)`.
  if (arguments.length >= 4) {
    cellSetStyle(getCell(ws, addr, styleOrCol as number), style as Partial<Style>);
    return;
  }
  cellSetStyle(getCell(ws, addr), styleOrCol as Partial<Style>);
}

// --- merge ---

export function isMerged(ws: Sheet, addr: Addr): boolean {
  return cellIsMerged(getCell(ws, addr));
}
export function getMergeMaster(ws: Sheet, addr: Addr): CellData {
  return cellMaster(getCell(ws, addr));
}

// --- hyperlink ---

export function getHyperlink(ws: Sheet, addr: Addr): string | undefined {
  return cellHyperlink(getCell(ws, addr));
}

// --- note ---

export function getNote(ws: Sheet, addr: Addr): string | NoteConfig | undefined {
  return cellNote(getCell(ws, addr));
}
export function setNote(ws: Sheet, addr: Addr, note: string | NoteConfig): void {
  cellSetNote(getCell(ws, addr), note);
}

// --- defined names ---

export function getNames(ws: Sheet, addr: Addr): string[] {
  return cellNames(getCell(ws, addr));
}
export function addName(ws: Sheet, addr: Addr, name: string): void {
  cellAddName(getCell(ws, addr), name);
}
export function removeName(ws: Sheet, addr: Addr, name: string): void {
  cellRemoveName(getCell(ws, addr), name);
}

// --- data validation ---

export function getValidation(ws: Sheet, addr: Addr): DataValidation | undefined {
  return cellDataValidation(getCell(ws, addr));
}
export function setValidation(ws: Sheet, addr: Addr, value: DataValidation): void {
  cellSetDataValidation(getCell(ws, addr), value);
}

// --- model (advanced / round-trip) ---

export function getModel(ws: Sheet, addr: Addr): CellModel {
  return cellGetModel(getCell(ws, addr));
}
