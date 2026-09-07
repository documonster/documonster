/**
 * Excel → formula snapshot/writeback adapter.
 *
 * The formula engine consumes an immutable workbook snapshot and emits a
 * declarative writeback plan. documonster's workbook/worksheet/cell values are
 * plain-data records, so this module captures their state through the flat
 * helpers and resolves live cells only during final writeback.
 *
 * Excel (layer 4) may import formula (layer 3); this is the sanctioned seam.
 */
import { captureFormulaSnapshot } from "@excel/core/formula-capture";
import { applyFormulaWriteback } from "@excel/core/formula-writeback";
import type { WorkbookFunctionDescriptor } from "@excel/core/workbook-core";
import type { WorkbookData } from "@excel/core/workbook.browser";
import type { FormulaFunction } from "@formula/integration/calculate-formulas";
import { calculateFormulas as calculateFormulasEngine } from "@formula/integration/calculate-formulas";
import { createFormulaCalculationState } from "@formula/integration/calculation-state";
import type { FormulaCalculationState } from "@formula/integration/calculation-state";
import type { RuntimeValue } from "@formula/runtime/values";
import { BLANK, rvBoolean, rvNumber, rvString } from "@formula/runtime/values";

const calculationStates = new WeakMap<WorkbookData, FormulaCalculationState>();

/**
 * Wrap what `Workbook.registerFunction` stored into what the engine calls.
 *
 * `registerFunction` accepts a function returning a `RuntimeValue` **or** a plain number, string,
 * boolean, `null` or `undefined`, because requiring callers to build engine values would make the
 * feature unusable without importing the engine. The conversion is here rather than at registration
 * time for a tree-shaking reason worth keeping: `rvNumber` and friends are the formula runtime, and
 * reaching for them from `core/workbook.browser.ts` charged every `Workbook` consumer for it. This
 * module is only reachable through `documonster/excel/formula`, which is where a caller has already
 * asked for the engine.
 */
function toEngineFunctions(
  registry: Map<string, WorkbookFunctionDescriptor> | undefined
): Map<string, FormulaFunction> | undefined {
  if (registry === undefined) {
    return undefined;
  }
  const wrapped = new Map<string, FormulaFunction>();
  for (const [name, fn] of registry) {
    wrapped.set(name, {
      minArity: fn.minArity,
      maxArity: fn.maxArity,
      ...(fn.volatile === undefined ? {} : { volatile: fn.volatile }),
      invoke: args => normalizeFunctionResult(fn.invoke(args as never[]))
    });
  }
  return wrapped;
}

/** A user function's return value as an engine value. `null`/`undefined` mean an empty cell. */
function normalizeFunctionResult(value: unknown): RuntimeValue {
  if (value === null || value === undefined) {
    return BLANK;
  }
  if (typeof value === "number") {
    return rvNumber(value);
  }
  if (typeof value === "string") {
    return rvString(value);
  }
  if (typeof value === "boolean") {
    return rvBoolean(value);
  }
  return value as RuntimeValue;
}

function stateFor(wb: WorkbookData): FormulaCalculationState {
  let state = calculationStates.get(wb);
  if (!state) {
    state = createFormulaCalculationState();
    calculationStates.set(wb, state);
  }
  return state;
}

/**
 * Recalculate all formulas in a workbook, mutating cached results in place.
 * Excel-side wrapper around the formula engine's snapshot/writeback boundary.
 */
export function calculateFormulas(wb: WorkbookData): void {
  const state = stateFor(wb);
  const snapshot = captureFormulaSnapshot(wb);
  // Capture this before evaluation. A custom function may unregister itself
  // after mutating the workbook; consulting the live registry afterwards
  // would incorrectly skip drift validation.
  const mayHaveMutated = (wb.userFunctions?.size ?? 0) > 0;
  const result = calculateFormulasEngine(snapshot, state, toEngineFunctions(wb.userFunctions));
  const { plan } = result;
  // The engine is pure, so the only way the workbook can change while a
  // calculation runs is a user-registered function writing to it. Re-capture
  // (and compare) only in that case; otherwise the plan is provably still
  // consistent with `snapshot` and a second full capture would be dead work.
  applyFormulaWriteback(wb, plan, snapshot, mayHaveMutated ? captureFormulaSnapshot(wb) : snapshot);
  calculationStates.set(wb, result.state);
}
