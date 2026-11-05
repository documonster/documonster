import type { ConditionalFormattingTarget } from "@excel/core/worksheet";
import {
  addConditionalFormatting,
  clearConditionalFormattings,
  getConditionalFormattings,
  removeConditionalFormattingEntry
} from "@excel/core/worksheet";
/**
 * `ConditionalFormatting` namespace surface — a sheet's conditional-formatting rules.
 *
 * `import { ConditionalFormatting } from "documonster/excel"` →
 *   `const cf = ConditionalFormatting.add(ws, { ref: "A1:A10", rules })`,
 *   `ConditionalFormatting.list(ws)`, `ConditionalFormatting.remove(ws, cf)`,
 *   `ConditionalFormatting.clear(ws)`.
 */
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkListTarget } from "@excel/surface/argument-checks";
import type { ConditionalFormattingOptions } from "@excel/types";
import { argumentChecks, normalizeReference } from "@excel/utils/arguments";
import { checkConditionalFormatting } from "@excel/utils/conditional-formatting-checks";

/**
 * Add a block of conditional-formatting rules to the sheet, and return the block the sheet holds.
 *
 * `ref` is one or more ranges separated by spaces. Each rule is checked against its type — an unknown type, a
 * comparison without operands, a colour scale without its thresholds, a priority below 1 — and its `style` as a
 * cell style is, so a rule Excel would drop or repair the file for is refused here, before anything is added.
 *
 * The sheet keeps a copy of `cf`, so the same options can be added to several sheets and edited afterwards without
 * reaching what was added. The returned block, like the entries {@link list} returns, is the one the sheet writes and
 * moves when rows or columns are inserted; give it to {@link remove} to take it away.
 */
export function add(
  ws: WorksheetData,
  cf: ConditionalFormattingOptions
): ConditionalFormattingOptions {
  checkWorksheet(ws);
  checkConditionalFormatting("ConditionalFormatting.add", cf);
  // Stored canonical, as every other reference is: `a1:b2` was written as `sqref="a1:b2"` and read back differently
  // after an unrelated structural edit upper-cased it.
  return addConditionalFormatting(ws, { ...structuredClone(cf), ref: normalizeReference(cf.ref) });
}

/** The sheet's blocks of conditional-formatting rules, in the order they were added. */
export function list(ws: WorksheetData): ConditionalFormattingOptions[] {
  checkWorksheet(ws);
  return getConditionalFormattings(ws);
}

/**
 * Remove a block of rules — an entry {@link list} listed, its position in that list, its `ref` (`"A1:B9"`, compared as
 * Excel reads a reference, so `"a1:b9"` is the same), or the first entry a predicate accepts. Returns `true` if one was
 * removed.
 *
 * @throws {ExcelArgumentError} when `target` is none of those — a fractional position, `null`, a boolean.
 * @throws {ExcelInvalidAddressError} when a string `target` is not a reference.
 */
export function remove(ws: WorksheetData, target: ConditionalFormattingTarget | string): boolean {
  checkWorksheet(ws);
  if (typeof target === "string") {
    argumentChecks("ConditionalFormatting.remove").reference("target", target, {
      range: true,
      lines: true,
      list: true
    });
    const ref = normalizeReference(target);
    return removeConditionalFormattingEntry(ws, entry => normalizeReference(entry.ref) === ref);
  }
  if (typeof target !== "function") {
    checkListTarget("ConditionalFormatting.remove", "target", target, "a block of rules");
  }
  return removeConditionalFormattingEntry(ws, target);
}

/** Remove every block of rules from the sheet. */
export function clear(ws: WorksheetData): void {
  checkWorksheet(ws);
  clearConditionalFormattings(ws);
}

/** What `ConditionalFormatting.remove` takes: an entry, its position, or a predicate. */
export type { ConditionalFormattingTarget as Target } from "@excel/core/worksheet";
