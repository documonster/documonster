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
 *
 * Each also takes a streaming worksheet writer (`Stream.WorksheetWriter`), whose rules are written when it is
 * committed: an edit after that is refused.
 */
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkListTarget } from "@excel/surface/argument-checks";
import type { StreamWorksheet } from "@excel/surface/stream-targets";
import { asStreamWorksheet } from "@excel/surface/stream-targets";
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
  ws: WorksheetData | StreamWorksheet,
  cf: ConditionalFormattingOptions
): ConditionalFormattingOptions {
  const writer = asStreamWorksheet(ws);
  if (writer) {
    writer._checkOpen("add conditional formatting");
    checkConditionalFormatting("ConditionalFormatting.add", cf);
    return writer._addConditionalFormatting(cf);
  }
  checkWorksheet(ws);
  checkConditionalFormatting("ConditionalFormatting.add", cf);
  // Stored canonical, as every other reference is: `a1:b2` was written as `sqref="a1:b2"` and read back differently
  // after an unrelated structural edit upper-cased it.
  return addConditionalFormatting(ws, { ...structuredClone(cf), ref: normalizeReference(cf.ref) });
}

/**
 * The sheet's blocks of conditional-formatting rules, in the order they were added: the blocks themselves, not copies,
 * which {@link remove} takes. An edit made to one directly is not checked as {@link add} checks a block.
 */
export function list(ws: WorksheetData | StreamWorksheet): ConditionalFormattingOptions[] {
  const writer = asStreamWorksheet(ws);
  if (writer) {
    return [...writer.conditionalFormatting];
  }
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
export function remove(
  ws: WorksheetData | StreamWorksheet,
  target: ConditionalFormattingTarget | string
): boolean {
  const writer = asStreamWorksheet(ws);
  if (!writer) {
    checkWorksheet(ws);
  }
  if (typeof target === "string") {
    argumentChecks("ConditionalFormatting.remove").reference("target", target, {
      range: true,
      lines: true,
      list: true
    });
    const ref = normalizeReference(target);
    return removeEntry(ws, entry => normalizeReference(entry.ref) === ref);
  }
  if (typeof target !== "function") {
    checkListTarget("ConditionalFormatting.remove", "target", target, "a block of rules");
  }
  return removeEntry(ws, target);
}

function removeEntry(
  ws: WorksheetData | StreamWorksheet,
  target: ConditionalFormattingTarget
): boolean {
  const writer = asStreamWorksheet(ws);
  if (!writer) {
    return removeConditionalFormattingEntry(ws as WorksheetData, target);
  }
  writer._checkOpen("remove conditional formatting");
  const entries = writer.conditionalFormatting;
  const index =
    typeof target === "function"
      ? entries.findIndex(target)
      : typeof target === "number"
        ? Number.isInteger(target) && target >= 0 && target < entries.length
          ? target
          : -1
        : entries.indexOf(target);
  if (index < 0) {
    return false;
  }
  entries.splice(index, 1);
  return true;
}

/** Remove every block of rules from the sheet. */
export function clear(ws: WorksheetData | StreamWorksheet): void {
  const writer = asStreamWorksheet(ws);
  if (writer) {
    writer._checkOpen("remove conditional formatting");
    writer.conditionalFormatting = [];
    return;
  }
  checkWorksheet(ws);
  clearConditionalFormattings(ws);
}

/** What `ConditionalFormatting.remove` takes: an entry, its position, or a predicate. */
export type { ConditionalFormattingTarget as Target } from "@excel/core/worksheet";
