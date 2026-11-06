import type { DataValidationRule } from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { slideFormula } from "@excel/utils/shared-formula";
import { deepEqual } from "@utils/object";

/**
 * Where a data validation's relative references are measured from, and how to move them.
 *
 * A file states a validation once for a whole `sqref`, its relative references measured from the top-left of the
 * bounding box of every range in it (`BrtDVal` in XLSB, `<dataValidation>` in XLSX).
 * The model keys a rule by each range on its own, so a rule in the model is relative to *its key's* top-left. Reading
 * `B2:B4 D6:D8` as two entries that both kept the formula stated for `B2` made `D6:D8` validate against the cells
 * beside `B2`; each entry has to be slid to its own corner on read, and back to the shared one on write.
 */

/**
 * The top-left cell, as an address, of the bounding box of `references` (`A1` or `A1:B2` each), or `undefined` when
 * one of them is not a cell or range.
 */
export function validationAnchor(references: readonly string[]): string | undefined {
  let row = Infinity;
  let col = Infinity;
  for (const reference of references) {
    for (const end of reference.replace(/\$/g, "").split(":")) {
      const m = /^([A-Za-z]{1,3})([1-9]\d*)$/.exec(end);
      if (m === null) {
        return undefined;
      }
      const column = colCache.decodeCol(m[1].toUpperCase());
      if (!(column >= 1 && column <= 16384) || Number(m[2]) > 1048576) {
        return undefined;
      }
      row = Math.min(row, Number(m[2]));
      col = Math.min(col, column);
    }
  }
  return row === Infinity ? undefined : colCache.encodeAddress(row, col);
}

/**
 * `rule` with its formulae moved from `from` to `to`, or `undefined` when a reference would leave the sheet.
 *
 * Returns `rule` itself when nothing moves, so a rule with no relative reference stays the one object it was.
 */
export function slideValidation<T extends DataValidationRule>(
  rule: T,
  from: string,
  to: string
): T | undefined {
  const formulae = (rule as { formulae?: unknown[] }).formulae;
  if (from === to || !Array.isArray(formulae)) {
    return rule;
  }
  let changed = false;
  const slid: unknown[] = [];
  for (const formula of formulae) {
    if (typeof formula !== "string") {
      slid.push(formula);
      continue;
    }
    const moved = slideFormula(formula, from, to);
    if (moved.includes("#REF!") && !formula.includes("#REF!")) {
      return undefined;
    }
    changed ||= moved !== formula;
    slid.push(moved);
  }
  return changed ? ({ ...rule, formulae: slid } as T) : rule;
}

/** Whether `a` at `anchorA` and `b` at `anchorB` are the same validation, once both are measured from `anchorA`. */
export function isSameValidation(
  a: DataValidationRule,
  anchorA: string,
  b: DataValidationRule,
  anchorB: string
): boolean {
  const moved = slideValidation(b, anchorB, anchorA);
  return moved !== undefined && deepEqual(a, moved);
}

/**
 * Validations that are one rule, merged into one entry over all their ranges.
 *
 * Each input's formulae are relative to the top-left of its own `ranges`; each output's are relative to the top-left
 * of its merged `ranges`, which is what a writer states for an `sqref`. Inputs are merged in order and only when the
 * result is exact — a rule that would need a reference moved off the sheet stays on its own.
 */
export function mergeValidations<T extends DataValidationRule>(
  entries: readonly { readonly ranges: readonly string[]; readonly rule: T }[]
): { ranges: string[]; rule: T }[] {
  const groups: { ranges: string[]; rule: T; anchor: string }[] = [];
  for (const entry of entries) {
    const anchor = validationAnchor(entry.ranges);
    if (anchor === undefined) {
      groups.push({ ranges: [...entry.ranges], rule: entry.rule, anchor: "" });
      continue;
    }
    let joined = false;
    for (const group of groups) {
      if (group.anchor === "") {
        continue;
      }
      const merged = validationAnchor([group.anchor, anchor])!;
      const groupRule = slideValidation(group.rule, group.anchor, merged);
      const entryRule = slideValidation(entry.rule, anchor, merged);
      if (groupRule === undefined || entryRule === undefined) {
        continue;
      }
      if (deepEqual(groupRule, entryRule)) {
        group.ranges.push(...entry.ranges);
        group.rule = groupRule;
        group.anchor = merged;
        joined = true;
        break;
      }
    }
    if (!joined) {
      groups.push({ ranges: [...entry.ranges], rule: entry.rule, anchor });
    }
  }
  return groups.map(({ ranges, rule }) => ({ ranges, rule }));
}
