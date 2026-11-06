/**
 * The conditional formats, data validations and defined names an XLSX write leaves out because a formula of theirs
 * cannot be written — see `unwritableFormulaReason` in `utils/array-evaluation.ts`, which is handed in as `check`.
 *
 * Handed in rather than imported because the check parses formulas, and the buffered writer loads the parser only when
 * a workbook has formulas. The XLSB writer needs none of this: its formula encoder refuses such a formula itself, and
 * each of these records already reports a formula that does not encode.
 *
 * Each function returns a new collection and leaves the one it was given alone: it may be the workbook's own.
 */

export type FormulaCheck = (formula: string) => string | undefined;

function reasonOf(
  formulae: readonly unknown[] | undefined,
  check: FormulaCheck
): string | undefined {
  for (const formula of formulae ?? []) {
    if (typeof formula === "string") {
      const reason = check(formula.startsWith("=") ? formula.slice(1) : formula);
      if (reason !== undefined) {
        return reason;
      }
    }
  }
  return undefined;
}

/**
 * `formattings` without the rules whose formulas cannot be written, each reported as `<prefix>range: …` — the prefix
 * being `Sheet!`, or nothing where the caller adds the sheet itself.
 */
export function writableConditionalFormattings<
  T extends {
    readonly ref?: string;
    readonly rules?: readonly unknown[];
  }
>(formattings: readonly T[], check: FormulaCheck, prefix: string, lost: string[]): T[] {
  return formattings.flatMap(block => {
    const rules = (block.rules ?? []).filter(rule => {
      const reason = reasonOf((rule as { formulae?: readonly unknown[] }).formulae, check);
      if (reason !== undefined) {
        lost.push(`${prefix}${block.ref}: conditional formatting ${reason}`);
      }
      return reason === undefined;
    });
    if (rules.length === (block.rules ?? []).length) {
      return [block];
    }
    return rules.length === 0 ? [] : [{ ...block, rules }];
  });
}

/** A data-validation model (rules by address) without the rules whose formulas cannot be written. */
export function writableDataValidations<T extends object>(
  validations: Readonly<Record<string, T | undefined>>,
  check: FormulaCheck,
  prefix: string,
  lost: string[]
): Record<string, T | undefined> {
  const kept: Record<string, T | undefined> = {};
  for (const [address, rule] of Object.entries(validations)) {
    const reason =
      rule === undefined
        ? undefined
        : reasonOf((rule as { formulae?: readonly unknown[] }).formulae, check);
    if (reason === undefined) {
      kept[address] = rule;
    } else {
      lost.push(`${prefix}${address.replace(/^range:/, "")}: data validation ${reason}`);
    }
  }
  return kept;
}

/** Defined names without those whose definition cannot be written, each reported as `Name: …`. */
export function writableDefinedNames<
  T extends {
    readonly name: string;
    readonly formulaExpression?: string;
    readonly rawText?: string;
  }
>(names: readonly T[], check: FormulaCheck, lost: string[]): T[] {
  return names.filter(name => {
    const reason = reasonOf([name.formulaExpression ?? name.rawText], check);
    if (reason !== undefined) {
      lost.push(`${name.name}: defined name ${reason}`);
    }
    return reason === undefined;
  });
}
