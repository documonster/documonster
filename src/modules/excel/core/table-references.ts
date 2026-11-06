import { sanitizeTableName } from "@excel/core/table-name";
import type { WorkbookModel } from "@excel/core/workbook-model";

/**
 * Structured references written to a table under the name it was *requested* under, restated under the name it has.
 *
 * A table a file names `T2` is stored as `_T2`, because Excel refuses a table whose name reads as a cell address
 * (see `sanitizeTableName`; `Table.add` refuses such a name rather than storing it changed). Lookups inside the library already accept either spelling, but a formula still saying
 * `T2[x]` names no table in the file, and Excel refuses the workbook. So every formula the model carries is
 * restated here, at the one point both writers take their model from.
 *
 * Only a name that is not itself a table, and whose sanitised form is one, is rewritten: such a reference cannot
 * mean anything else, and a formula that already names the stored table is left byte-for-byte alone.
 *
 * The XLSX reader applies it to the model it read, for a file whose table is named `T2`: the workbook stores that
 * table as `_T2`, and the formulas read with it say `_T2[x]` too, so the workbook a caller holds is the one the next
 * write produces. (The XLSB reader reaches the same result through the `getWorkbookModel` it loads the workbook from.)
 */
export function restateRenamedTableReferences(model: WorkbookModel): void {
  const tables = new Set<string>();
  for (const sheet of model.worksheets ?? []) {
    for (const table of sheet.tables ?? []) {
      // Under the name it is stored as. A table model the library built is sanitised already; one just read from a
      // file may still carry the name the file gave it (`T2`), which the workbook stores as `_T2` — and its formulas,
      // still saying `T2[x]`, are restated to match.
      for (const name of [table.name, table.displayName]) {
        if (typeof name === "string" && name !== "") {
          tables.add(sanitizeTableName(name).toLowerCase());
        }
      }
    }
  }
  if (tables.size === 0) {
    return;
  }
  const restate = (formula: string): string => restateFormula(formula, tables);
  const restateList = <T>(list: T[] | undefined): T[] | undefined => {
    if (!Array.isArray(list)) {
      return list;
    }
    let changed = false;
    const out = list.map(item => {
      if (typeof item !== "string") {
        return item;
      }
      const next = restate(item);
      changed ||= next !== item;
      return next as T;
    });
    return changed ? out : list;
  };

  model.worksheets = model.worksheets.map(sheet => {
    let next = sheet;
    const rows = mapChanged(sheet.rows, row => {
      const cells = mapChanged(row.cells, cell => {
        const formula = (cell as { formula?: unknown }).formula;
        if (typeof formula !== "string") {
          return cell;
        }
        const restated = restate(formula);
        return restated === formula ? cell : { ...cell, formula: restated };
      });
      return cells === row.cells ? row : { ...row, cells };
    });
    if (rows !== sheet.rows) {
      next = { ...next, rows };
    }
    const formattings = mapChanged(sheet.conditionalFormattings, formatting => {
      const rules = mapChanged(formatting.rules, rule => {
        const formulae = (rule as { formulae?: unknown[] }).formulae;
        const restated = restateList(formulae);
        return restated === formulae ? rule : ({ ...rule, formulae: restated } as typeof rule);
      });
      return rules === formatting.rules ? formatting : { ...formatting, rules };
    });
    if (formattings !== sheet.conditionalFormattings) {
      next = { ...next, conditionalFormattings: formattings };
    }
    if (sheet.dataValidations) {
      let validations: typeof sheet.dataValidations | undefined;
      for (const [key, rule] of Object.entries(sheet.dataValidations)) {
        const formulae = (rule as { formulae?: unknown[] } | undefined)?.formulae;
        const restated = restateList(formulae);
        if (restated !== formulae) {
          validations ??= { ...sheet.dataValidations };
          validations[key] = { ...rule!, formulae: restated } as typeof rule;
        }
      }
      if (validations !== undefined) {
        next = { ...next, dataValidations: validations };
      }
    }
    const sheetTables = mapChanged(sheet.tables, table => {
      const columns = mapChanged(table.columns, column => {
        const calculated = column.calculatedColumnFormula;
        const totals = column.totalsRowFormula;
        const nextCalculated = typeof calculated === "string" ? restate(calculated) : calculated;
        const nextTotals = typeof totals === "string" ? restate(totals) : totals;
        return nextCalculated === calculated && nextTotals === totals
          ? column
          : { ...column, calculatedColumnFormula: nextCalculated, totalsRowFormula: nextTotals };
      });
      return columns === table.columns ? table : { ...table, columns };
    });
    if (sheetTables !== sheet.tables) {
      next = { ...next, tables: sheetTables };
    }
    return next;
  });

  model.definedNames = mapChanged(model.definedNames, name => {
    const ranges = restateList(name.ranges);
    const expression =
      typeof name.formulaExpression === "string"
        ? restate(name.formulaExpression)
        : name.formulaExpression;
    return ranges === name.ranges && expression === name.formulaExpression
      ? name
      : { ...name, ranges: ranges!, formulaExpression: expression };
  });
}

/** `list` with `fn` applied, or `list` itself when `fn` returned every item unchanged. */
function mapChanged<T>(list: T[], fn: (item: T) => T): T[];
function mapChanged<T>(list: T[] | undefined, fn: (item: T) => T): T[] | undefined;
function mapChanged<T>(list: T[] | undefined, fn: (item: T) => T): T[] | undefined {
  if (!Array.isArray(list)) {
    return list;
  }
  let changed = false;
  const out = list.map(item => {
    const next = fn(item);
    changed ||= next !== item;
    return next;
  });
  return changed ? out : list;
}

const NAME_CHAR = /[\p{L}\p{N}_.\\]/u;

/**
 * `formula` with each `Name[` whose `Name` is not a table but sanitises to one restated under the sanitised name.
 * String literals and quoted sheet names are skipped, and so is the inside of every bracket, which is where a
 * structured reference keeps its column names.
 */
export function restateFormula(formula: string, tables: ReadonlySet<string>): string {
  if (!formula.includes("[")) {
    return formula;
  }
  let out = "";
  let last = 0;
  let i = 0;
  while (i < formula.length) {
    const c = formula[i];
    if (c === '"' || c === "'") {
      i = skipQuoted(formula, i, c);
      continue;
    }
    if (c !== "[") {
      i++;
      continue;
    }
    let start = i;
    while (start > 0 && NAME_CHAR.test(formula[start - 1])) {
      start--;
    }
    if (start < i) {
      const name = formula.slice(start, i);
      const lower = name.toLowerCase();
      if (!tables.has(lower)) {
        const stored = sanitizeTableName(name);
        if (stored !== name && tables.has(stored.toLowerCase())) {
          out += formula.slice(last, start) + stored;
          last = i;
        }
      }
    }
    i = skipBrackets(formula, i);
  }
  return last === 0 ? formula : out + formula.slice(last);
}

function skipQuoted(formula: string, start: number, quote: string): number {
  let i = start + 1;
  while (i < formula.length) {
    if (formula[i] === quote) {
      if (formula[i + 1] !== quote) {
        return i + 1;
      }
      i++;
    }
    i++;
  }
  return i;
}

function skipBrackets(formula: string, start: number): number {
  let depth = 0;
  let i = start;
  while (i < formula.length) {
    const c = formula[i];
    if (c === "'") {
      i += 2;
      continue;
    }
    if (c === "[") {
      depth++;
    } else if (c === "]" && --depth === 0) {
      return i + 1;
    }
    i++;
  }
  return i;
}
