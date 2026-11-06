import { definedNameListKey } from "@excel/utils/function-names";
import {
  bareFunctionName,
  isLambdaParameter,
  lambdaParameters,
  toFileFormula
} from "@excel/utils/future-functions";
import {
  etaStubName,
  futureFunctionStubName,
  hasNoXlsbForm,
  isFutureFunction,
  isUserFunction
} from "@excel/xlsb/formula/ptg";
import { NodeType, type AstNode } from "@formula/syntax/ast";
import { parseFormula } from "@formula/syntax/parser";

/** A defined name a formula needs before it can be encoded, in the shape `BrtName` writes. */
export interface FormulaNameStub {
  readonly name: string;
  readonly formula?: string;
  readonly hidden: boolean;
  /**
   * For a `_xlpm.` parameter, whether its `BrtName` sets `fCalcExp` — updated as later formulas are scanned, since the
   * record is written only once every formula has been seen. See `parameterDeclarations`.
   */
  calcExp?: boolean;
}

/**
 * The `_xlfn.*`, `_xludf.*` and `_xlpm.*` defined names a set of formulas will need.
 *
 * A function with no `Ftab` id — `XLOOKUP`, `TEXTJOIN`, anything newer than the enumeration — is called through a
 * `PtgName` naming a hidden stub, and a `LET`/`LAMBDA` parameter is read through a `PtgName` too. A `PtgName` is a
 * one-based index into the defined-name list, so every such name has to be in the list before the formula that
 * calls it is encoded.
 *
 * **One collector for both writers.** The buffered writer scans every formula in the model before encoding any; the
 * streaming writer scans each row as it is committed and appends what is new, because its rows are encoded before the
 * workbook part is written. Two scanners would be two answers to "which names does this formula need".
 *
 * The scan is a regular expression over the formula source rather than a parse, deliberately: a name that turns out
 * not to be needed costs one hidden defined name, while a missed one costs the formula. Existing names win — a
 * workbook that already defines `_xlfn.XLOOKUP` gets no second copy.
 */
export class FormulaNameCollector {
  private readonly seen: Set<string>;
  private readonly collected: FormulaNameStub[] = [];

  /**
   * `_xlpm.` name, by `definedNameListKey`, → whether any formula declares it outside a lambda handed to another function — the
   * `fCalcExp` bit of its `BrtName`; see {@link parameterDeclarations}. Kept for names already in the list too.
   */
  private readonly calcExp = new Map<string, boolean>();

  constructor(existing: readonly { readonly name: string }[] = []) {
    this.seen = new Set(existing.map(entry => definedNameListKey(entry.name)));
  }

  /** Treat these names as already defined, so a reference to one is not taken for a function passed by name. */
  know(names: Iterable<string>): void {
    for (const name of names) {
      this.seen.add(definedNameListKey(name));
    }
  }

  /** Whether the `BrtName` of the `_xlpm.` name `name` sets `fCalcExp`, from every formula scanned so far. */
  isCalcExpParameter(name: string): boolean {
    return this.calcExp.get(definedNameListKey(name)) ?? true;
  }

  /** Every name collected so far, in allocation order. */
  get stubs(): readonly FormulaNameStub[] {
    return this.collected;
  }

  /**
   * Scan a batch of formula sources, appending the names they need that are not yet known.
   *
   * Within a batch, functions come before user functions and those before parameters — the order Excel writes and the
   * order the buffered writer has always used, so a streamed workbook whose names all arrive in one batch carries the
   * identical list. Returns whether anything was added.
   */
  scan(sources: Iterable<unknown>): boolean {
    const etas = new Map<string, string>();
    const functions = new Map<string, string>();
    const userFunctions = new Map<string, string>();
    const parameters = new Map<string, string>();
    const want = (into: Map<string, string>, name: string): void => {
      if (!this.seen.has(definedNameListKey(name))) {
        into.set(definedNameListKey(name), name);
      }
    };
    for (const source of sources) {
      if (typeof source !== "string") {
        continue;
      }
      // The stored spelling, which is what the encoder writes: `FILTER` needs `_xlfn._xlws.FILTER`, and `LET(x,…)`
      // reads `_xlpm.x`.
      const formula = toFileFormula(source);
      for (const parameter of lambdaParameters(formula)) {
        want(parameters, parameter);
      }
      for (const name of etaNames(formula, this.seen)) {
        want(etas, name);
      }
      for (const [name, calcExp] of parameterDeclarations(formula)) {
        this.calcExp.set(name, (this.calcExp.get(name) ?? false) || calcExp);
      }
      // A function call in the source is a name followed by `(`. A name already carrying `_xlfn.` — how XLSX stores
      // such a call — needs its stub just the same; `futureFunctionStubName` normalises the prefix rather than adding
      // a second one.
      for (const match of formula.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) {
        const name = match[1]!;
        if (isUserFunction(name)) {
          want(userFunctions, name.slice("_xludf.".length));
          continue;
        }
        // A call with no XLSB form is refused by the encoder; its stub would still make Excel refuse the package.
        if (isLambdaParameter(name) || !isFutureFunction(name) || hasNoXlsbForm(name)) {
          continue;
        }
        want(functions, futureFunctionStubName(name));
      }
    }
    // `formula` is the text form of `PtgErr(#NAME?)`, which is the body Excel gives every stub. The record writer
    // re-encodes it, so the two agree by construction rather than by a second literal.
    // `_xleta.` names first: Excel lists names sorted, and they sort ahead of `_xlfn.`.
    const added: FormulaNameStub[] = [
      ...[...etas.values()].map(name => ({ name, formula: "#NAME?", hidden: true })),
      ...[...functions.values()].map(name => ({ name, formula: "#NAME?", hidden: true })),
      ...[...userFunctions.values()].map(name => ({ name, hidden: false })),
      ...[...parameters.values()].map(name => ({ name, hidden: false }))
    ];
    for (const stub of added) {
      this.seen.add(definedNameListKey(stub.name));
      this.collected.push(stub);
    }
    for (const stub of this.collected) {
      if (isLambdaParameter(stub.name)) {
        stub.calcExp = this.isCalcExpParameter(stub.name);
      }
    }
    return added.length > 0;
  }
}

/**
 * The `_xleta.` names `formula` needs: one for each function it passes by name — `SUM` in `GROUPBY(A2:A9,B2:B9,SUM)`,
 * or `_xleta.SUM` as a file spells it — by the rule {@link etaStubName} applies when the formula is encoded. `defined`
 * is the keys (`definedNameListKey`) of the names the workbook already has; a bare name among them is that name, not a function.
 */
function etaNames(formula: string, defined: ReadonlySet<string>): string[] {
  if (!/[\p{L}_]/u.test(formula)) {
    return [];
  }
  let root: AstNode;
  try {
    root = parseFormula(formula.startsWith("=") ? formula.slice(1) : formula);
  } catch {
    return [];
  }
  const found: string[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      value.forEach(visit);
      return;
    }
    if (typeof value !== "object" || value === null) {
      return;
    }
    const node = value as AstNode;
    if (node.type === NodeType.Name && node.sheet === undefined) {
      const eta = etaStubName(node.name, defined.has(definedNameListKey(node.name)));
      if (eta !== undefined) {
        found.push(eta);
      }
      return;
    }
    Object.values(node).forEach(visit);
  };
  visit(root);
  return found;
}

/**
 * Each `LET`/`LAMBDA` parameter `formula` declares (keyed by `definedNameListKey`, as stored with `_xlpm.`), and whether that declaration
 * sets `fCalcExp` on the parameter's `BrtName`.
 *
 * The rule is read off Excel's own XLSB, not off the specification: there, every `LET` name and every parameter of a
 * `LAMBDA` that is bound by `LET`, immediately invoked or the whole formula is `0x00020019`, and a parameter of a
 * `LAMBDA` handed straight to another function (`MAP`, `REDUCE`, `BYCOL`, `MAKEARRAY` …) is `0x00020009`. A name with
 * both kinds of declaration has the bit. A formula that does not parse is left out, and its names default to set.
 */
function parameterDeclarations(formula: string): Map<string, boolean> {
  const found = new Map<string, boolean>();
  let root: AstNode;
  try {
    root = parseFormula(formula.startsWith("=") ? formula.slice(1) : formula);
  } catch {
    return found;
  }
  const declare = (node: AstNode | undefined, calcExp: boolean): void => {
    if (node?.type === NodeType.Name && isLambdaParameter(node.name)) {
      const key = definedNameListKey(node.name);
      found.set(key, (found.get(key) ?? false) || calcExp);
    }
  };
  const visit = (node: AstNode, handedToFunction: boolean): void => {
    switch (node.type) {
      case NodeType.FunctionCall: {
        const name = bareFunctionName(node.name);
        if (name === "LET") {
          node.args.forEach((arg, index) => {
            if (index % 2 === 0 && index < node.args.length - 1) {
              declare(arg, true);
            }
          });
        } else if (name === "LAMBDA") {
          // Invoked, it is not handed to anything: `LAMBDA(x,x*2)(5)`.
          const handed = handedToFunction && node.invocations === undefined;
          node.args.slice(0, -1).forEach(arg => declare(arg, !handed));
        }
        // A lambda bound by `LET`, or returned from a `LAMBDA`, is not handed to a function.
        const handsOn = name !== "LET" && name !== "LAMBDA";
        node.args.forEach(arg => visit(arg, handsOn));
        for (const list of node.invocations ?? []) {
          list.forEach(arg => visit(arg, true));
        }
        return;
      }
      case NodeType.BinaryOp:
        visit(node.left, false);
        visit(node.right, false);
        return;
      case NodeType.UnaryOp:
      case NodeType.Percent:
        visit(node.operand, false);
        return;
      case NodeType.Array:
        node.rows.forEach(row => row.forEach(entry => visit(entry, false)));
        return;
      case NodeType.UnionRef:
        node.areas.forEach(area => visit(area, false));
        return;
      default:
        return;
    }
  };
  visit(root, false);
  return found;
}

/**
 * The names a streaming writer allocated as its rows arrived, handed to the package writer at the end.
 *
 * The stubs go into the defined-name list right after the workbook's own names and print names, which is where the
 * buffered writer puts them. `prefixes` records the signature of those leading names at every encoding that could
 * have referenced a stub, so the package writer can tell whether they changed after an index was written.
 */
export interface StreamedFormulaNames {
  readonly stubs: readonly FormulaNameStub[];
  readonly prefixes: ReadonlySet<string>;
}

/** A defined-name list's identity as far as `PtgName` indices are concerned: each name and its scope, in order. */
export function definedNamesSignature(
  names: readonly { readonly name: string; readonly localSheetId?: number }[]
): string {
  return JSON.stringify(names.map(entry => [entry.name, entry.localSheetId ?? null]));
}

/** Every formula source a worksheet model carries outside its cells: conditional formats and validations. */
export function* sheetLevelFormulas(worksheet: unknown): Iterable<unknown> {
  const sheet = worksheet as {
    conditionalFormattings?: readonly { rules?: readonly { formulae?: readonly unknown[] }[] }[];
    dataValidations?: Record<string, { formulae?: readonly unknown[] } | undefined>;
  };
  for (const formatting of sheet.conditionalFormattings ?? []) {
    for (const rule of formatting.rules ?? []) {
      yield* rule.formulae ?? [];
    }
  }
  for (const validation of Object.values(sheet.dataValidations ?? {})) {
    yield* validation?.formulae ?? [];
  }
}

/** The formula of every cell in a list of row models. */
export function* rowFormulas(
  rows: readonly { readonly cells?: readonly unknown[] }[] | undefined
): Iterable<unknown> {
  for (const row of rows ?? []) {
    for (const cell of row.cells ?? []) {
      yield (cell as { formula?: unknown }).formula;
    }
  }
}
