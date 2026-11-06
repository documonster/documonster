/**
 * Whether a formula has to be stored as a dynamic array for Excel to compute what this library computes.
 *
 * **Two meanings of one formula.** This library evaluates a cell's formula the way Excel 365 does when it is typed:
 * `MEDIAN(IF(T[Group]="A",T[Response]))` filters the column and takes the median of what is left. A formula stored
 * *without* dynamic-array metadata is a legacy formula, and Excel evaluates a legacy formula the way Excel 2016 did —
 * a multi-cell range in a place that wants one value is *implicitly intersected* with the formula's row or column. The
 * same text then compares one cell of `T[Group]` instead of all of them; Excel shows it as `IF(@T[Group]="A",…)`, and
 * in the reported workbook computed 130.01 where the library had computed 100.135. Excel 365 itself avoids this by
 * storing such a formula as a dynamic array (`t="array"` on its own cell, `cm` naming the `XLDAPR` metadata), which
 * every version evaluates over arrays. The writers do the same with the formulas this answers `true` for.
 *
 * **Where legacy Excel intersects** is where it puts a reference in the value class: the operand-class assignment the
 * XLSB encoder writes, from the parameter classes of `@formula/syntax/function-table` (`implicitIntersections` in
 * `xlsb/formula/operand-classes.ts`). A range handed straight to a reference parameter is fine (`SUM(A1:A3)`), and so
 * is one below a parameter that converts to arrays (`SUMPRODUCT(A1:A3*B1:B3)`, `ROWS(A1:A3+0)`); one that reaches a
 * value position otherwise — an operator's operand, a value parameter, `IF`'s condition — is intersected.
 *
 * **Which intersections change the answer.** One *below a parameter that consumes a range whole* — `MEDIAN`, `SUM`,
 * `CORREL` — does, because the library hands that function the whole array and legacy Excel hands it one cell. Those
 * are the formulas this answers `true` for. One on the path from the root changes it too (`ABS(A1:A3)` is the cell in
 * the formula's row, or `#VALUE!` outside the range), but storing such a formula as a dynamic array would make it
 * spill, which the library does not do for a formula it was not told is dynamic.
 *
 * **What the library computes is what Excel computes for the file.** A formula stored, or to be stored, as an ordinary
 * one — read from a file that way, or set through the API and not answered `true` here — is evaluated with Excel's
 * intersections made explicit ({@link ordinaryEvaluationFormula}). One read from a file is written back as it came —
 * see `stored-formulas.ts` — so a workbook read and written keeps the values Excel shows for it.
 *
 * **When it is not answered `true` even so.** A formula that also has a multi-cell range on the path from its root
 * would spill as a dynamic array, so it is left ordinary and evaluated with Excel's intersections. Ranges are recognised
 * syntactically here — areas, whole rows and columns and multi-cell structured references. A defined name or a
 * function returning a reference (`OFFSET`) is not: its extent is not known, so it neither asks for a dynamic array nor
 * prevents one. It is still intersected where Excel intersects it.
 *
 * **A formula that produces an array is a dynamic array.** `=SEQUENCE(3)`, `={1,2,3}`, `=FILTER(…)*2`, `=A1#*2` —
 * whatever `producesArray` (`@formula/syntax/array-result`) answers `true` for, which is the question the engine asks
 * too. Stored as an ordinary formula, Excel would compute it with implicit intersection — its top-left value, and an
 * `A1#` on it is `#REF!` — while this library spilled it; so it is written as a dynamic array and evaluated as one
 * ({@link formulaEvaluation}), and the engine spills a formula set through the API exactly when the writers store it as
 * a dynamic array. Whatever ranges sit beside the array on that path spill with it, as in Excel. A spill reference
 * `A1#` (`ANCHORARRAY(A1)`) is also a range, intersected where any range is.
 */
import { bareFunctionName } from "@excel/utils/function-names";
import { implicitIntersections, nonReferenceArgument } from "@excel/xlsb/formula/operand-classes";
import { producesArray, type RegisteredFunctions } from "@formula/syntax/array-result";
import { NodeType, type AstNode, type StructuredRefNode } from "@formula/syntax/ast";
import {
  FUNCTION_ID_BY_NAME,
  functionClasses,
  futureFunctionClasses,
  paramClassAt,
  REFERENCE_ONLY_FUNCTIONS
} from "@formula/syntax/function-table";
import { parse } from "@formula/syntax/parser";
import { printAst } from "@formula/syntax/print";
import { SPILL_OPERATOR } from "@formula/syntax/token-types";
import { tokenize } from "@formula/syntax/tokenizer";
import { escapeRegExp } from "@utils/regexp";

interface Analysis {
  /** What Excel intersects when the formula is stored as an ordinary one — see `implicitIntersections`. */
  readonly intersected: ReadonlySet<AstNode>;
  /** A range Excel intersects sits below a parameter that consumes it whole, so intersecting it changes the answer. */
  readonly changed: boolean;
  /** A multi-cell range sits on the path from the root, so the formula as a dynamic array could spill. */
  readonly rootRange: boolean;
  /** The formula produces an array (`producesArray`): dynamic-array Excel spills it, and so does the engine. */
  readonly producesArray: boolean;
}

/** Whether the analysis makes a formula set through the API a dynamic array. */
function isDynamicArray(analysis: Analysis): boolean {
  return analysis.producesArray || (analysis.changed && !analysis.rootRange);
}

function analyse(root: AstNode, functions: RegisteredFunctions | undefined): Analysis {
  const intersected = implicitIntersections(root);
  let changed = false;
  let rootRange = false;
  // `consumed`: below a parameter that takes its argument whole — a reference or array parameter.
  const visit = (node: AstNode, consumed: boolean): void => {
    if (isRange(node)) {
      if (!consumed) {
        // A spill reference on that path is an array the formula produces, not a range it intersects.
        if (!isSpillReference(node)) {
          rootRange = true;
        }
      } else if (intersected.has(node)) {
        changed = true;
      }
      return;
    }
    switch (node.type) {
      case NodeType.BinaryOp:
        visit(node.left, consumed);
        visit(node.right, consumed);
        return;
      case NodeType.UnaryOp:
        // `@` intersects on purpose, and this library does the same.
        if (node.op !== "@") {
          visit(node.operand, consumed);
        }
        return;
      case NodeType.Percent:
        visit(node.operand, consumed);
        return;
      case NodeType.UnionRef:
        for (const area of node.areas) {
          visit(area, consumed);
        }
        return;
      case NodeType.FunctionCall: {
        const name = node.name.toUpperCase();
        if (name === "_XLFN.SINGLE" || node.invocations !== undefined) {
          return;
        }
        const id = FUNCTION_ID_BY_NAME.get(name);
        const classes = id === undefined ? futureFunctionClasses(node.name) : functionClasses(id);
        node.args.forEach((arg, index) => {
          const param = paramClassAt(classes, index);
          const takesWhole = !param.valueType || param.conversion === "array";
          visit(arg, consumed || (takesWhole && !PASSES_THROUGH.has(name)));
        });
        return;
      }
      default:
        return;
    }
  };
  visit(root, false);
  return { intersected, changed, rootRange, producesArray: producesArray(root, functions) };
}

/** A call to a function with a reference-only parameter, so most formulas are never parsed for the check below. */
const CALLS_REFERENCE_ONLY = new RegExp(
  `(?:^|[^A-Z0-9_.]|_XLFN\\.)(?:${[...REFERENCE_ONLY_FUNCTIONS].map(escapeRegExp).join("|")})\\s*\\(`,
  "i"
);

/**
 * Why `formula` cannot be written — a constant or a computed value where only a reference is accepted, which makes a
 * workbook Excel will not open (see `nonReferenceArgument`) — as the reason the writers report, or `undefined`.
 *
 * The one check every writer asks, for every place a formula is stored: cells, shared and array formulas, conditional
 * formats, data validations and defined names, in both containers.
 */
export function unwritableFormulaReason(formula: string): string | undefined {
  if (!CALLS_REFERENCE_ONLY.test(formula)) {
    return undefined;
  }
  const root = parsed(formula);
  const reason = root === undefined ? undefined : nonReferenceArgument(root);
  return reason === undefined ? undefined : `formula (${reason})`;
}

function parsed(formula: string): AstNode | undefined {
  try {
    return parse(tokenize(formula));
  } catch {
    return undefined;
  }
}

/** Whether `formula`, written as an ordinary formula, would be computed differently by Excel — and can be stored as a
 * dynamic array instead without spilling where this library does not. */
export function needsArrayEvaluation(formula: string, functions?: RegisteredFunctions): boolean {
  const root = parsed(formula);
  if (root === undefined) {
    return false;
  }
  return isDynamicArray(analyse(root, functions));
}

/**
 * The text this library evaluates for a cell's ordinary (not array) formula, so it computes what Excel computes for
 * the formula as it is, or will be, stored: each operand Excel intersects is given an explicit `@`, which is how Excel
 * 365 displays it (`SUM(@A1:A5*2)`). `stored` says the formula came from a file as an ordinary one; one set through the
 * API is written as a dynamic array when {@link needsArrayEvaluation} says so, and is then evaluated over arrays as
 * written. `undefined` when the formula is evaluated as it is, which is almost always.
 *
 * The root itself is not marked: a cell intersects its formula's result anyway, and a reference result is intersected
 * the same way either side of that.
 */
export function ordinaryEvaluationFormula(formula: string, stored: boolean): string | undefined {
  return formulaEvaluation(formula, stored).formula;
}

/** How this library evaluates a cell's ordinary (not array) formula. */
export interface FormulaEvaluation {
  /** The text it evaluates — see {@link ordinaryEvaluationFormula}; `undefined` for the formula as written. */
  readonly formula: string | undefined;
  /**
   * Whether it is evaluated as a dynamic array, spilling an array result: a formula set through the API that the
   * writers store as one ({@link needsArrayEvaluation}). Never for one a file stored as an ordinary formula.
   */
  readonly dynamicArray: boolean;
}

const AS_WRITTEN: FormulaEvaluation = { formula: undefined, dynamicArray: false };

/** {@link ordinaryEvaluationFormula} and whether the formula spills, from one parse. */
export function formulaEvaluation(
  formula: string,
  stored: boolean,
  functions?: RegisteredFunctions
): FormulaEvaluation {
  const root = parsed(formula);
  if (root === undefined) {
    return AS_WRITTEN;
  }
  const analysis = analyse(root, functions);
  if (!stored && isDynamicArray(analysis)) {
    return { formula: undefined, dynamicArray: true };
  }
  const { intersected } = analysis;
  const marked = new Set(intersected);
  marked.delete(root);
  if (marked.size === 0) {
    return AS_WRITTEN;
  }
  const rebuild = (node: AstNode): AstNode => {
    if (marked.has(node)) {
      return { type: NodeType.UnaryOp, op: "@", operand: node };
    }
    switch (node.type) {
      case NodeType.BinaryOp:
        return { ...node, left: rebuild(node.left), right: rebuild(node.right) };
      case NodeType.UnaryOp:
        return node.op === "@" ? node : { ...node, operand: rebuild(node.operand) };
      case NodeType.Percent:
        return { ...node, operand: rebuild(node.operand) };
      case NodeType.UnionRef:
        return { ...node, areas: node.areas.map(rebuild) };
      case NodeType.FunctionCall:
        return { ...node, args: node.args.map(rebuild) };
      default:
        return node;
    }
  };
  return { formula: printAst(rebuild(root)), dynamicArray: false };
}

/**
 * Functions whose result is one of their reference arguments, unchanged: what reaches those arguments is on the path of
 * whatever consumes the call. Any other function taking a range whole consumes it — `INDEX` and `OFFSET` return a
 * reference too, but one they chose, and `INDEX(A:A,MATCH(…))` is a single cell.
 */
const PASSES_THROUGH = new Set(["IF", "CHOOSE"]);

/** `A1#`, or the `ANCHORARRAY(A1)` call a file stores it as: the anchor's spill range. */
function isSpillReference(node: AstNode): boolean {
  return (
    (node.type === NodeType.UnaryOp && node.op === SPILL_OPERATOR) ||
    (node.type === NodeType.FunctionCall &&
      node.invocations === undefined &&
      node.args.length === 1 &&
      bareFunctionName(node.name) === "ANCHORARRAY")
  );
}

function isRange(node: AstNode): boolean {
  if (isSpillReference(node)) {
    return true;
  }
  switch (node.type) {
    case NodeType.RangeRef:
    case NodeType.ColRangeRef:
    case NodeType.RowRangeRef:
      return true;
    case NodeType.CellRef:
      return node.endSheet !== undefined;
    case NodeType.BinaryOp:
      return node.op === ":";
    case NodeType.StructuredRef:
      return !isSingleCellStructuredRef(node);
    default:
      return false;
  }
}

function isSingleCellStructuredRef(node: StructuredRefNode): boolean {
  if (node.columns.length > 1) {
    return false;
  }
  if (node.tableName === "" || node.specials.some(s => s.toUpperCase() === "#THIS ROW")) {
    return true;
  }
  return (
    node.specials.length === 1 &&
    (node.specials[0]!.toUpperCase() === "#HEADERS" ||
      node.specials[0]!.toUpperCase() === "#TOTALS")
  );
}
