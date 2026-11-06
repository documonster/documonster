/**
 * The operand class Excel gives each token of a formula — reference, value or array — from the classes its functions
 * declare for their parameters (`@formula/syntax/function-table`).
 *
 * Apart from the codec in `ptg.ts` because the class is not only an encoding matter. A range token in the value class
 * of a cell formula is one Excel *implicitly intersects* with the formula's row or column when the formula is stored
 * as an ordinary one, and that decides what such a formula computes — see `legacyIntersections` and its use in
 * `utils/array-evaluation.ts`. One assignment serves both, so what is written and what is computed cannot disagree.
 */

import { NodeType, type AstNode, type FunctionCallNode } from "@formula/syntax/ast";
import {
  FUNCTION_ID_BY_NAME,
  functionClasses,
  futureFunctionClasses,
  paramClassAt,
  type FunctionClasses,
  type ParamClass
} from "@formula/syntax/function-table";
import { SPILL_OPERATOR } from "@formula/syntax/token-types";

/**
 * The operand class of a class-tagged token, added to the classless base values of the token ids in `ptg.ts`.
 *
 * `PtgRef` is `0x24` in the reference class, `0x44` in the value class and `0x64` in the array class; Excel's own
 * `=SUM(A1:A3)` is `25 …` and its `=A1+B1` is `44 … 44 … 03`.
 */
export const REFERENCE_CLASS = 0x20;
export const VALUE_CLASS = 0x40;
export const ARRAY_CLASS = 0x60;

/**
 * What a formula is for, which decides the class its root is evaluated in and what an unconverted value token
 * becomes.
 *
 * MS-XLSB states the classes per token and not the rule that assigns them; this follows the rule Excel's own
 * output shows, as LibreOffice's export implements it (`XclExpFmlaCompImpl::RecalcTokenClasses`): a cell or shared
 * formula is evaluated for one value; an array formula, a conditional format and a data-validation constraint are
 * evaluated over arrays; a defined name, a data-validation *list* and a sparkline range are references. Verified
 * against Excel: a data-validation list `$A$1:$A$10` is `25` (reference) where the same range as a `whole`
 * constraint operand is `44`.
 */
export type FormulaRole = "cell" | "array" | "conditional" | "validation" | "list" | "name";

/** `SUM`'s `iftab`: a one-argument call is written as `PtgAttrSum`. */
const SUM_ID = 4;

/** The class decided for each classed token: the node's own, and the `PtgMemFunc` wrapping a reference expression. */
export interface TokenClasses {
  readonly own: Map<AstNode, number>;
  readonly mem: Map<AstNode, number>;
}

/**
 * A parameter's conversion rule, plus `operator`: "whatever the parent's is", which is what an operator's operands
 * get (LibreOffice's `RPO`).
 */
type Conversion = ParamClass["conversion"] | "operator";
type ClassConversion = "original" | "value" | "array";

interface OperandRule {
  readonly conversion: Conversion;
  readonly valueType: boolean;
}

/**
 * Decide every classed token's class, from the root down.
 *
 * This is the algorithm Excel's own output follows, as LibreOffice's export implements it
 * (`XclExpFmlaCompImpl::RecalcTokenClass`): each token starts in its natural class — a reference in the reference
 * class, a function in its return class, an array constant in the array class — and is converted according to what
 * its consumer asked for and what *its* consumer asked for in turn. The cases that decide most formulas:
 *
 * - an operator's operands are wanted as values: `=A1+B1` is `44 … 44 … 03`, and `=ABS(A1-B1)` the same inside;
 * - `SUM`, `COUNT`, `MAX` and their kind take references: `=SUM(A1:A3)` is `25`;
 * - a function taking arrays converts everything below it: `=SUMPRODUCT(A1:A3*B1:B3)` is `65 … 65 … 05`;
 * - a reference-returning function stays a reference where one is wanted: `=SUM(OFFSET(A1,1,1,2,2))` is `22 05 4e`,
 *   and the same `OFFSET` as the whole formula is `42 05 4e`;
 * - outside a cell formula, a value that is not converted becomes an array: `AND($C5>0,C$5<10)` as a conditional
 *   format is `6c … 6c …` where the same formula in a cell is `44`.
 *
 * Each of these is asserted against Excel's bytes in `__tests__/operand-classes.node.test.ts`.
 */
export function assignClasses(
  root: AstNode,
  role: FormulaRole,
  bare: ReadonlySet<AstNode> = NONE
): TokenClasses {
  const classes: TokenClasses = { own: new Map(), mem: new Map() };
  const referenceRoot = role === "name" || role === "list";
  const cellType = role === "cell";
  const visit = (
    node: AstNode,
    rule: OperandRule,
    previous: Conversion,
    previousClass: ClassConversion,
    wasReference: boolean,
    memWrapped: boolean,
    parent: AstNode | undefined
  ): void => {
    if (bare.has(node) && node.type === NodeType.FunctionCall) {
      // An `_xlfn.SINGLE` call written as its operand alone takes the call's place — see `bareIntersections`.
      const operand = node.args[0]!;
      visit(
        operand,
        rule,
        previous,
        previousClass,
        wasReference,
        needsMemToken(operand, parent),
        parent
      );
      return;
    }
    if (memWrapped) {
      // The `PtgMemFunc` around a reference expression is itself a classed reference token, and the expression is
      // its operand, wanted as a reference.
      const memClass = convert(
        REFERENCE_CLASS,
        rule,
        previous,
        previousClass,
        wasReference,
        cellType
      );
      classes.mem.set(node, memClass.tokenClass);
      visit(
        node,
        { conversion: "value", valueType: false },
        memClass.conversion,
        memClass.classConversion,
        memClass.tokenClass === REFERENCE_CLASS,
        false,
        parent
      );
      return;
    }
    const natural = naturalClass(node);
    const decided = convert(natural, rule, previous, previousClass, wasReference, cellType);
    if (natural !== undefined) {
      classes.own.set(node, decided.tokenClass);
    }
    for (const operand of operandsOf(node)) {
      visit(
        operand.node,
        operand.rule,
        decided.conversion,
        decided.classConversion,
        natural !== undefined && decided.tokenClass === REFERENCE_CLASS,
        needsMemToken(operand.node, node),
        node
      );
    }
  };
  visit(
    root,
    referenceRoot
      ? { conversion: "array", valueType: false }
      : { conversion: "value", valueType: true },
    referenceRoot ? "array" : "value",
    referenceRoot ? "array" : "value",
    referenceRoot,
    needsMemToken(root, undefined),
    undefined
  );
  return classes;
}

const NONE: ReadonlySet<AstNode> = new Set();

/**
 * The `_xlfn.SINGLE` calls — the implicit-intersection operator `@` — of a formula that Excel writes as the operand
 * alone, in the value class: those of a cell formula whose operand, written bare in the call's place, is intersected
 * there and shown by Excel with an `@` anyway — a range, a name or a reference-returning call in the value class
 * (`@A1:A3+1`, `@Rng`, `@INDEX(A1:A3,0)`), or a call returning an array (`@SEQUENCE(3)`). Every other `@` is written as
 * the call, through the `_xlfn.SINGLE` stub: one whose operand would not be intersected (`SUM(@A1:A3)`,
 * `INDEX(@A1:A3,1)`), a single cell (`@A1`), and every `@` outside a cell formula, where a value-class operand is not
 * intersected but evaluated over its range (an array formula, a dynamic array, a conditional format, a validation, a
 * defined name). That is what Excel 16.114 wrote in all of them (`single-tokens.node.test.ts`).
 *
 * Decided from the root down: an `@`'s choice changes the classes below it, never those above.
 */
export function bareIntersections(root: AstNode, role: FormulaRole): ReadonlySet<AstNode> {
  const bare = new Set<AstNode>();
  if (role !== "cell") {
    return bare;
  }
  const visit = (node: AstNode): void => {
    if (node.type === NodeType.FunctionCall && isImplicitIntersection(node)) {
      bare.add(node);
      if (!isShownIntersected(node.args[0]!, assignClasses(root, role, bare))) {
        bare.delete(node);
      }
    }
    for (const operand of operandsOf(node)) {
      visit(operand.node);
    }
  };
  visit(root);
  return bare;
}

/** Whether Excel intersects `node` in its place and displays it with an `@` — see {@link bareIntersections}. */
function isShownIntersected(node: AstNode, classes: TokenClasses): boolean {
  if (classes.mem.get(node) === VALUE_CLASS) {
    return true;
  }
  if (classes.own.get(node) !== VALUE_CLASS) {
    return false;
  }
  const natural = naturalClass(node);
  return (
    (natural === REFERENCE_CLASS && !isSingleCell(node)) ||
    (natural === ARRAY_CLASS && node.type === NodeType.FunctionCall)
  );
}

/**
 * One token's conversion. `natural` is `undefined` for a classless token — an operator or a literal — which still
 * passes a conversion down to its operands.
 */
function convert(
  natural: number | undefined,
  rule: OperandRule,
  previous: Conversion,
  previousClass: ClassConversion,
  wasReference: boolean,
  cellType: boolean
): { tokenClass: number; conversion: Conversion; classConversion: ClassConversion } {
  let tokenClass = natural ?? 0;
  // A reference in a parameter that wants a value behaves like a value.
  if (rule.valueType && tokenClass === REFERENCE_CLASS) {
    tokenClass = VALUE_CLASS;
  }
  const conversion = rule.conversion === "operator" ? previous : rule.conversion;
  let classConversion: ClassConversion;
  switch (conversion) {
    case "original":
      classConversion = "original";
      break;
    case "value":
      classConversion = "value";
      break;
    case "array":
      classConversion = "array";
      break;
    case "repeat":
      // Repeat the parent's conversion — except below a reference that a parameter took as a reference, where it
      // falls back to value: a reference argument's own arguments are not evaluated over arrays.
      classConversion =
        previous === "repeat"
          ? previousClass
          : previous === "repeatReference"
            ? wasReference
              ? previousClass
              : "original"
            : wasReference
              ? "value"
              : previousClass;
      break;
    default:
      // `repeatReference`: a reference keeps the parent's conversion, and so does anything below an array one.
      classConversion =
        tokenClass === REFERENCE_CLASS || previousClass === "array" ? previousClass : "original";
      break;
  }
  if (natural !== undefined) {
    if (classConversion === "original") {
      // Only a cell formula leaves a value token alone; everywhere else it is evaluated as an array.
      if (!cellType && tokenClass === VALUE_CLASS) {
        tokenClass = ARRAY_CLASS;
      }
    } else if (classConversion === "value" && tokenClass === ARRAY_CLASS) {
      tokenClass = VALUE_CLASS;
    } else if (classConversion === "array" && tokenClass === VALUE_CLASS) {
      tokenClass = ARRAY_CLASS;
    }
  }
  return { tokenClass, conversion, classConversion };
}

/**
 * The spill-range operator `A1#` is the call `_xlfn.ANCHORARRAY(A1)` — the form a file stores and the encoder writes —
 * so its operand and result take that call's classes: a reference parameter, a reference result.
 */
const SPILL_CLASSES = futureFunctionClasses("_xlfn.ANCHORARRAY");

function isSpillOperation(node: AstNode): boolean {
  return node.type === NodeType.UnaryOp && node.op === SPILL_OPERATOR;
}

/** A classed token's class before conversion, or `undefined` for a token with no class. */
function naturalClass(node: AstNode): number | undefined {
  if (isSpillOperation(node)) {
    return SPILL_CLASSES.returns === "R" ? REFERENCE_CLASS : VALUE_CLASS;
  }
  switch (node.type) {
    case NodeType.CellRef:
    case NodeType.RangeRef:
    case NodeType.ColRangeRef:
    case NodeType.RowRangeRef:
    case NodeType.Name:
    case NodeType.StructuredRef:
      return REFERENCE_CLASS;
    case NodeType.Error:
      // `#REF!` is written as `PtgRefErr`, which is a reference; any other error is a classless `PtgErr`.
      return node.value === "#REF!" ? REFERENCE_CLASS : undefined;
    case NodeType.Array:
      return ARRAY_CLASS;
    case NodeType.FunctionCall: {
      if (!isInvocation(node) && isAttrSum(node)) {
        return undefined;
      }
      const classes = classesOfCall(node);
      return classes.returns === "R"
        ? REFERENCE_CLASS
        : classes.returns === "A"
          ? ARRAY_CLASS
          : VALUE_CLASS;
    }
    default:
      return undefined;
  }
}

/** Whether `node` is a reference operator — `:`, intersection, or a union. */
export function isReferenceOperator(node: AstNode): boolean {
  return (
    node.type === NodeType.UnionRef ||
    (node.type === NodeType.BinaryOp && (node.op === ":" || node.op === " "))
  );
}

/**
 * Whether `node` starts a reference expression that a `PtgMemFunc` must enclose.
 *
 * Excel encloses every expression built from reference operators in a mem token — `PtgMemArea` with the cached
 * result when it can compute one, `PtgMemFunc` when it cannot (`A1:INDEX(A1:C3,2,2)` is `49 1f 00 …`). This writer
 * always uses `PtgMemFunc`, the form that asks for evaluation rather than asserting a cached answer; Excel accepts
 * it for any reference expression and recomputes on load. Nested reference operators share their outermost one's
 * token, as Excel's do.
 */
export function needsMemToken(node: AstNode, parent: AstNode | undefined): boolean {
  return isReferenceOperator(node) && (parent === undefined || !isReferenceOperator(parent));
}

const OPERATOR_VALUE: OperandRule = { conversion: "operator", valueType: true };
const OPERATOR_REFERENCE: OperandRule = { conversion: "operator", valueType: false };

export function operandsOf(node: AstNode): { node: AstNode; rule: OperandRule }[] {
  switch (node.type) {
    case NodeType.BinaryOp: {
      const rule = node.op === ":" || node.op === " " ? OPERATOR_REFERENCE : OPERATOR_VALUE;
      return [
        { node: node.left, rule },
        { node: node.right, rule }
      ];
    }
    case NodeType.UnionRef:
      return node.areas.map(area => ({ node: area, rule: OPERATOR_REFERENCE }));
    case NodeType.UnaryOp:
      if (node.op === SPILL_OPERATOR) {
        const param = paramClassAt(SPILL_CLASSES, 0);
        return [
          { node: node.operand, rule: { conversion: param.conversion, valueType: param.valueType } }
        ];
      }
      // `@` is implicit intersection, which is what a value-class reference *is* in a cell formula. The encoder never
      // sees one — it writes `@x` as the `_xlfn.SINGLE` call (see `bareIntersections`) — but `implicitIntersections`
      // and `nonReferenceArgument` read the tree as parsed.
      return [
        {
          node: node.operand,
          rule: node.op === "@" ? { conversion: "value", valueType: true } : OPERATOR_VALUE
        }
      ];
    case NodeType.Percent:
      return [{ node: node.operand, rule: OPERATOR_VALUE }];
    case NodeType.FunctionCall: {
      if (isInvocation(node)) {
        return invocationOperands(node).map((operand, index) => {
          const param = paramClassAt(INVOCATION_CLASSES, index);
          return {
            node: operand,
            rule: { conversion: param.conversion, valueType: param.valueType }
          };
        });
      }
      const classes = classesOfCall(node);
      return node.args.map((arg, index) => {
        const param = paramClassAt(classes, index);
        return { node: arg, rule: { conversion: param.conversion, valueType: param.valueType } };
      });
    }
    default:
      return [];
  }
}

export function classesOfCall(node: FunctionCallNode): FunctionClasses {
  if (isInvocation(node)) {
    return INVOCATION_CLASSES;
  }
  const id = builtInId(node.name);
  return id === undefined ? futureFunctionClasses(node.name) : functionClasses(id);
}

/**
 * An argument list applied to the result of a call — `LAMBDA(x,x*2)(5)`, `f(1)(2)`.
 *
 * Excel's own XLSB of `_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(5)` is
 * `23 <LAMBDA> 23 <x> 43 <x> 1e 02 00 05 42 03 ff 00 1e 05 00 42 02 ff 00`: the `LAMBDA` call as usual, then the
 * applied arguments, then `PtgFuncVar` with `iftab` `0x00FF` — the same token that calls a function through a
 * `PtgName`, whose first operand is the callee. Here the callee is the expression before the arguments rather than a
 * name, and `cparams` counts it. Each further argument list applies another `PtgFuncVar` to the result.
 *
 * A call through a `LET`/`LAMBDA` parameter (`_xlpm.f(1,2)`) is the same kind of call and has the same classes —
 * see `futureFunctionClasses`. The callee keeps its own class (`LAMBDA` is `42` above), and the result is a
 * reference where one is wanted, `22`, like any call whose result is not known in advance.
 */
export const INVOCATION_CLASSES: FunctionClasses = {
  returns: "R",
  params: [{ valueType: false, conversion: "original" }],
  volatile: false,
  pairs: false
};

/**
 * The marker name of an inner level of a chain of applied argument lists, `f(1)(2)(3)`: the AST keeps the whole chain
 * on one node, and the token stream has a `PtgFuncVar` per list. Not a name any formula can spell.
 */
export const INVOKE = "\u0000INVOKE";

/** The callee node of each invoked call, and the levels of each chain — one object per node, so classes stick. */
const INVOCATION_LEVELS = new WeakMap<FunctionCallNode, FunctionCallNode>();

export function isInvocation(node: FunctionCallNode): boolean {
  return node.invocations !== undefined || node.name === INVOKE;
}

/**
 * The operands of the outermost `PtgFuncVar` of an invoked call: its callee, then the last applied argument list.
 * The callee is the same call with one list fewer — the bare call when none is left.
 */
export function invocationOperands(node: FunctionCallNode): AstNode[] {
  if (node.name === INVOKE) {
    return node.args;
  }
  let callee = INVOCATION_LEVELS.get(node);
  if (callee === undefined) {
    const lists = node.invocations ?? [];
    callee = { type: NodeType.FunctionCall, name: node.name, args: node.args };
    for (const list of lists.slice(0, -1)) {
      callee = { type: NodeType.FunctionCall, name: INVOKE, args: [callee, ...list] };
    }
    INVOCATION_LEVELS.set(node, callee);
  }
  return [callee, ...(node.invocations?.[node.invocations.length - 1] ?? [])];
}

/**
 * `_xlfn.SINGLE(x)` — how a file spells Excel 365's `@x`, implicit intersection. XLSB calls it through the
 * `_xlfn.SINGLE` stub, or writes `x` alone where that means the same — see {@link bareIntersections}.
 */
export function isImplicitIntersection(node: FunctionCallNode): boolean {
  return node.args.length === 1 && node.name.toUpperCase() === "_XLFN.SINGLE";
}

/** A one-argument `SUM`, which Excel writes as `PtgAttrSum` rather than as a call. */
export function isAttrSum(node: FunctionCallNode): boolean {
  return builtInId(node.name) === SUM_ID && node.args.length === 1;
}

/** The `iftab` of a built-in function, or `undefined` for one called through a name. */
export function builtInId(name: string): number | undefined {
  return FUNCTION_ID_BY_NAME.get(name.toUpperCase());
}

/**
 * The operands of a cell formula that Excel *implicitly intersects* with the formula's row or column when the formula
 * is stored as an ordinary one: every reference — a range, a name, a structured reference, a call returning one, a
 * reference expression — that the class assignment puts in the value class. That is the meaning of the value class
 * (Excel's own XLSB of `_xlfn.SINGLE(A1:A3)`, the `@` of Excel 365, is the bare `45` area), so the set is what Excel
 * would mark with `@` when it displays the formula: `SUM(@A1:A5*2)`, `ABS(@A1:A5)`, but `SUMPRODUCT(A1:A5*2)`,
 * `ROWS(A1:A5+0)` and `INDEX(A1:A5*3,4)` untouched.
 *
 * A single cell is left out: intersecting it changes nothing. So is anything already below an explicit `@`.
 */
export function implicitIntersections(root: AstNode): ReadonlySet<AstNode> {
  const classes = assignClasses(root, "cell");
  const found = new Set<AstNode>();
  const visit = (node: AstNode): void => {
    if (
      (node.type === NodeType.UnaryOp && node.op === "@") ||
      (node.type === NodeType.FunctionCall && isImplicitIntersection(node))
    ) {
      return;
    }
    if (
      classes.mem.get(node) === VALUE_CLASS ||
      (classes.own.get(node) === VALUE_CLASS &&
        naturalClass(node) === REFERENCE_CLASS &&
        !isSingleCell(node))
    ) {
      found.add(node);
    }
    for (const operand of operandsOf(node)) {
      visit(operand.node);
    }
  };
  visit(root);
  return found;
}

function isSingleCell(node: AstNode): boolean {
  switch (node.type) {
    case NodeType.CellRef:
      return node.endSheet === undefined;
    case NodeType.StructuredRef: {
      if (node.columns.length > 1) {
        return false;
      }
      const specials = node.specials.map(s => s.toUpperCase());
      return (
        node.tableName === "" ||
        specials.includes("#THIS ROW") ||
        (node.columns.length === 1 &&
          specials.length === 1 &&
          (specials[0] === "#HEADERS" || specials[0] === "#TOTALS"))
      );
    }
    case NodeType.Error:
      return true;
    default:
      return false;
  }
}

/**
 * Why a formula cannot be written: an argument that is plainly not a reference — a constant, an array constant, or the
 * result of an operator — in a parameter only a reference can fill (`referenceOnly` in `@formula/syntax/function-table`).
 * Excel refuses such a formula when it is typed, and refuses to open a file that holds one: `RANK(2,{2,2,1},0)`.
 *
 * Only what is certainly not a reference is refused. A call, a name or an `@` may well yield one, and is left to Excel.
 *
 * @returns `undefined` when every reference-only parameter is given something that can be a reference.
 */
export function nonReferenceArgument(root: AstNode): string | undefined {
  let found: string | undefined;
  const visit = (node: AstNode): void => {
    if (found !== undefined) {
      return;
    }
    if (node.type === NodeType.FunctionCall && !isInvocation(node)) {
      const classes = classesOfCall(node);
      node.args.forEach((arg, index) => {
        if (
          found === undefined &&
          paramClassAt(classes, index).referenceOnly &&
          isNotReference(arg)
        ) {
          found = `${node.name.replace(/^(?:_xlfn\.|_xlws\.)+/i, "").toUpperCase()} argument ${index + 1} must be a reference`;
        }
      });
    }
    for (const operand of operandsOf(node)) {
      visit(operand.node);
    }
  };
  visit(root);
  return found;
}

function isNotReference(node: AstNode): boolean {
  switch (node.type) {
    case NodeType.Number:
    case NodeType.String:
    case NodeType.Boolean:
    case NodeType.Array:
    case NodeType.Percent:
      return true;
    case NodeType.Error:
      return node.value !== "#REF!";
    case NodeType.BinaryOp:
      return !isReferenceOperator(node);
    case NodeType.UnaryOp:
      return node.op !== "@" && node.op !== SPILL_OPERATOR;
    default:
      return false;
  }
}
