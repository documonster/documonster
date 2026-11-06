/**
 * Whether a formula produces an array that dynamic-array Excel spills — decided from the formula's text alone.
 *
 * **Why statically.** A host has to store a formula either as a dynamic array or as an ordinary one before anything
 * is evaluated, and the engine has to evaluate it the way it is stored: Excel computes an ordinary formula with
 * implicit intersection (the top-left value of an array), and spills a dynamic one. Deciding from the evaluated result
 * would make a file depend on whether it was calculated before it was written. So the engine and the Excel writers
 * ask this one question, and a formula the engine spills is exactly one a file stores as a dynamic array.
 *
 * **What produces an array**, on the path from the root — through operators and the value parameters that lift a
 * function over an array (`ABS(SEQUENCE(3))`), not through a parameter that takes its argument whole
 * (`SUM(SEQUENCE(3))`):
 *
 * - a call to a function whose result is an array — the return classes of `function-table.ts`: `TRANSPOSE`, `MMULT`,
 *   `FREQUENCY`, `LINEST`, … and every future function that returns an array or a reference (`SEQUENCE`, `FILTER`,
 *   `TAKE`, `XLOOKUP`, `MAP`, …). Those return classes are LibreOffice's; three disagree with Microsoft's
 *   documentation, which describes `RTD`, `TEXTAFTER` and `TEXTBEFORE` as returning one value (they lift over an array
 *   argument like any value function), and `STOCKHISTORY` is documented as returning an array but has no entry;
 * - an array constant of more than one element, `{1,2,3}`;
 * - the spill-range operator `A1#` (`ANCHORARRAY(A1)`);
 * - what a `LET` or an invoked `LAMBDA` returns, with the names it binds standing for their values.
 *
 * **Where it is only known at evaluation time** — `XLOOKUP` returning one column or two, `REGEXEXTRACT` with a
 * computed return mode, `INDEX` into an array — the answer is "an array". A dynamic array whose result is one value
 * is that value in Excel, so the over-approximation costs nothing, where the under-approximation would store a
 * formula Excel intersects while the engine spills it.
 *
 * **A reference is not an array here**, whatever it covers: a range (`A1:A3`), a name, or a legacy function returning
 * one (`INDEX(A1:C3,0,1)`, `OFFSET`, `INDIRECT`). Its extent is known only when it is evaluated, an ordinary formula
 * intersects it, and a host that stores it as an ordinary formula has the engine intersect it too — the rule the Excel
 * writers have always applied to a range at the root.
 */
import { NodeType, type AstNode, type FunctionCallNode } from "@formula/syntax/ast";
import {
  FUNCTION_ID_BY_NAME,
  functionClasses,
  futureFunctionClasses,
  paramClassAt
} from "@formula/syntax/function-table";
import { SPILL_OPERATOR, stripFunctionPrefix } from "@formula/syntax/token-types";
import { sheetNameKey } from "@utils/name-key";

/** Functions documented as returning one value, which the return classes list as returning an array. */
const RETURNS_VALUE = new Set(["RTD", "TEXTAFTER", "TEXTBEFORE"]);

/** Functions documented as returning an array that have no return class of their own. */
const RETURNS_ARRAY = new Set(["STOCKHISTORY"]);

/** Functions whose result is one of their arguments: an array argument is what they return. */
const PASSES_THROUGH = new Set(["IF", "CHOOSE", "IFS", "SWITCH", "IFERROR", "IFNA"]);

const XLPM = "_XLPM.";

/** What a name bound by `LET` or a `LAMBDA` parameter stands for: an array or not, or a lambda to be invoked. */
type Binding = boolean | FunctionCallNode;

type Scope = ReadonlyMap<string, Binding>;

/**
 * The functions a host registered, by upper-case name, each with whether it returns an array. A registered function's
 * own declaration decides — over a built-in of the same name, whose behaviour it replaces — and one that says nothing
 * returns one value. `Workbook.registerFunction`'s registry and `calculateFormulas`' `userFunctions` are both one.
 */
export type RegisteredFunctions = ReadonlyMap<string, { readonly returnsArray?: boolean }>;

/**
 * Whether the function called `name` (any `_xlfn.` prefix, any case) returns an array whatever its arguments — as
 * `functions` declares it, when it registers the name.
 */
export function returnsArray(name: string, functions?: RegisteredFunctions): boolean {
  const bare = stripFunctionPrefix(name.toUpperCase());
  const registered = functions?.get(bare);
  if (registered !== undefined) {
    return registered.returnsArray === true;
  }
  if (RETURNS_VALUE.has(bare)) {
    return false;
  }
  if (RETURNS_ARRAY.has(bare)) {
    return true;
  }
  if (bare.startsWith(XLPM)) {
    return false;
  }
  const id = FUNCTION_ID_BY_NAME.get(bare);
  if (id !== undefined) {
    return functionClasses(id).returns === "A";
  }
  // A future function returning a reference is a dynamic-array function too — `TAKE`, `DROP`, `XLOOKUP`,
  // `TRIMRANGE`, `ANCHORARRAY` — and what it covers is decided by its arguments.
  return futureFunctionClasses(bare).returns !== "V";
}

/**
 * Whether `root`, as a cell's whole formula, produces an array — see the module documentation. `functions` are the
 * host's registered functions, whose declarations decide for the names they register.
 */
export function producesArray(root: AstNode, functions?: RegisteredFunctions): boolean {
  return produces(root, new Map(), functions);
}

function bindingKey(name: string): string {
  return sheetNameKey(name.toUpperCase().startsWith(XLPM) ? name.slice(XLPM.length) : name);
}

function produces(
  node: AstNode,
  scope: Scope,
  functions: RegisteredFunctions | undefined
): boolean {
  switch (node.type) {
    case NodeType.Array:
      return node.rows.length * (node.rows[0]?.length ?? 0) > 1
        ? true
        : node.rows.some(row => row.some(item => produces(item, scope, functions)));
    case NodeType.UnaryOp:
      if (node.op === SPILL_OPERATOR) {
        return true;
      }
      return node.op !== "@" && produces(node.operand, scope, functions);
    case NodeType.Percent:
      return produces(node.operand, scope, functions);
    case NodeType.BinaryOp:
      // `:` and intersection build references.
      return (
        node.op !== ":" &&
        node.op !== " " &&
        (produces(node.left, scope, functions) || produces(node.right, scope, functions))
      );
    case NodeType.Name: {
      const bound = scope.get(bindingKey(node.name));
      return bound === true;
    }
    case NodeType.FunctionCall:
      return callProduces(node, scope, functions);
    default:
      return false;
  }
}

function callProduces(
  node: FunctionCallNode,
  scope: Scope,
  functions: RegisteredFunctions | undefined
): boolean {
  const bare = stripFunctionPrefix(node.name.toUpperCase());
  if (node.invocations !== undefined) {
    return invocationProduces(node, scope, functions);
  }
  if (bare === "SINGLE") {
    return false;
  }
  if (bare === "LET") {
    return letProduces(node, scope, functions);
  }
  if (bare === "LAMBDA") {
    // An uninvoked lambda is not a value a cell can hold.
    return false;
  }
  const bound = scope.get(bindingKey(node.name));
  if (bound !== undefined && FUNCTION_ID_BY_NAME.get(bare) === undefined) {
    // A call through a `LET`-bound name: the lambda bound to it, when that is known; anything else could return
    // anything, and is taken to return an array.
    return typeof bound === "object" ? lambdaProduces(bound, node.args, scope, functions) : true;
  }
  if (bare.startsWith(XLPM)) {
    return true;
  }
  // A registered function returns what it declares, whatever a built-in of its name returns or how it lifts.
  const registered = functions?.get(bare);
  if (registered !== undefined) {
    return registered.returnsArray === true;
  }
  if (returnsArray(bare)) {
    return true;
  }
  if (bare === "REGEXEXTRACT") {
    // `return_mode` 1 and 2 return arrays (all matches, the capture groups); 0, the default, one string.
    const mode = node.args[2];
    if (mode !== undefined && mode.type !== NodeType.Missing) {
      if (mode.type !== NodeType.Number || mode.value !== 0) {
        return true;
      }
    }
  }
  if (bare === "INDEX" && node.args[0] !== undefined && produces(node.args[0], scope, functions)) {
    // A row or column of an array, or one element of it — which, depends on the indices and the array's shape.
    return true;
  }
  const id = FUNCTION_ID_BY_NAME.get(bare);
  const classes = id === undefined ? futureFunctionClasses(bare) : functionClasses(id);
  const passes = PASSES_THROUGH.has(bare);
  return node.args.some((arg, index) => {
    const param = paramClassAt(classes, index);
    return (
      (passes || (param.valueType && param.conversion !== "array")) &&
      produces(arg, scope, functions)
    );
  });
}

function letProduces(
  node: FunctionCallNode,
  outer: Scope,
  functions: RegisteredFunctions | undefined
): boolean {
  const scope = new Map(outer);
  const { args } = node;
  for (let i = 0; i + 1 < args.length - 1; i += 2) {
    const name = args[i]!;
    if (name.type !== NodeType.Name) {
      return true;
    }
    scope.set(bindingKey(name.name), bindingOf(args[i + 1]!, scope, functions));
  }
  const body = args[args.length - 1];
  return body !== undefined && produces(body, scope, functions);
}

function bindingOf(
  value: AstNode,
  scope: Scope,
  functions: RegisteredFunctions | undefined
): Binding {
  return isLambda(value) ? value : produces(value, scope, functions);
}

/** What `lambda` returns, invoked with `args`. */
function lambdaProduces(
  lambda: FunctionCallNode,
  args: readonly AstNode[],
  outer: Scope,
  functions: RegisteredFunctions | undefined
): boolean {
  return applyLambda(lambda, [args], outer, functions);
}

/**
 * What `lambda` returns, invoked with each argument list in turn: a curried `LAMBDA(x,LAMBDA(y,x+y))(1)(2)` applies
 * the second list to the lambda the first returns. A body that is not a lambda where another list follows could
 * return anything, and is taken to return an array.
 */
function applyLambda(
  lambda: FunctionCallNode,
  lists: readonly (readonly AstNode[])[],
  outer: Scope,
  functions: RegisteredFunctions | undefined
): boolean {
  let current = lambda;
  const scope = new Map(outer);
  for (let n = 0; n < lists.length; n++) {
    const args = lists[n]!;
    const params = current.args.slice(0, -1);
    const body = current.args[current.args.length - 1];
    if (body === undefined) {
      return false;
    }
    const bindings = new Map<string, Binding>();
    for (let i = 0; i < params.length; i++) {
      const param = params[i]!;
      if (param.type !== NodeType.Name) {
        return true;
      }
      const arg = args[i];
      bindings.set(
        bindingKey(param.name),
        arg === undefined ? false : bindingOf(arg, outer, functions)
      );
    }
    for (const [key, value] of bindings) {
      scope.set(key, value);
    }
    if (n === lists.length - 1) {
      return produces(body, scope, functions);
    }
    if (!isLambda(body)) {
      return true;
    }
    current = body;
  }
  return false;
}

function isLambda(node: AstNode): node is FunctionCallNode {
  return (
    node.type === NodeType.FunctionCall &&
    node.invocations === undefined &&
    stripFunctionPrefix(node.name.toUpperCase()) === "LAMBDA"
  );
}

function invocationProduces(
  node: FunctionCallNode,
  scope: Scope,
  functions: RegisteredFunctions | undefined
): boolean {
  // `LAMBDA(x,…)(args)`: what its body returns. A computed callee could return anything.
  if (stripFunctionPrefix(node.name.toUpperCase()) === "LAMBDA") {
    return applyLambda(node, node.invocations!, scope, functions);
  }
  return true;
}
