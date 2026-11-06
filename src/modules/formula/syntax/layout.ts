/**
 * Formula layout: the whitespace and redundant parentheses a formula's text carries and its AST does not.
 *
 * The AST is a statement of *meaning*. `(A1+B1)*2` and `( A1 + B1 ) * 2` are the same tree, and every consumer
 * that evaluates, rewrites or analyses a formula wants exactly that. A consumer that has to reproduce the text an
 * author wrote does not — and the XLSB codec is such a consumer: BIFF12 stores a formula as tokens, and Excel
 * records the author's spacing (`PtgAttrSpace`) and parentheses (`PtgParen`) as tokens of their own, so a file
 * Excel wrote shows `( A1 + B1 ) * 2` in Excel and must read back that way.
 *
 * Kept *beside* the tree rather than inside it, keyed by node identity, so the public `AstNode` union is untouched
 * and nothing that already walks a tree sees a difference. A layout is only produced on request
 * (`tokenizeWithGaps` + `parseWithLayout`) and only consumed on request (`printAstWithLayout`).
 *
 * What a layout can say is exactly what a token stream can carry, which is also what Excel preserves when it
 * re-saves a formula: whitespace before an operand, an operator, a function's closing parenthesis, and around a
 * display parenthesis. Whitespace before a comma is not representable and Excel drops it too.
 */

import type { AstNode } from "@formula/syntax/ast";

/** One display parenthesis around a node: the whitespace before its `(` and before its `)`. */
export interface ParenLayout {
  readonly open?: string;
  readonly close?: string;
}

/** What one node's text carries beyond its meaning. Every field is whitespace except `parens` and `bare`. */
export interface NodeLayout {
  /** Before the node's first character — a literal, a reference, a function's name, a unary operator. */
  lead?: string;
  /**
   * Before a binary operator, or before `%`.
   *
   * For the intersection operator, which *is* whitespace, this is the whole run: `A1:A3  B2:C2` has `"  "`.
   */
  op?: string;
  /** Before the `)` of a function call or of a reference union. */
  close?: string;
  /** Before the `)` of each argument list applied to a call's result, by position — see `FunctionCallNode.invocations`. */
  invocationClose?: string[];
  /** Parentheses the author wrote that precedence does not need, innermost first. */
  parens?: ParenLayout[];
  /**
   * A reference union written without its parentheses — legal only as a whole definition, which is how a
   * multi-area defined name is stored (`S!$A$1,S!$B$2`). Ignored anywhere else, where the parentheses are
   * what makes the commas a union rather than argument separators.
   */
  bare?: boolean;
}

/** Layout by node identity. */
export type FormulaLayout = Map<AstNode, NodeLayout>;

/** The layout entry for `node`, created on first use. */
export function layoutOf(layout: FormulaLayout, node: AstNode): NodeLayout {
  let entry = layout.get(node);
  if (entry === undefined) {
    entry = {};
    layout.set(node, entry);
  }
  return entry;
}
