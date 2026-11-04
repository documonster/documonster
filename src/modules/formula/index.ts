/**
 * Public entry for the documonster formula engine.
 *
 * This entry exposes syntax inspection only. Workbook recalculation lives at
 * `documonster/excel/formula`, where the Excel host adapter and the calculation
 * engine are linked without adding either to `documonster/excel` itself.
 *
 * To recalculate a workbook from `documonster/excel`, use the adapter exported
 * by `documonster/excel/formula`:
 * ```ts
 * import { ExcelFormula } from "documonster/excel/formula";
 * ExcelFormula.calculate(workbook);
 * ```
 */

// Public value API — the `Formula` domain namespace. Tree-shaken per-member
// on rolldown / rspack; a consumer that references only `Formula.tokenize`
// never pulls in the evaluator.
export * as Formula from "@formula/surface/formula";

// Errors — extend BaseError, consistent with every other module's errors.ts.
export {
  FormulaError,
  FormulaArgumentError,
  FormulaParseError,
  isFormulaError
} from "@formula/errors";

// Syntax types, so callers can name what `tokenize` / `parse` return and switch
// exhaustively over token / node kinds. Each member of the `Token` and `AstNode` unions is
// published too, so a visitor can name the node it handles. Members carry the module prefix
// (`FormulaCellRefNode`), as every flat type of an entry with generic vocabulary does;
// `Token`, `AstNode`, `TokenType` and `NodeType` were published before that rule.
export { TokenType } from "@formula/syntax/token-types";
export type {
  Token,
  NumberToken as FormulaNumberToken,
  StringToken as FormulaStringToken,
  BooleanToken as FormulaBooleanToken,
  ErrorToken as FormulaErrorToken,
  CellRefToken as FormulaCellRefToken,
  RangeToken as FormulaRangeToken,
  SheetRefToken as FormulaSheetRefToken,
  FunctionToken as FormulaFunctionToken,
  OperatorToken as FormulaOperatorToken,
  OpenParenToken as FormulaOpenParenToken,
  CloseParenToken as FormulaCloseParenToken,
  CommaToken as FormulaCommaToken,
  ColonToken as FormulaColonToken,
  PercentToken as FormulaPercentToken,
  OpenBraceToken as FormulaOpenBraceToken,
  CloseBraceToken as FormulaCloseBraceToken,
  SemicolonToken as FormulaSemicolonToken,
  UnaryPrefixToken as FormulaUnaryPrefixToken,
  NameToken as FormulaNameToken,
  ColRangeToken as FormulaColRangeToken,
  RowRangeToken as FormulaRowRangeToken,
  StructuredRefToken as FormulaStructuredRefToken,
  AtSignToken as FormulaAtSignToken,
  IntersectToken as FormulaIntersectToken
} from "@formula/syntax/token-types";
export { NodeType } from "@formula/syntax/ast";
export type {
  AstNode,
  NumberNode as FormulaNumberNode,
  StringNode as FormulaStringNode,
  BooleanNode as FormulaBooleanNode,
  ErrorNode as FormulaErrorNode,
  CellRefNode as FormulaCellRefNode,
  RangeRefNode as FormulaRangeRefNode,
  BinaryOpNode as FormulaBinaryOpNode,
  UnaryOpNode as FormulaUnaryOpNode,
  FunctionCallNode as FormulaFunctionCallNode,
  ArrayNode as FormulaArrayNode,
  PercentNode as FormulaPercentNode,
  NameNode as FormulaNameNode,
  ColRangeRefNode as FormulaColRangeRefNode,
  RowRangeRefNode as FormulaRowRangeRefNode,
  StructuredRefNode as FormulaStructuredRefNode,
  MissingNode as FormulaMissingNode,
  UnionRefNode as FormulaUnionRefNode
} from "@formula/syntax/ast";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
