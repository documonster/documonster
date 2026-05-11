/**
 * Word Document Field Calculation Engine
 *
 * Computes and updates field cachedValues based on layout (pagination) results.
 * Supports PAGE, NUMPAGES, SECTIONPAGES, SECTION, TOC, REF, PAGEREF, SEQ, IF, STYLEREF,
 * INDEX, = (formula), INCLUDETEXT.
 *
 * The engine performs a layout pass via layoutDocument, then traverses the document
 * body to update each field's cachedValue in an immutable-style manner (deep-cloning
 * modified portions of the tree).
 */

import { extractParagraphText, isRun } from "../core/text-utils";
import { walkBlocks } from "../core/walker";
import { layoutDocument, type LayoutOptions, type LayoutResult } from "../layout/layout";
import type {
  BodyContent,
  BookmarkStart,
  DocxDocument,
  FieldContent,
  Paragraph,
  ParagraphChild,
  Run,
  RunContent,
  Table,
  TableCell,
  TableOfContents
} from "../types";

// =============================================================================
// Public API Types
// =============================================================================

/** 字段计算选项 */
export interface FieldUpdateOptions {
  /** 布局选项（传递给 layoutDocument）。 */
  readonly layoutOptions?: LayoutOptions;
  /** 是否更新 TOC（默认 true）。 */
  readonly updateToc?: boolean;
  /** 是否更新页码字段（默认 true）。 */
  readonly updatePageFields?: boolean;
  /** 是否更新交叉引用（默认 true）。 */
  readonly updateReferences?: boolean;
  /** 是否更新序列号（默认 true）。 */
  readonly updateSequences?: boolean;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * 更新文档中所有字段的 cachedValue。
 *
 * 执行分页布局，然后遍历所有字段更新其缓存值。
 * 对于 TOC 字段，会生成完整的目录段落。
 *
 * @param doc - 文档模型（不会被修改，返回新对象）
 * @param options - 更新选项
 * @returns 更新后的文档模型
 */
export function updateFields(doc: DocxDocument, options?: FieldUpdateOptions): DocxDocument {
  const opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & {
    layoutOptions?: LayoutOptions;
  } = {
    layoutOptions: options?.layoutOptions,
    updateToc: options?.updateToc ?? true,
    updatePageFields: options?.updatePageFields ?? true,
    updateReferences: options?.updateReferences ?? true,
    updateSequences: options?.updateSequences ?? true
  };

  // Perform layout pass
  const layout = layoutDocument(doc, opts.layoutOptions);

  // Build bookmark info map
  const bookmarkInfo = opts.updateReferences ? collectBookmarkInfo(doc, layout) : new Map();

  // Build SEQ counters
  const seqValues = opts.updateSequences ? computeSeqValues(doc) : new Map();

  // Build style → paragraphs index for STYLEREF
  const styleIndex = buildStyleIndex(doc);

  // Collect INDEX entries (XE fields) from all body content
  const indexEntries = collectIndexEntries(doc);

  // Update body content
  const newBody = updateBody(doc, layout, bookmarkInfo, seqValues, styleIndex, indexEntries, opts);

  if (newBody === doc.body) {
    return doc;
  }

  return { ...doc, body: newBody };
}

/**
 * 仅更新 TOC（目录）字段。
 * 扫描文档中的标题（Heading1-9 样式或 outline level），
 * 生成目录缓存段落。
 */
export function updateTableOfContents(
  doc: DocxDocument,
  options?: FieldUpdateOptions
): DocxDocument {
  const layout = layoutDocument(doc, options?.layoutOptions);
  const headings = collectHeadings(doc, layout);

  const newBody = doc.body.map(item => {
    if (item.type === "tableOfContents") {
      return updateTocContent(item, headings);
    }
    return item;
  });

  const changed = newBody.some((item, i) => item !== doc.body[i]);
  if (!changed) {
    return doc;
  }

  return { ...doc, body: newBody };
}

// =============================================================================
// Internal Types
// =============================================================================

interface BookmarkData {
  readonly text: string;
  readonly page: number;
}

interface HeadingEntry {
  readonly text: string;
  readonly level: number;
  readonly page: number;
}

interface IndexEntry {
  readonly term: string;
  readonly page: number;
}

type BookmarkMap = ReadonlyMap<string, BookmarkData>;
type SeqMap = ReadonlyMap<number, number>;

// =============================================================================
// Field Instruction Parsing
// =============================================================================

/** Parse a field instruction string to get the field type and arguments. */
function parseFieldInstruction(instruction: string): {
  type: string;
  args: string;
} {
  const trimmed = instruction.trim();

  // Special case: formula fields start with "=" (may or may not have space after)
  if (trimmed.startsWith("=")) {
    return { type: "=", args: trimmed.slice(1).trim() };
  }

  const spaceIdx = trimmed.indexOf(" ");
  if (spaceIdx === -1) {
    return { type: trimmed.toUpperCase(), args: "" };
  }
  return {
    type: trimmed.slice(0, spaceIdx).toUpperCase(),
    args: trimmed.slice(spaceIdx + 1).trim()
  };
}

/** Parse TOC instruction to extract outline level range. */
function parseTocLevels(args: string): { min: number; max: number } {
  // Match \o "1-3" pattern
  const match = /\\o\s*"(\d+)-(\d+)"/.exec(args);
  if (match) {
    return { min: parseInt(match[1], 10), max: parseInt(match[2], 10) };
  }
  // Default: all heading levels 1-9
  return { min: 1, max: 9 };
}

/** Parse SEQ switches: \h (hidden), \r N (reset to N), \c (repeat). */
function parseSeqSwitches(args: string): {
  identifier: string;
  hidden: boolean;
  resetTo: number | null;
  repeat: boolean;
} {
  let identifier = "";
  let hidden = false;
  let resetTo: number | null = null;
  let repeat = false;

  // Extract identifier (first word before any switches)
  const parts = args.split(/\s+/);
  if (parts.length > 0 && !parts[0].startsWith("\\")) {
    identifier = parts[0];
  }

  if (/\\h\b/i.test(args)) {
    hidden = true;
  }
  if (/\\c\b/i.test(args)) {
    repeat = true;
  }
  const resetMatch = /\\r\s+(\d+)/i.exec(args);
  if (resetMatch) {
    resetTo = parseInt(resetMatch[1], 10);
  }

  return { identifier, hidden, resetTo, repeat };
}

/** Parse IF field: IF expr1 op expr2 "trueText" "falseText" */
function parseIfField(args: string): {
  left: string;
  operator: string;
  right: string;
  trueText: string;
  falseText: string;
} | null {
  // Pattern: expr1 = expr2 "trueText" "falseText"
  // Supports both quoted and unquoted operands.
  // Two-character operators must come first or `<=` would match `<` only.
  const match = /^"?([^"=<>!]*?)"?\s*(<=|>=|<>|=|<|>)\s*"?([^"]*?)"?\s+"([^"]*)"\s+"([^"]*)"/.exec(
    args
  );
  if (match) {
    return {
      left: match[1].trim(),
      operator: match[2],
      right: match[3].trim(),
      trueText: match[4],
      falseText: match[5]
    };
  }

  // Simpler pattern without quotes on operands
  const simpleMatch = /^(\S+)\s*(<=|>=|<>|=|<|>)\s*(\S+)\s+"([^"]*)"\s+"([^"]*)"/.exec(args);
  if (simpleMatch) {
    return {
      left: simpleMatch[1],
      operator: simpleMatch[2],
      right: simpleMatch[3],
      trueText: simpleMatch[4],
      falseText: simpleMatch[5]
    };
  }

  return null;
}

/** Parse STYLEREF field to get the style name. */
function parseStyleRef(args: string): string {
  // STYLEREF "StyleName" or STYLEREF StyleName
  const quotedMatch = /^"([^"]+)"/.exec(args);
  if (quotedMatch) {
    return quotedMatch[1];
  }
  const parts = args.split(/\s+/);
  return parts[0] ?? "";
}

// =============================================================================
// Bookmark Collection
// =============================================================================

/** Collect all bookmarks with their text and page numbers. */
function collectBookmarkInfo(doc: DocxDocument, layout: LayoutResult): Map<string, BookmarkData> {
  const map = new Map<string, BookmarkData>();
  const { contentPages } = layout;

  for (let i = 0; i < doc.body.length; i++) {
    const item = doc.body[i];
    if (item.type === "paragraph") {
      const text = extractParagraphText(item);
      for (const child of item.children) {
        if (isBookmarkStart(child)) {
          const page = layout.bookmarkPages.get(child.name) ?? contentPages[i] ?? 1;
          map.set(child.name, { text, page });
        }
      }
    }
  }

  return map;
}

function isBookmarkStart(child: ParagraphChild): child is BookmarkStart {
  return "type" in child && (child as { type: string }).type === "bookmarkStart";
}

// =============================================================================
// Heading Collection
// =============================================================================

/** Collect headings from the document for TOC generation. */
function collectHeadings(doc: DocxDocument, layout: LayoutResult): HeadingEntry[] {
  const headings: HeadingEntry[] = [];
  const { contentPages } = layout;

  for (let i = 0; i < doc.body.length; i++) {
    const item = doc.body[i];
    if (item.type !== "paragraph") {
      continue;
    }

    const level = getHeadingLevel(item, doc.styles);
    if (level !== null) {
      headings.push({
        text: extractParagraphText(item),
        level,
        page: contentPages[i] ?? 1
      });
    }
  }

  return headings;
}

/** Determine if a paragraph is a heading and return its level (1-9), or null. */
function getHeadingLevel(para: Paragraph, styles?: DocxDocument["styles"]): number | null {
  const props = para.properties;

  // Check outlineLevel directly on paragraph
  if (props?.outlineLevel != null && props.outlineLevel >= 0 && props.outlineLevel <= 8) {
    return props.outlineLevel + 1; // outlineLevel is 0-based, heading level is 1-based
  }

  // Check style name
  const styleId = props?.style;
  if (styleId) {
    // Direct pattern match: Heading1, Heading2, etc.
    const headingMatch = /^[Hh]eading(\d)$/.exec(styleId);
    if (headingMatch) {
      return parseInt(headingMatch[1], 10);
    }

    // Look up in styles array
    if (styles) {
      const styleDef = styles.find(s => s.styleId === styleId);
      if (styleDef) {
        // Check style name pattern
        const nameMatch = /^[Hh]eading\s*(\d)$/.exec(styleDef.name);
        if (nameMatch) {
          return parseInt(nameMatch[1], 10);
        }
        // Check outlineLevel on style
        if (
          styleDef.outlineLevel != null &&
          styleDef.outlineLevel >= 0 &&
          styleDef.outlineLevel <= 8
        ) {
          return styleDef.outlineLevel + 1;
        }
      }
    }
  }

  return null;
}

// =============================================================================
// SEQ Value Computation
// =============================================================================

/** Compute all SEQ field values by traversing the document in order. */
function computeSeqValues(doc: DocxDocument): Map<number, number> {
  const counters = new Map<string, number>();
  const values = new Map<number, number>();
  let fieldIndex = 0;

  for (const item of doc.body) {
    if (item.type === "paragraph") {
      fieldIndex = processSeqInParagraph(item, counters, values, fieldIndex);
    }
  }

  return values;
}

function processSeqInParagraph(
  para: Paragraph,
  counters: Map<string, number>,
  values: Map<number, number>,
  startIndex: number
): number {
  let fieldIndex = startIndex;

  for (const child of para.children) {
    if (!isRun(child)) {
      continue;
    }
    for (const content of child.content) {
      if (content.type !== "field") {
        continue;
      }

      const { type, args } = parseFieldInstruction(content.instruction);
      if (type !== "SEQ") {
        fieldIndex++;
        continue;
      }

      const { identifier, hidden: _hidden, resetTo, repeat } = parseSeqSwitches(args);
      if (!identifier) {
        fieldIndex++;
        continue;
      }

      if (resetTo !== null) {
        counters.set(identifier, resetTo);
      }

      if (repeat) {
        // \c: repeat last value without incrementing
        const current = counters.get(identifier) ?? 0;
        values.set(fieldIndex, current);
      } else {
        const current = (counters.get(identifier) ?? 0) + 1;
        counters.set(identifier, current);
        values.set(fieldIndex, current);
      }

      fieldIndex++;
    }
  }

  return fieldIndex;
}

// =============================================================================
// Style Index (for STYLEREF)
// =============================================================================

/** Build index from style ID → paragraph entries with their body index. */
function buildStyleIndex(doc: DocxDocument): Map<string, { index: number; text: string }[]> {
  const map = new Map<string, { index: number; text: string }[]>();

  for (let i = 0; i < doc.body.length; i++) {
    const item = doc.body[i];
    if (item.type !== "paragraph") {
      continue;
    }

    const styleId = item.properties?.style;
    if (!styleId) {
      continue;
    }

    let entries = map.get(styleId);
    if (!entries) {
      entries = [];
      map.set(styleId, entries);
    }
    entries.push({ index: i, text: extractParagraphText(item) });
  }

  return map;
}

/** Find the nearest paragraph with the given style before bodyIndex. */
function findStyleRef(
  styleIndex: Map<string, { index: number; text: string }[]>,
  styleName: string,
  bodyIndex: number
): string {
  const entries = styleIndex.get(styleName);
  if (!entries || entries.length === 0) {
    return "";
  }

  // Find the last entry whose index <= bodyIndex
  let best: { index: number; text: string } | undefined;
  for (const entry of entries) {
    if (entry.index <= bodyIndex) {
      best = entry;
    } else {
      break;
    }
  }

  return best?.text ?? entries[0].text ?? "";
}

// =============================================================================
// Index Entry Collection (for INDEX field)
// =============================================================================

/** Collect all XE (Index Entry) fields from the document body. */
function collectIndexEntries(doc: DocxDocument): IndexEntry[] {
  const entries: IndexEntry[] = [];

  walkBlocks(doc.body, {
    visitRunContent(content) {
      if (content.type !== "field") {
        return;
      }
      const { type, args } = parseFieldInstruction(content.instruction);
      if (type === "XE") {
        const term = parseXeTerm(args);
        if (term) {
          entries.push({ term, page: 1 });
        }
      }
    }
  });

  return entries;
}

/** Parse the term from an XE field argument: XE "term" or XE term. */
function parseXeTerm(args: string): string {
  const quotedMatch = /^"([^"]+)"/.exec(args);
  if (quotedMatch) {
    return quotedMatch[1];
  }
  const raw = args.split(/\s+/)[0];
  return raw ?? "";
}

/** Build index content from collected entries. */
function buildIndexContent(entries: IndexEntry[], args: string): string {
  if (entries.length === 0) {
    return "";
  }

  // Sort entries alphabetically by term
  const sorted = [...entries].sort((a, b) => a.term.localeCompare(b.term));

  // Check for \h switch (group by first letter with headings)
  const grouped = /\\h\b/i.test(args);

  if (!grouped) {
    return sorted.map(e => `${e.term}\t${e.page}`).join("\n");
  }

  // Group by first letter
  const lines: string[] = [];
  let currentLetter = "";

  for (const entry of sorted) {
    const letter = entry.term.charAt(0).toUpperCase();
    if (letter !== currentLetter) {
      currentLetter = letter;
      if (lines.length > 0) {
        lines.push("");
      }
      lines.push(currentLetter);
    }
    lines.push(`${entry.term}\t${entry.page}`);
  }

  return lines.join("\n");
}

// =============================================================================
// Formula Evaluation (for = field)
// =============================================================================

/**
 * Evaluate a simple math formula expression.
 * Supports: +, -, *, /, (), numbers, and SUM(ABOVE)/SUM(LEFT).
 */
function evaluateFormula(args: string, bodyIndex: number, doc: DocxDocument): string {
  // Extract the number format switch \# "format"
  const formatMatch = /\\#\s*"([^"]+)"/.exec(args);
  const format = formatMatch ? formatMatch[1] : null;

  // Remove format switch from expression
  let expr = args.replace(/\\#\s*"[^"]*"/, "").trim();

  // Handle SUM(ABOVE) and SUM(LEFT)
  expr = expr.replace(/SUM\s*\(\s*ABOVE\s*\)/gi, () => {
    return String(sumAbove(bodyIndex, doc));
  });
  expr = expr.replace(/SUM\s*\(\s*LEFT\s*\)/gi, () => {
    return String(sumLeft(bodyIndex, doc));
  });

  // Evaluate the arithmetic expression
  const result = evalArithmetic(expr);
  if (result === null) {
    return "!";
  }

  return format ? formatNumber(result, format) : String(result);
}

/** Evaluate a simple arithmetic expression with +, -, *, /, (). */
function evalArithmetic(expr: string): number | null {
  // Remove whitespace
  const tokens = tokenizeExpression(expr.replace(/\s+/g, ""));
  if (tokens === null) {
    return null;
  }

  try {
    const result = parseExpression(tokens, { pos: 0 });
    return result;
  } catch {
    return null;
  }
}

/** Tokenize arithmetic expression into numbers, operators, and parens. */
function tokenizeExpression(expr: string): string[] | null {
  const tokens: string[] = [];
  let i = 0;

  while (i < expr.length) {
    const ch = expr[i];

    if (ch === "+" || ch === "-" || ch === "*" || ch === "/" || ch === "(" || ch === ")") {
      // Handle negative numbers: minus after operator or at start
      if (
        ch === "-" &&
        (tokens.length === 0 ||
          tokens[tokens.length - 1] === "(" ||
          isOperator(tokens[tokens.length - 1]))
      ) {
        let num = "-";
        i++;
        while (i < expr.length && isDigitOrDot(expr[i])) {
          num += expr[i];
          i++;
        }
        if (num === "-") {
          return null;
        }
        tokens.push(num);
      } else {
        tokens.push(ch);
        i++;
      }
    } else if (isDigitOrDot(ch)) {
      let num = "";
      while (i < expr.length && isDigitOrDot(expr[i])) {
        num += expr[i];
        i++;
      }
      tokens.push(num);
    } else {
      return null; // Invalid character
    }
  }

  return tokens;
}

function isDigitOrDot(ch: string): boolean {
  return (ch >= "0" && ch <= "9") || ch === ".";
}

function isOperator(token: string | undefined): boolean {
  return token === "+" || token === "-" || token === "*" || token === "/";
}

/** Recursive descent parser for arithmetic expressions. */
function parseExpression(tokens: string[], ctx: { pos: number }): number {
  let left = parseTerm(tokens, ctx);

  while (ctx.pos < tokens.length) {
    const op = tokens[ctx.pos];
    if (op !== "+" && op !== "-") {
      break;
    }
    ctx.pos++;
    const right = parseTerm(tokens, ctx);
    left = op === "+" ? left + right : left - right;
  }

  return left;
}

function parseTerm(tokens: string[], ctx: { pos: number }): number {
  let left = parseFactor(tokens, ctx);

  while (ctx.pos < tokens.length) {
    const op = tokens[ctx.pos];
    if (op !== "*" && op !== "/") {
      break;
    }
    ctx.pos++;
    const right = parseFactor(tokens, ctx);
    left = op === "*" ? left * right : left / right;
  }

  return left;
}

function parseFactor(tokens: string[], ctx: { pos: number }): number {
  const token = tokens[ctx.pos];

  if (token === "(") {
    ctx.pos++;
    const result = parseExpression(tokens, ctx);
    if (tokens[ctx.pos] === ")") {
      ctx.pos++;
    }
    return result;
  }

  ctx.pos++;
  const num = parseFloat(token);
  if (isNaN(num)) {
    throw new Error("Invalid number");
  }
  return num;
}

/** Sum numeric values from cells above the current cell in a table. */
function sumAbove(bodyIndex: number, doc: DocxDocument): number {
  const location = findCellInTable(bodyIndex, doc);
  if (!location) {
    return 0;
  }

  const { table, rowIndex, colIndex } = location;
  let sum = 0;

  for (let r = 0; r < rowIndex; r++) {
    const cell = table.rows[r]?.cells[colIndex];
    if (cell) {
      sum += extractCellNumericValue(cell);
    }
  }

  return sum;
}

/** Sum numeric values from cells to the left of the current cell in a table. */
function sumLeft(bodyIndex: number, doc: DocxDocument): number {
  const location = findCellInTable(bodyIndex, doc);
  if (!location) {
    return 0;
  }

  const { table, rowIndex, colIndex } = location;
  const row = table.rows[rowIndex];
  let sum = 0;

  for (let c = 0; c < colIndex; c++) {
    const cell = row?.cells[c];
    if (cell) {
      sum += extractCellNumericValue(cell);
    }
  }

  return sum;
}

/**
 * Find which table/row/col a given body-level paragraph belongs to.
 * Searches top-level tables in the body for a paragraph that contains a formula field
 * at the given body index.
 */
function findCellInTable(
  bodyIndex: number,
  doc: DocxDocument
): { table: Table; rowIndex: number; colIndex: number } | null {
  // The formula field is inside a table — we scan all tables looking for a match.
  // We track body items: tables contain paragraphs internally but bodyIndex here is
  // the top-level index. When the body item at bodyIndex is a table itself, we search within.
  // However, our engine flattens paragraphs: if the body item IS a table, it won't be processed
  // as a paragraph. So we look at surrounding tables to find where the field might be.

  // Strategy: walk backwards from bodyIndex to find the enclosing table.
  // But in our document model, table cells' paragraphs are not at the top level.
  // Instead, formula fields inside tables will be visited via the body flattening that the
  // engine does not currently do. For simplicity, we search all tables for a cell matching.

  for (const item of doc.body) {
    if (item.type !== "table") {
      continue;
    }
    for (let r = 0; r < item.rows.length; r++) {
      const row = item.rows[r];
      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        let found = false;
        walkBlocks(cell.content as readonly BodyContent[], {
          visitRunContent(content) {
            if (content.type === "field") {
              const { type } = parseFieldInstruction(content.instruction);
              if (type === "=") {
                found = true;
              }
            }
          }
        });
        if (found) {
          // Heuristic: return first match since we can't precisely map bodyIndex to cell
          // In practice, the engine would need cell-level tracking; this is a best-effort approach.
          return { table: item, rowIndex: r, colIndex: c };
        }
      }
    }
  }

  return null;
}

/** Extract numeric value from a table cell (first number found in text). */
function extractCellNumericValue(cell: TableCell): number {
  let text = "";
  for (const content of cell.content) {
    if (content.type === "paragraph") {
      text += extractParagraphText(content);
    }
  }

  // Try to parse the entire text as a number (with possible comma separators)
  const cleaned = text.replace(/,/g, "").trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

/** Format a number according to a simple format string (e.g. "#,##0", "0.00"). */
function formatNumber(value: number, format: string): string {
  // Determine decimal places from format
  const dotIndex = format.indexOf(".");
  let decimals = 0;
  if (dotIndex !== -1) {
    decimals = format.length - dotIndex - 1;
  }

  // Check if format uses thousand separators
  const useThousands = format.includes(",");

  if (useThousands) {
    const fixed = value.toFixed(decimals);
    const [intPart, decPart] = fixed.split(".");
    const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return decPart ? `${withCommas}.${decPart}` : withCommas;
  }

  return value.toFixed(decimals);
}

// =============================================================================
// Body Update
// =============================================================================

function updateBody(
  doc: DocxDocument,
  layout: LayoutResult,
  bookmarkInfo: BookmarkMap,
  seqValues: SeqMap,
  styleIndex: Map<string, { index: number; text: string }[]>,
  indexEntries: IndexEntry[],
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions }
): readonly BodyContent[] {
  const body = doc.body;
  let changed = false;
  let globalFieldIndex = 0;
  const newBody: BodyContent[] = [];

  for (let i = 0; i < body.length; i++) {
    const item = body[i];

    if (item.type === "paragraph") {
      const result = updateParagraphFields(
        item,
        i,
        layout,
        bookmarkInfo,
        seqValues,
        styleIndex,
        indexEntries,
        doc,
        globalFieldIndex,
        opts
      );
      globalFieldIndex = result.nextFieldIndex;
      if (result.paragraph !== item) {
        changed = true;
      }
      newBody.push(result.paragraph);
    } else if (item.type === "tableOfContents" && opts.updateToc) {
      const headings = collectHeadingsFromLayout(body, layout);
      const updated = updateTocContent(item, headings);
      if (updated !== item) {
        changed = true;
      }
      newBody.push(updated);
    } else if (item.type === "table") {
      // Recurse into table cells to update fields within tables
      const updatedTable = updateTableFields(
        item,
        i,
        layout,
        bookmarkInfo,
        seqValues,
        styleIndex,
        indexEntries,
        doc,
        globalFieldIndex,
        opts
      );
      if (updatedTable.table !== item) {
        changed = true;
      }
      globalFieldIndex = updatedTable.nextFieldIndex;
      newBody.push(updatedTable.table);
    } else {
      newBody.push(item);
    }
  }

  return changed ? newBody : body;
}

/** Recursively update fields inside a table's cells. */
function updateTableFields(
  table: Table,
  bodyIndex: number,
  layout: LayoutResult,
  bookmarkInfo: BookmarkMap,
  seqValues: SeqMap,
  styleIndex: Map<string, { index: number; text: string }[]>,
  indexEntries: IndexEntry[],
  doc: DocxDocument,
  startFieldIndex: number,
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions }
): { table: Table; nextFieldIndex: number } {
  let fieldIndex = startFieldIndex;
  let tableChanged = false;
  const newRows = table.rows.map(row => {
    let rowChanged = false;
    const newCells = row.cells.map(cell => {
      let cellChanged = false;
      const newContent: (Paragraph | Table)[] = [];
      for (const block of cell.content) {
        if (block.type === "paragraph") {
          const result = updateParagraphFields(
            block,
            bodyIndex,
            layout,
            bookmarkInfo,
            seqValues,
            styleIndex,
            indexEntries,
            doc,
            fieldIndex,
            opts
          );
          fieldIndex = result.nextFieldIndex;
          if (result.paragraph !== block) {
            cellChanged = true;
          }
          newContent.push(result.paragraph);
        } else if (block.type === "table") {
          const nested = updateTableFields(
            block,
            bodyIndex,
            layout,
            bookmarkInfo,
            seqValues,
            styleIndex,
            indexEntries,
            doc,
            fieldIndex,
            opts
          );
          fieldIndex = nested.nextFieldIndex;
          if (nested.table !== block) {
            cellChanged = true;
          }
          newContent.push(nested.table);
        } else {
          newContent.push(block);
        }
      }
      if (cellChanged) {
        rowChanged = true;
        return { ...cell, content: newContent } as TableCell;
      }
      return cell;
    });
    if (rowChanged) {
      tableChanged = true;
      return { ...row, cells: newCells };
    }
    return row;
  });

  if (tableChanged) {
    return { table: { ...table, rows: newRows }, nextFieldIndex: fieldIndex };
  }
  return { table, nextFieldIndex: fieldIndex };
}

/** Collect headings directly from body with layout info (avoiding a second layout pass). */
function collectHeadingsFromLayout(
  body: readonly BodyContent[],
  layout: LayoutResult
): HeadingEntry[] {
  const headings: HeadingEntry[] = [];
  const { contentPages } = layout;

  for (let i = 0; i < body.length; i++) {
    const item = body[i];
    if (item.type !== "paragraph") {
      continue;
    }

    const level = getHeadingLevelSimple(item);
    if (level !== null) {
      headings.push({
        text: extractParagraphText(item),
        level,
        page: contentPages[i] ?? 1
      });
    }
  }

  return headings;
}

/** Simple heading level detection without full style lookup. */
function getHeadingLevelSimple(para: Paragraph): number | null {
  const props = para.properties;

  if (props?.outlineLevel != null && props.outlineLevel >= 0 && props.outlineLevel <= 8) {
    return props.outlineLevel + 1;
  }

  const styleId = props?.style;
  if (styleId) {
    const match = /^[Hh]eading(\d)$/.exec(styleId);
    if (match) {
      return parseInt(match[1], 10);
    }
  }

  return null;
}

// =============================================================================
// Paragraph Field Update
// =============================================================================

function updateParagraphFields(
  para: Paragraph,
  bodyIndex: number,
  layout: LayoutResult,
  bookmarkInfo: BookmarkMap,
  seqValues: SeqMap,
  styleIndex: Map<string, { index: number; text: string }[]>,
  indexEntries: IndexEntry[],
  doc: DocxDocument,
  startFieldIndex: number,
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions }
): { paragraph: Paragraph; nextFieldIndex: number } {
  let fieldIndex = startFieldIndex;
  let childrenChanged = false;
  const newChildren: ParagraphChild[] = [];

  for (const child of para.children) {
    if (!isRun(child)) {
      newChildren.push(child);
      continue;
    }

    const runResult = updateRunFields(
      child,
      bodyIndex,
      layout,
      bookmarkInfo,
      seqValues,
      styleIndex,
      indexEntries,
      doc,
      fieldIndex,
      opts
    );
    fieldIndex = runResult.nextFieldIndex;
    if (runResult.run !== child) {
      childrenChanged = true;
    }
    newChildren.push(runResult.run);
  }

  if (!childrenChanged) {
    return { paragraph: para, nextFieldIndex: fieldIndex };
  }

  return {
    paragraph: { ...para, children: newChildren },
    nextFieldIndex: fieldIndex
  };
}

// =============================================================================
// Run Field Update
// =============================================================================

function updateRunFields(
  run: Run,
  bodyIndex: number,
  layout: LayoutResult,
  bookmarkInfo: BookmarkMap,
  seqValues: SeqMap,
  styleIndex: Map<string, { index: number; text: string }[]>,
  indexEntries: IndexEntry[],
  doc: DocxDocument,
  startFieldIndex: number,
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions }
): { run: Run; nextFieldIndex: number } {
  let fieldIndex = startFieldIndex;
  let contentChanged = false;
  const newContent: RunContent[] = [];

  for (const item of run.content) {
    if (item.type !== "field") {
      newContent.push(item);
      continue;
    }

    const updated = computeFieldValue(
      item,
      bodyIndex,
      layout,
      bookmarkInfo,
      seqValues,
      styleIndex,
      indexEntries,
      doc,
      fieldIndex,
      opts
    );
    fieldIndex++;

    if (updated !== item) {
      contentChanged = true;
    }
    newContent.push(updated);
  }

  if (!contentChanged) {
    return { run, nextFieldIndex: fieldIndex };
  }

  return {
    run: { ...run, content: newContent },
    nextFieldIndex: fieldIndex
  };
}

// =============================================================================
// Field Value Computation
// =============================================================================

function computeFieldValue(
  field: FieldContent,
  bodyIndex: number,
  layout: LayoutResult,
  bookmarkInfo: BookmarkMap,
  seqValues: SeqMap,
  styleIndex: Map<string, { index: number; text: string }[]>,
  indexEntries: IndexEntry[],
  doc: DocxDocument,
  fieldIndex: number,
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions }
): FieldContent {
  const { type, args } = parseFieldInstruction(field.instruction);

  switch (type) {
    case "PAGE": {
      if (!opts.updatePageFields) {
        return field;
      }
      const page = layout.contentPages[bodyIndex] ?? 1;
      const value = String(page);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "NUMPAGES": {
      if (!opts.updatePageFields) {
        return field;
      }
      const value = String(layout.pageCount);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "SECTIONPAGES": {
      if (!opts.updatePageFields) {
        return field;
      }
      const sectionIdx = layout.contentSections[bodyIndex] ?? 0;
      const pages = layout.sectionPageCounts[sectionIdx] ?? 1;
      const value = String(pages);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "SECTION": {
      if (!opts.updatePageFields) {
        return field;
      }
      const sectionIdx = layout.contentSections[bodyIndex] ?? 0;
      const value = String(sectionIdx + 1);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "REF": {
      if (!opts.updateReferences) {
        return field;
      }
      const bookmarkName = args.split(/\s+/)[0] ?? "";
      const data = bookmarkInfo.get(bookmarkName);
      const value = data?.text ?? "";
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "PAGEREF": {
      if (!opts.updateReferences) {
        return field;
      }
      const bookmarkName = args.split(/\s+/)[0] ?? "";
      const page = layout.bookmarkPages.get(bookmarkName);
      const value = page != null ? String(page) : "?";
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "SEQ": {
      if (!opts.updateSequences) {
        return field;
      }
      const seqVal = seqValues.get(fieldIndex);
      if (seqVal == null) {
        return field;
      }
      const value = String(seqVal);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "IF": {
      const parsed = parseIfField(args);
      if (!parsed) {
        return field;
      }
      const result = evaluateCondition(parsed.left, parsed.operator, parsed.right);
      const value = result ? parsed.trueText : parsed.falseText;
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "STYLEREF": {
      const styleName = parseStyleRef(args);
      if (!styleName) {
        return field;
      }
      const value = findStyleRef(styleIndex, styleName, bodyIndex);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "INDEX": {
      const value = buildIndexContent(indexEntries, args);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "=": {
      const value = evaluateFormula(args, bodyIndex, doc);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "INCLUDETEXT": {
      // INCLUDETEXT requires file system access; mark as not evaluable
      const path = args.split(/\s+/)[0] ?? "";
      const value = `[INCLUDETEXT: ${path}]`;
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    default:
      return field;
  }
}

/** Evaluate a simple comparison between two string values. */
function evaluateCondition(left: string, operator: string, right: string): boolean {
  switch (operator) {
    case "=":
      return left === right;
    case "<>":
      return left !== right;
    case "<":
      return left < right;
    case ">":
      return left > right;
    case "<=":
      return left <= right;
    case ">=":
      return left >= right;
    default:
      return false;
  }
}

// =============================================================================
// TOC Update
// =============================================================================

/** Update a TableOfContents body content with generated cached paragraphs. */
function updateTocContent(toc: TableOfContents, headings: HeadingEntry[]): TableOfContents {
  const { min, max } = parseTocLevels(buildTocInstruction(toc));

  // Filter headings by level range
  const filtered = headings.filter(h => h.level >= min && h.level <= max);

  // Generate cached paragraphs for each heading
  const cachedParagraphs: Paragraph[] = filtered.map(heading => buildTocEntryParagraph(heading));

  // Check if content actually changed
  if (
    toc.cachedParagraphs &&
    toc.cachedParagraphs.length === cachedParagraphs.length &&
    cachedParagraphs.every(
      (p, i) => extractParagraphText(p) === extractParagraphText(toc.cachedParagraphs![i])
    )
  ) {
    return toc;
  }

  return { ...toc, cachedParagraphs };
}

/** Reconstruct a TOC instruction string from TOC properties. */
function buildTocInstruction(toc: TableOfContents): string {
  let instruction = "TOC";
  if (toc.headingStyleRange) {
    instruction += ` \\o "${toc.headingStyleRange}"`;
  }
  if (toc.hyperlink) {
    instruction += " \\h";
  }
  return instruction;
}

/** Build a single TOC entry paragraph. */
function buildTocEntryParagraph(heading: HeadingEntry): Paragraph {
  // Build: "HeadingText\tPageNumber"
  const textRun: Run = {
    content: [
      { type: "text", text: heading.text },
      { type: "tab" },
      { type: "text", text: String(heading.page) }
    ]
  };

  return {
    type: "paragraph",
    properties: {
      style: `TOC${heading.level}`
    },
    children: [textRun]
  };
}
