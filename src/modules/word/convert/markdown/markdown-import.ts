/**
 * DOCX Module - Markdown to DOCX Converter
 *
 * Converts a GFM (GitHub Flavored Markdown) string into DOCX document body content.
 * Handles common Markdown elements: headings, paragraphs, bold, italic, strikethrough,
 * code, links, images, lists, tables, blockquotes, horizontal rules, and fenced code blocks.
 *
 * @example
 * ```ts
 * import { markdownToDocx } from "excelts/word/markdown";
 * import { Document, toBuffer } from "excelts/word";
 *
 * const doc = markdownToDocx("# Hello\n\nWorld **bold**");
 * const buffer = await toBuffer(doc);
 * ```
 *
 * @stability experimental
 */

import { sanitizeUrl } from "../../core/internal-utils";
import type {
  AbstractNumbering,
  Alignment,
  BodyContent,
  DocxDocument,
  Hyperlink,
  LevelSuffix,
  NumberingInstance,
  Paragraph,
  ParagraphChild,
  ParagraphProperties,
  Run,
  RunProperties,
  Table,
  TableCell,
  TableCellProperties,
  TableProperties,
  TableRow,
  TableWidth
} from "../../types";

// =============================================================================
// Public API
// =============================================================================

/** Options for Markdown to DOCX conversion. */
export interface MarkdownImportOptions {
  /** Default font family for body text. */
  readonly defaultFont?: string;
  /** Default font size in half-points (default: 24 = 12pt). */
  readonly defaultFontSize?: number;
  /** Code font family (default: "Courier New"). */
  readonly codeFont?: string;
  /** Code font size in half-points (default: 20 = 10pt). */
  readonly codeFontSize?: number;
  /** Custom image resolver — given a URL, return image data or undefined to skip. */
  readonly resolveImage?: (
    url: string,
    alt: string
  ) => MarkdownImageData | undefined | Promise<MarkdownImageData | undefined>;
}

/** Resolved image data for embedding. */
export interface MarkdownImageData {
  readonly data: Uint8Array;
  readonly mediaType: "png" | "jpeg" | "gif" | "bmp" | "tiff" | "svg" | "webp";
  readonly width?: number; // EMU
  readonly height?: number; // EMU
}

/**
 * Convert a Markdown string into a complete DocxDocument.
 *
 * @param markdown - The GFM Markdown string.
 * @param options - Optional conversion settings.
 * @returns A DocxDocument ready to be packaged.
 */
export function markdownToDocx(markdown: string, options?: MarkdownImportOptions): DocxDocument {
  const { body, state } = markdownToDocxBodyInternal(markdown, options);
  return {
    body,
    styles: defaultMarkdownStyles(),
    abstractNumberings: state.abstractNumberings,
    numberingInstances: state.numberingInstances
  };
}

/**
 * Convert a Markdown string into an array of DOCX body content blocks.
 *
 * @param markdown - The GFM Markdown string.
 * @param options - Optional conversion settings.
 * @returns Array of BodyContent blocks.
 */
export function markdownToDocxBody(
  markdown: string,
  options?: MarkdownImportOptions
): BodyContent[] {
  return markdownToDocxBodyInternal(markdown, options).body;
}

/**
 * Internal implementation: converts markdown and returns both body and state.
 */
function markdownToDocxBodyInternal(
  markdown: string,
  options?: MarkdownImportOptions
): { body: BodyContent[]; state: ConversionState } {
  const state = createState();
  const opts: Required<Pick<MarkdownImportOptions, "codeFont" | "codeFontSize">> &
    MarkdownImportOptions = {
    codeFont: "Courier New",
    codeFontSize: 20,
    ...options
  };

  const lines = markdown.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const blocks = parseMarkdownBlocks(lines, 0, lines.length);
  const body = convertBlocks(blocks, opts, state);
  return { body, state };
}

// =============================================================================
// AST Types (internal)
// =============================================================================

interface HeadingBlock {
  type: "heading";
  level: 1 | 2 | 3 | 4 | 5 | 6;
  inlines: InlineNode[];
}

interface ParagraphBlock {
  type: "paragraph";
  inlines: InlineNode[];
}

interface BlockquoteBlock {
  type: "blockquote";
  children: Block[];
}

interface FencedCodeBlock {
  type: "fencedCode";
  language: string;
  code: string;
}

interface ThematicBreakBlock {
  type: "thematicBreak";
}

interface ListBlock {
  type: "list";
  ordered: boolean;
  start: number;
  items: ListItemBlock[];
}

interface ListItemBlock {
  type: "listItem";
  checked?: boolean; // for task lists
  children: Block[];
}

interface TableBlock {
  type: "table";
  headers: InlineNode[][];
  alignments: (Alignment | undefined)[];
  rows: InlineNode[][][];
}

interface HtmlBlock {
  type: "html";
  content: string;
}

type Block =
  | HeadingBlock
  | ParagraphBlock
  | BlockquoteBlock
  | FencedCodeBlock
  | ThematicBreakBlock
  | ListBlock
  | TableBlock
  | HtmlBlock;

// Inline AST
interface TextInline {
  type: "text";
  text: string;
}
interface BoldInline {
  type: "bold";
  children: InlineNode[];
}
interface ItalicInline {
  type: "italic";
  children: InlineNode[];
}
interface StrikethroughInline {
  type: "strikethrough";
  children: InlineNode[];
}
interface CodeInline {
  type: "code";
  text: string;
}
interface LinkInline {
  type: "link";
  url: string;
  title?: string;
  children: InlineNode[];
}
interface ImageInline {
  type: "image";
  url: string;
  alt: string;
  title?: string;
}
interface LineBreakInline {
  type: "lineBreak";
}

type InlineNode =
  | TextInline
  | BoldInline
  | ItalicInline
  | StrikethroughInline
  | CodeInline
  | LinkInline
  | ImageInline
  | LineBreakInline;

// =============================================================================
// Shared state for numbering (per-invocation)
// =============================================================================

interface ConversionState {
  abstractNumberings: AbstractNumbering[];
  numberingInstances: NumberingInstance[];
  nextNumId: number;
  bulletNumId: number | undefined;
  orderedNumId: number | undefined;
}

function createState(): ConversionState {
  return {
    abstractNumberings: [],
    numberingInstances: [],
    nextNumId: 1,
    bulletNumId: undefined,
    orderedNumId: undefined
  };
}

// =============================================================================
// Block Parser
// =============================================================================

function parseMarkdownBlocks(lines: string[], start: number, end: number): Block[] {
  const blocks: Block[] = [];
  let i = start;

  while (i < end) {
    const line = lines[i];

    // Blank line — skip
    if (line.trim() === "") {
      i++;
      continue;
    }

    // Fenced code block (``` or ~~~)
    const fenceMatch = line.match(/^(`{3,}|~{3,})(.*)$/);
    if (fenceMatch) {
      const fence = fenceMatch[1];
      const language = fenceMatch[2].trim();
      const codeLines: string[] = [];
      i++;
      while (i < end) {
        const trimmed = lines[i].trim();
        // CommonMark: closing fence must use same char and be at least as long
        if (
          trimmed.length >= fence.length &&
          new RegExp(`^${fence[0]}{${fence.length},}$`).test(trimmed)
        ) {
          i++;
          break;
        }
        codeLines.push(lines[i]);
        i++;
      }
      blocks.push({ type: "fencedCode", language, code: codeLines.join("\n") });
      continue;
    }

    // ATX Heading (# to ######)
    const headingMatch = line.match(/^(#{1,6})\s+(.+?)(?:\s+#+)?$/);
    if (headingMatch) {
      const level = headingMatch[1].length as 1 | 2 | 3 | 4 | 5 | 6;
      const inlines = parseInlines(headingMatch[2].trim());
      blocks.push({ type: "heading", level, inlines });
      i++;
      continue;
    }

    // Setext heading (underline with === or ---)
    if (i + 1 < end) {
      const nextLine = lines[i + 1];
      if (/^={3,}\s*$/.test(nextLine)) {
        blocks.push({ type: "heading", level: 1, inlines: parseInlines(line.trim()) });
        i += 2;
        continue;
      }
      if (/^-{3,}\s*$/.test(nextLine) && !/^\s*[-*]\s/.test(line)) {
        blocks.push({ type: "heading", level: 2, inlines: parseInlines(line.trim()) });
        i += 2;
        continue;
      }
    }

    // Thematic break (---, ***, ___)
    if (/^(\*{3,}|-{3,}|_{3,})\s*$/.test(line)) {
      blocks.push({ type: "thematicBreak" });
      i++;
      continue;
    }

    // Blockquote
    if (line.startsWith(">")) {
      const quoteLines: string[] = [];
      while (
        i < end &&
        (lines[i].startsWith(">") || (lines[i].trim() !== "" && !isBlockStart(lines[i])))
      ) {
        if (lines[i].startsWith(">")) {
          quoteLines.push(lines[i].replace(/^>\s?/, ""));
        } else {
          quoteLines.push(lines[i]);
        }
        i++;
      }
      const children = parseMarkdownBlocks(quoteLines, 0, quoteLines.length);
      blocks.push({ type: "blockquote", children });
      continue;
    }

    // Unordered list
    if (/^(\s*)([-*+])\s/.test(line)) {
      const result = parseList(lines, i, end, false);
      blocks.push(result.block);
      i = result.nextIndex;
      continue;
    }

    // Ordered list
    if (/^(\s*)(\d+)[.)]\s/.test(line)) {
      const result = parseList(lines, i, end, true);
      blocks.push(result.block);
      i = result.nextIndex;
      continue;
    }

    // GFM Table
    if (i + 1 < end && isTableRow(line) && isTableSeparator(lines[i + 1])) {
      const result = parseTable(lines, i, end);
      blocks.push(result.block);
      i = result.nextIndex;
      continue;
    }

    // HTML block (raw HTML starting with <)
    if (/^<[a-zA-Z]/.test(line) && !line.startsWith("<a ") && !line.startsWith("<img ")) {
      const htmlLines: string[] = [];
      while (i < end && lines[i].trim() !== "") {
        htmlLines.push(lines[i]);
        i++;
      }
      blocks.push({ type: "html", content: htmlLines.join("\n") });
      continue;
    }

    // Paragraph (default) — collect consecutive non-blank, non-block-start lines
    const paraLines: string[] = [];
    while (i < end && lines[i].trim() !== "" && !isBlockStart(lines[i])) {
      paraLines.push(lines[i]);
      i++;
    }
    if (paraLines.length > 0) {
      const text = paraLines.join("\n");
      blocks.push({ type: "paragraph", inlines: parseInlines(text) });
    }
  }

  return blocks;
}

function isBlockStart(line: string): boolean {
  if (/^#{1,6}\s/.test(line)) {
    return true;
  }
  if (/^(`{3,}|~{3,})/.test(line)) {
    return true;
  }
  if (/^(\*{3,}|-{3,}|_{3,})\s*$/.test(line)) {
    return true;
  }
  if (line.startsWith(">")) {
    return true;
  }
  if (/^(\s*)([-*+])\s/.test(line)) {
    return true;
  }
  if (/^(\s*)(\d+)[.)]\s/.test(line)) {
    return true;
  }
  if (/^<[a-zA-Z]/.test(line) && !line.startsWith("<a ") && !line.startsWith("<img ")) {
    return true;
  }
  return false;
}

// =============================================================================
// List Parser
// =============================================================================

function parseList(
  lines: string[],
  start: number,
  end: number,
  ordered: boolean
): { block: ListBlock; nextIndex: number } {
  const items: ListItemBlock[] = [];
  let i = start;
  const startNum = ordered ? parseInt(lines[i].match(/^(\s*)(\d+)/)?.[2] ?? "1", 10) : 1;

  // Determine the marker pattern for this list
  const baseIndent = lines[i].match(/^(\s*)/)?.[1].length ?? 0;

  while (i < end) {
    const line = lines[i];
    const itemMatch = ordered
      ? line.match(/^(\s*)(\d+)[.)]\s(.*)$/)
      : line.match(/^(\s*)([-*+])\s(.*)$/);

    if (!itemMatch) {
      break;
    }

    const indent = itemMatch[1].length;
    if (indent > baseIndent + 1) {
      break;
    } // Sub-list item belongs to parent

    const firstLine = itemMatch[3];
    i++;

    // Collect continuation lines (indented more than marker)
    const itemLines: string[] = [firstLine];
    const contIndent = indent + (ordered ? itemMatch[2].length + 2 : 2);

    while (i < end) {
      const nextLine = lines[i];
      if (nextLine.trim() === "") {
        // Check if next non-blank line continues this item
        if (i + 1 < end && lines[i + 1].startsWith(" ".repeat(contIndent))) {
          itemLines.push("");
          i++;
        } else {
          break;
        }
      } else if (nextLine.startsWith(" ".repeat(contIndent))) {
        itemLines.push(nextLine.slice(contIndent));
        i++;
      } else {
        break;
      }
    }

    // Check for task list checkbox
    let checked: boolean | undefined;
    if (itemLines[0].startsWith("[x] ") || itemLines[0].startsWith("[X] ")) {
      checked = true;
      itemLines[0] = itemLines[0].slice(4);
    } else if (itemLines[0].startsWith("[ ] ")) {
      checked = false;
      itemLines[0] = itemLines[0].slice(4);
    }

    // Parse item content — could contain sub-blocks
    const children = parseMarkdownBlocks(itemLines, 0, itemLines.length);
    // If no blocks were created but we have text, wrap in paragraph
    if (children.length === 0 && firstLine.trim()) {
      children.push({ type: "paragraph", inlines: parseInlines(firstLine) });
    }
    items.push({ type: "listItem", checked, children });
  }

  return {
    block: { type: "list", ordered, start: startNum, items },
    nextIndex: i
  };
}

// =============================================================================
// Table Parser
// =============================================================================

function isTableRow(line: string): boolean {
  return line.includes("|");
}

function isTableSeparator(line: string): boolean {
  return /^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/.test(line);
}

function parseTableCells(line: string): string[] {
  // Remove leading/trailing pipe
  let s = line.trim();
  if (s.startsWith("|")) {
    s = s.slice(1);
  }
  if (s.endsWith("|")) {
    s = s.slice(0, -1);
  }

  // Split by unescaped pipe
  const cells: string[] = [];
  let current = "";
  for (let i = 0; i < s.length; i++) {
    if (s[i] === "\\" && i + 1 < s.length) {
      current += s[i + 1];
      i++;
    } else if (s[i] === "|") {
      cells.push(current.trim());
      current = "";
    } else {
      current += s[i];
    }
  }
  cells.push(current.trim());
  return cells;
}

function parseTableAlignments(line: string): (Alignment | undefined)[] {
  const cells = parseTableCells(line);
  return cells.map(cell => {
    const trimmed = cell.trim();
    const left = trimmed.startsWith(":");
    const right = trimmed.endsWith(":");
    if (left && right) {
      return "center" as Alignment;
    }
    if (right) {
      return "end" as Alignment;
    }
    if (left) {
      return "start" as Alignment;
    }
    return undefined;
  });
}

function parseTable(
  lines: string[],
  start: number,
  end: number
): { block: TableBlock; nextIndex: number } {
  const headerCells = parseTableCells(lines[start]);
  const alignments = parseTableAlignments(lines[start + 1]);
  const rows: InlineNode[][][] = [];

  let i = start + 2;
  while (i < end && isTableRow(lines[i]) && lines[i].trim() !== "") {
    const cells = parseTableCells(lines[i]);
    rows.push(cells.map(c => parseInlines(c)));
    i++;
  }

  return {
    block: {
      type: "table",
      headers: headerCells.map(c => parseInlines(c)),
      alignments,
      rows
    },
    nextIndex: i
  };
}

// =============================================================================
// Inline Parser
// =============================================================================

function parseInlines(text: string): InlineNode[] {
  const nodes: InlineNode[] = [];
  let i = 0;

  while (i < text.length) {
    // Escaped character
    if (text[i] === "\\" && i + 1 < text.length && /[\\`*_{}[\]()#+\-.!|~>]/.test(text[i + 1])) {
      nodes.push({ type: "text", text: text[i + 1] });
      i += 2;
      continue;
    }

    // Line break (two trailing spaces + newline, or backslash + newline)
    if (text[i] === "\n") {
      // Check for hard break (two spaces before \n)
      const lastNode = nodes[nodes.length - 1];
      if (lastNode && lastNode.type === "text" && lastNode.text.endsWith("  ")) {
        lastNode.text = lastNode.text.slice(0, -2);
        nodes.push({ type: "lineBreak" });
      } else if (i > 0 && text[i - 1] === "\\") {
        // Backslash line break
        if (lastNode && lastNode.type === "text") {
          lastNode.text = lastNode.text.slice(0, -1);
        }
        nodes.push({ type: "lineBreak" });
      } else {
        // Soft line break → space
        nodes.push({ type: "text", text: " " });
      }
      i++;
      continue;
    }

    // Inline code (backtick)
    if (text[i] === "`") {
      const result = parseInlineCode(text, i);
      if (result) {
        nodes.push(result.node);
        i = result.end;
        continue;
      }
    }

    // Image ![alt](url "title")
    if (text[i] === "!" && text[i + 1] === "[") {
      const result = parseImage(text, i);
      if (result) {
        nodes.push(result.node);
        i = result.end;
        continue;
      }
    }

    // Link [text](url "title")
    if (text[i] === "[") {
      const result = parseLink(text, i);
      if (result) {
        nodes.push(result.node);
        i = result.end;
        continue;
      }
    }

    // Autolink <url>
    if (text[i] === "<") {
      const result = parseAutolink(text, i);
      if (result) {
        nodes.push(result.node);
        i = result.end;
        continue;
      }
    }

    // Bold/Italic with ** or __
    if ((text[i] === "*" || text[i] === "_") && i + 1 < text.length) {
      const result = parseEmphasis(text, i);
      if (result) {
        nodes.push(result.node);
        i = result.end;
        continue;
      }
    }

    // Strikethrough ~~text~~
    if (text[i] === "~" && text[i + 1] === "~") {
      const result = parseStrikethrough(text, i);
      if (result) {
        nodes.push(result.node);
        i = result.end;
        continue;
      }
    }

    // Plain text — accumulate until next special character
    let textEnd = i;
    while (textEnd < text.length && !isInlineSpecial(text, textEnd)) {
      textEnd++;
    }
    if (textEnd > i) {
      nodes.push({ type: "text", text: text.slice(i, textEnd) });
      i = textEnd;
    } else {
      // Single special char that didn't match any pattern — treat as text
      nodes.push({ type: "text", text: text[i] });
      i++;
    }
  }

  return mergeTextNodes(nodes);
}

function isInlineSpecial(text: string, i: number): boolean {
  const ch = text[i];
  if (ch === "\\" || ch === "`" || ch === "[" || ch === "!" || ch === "<" || ch === "\n") {
    return true;
  }
  if ((ch === "*" || ch === "_") && i + 1 < text.length) {
    return true;
  }
  if (ch === "~" && text[i + 1] === "~") {
    return true;
  }
  return false;
}

function mergeTextNodes(nodes: InlineNode[]): InlineNode[] {
  const result: InlineNode[] = [];
  for (const node of nodes) {
    const prev = result[result.length - 1];
    if (node.type === "text" && prev && prev.type === "text") {
      (prev as { text: string }).text += node.text;
    } else {
      result.push(node);
    }
  }
  return result;
}

function parseInlineCode(text: string, start: number): { node: CodeInline; end: number } | null {
  // Count opening backticks
  let ticks = 0;
  let i = start;
  while (i < text.length && text[i] === "`") {
    ticks++;
    i++;
  }

  // Find matching closing backticks
  const closePattern = "`".repeat(ticks);
  const closeIdx = text.indexOf(closePattern, i);
  if (closeIdx < 0) {
    return null;
  }

  // Verify it's exactly the right number of backticks
  const afterClose = closeIdx + ticks;
  if (afterClose < text.length && text[afterClose] === "`") {
    return null;
  }

  let code = text.slice(i, closeIdx);
  // Strip one leading and one trailing space if both exist (GFM rule)
  if (code.length >= 2 && code.startsWith(" ") && code.endsWith(" ")) {
    code = code.slice(1, -1);
  }

  return { node: { type: "code", text: code }, end: afterClose };
}

function parseImage(text: string, start: number): { node: ImageInline; end: number } | null {
  // ![alt](url "title")
  const altStart = start + 2; // skip "!["
  const altEnd = findClosingBracket(text, altStart - 1);
  if (altEnd < 0) {
    return null;
  }

  const alt = text.slice(altStart, altEnd);

  if (text[altEnd + 1] !== "(") {
    return null;
  }
  const urlResult = parseLinkDest(text, altEnd + 2);
  if (!urlResult) {
    return null;
  }

  return {
    node: { type: "image", url: urlResult.url, alt, title: urlResult.title },
    end: urlResult.end
  };
}

function parseLink(text: string, start: number): { node: LinkInline; end: number } | null {
  const textEnd = findClosingBracket(text, start);
  if (textEnd < 0) {
    return null;
  }

  const linkText = text.slice(start + 1, textEnd);

  if (text[textEnd + 1] !== "(") {
    return null;
  }
  const urlResult = parseLinkDest(text, textEnd + 2);
  if (!urlResult) {
    return null;
  }

  return {
    node: {
      type: "link",
      url: urlResult.url,
      title: urlResult.title,
      children: parseInlines(linkText)
    },
    end: urlResult.end
  };
}

function parseAutolink(text: string, start: number): { node: LinkInline; end: number } | null {
  const closeIdx = text.indexOf(">", start + 1);
  if (closeIdx < 0) {
    return null;
  }

  const content = text.slice(start + 1, closeIdx);
  // Must be a URL (https?://) or email
  if (/^https?:\/\//.test(content) || /^[^@]+@[^@]+\.[^@]+$/.test(content)) {
    const url = content.includes("@") && !content.includes("://") ? `mailto:${content}` : content;
    return {
      node: {
        type: "link",
        url,
        children: [{ type: "text", text: content }]
      },
      end: closeIdx + 1
    };
  }
  return null;
}

function parseLinkDest(
  text: string,
  start: number
): { url: string; title?: string; end: number } | null {
  let i = start;
  // Skip whitespace
  while (i < text.length && text[i] === " ") {
    i++;
  }

  if (i >= text.length) {
    return null;
  }

  // Parse URL (possibly in angle brackets)
  let url: string;
  if (text[i] === "<") {
    const closeAngle = text.indexOf(">", i + 1);
    if (closeAngle < 0) {
      return null;
    }
    url = text.slice(i + 1, closeAngle);
    i = closeAngle + 1;
  } else {
    const urlStart = i;
    let parens = 0;
    while (i < text.length && text[i] !== " " && text[i] !== "\t") {
      const ch = text[i];
      if (ch === "(") {
        parens++;
      } else if (ch === ")") {
        if (parens === 0) {
          break;
        }
        parens--;
      }
      i++;
    }
    url = text.slice(urlStart, i);
  }

  // Skip whitespace
  while (i < text.length && (text[i] === " " || text[i] === "\t")) {
    i++;
  }

  // Parse optional title
  let title: string | undefined;
  if (i < text.length && (text[i] === '"' || text[i] === "'")) {
    const quote = text[i];
    const titleStart = i + 1;
    const titleEnd = text.indexOf(quote, titleStart);
    if (titleEnd >= 0) {
      title = text.slice(titleStart, titleEnd);
      i = titleEnd + 1;
    }
  }

  // Skip whitespace and expect ")"
  while (i < text.length && text[i] === " ") {
    i++;
  }
  if (text[i] !== ")") {
    return null;
  }

  return { url, title, end: i + 1 };
}

function findClosingBracket(text: string, start: number): number {
  // start points to "[", find matching "]"
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    if (text[i] === "\\" && i + 1 < text.length) {
      i++; // skip escaped char
      continue;
    }
    if (text[i] === "[") {
      depth++;
    }
    if (text[i] === "]") {
      depth--;
      if (depth === 0) {
        return i;
      }
    }
  }
  return -1;
}

function parseEmphasis(
  text: string,
  start: number
): { node: BoldInline | ItalicInline; end: number } | null {
  const ch = text[start];
  const double = text[start + 1] === ch;
  const triple = double && start + 2 < text.length && text[start + 2] === ch;

  if (triple) {
    // ***bold italic*** or ___bold italic___
    const closeIdx = findDelimiterClose(text, start + 3, ch.repeat(3));
    if (closeIdx >= 0) {
      const inner = text.slice(start + 3, closeIdx);
      return {
        node: { type: "bold", children: [{ type: "italic", children: parseInlines(inner) }] },
        end: closeIdx + 3
      };
    }
  }

  if (double) {
    // **bold** or __bold__
    const closeIdx = findDelimiterClose(text, start + 2, ch.repeat(2));
    if (closeIdx >= 0) {
      const inner = text.slice(start + 2, closeIdx);
      return {
        node: { type: "bold", children: parseInlines(inner) },
        end: closeIdx + 2
      };
    }
  }

  // *italic* or _italic_
  // For underscore: must not be in the middle of a word
  if (ch === "_" && start > 0 && /\w/.test(text[start - 1])) {
    return null;
  }
  const closeIdx = findDelimiterClose(text, start + 1, ch);
  if (closeIdx >= 0) {
    if (ch === "_" && closeIdx + 1 < text.length && /\w/.test(text[closeIdx + 1])) {
      return null;
    }
    const inner = text.slice(start + 1, closeIdx);
    if (inner.trim() === "") {
      return null;
    }
    return {
      node: { type: "italic", children: parseInlines(inner) },
      end: closeIdx + 1
    };
  }

  return null;
}

function findDelimiterClose(text: string, start: number, delimiter: string): number {
  let i = start;
  while (i < text.length) {
    if (text[i] === "\\" && i + 1 < text.length) {
      i += 2;
      continue;
    }
    if (text.startsWith(delimiter, i)) {
      // Make sure it's not preceded by whitespace (for closing delimiter)
      if (i > start && text[i - 1] !== " ") {
        return i;
      }
    }
    i++;
  }
  return -1;
}

function parseStrikethrough(
  text: string,
  start: number
): { node: StrikethroughInline; end: number } | null {
  const closeIdx = text.indexOf("~~", start + 2);
  if (closeIdx < 0) {
    return null;
  }
  const inner = text.slice(start + 2, closeIdx);
  if (inner.trim() === "") {
    return null;
  }
  return {
    node: { type: "strikethrough", children: parseInlines(inner) },
    end: closeIdx + 2
  };
}

// =============================================================================
// AST to DOCX Conversion
// =============================================================================

type ConvertOpts = Required<Pick<MarkdownImportOptions, "codeFont" | "codeFontSize">> &
  MarkdownImportOptions;

function convertBlocks(blocks: Block[], opts: ConvertOpts, state: ConversionState): BodyContent[] {
  const result: BodyContent[] = [];
  for (const block of blocks) {
    const converted = convertBlock(block, opts, 0, state);
    result.push(...converted);
  }
  return result;
}

function convertBlock(
  block: Block,
  opts: ConvertOpts,
  listLevel: number,
  state: ConversionState
): BodyContent[] {
  switch (block.type) {
    case "heading":
      return [convertHeading(block)];
    case "paragraph":
      return [convertParagraph(block.inlines, opts)];
    case "blockquote":
      return convertBlockquote(block, opts, state);
    case "fencedCode":
      return [convertFencedCode(block, opts)];
    case "thematicBreak":
      return [convertThematicBreak()];
    case "list":
      return convertList(block, opts, listLevel, state);
    case "table":
      return [convertTable(block, opts)];
    case "html":
      // Pass through as plain text paragraph
      return [
        makeParagraph([makeRun(block.content, { font: opts.codeFont, size: opts.codeFontSize })])
      ];
    default:
      return [];
  }
}

function convertHeading(block: HeadingBlock): Paragraph {
  const children = inlinesToRuns(block.inlines, {});
  return {
    type: "paragraph",
    properties: {
      style: `Heading${block.level}`,
      outlineLevel: (block.level - 1) as 0 | 1 | 2 | 3 | 4 | 5
    },
    children
  };
}

function convertParagraph(
  inlines: InlineNode[],
  opts: ConvertOpts,
  props?: ParagraphProperties
): Paragraph {
  const children = inlinesToRuns(inlines, opts);
  return {
    type: "paragraph",
    properties: props,
    children
  };
}

function convertBlockquote(
  block: BlockquoteBlock,
  opts: ConvertOpts,
  state: ConversionState
): BodyContent[] {
  // Convert blockquote children with "Quote" style and left indent
  const result: BodyContent[] = [];
  for (const child of block.children) {
    const converted = convertBlock(child, opts, 0, state);
    for (const item of converted) {
      if (item.type === "paragraph") {
        result.push({
          ...item,
          properties: {
            ...item.properties,
            style: "Quote",
            indent: { left: 720 }, // 0.5 inch
            borders: {
              left: { style: "single", size: 18, color: "CCCCCC", space: 12 }
            }
          }
        });
      } else {
        result.push(item);
      }
    }
  }
  return result;
}

function convertFencedCode(block: FencedCodeBlock, opts: ConvertOpts): Paragraph {
  const runs: ParagraphChild[] = [];
  const lines = block.code.split("\n");
  for (let i = 0; i < lines.length; i++) {
    runs.push(makeRun(lines[i], { font: opts.codeFont, size: opts.codeFontSize }));
    if (i < lines.length - 1) {
      runs.push(makeRun("", undefined, [{ type: "break" }]));
    }
  }
  return {
    type: "paragraph",
    properties: {
      style: "CodeBlock",
      shading: { fill: "F5F5F5" },
      borders: {
        top: { style: "single", size: 4, color: "E0E0E0", space: 4 },
        bottom: { style: "single", size: 4, color: "E0E0E0", space: 4 },
        left: { style: "single", size: 4, color: "E0E0E0", space: 4 },
        right: { style: "single", size: 4, color: "E0E0E0", space: 4 }
      },
      spacing: { before: 120, after: 120 }
    },
    children: runs
  };
}

function convertThematicBreak(): Paragraph {
  return {
    type: "paragraph",
    properties: {
      thematicBreak: true,
      borders: {
        bottom: { style: "single", size: 12, color: "AAAAAA", space: 1 }
      },
      spacing: { before: 240, after: 240 }
    },
    children: []
  };
}

function convertList(
  block: ListBlock,
  opts: ConvertOpts,
  parentLevel: number,
  state: ConversionState
): BodyContent[] {
  const numId = getOrCreateNumbering(block.ordered, state);
  const result: BodyContent[] = [];

  for (const item of block.items) {
    let firstBlock = true;
    for (const child of item.children) {
      if (child.type === "list") {
        // Nested list — increase level
        const nested = convertList(child, opts, parentLevel + 1, state);
        result.push(...nested);
      } else if (child.type === "paragraph") {
        if (firstBlock) {
          // First paragraph gets list numbering
          const props: ParagraphProperties = {
            numbering: { numId, level: parentLevel }
          };
          const para = convertParagraph(child.inlines, opts, props);

          // Handle task list checkbox prefix
          if (item.checked !== undefined) {
            const checkbox = item.checked ? "☑ " : "☐ ";
            const existingChildren = [...para.children];
            const checkRun = makeRun(checkbox);
            result.push({
              ...para,
              children: [checkRun, ...existingChildren]
            });
          } else {
            result.push(para);
          }
          firstBlock = false;
        } else {
          // Continuation paragraphs — indented but no numbering
          const props: ParagraphProperties = {
            indent: { left: 720 * (parentLevel + 1) }
          };
          result.push(convertParagraph(child.inlines, opts, props));
        }
      } else {
        const converted = convertBlock(child, opts, parentLevel, state);
        result.push(...converted);
        firstBlock = false;
      }
    }
  }

  return result;
}

function convertTable(block: TableBlock, opts: ConvertOpts): Table {
  const colCount = block.headers.length;

  // Header row
  const headerCells: TableCell[] = block.headers.map((cell, ci) => {
    const cellProps: TableCellProperties = {
      shading: { fill: "F0F0F0" },
      verticalAlign: "center"
    };
    const paraProps: ParagraphProperties = {
      alignment: block.alignments[ci]
    };
    const para = convertParagraph(cell, opts, paraProps);
    // Bold header text
    const boldPara: Paragraph = {
      ...para,
      children: para.children.map(child => {
        if ("content" in child) {
          return { ...child, properties: { ...(child as Run).properties, bold: true } };
        }
        return child;
      })
    };
    return { properties: cellProps, content: [boldPara] };
  });

  // Data rows
  const dataRows: TableRow[] = block.rows.map(rowCells => {
    const cells: TableCell[] = [];
    for (let ci = 0; ci < colCount; ci++) {
      const cellInlines = ci < rowCells.length ? rowCells[ci] : [];
      const paraProps: ParagraphProperties = {
        alignment: block.alignments[ci]
      };
      const para = convertParagraph(cellInlines, opts, paraProps);
      cells.push({ content: [para] });
    }
    return { cells };
  });

  const allRows: TableRow[] = [
    { properties: { tableHeader: true }, cells: headerCells },
    ...dataRows
  ];

  // Table properties with borders
  const borders: Required<TableProperties>["borders"] = {
    top: { style: "single", size: 4, color: "CCCCCC" },
    bottom: { style: "single", size: 4, color: "CCCCCC" },
    left: { style: "single", size: 4, color: "CCCCCC" },
    right: { style: "single", size: 4, color: "CCCCCC" },
    insideH: { style: "single", size: 4, color: "CCCCCC" },
    insideV: { style: "single", size: 4, color: "CCCCCC" }
  };

  const tableWidth: TableWidth = { type: "pct", value: 5000 }; // 100%

  const tableProps: TableProperties = {
    width: tableWidth,
    borders,
    layout: "autofit"
  };

  return { type: "table", properties: tableProps, rows: allRows };
}

// =============================================================================
// Inline to Run Conversion
// =============================================================================

function inlinesToRuns(inlines: InlineNode[], opts: Partial<ConvertOpts>): ParagraphChild[] {
  const result: ParagraphChild[] = [];
  for (const node of inlines) {
    inlineToRuns(node, result, {}, opts);
  }
  return result;
}

function inlineToRuns(
  node: InlineNode,
  output: ParagraphChild[],
  inheritedProps: RunProperties,
  opts: Partial<ConvertOpts>
): void {
  switch (node.type) {
    case "text": {
      const textProps: RunProperties = {
        ...inheritedProps,
        ...(opts.defaultFont && !inheritedProps.font ? { font: opts.defaultFont } : {}),
        ...(opts.defaultFontSize && !inheritedProps.size ? { size: opts.defaultFontSize } : {})
      };
      output.push(makeRun(node.text, textProps));
      break;
    }

    case "bold":
      for (const child of node.children) {
        inlineToRuns(child, output, { ...inheritedProps, bold: true }, opts);
      }
      break;

    case "italic":
      for (const child of node.children) {
        inlineToRuns(child, output, { ...inheritedProps, italic: true }, opts);
      }
      break;

    case "strikethrough":
      for (const child of node.children) {
        inlineToRuns(child, output, { ...inheritedProps, strike: true }, opts);
      }
      break;

    case "code":
      output.push(
        makeRun(node.text, {
          ...inheritedProps,
          font: opts.codeFont ?? "Courier New",
          size: opts.codeFontSize ?? 20,
          shading: { fill: "F0F0F0" }
        })
      );
      break;

    case "link": {
      const linkChildren: Run[] = [];
      for (const child of node.children) {
        const tempOutput: ParagraphChild[] = [];
        inlineToRuns(
          child,
          tempOutput,
          { ...inheritedProps, color: "0563C1", underline: "single" },
          opts
        );
        for (const run of tempOutput) {
          if ("content" in run) {
            linkChildren.push(run as Run);
          }
        }
      }
      const safeUrl = sanitizeUrl(node.url);
      const link: Hyperlink = {
        type: "hyperlink",
        url: safeUrl ?? "",
        tooltip: node.title,
        children: linkChildren
      };
      output.push(link);
      break;
    }

    case "image":
      // Images require async resolution — for sync API we insert a placeholder
      // The resolveImage callback in options would be used in an async variant
      output.push(
        makeRun(`[Image: ${node.alt || node.url}]`, {
          ...inheritedProps,
          italic: true,
          color: "666666"
        })
      );
      break;

    case "lineBreak":
      output.push(makeRun("", undefined, [{ type: "break" }]));
      break;
  }
}

// =============================================================================
// Numbering Helpers
// =============================================================================

function getOrCreateNumbering(ordered: boolean, state: ConversionState): number {
  if (ordered && state.orderedNumId !== undefined) {
    return state.orderedNumId;
  }
  if (!ordered && state.bulletNumId !== undefined) {
    return state.bulletNumId;
  }

  const abstractNumId = state.nextNumId;
  const numId = state.nextNumId;
  state.nextNumId++;

  const levels = ordered
    ? [
        {
          level: 0,
          format: "decimal" as const,
          text: "%1.",
          start: 1,
          indent: { left: 720, hanging: 360 }
        },
        {
          level: 1,
          format: "lowerLetter" as const,
          text: "%2.",
          start: 1,
          indent: { left: 1440, hanging: 360 }
        },
        {
          level: 2,
          format: "lowerRoman" as const,
          text: "%3.",
          start: 1,
          indent: { left: 2160, hanging: 360 }
        },
        {
          level: 3,
          format: "decimal" as const,
          text: "%4.",
          start: 1,
          indent: { left: 2880, hanging: 360 }
        },
        {
          level: 4,
          format: "lowerLetter" as const,
          text: "%5.",
          start: 1,
          indent: { left: 3600, hanging: 360 }
        },
        {
          level: 5,
          format: "lowerRoman" as const,
          text: "%6.",
          start: 1,
          indent: { left: 4320, hanging: 360 }
        }
      ]
    : [
        {
          level: 0,
          format: "bullet" as const,
          text: "•",
          start: 1,
          indent: { left: 720, hanging: 360 }
        },
        {
          level: 1,
          format: "bullet" as const,
          text: "◦",
          start: 1,
          indent: { left: 1440, hanging: 360 }
        },
        {
          level: 2,
          format: "bullet" as const,
          text: "▪",
          start: 1,
          indent: { left: 2160, hanging: 360 }
        },
        {
          level: 3,
          format: "bullet" as const,
          text: "•",
          start: 1,
          indent: { left: 2880, hanging: 360 }
        },
        {
          level: 4,
          format: "bullet" as const,
          text: "◦",
          start: 1,
          indent: { left: 3600, hanging: 360 }
        },
        {
          level: 5,
          format: "bullet" as const,
          text: "▪",
          start: 1,
          indent: { left: 4320, hanging: 360 }
        }
      ];

  const tabSuffix: LevelSuffix = "tab";
  const abstractNumbering: AbstractNumbering = {
    abstractNumId,
    levels: levels.map(l => ({
      level: l.level,
      format: l.format,
      text: l.text,
      start: l.start,
      paragraphProperties: { indent: l.indent },
      suffix: tabSuffix
    }))
  };

  const numberingInstance: NumberingInstance = {
    numId,
    abstractNumId
  };

  state.abstractNumberings.push(abstractNumbering);
  state.numberingInstances.push(numberingInstance);

  if (ordered) {
    state.orderedNumId = numId;
  } else {
    state.bulletNumId = numId;
  }

  return numId;
}

// =============================================================================
// Run Construction Helpers
// =============================================================================

function makeRun(
  text: string,
  properties?: RunProperties,
  extraContent?: Array<{ type: string }>
): Run {
  const content: Array<{ type: "text"; text: string } | { type: "break"; breakType?: string }> = [];

  if (extraContent) {
    for (const item of extraContent) {
      content.push(item as { type: "break" });
    }
  }

  if (text) {
    content.push({ type: "text", text });
  }

  const run: Run = { content: content as Run["content"] };
  if (properties && Object.keys(properties).length > 0) {
    return { ...run, properties };
  }
  return run;
}

function makeParagraph(children: ParagraphChild[], properties?: ParagraphProperties): Paragraph {
  return { type: "paragraph", properties, children };
}

// =============================================================================
// Default Styles
// =============================================================================

function defaultMarkdownStyles() {
  return [
    {
      type: "paragraph" as const,
      styleId: "Heading1",
      name: "heading 1",
      basedOn: "Normal",
      next: "Normal",
      paragraphProperties: { spacing: { before: 480, after: 120 }, outlineLevel: 0 },
      runProperties: { size: 48, bold: true }
    },
    {
      type: "paragraph" as const,
      styleId: "Heading2",
      name: "heading 2",
      basedOn: "Normal",
      next: "Normal",
      paragraphProperties: { spacing: { before: 360, after: 120 }, outlineLevel: 1 },
      runProperties: { size: 36, bold: true }
    },
    {
      type: "paragraph" as const,
      styleId: "Heading3",
      name: "heading 3",
      basedOn: "Normal",
      next: "Normal",
      paragraphProperties: { spacing: { before: 240, after: 80 }, outlineLevel: 2 },
      runProperties: { size: 28, bold: true }
    },
    {
      type: "paragraph" as const,
      styleId: "Heading4",
      name: "heading 4",
      basedOn: "Normal",
      next: "Normal",
      paragraphProperties: { spacing: { before: 240, after: 80 }, outlineLevel: 3 },
      runProperties: { size: 24, bold: true }
    },
    {
      type: "paragraph" as const,
      styleId: "Heading5",
      name: "heading 5",
      basedOn: "Normal",
      next: "Normal",
      paragraphProperties: { spacing: { before: 200, after: 60 }, outlineLevel: 4 },
      runProperties: { size: 22, bold: true }
    },
    {
      type: "paragraph" as const,
      styleId: "Heading6",
      name: "heading 6",
      basedOn: "Normal",
      next: "Normal",
      paragraphProperties: { spacing: { before: 200, after: 60 }, outlineLevel: 5 },
      runProperties: { size: 20, bold: true, italic: true }
    },
    {
      type: "paragraph" as const,
      styleId: "Quote",
      name: "Quote",
      basedOn: "Normal",
      paragraphProperties: { indent: { left: 720 } },
      runProperties: { italic: true, color: "555555" }
    },
    {
      type: "paragraph" as const,
      styleId: "CodeBlock",
      name: "Code Block",
      basedOn: "Normal",
      runProperties: { font: "Courier New", size: 20 }
    }
  ];
}
