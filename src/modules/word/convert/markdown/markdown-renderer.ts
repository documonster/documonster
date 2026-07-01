/**
 * DOCX Module - Markdown Renderer
 *
 * Converts a DocxDocument model to a GFM-compatible Markdown string.
 * Supports headings, inline formatting, tables, lists, images, hyperlinks,
 * footnotes, code spans, horizontal rules, and more.
 */

import { extractMathText, isRun } from "@word/core/text-utils";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphChild,
  Run,
  RunContent,
  Table,
  TextBox,
  Hyperlink,
  InsertedRun,
  MovedToRun,
  TextContent,
  MathBlock,
  StructuredDocumentTag
} from "@word/types";

/** Options for Markdown rendering. */
export interface MarkdownRenderOptions {
  /** Include images as ![alt](filename). Default: true. */
  readonly includeImages?: boolean;
  /** Include footnotes as [^N]. Default: true. */
  readonly includeNotes?: boolean;
  /** Heading style: "atx" (# style) or "setext" (underline style). Default: "atx". */
  readonly headingStyle?: "atx" | "setext";
}

/**
 * Convert a DocxDocument to a GFM-compatible Markdown string.
 *
 * @param doc - The document model to convert.
 * @param options - Optional rendering settings.
 * @returns Markdown string.
 */
export function renderToMarkdown(doc: DocxDocument, options?: MarkdownRenderOptions): string {
  const opts: Required<MarkdownRenderOptions> = {
    includeImages: options?.includeImages ?? true,
    includeNotes: options?.includeNotes ?? true,
    headingStyle: options?.headingStyle ?? "atx"
  };

  const state: MdRenderState = {
    doc,
    options: opts,
    lines: [],
    footnotes: [],
    footnoteCounter: 0
  };

  for (const item of doc.body) {
    renderBlock(state, item);
  }

  // Append footnotes at the end
  if (opts.includeNotes && state.footnotes.length > 0) {
    state.lines.push("");
    for (const fn of state.footnotes) {
      state.lines.push(fn);
    }
  }

  return state.lines.join("\n").trim() + "\n";
}

// =============================================================================
// Internal state
// =============================================================================

interface MdRenderState {
  readonly doc: DocxDocument;
  readonly options: Required<MarkdownRenderOptions>;
  readonly lines: string[];
  readonly footnotes: string[];
  footnoteCounter: number;
}

// =============================================================================
// Block rendering
// =============================================================================

function renderBlock(state: MdRenderState, item: BodyContent): void {
  // GFM requires a blank line between block-level transitions (e.g. a
  // list followed by a table, or a paragraph followed by a code block).
  // List items deliberately do NOT push a trailing blank line so they
  // stack tightly; we instead inject one here whenever a non-list
  // block follows a list — and also between two adjacent lists of
  // different ordering (- vs 1.) since GFM otherwise merges them.
  const prev = state.lines.length > 0 ? state.lines[state.lines.length - 1] : "";
  const prevIsList = /^(\s*)([-*+]|\d+[.)])\s/.test(prev);
  const currentIsList = item.type === "paragraph" && isListItemParagraph(state.doc, item);
  if (prev !== "" && prevIsList) {
    if (!currentIsList) {
      state.lines.push("");
    } else if (item.type === "paragraph") {
      // Same-list-type tightness: if the previous list marker matches
      // (both bullet OR both ordered), keep them tight. Otherwise emit
      // a blank line so GFM doesn't merge them into one mixed list.
      const prevIsBullet = /^(\s*)[-*+]\s/.test(prev);
      const prevIsOrdered = /^(\s*)\d+[.)]\s/.test(prev);
      const numRef = item.properties?.numbering;
      if (numRef) {
        const format = getListFormat(state.doc, numRef.numId, numRef.level);
        const currentIsBullet = format === "bullet";
        if ((prevIsBullet && !currentIsBullet) || (prevIsOrdered && currentIsBullet)) {
          state.lines.push("");
        }
      }
    }
  }

  switch (item.type) {
    case "paragraph":
      renderParagraph(state, item);
      break;
    case "table":
      renderTable(state, item);
      break;
    case "textBox":
      renderTextBox(state, item);
      break;
    case "math":
      renderMathBlock(state, item);
      break;
    case "sdt":
      renderSdt(state, item);
      break;
    case "tableOfContents":
      if (item.cachedParagraphs) {
        for (const p of item.cachedParagraphs) {
          renderBlock(state, p);
        }
      }
      break;
    case "floatingImage":
      if (state.options.includeImages) {
        const alt = item.altText || "image";
        state.lines.push(`![${alt}](${item.rId})`);
        state.lines.push("");
      }
      break;
    default:
      break;
  }
}

function renderParagraph(state: MdRenderState, para: Paragraph): void {
  const props = para.properties;

  // Detect page break in children
  if (hasPageBreak(para)) {
    state.lines.push("---");
    state.lines.push("");
  }

  // Check if this is a thematic break (only bottom border, no content)
  if (isThematicBreak(para)) {
    state.lines.push("---");
    state.lines.push("");
    return;
  }

  // Determine heading level
  const headingLevel = getHeadingLevel(props?.style, props?.outlineLevel);

  // Check for blockquote style
  if (isBlockquoteStyle(props?.style)) {
    const text = renderInlineChildren(state, para.children);
    if (text.trim()) {
      state.lines.push("> " + text.trim());
    } else {
      state.lines.push(">");
    }
    state.lines.push("");
    return;
  }

  // Check for code block style
  if (isCodeBlockStyle(props?.style) || isEntireParagraphMonospace(para)) {
    const text = renderPlainInlineChildren(para.children);
    state.lines.push("```");
    state.lines.push(text);
    state.lines.push("```");
    state.lines.push("");
    return;
  }

  // Check for list
  const numRef = props?.numbering;

  const text = renderInlineChildren(state, para.children);

  // Skip empty non-heading paragraphs
  if (!text.trim() && headingLevel === 0 && !numRef) {
    // Emit blank line as paragraph separator
    if (state.lines.length > 0 && state.lines[state.lines.length - 1] !== "") {
      state.lines.push("");
    }
    return;
  }

  if (headingLevel > 0) {
    if (state.options.headingStyle === "setext" && headingLevel <= 2) {
      state.lines.push(text.trim());
      state.lines.push(headingLevel === 1 ? "===" : "---");
    } else {
      state.lines.push("#".repeat(headingLevel) + " " + text.trim());
    }
    state.lines.push("");
    return;
  }

  if (numRef) {
    const indent = "  ".repeat(numRef.level);
    const format = getListFormat(state.doc, numRef.numId, numRef.level);
    const bullet = format === "bullet" ? "-" : "1.";
    state.lines.push(`${indent}${bullet} ${text.trim()}`);
    return;
  }

  state.lines.push(text);
  state.lines.push("");
}

function renderTable(state: MdRenderState, table: Table): void {
  if (table.rows.length === 0) {
    return;
  }

  // Build cell text grid
  const grid: string[][] = [];
  for (const row of table.rows) {
    const rowTexts: string[] = [];
    for (const cell of row.cells) {
      const cellParts: string[] = [];
      for (const block of cell.content) {
        if (block.type === "paragraph") {
          cellParts.push(renderInlineChildren(state, block.children).trim());
        }
      }
      // Escape pipe characters to prevent table structure corruption.
      // Backslashes must be escaped *first*: replacing `|` first leaves
      // a literal `\|` in the source untouched, but a subsequent
      // `\` → `\\` pass would then double-escape it into `\\|`,
      // breaking GFM tables. CodeQL flags the single-pass form as
      // "Incomplete string escaping or encoding".
      rowTexts.push(cellParts.join(" ").replace(/\\/g, "\\\\").replace(/\|/g, "\\|"));
    }
    grid.push(rowTexts);
  }

  if (grid.length === 0) {
    return;
  }

  // Determine column count and widths
  const colCount = Math.max(...grid.map(r => r.length));
  const colWidths: number[] = new Array(colCount).fill(3);
  for (const row of grid) {
    for (let j = 0; j < row.length; j++) {
      colWidths[j] = Math.max(colWidths[j], row[j].length);
    }
  }

  const formatRow = (row: string[]): string => {
    const cells: string[] = [];
    for (let j = 0; j < colCount; j++) {
      cells.push((row[j] ?? "").padEnd(colWidths[j]));
    }
    return "| " + cells.join(" | ") + " |";
  };

  // Header row
  state.lines.push(formatRow(grid[0]));

  // Separator with alignment markers based on header cell paragraph alignment
  const sep: string[] = [];
  for (let j = 0; j < colCount; j++) {
    let alignment: string | undefined;
    if (table.rows.length > 0 && table.rows[0].cells[j]) {
      const cell = table.rows[0].cells[j];
      if (cell.content.length > 0 && cell.content[0].type === "paragraph") {
        alignment = cell.content[0].properties?.alignment;
      }
    }
    const w = colWidths[j];
    if (alignment === "center") {
      sep.push(":" + "-".repeat(Math.max(w - 2, 1)) + ":");
    } else if (alignment === "right") {
      sep.push("-".repeat(Math.max(w - 1, 1)) + ":");
    } else if (alignment === "left") {
      sep.push(":" + "-".repeat(Math.max(w - 1, 1)));
    } else {
      sep.push("-".repeat(w));
    }
  }
  state.lines.push("| " + sep.join(" | ") + " |");

  // Data rows
  for (let i = 1; i < grid.length; i++) {
    state.lines.push(formatRow(grid[i]));
  }
  state.lines.push("");
}

function renderTextBox(state: MdRenderState, textBox: TextBox): void {
  for (const p of textBox.content) {
    const text = renderInlineChildren(state, p.children);
    if (text.trim()) {
      state.lines.push("> " + text.trim());
    }
  }
  state.lines.push("");
}

function renderMathBlock(state: MdRenderState, block: MathBlock): void {
  const text = extractMathText(block.content);
  if (text.trim()) {
    state.lines.push(text);
    state.lines.push("");
  }
}

function renderSdt(state: MdRenderState, sdt: StructuredDocumentTag): void {
  for (const child of sdt.content) {
    if ("type" in child) {
      if (child.type === "paragraph" || child.type === "table") {
        renderBlock(state, child as BodyContent);
      }
    }
  }
}

// =============================================================================
// Inline rendering
// =============================================================================

function renderInlineChildren(state: MdRenderState, children: readonly ParagraphChild[]): string {
  let result = "";
  for (const child of children) {
    if ("type" in child) {
      switch (child.type) {
        case "hyperlink":
          result += renderHyperlink(state, child as Hyperlink);
          break;
        case "bookmarkStart":
        case "bookmarkEnd":
        case "commentRangeStart":
        case "commentRangeEnd":
        case "commentReference":
          break;
        case "insertedRun":
          result += renderRun(state, (child as InsertedRun).run);
          break;
        case "deletedRun":
          // Skip deleted content in markdown
          break;
        case "movedFromRun":
          break;
        case "movedToRun":
          result += renderRun(state, (child as MovedToRun).run);
          break;
        default:
          break;
      }
    } else if (isRun(child)) {
      result += renderRun(state, child);
    }
  }
  return result;
}

function renderHyperlink(state: MdRenderState, link: Hyperlink): string {
  const text = renderInlineChildren(state, link.children);
  const url = link.url ?? (link.anchor ? `#${link.anchor}` : "");
  if (url) {
    return `[${text}](${url})`;
  }
  return text;
}

function renderRun(state: MdRenderState, run: Run): string {
  let text = "";
  for (const content of run.content) {
    text += renderRunContent(state, content);
  }

  if (!text) {
    return "";
  }

  const props = run.properties;
  if (!props) {
    return text;
  }

  // Check for monospace font → inline code
  if (isMonospaceFont(props.font)) {
    return "`" + text + "`";
  }

  // Apply formatting cumulatively (supports combinations like bold+strike)
  let result = text;
  if (props.strike) {
    result = `~~${result}~~`;
  }
  if (props.bold && props.italic) {
    result = `***${result}***`;
  } else if (props.bold) {
    result = `**${result}**`;
  } else if (props.italic) {
    result = `*${result}*`;
  }

  return result;
}

function renderRunContent(state: MdRenderState, content: RunContent): string {
  switch (content.type) {
    case "text":
      return content.text;
    case "break":
      if (content.breakType === "page") {
        // Page breaks are emitted at the paragraph level (see renderParagraph
        // -> hasPageBreak). Skipping here avoids producing two thematic breaks
        // for the same page break.
        return "";
      }
      return "  \n";
    case "tab":
      return " ";
    case "ptab":
      return " ";
    case "carriageReturn":
      return "  \n";
    case "noBreakHyphen":
      return "\u2011";
    case "softHyphen":
      return "";
    case "symbol":
      try {
        const code = parseInt(content.char, 16);
        return String.fromCodePoint(code);
      } catch {
        return content.char;
      }
    case "footnoteRef":
      if (state.options.includeNotes) {
        state.footnoteCounter++;
        const noteId = content.id;
        const noteContent = getFootnoteText(state, noteId);
        state.footnotes.push(`[^${state.footnoteCounter}]: ${noteContent}`);
        return `[^${state.footnoteCounter}]`;
      }
      return "";
    case "endnoteRef":
      if (state.options.includeNotes) {
        state.footnoteCounter++;
        const noteContent = getEndnoteText(state, content.id);
        state.footnotes.push(`[^${state.footnoteCounter}]: ${noteContent}`);
        return `[^${state.footnoteCounter}]`;
      }
      return "";
    case "image":
      if (state.options.includeImages) {
        const alt = content.altText ?? "image";
        const imgDef = state.doc.images?.find(img => img.rId === content.rId);
        const filename = imgDef?.fileName ?? content.rId;
        return `![${alt}](${filename})`;
      }
      return "";
    case "field":
      return content.cachedValue ?? "";
    case "ruby":
      // Output base text only
      return content.baseText.map(r => renderRun(state, r)).join("");
    case "lastRenderedPageBreak":
    case "annotationReference":
      return "";
  }
  return "";
}

// =============================================================================
// Helpers
// =============================================================================

function getHeadingLevel(style: string | undefined, outlineLevel: number | undefined): number {
  if (style) {
    const styleId = style.toLowerCase();
    if (styleId === "heading1" || styleId === "heading 1" || styleId === "title") {
      return 1;
    }
    if (styleId === "heading2" || styleId === "heading 2") {
      return 2;
    }
    if (styleId === "heading3" || styleId === "heading 3") {
      return 3;
    }
    if (styleId === "heading4" || styleId === "heading 4") {
      return 4;
    }
    if (styleId === "heading5" || styleId === "heading 5") {
      return 5;
    }
    if (styleId === "heading6" || styleId === "heading 6") {
      return 6;
    }
    // Generic heading pattern
    const match = /^heading\s*(\d)$/i.exec(styleId);
    if (match) {
      return Math.min(parseInt(match[1], 10), 6);
    }
  }
  if (outlineLevel !== undefined && outlineLevel >= 0 && outlineLevel < 6) {
    return outlineLevel + 1;
  }
  return 0;
}

function getListFormat(doc: DocxDocument, numId: number, level: number): string {
  const instance = doc.numberingInstances?.find(n => n.numId === numId);
  if (!instance) {
    return "bullet";
  }
  const abstractNum = doc.abstractNumberings?.find(a => a.abstractNumId === instance.abstractNumId);
  if (!abstractNum) {
    return "bullet";
  }
  const levelDef = abstractNum.levels.find(l => l.level === level);
  if (!levelDef) {
    return "bullet";
  }
  return levelDef.format;
}

function isMonospaceFont(font: unknown): boolean {
  if (!font) {
    return false;
  }
  if (typeof font === "string") {
    return isMonospaceFontName(font);
  }
  // `!font` above already discarded `null`; `font !== null` here was
  // therefore always true and CodeQL flagged it as a comparison
  // between inconvertible types.
  if (typeof font === "object") {
    const f = font as Record<string, unknown>;
    return (
      isMonospaceFontName(f.ascii as string | undefined) ||
      isMonospaceFontName(f.hAnsi as string | undefined)
    );
  }
  return false;
}

function isMonospaceFontName(name: string | undefined | null): boolean {
  if (!name) {
    return false;
  }
  const lower = name.toLowerCase();
  return (
    lower === "courier new" ||
    lower === "consolas" ||
    lower === "menlo" ||
    lower === "monaco" ||
    lower === "source code pro" ||
    lower === "fira code" ||
    lower === "jetbrains mono"
  );
}

/** Check if a paragraph style is a code block style. */
function isCodeBlockStyle(style: string | undefined): boolean {
  if (!style) {
    return false;
  }
  const lower = style.toLowerCase();
  return lower === "code" || lower === "codeblock" || lower === "code block";
}

/** Check if a paragraph style indicates a blockquote. */
function isBlockquoteStyle(style: string | undefined): boolean {
  if (!style) {
    return false;
  }
  const lower = style.toLowerCase();
  return lower.includes("quote") || lower.includes("blockquote");
}

/**
 * Whether a body-level Paragraph would render as a Markdown list item.
 * Used by renderBlock to decide whether to inject a blank-line separator
 * between adjacent block types — list items stack tightly, but a list
 * must be followed by a blank line before any other block-level element.
 */
function isListItemParagraph(doc: DocxDocument, para: Paragraph): boolean {
  void doc;
  return para.properties?.numbering !== undefined;
}

/** Check if the entire paragraph uses a monospace font (all runs). */
function isEntireParagraphMonospace(para: Paragraph): boolean {
  const runs: Run[] = [];
  for (const child of para.children) {
    if (isRun(child)) {
      runs.push(child);
    }
  }
  if (runs.length === 0) {
    return false;
  }
  for (const run of runs) {
    if (!run.properties?.font || !isMonospaceFont(run.properties.font)) {
      return false;
    }
  }
  return true;
}

/** Render paragraph children as plain text (no markdown formatting). */
function renderPlainInlineChildren(children: readonly ParagraphChild[]): string {
  let result = "";
  for (const child of children) {
    if ("type" in child) {
      if (child.type === "insertedRun") {
        result += renderPlainRun((child as InsertedRun).run);
      } else if (child.type === "movedToRun") {
        result += renderPlainRun((child as MovedToRun).run);
      }
    } else if (isRun(child)) {
      result += renderPlainRun(child);
    }
  }
  return result;
}

/** Render a run as plain text without formatting. */
function renderPlainRun(run: Run): string {
  let text = "";
  for (const content of run.content) {
    if (content.type === "text") {
      text += content.text;
    } else if (content.type === "break") {
      text += "\n";
    } else if (content.type === "tab") {
      text += "\t";
    }
  }
  return text;
}

function isThematicBreak(para: Paragraph): boolean {
  const borders = para.properties?.borders;
  if (!borders) {
    return false;
  }
  // Only bottom border, no text content
  const hasBottom =
    borders.bottom && borders.bottom.style !== "none" && borders.bottom.style !== "nil";
  const hasTop = borders.top && borders.top.style !== "none" && borders.top.style !== "nil";
  const hasLeft = borders.left && borders.left.style !== "none" && borders.left.style !== "nil";
  const hasRight = borders.right && borders.right.style !== "none" && borders.right.style !== "nil";
  if (hasBottom && !hasTop && !hasLeft && !hasRight) {
    // Check if there's no meaningful text
    const text = para.children
      .filter((c): c is Run => "content" in c && !("type" in c))
      .map(r =>
        r.content
          .filter(c => c.type === "text")
          .map(c => (c as TextContent).text)
          .join("")
      )
      .join("");
    return text.trim() === "";
  }
  return false;
}

function hasPageBreak(para: Paragraph): boolean {
  for (const child of para.children) {
    if (isRun(child)) {
      for (const c of child.content) {
        if (c.type === "break" && c.breakType === "page") {
          return true;
        }
      }
    }
  }
  return false;
}

function getFootnoteText(state: MdRenderState, noteId: number): string {
  const note = state.doc.footnotes?.find(n => n.id === noteId);
  if (!note) {
    return "";
  }
  const parts: string[] = [];
  for (const p of note.content) {
    if (p.type === "paragraph") {
      parts.push(renderInlineChildren(state, p.children).trim());
    }
  }
  return parts.join(" ");
}

function getEndnoteText(state: MdRenderState, noteId: number): string {
  const note = state.doc.endnotes?.find(n => n.id === noteId);
  if (!note) {
    return "";
  }
  const parts: string[] = [];
  for (const p of note.content) {
    if (p.type === "paragraph") {
      parts.push(renderInlineChildren(state, p.children).trim());
    }
  }
  return parts.join(" ");
}
