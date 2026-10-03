/**
 * DOCX Module - Markdown Renderer
 *
 * Converts a DocxDocument model to a GFM-compatible Markdown string.
 * Supports headings, inline formatting, tables, lists, images, hyperlinks,
 * footnotes, code spans, horizontal rules, and more.
 */

import { extractMathText, isRun, symbolText } from "@word/core/text-utils";
import { finalViewRun, isHiddenRun } from "@word/query/final-view";
import { createNumberingCounter } from "@word/query/numbering-counter";
import type { NumberingCounter } from "@word/query/numbering-counter";
import { createStyleResolver } from "@word/query/style-resolver";
import type { StyleResolver } from "@word/query/style-resolver";
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
  RunProperties,
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
    footnoteCounter: 0,
    styles: createStyleResolver(doc),
    numbering: createNumberingCounter(doc),
    openLists: [],
    listIndents: []
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
  /** Style resolution bound to `doc` for this render. */
  readonly styles: StyleResolver;
  /** Real number of each list item, per numbering instance and level. */
  readonly numbering: NumberingCounter;
  /**
   * The numbering instance of the Markdown list currently open at each
   * level, so an adjacent item from a different instance is split off into
   * a list of its own instead of being merged by CommonMark.
   */
  openLists: number[];
  /**
   * Content column of the item open at each level. CommonMark nests a list
   * only when it is indented to its parent item's content, which for "10. "
   * is four columns, not two.
   */
  listIndents: number[];
  /** Run properties the paragraph being rendered passes down to its runs. */
  paragraphRunProperties?: RunProperties;
  /**
   * In a heading, the formatting its style already gives every run. The `#`
   * marker conveys it, so emphasis is emitted only where a run differs.
   */
  headingRunProperties?: RunProperties;
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
  const currentIsList = item.type === "paragraph" && state.styles.numbering(item) !== undefined;
  if (!currentIsList) {
    state.openLists = [];
    state.listIndents = [];
  }
  if (prev !== "" && prevIsList) {
    if (!currentIsList) {
      state.lines.push("");
    } else if (item.type === "paragraph") {
      // Same-list-type tightness: if the previous list marker matches
      // (both bullet OR both ordered), keep them tight. Otherwise emit
      // a blank line so GFM doesn't merge them into one mixed list.
      const prevIsBullet = /^(\s*)[-*+]\s/.test(prev);
      const prevIsOrdered = /^(\s*)\d+[.)]\s/.test(prev);
      const numRef = state.styles.numbering(item);
      if (numRef) {
        const format = state.numbering.levelDef(numRef.numId, numRef.level)?.format ?? "bullet";
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
  // Title renders as `#`; levels 7–9 have no Markdown marker and clamp to `######`.
  const heading = state.styles.heading(para);
  const headingLevel = heading ? Math.min(heading.level, 6) : 0;

  // Check for blockquote style
  if (isBlockquoteStyle(props?.style)) {
    const text = renderParagraphText(state, para);
    if (text.trim()) {
      state.lines.push("> " + text.trim());
    } else {
      state.lines.push(">");
    }
    state.lines.push("");
    return;
  }

  // Check for code block style
  if (isCodeBlockStyle(props?.style) || isEntireParagraphMonospace(state, para)) {
    const text = renderPlainInlineChildren(state, para);
    state.lines.push("```");
    state.lines.push(text);
    state.lines.push("```");
    state.lines.push("");
    return;
  }

  // Check for list
  const numRef = state.styles.numbering(para);

  const text = renderParagraphText(state, para, headingLevel > 0);

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
    const level = numRef.level;
    const indent = " ".repeat(level === 0 ? 0 : (state.listIndents[level - 1] ?? 2 * level));
    const levelDef = state.numbering.levelDef(numRef.numId, numRef.level);
    let bullet = "-";
    // Every item advances its instance's count, bullets included, so the
    // levels beneath it restart.
    const advanced = levelDef ? state.numbering.next(numRef.numId, numRef.level) : 0;
    if (levelDef && levelDef.format !== "bullet") {
      // CommonMark takes an ordered list's start from its first marker, so
      // every item carries its real number: an item resuming an instance
      // after an interruption then opens a list that starts where Word's
      // count stands (e.g. "3.").
      bullet = `${advanced}.`;
      const open = state.openLists[numRef.level];
      if (open !== undefined && open !== numRef.numId) {
        // A different instance directly after an open list would be merged
        // into it. An HTML comment is the CommonMark idiom for ending a list.
        if (state.lines[state.lines.length - 1] !== "") {
          state.lines.push("");
        }
        state.lines.push(`${indent}<!-- -->`, "");
      }
    }
    state.openLists = state.openLists.slice(0, level);
    state.openLists[level] = numRef.numId;
    state.listIndents = state.listIndents.slice(0, level);
    state.listIndents[level] = indent.length + bullet.length + 1;
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
          cellParts.push(renderParagraphText(state, block).trim());
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
    const text = renderParagraphText(state, p);
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

/** Render a paragraph's inline content with its style's run properties in scope. */
function renderParagraphText(state: MdRenderState, para: Paragraph, isHeading = false): string {
  const outer = state.paragraphRunProperties;
  const outerHeading = state.headingRunProperties;
  const paragraphRunProperties = state.styles.paragraph(para).runProperties;
  state.paragraphRunProperties = paragraphRunProperties;
  state.headingRunProperties = isHeading ? paragraphRunProperties : undefined;
  const text = renderInlineChildren(state, para.children);
  state.paragraphRunProperties = outer;
  state.headingRunProperties = outerHeading;
  return text;
}

function renderInlineChildren(state: MdRenderState, children: readonly ParagraphChild[]): string {
  let result = "";
  for (const child of children) {
    // Final view: inserted / moved-to text shown, deleted / moved-from hidden.
    const run = finalViewRun(child);
    if (run) {
      result += renderRun(state, run);
    } else if ("type" in child && child.type === "hyperlink") {
      result += renderHyperlink(state, child);
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
  const props = state.styles.run(run, state.paragraphRunProperties).runProperties;
  if (isHiddenRun(props)) {
    return "";
  }
  let text = "";
  for (const content of run.content) {
    text += renderRunContent(state, content);
  }

  if (!text) {
    return "";
  }

  // Check for monospace font → inline code
  if (isMonospaceFont(props.font)) {
    return "`" + text + "`";
  }

  // Apply formatting cumulatively (supports combinations like bold+strike)
  const base = state.headingRunProperties;
  const strike = props.strike === true && base?.strike !== true;
  const bold = props.bold === true && base?.bold !== true;
  const italic = props.italic === true && base?.italic !== true;
  let result = text;
  if (strike) {
    result = `~~${result}~~`;
  }
  if (bold && italic) {
    result = `***${result}***`;
  } else if (bold) {
    result = `**${result}**`;
  } else if (italic) {
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
      return symbolText(content);
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
 * Check if the entire paragraph uses a monospace font (all runs), judged on
 * the runs' resolved properties — the same view `renderRun` uses, so a font
 * inherited from a character or paragraph style counts.
 */
function isEntireParagraphMonospace(state: MdRenderState, para: Paragraph): boolean {
  const paragraphRunProperties = state.styles.paragraph(para).runProperties;
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
    if (!isMonospaceFont(state.styles.run(run, paragraphRunProperties).runProperties.font)) {
      return false;
    }
  }
  return true;
}

/** Render paragraph children as plain text (no markdown formatting). */
function renderPlainInlineChildren(state: MdRenderState, para: Paragraph): string {
  const paragraphRunProperties = state.styles.paragraph(para).runProperties;
  let result = "";
  for (const child of para.children) {
    const run = finalViewRun(child);
    if (run && !isHiddenRun(state.styles.run(run, paragraphRunProperties).runProperties)) {
      result += renderPlainRun(run);
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
      parts.push(renderParagraphText(state, p).trim());
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
      parts.push(renderParagraphText(state, p).trim());
    }
  }
  return parts.join(" ");
}
