/**
 * DOCX Module - Text Utilities & Type Guards
 *
 * Single source of truth for type guards and text extraction helpers
 * used across the word module. Eliminates duplicated implementations
 * of `isRun`, `extractParagraphText`, `extractMathText`, etc.
 */

import type {
  BodyContent,
  Paragraph,
  ParagraphChild,
  Run,
  Table,
  MathContent,
  MathBlock,
  StructuredDocumentTag,
  Hyperlink,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun,
  FloatingImage,
  TextBox
} from "../types";

// =============================================================================
// Type Guards — Body Content
// =============================================================================

/** Check if a body content block is a Paragraph. */
export function isParagraph(content: BodyContent): content is Paragraph {
  return content.type === "paragraph";
}

/** Check if a body content block is a Table. */
export function isTable(content: BodyContent): content is Table {
  return content.type === "table";
}

/** Check if a body content block is a StructuredDocumentTag. */
export function isSdt(content: BodyContent): content is StructuredDocumentTag {
  return content.type === "sdt";
}

/** Check if a body content block is a MathBlock. */
export function isMathBlock(content: BodyContent): content is MathBlock {
  return content.type === "math";
}

/** Check if a body content block is a TextBox. */
export function isTextBox(content: BodyContent): content is TextBox {
  return content.type === "textBox";
}

/** Check if a body content block is a FloatingImage. */
export function isFloatingImage(content: BodyContent): content is FloatingImage {
  return content.type === "floatingImage";
}

// =============================================================================
// Type Guards — Paragraph Children
// =============================================================================

/**
 * Check if a paragraph child is a Run.
 *
 * A Run is the only ParagraphChild without a `type` discriminator field.
 * All other paragraph children (Hyperlink, InsertedRun, DeletedRun, MovedFromRun,
 * MovedToRun, BookmarkStart/End, CommentRangeStart/End) have an explicit `type`
 * string. Therefore, a child is a Run iff it has a `content` array AND no `type`
 * field — this is more robust than checking `content` alone (since some typed
 * children also have `content`-like properties).
 */
export function isRun(child: ParagraphChild): child is Run {
  if ("type" in child) {
    return false;
  }
  return "content" in child && Array.isArray((child as Run).content);
}

/** Check if a paragraph child is a Hyperlink. */
export function isHyperlink(child: ParagraphChild): child is Hyperlink {
  return "type" in child && (child as { type: string }).type === "hyperlink";
}

/** Check if a paragraph child is an InsertedRun. */
export function isInsertedRun(child: ParagraphChild): child is InsertedRun {
  return "type" in child && (child as { type: string }).type === "insertedRun";
}

/** Check if a paragraph child is a DeletedRun. */
export function isDeletedRun(child: ParagraphChild): child is DeletedRun {
  return "type" in child && (child as { type: string }).type === "deletedRun";
}

/** Check if a paragraph child is a MovedFromRun. */
export function isMovedFromRun(child: ParagraphChild): child is MovedFromRun {
  return "type" in child && (child as { type: string }).type === "movedFromRun";
}

/** Check if a paragraph child is a MovedToRun. */
export function isMovedToRun(child: ParagraphChild): child is MovedToRun {
  return "type" in child && (child as { type: string }).type === "movedToRun";
}

// =============================================================================
// Text Extraction — Runs
// =============================================================================

/**
 * Extract plain text from a single run's content array.
 */
export function extractRunText(run: Run): string {
  let text = "";
  for (const c of run.content) {
    if (c.type === "text") {
      text += c.text;
    } else if (c.type === "tab") {
      text += "\t";
    } else if (c.type === "break") {
      text += "\n";
    } else if (c.type === "noBreakHyphen") {
      text += "-";
    } else if (c.type === "softHyphen") {
      text += "\u00AD";
    } else if (c.type === "field" && c.cachedValue) {
      text += c.cachedValue;
    }
  }
  return text;
}

// =============================================================================
// Text Extraction — Paragraphs
// =============================================================================

/**
 * Extract concatenated plain text from a paragraph's children (runs + hyperlinks).
 * This recursively handles hyperlinks and tracked changes.
 */
export function extractParagraphText(para: Paragraph): string {
  let text = "";
  for (const child of para.children) {
    text += extractChildText(child);
  }
  return text;
}

/**
 * Extract text from a single paragraph child (run, hyperlink, or tracked change).
 */
export function extractChildText(child: ParagraphChild): string {
  if (isRun(child)) {
    return extractRunText(child);
  }
  if ("type" in child) {
    const typed = child as { type: string };
    switch (typed.type) {
      case "hyperlink": {
        const hl = child as Hyperlink;
        let text = "";
        for (const c of hl.children) {
          text += extractChildText(c as ParagraphChild);
        }
        return text;
      }
      case "insertedRun":
        return extractRunText((child as InsertedRun).run);
      case "deletedRun":
        return ""; // Deleted content is not visible
      case "movedFromRun":
        return ""; // Moved-from is not visible at source
      case "movedToRun":
        return extractRunText((child as MovedToRun).run);
      default:
        return "";
    }
  }
  return "";
}

// =============================================================================
// Text Extraction — Math
// =============================================================================

/**
 * Extract plain text representation from math content elements.
 */
export function extractMathText(content: readonly MathContent[]): string {
  let text = "";
  for (const item of content) {
    switch (item.type) {
      case "mathRun":
        text += item.text;
        break;
      case "mathFraction":
        text += `(${extractMathText(item.numerator)})/(${extractMathText(item.denominator)})`;
        break;
      case "mathRadical":
        if (item.hideDegree) {
          text += `√(${extractMathText(item.content)})`;
        } else if (item.degree) {
          text += `${extractMathText(item.degree)}√(${extractMathText(item.content)})`;
        } else {
          text += `√(${extractMathText(item.content)})`;
        }
        break;
      case "mathSuperScript":
        text += `${extractMathText(item.base)}^(${extractMathText(item.superScript)})`;
        break;
      case "mathSubScript":
        text += `${extractMathText(item.base)}_(${extractMathText(item.subScript)})`;
        break;
      case "mathSubSuperScript":
        text += `${extractMathText(item.base)}_(${extractMathText(item.subScript)})^(${extractMathText(item.superScript)})`;
        break;
      case "mathPreSubSuperScript":
        text += `_(${extractMathText(item.preSubScript)})^(${extractMathText(item.preSuperScript)})${extractMathText(item.base)}`;
        break;
      case "mathDelimiter":
        text += item.beginChar ?? "(";
        text += item.content.map(c => extractMathText(c)).join(item.separatorChar ?? ",");
        text += item.endChar ?? ")";
        break;
      case "mathNary":
        text += item.char ?? "∑";
        if (item.sub) {
          text += `_(${extractMathText(item.sub)})`;
        }
        if (item.sup) {
          text += `^(${extractMathText(item.sup)})`;
        }
        text += extractMathText(item.content);
        break;
      case "mathFunction":
        text += extractMathText(item.name);
        text += extractMathText(item.content);
        break;
      case "mathLimit":
        text += extractMathText(item.base);
        text += item.limitType === "upper" ? "^" : "_";
        text += `(${extractMathText(item.limit)})`;
        break;
      case "mathMatrix":
        text += "[";
        text += item.rows.map(row => row.map(cell => extractMathText(cell)).join(",")).join(";");
        text += "]";
        break;
      case "mathAccent":
        text += extractMathText(item.content);
        if (item.char) {
          text += item.char;
        }
        break;
      case "mathBar":
        text += extractMathText(item.content);
        break;
      case "mathBox":
        text += extractMathText(item.content);
        break;
      case "mathEquationArray":
        text += item.rows.map(row => extractMathText(row)).join("\n");
        break;
      case "mathGroupChar":
        text += extractMathText(item.base);
        break;
      case "mathPhantom":
        text += extractMathText(item.content);
        break;
      case "mathBorderBox":
        text += extractMathText(item.content);
        break;
    }
  }
  return text;
}

// =============================================================================
// Text Extraction — Body Content (recursive)
// =============================================================================

/**
 * Extract plain text from an array of body content blocks.
 * Paragraphs are separated by newlines, table cells by tabs.
 */
export function extractBodyText(blocks: readonly BodyContent[]): string {
  const lines: string[] = [];
  collectBlockText(blocks, lines);
  return lines.join("\n");
}

/**
 * Recursively collect text lines from body content blocks into an output array.
 */
export function collectBlockText(blocks: readonly BodyContent[], lines: string[]): void {
  for (const block of blocks) {
    if (block.type === "paragraph") {
      lines.push(extractParagraphText(block));
    } else if (block.type === "table") {
      for (const row of block.rows) {
        const cellTexts: string[] = [];
        for (const cell of row.cells) {
          const cellLines: string[] = [];
          collectBlockText(cell.content as readonly BodyContent[], cellLines);
          cellTexts.push(cellLines.join(" "));
        }
        lines.push(cellTexts.join("\t"));
      }
    } else if (block.type === "sdt") {
      const filtered = block.content.filter(
        c => "type" in c && (c.type === "paragraph" || c.type === "table")
      );
      collectBlockText(filtered as readonly BodyContent[], lines);
    } else if (block.type === "textBox" && block.content) {
      collectBlockText(block.content as readonly BodyContent[], lines);
    }
  }
}
