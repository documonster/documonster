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
  StyleDef,
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

  // If TOC was updated, register the TOC1..TOCn paragraph styles so the
  // cached TOC entries don't reference undefined styles. Only do this when
  // the document actually contains a <toc> — otherwise we'd mutate doc
  // shape for no reason and break === comparisons in callers.
  let nextDoc: DocxDocument = doc;
  if (newBody !== doc.body) {
    nextDoc = { ...doc, body: newBody };
  }
  if (opts.updateToc && doc.body.some(item => item.type === "tableOfContents")) {
    const headings = collectHeadings(doc, layout);
    nextDoc = ensureTocStyles(nextDoc, headings);
  }
  return nextDoc;
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
  // Skip the entire pipeline (including TOC style registration) when the
  // document has no <toc> block at all — callers rely on `updateTableOfContents`
  // returning the same reference in that case so they can detect "nothing
  // changed" with a `===` check.
  const hasToc = doc.body.some(item => item.type === "tableOfContents");
  if (!hasToc) {
    return doc;
  }

  const layout = layoutDocument(doc, options?.layoutOptions);
  const headings = collectHeadings(doc, layout);

  const newBody = doc.body.map(item => {
    if (item.type === "tableOfContents") {
      return updateTocContent(item, headings);
    }
    return item;
  });

  const changed = newBody.some((item, i) => item !== doc.body[i]);
  const stage1: DocxDocument = changed ? { ...doc, body: newBody } : doc;
  return ensureTocStyles(stage1, headings);
}

/**
 * Register TOC1..TOCn paragraph styles for every heading level present in
 * the document. The cached TOC paragraphs reference these styles by id, so
 * if they aren't defined Word logs a "missing referenced style" warning on
 * every TOC entry.
 */
function ensureTocStyles(doc: DocxDocument, headings: HeadingEntry[]): DocxDocument {
  if (headings.length === 0) {
    return doc;
  }
  const usedTocLevels = new Set<number>();
  for (const h of headings) {
    usedTocLevels.add(Math.max(1, Math.min(9, h.level)));
  }
  const existingStyles = doc.styles ?? [];
  const definedIds = new Set(existingStyles.map(s => s.styleId));
  const stylesToAdd: StyleDef[] = [];
  for (const lvl of usedTocLevels) {
    const id = `TOC${lvl}`;
    if (definedIds.has(id)) {
      continue;
    }
    stylesToAdd.push({
      type: "paragraph",
      styleId: id,
      name: `toc ${lvl}`,
      basedOn: "Normal",
      next: "Normal",
      uiPriority: 39,
      unhideWhenUsed: true,
      paragraphProperties: {
        spacing: { after: 100 },
        indent: { left: (lvl - 1) * 220 }
      }
    });
  }
  if (stylesToAdd.length === 0) {
    return doc;
  }
  return { ...doc, styles: [...existingStyles, ...stylesToAdd] };
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
  // Hand-rolled parser used in place of the previous chained regex
  // (`/^"?([^"=<>!]*?)"?\s*(<=|>=|<>|=|<|>)\s*"?([^"]*?)"?\s+"([^"]*)"\s+"([^"]*)"/`).
  // CodeQL flagged that regex as polynomial-redos. The grammar is also
  // permissive in ways the regex captured implicitly: real Word IF
  // fields contain operands such as `MERGEFIELD foo` (with internal
  // whitespace) and operands wrapped in quotes. The scanner below
  // mirrors the regex's accepted shape:
  //
  //   args := SP* leftOperand SP* op SP* rightOperand SP+ "trueText" SP+ "falseText" …
  //
  // where `leftOperand` runs up to the first comparison operator that
  // is not inside a quoted span, and `rightOperand` runs up to the
  // first `"` that begins the trueText literal.

  // 1. Find the comparison operator outside any quoted span.
  const opPos = findIfOperator(args, 0);
  if (!opPos) {
    return null;
  }

  // 2. Left operand: everything before the operator, with surrounding
  //    whitespace and outer quotes stripped.
  const left = stripOuterQuotes(args.slice(0, opPos.start).trim());

  // 3. Right operand: text between the operator and the first quoted
  //    literal, with surrounding whitespace and outer quotes stripped.
  let cursor = opPos.next;
  // Scan to the next `"` that is not the immediate value-quoted operand.
  // We have to be careful: the right operand itself may be quoted, e.g.
  // `1 = "bar" "y" "n"`. To match the previous regex we adopt: skip
  // whitespace, optionally consume one quoted span as the right operand,
  // otherwise consume up to the next whitespace+`"` boundary.
  cursor = skipSpaces(args, cursor);
  let right: string;
  if (args.charCodeAt(cursor) === 0x22 /* '"' */) {
    const close = args.indexOf('"', cursor + 1);
    if (close < 0) {
      return null;
    }
    right = args.slice(cursor + 1, close);
    cursor = close + 1;
  } else {
    // Read until the next `"` (which begins the trueText literal).
    const nextQuote = args.indexOf('"', cursor);
    if (nextQuote < 0) {
      return null;
    }
    right = args.slice(cursor, nextQuote).trim();
    cursor = nextQuote;
  }

  // 4. trueText (required quoted string).
  cursor = skipSpaces(args, cursor);
  const trueRead = readQuotedString(args, cursor);
  if (!trueRead) {
    return null;
  }
  cursor = skipSpaces(args, trueRead.next);

  // 5. falseText (required quoted string).
  const falseRead = readQuotedString(args, cursor);
  if (!falseRead) {
    return null;
  }

  return {
    left,
    operator: opPos.value,
    right,
    trueText: trueRead.value,
    falseText: falseRead.value
  };
}

/**
 * Find the first IF-field comparison operator (`<=`, `>=`, `<>`, `=`,
 * `<`, `>`) starting at `from`, skipping over any quoted (`"…"`) spans
 * so that operators inside operand strings are not mistaken for the
 * top-level comparator.
 *
 * Bare `!` is reported as "no operator" rather than absorbed into the
 * preceding operand: the previous regex excluded `!` from the left
 * operand character class, so `1 != 1 …` was rejected outright. We
 * preserve that rejection here to avoid silently parsing `!=` (not a
 * Word IF-field operator) as `=` with `!` glued to the left operand.
 */
function findIfOperator(
  s: string,
  from: number
): { start: number; next: number; value: string } | null {
  const n = s.length;
  let i = from;
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c === 0x22 /* '"' */) {
      // Skip quoted span.
      const close = s.indexOf('"', i + 1);
      if (close < 0) {
        return null;
      }
      i = close + 1;
      continue;
    }
    if (c === 0x21 /* '!' */) {
      // Reject `!` outside quotes — matches the previous regex behaviour.
      return null;
    }
    if (c === 0x3c /* '<' */) {
      const next = s.charCodeAt(i + 1);
      if (next === 0x3d) {
        return { start: i, next: i + 2, value: "<=" };
      }
      if (next === 0x3e) {
        return { start: i, next: i + 2, value: "<>" };
      }
      return { start: i, next: i + 1, value: "<" };
    }
    if (c === 0x3e /* '>' */) {
      if (s.charCodeAt(i + 1) === 0x3d) {
        return { start: i, next: i + 2, value: ">=" };
      }
      return { start: i, next: i + 1, value: ">" };
    }
    if (c === 0x3d /* '=' */) {
      return { start: i, next: i + 1, value: "=" };
    }
    i++;
  }
  return null;
}

/**
 * Strip a single matched pair of outer quotes (`"…"`) from a trimmed
 * operand. Mirrors the implicit `"?…"?` shape of the original regex.
 */
function stripOuterQuotes(s: string): string {
  if (s.length >= 2 && s.charCodeAt(0) === 0x22 && s.charCodeAt(s.length - 1) === 0x22) {
    return s.slice(1, -1);
  }
  return s;
}

function skipSpaces(s: string, from: number): number {
  const n = s.length;
  let i = from;
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) {
      break;
    }
    i++;
  }
  return i;
}

/** Read a quoted (`"…"`) string starting at `from`, or return null. */
function readQuotedString(s: string, from: number): { value: string; next: number } | null {
  if (s.charCodeAt(from) !== 0x22 /* '"' */) {
    return null;
  }
  const close = s.indexOf('"', from + 1);
  if (close < 0) {
    return null;
  }
  return { value: s.slice(from + 1, close), next: close + 1 };
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

/** Collect all bookmarks with their text and page numbers.
 *
 * Walks the entire document (including tables, SDTs, text boxes, etc.) so
 * bookmarks placed inside nested structures still feed PAGEREF / TOC
 * resolution. Page numbers come from `layout.bookmarkPages` when the
 * layout engine has registered them; otherwise we fall back to the
 * top-level body block's `contentPages` entry.
 */
function collectBookmarkInfo(doc: DocxDocument, layout: LayoutResult): Map<string, BookmarkData> {
  const map = new Map<string, BookmarkData>();
  const { contentPages } = layout;

  // Track the current top-level body index so we can fall back to
  // contentPages[i] when the bookmark wasn't seen by the layout pass.
  let topLevelIndex = -1;

  walkBlocks(doc.body as BodyContent[], {
    enterParagraph(para) {
      const text = extractParagraphText(para);
      for (const child of para.children) {
        if (isBookmarkStart(child)) {
          const fallbackPage =
            topLevelIndex >= 0 && contentPages[topLevelIndex] !== undefined
              ? contentPages[topLevelIndex]!
              : 1;
          const page = layout.bookmarkPages.get(child.name) ?? fallbackPage;
          map.set(child.name, { text, page });
        }
      }
    }
  });

  // The walker doesn't expose top-level indexing directly, so re-walk only
  // body to keep `topLevelIndex` synchronised. The previous loop's results
  // already reflect bookmarks via `layout.bookmarkPages` for everything
  // layout could see; this second pass refines the fallback page for
  // bookmarks layout missed (very rare in practice).
  for (let i = 0; i < doc.body.length; i++) {
    topLevelIndex = i;
    const item = doc.body[i];
    if (item.type !== "paragraph") {
      continue;
    }
    for (const child of item.children) {
      if (isBookmarkStart(child)) {
        const existing = map.get(child.name);
        if (existing && existing.page === 1 && contentPages[i] !== undefined) {
          map.set(child.name, { ...existing, page: contentPages[i]! });
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

/** Collect headings from the document for TOC generation.
 *
 * Walks the entire document — headings inside tables (very common) or text
 * boxes / SDTs / TOC cached content also feed the TOC. Page numbers for
 * top-level headings come from `layout.contentPages`; nested headings use
 * the outer block's page or fall back to 1.
 */
function collectHeadings(doc: DocxDocument, layout: LayoutResult): HeadingEntry[] {
  const headings: HeadingEntry[] = [];
  const { contentPages } = layout;

  // Walk body twice to track the enclosing top-level body index. First pass:
  // record outer index per top-level block, then traverse its subtree.
  for (let i = 0; i < doc.body.length; i++) {
    const item = doc.body[i];
    const fallbackPage = contentPages[i] ?? 1;
    walkBlocks([item] as BodyContent[], {
      enterParagraph(para) {
        const level = getHeadingLevel(para, doc.styles);
        if (level !== null) {
          headings.push({
            text: extractParagraphText(para),
            level,
            page: fallbackPage
          });
        }
      }
    });
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

/** Compute all SEQ field values by traversing the document in order.
 *
 * Walks every paragraph reachable from the body — including ones inside
 * tables / SDTs / text boxes / TOC caches — so SEQ counters in tabular
 * figure / table captions advance correctly. Field index ordering across
 * nested blocks matches the document reading order produced by
 * walkBlocks (depth-first).
 */
function computeSeqValues(doc: DocxDocument): Map<number, number> {
  const counters = new Map<string, number>();
  const values = new Map<number, number>();
  const state = { fieldIndex: 0 };

  walkBlocks(doc.body as BodyContent[], {
    enterParagraph(para) {
      state.fieldIndex = processSeqInParagraph(para, counters, values, state.fieldIndex);
    }
  });

  return values;
}

function processSeqInParagraph(
  para: Paragraph,
  counters: Map<string, number>,
  values: Map<number, number>,
  startIndex: number
): number {
  let fieldIndex = startIndex;

  // Walk every visible Run reachable from the paragraph — including ones
  // nested inside hyperlinks (`<w:hyperlink>`), tracked-insertion
  // wrappers (`<w:ins>`) and moved-to wrappers (`<w:moveTo>`). The
  // previous implementation only saw top-level runs, so common cases
  // like a SEQ field placed inside `Figure 1` (typically wrapped in a
  // hyperlink for cross-references) silently skipped the SEQ counter
  // and produced wrong figure numbering throughout the document.
  forEachVisibleRun(para.children, run => {
    for (const content of run.content) {
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
  });

  return fieldIndex;
}

/**
 * Walk every "visible" Run reachable from a paragraph child list.
 *
 * Visible runs are:
 *   - bare Run children;
 *   - Run children inside a Hyperlink wrapper;
 *   - the inner `run` of an `insertedRun` or `movedToRun` wrapper.
 *
 * `deletedRun` / `movedFromRun` represent pending removals and so are
 * skipped, matching the convention used by `extractParagraphText` and
 * `replaceText`.
 */
function forEachVisibleRun(children: readonly ParagraphChild[], cb: (run: Run) => void): void {
  for (const child of children) {
    if (isRun(child)) {
      cb(child);
      continue;
    }
    if ("type" in child) {
      const t = (child as { type: string }).type;
      if (t === "hyperlink") {
        forEachVisibleRun((child as { children: readonly ParagraphChild[] }).children, cb);
        continue;
      }
      if (t === "insertedRun" || t === "movedToRun") {
        const inner = (child as { run?: Run }).run;
        if (inner) {
          cb(inner);
        }
      }
    }
  }
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
function evaluateFormula(
  args: string,
  bodyIndex: number,
  doc: DocxDocument,
  cellCtx: CellContext | undefined
): string {
  // Extract the number format switch \# "format"
  const formatMatch = /\\#\s*"([^"]+)"/.exec(args);
  const format = formatMatch ? formatMatch[1] : null;

  // Remove format switch from expression
  let expr = args.replace(/\\#\s*"[^"]*"/, "").trim();

  // Handle SUM(ABOVE) and SUM(LEFT)
  expr = expr.replace(/SUM\s*\(\s*ABOVE\s*\)/gi, () => {
    return String(sumAbove(bodyIndex, doc, cellCtx));
  });
  expr = expr.replace(/SUM\s*\(\s*LEFT\s*\)/gi, () => {
    return String(sumLeft(bodyIndex, doc, cellCtx));
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

  // The parser surfaces malformed input via `ctx.error` instead of an
  // exception so the field engine doesn't need a try/catch around what is
  // already a control-flow path. Any token that fails to parse as a number
  // sets the flag and the result is discarded.
  const ctx: ParseCtx = { pos: 0, error: false };
  const result = parseExpression(tokens, ctx);
  if (ctx.error) {
    return null;
  }
  return result;
}

interface ParseCtx {
  pos: number;
  error: boolean;
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
function parseExpression(tokens: string[], ctx: ParseCtx): number {
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

function parseTerm(tokens: string[], ctx: ParseCtx): number {
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

function parseFactor(tokens: string[], ctx: ParseCtx): number {
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
    ctx.error = true;
    return 0;
  }
  return num;
}

/**
 * Sum numeric values from cells above the current cell in a table.
 * `cellCtx` identifies which cell the formula is being evaluated in.
 * Without it the previous heuristic returned the first formula-bearing
 * cell of the first table in the document, producing wrong sums whenever
 * more than one cell carried a formula.
 */
function sumAbove(bodyIndex: number, doc: DocxDocument, cellCtx: CellContext | undefined): number {
  const location = cellCtx ?? findCellInTable(bodyIndex, doc);
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

/** Sum numeric values from cells to the left of the current cell. */
function sumLeft(bodyIndex: number, doc: DocxDocument, cellCtx: CellContext | undefined): number {
  const location = cellCtx ?? findCellInTable(bodyIndex, doc);
  if (!location) {
    return 0;
  }

  const { table, rowIndex, colIndex } = location;
  let sum = 0;

  const row = table.rows[rowIndex];
  if (row) {
    for (let c = 0; c < colIndex; c++) {
      const cell = row.cells[c];
      if (cell) {
        sum += extractCellNumericValue(cell);
      }
    }
  }

  return sum;
}

/** Cell coordinates carried into formula evaluation. */
interface CellContext {
  readonly table: Table;
  readonly rowIndex: number;
  readonly colIndex: number;
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
        opts,
        // No cell context at the body level — formulas here can still
        // call SUM(ABOVE/LEFT) but findCellInTable will be used as a
        // fallback (likely returning null at the body level).
        undefined
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
  const newRows = table.rows.map((row, rowIndex) => {
    let rowChanged = false;
    const newCells = row.cells.map((cell, colIndex) => {
      let cellChanged = false;
      const newContent: (Paragraph | Table)[] = [];
      // Build the cell context once per cell so SUM(ABOVE)/SUM(LEFT)
      // resolves against THIS cell's coordinates, not whichever cell
      // happened to be the first formula-bearing cell in the table.
      const cellCtx: CellContext = { table, rowIndex, colIndex };
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
            opts,
            cellCtx
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
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions },
  cellCtx: CellContext | undefined
): { paragraph: Paragraph; nextFieldIndex: number } {
  let fieldIndex = startFieldIndex;
  let childrenChanged = false;
  const newChildren: ParagraphChild[] = [];

  // Helper that runs `updateRunFields` against a single Run and returns
  // the (possibly new) Run reference.
  const updateRun = (r: Run): Run => {
    const runResult = updateRunFields(
      r,
      bodyIndex,
      layout,
      bookmarkInfo,
      seqValues,
      styleIndex,
      indexEntries,
      doc,
      fieldIndex,
      opts,
      cellCtx
    );
    fieldIndex = runResult.nextFieldIndex;
    return runResult.run;
  };

  for (const child of para.children) {
    if (isRun(child)) {
      const updated = updateRun(child);
      if (updated !== child) {
        childrenChanged = true;
      }
      newChildren.push(updated);
      continue;
    }
    // Hyperlink: descend into its run children so fields like SEQ/REF
    // wrapped in a hyperlink are still evaluated.
    if ("type" in child && (child as { type: string }).type === "hyperlink") {
      const hl = child as { type: "hyperlink"; children: readonly Run[] } & object;
      let hlChanged = false;
      const newRuns = hl.children.map(r => {
        const upd = updateRun(r);
        if (upd !== r) {
          hlChanged = true;
        }
        return upd;
      });
      if (hlChanged) {
        childrenChanged = true;
        newChildren.push({ ...(hl as object), children: newRuns } as unknown as ParagraphChild);
      } else {
        newChildren.push(child);
      }
      continue;
    }
    // Tracked-insert / moved-to wrappers: update the inner run in place.
    if (
      "type" in child &&
      ((child as { type: string }).type === "insertedRun" ||
        (child as { type: string }).type === "movedToRun")
    ) {
      const wrap = child as { type: string; run: Run } & object;
      const updated = updateRun(wrap.run);
      if (updated !== wrap.run) {
        childrenChanged = true;
        newChildren.push({ ...(wrap as object), run: updated } as unknown as ParagraphChild);
      } else {
        newChildren.push(child);
      }
      continue;
    }
    newChildren.push(child);
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
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions },
  cellCtx: CellContext | undefined
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
      opts,
      cellCtx
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
  opts: Required<Omit<FieldUpdateOptions, "layoutOptions">> & { layoutOptions?: LayoutOptions },
  cellCtx: CellContext | undefined
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
      const value = evaluateFormula(args, bodyIndex, doc, cellCtx);
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
