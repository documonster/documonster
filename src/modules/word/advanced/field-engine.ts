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

import { planTableOfContents } from "@word/advanced/toc-update";
import type { TocPlan } from "@word/advanced/toc-update";
import { checkBoolean, checkHandle, checkRecord } from "@word/core/argument-checks";
import { builtInStyle, tocStyle } from "@word/core/builtin-styles";
import { formatFieldResult, tokenizeFieldCode } from "@word/core/field-format";
import { evaluateIfField } from "@word/core/field-if";
import {
  evaluateSplitFields,
  forEachField,
  resolveSplitFields,
  rewriteSplitFields
} from "@word/core/field-view";
import type { Mutable } from "@word/core/internal-utils";
import { collectRangeMarkers, settleRangeMarkers } from "@word/core/range-markers";
import { extractParagraphText, isRun } from "@word/core/text-utils";
import { buildTocInstruction } from "@word/core/toc-instruction";
import { walkBlocks } from "@word/core/walker";
import type { LayoutOptions, LayoutResult } from "@word/layout/layout";
import { layoutDocument } from "@word/layout/layout";
import { checkLayoutOptions, layoutFootprint, paginateDocument } from "@word/layout/layout-full";
import {
  collectSectionProperties,
  displayedPageNumbers,
  formatPageField,
  type PageSectionSlot,
  parsePageField
} from "@word/query/page-field";
import type {
  BodyContent,
  BookmarkStart,
  DocxDocument,
  FieldContent,
  PageNumberFormat,
  Paragraph,
  ParagraphChild,
  Run,
  SectionProperties,
  RunContent,
  StyleDef,
  Table,
  TableCell,
  TableOfContents
} from "@word/types";

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
 * 嵌套字段先于外层字段计算（`IF` 比较的是嵌套字段的新结果）；结果按字段的
 * 通用开关（`\\*`、`\\#`、`\\@`，见 `formatFieldResult`）格式化。
 *
 * 页码（PAGE、NUMPAGES、SECTIONPAGES、PAGEREF、INDEX、目录条目）来自本库的
 * 分页布局（目录按其新条目参与分页）。不会设置 `settings.updateFieldsOnOpen`。
 *
 * Does not modify `doc`; the result may share unchanged nodes with it — see
 * `Document.clone`.
 *
 * @param doc - 文档模型（不会被修改，返回新对象）
 * @param options - 更新选项
 * @returns 更新后的文档模型
 */
function checkFieldUpdateOptions(options: unknown, where: string): void {
  checkRecord(options, "options", where, { optional: true });
  const o = options as FieldUpdateOptions | undefined;
  checkLayoutOptions(o?.layoutOptions, "options.layoutOptions", where);
  for (const key of [
    "updateToc",
    "updatePageFields",
    "updateReferences",
    "updateSequences"
  ] as const) {
    checkBoolean(o?.[key], `options.${key}`, where, { optional: true });
  }
}

export function updateFields(doc: DocxDocument, options?: FieldUpdateOptions): DocxDocument {
  checkHandle(doc, "document", "doc", "Transform.updateFields");
  checkFieldUpdateOptions(options, "Transform.updateFields");
  const opts: ResolvedFieldOptions = {
    layoutOptions: options?.layoutOptions,
    updateToc: options?.updateToc ?? true,
    updatePageFields: options?.updatePageFields ?? true,
    updateReferences: options?.updateReferences ?? true,
    updateSequences: options?.updateSequences ?? true
  };

  // The tables of contents first: they bookmark what they list, which the
  // layout then places.
  const tocInstructions = opts.updateToc ? listTocInstructions(doc) : [];
  const plan = tocInstructions.length > 0 ? planTableOfContents(doc, tocInstructions) : undefined;
  const { doc: source, layout } = plan
    ? settleTableOfContents(plan, tocInstructions, opts.layoutOptions)
    : { doc, layout: paginateDocument(doc, opts.layoutOptions) };

  const env: UpdateEnv = {
    doc: source,
    layout,
    // Build bookmark info map
    bookmarkInfo: opts.updateReferences ? collectBookmarkInfo(source, layout) : new Map(),
    // Build SEQ counters
    seqValues: opts.updateSequences ? computeSeqValues(source) : new Map(),
    // Build style → paragraphs index for STYLEREF
    styleIndex: buildStyleIndex(source),
    // Collect INDEX entries (XE fields) from all body content
    indexEntries: collectIndexEntries(source, layout),
    opts,
    toc: plan
  };

  // Update body content
  const newBody = updateBody(env);

  // If TOC was updated, register the TOC1..TOCn paragraph styles so the
  // cached TOC entries don't reference undefined styles. Only do this when
  // the document actually contains a TOC — otherwise we'd mutate doc
  // shape for no reason and break === comparisons in callers.
  let nextDoc: DocxDocument = source;
  if (newBody !== source.body) {
    nextDoc = { ...source, body: newBody };
  }
  if (plan) {
    nextDoc = ensureTocStyles(nextDoc, plan, tocInstructions);
  }
  return nextDoc;
}

/**
 * 仅更新 TOC（目录）字段。
 * 扫描文档中的标题（Heading1-9 样式或 outline level），
 * 生成目录缓存段落。
 * 目录页码来自本库的分页布局，见 {@link updateFields}。
 *
 * Does not modify `doc`; the result may share unchanged nodes with it — see
 * `Document.clone`.
 */
export function updateTableOfContents(
  doc: DocxDocument,
  options?: FieldUpdateOptions
): DocxDocument {
  checkHandle(doc, "document", "doc", "Transform.updateTableOfContents");
  checkFieldUpdateOptions(options, "Transform.updateTableOfContents");
  // Skip the entire pipeline (including TOC style registration) when the
  // document has no table of contents at all — callers rely on
  // `updateTableOfContents` returning the same reference in that case so they
  // can detect "nothing changed" with a `===` check.
  const instructions = listTocInstructions(doc);
  if (instructions.length === 0) {
    return doc;
  }

  const plan = planTableOfContents(doc, instructions);
  return settleTableOfContents(plan, instructions, options?.layoutOptions).doc;
}

/**
 * The document of `plan` with its tables of contents regenerated, and the
 * layout their page numbers come from — the layout *of that document*, entries
 * included, as Word paginates it: a table of contents that grows by a page
 * moves every heading after it by a page. Laid out again until the entries
 * read the same as the layout they were numbered from; a page number changing
 * width can rewrap an entry, so this can take a second round, rarely a third.
 */
function settleTableOfContents(
  plan: TocPlan,
  instructions: readonly string[],
  layoutOptions: LayoutOptions | undefined
): { doc: DocxDocument; layout: LayoutResult } {
  // The entries are laid out in the styles they will be written with, numbered
  // first from the quick estimate so the first full layout already holds them.
  // Each round still numbers the entries of the document as it stands, so a
  // table of contents that was already right comes back untouched.
  let doc = ensureTocStyles(plan.doc, plan, instructions);
  let laidOut = withTocEntries(doc, plan, layoutDocument(doc, layoutOptions));
  let layout = paginateDocument(laidOut, layoutOptions);
  for (let round = 0; round <= MAX_TOC_ROUNDS; round++) {
    const next = withTocEntries(doc, plan, layout);
    // New numbers that leave every line where it was move nothing after
    // them: the layout of `laidOut` is the layout of `next`.
    if (next === laidOut || samePrefixLayout(laidOut, next, layoutOptions)) {
      return { doc: next, layout };
    }
    if (round === MAX_TOC_ROUNDS) {
      break;
    }
    doc = laidOut = next;
    layout = paginateDocument(doc, layoutOptions);
  }
  return { doc, layout };
}

/**
 * Whether `a` and `b`, which differ only in some body items, lay out the same
 * up to the last of those — compared on the body cut after it, set in the
 * section it belongs to. Content after it then lands the same in both.
 */
function samePrefixLayout(
  a: DocxDocument,
  b: DocxDocument,
  layoutOptions: LayoutOptions | undefined
): boolean {
  if (a.body.length !== b.body.length || a.styles !== b.styles) {
    return false;
  }
  let last = -1;
  for (let i = 0; i < a.body.length; i++) {
    if (a.body[i] !== b.body[i]) {
      last = i;
    }
  }
  const cut = (doc: DocxDocument): DocxDocument => {
    let sectionProperties = doc.sectionProperties;
    for (let i = last; i < doc.body.length; i++) {
      const item = doc.body[i]!;
      if (item.type === "paragraph" && item.properties?.sectionProperties) {
        sectionProperties = item.properties.sectionProperties;
        break;
      }
    }
    return { ...doc, body: doc.body.slice(0, last + 1), sectionProperties };
  };
  return layoutFootprint(cut(a), layoutOptions) === layoutFootprint(cut(b), layoutOptions);
}

/** Rounds of {@link settleTableOfContents} after the first layout. */
const MAX_TOC_ROUNDS = 4;

/** `doc` with its tables of contents numbered from `layout`, or `doc` when they already are. */
function withTocEntries(doc: DocxDocument, plan: TocPlan, layout: LayoutResult): DocxDocument {
  const updated = doc.body.map(item =>
    item.type === "tableOfContents" ? updateTocContent(item, plan, layout) : item
  );
  const changed = updated.some((item, i) => item !== doc.body[i]);
  const body = updateSplitTocFields(changed ? updated : doc.body, plan, layout);
  return body !== doc.body ? { ...doc, body } : doc;
}

/**
 * The instruction of every table of contents in the body: its
 * `TableOfContents` blocks, and the `TOC` fields held in pieces among its
 * paragraphs.
 */
function listTocInstructions(doc: DocxDocument): string[] {
  const out: string[] = [];
  for (const item of doc.body) {
    if (item.type === "tableOfContents") {
      out.push(buildTocInstruction(item));
    }
  }
  for (const field of resolveSplitFields(doc.body)?.fields ?? []) {
    if (isTocInstruction(field.instruction) && field.end && !field.deleted) {
      out.push(field.instruction);
    }
  }
  return out;
}

function isTocInstruction(instruction: string): boolean {
  return /^\s*TOC\b/i.test(instruction);
}

/**
 * Register the paragraph styles the entries use — `TOC1`..`TOCn` for every
 * heading level listed, `TableofFigures` for captions — so the cached entries
 * don't reference undefined styles. Word logs a "missing referenced style"
 * warning on every entry otherwise.
 */
function ensureTocStyles(
  doc: DocxDocument,
  plan: TocPlan,
  instructions: readonly string[]
): DocxDocument {
  const usedTocLevels = new Set<number>();
  let figures = false;
  for (const instruction of instructions) {
    const used = plan.styles(instruction);
    for (const level of used.levels) {
      usedTocLevels.add(Math.max(1, Math.min(9, level)));
    }
    figures ||= used.figures;
  }
  const existingStyles = doc.styles ?? [];
  const definedIds = new Set(existingStyles.map(s => s.styleId));
  const stylesToAdd: StyleDef[] = [];
  for (const lvl of [...usedTocLevels].sort((a, b) => a - b)) {
    const id = `TOC${lvl}`;
    if (definedIds.has(id)) {
      continue;
    }
    stylesToAdd.push(tocStyle(lvl));
  }
  if (figures && !definedIds.has("TableofFigures")) {
    stylesToAdd.push(builtInStyle("TableofFigures")!);
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

interface IndexEntry {
  readonly term: string;
  readonly page: number;
}

type BookmarkMap = ReadonlyMap<string, BookmarkData>;
/** A `SEQ` field's number, by the field (whole) or its begin piece (in pieces). */
type SeqMap = ReadonlyMap<FieldContent, number>;

type ResolvedFieldOptions = Required<Omit<FieldUpdateOptions, "layoutOptions">> & {
  layoutOptions?: LayoutOptions;
};

/** What updating one field needs, computed once per update. */
interface UpdateEnv {
  readonly doc: DocxDocument;
  readonly layout: LayoutResult;
  readonly bookmarkInfo: BookmarkMap;
  readonly seqValues: SeqMap;
  readonly styleIndex: Map<string, { index: number; text: string }[]>;
  readonly indexEntries: IndexEntry[];
  readonly opts: ResolvedFieldOptions;
  /** The tables of contents to regenerate, when they are updated. */
  readonly toc: TocPlan | undefined;
}

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

  // A bookmark that starts between blocks (before a table, between rows…)
  // takes the text of the paragraph it opens before, as one inside a
  // paragraph takes that paragraph's text.
  let inParagraph = 0;
  let pending: string[] = [];
  walkBlocks(doc.body as BodyContent[], {
    visitBookmarkStart(bookmark) {
      if (inParagraph === 0) {
        pending.push(bookmark.name);
      }
    },
    enterParagraph(para) {
      inParagraph++;
      const text = extractParagraphText(para);
      for (const name of pending) {
        map.set(name, { text, page: layout.bookmarkPages.get(name) ?? 1 });
      }
      pending = [];
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
    },
    leaveParagraph() {
      inParagraph--;
    }
  });
  for (const marker of doc.bodyRangeMarkersAfter ?? []) {
    if (marker.type === "bookmarkStart") {
      pending.push(marker.name);
    }
  }
  for (const name of pending) {
    map.set(name, { text: "", page: layout.bookmarkPages.get(name) ?? 1 });
  }

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
// SEQ Value Computation
// =============================================================================

/** Compute all SEQ field values by traversing the document in order.
 *
 * Walks every field reachable from the body — whole or in pieces, inside
 * hyperlinks and tracked insertions, tables, SDTs, text boxes and TOC caches
 * — in reading order (`forEachField`), so SEQ counters in tabular figure /
 * table captions advance correctly. A field in a tracked deletion is not
 * counted. Each value is keyed by the field itself, so the update finds it
 * wherever it walks.
 */
function computeSeqValues(doc: DocxDocument): Map<FieldContent, number> {
  const counters = new Map<string, number>();
  const values = new Map<FieldContent, number>();

  forEachField(doc.body, field => {
    if (field.deleted) {
      return;
    }
    const { type, args } = parseFieldInstruction(field.instruction);
    if (type !== "SEQ") {
      return;
    }
    const { identifier, resetTo, repeat } = parseSeqSwitches(args);
    if (!identifier) {
      return;
    }
    if (resetTo !== null) {
      counters.set(identifier, resetTo);
    }
    const key = field.kind === "whole" ? field.content : field.field.begin;
    if (repeat) {
      // \c: repeat last value without incrementing
      values.set(key, counters.get(identifier) ?? 0);
    } else {
      const current = (counters.get(identifier) ?? 0) + 1;
      counters.set(identifier, current);
      values.set(key, current);
    }
  });

  return values;
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

/**
 * Collect all XE (Index Entry) fields from the document body, tagging each
 * with the page it falls on. Page numbers come from `layout.contentPages`
 * keyed by the enclosing top-level body index (the same mechanism used for
 * headings/TOC); entries inside nested structures inherit their outer block's
 * page, falling back to 1 when layout produced no page for that block.
 */
function collectIndexEntries(doc: DocxDocument, layout: LayoutResult): IndexEntry[] {
  const entries: IndexEntry[] = [];
  const { contentPages } = layout;

  for (let i = 0; i < doc.body.length; i++) {
    const page = contentPages[i] ?? 1;
    forEachField([doc.body[i]!], field => {
      if (field.deleted) {
        return;
      }
      const { type, args } = parseFieldInstruction(field.instruction);
      if (type === "XE") {
        const term = parseXeTerm(args);
        if (term) {
          entries.push({ term, page });
        }
      }
    });
  }

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

  // Merge entries that share the same term: a term marked on several pages
  // produces a single index line listing each distinct page (Word behaviour),
  // e.g. "widget, 1, 4" rather than two separate "widget" rows.
  const byTerm = new Map<string, Set<number>>();
  for (const entry of entries) {
    let pages = byTerm.get(entry.term);
    if (!pages) {
      pages = new Set<number>();
      byTerm.set(entry.term, pages);
    }
    pages.add(entry.page);
  }

  // Sort terms alphabetically; within each term sort pages numerically.
  const merged = [...byTerm.entries()]
    .map(([term, pages]) => ({ term, pages: [...pages].sort((a, b) => a - b) }))
    .sort((a, b) => a.term.localeCompare(b.term));

  const formatEntry = (e: { term: string; pages: number[] }): string =>
    `${e.term}\t${e.pages.join(", ")}`;

  // Check for \h switch (group by first letter with headings)
  const grouped = /\\h\b/i.test(args);

  if (!grouped) {
    return merged.map(formatEntry).join("\n");
  }

  // Group by first letter
  const lines: string[] = [];
  let currentLetter = "";

  for (const entry of merged) {
    const letter = entry.term.charAt(0).toUpperCase();
    if (letter !== currentLetter) {
      currentLetter = letter;
      if (lines.length > 0) {
        lines.push("");
      }
      lines.push(currentLetter);
    }
    lines.push(formatEntry(entry));
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
  instruction: string,
  args: string,
  bodyIndex: number,
  doc: DocxDocument,
  cellCtx: CellContext | undefined
): string {
  // The expression ends where the first switch begins.
  const switchAt = args.search(/\\[#*@]/);
  let expr = (switchAt < 0 ? args : args.slice(0, switchAt)).trim();

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

  return formatFieldResult(String(result), instruction);
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
        forEachField(cell.content, field => {
          found ||= parseFieldInstruction(field.instruction).type === "=";
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

// =============================================================================
// Body Update
// =============================================================================

function updateBody(env: UpdateEnv): readonly BodyContent[] {
  const body = env.doc.body;
  let changed = false;
  const newBody: BodyContent[] = [];

  for (let i = 0; i < body.length; i++) {
    const item = body[i];

    if (item.type === "paragraph") {
      // No cell context at the body level — formulas here can still call
      // SUM(ABOVE/LEFT) but findCellInTable will be used as a fallback
      // (likely returning null at the body level).
      const updated = updateParagraphFields(item, i, env, undefined);
      if (updated !== item) {
        changed = true;
      }
      newBody.push(updated);
    } else if (item.type === "tableOfContents" && env.toc) {
      const updated = updateTocContent(item, env.toc, env.layout);
      if (updated !== item) {
        changed = true;
      }
      newBody.push(updated);
    } else if (item.type === "table") {
      // Recurse into table cells to update fields within tables
      const updated = updateTableFields(item, i, env);
      if (updated !== item) {
        changed = true;
      }
      newBody.push(updated);
    } else {
      newBody.push(item);
    }
  }

  const blocks = changed ? newBody : body;
  // The fields held in pieces, across the paragraphs they span.
  return updateSplitFields(blocks, block => block, env, undefined, true);
}

/**
 * `blocks` — one list of blocks — with the result of every field it holds in
 * pieces updated: across the runs, wrappers and paragraphs between its
 * separate and end characters, which stay. `bodyIndex` maps the index of the
 * paragraph holding a field's begin to the body block it is in.
 */
function updateSplitFields<T extends object>(
  blocks: readonly T[],
  bodyIndex: (block: number) => number,
  env: UpdateEnv,
  cellCtx: CellContext | undefined,
  inBody = false
): readonly T[] {
  const resolved = resolveSplitFields(blocks);
  if (!resolved) {
    return blocks;
  }
  // Inner fields first: an `IF` testing a nested field tests what it updates to.
  const values = evaluateSplitFields(
    resolved,
    (field, instruction) => {
      if (field.deleted || !field.end || isTocInstruction(field.instruction)) {
        return undefined;
      }
      const current: FieldContent = { type: "field", instruction, cachedValue: field.result };
      return computeFieldValue(current, field.begin, bodyIndex(field.block), env, cellCtx)
        .cachedValue;
    },
    (content, within) =>
      computeFieldValue(content, content, bodyIndex(within.block), env, cellCtx).cachedValue ?? ""
  );
  return rewriteSplitFields(
    blocks,
    field => {
      if (field.deleted || !field.end) {
        return undefined;
      }
      if (isTocInstruction(field.instruction)) {
        // A table of contents lists the body: one in the body regenerates.
        if (!env.toc || !inBody) {
          return undefined;
        }
        const paragraphs = env.toc.entries(field.instruction, env.layout, page =>
          shownPageText(env.layout, env.doc, page)
        );
        const old = field.result.replace(/^\n+|\n+$/g, "");
        if (paragraphs.map(extractParagraphText).join("\n") === old) {
          return undefined;
        }
        return { replace: "result", paragraphs };
      }
      const value = values.get(field);
      return value === undefined || value === field.result
        ? undefined
        : { replace: "result", text: value };
    },
    resolved
  );
}

/**
 * The `TOC` fields held in pieces among the body's paragraphs, regenerated
 * (`updateTableOfContents`, which updates no other field).
 */
function updateSplitTocFields(
  body: readonly BodyContent[],
  plan: TocPlan,
  layout: LayoutResult
): readonly BodyContent[] {
  return rewriteSplitFields(body, field => {
    if (field.deleted || !field.end || !isTocInstruction(field.instruction)) {
      return undefined;
    }
    const paragraphs = plan.entries(field.instruction, layout, page =>
      shownPageText(layout, plan.doc, page)
    );
    const old = field.result.replace(/^\n+|\n+$/g, "");
    if (paragraphs.map(extractParagraphText).join("\n") === old) {
      return undefined;
    }
    return { replace: "result", paragraphs };
  });
}

/** Recursively update fields inside a table's cells. */
function updateTableFields(table: Table, bodyIndex: number, env: UpdateEnv): Table {
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
          const updated = updateParagraphFields(block, bodyIndex, env, cellCtx);
          if (updated !== block) {
            cellChanged = true;
          }
          newContent.push(updated);
        } else if (block.type === "table") {
          const nested = updateTableFields(block, bodyIndex, env);
          if (nested !== block) {
            cellChanged = true;
          }
          newContent.push(nested);
        } else {
          newContent.push(block);
        }
      }
      const content = updateSplitFields(
        cellChanged ? newContent : cell.content,
        () => bodyIndex,
        env,
        cellCtx
      );
      if (content !== cell.content) {
        rowChanged = true;
        return { ...cell, content } as TableCell;
      }
      return cell;
    });
    if (rowChanged) {
      tableChanged = true;
      return { ...row, cells: newCells };
    }
    return row;
  });

  return tableChanged ? { ...table, rows: newRows } : table;
}

// =============================================================================
// Paragraph Field Update
// =============================================================================

function updateParagraphFields(
  para: Paragraph,
  bodyIndex: number,
  env: UpdateEnv,
  cellCtx: CellContext | undefined
): Paragraph {
  let childrenChanged = false;
  const newChildren: ParagraphChild[] = [];

  const updateRun = (r: Run): Run => updateRunFields(r, bodyIndex, env, cellCtx);

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
        newChildren.push({ ...hl, children: newRuns } as ParagraphChild);
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
        newChildren.push({ ...wrap, run: updated } as ParagraphChild);
      } else {
        newChildren.push(child);
      }
      continue;
    }
    newChildren.push(child);
  }

  return childrenChanged ? { ...para, children: newChildren } : para;
}

// =============================================================================
// Run Field Update
// =============================================================================

function updateRunFields(
  run: Run,
  bodyIndex: number,
  env: UpdateEnv,
  cellCtx: CellContext | undefined
): Run {
  let contentChanged = false;
  const newContent: RunContent[] = [];

  for (const item of run.content) {
    // A field read in pieces is updated across the runs it spans
    // (`updateSplitFields`); its pieces stay as they are.
    if (item.type !== "field" || item.fieldChar) {
      newContent.push(item);
      continue;
    }

    const updated = computeFieldValue(item, item, bodyIndex, env, cellCtx);
    if (updated !== item) {
      contentChanged = true;
    }
    newContent.push(updated);
  }

  return contentChanged ? { ...run, content: newContent } : run;
}

// =============================================================================
// Field Value Computation
// =============================================================================

/**
 * A page-number field's text for `n`, honouring its `\\*` format switch and,
 * for `PAGE`, the section's `w:pgNumType w:fmt` (shared with the layout's
 * headers and footers). Any other switch (`\\#`, `\\* Ordinal`…) goes through
 * the general field switches (`formatFieldResult`).
 */
function pageFieldText(
  field: FieldContent,
  n: number,
  sectionFormat: PageNumberFormat | undefined
): string {
  const parsed = parsePageField(field.instruction);
  return (
    (parsed && formatPageField(parsed, n, sectionFormat)) ??
    formatFieldResult(String(n), field.instruction)
  );
}

/** The number `PAGE` shows on physical `page` (1-based) of `layout`. */
function shownPageNumber(layout: LayoutResult, doc: DocxDocument, page: number): number {
  return layoutPageNumbering(layout, doc).shown[page - 1] ?? page;
}

/**
 * Physical `page` of `layout` as Word refers to it from a `PAGEREF` or a
 * table of contents: the number `PAGE` shows there, in its section's format
 * (a front matter numbered i, ii, … lists its pages so, and a section that
 * restarts at 1 is listed from 1).
 */
function shownPageText(layout: LayoutResult, doc: DocxDocument, page: number): string {
  const numbering = layoutPageNumbering(layout, doc);
  const n = numbering.shown[page - 1] ?? page;
  const section = numbering.sections[numbering.pageSections[page - 1] ?? 0];
  return (
    formatPageField({ name: "PAGE", format: undefined }, n, section?.pageNumbering?.format) ??
    String(n)
  );
}

/** What `PAGE` needs from one layout, computed once per layout. */
interface LayoutPageNumbering {
  /** The number `PAGE` shows on each page (index 0 is page 1). */
  readonly shown: readonly number[];
  /** `collectSectionProperties(doc)`, indexed like `layout.contentSections`. */
  readonly sections: readonly (SectionProperties | undefined)[];
  /** The section each page is in (index 0 is page 1). */
  readonly pageSections: readonly number[];
}

/**
 * The number `PAGE` shows on each page of the layout, by the same
 * rule the page layout applies to headers and footers
 * (`displayedPageNumbers`). A page takes the section of the first body item
 * that lands on it; a page no item starts on continues the previous one's.
 */
function layoutPageNumbering(layout: LayoutResult, doc: DocxDocument): LayoutPageNumbering {
  const cached = pageNumberingCache.get(layout);
  if (cached) {
    return cached;
  }
  const firstSection: (number | undefined)[] = [];
  for (let i = 0; i < layout.contentPages.length; i++) {
    const page = layout.contentPages[i];
    if (page !== undefined && page >= 1 && firstSection[page - 1] === undefined) {
      firstSection[page - 1] = layout.contentSections[i];
    }
  }
  const slots: PageSectionSlot[] = [];
  for (let p = 0; p < layout.pageCount; p++) {
    const previous = slots[p - 1];
    const sectionIndex = firstSection[p] ?? previous?.sectionIndex ?? 0;
    const pageInSection =
      previous !== undefined && previous.sectionIndex === sectionIndex
        ? previous.pageInSection + 1
        : 1;
    slots.push({ sectionIndex, pageInSection });
  }
  const sections = collectSectionProperties(doc);
  const result = {
    shown: displayedPageNumbers(slots, sections),
    sections,
    pageSections: slots.map(slot => slot.sectionIndex)
  };
  pageNumberingCache.set(layout, result);
  return result;
}

/**
 * Per layout, so a document with many `PAGE` fields computes it
 * once. A layout is produced from one document snapshot, so it is a safe key.
 */
const pageNumberingCache = new WeakMap<LayoutResult, LayoutPageNumbering>();

/**
 * `field` with the result it computes to, or `field` itself when that is its
 * result already (or it is not one computed here). `key` is the field in the
 * document — a whole field, or a field in pieces' begin — that `SEQ` values
 * are kept by.
 */
function computeFieldValue(
  field: FieldContent,
  key: FieldContent,
  bodyIndex: number,
  env: UpdateEnv,
  cellCtx: CellContext | undefined
): FieldContent {
  const { layout, bookmarkInfo, seqValues, styleIndex, indexEntries, doc, opts } = env;
  const { type, args } = parseFieldInstruction(field.instruction);

  switch (type) {
    case "PAGE": {
      if (!opts.updatePageFields) {
        return field;
      }
      const page = layout.contentPages[bodyIndex] ?? 1;
      const sectionIdx = layout.contentSections[bodyIndex] ?? 0;
      const { shown: displayed, sections } = layoutPageNumbering(layout, doc);
      const numbering = sections[sectionIdx]?.pageNumbering;
      const shown = displayed[page - 1] ?? page;
      const value = pageFieldText(field, shown, numbering?.format);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "NUMPAGES": {
      if (!opts.updatePageFields) {
        return field;
      }
      const value = pageFieldText(field, layout.pageCount, undefined);
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
      const value = pageFieldText(field, pages, undefined);
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
      const value = formatFieldResult(data?.text ?? "", field.instruction);
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
      // The number the page shows; a `\*` switch formats the plain number instead.
      const value =
        page == null
          ? "?"
          : /\\\*/.test(args)
            ? formatFieldResult(String(shownPageNumber(layout, env.doc, page)), field.instruction)
            : shownPageText(layout, env.doc, page);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "SEQ": {
      if (!opts.updateSequences) {
        return field;
      }
      const seqVal = seqValues.get(key);
      if (seqVal == null) {
        return field;
      }
      const value = formatFieldResult(String(seqVal), field.instruction);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "IF": {
      const value = evaluateIfField(field.instruction);
      if (value === undefined || field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "QUOTE": {
      const tokens = tokenizeFieldCode(args);
      const end = tokens.findIndex(t => !t.quoted && t.text.startsWith("\\"));
      const text = (end < 0 ? tokens : tokens.slice(0, end)).map(t => t.text).join(" ");
      const value = formatFieldResult(text, field.instruction);
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
      const value = formatFieldResult(
        findStyleRef(styleIndex, styleName, bodyIndex),
        field.instruction
      );
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "INDEX": {
      // An index lists the pages its entries are on.
      const value = buildIndexContent(indexEntries, args);
      if (field.cachedValue === value) {
        return field;
      }
      return { ...field, cachedValue: value };
    }

    case "=": {
      const value = evaluateFormula(field.instruction, args, bodyIndex, doc, cellCtx);
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

// =============================================================================
// TOC Update
// =============================================================================

/**
 * A TableOfContents block with its entries regenerated as its instruction
 * asks (`planTableOfContents`), or the block itself when they read the same.
 */
function updateTocContent(
  toc: TableOfContents,
  plan: TocPlan,
  layout: LayoutResult
): TableOfContents {
  const cachedParagraphs = plan.entries(
    buildTocInstruction(toc),
    layout,
    page => shownPageText(layout, plan.doc, page),
    toc.leader
  );

  // Check if content actually changed. Entries holding the block's field
  // (`field: "entries"`) also hold the paragraphs it opens and closes in,
  // which carry no entry.
  const old = (toc.cachedParagraphs ?? []).filter(
    p => toc.field !== "entries" || extractParagraphText(p) !== ""
  );
  if (
    toc.cachedParagraphs &&
    old.length === cachedParagraphs.length &&
    cachedParagraphs.every((p, i) => extractParagraphText(p) === extractParagraphText(old[i]!))
  ) {
    return toc;
  }

  // The ranges the old entries opened or closed stay paired around the new ones.
  const { leading, trailing } = collectRangeMarkers(toc.cachedParagraphs ?? []);
  const updated: Mutable<TableOfContents> = {
    ...toc,
    cachedParagraphs: settleRangeMarkers<Paragraph>([...leading, ...cachedParagraphs, ...trailing])
  };
  // The old entries took the field they held with them: the new ones get the
  // block's own, with the instruction it had.
  if (updated.field === "entries") {
    delete updated.field;
  }
  return updated;
}
