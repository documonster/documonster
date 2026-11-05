/**
 * DOCX Module - Template Engine
 *
 * A template engine for DOCX documents supporting:
 * - Variable interpolation: {{name}}, {{user.name}}
 * - Conditionals: {{#if cond}}...{{/if}}, {{#if cond}}...{{else}}...{{/if}}
 * - Loops: {{#each items}}...{{/each}} with {{.}}, {{.prop}}, {{@index}}
 * - Table row loops: auto-duplicates table rows when {{#each}} is in a row
 * - Table column loops: {{#cols items}}...{{/cols}} duplicates columns
 * - Image placeholders: {{%image}} resolves to an inline image
 * - RichText placeholders: {{&richText}} resolves to paragraph children
 * - Sub-document insertion: {{>subDoc}} resolves to body content blocks
 * - Chart placeholders: {{^chart}} resolves to chart content
 *
 * Handles cross-run placeholders (Word often splits {{name}} across runs).
 */

import { isForbiddenKey } from "@utils/object";
import {
  checkBoolean,
  checkHandle,
  checkNamedValues,
  checkRecord,
  invalidArgument
} from "@word/core/argument-checks";
import { imageKeys, isSamePicture } from "@word/core/image-keys";
import { assertImageMediaType } from "@word/core/image-media";
import { cloneModel, utf8Encoder } from "@word/core/internal-utils";
import type { RangeAnchor } from "@word/core/range-markers";
import {
  collectRangeAnchors,
  keepClonedRangeMarkers,
  listRangeAnchors,
  repeatedRanges,
  replaceParagraphChildren,
  settleRangeMarkers,
  withParagraphChildren,
  withRangeMarkersBefore
} from "@word/core/range-markers";
import { isRun } from "@word/core/text-utils";
import { DocxTemplateError } from "@word/errors";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphChild,
  Run,
  RunProperties,
  Table,
  TableRow,
  TableCell,
  ImageDef,
  InlineImageContent,
  Chart,
  ChartContent,
  AltChunk,
  CommentReference,
  RangeMarker,
  StructuredDocumentTag
} from "@word/types";

// =============================================================================
// Public Types
// =============================================================================

/** Options for the template engine. */
export interface TemplateOptions {
  /** Custom delimiters (default: ["{{", "}}"]). */
  readonly delimiters?: readonly [string, string];
  /**
   * If true (default), throw `DocxTemplateError` when a placeholder cannot be
   * resolved. If false, unresolved variables render as the empty string and
   * unresolved control directives are skipped silently.
   */
  readonly strict?: boolean;
  /**
   * Also resolve the rich placeholders — image (`{{%path}}`), rich text
   * (`{{&path}}`), sub-document (`{{>path}}`), chart (`{{^path}}`) and HTML
   * chunks (`{{!path}}`) — in addition to variables, conditionals and loops,
   * wherever those can appear, loop bodies included. Each must be the whole
   * text of its paragraph. Default: false.
   *
   * Read by `Template.fill` and the entry points built on it; it does not
   * change whether the input is modified (`fill` never modifies it).
   */
  readonly richPlaceholders?: boolean;
}

/**
 * {@link TemplateOptions} as the deprecated entry points (`fillTemplate`,
 * `fillTemplateEnhanced` and the `…FromSource` / `…FromBuffer` forms) take
 * them. They predate `richPlaceholders`, and which placeholders each resolves
 * is fixed by its name, so the option is not part of their signature.
 *
 * @deprecated Use {@link TemplateOptions} with `Template.fill`. Removed in 1.0.
 */
export type LegacyTemplateOptions = Omit<TemplateOptions, "richPlaceholders">;

/** An image value to be inserted by the template engine. */
export interface TemplateImage {
  /** Image definition (data, fileName, mediaType). */
  readonly image: ImageDef;
  /** Width in EMU. */
  readonly width: number;
  /** Height in EMU. */
  readonly height: number;
  /** Optional alt text. */
  readonly altText?: string;
}

/** Rich text value: an array of runs to insert. */
export type TemplateRichText = readonly Run[];

/** Sub-document content: body blocks to insert. */
export type TemplateSubDocument = readonly BodyContent[];

/** Chart value to be inserted by the template engine. */
export interface TemplateChart {
  readonly chart: Chart;
  readonly altText?: string;
  readonly name?: string;
}

/** HTML chunk value to be inserted as an altChunk body content. */
export interface TemplateHtmlChunk {
  /** HTML content string. */
  readonly html: string;
  /** Optional content type (default: "text/html"). */
  readonly contentType?: string;
}

/** Information about a template tag found in the document. */
export interface TemplateTag {
  /** The raw tag expression (e.g. "name", "#each items", "%image"). */
  readonly expression: string;
  /** Tag type. */
  readonly type:
    | "variable"
    | "image"
    | "richText"
    | "subDocument"
    | "chart"
    | "htmlChunk"
    | "ifOpen"
    | "ifClose"
    | "else"
    | "eachOpen"
    | "eachClose"
    | "colsOpen"
    | "colsClose";
  /** Location hint (e.g. "body paragraph 3", "header"). */
  readonly location: string;
}

// =============================================================================
// Token Types (internal)
// =============================================================================

const enum TokenType {
  Text = 0,
  Variable = 1,
  IfOpen = 2,
  Else = 3,
  IfClose = 4,
  EachOpen = 5,
  EachClose = 6
}

interface TextToken {
  readonly type: TokenType.Text;
  readonly value: string;
}

interface VariableToken {
  readonly type: TokenType.Variable;
  readonly path: string;
}

interface IfOpenToken {
  readonly type: TokenType.IfOpen;
  readonly condition: string;
}

interface ElseToken {
  readonly type: TokenType.Else;
}

interface IfCloseToken {
  readonly type: TokenType.IfClose;
}

interface EachOpenToken {
  readonly type: TokenType.EachOpen;
  readonly collection: string;
}

interface EachCloseToken {
  readonly type: TokenType.EachClose;
}

type Token =
  | TextToken
  | VariableToken
  | IfOpenToken
  | ElseToken
  | IfCloseToken
  | EachOpenToken
  | EachCloseToken;

// =============================================================================
// Core: fillTemplate
// =============================================================================

/**
 * Fill a DOCX template document with data.
 *
 * Resolves variables, conditionals and loops in the body, headers, footers,
 * footnotes and endnotes — not rich placeholders, and not comment bodies.
 * Operates on the document model in place and returns it.
 *
 * @param doc - The parsed DocxDocument model.
 * @param data - Data to fill into the template.
 * @param options - Optional template settings.
 * @returns The same DocxDocument with placeholders resolved.
 */
export function fillTemplate(
  doc: DocxDocument,
  data: Record<string, unknown>,
  options?: LegacyTemplateOptions
): DocxDocument {
  checkFillArguments(data, options, "Template.fillTemplate");
  return fillInPlace(doc, data, options, {
    where: "Template.fillTemplate",
    rich: false,
    comments: false
  });
}

/**
 * The delimiters a template is read with: two non-empty strings that differ.
 * An empty one matches everywhere, and a pair that is one string cannot tell
 * where a tag ends.
 */
function checkDelimiters(options: unknown, where: string): void {
  checkRecord(options, "options", where, { optional: true });
  const delimiters = (options as TemplateOptions | undefined)?.delimiters;
  if (delimiters === undefined) {
    return;
  }
  if (
    !Array.isArray(delimiters) ||
    delimiters.length !== 2 ||
    typeof delimiters[0] !== "string" ||
    typeof delimiters[1] !== "string" ||
    delimiters[0] === "" ||
    delimiters[1] === "" ||
    delimiters[0] === delimiters[1]
  ) {
    invalidArgument(
      where,
      "options.delimiters",
      "two different non-empty strings [open, close]",
      delimiters
    );
  }
}

/** What every fill takes: data to fill with, and options. */
/**
 * The data a fill is given, as an object: `Template.fill` and
 * `Package.fillTemplate` take a plain object or a `Map` keyed by name, as the
 * other filling functions (`Document.mailMerge`, `Transform.fillFormFields`)
 * do. A `Map` becomes the top level of the data; nested values are read as
 * they are.
 */
export function fillData(data: unknown, where: string): Record<string, unknown> {
  const entries = checkNamedValues(data, "data", where);
  return data instanceof Map ? Object.fromEntries(entries) : (data as Record<string, unknown>);
}

function checkFillArguments(data: unknown, options: unknown, where: string): void {
  checkRecord(data, "data", where);
  checkDelimiters(options, where);
  const o = options as TemplateOptions | undefined;
  checkBoolean(o?.strict, "options.strict", where, { optional: true });
  checkBoolean(o?.richPlaceholders, "options.richPlaceholders", where, { optional: true });
}

/** Which parts of the template language a fill resolves, beyond the basic one. */
interface FillScope {
  /** The public call filling, which an error names. */
  readonly where: string;
  /** Rich placeholders (images, rich text, sub-documents, charts, HTML chunks). */
  readonly rich: boolean;
  /** Comment bodies, as well as the body, headers, footers and notes. */
  readonly comments: boolean;
}

/**
 * The fill itself: replaces the body, the children of every header and
 * footer, and the content of every note (and comment, when `scope` says so)
 * of `doc` in place, and adds the pictures rich placeholders inserted to its
 * images.
 */
function fillInPlace(
  doc: DocxDocument,
  data: Record<string, unknown>,
  options: LegacyTemplateOptions | undefined,
  scope: FillScope
): DocxDocument {
  const open = options?.delimiters?.[0] ?? "{{";
  const close = options?.delimiters?.[1] ?? "}}";
  const strict = options?.strict ?? true;

  const rich: RichState | undefined = scope.rich
    ? { images: createTemplateImageSink(doc.images, scope.where), depth: 0 }
    : undefined;
  const ctx: TemplateContext = { open, close, strict, data, rich };

  // Process body
  const newBody = processBodyContent(doc.body as BodyContent[], ctx, "body");
  (doc as { body: readonly BodyContent[] }).body = newBody;

  // Process headers
  if (doc.headers) {
    for (const [key, headerDef] of doc.headers) {
      const newChildren = processBlockList(
        headerDef.content.children as (Paragraph | Table)[],
        ctx,
        `header:${key}`
      );
      (headerDef.content as { children: readonly (Paragraph | Table)[] }).children = newChildren;
    }
  }

  // Process footers
  if (doc.footers) {
    for (const [key, footerDef] of doc.footers) {
      const newChildren = processBlockList(
        footerDef.content.children as (Paragraph | Table)[],
        ctx,
        `footer:${key}`
      );
      (footerDef.content as { children: readonly (Paragraph | Table)[] }).children = newChildren;
    }
  }

  // Process footnotes
  if (doc.footnotes) {
    for (const note of doc.footnotes) {
      if (note.id <= 0) {
        continue;
      }
      const processed = processBlockList(note.content as Paragraph[], ctx, `footnote:${note.id}`);
      (note as { content: readonly Paragraph[] }).content = processed as Paragraph[];
    }
  }

  // Process endnotes
  if (doc.endnotes) {
    for (const note of doc.endnotes) {
      if (note.id <= 0) {
        continue;
      }
      const processed = processBlockList(note.content as Paragraph[], ctx, `endnote:${note.id}`);
      (note as { content: readonly Paragraph[] }).content = processed as Paragraph[];
    }
  }

  if (scope.comments && doc.comments) {
    for (const comment of doc.comments) {
      const processed = processBlockList(
        comment.content as Paragraph[],
        ctx,
        `comment:${comment.id}`
      );
      (comment as { content: readonly Paragraph[] }).content = processed as Paragraph[];
    }
  }

  if (rich) {
    // Each added picture already has a key and a file name of its own (see
    // `createTemplateImageSink`).
    const images = [...(doc.images ?? []), ...rich.images.added];
    (doc as { images?: readonly ImageDef[] }).images = images.length > 0 ? images : undefined;
  }

  return doc;
}

// =============================================================================
// Internal Context
// =============================================================================

interface TemplateContext {
  readonly open: string;
  readonly close: string;
  readonly strict: boolean;
  readonly data: Record<string, unknown>;
  /** Set when rich placeholders are resolved. */
  readonly rich?: RichState;
}

/** What resolving rich placeholders accumulates over one fill. */
interface RichState {
  /**
   * Pictures inserted, to be added to the document's images. A picture
   * placed many times — by a loop, say — is stored once.
   */
  readonly images: TemplateImageSink;
  /** How many sub-documents the content being filled is nested in. */
  readonly depth: number;
}

/**
 * How deeply sub-document placeholders may nest. An inserted sub-document is
 * filled like the rest of the template, so one that inserts itself would
 * otherwise never finish.
 */
const MAX_SUB_DOCUMENT_DEPTH = 32;

// =============================================================================
// Cross-run text merging
// =============================================================================

/**
 * Merge all text content from a paragraph's runs into a single string.
 * This is intentionally simple — only extracts raw text from runs (no hyperlinks,
 * no tracked changes) as template placeholders only exist in direct run content.
 */
function extractParagraphText(para: Paragraph): string {
  let result = "";
  for (const child of para.children) {
    if (!isRun(child)) {
      continue;
    }
    for (const c of child.content) {
      if (c.type === "text" && "text" in c) {
        result += (c as { type: "text"; text: string }).text;
      }
    }
  }
  return result;
}

// =============================================================================
// Tokenizer
// =============================================================================

/**
 * The tokens of `text`. With `spans`, the source range of each token is
 * pushed to it, in step with the tokens.
 */
function tokenize(
  text: string,
  open: string,
  close: string,
  spans?: [start: number, end: number][]
): Token[] {
  const tokens: Token[] = [];
  let pos = 0;

  while (pos < text.length) {
    const startIdx = text.indexOf(open, pos);
    if (startIdx === -1) {
      tokens.push({ type: TokenType.Text, value: text.slice(pos) });
      spans?.push([pos, text.length]);
      break;
    }

    if (startIdx > pos) {
      tokens.push({ type: TokenType.Text, value: text.slice(pos, startIdx) });
      spans?.push([pos, startIdx]);
    }

    const endIdx = text.indexOf(close, startIdx + open.length);
    if (endIdx === -1) {
      // Unclosed delimiter — treat rest as text
      tokens.push({ type: TokenType.Text, value: text.slice(startIdx) });
      spans?.push([startIdx, text.length]);
      break;
    }

    const expr = text.slice(startIdx + open.length, endIdx).trim();
    pos = endIdx + close.length;
    spans?.push([startIdx, pos]);

    if (expr.startsWith("#if ")) {
      tokens.push({ type: TokenType.IfOpen, condition: expr.slice(4).trim() });
    } else if (expr === "else") {
      tokens.push({ type: TokenType.Else });
    } else if (expr === "/if") {
      tokens.push({ type: TokenType.IfClose });
    } else if (expr.startsWith("#each ")) {
      tokens.push({ type: TokenType.EachOpen, collection: expr.slice(6).trim() });
    } else if (expr === "/each") {
      tokens.push({ type: TokenType.EachClose });
    } else {
      tokens.push({ type: TokenType.Variable, path: expr });
    }
  }

  return tokens;
}

// =============================================================================
// Value resolution
// =============================================================================

function resolvePath(data: unknown, path: string): unknown {
  if (path === ".") {
    // In loop context, "." refers to the current item stored at key "."
    if (data != null && typeof data === "object" && Object.hasOwn(data as object, ".")) {
      return (data as Record<string, unknown>)["."];
    }
    return data;
  }
  const startPath = path.startsWith(".") ? path.slice(1) : path;
  const parts = startPath.split(".");
  let current: unknown = data;
  for (const part of parts) {
    if (current == null || typeof current !== "object") {
      return undefined;
    }
    // Use Object.hasOwn so a template path can never traverse prototype
    // properties (`{{constructor}}`, `{{__proto__.something}}`, …) and
    // surface JS internals into the rendered document. Arrays are
    // handled the same way — numeric indices are own properties.
    if (!Object.hasOwn(current as object, part)) {
      return undefined;
    }
    current = (current as Record<string, unknown>)[part];
  }
  return current;
}

function isTruthy(value: unknown): boolean {
  if (Array.isArray(value)) {
    return value.length > 0;
  }
  return !!value;
}

function valueToString(value: unknown): string {
  if (value == null) {
    return "";
  }
  if (typeof value === "object") {
    return JSON.stringify(value);
  }
  return String(value);
}

// =============================================================================
// Processing body content (handles block-level #each and #if)
// =============================================================================

function processBodyContent(
  blocks: BodyContent[],
  ctx: TemplateContext,
  location: string
): BodyContent[] {
  return settleRangeMarkers(processBodyItems(blocks, ctx, location));
}

/**
 * {@link processBodyContent} without settling: the range markers and comment
 * references of the blocks it removes (directive paragraphs, branches not
 * taken, loop bodies with no items) stay in the result as items of their own,
 * where the blocks were, so the ranges they open or close stay paired and no
 * comment loses its reference.
 */
function processBodyItems(
  blocks: BodyContent[],
  ctx: TemplateContext,
  location: string
): (BodyContent | RangeAnchor)[] {
  const result: (BodyContent | RangeAnchor)[] = [];
  let i = 0;

  while (i < blocks.length) {
    const block = blocks[i];

    if (block.type === "paragraph") {
      const text = extractParagraphText(block);

      // Check if this paragraph starts a block-level #each
      if (hasBlockDirective(text, ctx.open, ctx.close, "#each ")) {
        const collectionExpr = extractDirectiveArg(text, ctx.open, ctx.close, "#each ");
        const endIdx = findClosingBlock(blocks, i, ctx, "#each ", "/each");
        if (endIdx === -1) {
          throw new DocxTemplateError(
            `Unclosed {{#each ${collectionExpr}}}`,
            `#each ${collectionExpr}`,
            location,
            {
              tagName: `#each ${collectionExpr}`,
              paragraphIndex: i,
              sectionPath: location
            }
          );
        }

        const innerBlocks = blocks.slice(i + 1, endIdx);
        const items = resolvePath(ctx.data, collectionExpr);

        result.push(...listRangeAnchors([block]));
        if (Array.isArray(items) && items.length > 0) {
          const ranges = repeatedRanges(innerBlocks);
          for (let idx = 0; idx < items.length; idx++) {
            const itemData = buildLoopData(items[idx], idx, ctx.data);
            const innerCtx: TemplateContext = { ...ctx, data: itemData };
            const processed = processBodyItems(
              keepClonedRangeMarkers(cloneBlocks(innerBlocks), idx, items.length, ranges),
              innerCtx,
              `${location}[${idx}]`
            );
            result.push(...processed);
          }
        } else {
          result.push(...listRangeAnchors(innerBlocks));
        }
        result.push(...listRangeAnchors([blocks[endIdx]!]));
        if (!Array.isArray(items) && ctx.strict && items !== undefined && items !== null) {
          throw new DocxTemplateError(
            `{{#each ${collectionExpr}}} expects an array, got ${typeof items}`,
            `#each ${collectionExpr}`,
            location,
            {
              tagName: `#each ${collectionExpr}`,
              paragraphIndex: i,
              sectionPath: location
            }
          );
        }

        i = endIdx + 1;
        continue;
      }

      // Check if this paragraph starts a block-level #if
      if (hasBlockDirective(text, ctx.open, ctx.close, "#if ")) {
        const condExpr = extractDirectiveArg(text, ctx.open, ctx.close, "#if ");
        const { elseIdx, endIdx } = findIfBlock(blocks, i, ctx);
        if (endIdx === -1) {
          throw new DocxTemplateError(`Unclosed {{#if ${condExpr}}}`, `#if ${condExpr}`, location, {
            tagName: `#if ${condExpr}`,
            paragraphIndex: i,
            sectionPath: location
          });
        }

        const condValue = resolvePath(ctx.data, condExpr);
        const thenBlocks = blocks.slice(i + 1, elseIdx !== -1 ? elseIdx : endIdx);
        const elseBlocks = elseIdx !== -1 ? blocks.slice(elseIdx + 1, endIdx) : [];
        const taken = isTruthy(condValue);

        result.push(...listRangeAnchors([block]));
        if (taken) {
          result.push(...processBodyItems(cloneBlocks(thenBlocks), ctx, location));
        } else {
          result.push(...listRangeAnchors(thenBlocks));
        }
        if (elseIdx !== -1) {
          result.push(...listRangeAnchors([blocks[elseIdx]!]));
        }
        if (taken) {
          result.push(...listRangeAnchors(elseBlocks));
        } else {
          result.push(...processBodyItems(cloneBlocks(elseBlocks), ctx, location));
        }
        result.push(...listRangeAnchors([blocks[endIdx]!]));

        i = endIdx + 1;
        continue;
      }

      // Regular paragraph — process inline templates
      try {
        const replaced = ctx.rich ? resolveRichParagraph(block, ctx, location) : undefined;
        if (replaced) {
          result.push(...replaced);
        } else {
          result.push(processParagraph(block, ctx, `${location} para ${i}`));
        }
      } catch (err) {
        if (err instanceof DocxTemplateError && err.tagName === undefined) {
          throw new DocxTemplateError(err.message, err.placeholder, err.location, {
            cause: err.cause,
            tagName: err.placeholder,
            paragraphIndex: i,
            sectionPath: location
          });
        }
        throw err;
      }
      i++;
    } else if (block.type === "table") {
      result.push(...processTable(block, ctx, `${location} table ${i}`));
      i++;
    } else if (block.type === "sdt") {
      // Recursively process SDT content (StructuredDocumentTag)
      const sdt = block as StructuredDocumentTag;
      const processedContent = processBodyContent(
        sdt.content as BodyContent[],
        ctx,
        `${location} sdt ${i}`
      );
      result.push({
        ...sdt,
        content: processedContent as readonly (Paragraph | Run | Table)[]
      });
      i++;
    } else {
      result.push(block);
      i++;
    }
  }

  return result;
}

function processBlockList(
  blocks: (Paragraph | Table)[],
  ctx: TemplateContext,
  location: string
): (Paragraph | Table)[] {
  // Reuse body content processing, filtering to Paragraph | Table
  return processBodyContent(blocks as BodyContent[], ctx, location) as (Paragraph | Table)[];
}

// =============================================================================
// Paragraph processing (inline variable substitution)
// =============================================================================

function processParagraph(para: Paragraph, ctx: TemplateContext, location: string): Paragraph {
  // Merge all run text, resolve inline variables, then rebuild runs
  const fullText = extractParagraphText(para);

  // Quick check: does this paragraph contain any delimiters?
  if (!fullText.includes(ctx.open)) {
    return para;
  }

  // A field keeps its place among the text around it.
  if (holdsField(para)) {
    return processTextAroundFields(para, ctx, location);
  }

  // Tokenize and evaluate inline expressions
  const tokens = tokenize(fullText, ctx.open, ctx.close);
  const resolved = evaluateInlineTokens(tokens, ctx, location);

  // If the text didn't change, return as-is
  if (resolved === fullText) {
    return para;
  }

  // Rebuild the paragraph runs with the resolved text
  return rebuildParagraphText(para, resolved);
}

/** Whether any run of `para` holds a field, whole or in pieces. */
function holdsField(para: Paragraph): boolean {
  return para.children.some(child => isRun(child) && child.content.some(c => c.type === "field"));
}

/** A text content item of a run directly in the paragraph, outside every field. */
interface FreeText {
  readonly child: number;
  readonly content: number;
  readonly start: number;
  readonly end: number;
}

/**
 * {@link processParagraph} for a paragraph holding fields: the placeholders in
 * the text of its runs outside every field are resolved in place, each run
 * keeping its formatting and every field — whole, or in pieces with its
 * instruction and result — staying where it is, as does the text in a field's
 * result (which Word computes). A placeholder's value goes where it starts.
 */
function processTextAroundFields(
  para: Paragraph,
  ctx: TemplateContext,
  location: string
): Paragraph {
  const items: FreeText[] = [];
  let text = "";
  let depth = 0;
  para.children.forEach((child, ci) => {
    if (!isRun(child)) {
      return;
    }
    child.content.forEach((c, k) => {
      if (c.type === "field") {
        if (c.fieldChar === "begin") {
          depth++;
        } else if (c.fieldChar === "end") {
          depth = Math.max(0, depth - 1);
        }
      } else if (c.type === "text" && depth === 0) {
        items.push({ child: ci, content: k, start: text.length, end: text.length + c.text.length });
        text += c.text;
      }
    });
  });
  if (!text.includes(ctx.open)) {
    return para;
  }
  const spans: [number, number][] = [];
  const tokens = tokenize(text, ctx.open, ctx.close, spans);
  // Top-level segments of the source: literal text, or what an expression
  // (an inline block from its open tag to its close) evaluates to.
  const segments: { start: number; end: number; output?: string }[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    const [start, end] = spans[i]!;
    if (token.type === TokenType.Text) {
      segments.push({ start, end });
      continue;
    }
    let last = i;
    if (token.type === TokenType.IfOpen) {
      last = collectInlineIf(tokens, i).endIndex;
    } else if (token.type === TokenType.EachOpen) {
      last = collectInlineEach(tokens, i).endIndex;
    }
    // An unclosed block throws as it does in a paragraph without fields.
    const slice = last === -1 ? tokens.slice(i) : tokens.slice(i, last + 1);
    const output = evaluateInlineTokens(slice, ctx, location);
    segments.push({ start, end: last === -1 ? text.length : spans[last]![1], output });
    if (last === -1) {
      break;
    }
    i = last;
  }
  const replaced = new Map<string, string>();
  for (const item of items) {
    let out = "";
    for (const seg of segments) {
      if (seg.output === undefined) {
        const from = Math.max(seg.start, item.start);
        const to = Math.min(seg.end, item.end);
        if (from < to) {
          out += text.slice(from, to);
        }
      } else if (seg.start >= item.start && seg.start < item.end) {
        out += seg.output;
      }
    }
    if (out !== text.slice(item.start, item.end)) {
      replaced.set(`${item.child}:${item.content}`, out);
    }
  }
  if (replaced.size === 0) {
    return para;
  }
  const children: (ParagraphChild | RangeMarker)[] = [];
  para.children.forEach((child, ci) => {
    if (!isRun(child)) {
      children.push(child);
      return;
    }
    let changed = false;
    const content: Run["content"][number][] = [];
    child.content.forEach((c, k) => {
      const next = replaced.get(`${ci}:${k}`);
      if (next === undefined || c.type !== "text") {
        content.push(c);
        return;
      }
      changed = true;
      if (next !== "") {
        content.push({ ...c, text: next });
      }
    });
    if (!changed) {
      children.push(child);
    } else if (content.length > 0) {
      children.push({ ...child, content });
    } else {
      children.push(...(child.rangeMarkersBefore ?? []));
    }
  });
  return withParagraphChildren({ ...para, children: [] }, children);
}

/**
 * Evaluate inline tokens (only variables and inline if/each within a single paragraph line).
 * For simplicity, inline conditionals and loops are evaluated as text.
 */
function evaluateInlineTokens(tokens: Token[], ctx: TemplateContext, location: string): string {
  let result = "";
  let i = 0;

  while (i < tokens.length) {
    const token = tokens[i];

    switch (token.type) {
      case TokenType.Text:
        result += token.value;
        i++;
        break;

      case TokenType.Variable: {
        const path = token.path;
        let value: unknown;
        if (path === "@index") {
          value = resolvePath(ctx.data, "@index");
        } else {
          value = resolvePath(ctx.data, path);
        }
        if (value === undefined && ctx.strict) {
          throw new DocxTemplateError(`Unresolved variable: {{${path}}}`, path, location, {
            tagName: path,
            sectionPath: location
          });
        }
        result += valueToString(value);
        i++;
        break;
      }

      case TokenType.IfOpen: {
        // Inline if: collect tokens until matching /if
        const condValue = resolvePath(ctx.data, token.condition);
        const { trueBranch, falseBranch, endIndex } = collectInlineIf(tokens, i);
        if (endIndex === -1) {
          throw new DocxTemplateError(
            `Unclosed inline {{#if ${token.condition}}}`,
            `#if ${token.condition}`,
            location,
            { tagName: `#if ${token.condition}`, sectionPath: location }
          );
        }

        if (isTruthy(condValue)) {
          result += evaluateInlineTokens(trueBranch, ctx, location);
        } else {
          result += evaluateInlineTokens(falseBranch, ctx, location);
        }
        i = endIndex + 1;
        break;
      }

      case TokenType.EachOpen: {
        // Inline each: collect tokens until matching /each
        const items = resolvePath(ctx.data, token.collection);
        const { body, endIndex } = collectInlineEach(tokens, i);
        if (endIndex === -1) {
          throw new DocxTemplateError(
            `Unclosed inline {{#each ${token.collection}}}`,
            `#each ${token.collection}`,
            location,
            { tagName: `#each ${token.collection}`, sectionPath: location }
          );
        }

        if (Array.isArray(items)) {
          for (let idx = 0; idx < items.length; idx++) {
            const itemData = buildLoopData(items[idx], idx, ctx.data);
            const innerCtx: TemplateContext = { ...ctx, data: itemData };
            result += evaluateInlineTokens(body, innerCtx, location);
          }
        }
        i = endIndex + 1;
        break;
      }

      default:
        // Stray else/close tokens — skip
        i++;
        break;
    }
  }

  return result;
}

function collectInlineIf(
  tokens: Token[],
  startIdx: number
): {
  trueBranch: Token[];
  falseBranch: Token[];
  /** Index of the matching `{{/if}}` token, or -1 if no closing tag exists. */
  endIndex: number;
} {
  let depth = 0;
  const trueBranch: Token[] = [];
  const falseBranch: Token[] = [];
  let inElse = false;
  let endIndex = -1;

  for (let i = startIdx + 1; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === TokenType.IfOpen) {
      depth++;
      (inElse ? falseBranch : trueBranch).push(t);
    } else if (t.type === TokenType.IfClose) {
      if (depth === 0) {
        endIndex = i;
        break;
      }
      depth--;
      (inElse ? falseBranch : trueBranch).push(t);
    } else if (t.type === TokenType.Else && depth === 0) {
      inElse = true;
    } else {
      (inElse ? falseBranch : trueBranch).push(t);
    }
  }

  return { trueBranch, falseBranch, endIndex };
}

function collectInlineEach(
  tokens: Token[],
  startIdx: number
): {
  body: Token[];
  /** Index of the matching `{{/each}}` token, or -1 if no closing tag exists. */
  endIndex: number;
} {
  let depth = 0;
  const body: Token[] = [];
  let endIndex = -1;

  for (let i = startIdx + 1; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === TokenType.EachOpen) {
      depth++;
      body.push(t);
    } else if (t.type === TokenType.EachClose) {
      if (depth === 0) {
        endIndex = i;
        break;
      }
      depth--;
      body.push(t);
    } else {
      body.push(t);
    }
  }

  return { body, endIndex };
}

// =============================================================================
// Rebuild paragraph with new text content
// =============================================================================

/**
 * Rebuild a paragraph's run text content with a new resolved string.
 * Preserves the formatting of the first text run.
 */
function rebuildParagraphText(para: Paragraph, newText: string): Paragraph {
  // Find the first run with text content to use as formatting reference
  let refProperties: RunProperties | undefined;
  for (const child of para.children) {
    if (!isRun(child)) {
      continue;
    }
    for (const c of child.content) {
      if (c.type === "text") {
        refProperties = child.properties;
        break;
      }
    }
    if (refProperties !== undefined) {
      break;
    }
  }

  // If no text runs found, get properties from first run
  if (refProperties === undefined) {
    for (const child of para.children) {
      if (isRun(child)) {
        refProperties = child.properties;
        break;
      }
    }
  }

  // Build new children:
  //   - replace the first Run we see with `newRun` (carrying the resolved text);
  //   - drop every other Run (their text was already concatenated into newText);
  //   - keep non-Run children (bookmarkStart / bookmarkEnd / commentRangeStart /
  //     commentRangeEnd / hyperlink / insertedRun / movedToRun / etc.) at
  //     their original index. Reordering these silently miswires bookmark and
  //     comment ranges, since the start/end markers stop bracketing the
  //     intended content.
  // The markers that rode on the absorbed runs (range permissions) move onto
  // the surviving run, in source order, so their ranges stay paired.
  const absorbedMarkers: RangeMarker[] = [];
  for (const child of para.children) {
    if (isRun(child) && child.rangeMarkersBefore) {
      absorbedMarkers.push(...child.rangeMarkersBefore);
    }
  }
  const newRun: Run = {
    ...(absorbedMarkers.length > 0 ? { rangeMarkersBefore: absorbedMarkers } : {}),
    properties: refProperties,
    content: [{ type: "text", text: newText }]
  };

  const newChildren: ParagraphChild[] = [];
  let runReplaced = false;
  for (const child of para.children) {
    if (isRun(child)) {
      if (!runReplaced) {
        newChildren.push(newRun);
        runReplaced = true;
      }
      // Subsequent runs are absorbed into newRun's text — drop them.
    } else {
      newChildren.push(child);
    }
  }
  // If the paragraph had no runs at all (rare — typically all bookmark
  // markers), append newRun at the end so the rendered text still appears.
  if (!runReplaced) {
    newChildren.push(newRun);
  }

  return {
    ...para,
    children: newChildren
  };
}

// =============================================================================
// Table processing (row-level loops)
// =============================================================================

/**
 * The table filled — its row loops repeated, and with rich placeholders its
 * column loops too — followed by the comment references of any rows it
 * removed. A table cannot hold a reference outside a cell, so they are left
 * for the content around the table to take (see {@link settleRangeMarkers}).
 */
function processTable(
  source: Table,
  ctx: TemplateContext,
  location: string
): (Table | CommentReference)[] {
  const references: CommentReference[] = [];
  const table = ctx.rich ? processTableColumnLoop(source, ctx, references) : source;
  const newRows: TableRow[] = [];
  // Range markers of removed rows, waiting for the next row that is kept.
  let pending: RangeMarker[] = [];
  const pushRow = (row: TableRow): void => {
    newRows.push(withRangeMarkersBefore(row, pending));
    pending = [];
  };
  const dropRows = (rows: readonly TableRow[]): void => {
    for (const anchor of listTableRowAnchors(rows)) {
      if (anchor.type === "commentReference") {
        references.push(anchor);
      } else {
        pending.push(anchor);
      }
    }
  };

  let i = 0;
  while (i < table.rows.length) {
    const row = table.rows[i];
    const rowText = extractRowText(row);

    // Check if this row contains a #each directive
    const openTag = [...rowText.matchAll(eachTagPatterns(ctx).open)][0];
    if (openTag) {
      const collectionExpr = openTag[1]!.trim();

      if (collectionExpr) {
        // Find the closing row
        const endRowIdx = findClosingTableRow(table.rows, i, ctx);
        if (endRowIdx === -1) {
          throw new DocxTemplateError(
            `Unclosed {{#each ${collectionExpr}}} in table row`,
            `#each ${collectionExpr}`,
            location,
            {
              tagName: `#each ${collectionExpr}`,
              paragraphIndex: i,
              sectionPath: location
            }
          );
        }

        const items = resolvePath(ctx.data, collectionExpr);

        if (!Array.isArray(items) || items.length === 0) {
          // The loop rows go; the ranges they opened or closed stay paired.
          dropRows(table.rows.slice(i, endRowIdx + 1));
        } else {
          // If the loop is on a single row (open and close on same row)
          if (endRowIdx === i) {
            // Tags in one paragraph are filled as an inline loop, as before;
            // tags in different paragraphs go before the row is filled, as
            // over several rows, or they read as unclosed inline loops.
            const tags = eachTagPatterns(ctx);
            const templateRow = loopTagsShareParagraph(row, tags)
              ? row
              : removeRowTag(removeMatchingRowClose(row, tags), tags.open, "first");
            const ranges = repeatedRanges(templateRow);
            for (let idx = 0; idx < items.length; idx++) {
              const itemData = buildLoopData(items[idx], idx, ctx.data);
              const innerCtx: TemplateContext = { ...ctx, data: itemData };
              const clonedRow = keepClonedRangeMarkers(
                cloneRow(templateRow),
                idx,
                items.length,
                ranges
              );
              const processedRow = processTableRow(clonedRow, innerCtx, location);
              // Strip the #each and /each tags from the row
              const cleanedRow = stripDirectivesFromRow(processedRow, ctx);
              pushRow(cleanedRow);
            }
          } else {
            // Multi-row loop: duplicate all rows between open and close
            // The opening tag sits in the first row and the closing one in the
            // last; each row is processed on its own, so remove the loop's own
            // tags first or they read as an unclosed inline loop.
            const templateRows = table.rows.slice(i, endRowIdx + 1);
            const tags = eachTagPatterns(ctx);
            templateRows[0] = removeRowTag(templateRows[0]!, tags.open, "first");
            templateRows[templateRows.length - 1] = removeRowTag(
              templateRows[templateRows.length - 1]!,
              tags.close,
              "last"
            );
            const ranges = repeatedRanges(templateRows);
            for (let idx = 0; idx < items.length; idx++) {
              const itemData = buildLoopData(items[idx], idx, ctx.data);
              const innerCtx: TemplateContext = { ...ctx, data: itemData };
              // One loop iteration's rows are one copy of the repeated content.
              const clonedRows = keepClonedRangeMarkers(
                templateRows.map(cloneRow),
                idx,
                items.length,
                ranges
              );
              for (const clonedRow of clonedRows) {
                const processedRow = processTableRow(clonedRow, innerCtx, location);
                const cleanedRow = stripDirectivesFromRow(processedRow, ctx);
                pushRow(cleanedRow);
              }
            }
          }
        }

        i = endRowIdx + 1;
        continue;
      }
    }

    // Regular row: process inline templates in each cell
    const processedRow = processTableRow(row, ctx, location);
    pushRow(processedRow);
    i++;
  }

  const filled: Table = {
    ...table,
    rows: newRows,
    ...(pending.length > 0
      ? { rangeMarkersAfter: [...pending, ...(table.rangeMarkersAfter ?? [])] }
      : {})
  };
  return [filled, ...references];
}

/**
 * Every range marker and comment reference in `rows` — on and between them
 * and in their cells — in source order.
 */
function listTableRowAnchors(rows: readonly TableRow[]): RangeAnchor[] {
  return listRangeAnchors([{ type: "table", rows } as unknown as Table]);
}

function processTableRow(row: TableRow, ctx: TemplateContext, location: string): TableRow {
  const newCells: TableCell[] = [];
  for (const cell of row.cells) {
    newCells.push({ ...cell, content: processCellContent(cell.content, ctx, location) });
  }
  return { ...row, cells: newCells };
}

/**
 * A cell's paragraphs filled inline (block directives are not looked for in a
 * cell; a row's loop tags are handled by its table) and its nested tables
 * filled as tables.
 */
function processCellContent(
  content: readonly (Paragraph | Table)[],
  ctx: TemplateContext,
  location: string
): (Paragraph | Table)[] {
  const items: (BodyContent | RangeAnchor)[] = [];
  for (const block of content) {
    if (block.type === "paragraph") {
      const replaced = ctx.rich ? resolveRichParagraph(block, ctx, location) : undefined;
      if (replaced) {
        items.push(...replaced);
      } else {
        items.push(processParagraph(block, ctx, location));
      }
    } else if (block.type === "table") {
      items.push(...processTable(block, ctx, location));
    } else {
      items.push(block as BodyContent);
    }
  }
  return settleRangeMarkers(items) as (Paragraph | Table)[];
}

function extractRowText(row: TableRow): string {
  return rowParagraphs(row)
    .map(({ para }) => extractParagraphText(para))
    .join("");
}

/**
 * The paragraphs a table row's loop tags are looked for in: the top-level
 * paragraphs of every cell. A nested table is a scope of its own — its loops
 * are expanded when that table is processed — so detection, matching and
 * removal all leave it alone.
 */
function rowParagraphs(row: TableRow): { c: number; b: number; para: Paragraph }[] {
  const out: { c: number; b: number; para: Paragraph }[] = [];
  row.cells.forEach((cell, c) => {
    cell.content.forEach((block, b) => {
      if (block.type === "paragraph") {
        out.push({ c, b, para: block });
      }
    });
  });
  return out;
}

/** The `{{#each …}}` (capturing its argument) and `{{/each}}` tags, spacing inside the delimiters allowed. */
function eachTagPatterns(ctx: TemplateContext): { open: RegExp; close: RegExp } {
  const open = escapeRegex(ctx.open);
  const close = escapeRegex(ctx.close);
  return {
    open: new RegExp(`${open}\\s*#each\\s+([^${escapeRegex(ctx.close.charAt(0))}]*)${close}`, "g"),
    close: new RegExp(`${open}\\s*/each\\s*${close}`, "g")
  };
}

/**
 * `para` without the run text at the offsets `mask` marks, where offsets
 * count the text of its top-level runs ({@link extractParagraphText}). Each
 * run keeps its own formatting, and every other child (a hyperlink, a field,
 * a range marker) stays as it is. A run left empty is dropped; the range
 * markers it carried stay paired.
 */
function removeMaskedText(para: Paragraph, mask: readonly boolean[]): Paragraph {
  let offset = 0;
  const items: (ParagraphChild | RangeMarker)[] = [];
  let changed = false;
  for (const child of para.children) {
    if (!isRun(child)) {
      items.push(child);
      continue;
    }
    const content: Run["content"][number][] = [];
    let runChanged = false;
    for (const c of child.content) {
      if (c.type !== "text") {
        content.push(c);
        continue;
      }
      let text = "";
      for (let i = 0; i < c.text.length; i++) {
        if (!mask[offset + i]) {
          text += c.text[i];
        }
      }
      offset += c.text.length;
      if (text === c.text) {
        content.push(c);
        continue;
      }
      runChanged = true;
      if (text.length > 0) {
        content.push({ ...c, text });
      }
    }
    if (!runChanged) {
      items.push(child);
      continue;
    }
    changed = true;
    if (content.length > 0) {
      items.push({ ...child, content });
    } else {
      items.push(...(child.rangeMarkersBefore ?? []));
    }
  }
  return changed ? withParagraphChildren(para, items) : para;
}

/** A mask over `text` marking every match of `patterns`. */
function maskMatches(text: string, matches: Iterable<RegExpMatchArray>): boolean[] {
  const mask = Array.from({ length: text.length }, () => false);
  for (const m of matches) {
    mask.fill(true, m.index!, m.index! + m[0].length);
  }
  return mask;
}

function findClosingTableRow(
  rows: readonly TableRow[],
  startIdx: number,
  ctx: TemplateContext
): number {
  // We start AFTER having seen exactly one `{{#each ...}}` token, so the
  // virtual depth at the boundary into row `startIdx` is 1. Each row can
  // contain any mix of `{{#each ...}}` opens and `{{/each}}` closes; we
  // increment the depth by `opens` and decrement by `closes`. The first
  // row where depth hits zero is the closing row.
  //
  // A previous implementation lumped both adjustments together and
  // returned as soon as `depth === 0 && closes > 0`, which mis-handled
  // rows containing two opens and one close (depth ends at 1, but the
  // intermediate calculation hit a transient zero).
  let depth = 1;
  const tags = eachTagPatterns(ctx);
  for (let i = startIdx; i < rows.length; i++) {
    const text = extractRowText(rows[i]);
    let opens = [...text.matchAll(tags.open)].length;
    const closes = [...text.matchAll(tags.close)].length;
    // The opening `{{#each ...}}` of the current loop is in row
    // startIdx; don't count it as opening a new nested loop.
    if (i === startIdx && opens > 0) {
      opens--;
    }

    // Walk the row token-by-token in opens-then-closes order. We don't
    // know the actual interleave from raw counts, so we apply opens
    // first, then closes — opens always come before their matching
    // close in valid templates, and a row that closes more than it
    // opens (other than the top-level close in this very row) is
    // malformed input we should still terminate on.
    depth += opens;
    if (depth <= 0 && opens === 0 && closes === 0) {
      // Defensive: should never happen because we entered with depth=1.
      return -1;
    }
    if (closes >= depth) {
      // Consume `depth` worth of closes in this row; the remainder, if
      // any, is malformed but stops the search here.
      return i;
    }
    depth -= closes;
  }
  return -1;
}

/**
 * `row` without the first (`which` "first") or last match of `tag` — the
 * loop's own `{{#each …}}` or `{{/each}}` — in its {@link rowParagraphs}.
 * The tag is removed from the runs that hold it, so the rest of the
 * paragraph keeps its formatting.
 */
function removeRowTag(row: TableRow, tag: RegExp, which: "first" | "last"): TableRow {
  const order = rowParagraphs(row);
  if (which === "last") {
    order.reverse();
  }
  for (const { c, b, para } of order) {
    const matches = [...extractParagraphText(para).matchAll(tag)];
    const match = which === "first" ? matches[0] : matches[matches.length - 1];
    if (match) {
      const text = extractParagraphText(para);
      return replaceRowParagraph(row, c, b, removeMaskedText(para, maskMatches(text, [match])));
    }
  }
  return row;
}

/**
 * Where the `{{/each}}` matching a row's first `{{#each …}}` sits: the
 * row's tags are walked in document order with a depth counter, so a loop
 * nested inside the first one does not stand in for its close. `undefined`
 * when the row does not close it.
 */
function findMatchingRowClose(
  row: TableRow,
  tags: { open: RegExp; close: RegExp }
): { c: number; b: number; index: number; length: number; sameParagraph: boolean } | undefined {
  let depth = 0;
  let openPara: { c: number; b: number } | undefined;
  for (const { c, b, para } of rowParagraphs(row)) {
    const text = extractParagraphText(para);
    const found = [
      ...[...text.matchAll(tags.open)].map(m => ({ m, delta: 1 })),
      ...[...text.matchAll(tags.close)].map(m => ({ m, delta: -1 }))
    ].sort((x, y) => x.m.index! - y.m.index!);
    for (const { m, delta } of found) {
      if (!openPara) {
        if (delta < 0) {
          continue;
        }
        openPara = { c, b };
      }
      depth += delta;
      if (depth === 0) {
        return {
          c,
          b,
          index: m.index!,
          length: m[0].length,
          sameParagraph: openPara.c === c && openPara.b === b
        };
      }
    }
  }
  return undefined;
}

/** Whether a row's first `{{#each …}}` is closed in its own paragraph. */
function loopTagsShareParagraph(row: TableRow, tags: { open: RegExp; close: RegExp }): boolean {
  return findMatchingRowClose(row, tags)?.sameParagraph ?? false;
}

/** `row` without the `{{/each}}` matching its first `{{#each …}}`. */
function removeMatchingRowClose(row: TableRow, tags: { open: RegExp; close: RegExp }): TableRow {
  const close = findMatchingRowClose(row, tags);
  if (!close) {
    return removeRowTag(row, tags.close, "last");
  }
  const para = rowParagraphs(row).find(p => p.c === close.c && p.b === close.b)!.para;
  const text = extractParagraphText(para);
  const match = [...text.matchAll(tags.close)].find(m => m.index === close.index)!;
  return replaceRowParagraph(
    row,
    close.c,
    close.b,
    removeMaskedText(para, maskMatches(text, [match]))
  );
}

function replaceRowParagraph(row: TableRow, c: number, b: number, para: Paragraph): TableRow {
  return {
    ...row,
    cells: row.cells.map((cell, i) =>
      i === c
        ? { ...cell, content: cell.content.map((block, j) => (j === b ? para : block)) }
        : cell
    )
  };
}

/**
 * `row` without any `{{#each …}}` or `{{/each}}` tag in its
 * {@link rowParagraphs}, and without the whitespace a removed tag leaves at
 * the start or end of its paragraph. The tags are removed from the runs that
 * hold them, so the rest of each paragraph keeps its formatting.
 */
function stripDirectivesFromRow(row: TableRow, ctx: TemplateContext): TableRow {
  const tags = eachTagPatterns(ctx);
  let out = row;
  for (const { c, b, para } of rowParagraphs(row)) {
    const text = extractParagraphText(para);
    const matches = [...text.matchAll(tags.open), ...text.matchAll(tags.close)];
    if (matches.length === 0) {
      continue;
    }
    const mask = maskMatches(text, matches);
    for (let i = 0; i < text.length && (mask[i] || /\s/.test(text[i]!)); i++) {
      mask[i] = true;
    }
    for (let i = text.length - 1; i >= 0 && (mask[i] || /\s/.test(text[i]!)); i--) {
      mask[i] = true;
    }
    out = replaceRowParagraph(out, c, b, removeMaskedText(para, mask));
  }
  return out;
}

// =============================================================================
// Block-level directive helpers
// =============================================================================

function hasBlockDirective(text: string, open: string, close: string, directive: string): boolean {
  // Block-level directives must occupy the *entire* paragraph (modulo
  // surrounding whitespace). A paragraph like "Hello {{#if x}}A{{/if}}"
  // contains an *inline* directive and must not be re-classified as a block
  // directive — otherwise the engine tries to find the block close on
  // following paragraphs and emits "Unclosed {{#if x}}" errors.
  const trimmed = text.trim();
  const opener = open + directive;
  if (!trimmed.startsWith(opener)) {
    return false;
  }
  // The opener tag must close before any other text follows. e.g. valid:
  //   "{{#if x}}"          → block
  //   "{{#if x}}\n"        → block
  //   "{{#if x}}foo"       → inline (has trailing content on same paragraph)
  const tagEnd = trimmed.indexOf(close, opener.length);
  if (tagEnd === -1) {
    return false;
  }
  const afterTag = trimmed.slice(tagEnd + close.length).trim();
  return afterTag.length === 0;
}

function extractDirectiveArg(text: string, open: string, close: string, directive: string): string {
  const startIdx = text.indexOf(open + directive);
  if (startIdx === -1) {
    // Try with space variations
    const idx = text.indexOf(open + " " + directive);
    if (idx === -1) {
      return "";
    }
    const argStart = idx + open.length + 1 + directive.length;
    const endIdx = text.indexOf(close, argStart);
    if (endIdx === -1) {
      return "";
    }
    return text.slice(argStart, endIdx).trim();
  }
  const argStart = startIdx + open.length + directive.length;
  const endIdx = text.indexOf(close, argStart);
  if (endIdx === -1) {
    return "";
  }
  return text.slice(argStart, endIdx).trim();
}

function extractDirectiveFromText(
  text: string,
  open: string,
  close: string,
  directive: string
): string | null {
  const pattern = open + directive;
  const idx = text.indexOf(pattern);
  if (idx === -1) {
    return null;
  }
  const argStart = idx + pattern.length;
  const endIdx = text.indexOf(close, argStart);
  if (endIdx === -1) {
    return null;
  }
  return text.slice(argStart, endIdx).trim();
}

function findClosingBlock(
  blocks: readonly BodyContent[],
  startIdx: number,
  ctx: TemplateContext,
  openDir: string,
  closeDir: string
): number {
  let depth = 0;
  for (let i = startIdx + 1; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type !== "paragraph") {
      continue;
    }
    const text = extractParagraphText(block as Paragraph);
    if (hasBlockDirective(text, ctx.open, ctx.close, openDir)) {
      depth++;
    }
    if (
      text.includes(ctx.open + closeDir + ctx.close) ||
      text.includes(ctx.open + " " + closeDir + " " + ctx.close)
    ) {
      if (depth === 0) {
        return i;
      }
      depth--;
    }
  }
  return -1;
}

function findIfBlock(
  blocks: readonly BodyContent[],
  startIdx: number,
  ctx: TemplateContext
): { elseIdx: number; endIdx: number } {
  let depth = 0;
  let elseIdx = -1;

  for (let i = startIdx + 1; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type !== "paragraph") {
      continue;
    }
    const text = extractParagraphText(block as Paragraph);

    // A single paragraph can legitimately contain both `{{#if cond}}` and
    // `{{/if}}` (an inline if). Count opens and closes independently so the
    // pairing logic stays correct in that case — the previous else-if chain
    // would advance `depth` and then skip the close, leaving the block
    // permanently unterminated.
    const openCount = countOccurrences(text, ctx.open + "#if ");
    const closeCount =
      countOccurrences(text, ctx.open + "/if" + ctx.close) +
      countOccurrences(text, ctx.open + " /if " + ctx.close);

    if (closeCount > 0) {
      // Account for inline opens first, then balance closes against the
      // outstanding depth.
      depth += openCount;
      for (let c = 0; c < closeCount; c++) {
        if (depth === 0) {
          return { elseIdx, endIdx: i };
        }
        depth--;
      }
    } else if (openCount > 0) {
      depth += openCount;
    } else if (
      depth === 0 &&
      (text.trim() === ctx.open + "else" + ctx.close ||
        text.includes(ctx.open + "else" + ctx.close))
    ) {
      elseIdx = i;
    }
  }

  return { elseIdx, endIdx: -1 };
}

// =============================================================================
// Loop data construction
// =============================================================================

function buildLoopData(
  item: unknown,
  index: number,
  parentData: Record<string, unknown>
): Record<string, unknown> {
  // Build the loop context on a null-prototype object so that an item
  // like `{ "__proto__": { injected: "X" } }` (which JSON.parse will
  // produce as an own enumerable data property — a known
  // prototype-pollution vector) cannot retarget the result's prototype
  // and leak hidden fields into the rendered document.
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Object.keys(parentData)) {
    if (isForbiddenKey(key)) {
      // Skip dangerous keys at the parent level too.
      continue;
    }
    result[key] = parentData[key];
  }
  result["@index"] = index;
  result["."] = item;

  // If item is an object, spread its OWN properties so {{.name}} and
  // {{name}} both work. Skip the same dangerous key set defensively.
  if (item != null && typeof item === "object" && !Array.isArray(item)) {
    const obj = item as Record<string, unknown>;
    for (const key of Object.keys(obj)) {
      if (isForbiddenKey(key)) {
        continue;
      }
      // Object.hasOwn ensures we never copy something that came from the
      // prototype (Object.keys already returns only own enumerables, but
      // belt-and-braces).
      if (Object.hasOwn(obj, key)) {
        result[key] = obj[key];
      }
    }
  }

  return result;
}

// =============================================================================
// Cloning helpers
// =============================================================================

function cloneBlocks(blocks: BodyContent[]): BodyContent[] {
  return cloneModel(blocks);
}

function cloneRow(row: TableRow): TableRow {
  return cloneModel(row);
}

// =============================================================================
// Utility
// =============================================================================

function countOccurrences(text: string, search: string): number {
  let count = 0;
  let pos = 0;
  while (true) {
    const idx = text.indexOf(search, pos);
    if (idx === -1) {
      break;
    }
    count++;
    pos = idx + search.length;
  }
  return count;
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// =============================================================================
// Tag Listing: discover all template placeholders in a document
// =============================================================================

/**
 * List all template tags found in a document.
 *
 * Scans body, headers, footers, footnotes, and endnotes for template
 * placeholders — not comment bodies, which the deprecated fills it pairs with
 * leave alone. Useful for validating data objects against templates.
 *
 * @param doc - The parsed DocxDocument model.
 * @param options - Optional settings (custom delimiters).
 * @returns Array of all discovered template tags.
 */
export function listTemplateTags(
  doc: DocxDocument,
  options?: Pick<TemplateOptions, "delimiters">
): TemplateTag[] {
  return scanTemplateTags(doc, options, false);
}

/**
 * List every template tag in a document: in the body, headers, footers,
 * footnotes, endnotes and comment bodies — everywhere `fill` fills.
 * Useful for validating data objects against templates.
 *
 * @param doc - The parsed DocxDocument model.
 * @param options - Optional settings (custom delimiters).
 * @returns Every tag, with where it was found (`"comment:3 paragraph 0"`, …).
 */
export function listTags(
  doc: DocxDocument,
  options?: Pick<TemplateOptions, "delimiters">
): TemplateTag[] {
  checkHandle(doc, "document", "doc", "Template.listTags");
  return scanTemplateTags(doc, options, true);
}

function scanTemplateTags(
  doc: DocxDocument,
  options: Pick<TemplateOptions, "delimiters"> | undefined,
  includeComments: boolean
): TemplateTag[] {
  checkDelimiters(options, includeComments ? "Template.listTags" : "Template.listTemplateTags");
  const open = options?.delimiters?.[0] ?? "{{";
  const close = options?.delimiters?.[1] ?? "}}";
  const tags: TemplateTag[] = [];

  function scanBlocks(blocks: readonly BodyContent[], location: string): void {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === "paragraph") {
        scanParagraph(block, `${location} paragraph ${i}`);
      } else if (block.type === "table") {
        scanTable(block, `${location} table ${i}`);
      }
    }
  }

  function scanParagraph(para: Paragraph, location: string): void {
    const text = extractParagraphText(para);
    scanText(text, location);
  }

  function scanTable(table: Table, location: string): void {
    for (let r = 0; r < table.rows.length; r++) {
      const row = table.rows[r];
      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        for (let b = 0; b < cell.content.length; b++) {
          const block = cell.content[b];
          if (block.type === "paragraph") {
            scanParagraph(block as Paragraph, `${location} row ${r} cell ${c}`);
          } else if (block.type === "table") {
            scanTable(block as Table, `${location} row ${r} cell ${c}`);
          }
        }
      }
    }
  }

  function scanText(text: string, location: string): void {
    let pos = 0;
    while (pos < text.length) {
      const startIdx = text.indexOf(open, pos);
      if (startIdx === -1) {
        break;
      }
      const endIdx = text.indexOf(close, startIdx + open.length);
      if (endIdx === -1) {
        break;
      }
      const expr = text.slice(startIdx + open.length, endIdx).trim();
      pos = endIdx + close.length;

      if (expr.startsWith("#if ")) {
        tags.push({ expression: expr, type: "ifOpen", location });
      } else if (expr === "else") {
        tags.push({ expression: expr, type: "else", location });
      } else if (expr === "/if") {
        tags.push({ expression: expr, type: "ifClose", location });
      } else if (expr.startsWith("#each ")) {
        tags.push({ expression: expr, type: "eachOpen", location });
      } else if (expr === "/each") {
        tags.push({ expression: expr, type: "eachClose", location });
      } else if (expr.startsWith("#cols ")) {
        tags.push({ expression: expr, type: "colsOpen", location });
      } else if (expr === "/cols") {
        tags.push({ expression: expr, type: "colsClose", location });
      } else if (expr.startsWith("%")) {
        tags.push({ expression: expr, type: "image", location });
      } else if (expr.startsWith("&")) {
        tags.push({ expression: expr, type: "richText", location });
      } else if (expr.startsWith(">")) {
        tags.push({ expression: expr, type: "subDocument", location });
      } else if (expr.startsWith("^")) {
        tags.push({ expression: expr, type: "chart", location });
      } else if (expr.startsWith("!")) {
        tags.push({ expression: expr, type: "htmlChunk", location });
      } else {
        tags.push({ expression: expr, type: "variable", location });
      }
    }
  }

  // Scan body
  scanBlocks(doc.body, "body");

  // Scan headers
  if (doc.headers) {
    for (const [key, headerDef] of doc.headers) {
      scanBlocks(headerDef.content.children as BodyContent[], `header:${key}`);
    }
  }

  // Scan footers
  if (doc.footers) {
    for (const [key, footerDef] of doc.footers) {
      scanBlocks(footerDef.content.children as BodyContent[], `footer:${key}`);
    }
  }

  // Scan footnotes
  if (doc.footnotes) {
    for (const note of doc.footnotes) {
      if (note.id <= 0) {
        continue;
      }
      scanBlocks(note.content as BodyContent[], `footnote:${note.id}`);
    }
  }

  // Scan endnotes
  if (doc.endnotes) {
    for (const note of doc.endnotes) {
      if (note.id <= 0) {
        continue;
      }
      scanBlocks(note.content as BodyContent[], `endnote:${note.id}`);
    }
  }

  if (includeComments && doc.comments) {
    for (const comment of doc.comments) {
      scanBlocks(comment.content as BodyContent[], `comment:${comment.id}`);
    }
  }

  return tags;
}

// =============================================================================
// Enhanced value resolution for image/richText/subDocument
// =============================================================================

function isTemplateImage(v: unknown): v is TemplateImage {
  return v != null && typeof v === "object" && "image" in v && "width" in v && "height" in v;
}

function isTemplateRichText(v: unknown): v is TemplateRichText {
  if (!Array.isArray(v)) {
    return false;
  }
  // Check that it looks like an array of Run objects (has .content array)
  return v.length > 0 && "content" in v[0] && Array.isArray(v[0].content);
}

function isTemplateSubDocument(v: unknown): v is TemplateSubDocument {
  if (!Array.isArray(v)) {
    return false;
  }
  // Check that it looks like body content blocks (has .type)
  return v.length > 0 && "type" in v[0];
}

export function isTemplateChart(v: unknown): v is TemplateChart {
  return v != null && typeof v === "object" && "chart" in v && !("type" in v);
}

function isTemplateHtmlChunk(v: unknown): v is TemplateHtmlChunk {
  return (
    v != null &&
    typeof v === "object" &&
    "html" in v &&
    typeof (v as TemplateHtmlChunk).html === "string"
  );
}

/**
 * Fill a DOCX template document with data.
 *
 * Resolves variables, conditionals and loops in the body, headers, footers,
 * footnotes, endnotes and comment bodies. Pass `richPlaceholders: true` to
 * also resolve image, rich-text, sub-document, chart and HTML-chunk
 * placeholders — inside loops too, against the item being repeated.
 *
 * Returns a new document; the input is never modified, and neither is the
 * template when the result is edited afterwards: every top-level collection of
 * the result is its own, so filling one template many times and calling
 * `Document.addStyle` / `Document.addImage` / … on each result never reaches
 * the template or the other results. Unchanged nodes inside those collections
 * are shared, which no `Document.*` helper can observe.
 *
 * The deprecated names keep what they did: `fillTemplate` fills in place and
 * resolves no rich placeholder; `fillTemplateEnhanced` returns a new document
 * and always resolves them; neither fills comment bodies.
 *
 * Does not modify `doc`; the result may share unchanged nodes with it — see
 * `Document.clone`.
 *
 * @example
 * ```ts
 * import { Io, Template } from "documonster/word";
 *
 * const doc = await Io.read(bytes);
 * const filled = Template.fill(doc, { name: "Ada" });
 * ```
 */
export function fill(
  doc: DocxDocument,
  data: Readonly<Record<string, unknown>> | ReadonlyMap<string, unknown>,
  options?: TemplateOptions
): DocxDocument {
  return fillDocument(doc, data, options, "Template.fill");
}

/** {@link fill} in the name of `where` (`Package.fillTemplate`, `Template.fillFromSource`). */
export function fillDocument(
  doc: DocxDocument,
  data: unknown,
  options: TemplateOptions | undefined,
  where: string
): DocxDocument {
  checkHandle(doc, "document", "doc", where);
  const record = fillData(data, where);
  checkFillArguments(record, options, where);
  return detachCollections(
    fillInPlace(copyFillTargets(doc), record, options, {
      where,
      rich: options?.richPlaceholders === true,
      comments: true
    })
  );
}

/**
 * Give a filled document its own copy of every top-level collection (styles,
 * images, numbering, comments, fonts, custom properties, headers, …), so the
 * `Document.*` helpers — which add to and replace entries in exactly those
 * collections, and assign top-level fields — reach the result alone.
 *
 * The nodes inside are shared with the template wherever the fill left them
 * unchanged (a style, an image's bytes, an untouched paragraph). Nothing in
 * this library edits a node in place, so sharing is invisible to its API; a
 * caller who mutates a node object directly (`filled.styles[0].name = …`
 * through a cast) reaches the template, as with any shallow copy. Copying the
 * collections is proportional to their length, not to the document; copying
 * every node would make each fill as expensive as a deep clone.
 */
function detachCollections(filled: DocxDocument): DocxDocument {
  const out: Record<string, unknown> = { ...filled };
  for (const key in out) {
    const value = out[key];
    if (Array.isArray(value)) {
      out[key] = [...value];
    } else if (value instanceof Map) {
      out[key] = new Map(value);
    }
  }
  return out as unknown as DocxDocument;
}

/**
 * A copy of `doc` that {@link fillInPlace} can write into without reaching the
 * original: it replaces the body, the children of each header and footer, the
 * content of each note and comment, and the images, so those containers are
 * copied; the nodes inside them are rebuilt by the fill, not edited.
 */
function copyFillTargets(doc: DocxDocument): DocxDocument {
  const copyParts = <T extends { readonly content: object }>(
    parts: ReadonlyMap<string, T> | undefined
  ): Map<string, T> | undefined =>
    parts && new Map(Array.from(parts, ([k, v]) => [k, { ...v, content: { ...v.content } }]));
  return {
    ...doc,
    body: [...doc.body],
    headers: copyParts(doc.headers),
    footers: copyParts(doc.footers),
    footnotes: doc.footnotes?.map(n => ({ ...n })),
    endnotes: doc.endnotes?.map(n => ({ ...n })),
    comments: doc.comments?.map(c => ({ ...c }))
  };
}

/**
 * Fill a template, resolving image (`{{%img}}`), rich-text (`{{&rt}}`),
 * sub-document (`{{>sub}}`), chart (`{{^chart}}`) and HTML-chunk
 * (`{{!html}}`) placeholders as well as variables, conditionals and loops.
 * Returns a new document; the input is left untouched. Comment bodies are not
 * filled.
 *
 * Image placeholder: {{%imagePath}} — value must be a TemplateImage object.
 * RichText placeholder: {{&richPath}} — value must be TemplateRichText (Run[]).
 * SubDocument placeholder: {{>subPath}} — value must be TemplateSubDocument (BodyContent[]).
 * Chart placeholder: {{^chartPath}} — value must be a TemplateChart object.
 */
export function fillTemplateEnhanced(
  doc: DocxDocument,
  data: Record<string, unknown>,
  options?: LegacyTemplateOptions
): DocxDocument {
  checkFillArguments(data, options, "Template.fillTemplateEnhanced");
  return fillInPlace(copyFillTargets(doc), data, options, {
    where: "Template.fillTemplateEnhanced",
    rich: true,
    comments: false
  });
}

/**
 * What a paragraph that is one rich placeholder becomes — the blocks and
 * anchors to put in its place — or `undefined` when it is not one, or its
 * value is not of the placeholder's kind (it is then filled as ordinary text).
 * Resolved against `ctx.data`, so inside a loop `{{%.photo}}` is the item's.
 *
 * A placeholder replaced by something other than a paragraph (or by nothing)
 * leaves the paragraph's range markers and comment references where it was,
 * around what replaced it.
 */
function resolveRichParagraph(
  para: Paragraph,
  ctx: TemplateContext,
  location: string
): (BodyContent | RangeAnchor)[] | undefined {
  const rich = ctx.rich!;
  const { open, close, data } = ctx;
  const text = extractParagraphText(para);
  if (!text.includes(open)) {
    return undefined;
  }
  const replaceBlock = (replacement: readonly (BodyContent | RangeAnchor)[]) => {
    const { leading, trailing } = collectRangeAnchors([para]);
    return [...leading, ...replacement, ...trailing];
  };

  // HTML chunk placeholder ({{!htmlPath}})
  const htmlMatch = matchSinglePlaceholder(text, open, close, "!");
  if (htmlMatch) {
    const value = resolvePath(data, htmlMatch);
    if (isTemplateHtmlChunk(value)) {
      const seq = nextAltChunkSeq();
      const altChunk: AltChunk = {
        type: "altChunk",
        // A process-monotonic sequence, so two fills of the same template
        // don't collide on rId or path.
        rId: `rId_altchunk_${seq}`,
        contentType: value.contentType ?? "text/html",
        data: utf8Encoder.encode(value.html),
        fileName: `afchunk${seq}.html`
      };
      return replaceBlock([altChunk]);
    }
  }
  // Sub-document placeholder: the blocks are inserted, then filled like the
  // rest of the template.
  const subMatch = matchSinglePlaceholder(text, open, close, ">");
  if (subMatch) {
    const value = resolvePath(data, subMatch);
    if (isTemplateSubDocument(value)) {
      if (rich.depth >= MAX_SUB_DOCUMENT_DEPTH) {
        throw new DocxTemplateError(
          `Sub-document placeholders nest more than ${MAX_SUB_DOCUMENT_DEPTH} deep: {{>${subMatch}}}`,
          `>${subMatch}`,
          location,
          { tagName: `>${subMatch}`, sectionPath: location }
        );
      }
      const nested: TemplateContext = { ...ctx, rich: { ...rich, depth: rich.depth + 1 } };
      return replaceBlock(processBodyItems(cloneModel([...value]), nested, location));
    }
  }
  // Image placeholder, optionally conditional: {{%?condPath.imagePath}}
  const imgMatch = matchSinglePlaceholder(text, open, close, "%");
  if (imgMatch) {
    let imagePath = imgMatch;
    if (imgMatch.startsWith("?")) {
      const rest = imgMatch.slice(1);
      const dotIdx = rest.indexOf(".");
      if (dotIdx !== -1) {
        if (!isTruthy(resolvePath(data, rest.slice(0, dotIdx)))) {
          // Condition is falsy — nothing is inserted.
          return replaceBlock([]);
        }
        imagePath = rest.slice(dotIdx + 1);
      }
    }
    const value = resolvePath(data, imagePath);
    if (isTemplateImage(value)) {
      return [replaceParagraphChildren(para, [{ content: [templateImageContent(value, rich)] }])];
    }
  }
  // Rich text placeholder: the runs are inserted as they are — they are data,
  // not template text, so they are not filled again.
  const rtMatch = matchSinglePlaceholder(text, open, close, "&");
  if (rtMatch) {
    const value = resolvePath(data, rtMatch);
    if (isTemplateRichText(value)) {
      return [replaceParagraphChildren(para, cloneModel([...value]) as ParagraphChild[])];
    }
  }
  // Chart placeholder
  const chartMatch = matchSinglePlaceholder(text, open, close, "^");
  if (chartMatch) {
    const value = resolvePath(data, chartMatch);
    if (isTemplateChart(value)) {
      const chartContent: ChartContent = {
        type: "chart",
        chart: value.chart,
        altText: value.altText,
        name: value.name
      };
      return replaceBlock([chartContent]);
    }
  }
  return undefined;
}

/**
 * The inline image a picture placeholder becomes, registering the picture for
 * packaging. The same supplied picture placed many times — by a loop, say —
 * is stored once and referenced from every place.
 */
function templateImageContent(value: TemplateImage, rich: RichState): InlineImageContent {
  const safeFileName = sanitizeTemplateImageFileName(value.image.fileName);
  const rId = rich.images.place(value.image);
  return {
    type: "image",
    rId,
    width: value.width,
    height: value.height,
    altText: value.altText ?? safeFileName,
    name: safeFileName
  };
}

/** Match a single-placeholder paragraph: {{<prefix><path>}} */
function matchSinglePlaceholder(
  text: string,
  open: string,
  close: string,
  prefix: string
): string | null {
  const trimmed = text.trim();
  if (!trimmed.startsWith(open) || !trimmed.endsWith(close)) {
    return null;
  }
  const inner = trimmed.slice(open.length, trimmed.length - close.length).trim();
  if (!inner.startsWith(prefix)) {
    return null;
  }
  return inner.slice(prefix.length).trim();
}

// =============================================================================
// Table Column Loop: {{#cols items}}...{{/cols}}
// =============================================================================

/**
 * `table` with its `{{#cols items}}…{{/cols}}` columns repeated once per item
 * and each copy filled for that item (leniently: an unresolved variable in a
 * column copy renders empty). The comment references of columns removed for
 * an empty collection go to `references`.
 */
function processTableColumnLoop(
  table: Table,
  ctx: TemplateContext,
  references: CommentReference[]
): Table {
  const { open, close, data } = ctx;
  // Check if any cell in the first row contains a #cols directive
  if (table.rows.length === 0) {
    return table;
  }

  const firstRow = table.rows[0];
  let colsExpr: string | null = null;
  let templateColIdx = -1;

  for (let c = 0; c < firstRow.cells.length; c++) {
    const cell = firstRow.cells[c];
    for (const block of cell.content) {
      if (block.type === "paragraph") {
        const text = extractParagraphText(block as Paragraph);
        const pattern = open + "#cols ";
        if (text.includes(pattern)) {
          colsExpr = extractDirectiveFromText(text, open, close, "#cols ");
          templateColIdx = c;
          break;
        }
      }
    }
    if (colsExpr) {
      break;
    }
  }

  if (!colsExpr || templateColIdx === -1) {
    return table;
  }

  // Find closing {{/cols}} column
  let closeColIdx = templateColIdx;
  for (let c = templateColIdx; c < firstRow.cells.length; c++) {
    const cell = firstRow.cells[c];
    for (const block of cell.content) {
      if (block.type === "paragraph") {
        const text = extractParagraphText(block as Paragraph);
        if (text.includes(open + "/cols" + close) || text.includes(open + " /cols " + close)) {
          closeColIdx = c;
          break;
        }
      }
    }
  }

  const items = resolvePath(data, colsExpr);
  if (!Array.isArray(items)) {
    return table;
  }

  const colsOpenRegex = new RegExp(
    escapeRegex(open) +
      "\\s*#cols\\s+[^" +
      escapeRegex(close.charAt(0)) +
      "]*" +
      escapeRegex(close),
    "g"
  );
  const colsCloseRegex = new RegExp(escapeRegex(open) + "\\s*/cols\\s*" + escapeRegex(close), "g");

  // For each row, duplicate the template columns for each item
  const newRows: TableRow[] = [];
  for (const row of table.rows) {
    const beforeCells = row.cells.slice(0, templateColIdx);
    const templateCells = row.cells.slice(templateColIdx, closeColIdx + 1);
    const afterCells = row.cells.slice(closeColIdx + 1);

    const expandedCells: TableCell[] = [...beforeCells];
    const ranges = repeatedRanges(templateCells);

    for (let idx = 0; idx < items.length; idx++) {
      const itemCtx: TemplateContext = {
        ...ctx,
        strict: false,
        data: buildLoopData(items[idx], idx, data)
      };
      // One iteration's cells are one copy of the repeated content.
      const clonedCells = keepClonedRangeMarkers(
        cloneModel(templateCells),
        idx,
        items.length,
        ranges
      );
      for (const clonedCell of clonedCells) {
        // Strip #cols and /cols directives from the text, then fill the copy.
        const stripped = clonedCell.content.map(block => {
          if (block.type !== "paragraph") {
            return block;
          }
          const text = extractParagraphText(block as Paragraph);
          const cleaned = text.replace(colsOpenRegex, "").replace(colsCloseRegex, "");
          return cleaned === text ? block : rebuildParagraphText(block as Paragraph, cleaned);
        });
        expandedCells.push({
          ...clonedCell,
          content: processCellContent(stripped, itemCtx, "colLoop")
        });
      }
    }

    if (items.length === 0) {
      // The template columns go; the ranges their cells opened or closed stay
      // paired, and the comments they referenced keep their references.
      const dropped: RangeMarker[] = [];
      for (const anchor of listTableRowAnchors([
        {
          ...row,
          cells: templateCells,
          rangeMarkersBefore: undefined,
          rangeMarkersAfter: undefined
        }
      ])) {
        if (anchor.type === "commentReference") {
          references.push(anchor);
        } else {
          dropped.push(anchor);
        }
      }
      if (dropped.length > 0) {
        if (afterCells.length > 0) {
          afterCells[0] = withRangeMarkersBefore(afterCells[0]!, dropped);
        } else {
          newRows.push({
            ...row,
            cells: expandedCells,
            rangeMarkersAfter: [...dropped, ...(row.rangeMarkersAfter ?? [])]
          });
          continue;
        }
      }
    }
    expandedCells.push(...afterCells);

    // Update grid (tblGrid) if needed: we'll leave that to the caller
    newRows.push({ ...row, cells: expandedCells });
  }

  return { ...table, rows: newRows };
}

// =============================================================================
// rId allocation helpers
// =============================================================================
//
// Template-generated parts (alt chunks, dynamic images) need stable but
// non-colliding rIds. Earlier versions of this module derived the rId
// from `blockIndex` or the image's `fileName`, which collided as soon as
// the same template was filled twice or two distinct images shared a
// file name. We now use module-level monotonic counters so each generated
// part gets a unique rId for the lifetime of the process.

let _altChunkSeq = 0;
let _templateImageSeq = 0;
function nextAltChunkSeq(): number {
  _altChunkSeq++;
  return _altChunkSeq;
}
function nextTemplateImageSeq(): number {
  _templateImageSeq++;
  return _templateImageSeq;
}

/** Where a fill registers the pictures its placeholders show. */
interface TemplateImageSink {
  /**
   * Register a picture a placeholder shows and return the key its run names.
   * The key is the picture's own `rId` when that is free or already names the
   * same picture — supplying one of the template's own images reuses it — and
   * a fresh one otherwise, so a caller's picture never takes over a key the
   * template uses for another.
   */
  place(image: ImageDef): string;
  /** The pictures registered, for the document's image table. */
  readonly added: ImageDef[];
}

function createTemplateImageSink(
  existing: readonly ImageDef[] | undefined,
  where: string
): TemplateImageSink {
  const owners = new Map<string, ImageDef>();
  for (const img of existing ?? []) {
    for (const key of imageKeys(img)) {
      if (!owners.has(key)) {
        owners.set(key, img);
      }
    }
  }
  const usedFileNames = new Set((existing ?? []).map(img => img.fileName));
  const placed = new Map<ImageDef, string>();
  const added: ImageDef[] = [];
  return {
    added,
    place(image) {
      const again = placed.get(image);
      if (again !== undefined) {
        return again;
      }
      // The picture came from the fill's data.
      assertImageMediaType(image.mediaType, where, "data");
      const wanted = image.rId;
      const owner = wanted !== undefined ? owners.get(wanted) : undefined;
      if (wanted !== undefined && owner && isSamePicture(owner, image)) {
        placed.set(image, wanted);
        return wanted;
      }
      let key = wanted !== undefined && !owner ? wanted : `rId_tpl_${nextTemplateImageSeq()}`;
      while (owners.has(key)) {
        key = `rId_tpl_${nextTemplateImageSeq()}`;
      }
      const safe = sanitizeTemplateImageFileName(image.fileName);
      let fileName = safe;
      for (let n = 2; usedFileNames.has(fileName); n++) {
        const dot = safe.lastIndexOf(".");
        fileName = dot > 0 ? `${safe.slice(0, dot)}_${n}${safe.slice(dot)}` : `${safe}_${n}`;
      }
      const { aliasRIds: _aliases, ...rest } = image;
      const def: ImageDef = { ...rest, rId: key, fileName };
      added.push(def);
      owners.set(key, def);
      usedFileNames.add(fileName);
      placed.set(image, key);
      return key;
    }
  };
}

/**
 * Strip path-traversal segments and other unsafe characters from a
 * caller-supplied image file name so it can be safely used as a ZIP
 * entry path inside the package. Mirrors the equivalent helper in
 * `convert/odt/odt.ts`.
 */
function sanitizeTemplateImageFileName(raw: string | undefined): string {
  if (!raw) {
    return "image.bin";
  }
  const lastSep = Math.max(raw.lastIndexOf("/"), raw.lastIndexOf("\\"));
  let leaf = lastSep >= 0 ? raw.substring(lastSep + 1) : raw;
  while (leaf.startsWith(".")) {
    leaf = leaf.substring(1);
  }
  leaf = leaf.replace(/[^A-Za-z0-9._-]/g, "_");
  leaf = leaf.replace(/\.{2,}/g, ".");
  return leaf || "image.bin";
}
