/**
 * DOCX → Semantic Document Converter
 *
 * Transforms a DocxDocument into a format-agnostic SemanticDocument.
 * This is the bridge between OOXML-specific types and the universal IR
 * that HTML/Markdown/other renderers can consume.
 *
 * Handles:
 * - Heading detection (via style name or outlineLevel)
 * - Inline formatting resolution
 * - Hyperlink extraction
 * - Image registration into ConversionContext
 * - Table structure with merge (colSpan/rowSpan)
 * - List/numbering detection: consecutive numbered paragraphs are aggregated
 *   into ordered/unordered `list` blocks with nested sub-lists by level
 * - Footnote/endnote reference and content
 * - Math content (text fallback)
 */

import { ommlToMathML } from "@word/advanced/math-convert";
import type {
  ConversionContext,
  ResolvedFormatting,
  SemanticBlock,
  SemanticDocument,
  SemanticInline,
  SemanticListItem,
  SemanticNote,
  SemanticParagraphStyle,
  SemanticTableCell,
  SemanticTableRow
} from "@word/convert/conversion-ir";
import { createConversionContext } from "@word/convert/conversion-ir";
import { extractMathText, isRun } from "@word/core/text-utils";
import type {
  AltChunk,
  BodyContent,
  ChartContent,
  ChartExContent,
  CheckBox,
  ColorSpec,
  DocxDocument,
  FontSpec,
  Hyperlink,
  InlineImageContent,
  MathBlock,
  OpaqueDrawing,
  Paragraph,
  ParagraphChild,
  Run,
  RunContent,
  RunProperties,
  Table
} from "@word/types";
import { EMU_PER_INCH } from "@word/units";

// =============================================================================
// Public API
// =============================================================================

/** Options for the semantic conversion. */
export interface DocxToSemanticOptions {
  /** Include footnotes in the output. Default: true. */
  readonly includeFootnotes?: boolean;
  /** Include endnotes in the output. Default: true. */
  readonly includeEndnotes?: boolean;
  /** Extract images and register as assets. Default: true. */
  readonly extractImages?: boolean;
}

/**
 * Convert a DocxDocument to a SemanticDocument (format-agnostic IR).
 *
 * @param doc - The parsed DOCX document.
 * @param options - Conversion options.
 * @returns The semantic document with warnings and assets.
 */
export function docxToSemantic(
  doc: DocxDocument,
  options?: DocxToSemanticOptions
): { document: SemanticDocument; context: ConversionContext } {
  const ctx = createConversionContext();
  const opts: Required<DocxToSemanticOptions> = {
    includeFootnotes: options?.includeFootnotes ?? true,
    includeEndnotes: options?.includeEndnotes ?? true,
    extractImages: options?.extractImages ?? true
  };

  // Build image lookup by rId for quick reference
  const imageMap = new Map<string, { data: Uint8Array; mimeType: string; fileName: string }>();
  if (doc.images && opts.extractImages) {
    for (const img of doc.images) {
      if (img.rId && img.data) {
        const ext = img.fileName?.split(".").pop()?.toLowerCase() ?? "png";
        imageMap.set(img.rId, {
          data: img.data,
          mimeType: extToMimeType(ext),
          fileName: img.fileName
        });
      }
    }
  }

  // Convert body
  const blocks = convertBodyContent(doc.body, doc, ctx, imageMap);

  // Convert footnotes
  const footnotes: SemanticNote[] = [];
  if (opts.includeFootnotes && doc.footnotes) {
    for (const note of doc.footnotes) {
      if (note.id <= 0) {
        continue;
      }
      const noteBlocks = convertBodyContent(
        note.content as readonly BodyContent[],
        doc,
        ctx,
        imageMap
      );
      footnotes.push({ id: note.id, children: noteBlocks });
    }
  }

  // Convert endnotes
  const endnotes: SemanticNote[] = [];
  if (opts.includeEndnotes && doc.endnotes) {
    for (const note of doc.endnotes) {
      if (note.id <= 0) {
        continue;
      }
      const noteBlocks = convertBodyContent(
        note.content as readonly BodyContent[],
        doc,
        ctx,
        imageMap
      );
      endnotes.push({ id: note.id, children: noteBlocks });
    }
  }

  const document: SemanticDocument = {
    blocks,
    assets: ctx.assets,
    footnotes,
    endnotes,
    metadata: {
      title: doc.coreProperties?.title,
      author: doc.coreProperties?.creator,
      subject: doc.coreProperties?.subject
    }
  };

  return { document, context: ctx };
}

// =============================================================================
// Internal: Body Content Conversion
// =============================================================================

function convertBodyContent(
  body: readonly BodyContent[],
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticBlock[] {
  const blocks: SemanticBlock[] = [];

  for (let bodyIndex = 0; bodyIndex < body.length; bodyIndex++) {
    const item = body[bodyIndex];
    switch (item.type) {
      case "paragraph":
        // A run of consecutive list-item paragraphs (each carrying a
        // numbering reference, and not a heading) is aggregated into a single
        // semantic `list` block with nested sub-lists driven by the numbering
        // level. This is what turns Word numbering into real <ul>/<ol> in
        // HTML and `-`/`1.` markers in Markdown when downstream renderers
        // consume the IR.
        if (isListItemParagraph(item)) {
          let end = bodyIndex;
          while (end < body.length) {
            const next = body[end];
            if (next.type !== "paragraph" || !isListItemParagraph(next)) {
              break;
            }
            end++;
          }
          const listParas = body.slice(bodyIndex, end) as Paragraph[];
          blocks.push(...buildListBlocks(listParas, doc, ctx, imageMap));
          bodyIndex = end - 1; // loop's ++ advances past the consumed run
          break;
        }
        blocks.push(convertParagraph(item, doc, ctx, imageMap));
        break;
      case "table":
        blocks.push(convertTable(item, doc, ctx, imageMap));
        break;
      case "floatingImage": {
        if (item.rId && imageMap.has(item.rId)) {
          const img = imageMap.get(item.rId)!;
          const assetId = ctx.registerAsset(img.mimeType, img.data, item.altText);
          blocks.push({
            type: "image",
            assetId,
            alt: item.altText,
            width: item.width ? item.width / EMU_PER_INCH : undefined, // EMU to inches
            height: item.height ? item.height / EMU_PER_INCH : undefined
          });
        }
        break;
      }
      case "sdt": {
        // Structured document tags can wrap whole sections of body content
        // (Paragraph | Table | Run). Pass through paragraph/table children
        // recursively; bare Run children are skipped here because the
        // semantic model has no inline-only equivalent — the warning
        // surfaces them so downstream tooling knows to expect missing text.
        const inner: BodyContent[] = [];
        let dropped = 0;
        for (const c of item.content) {
          if ("type" in c && (c.type === "paragraph" || c.type === "table")) {
            inner.push(c as BodyContent);
          } else {
            dropped++;
          }
        }
        if (dropped > 0) {
          ctx.addWarning(
            "info",
            "sdt-inline-runs-dropped",
            `${dropped} inline Run children of an SDT had no semantic equivalent and were skipped`
          );
        }
        blocks.push(...convertBodyContent(inner, doc, ctx, imageMap));
        break;
      }
      case "textBox": {
        // Render textBox content as inline paragraphs in reading order.
        // The semantic model has no "frame" concept; flattening preserves
        // the text rather than dropping it.
        blocks.push(
          ...convertBodyContent(item.content as readonly BodyContent[], doc, ctx, imageMap)
        );
        break;
      }
      case "drawingShape": {
        if (item.textContent && item.textContent.length > 0) {
          blocks.push(
            ...convertBodyContent(item.textContent as readonly BodyContent[], doc, ctx, imageMap)
          );
        }
        break;
      }
      case "tableOfContents": {
        if (item.cachedParagraphs && item.cachedParagraphs.length > 0) {
          blocks.push(
            ...convertBodyContent(
              item.cachedParagraphs as readonly BodyContent[],
              doc,
              ctx,
              imageMap
            )
          );
        } else {
          // No cached paragraphs (TOC was never updated). Emit a placeholder
          // so the user-visible flow still acknowledges the TOC; encourage
          // calling `updateTableOfContents()` before conversion.
          blocks.push({
            type: "paragraph",
            children: [{ type: "text", text: "[Table of Contents]" }]
          });
          ctx.addWarning(
            "info",
            "toc-not-cached",
            "Table of contents has no cached paragraphs; output contains a placeholder. Run updateTableOfContents() before conversion to populate."
          );
        }
        break;
      }
      case "math": {
        // Block-level math equation. Always provide a plain-text fallback so
        // markdown / plain-text consumers have something to emit; attach
        // MathML when conversion succeeds.
        const mb = item as MathBlock;
        const text = extractMathText(mb.content);
        let mathML: string | undefined;
        try {
          mathML = ommlToMathML(mb.content);
        } catch (err) {
          ctx.addWarning(
            "info",
            "math-mathml-failed",
            `Failed to convert math block to MathML: ${(err as Error).message}`
          );
        }
        blocks.push({ type: "math", text, mathML });
        break;
      }
      case "chart":
      case "chartEx": {
        const c = item as ChartContent | ChartExContent;
        // chartId is meant to be a stable, unique reference handle —
        // not a human-readable label. Use the body position so two
        // charts with the same title still get distinct ids; fall
        // back to the source `name` when present (docx authoring
        // tools typically set a stable name like "Chart 3"); finally
        // synthesise from the position.
        const chartId = c.name ?? `chart-body-${bodyIndex}`;
        const title = c.type === "chart" ? c.chart?.title : undefined;
        blocks.push({
          type: "chart",
          chartId,
          title,
          altText: c.altText
        });
        break;
      }
      case "checkBox": {
        const cb = item as CheckBox;
        blocks.push({ type: "checkBox", checked: cb.checked === true });
        break;
      }
      case "altChunk": {
        const ac = item as AltChunk;
        if (!ac.contentType) {
          // ECMA-376 §17.17.1: the altChunk's content type comes from
          // the related part's `[Content_Types].xml` override. A
          // missing `contentType` after parsing means the source
          // document is malformed (the override isn't there) — surface
          // a warning so callers don't silently consume the
          // octet-stream fallback as if it were the real type.
          ctx.addWarning(
            "warning",
            "altchunk-missing-content-type",
            `altChunk (rId=${ac.rId}) has no contentType; defaulted to application/octet-stream. Source document is missing the <Override ContentType="…"/> entry.`
          );
        }
        blocks.push({
          type: "embed",
          contentType: ac.contentType ?? "application/octet-stream",
          data: ac.data,
          fileName: ac.fileName
        });
        break;
      }
      case "opaqueDrawing": {
        const od = item as OpaqueDrawing;
        blocks.push({ type: "raw", format: "ooxml-drawing", xml: od.rawXml });
        break;
      }
      default: {
        // Compile-time exhaustiveness guard: adding a new BodyContent variant
        // without a corresponding case here is now a build error rather than
        // a silent runtime drop.
        const _exhaustive: never = item;
        ctx.addWarning(
          "warning",
          "internal-bug",
          `Unhandled BodyContent variant: ${(_exhaustive as { type: string }).type}`
        );
        break;
      }
    }
  }

  return blocks;
}

// =============================================================================
// Internal: List Aggregation
// =============================================================================

/**
 * Whether a body paragraph should render as a list item: it carries a
 * numbering reference and is not itself a heading (a numbered heading stays a
 * heading, mirroring the markdown/html renderers).
 */
function isListItemParagraph(item: BodyContent): boolean {
  if (item.type !== "paragraph") {
    return false;
  }
  return item.properties?.numbering !== undefined && detectHeadingLevel(item) === null;
}

/**
 * Resolve a numbering reference to its number format string (e.g. "decimal",
 * "bullet"). Mirrors the lookup in the markdown/html renderers so the three
 * surfaces classify ordered vs. unordered lists identically. Defaults to
 * "bullet" when the numbering definition can't be resolved.
 */
function getNumberingFormat(doc: DocxDocument, numId: number, level: number): string {
  const instance = doc.numberingInstances?.find(n => n.numId === numId);
  if (!instance) {
    return "bullet";
  }
  const abstractNum = doc.abstractNumberings?.find(a => a.abstractNumId === instance.abstractNumId);
  if (!abstractNum) {
    return "bullet";
  }
  const levelDef = abstractNum.levels.find(l => l.level === level);
  return levelDef?.format ?? "bullet";
}

/** A number format other than "bullet"/"none" denotes an ordered list. */
function isOrderedFormat(format: string): boolean {
  return format !== "bullet" && format !== "none";
}

/**
 * Build one or more semantic `list` blocks from a contiguous run of list-item
 * paragraphs. Paragraphs are nested by their numbering `level`; a deeper level
 * becomes a `subList` of the preceding shallower item. Adjacent items that
 * switch between ordered and unordered at the same level start a new sibling
 * list so the ordered/unordered distinction is preserved.
 */
function buildListBlocks(
  paras: readonly Paragraph[],
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticBlock[] {
  const { blocks } = buildListLevel(paras, 0, 0, doc, ctx, imageMap);
  return blocks;
}

/**
 * Consume paragraphs starting at `start` that belong to `level` (or deeper),
 * emitting sibling lists for this level. Deeper-level paragraphs are folded
 * into the current item's `subList` via recursion. Returns the produced blocks
 * and the index of the first paragraph that no longer belongs to this level.
 */
function buildListLevel(
  paras: readonly Paragraph[],
  start: number,
  level: number,
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): { blocks: SemanticBlock[]; next: number } {
  const blocks: SemanticBlock[] = [];
  let i = start;
  let currentOrdered: boolean | null = null;
  let items: SemanticListItem[] = [];

  const flush = (): void => {
    if (items.length > 0 && currentOrdered !== null) {
      blocks.push({ type: "list", ordered: currentOrdered, items });
      items = [];
    }
  };

  while (i < paras.length) {
    const para = paras[i];
    const num = para.properties?.numbering;
    // Defensive: callers only pass list-item paragraphs, but guard anyway.
    if (!num) {
      break;
    }
    if (num.level < level) {
      // Belongs to a shallower list — let the caller handle it.
      break;
    }
    if (num.level > level) {
      // Deeper item with no shallower parent at this position: descend and
      // attach the nested list to the most recent item, or synthesise an
      // empty item to host it when there is no parent.
      const { blocks: subBlocks, next } = buildListLevel(paras, i, num.level, doc, ctx, imageMap);
      const subList = subBlocks[0];
      if (items.length > 0) {
        const last = items[items.length - 1];
        items[items.length - 1] = { ...last, subList };
      } else if (subList) {
        // Promote the deeper list to this level when there is no parent item.
        if (currentOrdered === null && subList.type === "list") {
          currentOrdered = subList.ordered;
        }
        items.push({ children: [], subList });
      }
      i = next;
      continue;
    }

    // num.level === level
    const format = getNumberingFormat(doc, num.numId, num.level);
    const ordered = isOrderedFormat(format);
    if (currentOrdered === null) {
      currentOrdered = ordered;
    } else if (ordered !== currentOrdered) {
      // Ordered/unordered switch at the same level → start a new sibling list.
      flush();
      currentOrdered = ordered;
    }

    const children = convertParagraphChildren(para.children, doc, ctx, imageMap);
    items.push({ children });
    i++;
  }

  flush();
  return { blocks, next: i };
}

// =============================================================================
// Internal: Paragraph Conversion
// =============================================================================

function convertParagraph(
  para: Paragraph,
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticBlock {
  // Detect heading
  const headingLevel = detectHeadingLevel(para);
  if (headingLevel !== null) {
    const children = convertParagraphChildren(para.children, doc, ctx, imageMap);
    return {
      type: "heading",
      level: headingLevel as 1 | 2 | 3 | 4 | 5 | 6,
      children
    };
  }

  const children = convertParagraphChildren(para.children, doc, ctx, imageMap);

  // Convert paragraph style
  const style = convertParagraphStyle(para.properties);

  return { type: "paragraph", children, style };
}

function detectHeadingLevel(para: Paragraph): number | null {
  const props = para.properties;
  if (!props) {
    return null;
  }

  // Check outlineLevel
  if (props.outlineLevel !== undefined && props.outlineLevel >= 0 && props.outlineLevel <= 5) {
    return props.outlineLevel + 1;
  }

  // Check style name
  if (props.style) {
    const match = /^[Hh]eading\s*(\d)$/i.exec(props.style);
    if (match) {
      const level = parseInt(match[1], 10);
      if (level >= 1 && level <= 6) {
        return level;
      }
    }
  }

  return null;
}

function convertParagraphStyle(props: Paragraph["properties"]): SemanticParagraphStyle | undefined {
  if (!props) {
    return undefined;
  }

  const style: SemanticParagraphStyle = {
    alignment: convertAlignment(props.alignment),
    indentLeft: props.indent?.left ? props.indent.left / 20 : undefined, // twips to points
    indentRight: props.indent?.right ? props.indent.right / 20 : undefined,
    spaceBefore: props.spacing?.before ? props.spacing.before / 20 : undefined,
    spaceAfter: props.spacing?.after ? props.spacing.after / 20 : undefined
  };

  // Only return if at least one property is set
  if (
    style.alignment === undefined &&
    style.indentLeft === undefined &&
    style.indentRight === undefined &&
    style.spaceBefore === undefined &&
    style.spaceAfter === undefined
  ) {
    return undefined;
  }

  return style;
}

function convertAlignment(
  alignment: string | undefined
): "left" | "center" | "right" | "justify" | undefined {
  switch (alignment) {
    case "left":
    case "start":
      return "left";
    case "center":
      return "center";
    case "right":
    case "end":
      return "right";
    case "both":
    case "justify":
      return "justify";
    default:
      return undefined;
  }
}

// =============================================================================
// Internal: Paragraph Children (Inline Content)
// =============================================================================

function convertParagraphChildren(
  children: readonly ParagraphChild[],
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticInline[] {
  const inlines: SemanticInline[] = [];

  for (const child of children) {
    if ("type" in child) {
      const typed = child as { type: string };
      if (typed.type === "hyperlink") {
        const hl = child as Hyperlink;
        const linkChildren: SemanticInline[] = [];
        for (const run of hl.children) {
          linkChildren.push(...convertRun(run, ctx, imageMap));
        }
        inlines.push({
          type: "link",
          href: hl.url ?? hl.anchor ?? "",
          children: linkChildren
        });
        continue;
      }
      // Skip other typed children (bookmark, comment range, etc.)
      continue;
    }

    // Default: treat as Run
    if (isRun(child)) {
      inlines.push(...convertRun(child, ctx, imageMap));
    }
  }

  return inlines;
}

function convertRun(
  run: Run,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticInline[] {
  const inlines: SemanticInline[] = [];
  const format = resolveRunFormatting(run.properties);

  for (const content of run.content) {
    const inline = convertRunContent(content, format, ctx, imageMap);
    if (inline) {
      inlines.push(inline);
    }
  }

  return inlines;
}

function convertRunContent(
  content: RunContent,
  format: ResolvedFormatting | undefined,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticInline | null {
  switch (content.type) {
    case "text":
      return { type: "text", text: content.text, format };
    case "break":
      return { type: "lineBreak" };
    case "tab":
      return { type: "text", text: "\t", format };
    case "image": {
      const img = content as InlineImageContent;
      if (img.rId && imageMap.has(img.rId)) {
        const imgData = imageMap.get(img.rId)!;
        const assetId = ctx.registerAsset(imgData.mimeType, imgData.data, img.altText);
        return {
          type: "image",
          assetId,
          alt: img.altText,
          width: img.width ? img.width / EMU_PER_INCH : undefined,
          height: img.height ? img.height / EMU_PER_INCH : undefined
        };
      }
      return null;
    }
    case "footnoteRef":
      return { type: "footnoteRef", id: content.id };
    case "endnoteRef":
      return { type: "endnoteRef", id: content.id };
    case "field":
      // Render field cached value as plain text
      if (content.cachedValue) {
        return { type: "text", text: content.cachedValue, format };
      }
      return null;
    case "symbol":
      return { type: "text", text: content.char ?? "", format };
    default:
      return null;
  }
}

// =============================================================================
// Internal: Run Formatting Resolution
// =============================================================================

function resolveRunFormatting(props: RunProperties | undefined): ResolvedFormatting | undefined {
  if (!props) {
    return undefined;
  }

  const format: ResolvedFormatting = {
    bold: props.bold || undefined,
    italic: props.italic || undefined,
    underline: props.underline !== undefined ? true : undefined,
    strikethrough: props.strike || undefined,
    superscript: props.vertAlign === "superscript" || undefined,
    subscript: props.vertAlign === "subscript" || undefined,
    fontFamily:
      typeof props.font === "string" ? props.font : (props.font as FontSpec | undefined)?.ascii,
    fontSize: props.size ? props.size / 2 : undefined, // half-points to points
    color:
      typeof props.color === "string" ? props.color : (props.color as ColorSpec | undefined)?.val,
    code: undefined
  };

  // Only return if at least one property is set
  if (
    !format.bold &&
    !format.italic &&
    !format.underline &&
    !format.strikethrough &&
    !format.superscript &&
    !format.subscript &&
    !format.fontFamily &&
    !format.fontSize &&
    !format.color
  ) {
    return undefined;
  }

  return format;
}

// =============================================================================
// Internal: Table Conversion
// =============================================================================

function convertTable(
  table: Table,
  doc: DocxDocument,
  ctx: ConversionContext,
  imageMap: Map<string, { data: Uint8Array; mimeType: string; fileName: string }>
): SemanticBlock {
  const rows: SemanticTableRow[] = [];

  for (const row of table.rows) {
    const cells: SemanticTableCell[] = [];
    const isHeader = row.properties?.tableHeader ?? false;

    for (const cell of row.cells) {
      const children = convertBodyContent(
        cell.content as readonly BodyContent[],
        doc,
        ctx,
        imageMap
      );
      const cellAlignment =
        cell.content?.[0]?.type === "paragraph"
          ? (cell.content[0] as Paragraph).properties?.alignment
          : undefined;
      const resolvedAlign = convertAlignment(cellAlignment);
      const cellResult: SemanticTableCell = {
        children,
        colSpan: cell.properties?.gridSpan,
        alignment: resolvedAlign === "justify" ? undefined : resolvedAlign
      };
      cells.push(cellResult);
    }

    rows.push({ cells, isHeader });
  }

  return { type: "table", rows };
}

/**
 * Map a file extension to the canonical IANA MIME type.
 *
 * The previous implementation just did `image/${ext}` which produced
 * non-canonical types like `image/tif` or `image/wmf` — those are
 * silently rejected by browsers when the semantic document is rendered
 * to HTML/PDF. Here we map the handful of extensions DOCX images
 * actually use to their correct types and refuse to forward unknown
 * extensions verbatim — they're returned as
 * `application/octet-stream` instead so attacker-controlled file names
 * can't smuggle a fabricated `image/<script>` Content-Type into
 * downstream renderers.
 */
function extToMimeType(ext: string): string {
  switch (ext) {
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "tif":
    case "tiff":
      return "image/tiff";
    case "svg":
      return "image/svg+xml";
    case "webp":
      return "image/webp";
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    default:
      // Unknown / hostile extension: don't synthesise an `image/<x>`
      // Content-Type from raw input.
      return "application/octet-stream";
  }
}
